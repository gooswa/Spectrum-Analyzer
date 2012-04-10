//
//  LMX_PLL.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/21/12.
//  Copyright (c) 2012. All rights reserved.
//

#import "LMX_PLL.h"
#import "HardwareInterface.h"

@implementation LMX_PLL

@synthesize delegate;
@synthesize CS;
@synthesize Data;
@synthesize Clock;
@synthesize FoLD_output;
@synthesize invertingCP;
@synthesize initialize;

#define R_DIVIDER_MASK 0x3FFF
#define R_DIVIDER_REGISTER 0x0
#define A_COUNTER_MASK 0x1F
#define B_COUNTER_MASK 0x1FFF
#define N_DIVIDER_REGISTER 0x1

#define  FUNCTION_REGISTER 0x2
#define INITIALIZATION          (1 << 0)
#define COUNTER_RESET           (1 << 2)
#define POWER_DOWN              (1 << 3)
#define FOLD_CONTROL_OFFSET     4
#define FOLD_CONTROL_MASK       0x7
#define PD_POLARITY             (1 << 7)
#define CP_TRISTATE             (1 << 8)
#define FASTLOCK_ENABLE         (1 << 9)
#define FASTLOCK_CONTROL        (1 << 10)
#define TIMEOUT_ENABLE          (1 << 11)
#define TIMEOUT_COUNTER_OFFSET  12
#define TIMEOUT_COUNTER_MASK    0xF
#define POWER_DOWN_MODE         (1 << 19)

- (id)init
{
    self = [super init];
    
    if (self) {
        // Set the defaults
        FoLD_output = FoLD_TRIS;
        invertingCP = YES;
        initialize  = YES;
        
        shadow[0] = 0;
        shadow[1] = 0;
        shadow[2] = 0;
    }
    
    return self;
}

- (double)refFreq
{
    return refFreq;
}

- (void)setRefFreq:(double)newRefFreq
{
    if (refFreq != newRefFreq) {
        refFreq  = newRefFreq;
        
        [self recalculate];
        
        if (delegate) {
            if ([(id)delegate respondsToSelector:@selector(moduleDidBecomeStale:)]) {
                [(id)delegate moduleDidBecomeStale:self];
            }
        }
    }
}

- (double)phaseFreq
{
    return phaseFreq;
}

- (void)setPhaseFreq:(double)newPhaseFreq
{
    if (newPhaseFreq != requestedPhaseFreq) {
        requestedPhaseFreq = newPhaseFreq;

        [self recalculate];

        if (delegate) {
            if ([(id)delegate respondsToSelector:@selector(moduleDidBecomeStale:)]) {
                [(id)delegate moduleDidBecomeStale:self];
            }
        }
    }
}

- (double)outputFreq 
{
    return outputFreq;
}

- (void)setOutputFreq:(double)newOutputFreq
{
    if (newOutputFreq != requestedPhaseFreq) {
        requestedOutputFreq = newOutputFreq;
        
        [self recalculate];
        
        if (delegate) {
            if ([(id)delegate respondsToSelector:@selector(moduleDidBecomeStale:)]) {
                [(id)delegate moduleDidBecomeStale:self];
            }
        }
    }
}

- (int)n_divider
{
    return int_n_divider;
}

- (int)r_divider
{
    return int_r_divider;
}

- (void)recalculate
{
    // Calculate dividers and calculate new frequencies
    double best_r_divider = refFreq / requestedPhaseFreq;
    int_r_divider = (uint32)round(best_r_divider);
    phaseFreq = refFreq / int_r_divider;

    double best_n_divider = requestedOutputFreq / phaseFreq;
    int_n_divider = (uint32)round(best_n_divider);
    outputFreq = (double)int_n_divider * phaseFreq;
}

// This function updates the N R and A divider values
// and applies the values to the hardware through the interface
// If the provided interface is 'nil' the values will be converted
// but they won't be transfered to the device.  This can be useful
// when you're iterating to get an acceptable value.
//
// Also the frequency values will be reset with the value they would
// have in the real world, rather than the idealized value.
- (void)updateHardware:(HardwareInterface *)interface
{
    if (delegate) {
        if ([(id)delegate respondsToSelector:@selector(moduleWillUpdateHardware)]) {
            [(id)delegate moduleWillUpdateHardware:self];
        }
    }

    // These are the 3 registers that control the PLL
    uint32 registers[3];

    uint32 prescaler = 32;
    if (type == 2306) {
        prescaler = 8;
    }
    
    uint32 int_b_divider = int_n_divider / prescaler;
    uint32 int_a_divider = int_n_divider % prescaler;
    
    // Generate register values
    // The left-shift is for the control bits
    registers[R_DIVIDER_REGISTER]  = (int_r_divider & R_DIVIDER_MASK) << 2;
    registers[R_DIVIDER_REGISTER] |= R_DIVIDER_REGISTER;
    
    // The N divider is made up of the A and B dividers
    registers[N_DIVIDER_REGISTER]  = (int_b_divider & B_COUNTER_MASK) << 7;
    registers[N_DIVIDER_REGISTER] |= (int_a_divider & A_COUNTER_MASK) << 2;
    registers[N_DIVIDER_REGISTER] |= N_DIVIDER_REGISTER;
    
    // This register controls all the functions of the PLL
    registers[FUNCTION_REGISTER]   = (invertingCP)? 0 : PD_POLARITY;
    registers[FUNCTION_REGISTER]  |= (FoLD_output & FOLD_CONTROL_MASK) << FOLD_CONTROL_OFFSET;
    registers[FUNCTION_REGISTER]  |= (initialize)?  INITIALIZATION : 0;
    registers[FUNCTION_REGISTER]  |= FUNCTION_REGISTER;
    // We don't need to save the initalization command in the shadow,
    // if we're initializing, create an equivalent setting and save it instead
    if (initialize) {
        // This bit is 0 = inverting, 1 = non-inverting
        shadow[FUNCTION_REGISTER]   = (invertingCP)? 0 : PD_POLARITY;
        shadow[FUNCTION_REGISTER]  |= (FoLD_output & FOLD_CONTROL_MASK) << FOLD_CONTROL_OFFSET;
        shadow[FUNCTION_REGISTER]  |= FUNCTION_REGISTER;
    }
    
    // If the interface is non-null, update the hardware
    if (interface) {

        // Send the (changed) registers to the PLL
        if (registers[FUNCTION_REGISTER] != shadow[FUNCTION_REGISTER]) {
            [interface sendSPIwithPin:Data
                                Clock:Clock
                                   CS:CS
                                 data:registers[FUNCTION_REGISTER]
                                 bits:21];
            shadow[FUNCTION_REGISTER] = registers[FUNCTION_REGISTER];
            if (initialize) {
                initialize = NO;
            } else {
                shadow[FUNCTION_REGISTER] = registers[FUNCTION_REGISTER];
            }
        }
        
        if (registers[R_DIVIDER_REGISTER] != shadow[R_DIVIDER_REGISTER]) {
            [interface sendSPIwithPin:Data
                                Clock:Clock
                                   CS:CS
                                 data:registers[R_DIVIDER_REGISTER]
                                 bits:21];
            shadow[R_DIVIDER_REGISTER] = registers[R_DIVIDER_REGISTER];
        }
        
        if (registers[N_DIVIDER_REGISTER] != shadow[N_DIVIDER_REGISTER]) {
            [interface sendSPIwithPin:Data
                                Clock:Clock
                                   CS:CS
                                 data:registers[N_DIVIDER_REGISTER]
                                 bits:21];
            shadow[N_DIVIDER_REGISTER] = registers[N_DIVIDER_REGISTER];
        }
    } else {
//        printf("Tuning to %f (ref(%f) / r(%d)) * n(%d)) n -> ((b)%d * 32 + (a)%d)\n",
//               outputFreq, refFreq, int_r_divider, int_n_divider, int_b_divider, int_a_divider);
        if (registers[FUNCTION_REGISTER] != shadow[FUNCTION_REGISTER]) {
//            printf("Updating the function register: 0x%x\n", registers[FUNCTION_REGISTER]);
            // If we're initializing, don't copy the register into the shadow
            // there is a better version already there.
            if (initialize) {
                initialize = NO;
            } else {
                shadow[FUNCTION_REGISTER] = registers[FUNCTION_REGISTER];
            }
        }
        
        if (registers[R_DIVIDER_REGISTER] != shadow[R_DIVIDER_REGISTER]) {
            printf("Updating the R divider: 0x%x\n", registers[R_DIVIDER_REGISTER]);
            shadow[R_DIVIDER_REGISTER] = registers[R_DIVIDER_REGISTER];
        }
        
        if (registers[N_DIVIDER_REGISTER] != shadow[N_DIVIDER_REGISTER]) {
            printf("Updating the N divider: 0x%x\n", registers[N_DIVIDER_REGISTER]);
            shadow[N_DIVIDER_REGISTER] = registers[N_DIVIDER_REGISTER];
        }
    }
}

@end
