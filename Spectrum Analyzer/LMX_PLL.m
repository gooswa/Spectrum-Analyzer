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
    }
    
    return self;
}

- (double)refFreq
{
    return refFreq;
}

- (void)setRefFreq:(double)newRefFreq
{
    refFreq = newRefFreq;
    
    // Update computed values with the new reference frequency
    [self setPhaseFreq:requestedPhaseFreq];
}

- (double)phaseFreq
{
    return phaseFreq;
}

- (void)setPhaseFreq:(double)newPhaseFreq
{
    requestedPhaseFreq = newPhaseFreq;
    
    // Calculate dividers and calculate new frequencies
    double best_r_divider = refFreq / newPhaseFreq;
    int_r_divider = (uint32)round(best_r_divider);
    phaseFreq = refFreq / int_r_divider;
    
    // Update the output frequency using the requested value
    // This has the effect of re-computing its actual value given
    // the new phase frequency
    [self setOutputFreq:requestedOutputFreq];
}

- (double)outputFreq 
{
    return outputFreq;
}

- (void)setOutputFreq:(double)newOutputFreq
{
    requestedOutputFreq = newOutputFreq;
    
    double best_n_divider = newOutputFreq / phaseFreq;
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
    registers[N_DIVIDER_REGISTER] |= (int_a_divider & A_COUNTER_MASK);
    registers[N_DIVIDER_REGISTER] |= N_DIVIDER_REGISTER;
    
    // This register controls all the functions of the PLL
    registers[FUNCTION_REGISTER]   = (invertingCP)? PD_POLARITY : 0;
    registers[FUNCTION_REGISTER]  |= (FoLD_output & FOLD_CONTROL_MASK) << FOLD_CONTROL_OFFSET;
    registers[FUNCTION_REGISTER]  |= (initialize)?  INITIALIZATION : 0;
    registers[FUNCTION_REGISTER]  |= FUNCTION_REGISTER;
    
    // If the interface is non-null, update the hardware
    if (interface) {
        // Send the registers to the PLL
        [interface sendSPIwithPin:Data
                            Clock:Clock
                               CS:CS
                             data:registers[FUNCTION_REGISTER]
                             bits:21];
        [interface sendSPIwithPin:Data
                            Clock:Clock
                               CS:CS
                             data:registers[R_DIVIDER_REGISTER]
                             bits:21];
        [interface sendSPIwithPin:Data
                            Clock:Clock
                               CS:CS
                             data:registers[N_DIVIDER_REGISTER]
                             bits:21];
    }
}

@end
