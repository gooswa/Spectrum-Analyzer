//
//  AD_DDS.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/22/12.
//  Copyright (c) 2012. All rights reserved.
//

#import "AD_DDS.h"
#import "HardwareInterface.h"

@implementation AD_DDS

@synthesize CS;
@synthesize Data;
@synthesize Clock;
@synthesize refFreq;
@synthesize phase;
@synthesize outputFreq;
@synthesize powerDown;

static uint8 reverse[16] = {
    0x0, // 0000 -> 0000
    0x8, // 0001 -> 1000
    0x4, // 0010 -> 0100
    0xC, // 0011 -> 1100
    0x2, // 0100 -> 0010
    0xA, // 0101 -> 1010
    0x6, // 0110 -> 0110
    0xE, // 0111 -> 1110
    0x1, // 1000 -> 0001
    0x9, // 1001 -> 1001
    0x5, // 1010 -> 0101
    0xD, // 1011 -> 1101
    0x3, // 1100 -> 0011
    0xB, // 1101 -> 1011
    0x7, // 1110 -> 0111
    0xF, // 1111 -> 1111
};

#define RESERVED_WORD_MASK 0xFC

- (id)init
{
    self = [super init];
    
    if (self) {
        // Set the defaults
        outputFreq = 10.7;
    }
    
    return self;
}

-(void)initHardware:(HardwareInterface *)interface
{
    // It should be safe to assume the pins are already set low.
    // Clock the parallel word once (loads serial mode into config)
    [interface setPin:Clock Value:1];
    [interface setPin:Clock Value:0];
    
    // Set FqUD high to latch into serial mode
    [interface setPin:CS Value:1];
    
    // Now, the device is in serial mode, ready for loading
    // If, for some reason, the device was already in serial mode,
    // or is otherwise FUBAR'd, try to clear out the registers by
    // loading all 0's...
    [interface sendSPIwithPin:Data
                        Clock:Clock
                           CS:CS
                         data:0
                         bits:40];
}

// This is applied to the phase/control word (bits 33 through 40)

- (void)updateHardware:(HardwareInterface *)interface
{
    // Calculate the delta phase value
    // This is the equation from the data sheet, but
    // re-arranged to solve for delta phase rather than Fout
    double floatPhase = (outputFreq * exp2(32)) / refFreq;
    if (floatPhase > UINT32_MAX) {
        NSLog(@"The calculated delta phase is greater than uint32");
    }
    
    deltaPhase = (uint32)round(floatPhase);
    
    // Re-calculate the output frequency from the delta phase
    // This is the same equation that's in the datasheet
    outputFreq = ((double)deltaPhase * refFreq) / exp2(32);
    
    // If we have a hardware interface, update the device
    if (interface) {
        uint64 reg = deltaPhase;
        
        // Because the DDS is loaded LSB-first,
        // we need to convert the register here.
        // To do this, we'll reverse each nibble,
        // then load them into a new register backwards.
        uint64 reg2 = 0;
        
        for (int i = 0; i < 10; i++) {
            reg2  = reg2 << 4;
            reg2 |= reverse[(reg >> i*4) & 0xf];
        }
        
        
        [interface sendSPIwithPin:Data
                            Clock:Clock
                               CS:CS
                             data:reg2
                             bits:40];
    }
}

@end
