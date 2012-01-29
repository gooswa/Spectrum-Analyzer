//
//  AD_DDS.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/22/12.
//  Copyright (c) 2012. All rights reserved.
//

#import <Foundation/Foundation.h>

@class HardwareInterface;

@interface AD_DDS : NSObject
{
    // Pin numbers for use with the hardware interface
    int CS, Data, Clock;
    
    // Frequencies
    double refFreq;
    double phase;
    double outputFreq;
    
    // Chip options
    bool powerDown;
    
    // Delta phase value
    uint32 deltaPhase;
}

@property(readwrite) int CS;
@property(readwrite) int Data;
@property(readwrite) int Clock;
@property(readwrite) double refFreq;
@property(readwrite) double phase;
@property(readwrite) double outputFreq;
@property(readwrite) bool powerDown;

- (void)initHardware:(HardwareInterface *)interface;
- (void)updateHardware:(HardwareInterface *)interface;

@end
