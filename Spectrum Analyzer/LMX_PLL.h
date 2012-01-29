//
//  LMX_PLL.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/21/12.
//  Copyright (c) 2012. All rights reserved.
//

#import <Foundation/Foundation.h>

@class HardwareInterface;

#define FoLD_TRIS 0 // Tri-state
#define FoLD_RDIV 1 // R Divider output
#define FoLD_NDIV 2 // N Divider output
#define FoLD_SDAT 3 // Serial Data Out
#define FoLD_LDET 4 // Digital lock detect
#define FoLD_LKOD 5 // n Channel Open Drain
#define FoLD_LKAH 6 // Active HIGH
#define FoLD_LKAL 7 // Active LOW

@interface LMX_PLL : NSObject
{
    // Pin numbers for use with the hardware interface
    int CS, Data, Clock;
    
    // LMX chip type
    int type;
    
    // Frequencies
    double refFreq;
    double phaseFreq;
    double outputFreq;
    
    // Chip options
    bool invertingCP;
    char  FoLD_output;
    bool initialize;
    
    @private
    double requestedPhaseFreq;
    double requestedOutputFreq;
    
    uint32 int_r_divider;
    uint32 int_n_divider;
}

@property(readwrite) int CS;
@property(readwrite) int Data;
@property(readwrite) int Clock;
@property(readwrite) double refFreq;
@property(readwrite) double phaseFreq;
@property(readwrite) double outputFreq;
@property(readwrite) bool invertingCP;
@property(readwrite) bool initialize;
@property(readwrite) char FoLD_output;

- (void)updateHardware:(HardwareInterface *)interface;

@end
