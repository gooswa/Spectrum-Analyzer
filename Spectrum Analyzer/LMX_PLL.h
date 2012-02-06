//
//  LMX_PLL.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/21/12.
//  Copyright (c) 2012. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "ModuleDelegate.h"

@class HardwareInterface;

#define FoLD_TRIS 0 // Tri-state
#define FoLD_LDET 1 // Digital lock detect
#define FoLD_NDIV 2 // N Divider output
#define FoLD_LKAH 3 // Active HIGH
#define FoLD_RDIV 4 // R Divider output
#define FoLD_LKOD 5 // n Channel Open Drain
#define FoLD_SDAT 6 // Serial Data Out
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
    char FoLD_output;
    bool initialize;
    
    @private
    double requestedPhaseFreq;
    double requestedOutputFreq;
    
    uint32 int_r_divider;
    uint32 int_n_divider;
    
    uint32 shadow[3];
    
    id<ModuleDelegate> *delegate;
}

@property(assign) id<ModuleDelegate> *delegate;

@property(readwrite) int CS;
@property(readwrite) int Data;
@property(readwrite) int Clock;
@property(readwrite) double refFreq;
@property(readwrite) double phaseFreq;
@property(readwrite) double outputFreq;
@property(readwrite) bool invertingCP;
@property(readwrite) bool initialize;
@property(readwrite) char FoLD_output;
@property(readonly)  int n_divider;
@property(readonly)  int r_divider;

- (void)updateHardware:(HardwareInterface *)interface;

- (void)recalculate;

@end
