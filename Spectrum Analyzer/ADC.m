//
//  ADC.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/29/12.
//  Copyright (c) 2012. All rights reserved.
//

#import "ADC.h"
#include "HardwareInterface.h"

@implementation ADC

-(void)sampleCount:(NSInteger)counts
             Delay:(NSInteger)delay
               Mag:(double *)magValue 
             Phase:(double *)phaseValue
         Interface:(HardwareInterface *)interface
{
    if (interface) {
        [interface readADCdelay:delay
                        samples:counts
                            mag:magValue
                          phase:phaseValue];
    } else {
        usleep((int)delay * 1000);
        
        float range = ((float)rand() / RAND_MAX) * 500.;
        
        // Simulate some values
        *magValue = 8000 + range;
        
        range = ((float)rand() / RAND_MAX) * 250.;
        *phaseValue = 8192 + range;
    }
    
    return;
}

@end
