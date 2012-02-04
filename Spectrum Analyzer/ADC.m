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
        *magValue = 0.;
        *phaseValue = 0.;
    }
    
    return;
}

@end
