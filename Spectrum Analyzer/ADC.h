//
//  ADC.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/29/12.
//  Copyright (c) 2012. All rights reserved.
//

#import <Foundation/Foundation.h>

@class HardwareInterface;

@interface ADC : NSObject

- (void)sampleCount:(NSInteger)counts
              Delay:(NSInteger)delay
                Mag:(double *)magValue
              Phase:(double *)phaseValue
          Interface:(HardwareInterface *)interface;

@end
