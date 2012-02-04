//
//  HardwareInterface.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/21/12.
//  Copyright (c) 2012. All rights reserved.
//

#import "HardwareInterface.h"

@implementation HardwareInterface

- (id)initWithPort:(NSString *)path
{
    return nil;
}

-(void)sendSPIwithPin:(NSInteger)dataPin
                Clock:(NSInteger)clockPin
                   CS:(NSInteger)CSPin
                 data:(long long)data
                 bits:(NSInteger)numBits
{
    return;
}

-(void)setPin:(NSInteger)pin Value:(_Bool)value
{
    return;
}

-(void)readADCdelay:(NSInteger)mSDelay
            samples:(NSInteger)samples
                mag:(double *)magData
              phase:(double *)phaseData
{
    // At least keep the minimum timing characteristics the same
    usleep((int)mSDelay * 1000);
    
    *magData = 0.;
    *phaseData = 0.;
    
    return;
}

@end
