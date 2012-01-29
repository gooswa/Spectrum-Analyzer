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

@end
