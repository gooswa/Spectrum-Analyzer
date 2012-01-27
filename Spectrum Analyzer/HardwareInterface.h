//
//  HardwareInterface.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/21/12.
//  Copyright (c) 2012 Oregon State University (COAS). All rights reserved.
//

#import <Foundation/Foundation.h>

// This is more like a 'virtual class'
// it doesn't do anything on its own.
@interface HardwareInterface : NSObject

- (id)initWithPort:(NSString *)path;

// This allows up to 64 bit SPI transfers
// All transfers must be right-justified
- (void)sendSPIwithPin:(NSInteger)dataPin
                 Clock:(NSInteger)clockPin
                    CS:(NSInteger)CSPin
                  data:(long long)data
                  bits:(NSInteger)numBits;

- (void)setPin:(NSInteger)pin Value:(bool)value;

@end
