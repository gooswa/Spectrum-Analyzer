//
//  BusPirate.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/18/12.
//  Copyright (c) 2012 Oregon State University (COAS). All rights reserved.
//

#import <Foundation/Foundation.h>
#import <SerialPort/SerialPort.h>

@interface BusPirate : NSObject
{
    SerialPort *port;
}

- (id)initWithPort:(NSString *)path;

// This allows up to 64 bit SPI transfers
// All transfers must be right-justified
- (void)sendSPIwithPin:(NSInteger)dataPin
                 Clock:(NSInteger)clockPin
                    CS:(NSInteger)CSPin
                  data:(long long)data
                  bits:(NSInteger)numBits;

@end
