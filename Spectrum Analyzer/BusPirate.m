//
//  BusPirate.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/18/12.
//  Copyright (c) 2012. All rights reserved.
//

#import "BusPirate.h"

// From the Dangerous prototypes website, for Buspirate 2.3+

//00000000 – Enter raw bitbang mode, reset to raw bitbang mode
#define BINARY_MODE 0x00

//00000001 – SPI mode/rawSPI version string (SPI1)
#define SPI_MODE 0x01

//00000010 – CS low (0)
#define CS_LOW  0x02
#define CS_HIGH 0x03

//00000011 – CS high (1)
#define CS_HIGH 0x03

//00001101 – Sniff all SPI traffic
//00001110 – Sniff when CS low
//00001111 – Sniff when CS high

//0001xxxx – Bulk SPI transfer, send 1-16 bytes (0=1byte!)
#define BULK_SPI 0x10

//0100wxyz – Configure peripherals, w=power, x=pullups, y=AUX, z=CS
#define SPI_CONFIG_1 0x40
#define  POWER_ON 0x08
#define PULLUP_ON 0x04
#define   USE_AUX 0x02
#define   USE_CS  0x01

//01100xxx – Set SPI speed, 30, 125, 250khz; 1, 2, 2.6, 4, 8MHz
// bitwise 'or' one of these speed values with the SPI_SPEED define
// to set the SPI Speed value
#define SPI_SPEED  0x60
#define SPI_30khz  0x00
#define SPI_125khz 0x01
#define SPI_250khz 0x02
#define SPI_1mhz   0x03
#define SPI_2mhz   0x04
#define SPI_2m6hz  0x05
#define SPI_4mhz   0x06
#define SPI_8mhz   0x07

//1000wxyz – SPI config, w=output type, x=idle, y=clock edge, z=sample
#define SPI_CONFIG_2  0x80
#define SPI_PUSH_PULL 0x08
#define SPI_IDLE_LOW  0x04
#define SPI_CLOCK_ACTIVE_TO_LOW 0x02
#define SPI_SAMPLE_END 0x01

@implementation BusPirate


-(id)initWithPort:(NSString *)path {
    self = [super init];
    
    if (self) {
        // Open the serial port
        port = [[SerialPort alloc] init:path
                              withSpeed:115200
                               andFlags:SERIAL_PORT_N81];
        
        // Enter binary mode
        for (int i = 0; i < 20; i++) {
            [port putByte:0x00];
        }
        
        // Wait for acknowledgement of Binary mode
        NSString *response = [NSString stringWithUTF8String:[[port getData:5] bytes]];
        
        if ([response caseInsensitiveCompare:@"BBIO1"] != NSOrderedSame) {
            NSLog(@"Confusing response from Bus Pirate: %@", response);
            [self release];
            self = nil;
            return self;
        }
        
        // Enter SPI mode
        [port putByte:0x01];
        
        response = [NSString stringWithUTF8String:[[port getData:4] bytes]];
        if ([response caseInsensitiveCompare:@"SPI1"] != NSOrderedSame) {
            NSLog(@"SPI version mismatch, may cause problems.");
        }
        
        // Set up the SPI bus
        [port putByte:(SPI_CONFIG_1 | POWER_ON | USE_CS)];
        unsigned char retval = [port getByte];
        if (retval != 0x01) {
            NSLog(@"Unable to configure (1) SPI mode: %d", retval);
        }
        
        [port putByte:(SPI_SPEED | 4)];
        retval = [port getByte];
        if (retval != 0x01) {
            NSLog(@"Unable to set SPI speed: %d", retval);
        }
        
        [port putByte:(SPI_CONFIG_2 | SPI_PUSH_PULL | SPI_IDLE_LOW)];
        retval = [port getByte];
        if (retval != 0x01) {
            NSLog(@"Unable to configure (2) SPI mode: %d", retval);
        }
    }
    
    return self;
}

-(void)sendSPIwithPin:(NSInteger)dataPin
                Clock:(NSInteger)clockPin
                   CS:(NSInteger)CSPin
                 data:(long long)data
                 bits:(NSInteger)numBits
{
    // SPI transfers are made up of several spi transfers to load into the
    // shift registers.
    
//    __|-|_|-|_|-|_|-|_|-|_|-|_|-|__ Clock
//    ___|---|___|---|___|---|___|--- Data (10101010...)
//    -|___________________________|- Chip select
//     **** *** *** *** *** *** ****  SPI transfers to the shift registers

    // For every bit of data, there are 3 SPI transfers
    // Allocate data for each of the states
    
    // Generate the command stream
    for (int i = 0; i < numBits; i++) {
        long long mask = (1 << (numBits - i - 1));
        // Raise the clock pin
        [port putByte:CS_LOW];
        [port putByte:BULK_SPI];

        // If this isn't the first data byte, keep the state from the last.
        if (i > 0) {
            // If the last byte was '1' bitwise or that with the clock
            if (data & (mask << 1)) {
                [port putByte:(1 << clockPin) | (1 << dataPin)];
            } else {
                [port putByte:(1 << clockPin)];                
            }
        } else {
            [port putByte:(1 << clockPin)];
        }
        [port putByte:CS_HIGH];
        [port getByte];

        // Set the data pin
        // This will be true if the desired bit is high
        if (data & mask) {
            [port putByte:CS_LOW];
            [port putByte:BULK_SPI];
            [port putByte:(1 << dataPin) | (1 << clockPin)];
            [port putByte:CS_HIGH];
            [port getByte];
            
            // Lower the clock pin
            [port putByte:CS_LOW];
            [port putByte:BULK_SPI];
            [port putByte:(1 << dataPin)];
            [port putByte:CS_HIGH];
            [port getByte];

        } else {
            [port putByte:CS_LOW];
            [port putByte:BULK_SPI];
            [port putByte:(1 << clockPin)];
            [port putByte:CS_HIGH];
            [port getByte];
            
            // Lower the clock pin
            [port putByte:CS_LOW];
            [port putByte:BULK_SPI];
            [port putByte:0x00];
            [port putByte:CS_HIGH];
            [port getByte];
        }        
    }
    
    // Pulse the CS pin high
    [port putByte:CS_LOW];
    [port putByte:BULK_SPI];
    [port putByte:(1 << CSPin)];
    [port putByte:CS_HIGH];
    [port getByte];
    
    // Delay
    usleep(100);
    
    // Bring the CS back down
    [port putByte:CS_LOW];
    [port putByte:BULK_SPI];
    [port putByte:0x00];
    [port putByte:CS_HIGH];
    [port getByte];    
}

@end
