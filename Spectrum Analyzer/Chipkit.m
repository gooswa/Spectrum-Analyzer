//
//  Chipkit.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/21/12.
//  Copyright (c) 2012 Oregon State University (COAS). All rights reserved.
//

#import "Chipkit.h"

@implementation Chipkit

-(id)initWithPort:(NSString *)path {
    self = [super init];
    
    if (self) {
        // Open the serial port
        port = [[SerialPort alloc] init:path
                              withSpeed:9600
                               andFlags:SERIAL_PORT_N81];
    }
    
    return self;
}

void hexToString(char *string, int len, uint64 hex) {
    // Buffer for hex characters (2 per byte)
    // this is in reversed order
    char *buffer = (char *)malloc(len * 2);
    
    // Calculate each bytes hex value
    char value;
    for (int i = 0; i < len; i++) {
        value = hex & 0x0F;
        hex = hex >> 4;
        if (value < 10) {
            buffer[i*2+0] = value + '0';            
        } else {
            buffer[i*2+0] = value - 10 + 'A';
        }
        value = hex & 0x0F;
        hex = hex >> 4;
        if (value < 10) {
            buffer[i*2+1] = value + '0';            
        } else {
            buffer[i*2+1] = value - 10 + 'A';
        }
    }
    
    // Copy the results into the provided string
    for (int i = 0; i < len*2; i++) {
        string[i] = buffer[(len*2)-i-1];
    }
    
    string[len*2] = '\0';
    free(buffer);
    
    return;
}

-(void)sendSPIwithPin:(NSInteger)dataPin
                Clock:(NSInteger)clockPin
                   CS:(NSInteger)CSPin
                 data:(long long)data
                 bits:(NSInteger)numBits
{
    // Calculate the number of characters we need to send the command
    // format: $s,S,D,C,B,xxxx....
    int bytes = ((int)numBits / 8);
    if (numBits % 8) bytes += 1;

    char buffer[256];
    
    NSLog(@"Number of bytes needed to store %ld bits: %d.", numBits, bytes);
    
    // This stores the converted hex representation of the data.  Includes a null
    char *hexBuffer = (char *)malloc(bytes * 2 + 1);
    
    // Mask out the data value (to make sure that we get the correct
    // result from printf.  I think all that's required is 0xFF for each byte
    long long mask = 0x00;
    for (int i = 0; i < bytes; i++) {
        mask  = mask << 8;
        mask |= 0xFF; 
    }
    
    data &= mask;
    hexToString(hexBuffer, bytes, data);
    
    snprintf(buffer, 256, "$s,%ld,%ld,%ld,%d,%s\n", 
             CSPin, dataPin, clockPin, bytes, hexBuffer);
    
    NSLog(@"Writing \"%s\" to the analyzer.", buffer);
    
    // Send to the device
    [port putString:[NSString stringWithCString:buffer encoding:NSUTF8StringEncoding]];
}

-(void)setPin:(NSInteger)pin Value:(bool)value
{
    char buffer[255];

    if (value) {
        snprintf(buffer, 255, "$p,%ld,1\n", pin);
    } else {
        snprintf(buffer, 255, "$p,%ld,0\n", pin);        
    }
    
    NSLog(@"Writing \"%s\" to the analyzer.", buffer);
    [port putString:[NSString stringWithCString:buffer encoding:NSUTF8StringEncoding]];
    
}

@end
