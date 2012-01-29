//
//  Chipkit.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/21/12.
//  Copyright (c) 2012. All rights reserved.
//

#import <Foundation/Foundation.h>
#import <SerialPort/SerialPort.h>
#import "HardwareInterface.h"

@interface Chipkit : HardwareInterface
{
    SerialPort *port;
}

@end
