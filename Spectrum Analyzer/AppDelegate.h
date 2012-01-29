//
//  AppDelegate.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/18/12.
//  Copyright (c) 2012. All rights reserved.
//

#import <Cocoa/Cocoa.h>
//#import "BusPirate.h"
#import "Chipkit.h"
#import "LMX_PLL.h"
#import "AD_DDS.h"

@interface AppDelegate : NSObject <NSApplicationDelegate, NSTextFieldDelegate>
{
//    BusPirate *interface;
    Chipkit *interface;

    LMX_PLL *PLO1, *PLO2, *PLO3;
    AD_DDS  *DDS1, *DDS3;
    
    IBOutlet NSTextField *DDS1_freq;
    IBOutlet NSTextField *DDS3_freq;
    IBOutlet NSTextField *PLO1_oFreq;
    IBOutlet NSTextField *PLO2_oFreq;
    IBOutlet NSTextField *PLO3_oFreq;
    IBOutlet NSTextField *PLO1_pFreq;
    IBOutlet NSTextField *PLO2_pFreq;
    IBOutlet NSTextField *PLO3_pFreq;
}

@property (assign) IBOutlet NSWindow *window;

@end
