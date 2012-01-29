//
//  AppDelegate.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/18/12.
//  Copyright (c) 2012. All rights reserved.
//

#import "AppDelegate.h"

#define REF_CLOCK 64

@implementation AppDelegate

@synthesize window = _window;

- (void)dealloc
{
    [super dealloc];
}

- (void)applicationDidFinishLaunching:(NSNotification *)aNotification
{
    // Insert code here to initialize your application
    interface = [[Chipkit alloc] initWithPort:@"/dev/tty.usbserial-A6009B2C"];
    
    // Set this class to be the delegate for all the text fields
    [PLO1_oFreq setDelegate:self];
    [PLO1_pFreq setDelegate:self];
    [PLO2_oFreq setDelegate:self];
    [PLO2_pFreq setDelegate:self];
    [PLO3_oFreq setDelegate:self];
    [PLO3_pFreq setDelegate:self];
    [DDS1_freq  setDelegate:self];
    [DDS3_freq  setDelegate:self];
    
    sleep(5);
    
    // Setup PLO2
    PLO2 = [[LMX_PLL alloc] init];
    PLO2.CS    = 4;
    PLO2.Data  = 3;
    PLO2.Clock = 2;
    PLO2.FoLD_output = FoLD_RDIV;
    PLO2.refFreq = 64.0;
    PLO2.phaseFreq = 4.0;
    PLO2.outputFreq = 1024.0;
    [PLO2 updateHardware:interface];
    [PLO2_pFreq setDoubleValue:[PLO2 phaseFreq]];
    [PLO2_oFreq setDoubleValue:[PLO2 outputFreq]];
    
    // Setup DDS1
    DDS1 = [[AD_DDS alloc] init];
    DDS1.CS    = 6;
    DDS1.Data  = 7;
    DDS1.Clock = 5;
    DDS1.refFreq = 64.0;
    DDS1.outputFreq = 10.7;
    [DDS1 initHardware:interface];
    [DDS1 updateHardware:interface];
    [DDS1_freq setDoubleValue:[DDS1 outputFreq]];
    
    // Setup PLO1
    PLO1 = [[LMX_PLL alloc] init];
    PLO1.CS    = 12;
    PLO1.Data  = 11;
    PLO1.Clock = 10;
    PLO1.FoLD_output = FoLD_RDIV;
    PLO1.refFreq = 10.7;
    PLO1.phaseFreq = 0.972;
    PLO1.outputFreq = 1450.;
    [PLO1 updateHardware:interface];
    [PLO1_pFreq setDoubleValue:[PLO1 phaseFreq]];
    [PLO1_oFreq setDoubleValue:[PLO1 outputFreq]];

    // Setup DDS3
    DDS3 = [[AD_DDS alloc] init];
    DDS3.CS    = 14;
    DDS3.Data  = 15;
    DDS3.Clock = 13;
    DDS3.refFreq = 64.0;
    DDS3.outputFreq = 10.7;
    [DDS3 initHardware:interface];
    [DDS3 updateHardware:interface];
    [DDS3_freq setDoubleValue:[DDS3 outputFreq]];

    // Setup PLO3
    PLO3 = [[LMX_PLL alloc] init];
    PLO3.CS    = 19;
    PLO3.Data  = 18;
    PLO3.Clock = 17;
    PLO3.FoLD_output = FoLD_RDIV;
    PLO3.refFreq = 10.7;
    PLO3.phaseFreq = 0.972;
    PLO3.outputFreq = 1450;
    [PLO3 updateHardware:interface];
    [PLO3_pFreq setDoubleValue:[PLO3 phaseFreq]];
    [PLO3_oFreq setDoubleValue:[PLO3 outputFreq]];
}

// 
- (bool)textShouldEndEditing:(id)sender
{
    double value = [sender doubleValue];
    
    // If this a DDS text field, the number can't be more than 1/3 of reference
    if (sender == DDS1_freq) {
        
        // Invalid value, restore old value
        if (value > ([DDS1 refFreq] * (1./3.))) {
            [sender setDoubleValue:[DDS1 outputFreq]];
            return NO;
        } else {
            [DDS1 setOutputFreq:value];
            [DDS1 updateHardware:interface];
            // DDS1 is the reference clock for PLO1
            [PLO1 setRefFreq:[DDS1 outputFreq]];
            [PLO1_oFreq setDoubleValue:[PLO1 outputFreq]];
            [PLO1_pFreq setDoubleValue:[PLO1 phaseFreq]];
            [PLO1 updateHardware:interface];
        }
    }
    
    if (sender == DDS3_freq) {
        
        if (value > ([DDS3 refFreq] * (1./3.))) {
            [sender setDoubleValue:[DDS3 outputFreq]];
            return NO;
        } else {
            [DDS3 setOutputFreq:value];
            [DDS3 updateHardware:interface];
            [PLO3 setRefFreq:[DDS3 outputFreq]];
            [PLO3_oFreq setDoubleValue:[PLO3 outputFreq]];
            [PLO3_pFreq setDoubleValue:[PLO3 phaseFreq]];
            [PLO3 updateHardware:interface];
        }
    }
    
    if (sender == PLO2_oFreq) {
        [PLO2 setOutputFreq:value];
        [PLO2_oFreq setDoubleValue:[PLO2 outputFreq]];
        [PLO2 updateHardware:interface];
    }
    
    if (sender == PLO2_pFreq) {
        [PLO2 setPhaseFreq:value];
        [PLO2_oFreq setDoubleValue:[PLO2 outputFreq]];
        [PLO2_pFreq setDoubleValue:[PLO2 phaseFreq]];
        [PLO2 updateHardware:interface];
    }    

    return NO;
}

- (void)controlTextDidEndEditing:(NSNotification *)obj
{
    id sender = [obj object];
    
    [self textShouldEndEditing:sender];
}

@end
