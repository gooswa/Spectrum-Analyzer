//
//  AppDelegate.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/18/12.
//  Copyright (c) 2012. All rights reserved.
//

#import <Cocoa/Cocoa.h>
#import "SpectrumAnalyzer.h"

@interface AppDelegate : NSObject <NSApplicationDelegate,
                                   NSTextFieldDelegate,
                                   NSComboBoxDelegate,
                                   SpectrumAnalyerDelegate>
{
    IBOutlet NSTextField *DDS1_freq;
    IBOutlet NSTextField *DDS3_freq;
    IBOutlet NSTextField *PLO1_oFreq;
    IBOutlet NSTextField *PLO2_oFreq;
    IBOutlet NSTextField *PLO3_oFreq;
    IBOutlet NSTextField *PLO1_pFreq;
    IBOutlet NSTextField *PLO2_pFreq;
    IBOutlet NSTextField *PLO3_pFreq;
    IBOutlet NSButton    *PLO1_invert;
    IBOutlet NSButton    *PLO2_invert;
    IBOutlet NSButton    *PLO3_invert;
    IBOutlet NSComboBox  *PLO1_FoLD;
    IBOutlet NSComboBox  *PLO2_FoLD;
    IBOutlet NSComboBox  *PLO3_FoLD;
    
    IBOutlet NSTextField *ADC_phase;
    IBOutlet NSTextField *ADC_mag;
    IBOutlet NSTextField *ADC_count;
    IBOutlet NSTextField *ADC_delay;

    IBOutlet NSTextField *DDS_start;
    IBOutlet NSTextField *DDS_stop;
    IBOutlet NSTextField *DDS_steps;
    IBOutlet NSTextField *DDS_delay;

    IBOutlet NSTextField *Sweep_start;
    IBOutlet NSTextField *Sweep_stop;
    IBOutlet NSTextField *Sweep_steps;
    IBOutlet NSTextField *Sweep_delay;

    @private
    HardwareInterface *interface;
    SpectrumAnalyzer *analyzer;
}

@property (assign) IBOutlet NSWindow *window;

- (IBAction)ADC_go:(id)sender;
- (IBAction)DDS1_Scan:(id)sender;
- (IBAction)Sweep_go:(id)sender;

- (IBAction)invertChanged:(id)sender;

@end
