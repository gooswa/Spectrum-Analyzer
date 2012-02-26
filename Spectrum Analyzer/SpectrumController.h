//
//  SpectrumController.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 2/23/12.
//  Copyright (c) 2012 Oregon State University (COAS). All rights reserved.
//

#import <Foundation/Foundation.h>

@class SpectrumView;
@class SpectrumAnalyzer;
@class HardwareInterface;

@interface SpectrumController : NSObject <NSTextFieldDelegate, NSComboBoxDelegate>
{
    HardwareInterface *interface;
    SpectrumAnalyzer  *analyzer;
    
    // Scan
    int   rbwFilterIndex;
    int   videoFilterIndex;
    int   scanSteps;
    int   scanStepDelay;
    
    // Vertical scale
    BOOL  vAutoScale;
    float vTopLevel;
    int   vDivisions;
    
    // Tracking Generator
    float tgOffset;
    BOOL  tgInvert;
}

// Accessors
@property (readonly) int   rbwFilterIndex;
@property (readonly) int   videoFilterIndex;
@property (readonly) int   scanSteps;
@property (readonly) int   scanStepDelay;

@property (readonly) BOOL  vAutoScale;
@property (readonly) float vTopLevel;
@property (readonly) int   vDivisions;

@property (readonly) float tgOffset;
@property (readonly) BOOL  tgInvert;

// UI outlets
// Spectrum analyzer display window
@property (assign) IBOutlet SpectrumView *spectrumView;

// Scan parameters
@property (assign) IBOutlet NSComboBox  *rbwFilterComboBox;
@property (assign) IBOutlet NSComboBox  *videoFilterComboBox;
@property (assign) IBOutlet NSTextField *stepTextField;
@property (assign) IBOutlet NSTextField *delayTextField;

// Vertical scale stuff
@property (assign) IBOutlet NSButton    *autoScaleButton;
@property (assign) IBOutlet NSTextField *topLevelTextField;
@property (assign) IBOutlet NSTextField *devisionsTextField;

// Tracking generator stuff
@property (assign) IBOutlet NSTextField *offsetTextField;
@property (assign) IBOutlet NSMatrix    *tgModeRadioBox;

// Button actions
- (IBAction)rbwFilterAutoChange:(id)sender;
- (IBAction)stepAutoChange:(id)sender;
- (IBAction)delayAutoChange:(id)sender;
- (IBAction)videoFilterAutoChange:(id)sender;
- (IBAction)tgModeChange:(id)sender;
- (IBAction)vAutoScaleChange:(id)sender;

@end
