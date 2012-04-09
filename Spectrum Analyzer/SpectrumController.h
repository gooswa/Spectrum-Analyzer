//
//  SpectrumController.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 2/23/12.
//  Copyright (c) 2012 Oregon State University (COAS). All rights reserved.
//

#import <Foundation/Foundation.h>
#import "SpectrumAnalyzer.h"

@class SpectrumView;
@class SpectrumAnalyzer;
@class HardwareInterface;

@interface SpectrumController : NSObject <NSTextFieldDelegate,
                                          NSComboBoxDelegate,
                                          SpectrumAnalyerDelegate>
{
    HardwareInterface *interface;
    SpectrumAnalyzer  *analyzer;
    
    // Scan
    int   rbwFilterIndex;
    int   videoFilterIndex;
    int   scanSteps;
    int   scanStepDelay;

    // These numbers are derived from the center/span setting
    // if we're using that mode
    BOOL  centerSpanMode;
    double startFreq;
    double stopFreq;
    
    // Vertical scale
    BOOL  vAutoScale;
    float vTopLevel;
    int   vDevisions;
    
    // Tracking Generator
    float tgOffset;
    BOOL  tgInvert;
}

- (void)didLoadNib;

// Accessors
@property (readonly)  int   rbwFilterIndex;
@property (readonly)  int   videoFilterIndex;
@property (readonly)  int   scanSteps;
@property (readonly)  int   scanStepDelay;

@property (readonly)  BOOL  centerSpanMode;
@property (readwrite) double startFreq;
@property (readwrite) double stopFreq;

@property (readonly)  BOOL  vAutoScale;
@property (readwrite) float vTopLevel;
@property (readonly)  int   vDevisions;

@property (readonly)  float tgOffset;
@property (readonly)  BOOL  tgInvert;

@property (readonly) SpectrumAnalyzer *analyzer;

// UI outlets
// Spectrum analyzer display window
@property (assign) IBOutlet SpectrumView *spectrumView;
@property (assign) IBOutlet NSTabViewItem *spanTabView;
@property (assign) IBOutlet NSTabViewItem *rangeTabView;

// Scan parameters
@property (assign) IBOutlet NSButton    *autoStepButton;
@property (assign) IBOutlet NSComboBox  *rbwFilterComboBox;
@property (assign) IBOutlet NSComboBox  *videoFilterComboBox;
@property (assign) IBOutlet NSTextField *stepTextField;
@property (assign) IBOutlet NSTextField *delayTextField;
@property (assign) IBOutlet NSTabView   *scanSpanRangeSelector;
@property (assign) IBOutlet NSTextField *scanCenterTextField;
@property (assign) IBOutlet NSTextField *scanFDivTextField;
@property (assign) IBOutlet NSTextField *scanStartTextField;
@property (assign) IBOutlet NSTextField *scanEndTextField;

// Vertical scale stuff
@property (assign) IBOutlet NSButton    *autoScaleButton;
@property (assign) IBOutlet NSTextField *topLevelTextField;
@property (assign) IBOutlet NSTextField *devisionsTextField;

// Tracking generator stuff
@property (assign) IBOutlet NSTextField *offsetTextField;
@property (assign) IBOutlet NSMatrix    *tgModeRadioBox;

- (bool)textShouldEndEditing:(id)sender;

// Button actions
- (IBAction)rbwFilterAutoChange:(id)sender;
- (IBAction)stepAutoChange:(id)sender;
- (IBAction)delayAutoChange:(id)sender;
- (IBAction)videoFilterAutoChange:(id)sender;
- (IBAction)tgModeChange:(id)sender;
- (IBAction)vAutoScaleChange:(id)sender;

- (void)updateAutoSteps;

@end
