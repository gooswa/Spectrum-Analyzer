//
//  SpectrumController.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 2/23/12.
//  Copyright (c) 2012 Oregon State University (COAS). All rights reserved.
//

#import <Foundation/Foundation.h>

@class SpectrumView;

@interface SpectrumController : NSObject

@property (assign) IBOutlet SpectrumView *spectrumView;

// Scan parameters
@property (assign) IBOutlet NSComboBox *rbwFilterComboBox;
@property (assign) IBOutlet NSTextField *stepTextField;
@property (assign) IBOutlet NSTextField *delayTextField;
@property (assign) IBOutlet NSComboBox *videoFilterComboBox;

// Vertical scale stuff
@property (assign) IBOutlet NSButton *autoScaleButton;
@property (assign) IBOutlet NSTextField *topLevelTextField;
@property (assign) IBOutlet NSTextField *devisionsTextField;

// Tracking generator stuff
@property (assign) IBOutlet NSTextField *offsetTextField;
@property (assign) IBOutlet NSMatrix *tgModeRadioBox;

@end
