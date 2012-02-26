//
//  SpectrumController.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 2/23/12.
//  Copyright (c) 2012 Oregon State University (COAS). All rights reserved.
//

#import "SpectrumController.h"
#import "SpectrumAnalyzer.h"
#import "HardwareInterface.h"

@implementation SpectrumController

// Accessors
@synthesize rbwFilterIndex;
@synthesize videoFilterIndex;
@synthesize scanSteps;
@synthesize scanStepDelay;
@synthesize vAutoScale;
@synthesize vTopLevel;
@synthesize vDivisions;
@synthesize tgOffset;
@synthesize tgInvert;

// UI accessors
@synthesize spectrumView;
@synthesize rbwFilterComboBox;
@synthesize stepTextField;
@synthesize delayTextField;
@synthesize videoFilterComboBox;
@synthesize autoScaleButton;
@synthesize topLevelTextField;
@synthesize devisionsTextField;
@synthesize offsetTextField;
@synthesize tgModeRadioBox;
@synthesize rbwFilterAutoChange;
@synthesize stepAutoChange;
@synthesize delayAutoChange;
@synthesize videoFilterAutoChange;
@synthesize tgModeChange;
@synthesize vAutoScaleChange;
@synthesize divisionsTextField;

- (id)init
{
    self = [super init];
    if (self) {
        
        // Subscribe to notifications of the window resizing
        NSNotificationCenter *center;
        center = [NSNotificationCenter defaultCenter];
        [center addObserver:self selector:@selector(windowResized)
                       name:NSWindowDidResizeNotification
                     object:nil];
        
    }
}

// This method is called whenever the window is resized
- (void)windowResized
{
    // If the auto setting is turned on for setting the number of scan steps
    // then, once the window is resized, we need to get the width (in points)
    // from the internal part of the spectrum view.  This way we have a perfect
    // match between the frame width and the samples.
    
    
    // The only exception to this is when the RBW filter is too narrow to
    // cover the steps.  In this case, the number of steps should be the minimum
    // number required to cover the span with the 3dB points of the filter
    // overlapping
    
    return;
}
         
// This is for either of the combo boxes (filters, for now)
- (void)comboBoxSelectionDidChange:(NSNotification *)notification
{
    NSComboBox *sender = (NSComboBox *)[notification object];
    
    if (sender == videoFilterComboBox) {
    }
    
    if (sender == rbwFilterComboBox) {
    }    
}

// Modify hardware state for user initiated change
// If the change doesn't make sense, return NO to revert
- (bool)textShouldEndEditing:(id)sender
{
    if (sender == stepTextField) {        
    }
    
    if (sender == delayTextField) {
    }
    
    if (sender == topLevelTextField) {
    }
    
    if (sender == devisionsTextField) {
    }    
    
    if (sender == offsetTextField) {
    }
    
    if (sender == divisionsTextField) {
    }    
    
    return NO;
}

- (void)controlTextDidEndEditing:(NSNotification *)obj
{
    id sender = [obj object];
    
    [self textShouldEndEditing:sender];
}


- (IBAction)rbwFilterAutoChange:(id)sender
{

}

- (IBAction)stepAutoChange:(id)sender
{

}

- (IBAction)delayAutoChange:(id)sender
{

}

- (IBAction)videoFilterAutoChange:(id)sender
{

}

- (IBAction)tgModeChange:(id)sender
{

}

- (IBAction)vAutoScaleChange:(id)sender
{

}

@end
