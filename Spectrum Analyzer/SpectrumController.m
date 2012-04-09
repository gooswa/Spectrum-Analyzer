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
#import "SpectrumView.h"

@implementation SpectrumController

// Accessors
@synthesize rbwFilterIndex;
@synthesize videoFilterIndex;

@synthesize scanSteps;
@synthesize scanStepDelay;

@synthesize centerSpanMode;
@synthesize startFreq;
@synthesize stopFreq;

@synthesize vAutoScale;
@synthesize vTopLevel;
@synthesize vDevisions;

@synthesize tgOffset;
@synthesize tgInvert;

@synthesize analyzer;

// UI accessors
@synthesize spectrumView;

@synthesize spanTabView;
@synthesize rangeTabView;

@synthesize stepTextField;
@synthesize delayTextField;

@synthesize autoStepButton;
@synthesize scanSpanRangeSelector;
@synthesize scanCenterTextField;
@synthesize scanFDivTextField;
@synthesize scanStartTextField;
@synthesize scanEndTextField;

@synthesize rbwFilterComboBox;
@synthesize videoFilterComboBox;

@synthesize autoScaleButton;
@synthesize topLevelTextField;
@synthesize devisionsTextField;

@synthesize offsetTextField;
@synthesize tgModeRadioBox;

- (id)init
{
    self = [super init];
    if (self) {
        
        // Insert code here to initialize your application
        interface = [[Chipkit alloc] initWithPort:@"/dev/tty.usbserial-A6009B2C"];
        analyzer = [[SpectrumAnalyzer alloc] initWithInterface:interface];
        [analyzer setDelegate:self];

        // Subscribe to notifications of the window resizing
        NSNotificationCenter *center;
        center = [NSNotificationCenter defaultCenter];
        [center addObserver:self selector:@selector(windowResized)
                       name:NSWindowDidResizeNotification
                     object:nil];
        
        // Setup defaults (this should be improved)
        rbwFilterIndex   = 0;
        videoFilterIndex = 0;
        scanStepDelay    = 15;
        
        centerSpanMode = YES;
        startFreq = -0.3;
        stopFreq  =  0.3;
        
        vAutoScale = YES;
        vTopLevel  = 0.;
        vDevisions = 12.;

        tgOffset = 0.;
        tgInvert = NO;
        
        [self configurationChanged];
    }
    
    return self;
}

- (void)didLoadNib
{
    [self updateAutoSteps];

    // Setup the text in each field based upon the defaults
    // The easiest way to do this is probably to run through the
    // UI objects and fake like they were just edited
    // Because they start blank, it'll revert to the default
    [self textShouldEndEditing:stepTextField];
    [self textShouldEndEditing:delayTextField];
    
    // Behave differently whether the span or range tab is selected
    if ([spanTabView tabState] == NSSelectedTab) {
        [self textShouldEndEditing:scanCenterTextField];
        [self textShouldEndEditing:scanFDivTextField];
    } else {
        [self textShouldEndEditing:scanStartTextField];
        [self textShouldEndEditing:scanEndTextField];
    }
    
    [self textShouldEndEditing:topLevelTextField];
    [self textShouldEndEditing:devisionsTextField];
    [self textShouldEndEditing:offsetTextField];
    
    [analyzer startScanning];
}

// This method is called whenever the window is resized
- (void)windowResized
{
    // If the auto setting is turned on for setting the number of scan steps
    // then, once the window is resized, we need to get the width (in points)
    // from the internal part of the spectrum view.  This way we have a perfect
    // match between the frame width and the samples.
    // If the scan steps are set to auto, read the width of the spectrum view
    if ([autoStepButton state] == NSOnState) {
        NSSize size = [spectrumView nativePixelsInGraph];
        scanSteps = size.width;
        [stepTextField setStringValue:[NSString stringWithFormat:@"%d", scanSteps]];
        [self textShouldEndEditing:stepTextField];
    }
    
    // The only exception to this is when the RBW filter is too narrow to
    // cover the steps.  In this case, the number of steps should be the minimum
    // number required to cover the span with the 3dB points of the filter
    // overlapping
    
    return;
}

- (void)configurationChanged
{
    // Not sure what this should do for now...
}

- (void)sampleCollected:(AnalyzerSample_t)sample atStep:(int)step
{
    // This will get called every time a new sample is collected during scanning
    
    [spectrumView setNeedsDisplay:YES];
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
    double doubleValue = [(NSTextField*)sender doubleValue];
    float intValue = [(NSTextField*)sender intValue];
    
    if (sender == stepTextField) {
        if (intValue != [analyzer steps]) {
            [analyzer setSteps:intValue];
            [spectrumView setNeedsDisplay:YES];
        }
        
        [spectrumView updateTextFields];
        return YES;
    }
    
    if (sender == delayTextField) {
        scanStepDelay = intValue;
        [analyzer setDelayMs:scanStepDelay];
        [delayTextField setStringValue:[NSString stringWithFormat:@"%d mS", scanStepDelay]];
        [spectrumView updateTextFields];
        return YES;
    }
    
    if (sender == topLevelTextField) {
        vTopLevel = doubleValue;
        [spectrumView setNeedsDisplay:YES];
        [spectrumView updateTextFields];
        return YES;
    }
    
    if (sender == devisionsTextField) {
        vDevisions = intValue;
        [spectrumView setNeedsDisplay:YES];
        [spectrumView updateTextFields];
        return YES;
    }    
    
    if (sender == scanFDivTextField) {
        startFreq = [scanCenterTextField doubleValue] - doubleValue * 5.;
        stopFreq  = [scanCenterTextField doubleValue] + doubleValue * 5.;
        [analyzer setStartFreq:startFreq];
        [analyzer  setStopFreq:stopFreq];
        
        [self updateAutoSteps];
        
        [spectrumView setNeedsDisplay:YES];
        [spectrumView updateTextFields];
        return YES;
    }
    
    if (sender == scanCenterTextField) {
        startFreq = doubleValue - [scanFDivTextField doubleValue] * .5;
        stopFreq  = doubleValue + [scanFDivTextField doubleValue] * .5;
        [analyzer setStartFreq:startFreq];
        [analyzer  setStopFreq:stopFreq];
        [spectrumView setNeedsDisplay:YES];
        [spectrumView updateTextFields];
        return YES;
    }

    if (sender == offsetTextField) {
        
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
    if ([sender state] == NSOnState) {
        [stepTextField setEditable:NO];
        [self windowResized];
    } else {
        [stepTextField setEditable:YES];
    }
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

- (void)updateAutoSteps
{
    // If the scan steps are set to auto, read the width of the spectrum view
    if ([autoStepButton state] == NSOnState) {
        NSSize size = [spectrumView nativePixelsInGraph];

        int pixelSteps = size.width;

        int resolutionSteps = (stopFreq - startFreq) / ([analyzer RBW] / 2.);
        
        scanSteps = MAX(pixelSteps, resolutionSteps);
        
        [stepTextField setStringValue:[NSString stringWithFormat:@"%d", scanSteps]];
        [self textShouldEndEditing:stepTextField];
    }
    
    
}

@end
