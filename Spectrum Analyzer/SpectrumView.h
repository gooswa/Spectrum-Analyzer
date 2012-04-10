//
//  SpectrumView.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 2/7/12.
//  Copyright (c) 2012. All rights reserved.
//

#import <Cocoa/Cocoa.h>
#import "SpectrumAnalyzer.h"

@class Document;
@class SpectrumController;

@interface SpectrumView : NSView <NSTextFieldDelegate>
{
    IBOutlet SpectrumController *controller;
    IBOutlet Document *document;
    
    @private
    
    // Cached value
    NSSize nativePixelsInGraph;
    
    // Text fields 
    NSTextField *titleTextField;

    // Text fields for the frequency span
    NSTextField *leftTextField;
    NSTextField *rightTextField;
    NSTextField *centerTextField;
    NSTextField *mhzPerDivTextField;
    
    // Top magnitude field
    NSTextField *topMagTextField;
}

@property(assign) Document *document;
@property(readonly) NSSize nativePixelsInGraph;

// This provides access to the number of native device pixels
// in the interior of the graph.  It can be used to sample the
// experiment at the perfect resolution

- (void)sampleCollected:(AnalyzerSample_t)sample atStep:(int)step;

- (void)updateTextFields;

@end
