//
//  SpectrumView.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 2/7/12.
//  Copyright (c) 2012. All rights reserved.
//

#import <Cocoa/Cocoa.h>

@class Document;
@class SpectrumController;

@interface SpectrumView : NSView <NSTextFieldDelegate>
{
    IBOutlet SpectrumController *controller;
    IBOutlet Document *document;
    
    @private
    
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

// This provides access to the number of native device pixels
// in the interior of the graph.  It can be used to sample the
// experiment at the perfect resolution
- (NSSize)nativePixelsInGraph;

- (void)updateTextFields;

@end
