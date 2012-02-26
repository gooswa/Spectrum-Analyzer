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

@interface SpectrumView : NSView
{
    IBOutlet SpectrumController *controller;
    IBOutlet Document *document;
}

@property(assign) Document *document;

@end
