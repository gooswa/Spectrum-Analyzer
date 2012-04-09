//
//  Document.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 2/6/12.
//  Copyright (c) 2012. All rights reserved.
//

#import <Cocoa/Cocoa.h>
@class SpectrumController;

@interface Document : NSDocument

@property (assign) IBOutlet SpectrumController *controller;


@end
