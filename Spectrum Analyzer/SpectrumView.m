//
//  SpectrumView.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 2/7/12.
//  Copyright (c) 2012. All rights reserved.
//

#import "SpectrumView.h"

@implementation SpectrumView

@synthesize document;

- (id)initWithFrame:(NSRect)frame
{
    self = [super initWithFrame:frame];
    if (self) {
        // Initialization code here.
    }
    
    return self;
}

- (void)drawRect:(NSRect)dirtyRect
{
    [[NSColor blackColor] set];

    float borderWidth = 0;
    float halfBorderWidth = 0.5 * borderWidth;
    NSRect borderRect = NSInsetRect([self bounds],
                                    halfBorderWidth,
                                    halfBorderWidth);

    NSBezierPath *borderPath = [NSBezierPath bezierPathWithRect:borderRect];
    [borderPath fill];
}

@end
