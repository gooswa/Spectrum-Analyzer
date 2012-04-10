//
//  SpectrumView.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 2/7/12.
//  Copyright (c) 2012. All rights reserved.
//

#import "SpectrumView.h"
#import "SpectrumController.h"
#import "SpectrumAnalyzer.h"

@implementation SpectrumView

@synthesize document;

// This method determines the location of the point in screen space
// and rounds the value to yield pixel-aligned lines, which are sharp.
// The size field is only used to compute a half-pixel offset in the case
// of objects that are an odd number of pixels in size, so they fit perfectly.
- (NSPoint)pixelAlignPoint:(NSPoint)point withSize:(NSSize)size
{
    NSSize sizeInPixels = [self convertSizeToBase:size];
    CGFloat halfWidthInPixels  = sizeInPixels.width * 0.5;
    CGFloat halfHeightInPixels = sizeInPixels.height * 0.5;
    
    // Is the width an odd number of pixels?
    NSPoint adjustmentInPixels = NSMakePoint(0., 0.);
    if (fabs(halfWidthInPixels - floor(halfWidthInPixels)) > 0.0001 ) {
        adjustmentInPixels.x = 0.5;
    } else {
        adjustmentInPixels.x = 0.;
    }
    
    // Is the height an odd number of pixels?
    if (fabs(halfHeightInPixels - floor(halfHeightInPixels)) > 0.0001 ) {
        adjustmentInPixels.y = 0.5;
    } else {
        adjustmentInPixels.y = 0.;
    }
    
    // This is the adjustment needed for odd or even sizes
//    NSPoint adjustment = [self convertPointFromBase:adjustmentInPixels];
    
    NSPoint basePoint = [self convertPointToBase:point];
    basePoint.x = round(basePoint.x) + adjustmentInPixels.x;
    basePoint.y = round(basePoint.y) + adjustmentInPixels.y;
    
    return [self convertPointFromBase:basePoint];
}

- (NSTextField *)createTextFieldInRect:(NSRect)frame
{
    NSTextField *field;
    
    field = [[NSTextField alloc] initWithFrame:frame];
    [field setBordered:NO];
    [field setDrawsBackground:NO];
    [field setTextColor:[NSColor whiteColor]];
    [field setFont:[NSFont fontWithName:@"Helvetica Neue"
                                   size:12.]];
    [field setBackgroundColor:[NSColor blackColor]];
    [field setDelegate:self];
    
    return field;
}

- (NSSize)nativePixelsInGraph
{
    float borderWidth = 60;
    NSRect borderRect = NSInsetRect([self bounds],
                                    borderWidth,
                                    borderWidth);
    
    // Pixel-align the rect
    borderRect.origin = [self pixelAlignPoint:borderRect.origin
                                     withSize:NSMakeSize(1., 1.)];    

    // Assume that points=pixels for now (not true in HiRes mode)
    return borderRect.size;
}

- (void)viewBoundsChanged
{
    NSRect frame = [self frame];
    NSRect textFrame;
    
    textFrame.origin.x = 60. - 50.;
    textFrame.origin.y = 35.;
    textFrame.size.width  = 100.;
    textFrame.size.height = 25.;
    [leftTextField setFrame:textFrame];

    textFrame.origin.x = frame.size.width / 2. - 50.;
    textFrame.origin.y = 35.;
    [centerTextField setFrame:textFrame];

    textFrame.origin.y = 15.;
    [mhzPerDivTextField setFrame:textFrame];
    
    textFrame.origin.x = frame.size.width - 60. - 50.;
    textFrame.origin.y = 35.;
    [rightTextField setFrame:textFrame];

    textFrame.origin.x = 0.;
    textFrame.origin.y = frame.size.height - 60. - 12.;
    textFrame.size.width  = 55.;
    [topMagTextField setFrame:textFrame];
    
    textFrame.origin.x = 0.;
    textFrame.origin.y = frame.size.height - 25.;
    textFrame.size.width = frame.size.width;
    textFrame.size.height = 25;
    [titleTextField setFrame:textFrame];
    
}

- (id)initWithFrame:(NSRect)frame
{
    self = [super initWithFrame:frame];
    if (self) {

        // Listen to notifications of views frame being resized
        NSNotificationCenter *center;
        center = [NSNotificationCenter defaultCenter];
        [center addObserverForName:NSViewFrameDidChangeNotification
                            object:nil
                             queue:[NSOperationQueue mainQueue]
                        usingBlock:
         ^(NSNotification *event) {
             if ([[event object] isEqual:self]) {
                 [self viewBoundsChanged];
             }
         }];
        
        leftTextField = [self createTextFieldInRect:NSMakeRect(0., 0., 0., 0.)];
        [leftTextField setAlignment:NSCenterTextAlignment];
        [leftTextField setEditable:NO];
        [leftTextField setFont:[NSFont fontWithName:@"Helvetica Neue"
                                               size:11.]];
        [leftTextField setDelegate:self];
        [self addSubview:leftTextField];

        centerTextField = [self createTextFieldInRect:NSMakeRect(0., 0., 0., 0.)];
        [centerTextField setAlignment:NSCenterTextAlignment];
        [centerTextField setFont:[NSFont fontWithName:@"Helvetica Neue"
                                                 size:11.]];
        [centerTextField setDelegate:self];
        [self addSubview:centerTextField];

        rightTextField = [self createTextFieldInRect:NSMakeRect(0., 0., 0., 0.)];
        [rightTextField setAlignment:NSCenterTextAlignment];
        [rightTextField setEditable:NO];
        [rightTextField setFont:[NSFont fontWithName:@"Helvetica Neue"
                                                 size:11.]];
        [rightTextField setDelegate:self];
        [self addSubview:rightTextField];
        
        mhzPerDivTextField = [self createTextFieldInRect:NSMakeRect(0., 0., 0., 0.)];
        [mhzPerDivTextField setAlignment:NSCenterTextAlignment];
        [mhzPerDivTextField setFont:[NSFont fontWithName:@"Helvetica Neue"
                                                 size:11.]];
        [mhzPerDivTextField setDelegate:self];
        [self addSubview:mhzPerDivTextField];
        
        topMagTextField = [self createTextFieldInRect:NSMakeRect(0., 0., 0., 0.)]; 
        [topMagTextField setAlignment:NSRightTextAlignment];
        [topMagTextField setFont:[NSFont fontWithName:@"Helvetica Neue"
                                                 size:11.]];
        [topMagTextField setDelegate:self];
        [self addSubview:topMagTextField];

        titleTextField = [self createTextFieldInRect:NSMakeRect(0., 0., 0., 0.)];
        [titleTextField setAlignment:NSCenterTextAlignment];
        [titleTextField setFont:[NSFont fontWithName:@"Helvetica Neue"
                                                size:15.]];
        [titleTextField setDelegate:self];
        [self addSubview:titleTextField];

        [self updateTextFields];
        
        [titleTextField  setStringValue:@"Untitled experiment"];
        
        [self viewBoundsChanged];
        
    }
    
    return self;
}

#pragma mark -
#pragma mark Drawing code

// This method draws the horizontal gridlines.
// Each line is at exactly at 10 dB points.
- (void)drawHorizGridsInRect:(NSRect)rect
{
    float heightPerDiv = rect.size.height / [controller vDevisions];
    NSBezierPath *path = [[NSBezierPath alloc] init];
    [path setLineWidth:1.];
    
    for (int i = 0; i < [controller vDevisions]; i++) {
        NSPoint leftPoint  = NSMakePoint(rect.origin.x,
                                         i * heightPerDiv + rect.origin.y);
        NSPoint rightPoint = NSMakePoint(rect.origin.x + rect.size.width,
                                         i * heightPerDiv + rect.origin.y);
        
        leftPoint  = [self pixelAlignPoint:leftPoint  withSize:NSMakeSize(1., 1.)];
        rightPoint = [self pixelAlignPoint:rightPoint withSize:NSMakeSize(1., 1.)];
        
        [path moveToPoint:leftPoint];
        [path lineToPoint:rightPoint];
    }
    
    [path stroke];
}

- (void)drawVertGridsInRect:(NSRect)rect
{
    [[NSColor grayColor] set];
    NSBezierPath *path = [NSBezierPath bezierPath];

    float deltaPixels = rect.size.width / 10.;
    for (int i = 1; i < 10; i++) {
        // iterate through the horizontal rules
        NSPoint topPoint    = NSMakePoint(i * deltaPixels + rect.origin.x,
                                          rect.origin.y + rect.size.height);
        NSPoint bottomPoint = NSMakePoint(i * deltaPixels + rect.origin.x,
                                          rect.origin.y);
        
        // Make the center line a little lower than the bottom line
        if (i == 5) {
            bottomPoint.y -= 5.;
        }
        
        topPoint    = [self pixelAlignPoint:topPoint    withSize:NSMakeSize(1., 1.)];
        bottomPoint = [self pixelAlignPoint:bottomPoint withSize:NSMakeSize(1., 1.)];
        
        [path moveToPoint:topPoint];
        [path lineToPoint:bottomPoint];
    }
    
    [path stroke];
}

- (void)drawDataInRect:(NSRect)rect
{
    // All this stuff doesn't change with the different "modes"
    int steps = [[controller analyzer] steps];
    AnalyzerSample_t *samples = [[controller analyzer] samples];
    
    float pixelsPerStep = (rect.size.width - 1) / steps;
    float dBRange = [controller vDevisions] * 10.;
    float pixelsPerDb = rect.size.height / dBRange;
    float bottomLevel = [controller vTopLevel] - ([controller vDevisions] * 10.);
    
    [[NSColor yellowColor] set];
    NSBezierPath *path = [[NSBezierPath alloc] init];
    [path setLineWidth:1.];
    
    // If pixels per step is less than one we're in "oversample" mode
    // If it's equal when we're in a 1-to-1 mode
    // If it's more than one we're in undersamples mode
    
    // Over sample mode works by finding the high and low values for each pixel
    // draw a line from the low value to the high at the pixel.  If it works
    // out that there is only one sample for the pixel.  It is both the high
    // and the low value.
    if (fabs(pixelsPerStep - 1.) < FLT_EPSILON ) {
        float min = FLT_MAX;
        float max = FLT_MIN;
        
        int lastSample = 0;
        
        // Iterate through the pixels
        for (int i = 0; i < rect.size.width; i++) {
            // Accumulate min and maxes for all steps that fall within this pixel
            NSPoint thisPixel = [self pixelAlignPoint:NSMakePoint(rect.origin.x + i, 0.)
                                             withSize:NSMakeSize(1., 1.)];

            for (int j = lastSample; j < steps; j++) {
                NSPoint testPixel = [self pixelAlignPoint:NSMakePoint(j * pixelsPerStep + rect.origin.x, 0.)
                                                 withSize:NSMakeSize(1., 1.)];
                
                // If we've left the pixel, break out and start again
                if (!NSEqualPoints(thisPixel, testPixel)) {
                    lastSample = j;
                    break;
                }
                
                // If the sample is NAN don't count it
                if (samples[i].magnitude == NAN) {
                    continue;
                }

                // Collect the minimum and maximum values
                float magnitude = samples[i].magnitude;
                min = (min < magnitude)? min : magnitude;
                max = (max > magnitude)? max : magnitude;
            }
            
            // Draw a line from the minimum to maximum
            float yMin = (min - bottomLevel) * pixelsPerDb + rect.origin.y;
            float yMax = (max - bottomLevel) * pixelsPerDb + rect.origin.y;
            [path moveToPoint:NSMakePoint(thisPixel.x, yMin)];
            [path lineToPoint:NSMakePoint(thisPixel.x, yMax)];
        }
    }
    
    // One-to-one mode means that each sample is exactly a single pixel
    // Undersample mode works the same as one-to-on mode.
    else {
        for (int i = 0; i < steps; i++) {
            // Cycle through the points in the graph converting the dBs to pixels
            float x = i * pixelsPerStep + rect.origin.x;
            
            // If the sample is NAN don't draw it
            if (samples[i].magnitude == NAN) {
                [path moveToPoint:NSMakePoint(x, rect.origin.y)];
                continue;
            }
            
            // Devide the number of steps into the width.
            // Snap all points to device pixels.
            float mag = samples[i].magnitude;
            float magAbove = mag - bottomLevel;
            float y = magAbove * pixelsPerDb + rect.origin.y;
            
            NSPoint point = [self pixelAlignPoint:NSMakePoint(x, y)
                                         withSize:NSMakeSize(1., 1.)];
            
            if (i == 0) {
                [path moveToPoint:point];
            } else {
                [path lineToPoint:point];
            }
        }
    }
    
    [path stroke];
}

- (void)drawRect:(NSRect)dirtyRect
{
    [[NSColor blackColor] set];

    NSBezierPath *framePath = [NSBezierPath bezierPathWithRect:[self bounds]];
    [framePath fill];
    
    float borderWidth = 60;
    NSRect borderRect = NSInsetRect([self bounds],
                                    borderWidth,
                                    borderWidth);

    // Pixel-align the rect
    borderRect.origin = [self pixelAlignPoint:borderRect.origin
                                     withSize:NSMakeSize(1., 1.)];    
    
    // Draw the vertical lines
    [self drawVertGridsInRect:borderRect];
    
    // Draw horizontal lines
    [self drawHorizGridsInRect:borderRect];
    
    NSBezierPath *borderPath = [NSBezierPath bezierPathWithRect:borderRect];
    [[NSColor whiteColor] set];
    [borderPath stroke];

    [self drawHorizGridsInRect:borderRect];
    
    // Draw the actual data
    [self drawDataInRect:borderRect];
    
}

#pragma mark -
#pragma mark UI Code

- (void)updateTextFields
{
    double centerFreq = ([controller startFreq] + [controller stopFreq])  /  2.;
    double freqPerDiv = ([controller stopFreq]  - [controller startFreq]) / 10.;
    [leftTextField   setStringValue:[NSString stringWithFormat:@"%0.6f Mhz", [controller startFreq]]];
    [centerTextField setStringValue:[NSString stringWithFormat:@"%0.6f Mhz", centerFreq]];
    [rightTextField  setStringValue:[NSString stringWithFormat:@"%0.6f Mhz", [controller stopFreq]]];
    [mhzPerDivTextField  setStringValue:[NSString stringWithFormat:@"%0.6f Mhz/Div", freqPerDiv]];
    [topMagTextField setStringValue:[NSString stringWithFormat:@"%1.0f dB", [controller vTopLevel]]];

}

// Pipe these through to the controller
- (bool)textShouldEndEditing:(id)sender
{
    bool result = NO;
    if (sender == leftTextField) {
        result = [controller textShouldEndEditing:[controller scanStartTextField]];
        if (result) {
            [self updateTextFields];
        }
    }
    
    if (sender == centerTextField) {
        result = [controller textShouldEndEditing:[controller scanCenterTextField]];
        if (result) {
            [self updateTextFields];
        }
    }
    
    if (sender == rightTextField) {
        result = [controller textShouldEndEditing:[controller scanEndTextField]];
        if (result) {
            [self updateTextFields];
        }
    }
    
    if (sender == mhzPerDivTextField) {
        result = [controller textShouldEndEditing:[controller scanFDivTextField]];
        if (result) {
            [self updateTextFields];
        }
    }
    
    if (sender == topMagTextField) {
        result = [controller textShouldEndEditing:[controller topLevelTextField]];
        if (result) {
            [self updateTextFields];
        }
    }
        
    return result;
}


@end
