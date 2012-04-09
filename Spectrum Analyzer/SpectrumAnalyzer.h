//
//  SpectrumAnalyzer.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 2/3/12.
//  Copyright (c) 2012. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "ModuleDelegate.h"
#import "BusPirate.h"
#import "Chipkit.h"
#import "LMX_PLL.h"
#import "AD_DDS.h"
#import "ADC.h"

typedef struct {
    float frequency;
    float magnitude;
    float phase;
} AnalyzerSample_t;

@protocol SpectrumAnalyerDelegate <NSObject>

@optional

- (void)configurationChanged;
- (void)sampleCollected:(AnalyzerSample_t)sample atStep:(int)step;

@end

@interface SpectrumAnalyzer : NSObject <ModuleDelegate>
{
    HardwareInterface *interface;
    
    LMX_PLL *PLO1, *PLO2, *PLO3;
    AD_DDS  *DDS1, *DDS3;
    ADC     *adc;
    
    double   IF1, IF2, LO2;
    double   RBW;
    
    id <SpectrumAnalyerDelegate> delegate;
    
    int steps, delayMs;
    double startFreq, stopFreq;
    bool scanning;
    
    int currentStep;
    AnalyzerSample_t *samples;
    
    dispatch_queue_t scanQueue;
    
    @private
    BOOL inhibit_callbacks;
}

@property(assign) id<SpectrumAnalyerDelegate> delegate;
@property(readonly) LMX_PLL *PLO1;
@property(readonly) LMX_PLL *PLO2;
@property(readonly) LMX_PLL *PLO3;
@property(readonly) AD_DDS  *DDS1;
@property(readonly) AD_DDS  *DDS3;
@property(readonly) ADC     *adc;
@property(readonly) HardwareInterface *interface;

@property(readonly) int currentStep;
@property(readonly) bool scanning;

@property(readonly) double RBW;
@property(readwrite) int steps;
@property(readwrite) int delayMs;
@property(readwrite) double startFreq;
@property(readwrite) double stopFreq;

// Continuously updated array of samples
// the size of this array matches the number of steps
@property(readonly) AnalyzerSample_t *samples;

- (id)initWithInterface:(HardwareInterface *)interface;

// This is the method that's use to tune the entire
// analyzer to a given frequency.  It does its best
// to reduce the occurance of spurs, etc.
// It also assumes the values set for LO1, LO2, etc.
- (void)tuneTo:(double)frequency;

// Commands to control continuous scanning.
// The scan once commands stop continuous scanning
- (void)startScanning;
- (void)stopScanning;

// This method will perform a sweep using the entire
// analyzer across a given interval using a given number of steps
- (AnalyzerSample_t *)scanOnce;

// This is more of a diagnostics method.  It will
// execute a scan using the DDS module.  It will
// collect ADC data using the magnitude and phase
// ADCs.
- (AnalyzerSample_t *)scanOnceWithDDS:(AD_DDS *)dds;

// This is more of a diagnostics method.  It will
// execute a scan using the PLO module.  It will
// collect ADC data using the magnitude and phase
// ADCs.
- (AnalyzerSample_t *)scanOnceWithPLO:(LMX_PLL *)plo;

@end
