//
//  SpectrumAnalyzer.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 2/3/12.
//  Copyright (c) 2012). All rights reserved.
//

#import "SpectrumAnalyzer.h"

@implementation SpectrumAnalyzer

@synthesize delegate;
@synthesize PLO1;
@synthesize PLO2;
@synthesize PLO3;
@synthesize DDS1;
@synthesize DDS3;
@synthesize adc;
@synthesize interface;

-(id)initWithInterface:(HardwareInterface *)newInterface
{
    self = [super init];
    if (self) {
        interface = newInterface;
        
        // Setup the ADC
        adc = [[ADC alloc] init];
        
        // Setup PLO2
        PLO2 = [[LMX_PLL alloc] init];
        PLO2.CS    = 4;
        PLO2.Data  = 3;
        PLO2.Clock = 2;
        PLO2.FoLD_output = FoLD_RDIV;
        PLO2.refFreq = 64.0;
        PLO2.phaseFreq = 4.0;
        PLO2.outputFreq = 1024.0;
        [PLO2 setDelegate:(id <ModuleDelegate>*)self];
        [PLO2 updateHardware:interface];
        
        // Setup DDS1
        DDS1 = [[AD_DDS alloc] init];
        DDS1.CS    = 6;
        DDS1.Data  = 7;
        DDS1.Clock = 5;
        DDS1.refFreq = 64.0;
        DDS1.outputFreq = 10.7;
        [DDS1 setDelegate:(id <ModuleDelegate>*)self];
        [DDS1 initHardware:interface];
        [DDS1 updateHardware:interface];
        
        // Setup PLO1
        PLO1 = [[LMX_PLL alloc] init];
        PLO1.CS    = 12;
        PLO1.Data  = 11;
        PLO1.Clock = 10;
        PLO1.FoLD_output = FoLD_RDIV;
        PLO1.refFreq = 10.7;
        PLO1.phaseFreq = 0.972;
        PLO1.outputFreq = 1450.;
        [PLO1 setDelegate:(id <ModuleDelegate>*)self];
        [PLO1 updateHardware:interface];
        
        // Setup DDS3
        DDS3 = [[AD_DDS alloc] init];
        DDS3.CS    = 14;
        DDS3.Data  = 15;
        DDS3.Clock = 13;
        DDS3.refFreq = 64.0;
        DDS3.outputFreq = 10.7;
        [DDS3 setDelegate:(id <ModuleDelegate>*)self];
        [DDS3 initHardware:interface];
        [DDS3 updateHardware:interface];
        
        // Setup PLO3
        PLO3 = [[LMX_PLL alloc] init];
        PLO3.CS    = 19;
        PLO3.Data  = 18;
        PLO3.Clock = 17;
        PLO3.FoLD_output = FoLD_RDIV;
        PLO3.refFreq = 10.7;
        PLO3.phaseFreq = 0.972;
        PLO3.outputFreq = 1450;
        [PLO3 setDelegate:(id <ModuleDelegate>*)self];
        [PLO3 updateHardware:interface];
    }
        
    return self;
}

- (void)tuneTo:(double)frequency
{
    
}

- (AnalyzerSample_t *)scanWithDDS:(AD_DDS *)dds
                         fromFreq:(float)startFreq
                           toFreq:(float)endFreq
                        withSteps:(NSInteger)steps
                         andDelay:(NSInteger)mS
{
    double mag, phase;

    // We want to make sure that we actually hit the last frequency
    // therefore, we want to subtract one from the steps.
    double deltaFreq = (endFreq - startFreq) / (float)(steps - 1);
    
    AnalyzerSample_t *results;
    results = (AnalyzerSample_t *)malloc(sizeof(AnalyzerSample_t) * steps);
    if (!results) {
        return nil;
    }
    
    for (int i = 0; i < steps; i++) {
        double freq = startFreq + (deltaFreq * i);
        [DDS1 setOutputFreq:freq];
        [DDS1 updateHardware:interface];
        [adc sampleCount:1
                   Delay:mS
                     Mag:&mag Phase:&phase
               Interface:interface];
        
        results[i].frequency = freq;
        results[i].magnitude = mag;
        results[i].phase = phase;
    }

    return results;
}



-(void)moduleWillUpdateHardware:(id)sender
{
    return;
}

-(void)moduleDidBecomeStale:(id)sender
{
    // Because the DDS1 state changed, we need to re-compute the PLO1 stuff
    if (sender == DDS1) {
        [PLO1 setRefFreq:[DDS1 outputFreq]];
        [DDS1 updateHardware:interface];
        [PLO1 updateHardware:interface];
    }

    // PLO3 is connected to DDS3
    if (sender == DDS3) {
        [PLO3 setRefFreq:[DDS3 outputFreq]];
        [DDS3 updateHardware:interface];
        [PLO3 updateHardware:interface];
    }
    
    if (delegate) {
        if ([delegate respondsToSelector:@selector(configurationChanged)]) {
            [delegate configurationChanged];
        }
    }
}

@end
