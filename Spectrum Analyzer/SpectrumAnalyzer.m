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
        
        IF1 = 1013.3;
        IF2 = 10.7;
        LO2 = 1024;
        RBW = .002;
        
        inhibit_callbacks = NO;
        
        // Setup the ADC
        adc = [[ADC alloc] init];
        
        // Setup PLO2
        PLO2 = [[LMX_PLL alloc] init];
        PLO2.CS    = 4;
        PLO2.Data  = 3;
        PLO2.Clock = 2;
        PLO2.FoLD_output = FoLD_NDIV;
        PLO2.invertingCP = NO;
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
        PLO1.FoLD_output = FoLD_NDIV;
        PLO1.invertingCP = YES;
        PLO1.refFreq = 10.7;
//        PLO1.refFreq = 64;
        PLO1.phaseFreq = 0.972;
        PLO1.outputFreq = 1034;
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
        PLO3.invertingCP = YES;
        PLO3.refFreq = 10.7;
        PLO3.phaseFreq = 0.972;
        PLO3.outputFreq = 1034;
        [PLO3 setDelegate:(id <ModuleDelegate>*)self];
        [PLO3 updateHardware:interface];
    }
        
    return self;
}

- (void)tuneTo:(double)frequency
{
    inhibit_callbacks = YES;
    // The first IF comes from mixing LO1 with the input (PLO1)
    // LO2 is fixed at 1024 Mhz
    double LO1 = frequency + IF1;

    // Because we can only tune PLO1 in .97 (PLO1.phaseFreq) steps,
    // we need to approximate the LO1 tuning value.
    PLO1.refFreq = IF2;
    PLO1.outputFreq = LO1;
    
    // Now, PLO1 is "stale" (which means that its software representation
    // doesn't match the hardware.  That's OK because we only want the
    // calculated tuning value.
    double PLO1_error = LO1 - PLO1.outputFreq;
    
    // Next, using the frequency error, we can figure out a new fine tuning
    // of the DDS to make the PLO1 output frequency very close to desired
    double correctedDDS1 = IF2 + (PLO1_error/PLO1.n_divider) * PLO1.r_divider;
    DDS1.outputFreq = correctedDDS1;
    PLO1.refFreq = DDS1.outputFreq;
    [DDS1 updateHardware:interface];
    [PLO1 updateHardware:interface];

    // Check for the possibility for spurs (this may be garbage)
/*    double firstIF   = LO2 - IF2;
    double fractFreq = DDS1.outputFreq / (PLO1.r_divider * 16);
    double harmonicb = round(firstIF/fractFreq);
    double harmonica = harmonicb - 1;
    double harmonicc = harmonicb + 1;
    double firstIF_low  = LO2 - (IF2 + RBW);
    double firstIF_high = LO2 - (IF2 - RBW);
    
    BOOL spur = NO;
    if (harmonica * fractFreq > firstIF_low &&
        harmonica * fractFreq < firstIF_high) {
        spur = YES;
    }

    if (harmonicb * fractFreq > firstIF_low &&
        harmonicb * fractFreq < firstIF_high) {
        spur = YES;
    }

    if (harmonicc * fractFreq > firstIF_low &&
        harmonicc * fractFreq < firstIF_high) {
        spur = YES;
    }

    // There was a spur.  This means we need to tweak the value of the
    // DDS such that we move the spur out of the passband of the final filter
    // Do do this, we'll either increment or decrememnt the N counter of the
    // PLO, and compute an appropriate DDS frequency using it.
    if (spur) {
        printf("SPUR DETECTED!!\n");
    }
*/    
    
    inhibit_callbacks = NO;
    
    printf("Tried to tune the analyzer to %f, achieved %f.  Error of %f hz.\n",
           frequency,
           PLO1.outputFreq - IF1,
           (frequency - (PLO1.outputFreq - IF1)) * 1000000);
}

-(AnalyzerSample_t *)scanFrom:(double)startFreq
                           To:(double)stopFreq
                    withSteps:(NSInteger)steps
                     andDelay:(NSInteger)delay
{
    double mag, phase;
    
    // We want to make sure that we actually hit the last frequency
    // therefore, we want to subtract one from the steps.
    double deltaFreq = (stopFreq - startFreq) / (float)(steps - 1);
    
    AnalyzerSample_t *results;
    results = (AnalyzerSample_t *)malloc(sizeof(AnalyzerSample_t) * steps);
    if (!results) {
        return nil;
    }
    
    for (int i = 0; i < steps; i++) {
        double freq = startFreq + (deltaFreq * i);
        
        [self tuneTo:freq];
        [adc sampleCount:1
                   Delay:delay
                     Mag:&mag
                   Phase:&phase
               Interface:interface];
        
        results[i].frequency = freq;
        results[i].magnitude = mag;
        results[i].phase = phase;
    }
    
    return results;
}

- (AnalyzerSample_t *)scanWithDDS:(AD_DDS *)dds
                         fromFreq:(float)startFreq
                           toFreq:(float)stopFreq
                        withSteps:(NSInteger)steps
                         andDelay:(NSInteger)mS
{
    double mag, phase;

    // We want to make sure that we actually hit the last frequency
    // therefore, we want to subtract one from the steps.
    double deltaFreq = (stopFreq - startFreq) / (float)(steps - 1);
    
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

- (AnalyzerSample_t *)scanWithPLO:(LMX_PLL *)plo
                         fromFreq:(float)startFreq
                           toFreq:(float)stopFreq
                        withSteps:(NSInteger)steps
                         andDelay:(NSInteger)mS
{
    double mag, phase;
    
    // We want to make sure that we actually hit the last frequency
    // therefore, we want to subtract one from the steps.
    double deltaFreq = (stopFreq - startFreq) / (float)(steps - 1);
    
    AnalyzerSample_t *results;
    results = (AnalyzerSample_t *)malloc(sizeof(AnalyzerSample_t) * steps);
    if (!results) {
        return nil;
    }
    
    for (int i = 0; i < steps; i++) {
        double freq = startFreq + (deltaFreq * i);

        inhibit_callbacks = YES;

        // Because we can only tune PLO1 in .97 (PLO1.phaseFreq) steps,
        // we need to approximate the LO1 tuning value.
        PLO1.refFreq = IF2;
        PLO1.outputFreq = freq;
        
        // Now, PLO1 is "stale" (which means that its software representation
        // doesn't match the hardware.  That's OK because we only want the
        // calculated tuning value.
        double PLO1_error = freq - PLO1.outputFreq;
        
        // Next, using the frequency error, we can figure out a new fine tuning
        // of the DDS to make the PLO1 output frequency very close to desired
        double correctedDDS1 = IF2 + (PLO1_error/PLO1.n_divider) * PLO1.r_divider;
        DDS1.outputFreq = correctedDDS1;
        PLO1.refFreq = DDS1.outputFreq;
        [DDS1 updateHardware:interface];
        [PLO1 updateHardware:interface];

        inhibit_callbacks = NO;
        
        [adc sampleCount:1
                   Delay:mS
                     Mag:&mag Phase:&phase
               Interface:interface];

//        printf("Tried to tune the analyzer to %f, achieved %f.  Error of %f hz.\n",
//               freq,
//               PLO1.outputFreq,
//               (freq - PLO1.outputFreq) * 1000000);        

        results[i].frequency = PLO1.outputFreq;
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
    if (inhibit_callbacks == NO) {
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
        
        if (sender == PLO1) {
            [PLO1 updateHardware:interface];
        }
        
        if (sender == PLO2) {
            [PLO2 updateHardware:interface];
        }
        
        if (sender == PLO3) {
            [PLO3 updateHardware:interface];
        }

        if (delegate) {
            if ([delegate respondsToSelector:@selector(configurationChanged)]) {
                [delegate configurationChanged];
            }
        }
    }
}

@end
