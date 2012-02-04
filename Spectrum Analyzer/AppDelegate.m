//
//  AppDelegate.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/18/12.
//  Copyright (c) 2012. All rights reserved.
//

#import "AppDelegate.h"

#define REF_CLOCK 64

@implementation AppDelegate

@synthesize window = _window;

- (void)dealloc
{
    [analyzer release];

    [super dealloc];
}

- (void)applicationDidFinishLaunching:(NSNotification *)aNotification
{
    // Insert code here to initialize your application
    HardwareInterface *interface;
    interface = [[Chipkit alloc] initWithPort:@"/dev/tty.usbserial-A6009B2C"];
    analyzer = [[SpectrumAnalyzer alloc] initWithInterface:interface];
    
    // Set this class to be the delegate for all the text fields
    [PLO1_oFreq setDelegate:self];
    [PLO1_pFreq setDelegate:self];
    [PLO2_oFreq setDelegate:self];
    [PLO2_pFreq setDelegate:self];
    [PLO3_oFreq setDelegate:self];
    [PLO3_pFreq setDelegate:self];
    [DDS1_freq  setDelegate:self];
    [DDS3_freq  setDelegate:self];
}

// 
- (bool)textShouldEndEditing:(id)sender
{
    double value = [sender doubleValue];
    
    // If this a DDS text field, the number can't be more than 1/3 of reference
    if (sender == DDS1_freq) {
        
        // Affect the change, any error checking will happen within the class
        // The analyzer will subscribe to any changes, and propogate changes
        // across the other modules when the change happens
        [[analyzer DDS1] setOutputFreq:value];
        [sender setDoubleValue:[[analyzer DDS1] outputFreq]];
    }
    
    if (sender == DDS3_freq) {
        [[analyzer DDS3] setOutputFreq:value];
        [sender setDoubleValue:[[analyzer DDS3] outputFreq]];
    }
    
    if (sender == PLO2_oFreq) {
        [[analyzer PLO2] setOutputFreq:value];
        [PLO2_oFreq setDoubleValue:[[analyzer PLO2] outputFreq]];
    }
    
    if (sender == PLO2_pFreq) {
        [[analyzer PLO2] setPhaseFreq:value];
        [PLO2_oFreq setDoubleValue:[[analyzer PLO2] outputFreq]];
        [PLO2_pFreq setDoubleValue:[[analyzer PLO2] phaseFreq]];        
    }    

    return NO;
}

- (void)controlTextDidEndEditing:(NSNotification *)obj
{
    id sender = [obj object];
    
    [self textShouldEndEditing:sender];
}

-(IBAction)DDS1_Scan:(id)sender
{
    AnalyzerSample_t *results = [analyzer scanWithDDS:[analyzer DDS1]
                                             fromFreq:[DDS_start doubleValue]
                                               toFreq:[DDS_stop  doubleValue]
                                            withSteps:[DDS_steps intValue]
                                             andDelay:[DDS_delay intValue]];
    
    // Make sure it worked.  If it did, pop up a window to save the CSV
    // file to disk somewhere.  Eventually we should be able to make a graph
    if (results) {
        // Open the save file as dialog box, concrrent with string processing
        NSSavePanel *savePanel = [NSSavePanel savePanel];
        [savePanel setCanSelectHiddenExtension:YES];
        [savePanel setTitle:@"Save as a CSV"];
        [savePanel setPrompt:@"Save"];
        NSInteger result = [savePanel runModal];        
        
        // Get the result of the users selection
        if (result == NSFileHandlingPanelOKButton) {
            NSError *error = nil;
            
            // Create an NSString with the CSV data
            NSMutableString *tempString = [[NSMutableString alloc] init];
            for (int i = 0; i < [DDS_steps intValue]; i++) {
                [tempString appendFormat:@"%f,%f,%f\n",
                 results[i].frequency,
                 results[i].magnitude,
                 results[i].phase];
            }
            
            // Save the CSV into the file
            bool result = [tempString writeToURL:[savePanel URL]
                                      atomically:NO
                                        encoding:NSUTF8StringEncoding
                                           error:&error];
            
            [tempString release];
            
            // Pop-up a status message
            NSAlert *completionAlert = [[NSAlert alloc] init];
            [completionAlert addButtonWithTitle:@"OK"];
            
            if (result == NO) {
                NSLog(@"Unable to save log CSV to %@: %@",
                      [savePanel URL], [error localizedDescription]);
                
                NSString *errorString;
                errorString = [NSString stringWithFormat:@"Error saving file: %@",
                               [error localizedDescription]];
                
                [completionAlert setMessageText:@"Unable to save"];
                [completionAlert setMessageText:errorString];
                [completionAlert setAlertStyle:NSWarningAlertStyle];
            } else {
                [completionAlert setMessageText:@"File saved"];
                [completionAlert setMessageText:@"Cluster log saved"];
                [completionAlert setAlertStyle:NSInformationalAlertStyle];			
            }
            
            [completionAlert runModal];
            [completionAlert release];
        }
    }
}

- (IBAction)ADC_go:(id)sender
{
    double mag, phase;
    
    [[analyzer adc] sampleCount:[ADC_count intValue]
                          Delay:[ADC_delay intValue]
                            Mag:&mag Phase:&phase
                      Interface:[analyzer interface]];
    
    [ADC_mag   setFloatValue:mag];
    [ADC_phase setFloatValue:phase];
}

@end
