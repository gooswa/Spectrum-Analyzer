//
//  AppDelegate.m
//  Spectrum Analyzer
//
//  Created by William Dillon on 1/18/12.
//  Copyright (c) 2012. All rights reserved.
//

#import "AppDelegate.h"

#define REF_CLOCK 64

void loadComboBoxForPLL(NSComboBox *box);

@implementation AppDelegate

@synthesize window = _window;

- (void)dealloc
{
    [analyzer release];

    [super dealloc];
}

void loadComboBoxForPLL(NSComboBox *box)
{
    NSArray *array = [NSArray arrayWithObjects:
                      @"Tri-state",
                      @"Lock Detect",
                      @"N Divider",
                      @"Active High",
                      @"R Divider",
                      @"Open Drain",
                      @"Serial Data",
                      @"Active Low", nil];

    [box removeAllItems];
    [box addItemsWithObjectValues:array];
}

- (void)applicationDidFinishLaunching:(NSNotification *)aNotification
{
    // Insert code here to initialize your application
    interface = [[Chipkit alloc] initWithPort:@"/dev/tty.usbserial-A6009B2C"];
    analyzer = [[SpectrumAnalyzer alloc] initWithInterface:interface];
    [analyzer setDelegate:self];
    
    // Set this class to be the delegate for all the text fields
    [PLO1_oFreq setDelegate:self];
    [PLO1_pFreq setDelegate:self];
    [PLO1_FoLD  setDelegate:self];
    [PLO2_oFreq setDelegate:self];
    [PLO2_pFreq setDelegate:self];
    [PLO2_FoLD  setDelegate:self];
    [PLO3_oFreq setDelegate:self];
    [PLO3_pFreq setDelegate:self];
    [PLO3_FoLD  setDelegate:self];
    [DDS1_freq  setDelegate:self];
    [DDS3_freq  setDelegate:self];
    
    loadComboBoxForPLL(PLO1_FoLD);
    loadComboBoxForPLL(PLO2_FoLD);
    loadComboBoxForPLL(PLO3_FoLD);
    
    [self configurationChanged];
}

- (void)comboBoxSelectionDidChange:(NSNotification *)notification
{
    NSComboBox *sender = (NSComboBox *)[notification object];
    
    if (sender == PLO1_FoLD) {
        [[analyzer PLO1] setFoLD_output:[sender indexOfSelectedItem]];
        [[analyzer PLO1] updateHardware:interface];
    }

    if (sender == PLO2_FoLD) {
        [[analyzer PLO2] setFoLD_output:[sender indexOfSelectedItem]];
        [[analyzer PLO2] updateHardware:interface];
    }

    if (sender == PLO3_FoLD) {
        [[analyzer PLO3] setFoLD_output:[sender indexOfSelectedItem]];
        [[analyzer PLO3] updateHardware:interface];
    }
}

-(void)invertChanged:(id)sender
{
    if (sender == PLO1_invert) {
        [[analyzer PLO1] setInvertingCP:([sender state] == NSOnState)? YES : NO];
        [[analyzer PLO1] updateHardware:interface];
    }

    if (sender == PLO2_invert) {
        [[analyzer PLO2] setInvertingCP:([sender state] == NSOnState)? YES : NO];
        [[analyzer PLO2] updateHardware:interface];
    }

    if (sender == PLO3_invert) {
        [[analyzer PLO3] setInvertingCP:([sender state] == NSOnState)? YES : NO];
        [[analyzer PLO3] updateHardware:interface];
    }
}

// Modify hardware state for user initiated change
- (bool)textShouldEndEditing:(id)sender
{
    double value = [sender doubleValue];
    
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
    
    if (sender == PLO1_oFreq) {
        [[analyzer PLO1] setOutputFreq:value];
        [PLO1_oFreq setDoubleValue:[[analyzer PLO1] outputFreq]];
    }
    
    if (sender == PLO1_pFreq) {
        [[analyzer PLO1] setPhaseFreq:value];
        [PLO1_oFreq setDoubleValue:[[analyzer PLO1] outputFreq]];
        [PLO1_pFreq setDoubleValue:[[analyzer PLO1] phaseFreq]];        
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

    if (sender == PLO3_oFreq) {
        [[analyzer PLO3] setOutputFreq:value];
        [PLO3_oFreq setDoubleValue:[[analyzer PLO3] outputFreq]];
    }
    
    if (sender == PLO3_pFreq) {
        [[analyzer PLO3] setPhaseFreq:value];
        [PLO3_oFreq setDoubleValue:[[analyzer PLO3] outputFreq]];
        [PLO3_pFreq setDoubleValue:[[analyzer PLO3] phaseFreq]];        
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

-(void)Sweep_go:(id)sender
{
    AnalyzerSample_t *results = [analyzer scanFrom:[Sweep_start doubleValue]
                                                To:[Sweep_stop  doubleValue]
                                         withSteps:[Sweep_steps intValue]
                                          andDelay:[Sweep_delay intValue]];
    
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

-(void)configurationChanged
{
    // For now, we'll just refreah all of the text feilds
    // it's less efficient, but much, much simpler.
    [PLO1_oFreq  setDoubleValue:[[analyzer PLO1] outputFreq]];
    [PLO1_pFreq  setDoubleValue:[[analyzer PLO1] phaseFreq]];
    [PLO1_invert setState:[[analyzer PLO1] invertingCP]? NSOnState : NSOffState];
    [PLO1_FoLD   selectItemAtIndex:[[analyzer PLO1] FoLD_output]];

    [PLO2_oFreq setDoubleValue:[[analyzer PLO2] outputFreq]];
    [PLO2_pFreq setDoubleValue:[[analyzer PLO2] phaseFreq]];
    [PLO2_invert setState:[[analyzer PLO2] invertingCP]? NSOnState : NSOffState];
    [PLO2_FoLD   selectItemAtIndex:[[analyzer PLO2] FoLD_output]];

    [PLO3_oFreq setDoubleValue:[[analyzer PLO3] outputFreq]];
    [PLO3_pFreq setDoubleValue:[[analyzer PLO3] phaseFreq]];
    [PLO3_invert setState:[[analyzer PLO3] invertingCP]? NSOnState : NSOffState];
    [PLO3_FoLD   selectItemAtIndex:[[analyzer PLO3] FoLD_output]];

    [DDS1_freq  setDoubleValue:[[analyzer DDS1] outputFreq]];
    [DDS3_freq  setDoubleValue:[[analyzer DDS3] outputFreq]];
}

@end
