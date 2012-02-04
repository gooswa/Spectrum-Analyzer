'version 116-4q
'Created from version 116-4p
'Reason for change: 

'------------Changes and additions ------------
'116-4q was released as the 8th beta
'116-4q
'  Change revision to H
'  Change boolean to long in Windows API calls (but not USB driver calls).
'  Adjust AutoWait box.
'  Add call to CoInitialize to fix tooltip crash in File Dialog
'
'116-4p was released as the 7th beta
'116-4p
'   Remove debug code for testing OSL
'   Fix sig gen. It does not need to be in band. Whatever it is, we generate as LO3-LO2, LO3+LO2 or plain LO3, whichever
'   one works for the given sgout.
'116-4o
'   Add Scotty changes for USB without control board, per his Lrev1
'   (Set-up the USB using the N bits rather than assuming none of them will ever change. They will if user plays with the LB code.)
'116-4n
'   Restore Reference to OSL. It is needed for Update Cal.
'   Apply OSL even with series and shunt fixture--just create default values for the missing cal items.
'   Change display of OSL data window to delete tiny values.
'   OSLBandApplyFull, OSLApplyFull and OSLBaseApply are no longer be used. We always do full OSL.
'116-4m
'   Change revision to G
'   Change some window text to use sans serif 9 rather than default values.
'   Fix group delay calc in uBestFitLine
'   Fix to CopyModeDataToVNAData to get freq correct
'   Fix very original startup data type 
'   Fix MouseOver in Two Port. Various Two Port tweaks.
'   Fix application of line cal Through delay
'   Remove TwoPort Matrix multiply and invert, which were not used.
'   When leaving Two Port, minimize array size only if we have no entries.
'
'Note: 116-4L was released as the 6th beta
'116-4L
'   Change revision to F
'   Change text sizes in PDM cal
'   When trace preference is loaded, set axis label to match trace color.
'   Change some colors and text in interface windows.
'   Move point at which phase is negated in 3G so it occurs after adjustment for PDM and difPhase.
'   Fix LO3 calculation regarding use of baseFrequency
'
'Note: 116-4k was released as the fifth beta
'116-4k
'   Modify reference cal in Reflection mode so the Open or Short does not have to be ideal.
'   When doing full OSL, treat all fixtures as Bridges.
'   Fixes to interpolation of angles. In graph modules, angles are kept in the graph range. Elsewhere they
'   are kept in the range -180 to +180.
'   Fix Smith Chart interpolation to match method for main graph. This means the interpolated point may not quite
'       match the graph on the Smith Chart, which linearly interpolates pixels.
'   Round markers to integral point if they are manually placed very close to an integral point.
'   Allow log sweeps to go to zero or negative, but only Y-axis can cross zero.
'   Add baseFrequency, which gets added to the current frequency when commanding the hardware.
'   Change Y pixel calculations for log scale to use gConvertYxToPix
'   In Reflection mode, delete printing of bridge/series/shunt in the setup info.
'   Add warning if number of steps exceeds 2000. Show hourglass after while calculating steps after restart.
'   Relocate printing of sweep time so it does not immediately get erased by refresh action.
'   In config manager, clarify listing of PLL types. For example, there is now an entry for 2326,4118 which
'     includes both types, though internally it is still stored as simply 2326.
'   Fix [CreateCmdAllArray]--last line got incorporated into a comment
'
'116-4j
'   Change revision to E
'   When restoring VNA data in reflection mode, re-calculate all data derived from S11
'   We now set both S11BridgeR0 and S21JigR0 to the same value in reflection mode. When doing full OSL, the user does not choose the
'      fixture type so we wouldn't know which one to set if we didn't set both. Certain routines refer to one or the other,
'      depending on fixture type, but they will get the same value either way. 
'   Fixes to use of [RestoreVNAData] regarding application of R0 transform and plane extension.
'   Adjust AutoWait to use max of mag and phase time constants
'   Fix PLL configuration labels
'   Fix crash when placing marker with Variables window open
'   Turn RBW filter selection into subroutines, and call when necessary from DetectChanges. Previously, necessary
'     recalculations were not always done when filter changed. Now we require [Restart] or do [PartialRestart].
'   Change Initial Cal so wait time is based on video time constants, and in Measure do a one-time [CommandCurrentStep]
'    to be sure we are properly commanded per current path frequency. Otherwise we get messed up when
'    changing paths.
'   Make filter A0 and A1 into globals, renamed FiltA0 and FiltA1.
'   Remove surplus code at end of sweep loop relating to haltAtEnd=0. It followed similar code that prevented
'    it from ever being executed.
'   Change mag cal to average fewer points gathered over a longer time.
'   Add item in Initial Cal to allow user to force phase correction to a set value (normally 0 or 180)
'
'Note: 116-4i was posted as the fourth beta
'116-4i
'   Change OSL standards specs to use regular RLC specs.
'   Add MouseOver to TwoPort window
'
'116-4h
'   Revise transmission mode RLC analysis to allow user choice as to whether to analyze notches at top or bottom.
'   Modify latched switches for USB.
'   When saving bitmaps with BMPSAVE, don't open and close file.
'   Modify transmission mode RLC Analysis
'   Add routine to calc RLC values from impedances at two frequencies
'   Delete lossy inductor option from RLC analysis
'   Add "mouse over" feature to show marker at current mouse frequency and allow setting of
'     markers at that frequency from the keyboard.
'   Fixes to doSpecialGraph so its phase does not get overridden by phase adjustments that apply to normal graphs.
'   Add Default Graph option to Y-axis window
'
'116-4g
'   changed revision to D
'   Clear markers when altering Two-Port data
'   Revise transmission RLC analysis for notches to use points 3 dB above the notch.
'   Clarify filter analysis instructions.
'   Delete autoWaitSettlingTime routine; it is no longer used
'
'Note: 116-4f was posted as the third beta
'116-4f
'   When selecting video filter, be sure capacitor values are not zero
'   When loading config file, be sure we have a Wide filter
'   Don't initialize relays if suppressHardware=1
'   Move a USB command line from DDS3Track to DDS1Sweep
'
'116-4e
'   change revision to C
'   Add latch numbers to latch enable pins in LPT tester.
'   Fix filter analysis reference to #handle.g
'   Fix [ReadStep] so phase is not read unnecessarily, and zero phase array if not doing phase.
'   Fixes to selection of auto wait when auto wait is disabled.
'   Change two port matching instructions.
'   Fix axis ranges in Two Port when changing data types
'   Save Two-Port params as .s2p rather than .s1p.
'   Provide for single port matching in AutoMatch. Start Two-Port with default parameter values.
'
'Note: 116-4d was posted as the second beta version
'116-4d
'   Make changes so hardware initilization always occurs on [PartialRestart] unless
'     suppressHardwareInitOnRestart is expressly set to 1. Set it to 1 only when
'     hardware init is pointless and we want to save time.
'   Fix uSaveFileDialog and uOpenFileDialog
'   Add switchTR to sweep context for saving and restoring
'   In display, abbreviate Sig Gen= to SG=
'
'116-4c
'   Changed Revision to B
'   Fix reference to "handle" in Filter function
'   Add Coaxial Cavity Sweep Test in Special Tests window

'Note: 116-4b was posted as beta software
'116-4b
'   Deleted option to graph Z parameters. Z parameters aren't that helpful. Instead,
'     we have added the ability to graph S11 and S22 as impedance and other items.
'   Moved Two Port module so it is above smith chart module.
'   Add resizing to Two Port Window.
'   Various improvements to Two Port module
'   Deleted comments on version 115 changes. File got too big for LB Workshop.
'   Add AutoMatch feature to Two-Port Window
'   Add display of mouse info on Smith Chart when left button is down
'   Allow trace widths of 0.
'   Modify refresh at end of each scan
'   Have functions and calibrations use AutoWait
'   Delete commented-out code for GraphCoaxLossAndZ0 to get more file space.
'   Fix to gGetCustomPresetColors
'   When saving grid context, put LastPresets at end, after the options have been recorded.
'   Add colors for two supplemental sets of traces: 1A, 2A, 1B and 2B
'   Tie reference trace colors to the 1A and 2A colors.
'   Don't set suppressHardware=1 if we have USB control board
'   Show "Show Variables" menu item for basic MSA build
'   Modify uOpenFileDialog to use up-to-date struct. Made no difference to occassional crash
'     that occurs when tool tips are displayed the second time through. Turns out it only
'     occurs when the target file (that mouse is hovering over) is on the desktop.

'116-4a
'   Change constDatatable to constModeData because it is not always data from datatable.
'   Change touchWriteParameters to touchWriteOnePortParameters because it only writes a single parameter,
'   and have it save data from VNAData as an intermediary rather than from specific scan data arrays.
'   Various improvements to Two Port module
'   Add ability to graph Z parameters in Two Port
'
'116-3a
'   Incorporate USB changes since 116-1b from Dave Roberts' version 116-1g
'   '  USB:02-08-2010  USB:05/12/2010 ' USB:15/08/10 'USB:06-08-2010
'116-2a
'   Reorganize order of routines to put them in logical groups
'   Expand Two Port module to display performance with specified impedance matching.
'   Scotty's changes to fix to use ADF4112. Changes marked '12-3-10
'   Fix uExtractDataItem so multiple delimiters are not deleted at once
'   Adjust marker routines so markers work on Two Port Window

'116-1b
'   Add extra element to configControlBoards for cb=3 (USB)
'   Change Data Window to display only the actual number of steps actually scanned.
'   Physically select path filter in Calibration Manager
'   Fix setup info so when reference line is turned off, no reference info is displayed.
'   Change interpolation of path calibration so use of 180 degree phase correction to mark "phase invalid"
'     does not mess up the interpolation.
'   When saving graph or smith chart image, don't try to unload bitmap if operation cancelled.
'   When Initial Cal Manager is invoked, do nothing if already running.
'   Fix file saving when leaving Initial Cal Manager
'   Fix saving of PLL1 and PLL2 in Configuration Manager
'   Fix 12-bit ADC reading
'   Add switching of video filter, band switch, forward/reverse and transmission/reflection switches.
'   create validPhaseThreshold so PDM inversion is not done at low levels where phase is forced to zero.
'   Add features for Auto Wait Time.
'   Add TwoPort module
'   change topref, botref, topphase,botphase to Y2Top, Y2Bot, Y1Top and Y1Bot
'   Multiply raw phase by -1 in 3G mode
'   Force sig gen to be in-band when mode is changed
'   For doSpecialGraph=3(noisy sine), increase transmission values of dB
'   Fix clearing of ReflectArray before installing a step of data--step 0 was always cleared.
'
'116-1a
'   Change Version to 116 and Revision to A
'   USB test version 01/08/2010 copied from Dave Roberts
'   all USB changes tagged with code "USB:date"
'   so search for the string "USB:" to find changes
'   Additional USB changes 02-08-2010 by S. Wetterlin
'115-9f
'    Change Revision to H
'    Fix calc of power and voltage for graphing.
'115-9e
'    Change Revision to G
'    When resizing window, don't require restart if sweep is not in progress.
'    Fix display of Initial Cal Manager in non-VNA modes
'Note: 115-9d was released as Revision F
'115-9d
'    Add routines for front-end adjustments of SA readings.
'    Fix saving of reflection Test Setups to include R0
'115-9c
'    Change Revision to F.
'    Adjust Zoom in crystal analysis.
'    Change crystal Cm to fF and Lm to mH
'Note: 115-9b was released as Revision E.
'---------------------------------------------------------------------------------------------

    global msaVersion$, msaRevision$  'Version and revision numbers of this release
    msaVersion$="116"   'ver116-1a
    msaRevision$="H"    'ver116-4m

'-------------Sequence of Main Routine for Original or Slim MSA/TG/VNA---------------
'Notes: The SLIM MSA (SLIM Control Board) can command all 6 modules at one time (PLO1,DDS1,PLO3,DDS3,PDM,FilterBank).
'        However, the FilterBank is commanded independently. So is PLO2.
'       The Original MSA (original Control Board) must command the modules independently.

'1.Establish User Global variables from external msaconfig.txt
'2.Establish hard Global variables
'3.Create Working Window, for Spectrum Analyzer Mode, and insert the Default Global Variables
'4.measure computer speed and update global, glitchtime
'5.Command Filter Bank to Path one
'   access the Path 1 Magnitude Calibration Table
'   access the MagError vs Freq Calibration Table
'6.if configured, initialize DDS3 by resetting to serial mode. Frequency is commanded to zero
'7.if configured, initialize PLO3. No frequency command yet.
'8.initialize and command PLO2 to proper frequency
'9.Initialize PLO 1. No frequency command yet.
'10.initialize DDS1 by reseting to serial mode. Frequency is commanded to zero
'11.[GrabWorkingWindowData] get info from Working Window and update variables
'12.[CreateGraphWindow], using Working Window data
'13.Calculate the command information for first step through last step of the sweep and put in arrays
'14.[StartSweep]'Begin sweeping from step 0
'15.[CommandThisStep](1.9ms). command relevant Control Board and modules
'16.Determine sequence of operations after commanding the modules
'    a. if in OneStep mode, then
'       add extra settling time
'       read data of thisstep. gosub [ReadStep]
'       process data of thisstep. gosub [ProcessAndPrint]
'       wait here, until the next button push
'    b. if ThisStep is the first step after a halt, then
'        add extra settling time
'        read data of thisstep. gosub [ReadStep]
'        goto [Scan]
'    c. if ThisStep is not the first step after halt (middle of sweep), then
'        process and print the Previous Step. gosub [ProcessAndPrintLastStep](3.8ms)
'           [ProcessAndPrint](3.8ms)                           /1.9
'               [ConvertMagPhaseData](2.5ms)                      /.53
'                   call calConvertMagPhase(1.2ms)
'                   freqerror=calConvertFreqError(thisfreq)(0.9ms)  /.085
'                   [CalcMagpowerPixel](0.4ms)
'               [PlotDataToScreen](1.3ms)                         /1.3
'        read data of thisstep. gosub [ReadStep](1.9ms)
'        goto [Scan]
'17.[Scan] Check to see if a button has been pushed
'   If a button was pushed (other than OneStep) goto [Halted]
'   If not, continue to [IncrementOneStep]
'18.[IncrementOneStep]
'   add 1 to the value of thisstep
'   if the new value of thisstep exceeds the number of steps in this sweep
'       "Glue" the full sweep plot into memory
'       go back to [StartSweep]
'   if not, go back to [CommandThisStep] and continue sweeping
'19.[Halted]
'   process and print ThisStep. gosub [ProcessAndPrint]
'   reprint Graph lines and text. gosub [PrintGraph]
'   "Glue" Graph into memory
'   wait for operator action.

'--------Start of Code, Main Routine---------
'1.Establish User Global variables from external msaconfig.txt. all of the following are in msaconfig.txt

'[EstablishUserVariables]
    '----Start of global variables set from a configuration file; some depend on construction of the
    'spectrum analyzer; others are just convenient defaults that can be changed at runtime.---------
    '[EstablishUserVariables]  'all of the following are "default" values and are dependent on the construction of your Spectrum Analyzer
    global masterclock  'Exact frequency of the Master Clock (in MHz).  Example: 64.000056 or 63.999937
                  'You can start with default configuration and change after calibration.
    global centfreq     'Sweep center frequency, in MHz. For initial set-up use "0"
    global sweepwidth   'Sweep width in MHz. For initial set-up use 10 times the BW of Final Xtal Filter
'delver113-7c    global wate         'value to "slow" the sweep speed for more accurate data.  Use "0" as default.
'delver113-7c    global glitchtime   'default=0, causing a self determination at startup.
    global adconv      'AtoD topology."8" for original 8 bit,"12" for optional 12 bit ladder,"16" for serial 16 bit AtoD, or "22" for serial 12 bit AtoD  ver111-36e
    global Y1Top, Y1Bot     'Top and bottom of Y1 (right) axis
    global Y2Top, Y2Bot      'Top and bottom of Y2 (right) axis
    global cb          '0= old Control Board, 1= old Control Board with new harness, 2 = SLIM Control Board ver111-22
    global dds1parser  '0 if DDS 1 is in parallel mode, 1 if serial 'ver111-21."0" not allowed on SLIM Control Board ver111-29
    global appxdds1    'nominal DDS1 output frequency (in MHz) that steers PLL 1; . Near 10.7.
                        'appxdds1 must be the center freq. of DDS1 xtal filter; exact value determined in calibration.
    global dds1filbw   'DDS1 xtal filter bandwidth (in MHz), at the 3 dB points).
    global PLL1    'PLL 1 model. Allowed values are 2325, 2326, 2350, 2353, or 4112
    global PLL1phasefreq   'approx. Phase Detector Frequency (MHz) for PLL 1. Use .974 when DDS1 filter is 15 KHz wide
                            'PLL1phasefreq must be less than the following formula:
                            'PLL1phasefreq < (VCO 1 minimum frequency) x dds1filbw/appxdds1
    global PLL1mode      '0 = Integer Mode, 1 = Fractional Mode for PLL 1.
                  'I don't recommend Fractional Mode for PLL 1, although it will work (noiser)
    global PLL1phasepolarity  '1 for non-inverting loop filter1; 0 inverting op amp; enter 0 for SLIM MSA
    global PLL2        'PLL 2 model. Allowed values are 2325, 2326, 2350, 2353, or 4112, or 0 for SRD multiplier
    global appxLO2      '2nd LO frequency (MHz). 1024 is nominal, Must be integer multiple of PLL2phasefreq
    global PLL2phasefreq   'PLL2 phase frequency (MHz).  See appxLO2". 4 is nominal
    global PLL2phasepolarity   'for non-inverting loop filter, enter 1(SLIM MSA); for inverting op amp, enter 0
    global TGtop        'Tracking Generator Topology:"0" for not installed, "1" for original Trk Gen,
                        '"2" for New TG (DDS3/PLL3 combination) ver111-18
    global PLL3        'PLL 3 model. Allowed values are 2325, 2326, 2350, 2353, or 4112, or 0 for no Trk Gen
                ' 2350 and 2353 can be used as fractional-N
    global appxdds3    'nominal DDS3 output frequency (in MHz) that steers PLL 1;  Near 10.7.
                'appxdds3 must be the center freq. of DDS3 xtal filter; exact value determined in calibration.
                'Enter "0" if TGtop = 0(no TG) or 1.  The original Trk Gen does not use a DDS3. ver111-17

    global dds3filbw   'DDS3 xtal filter bandwidth (in MHz), at the 3 dB points).
    global PLL3phasepolarity  '1 for non-inverting loop filter; 0 for inverting op amp; enter 0 for SLIM MSA
    global PLL3mode     '0 = Integer Mode, 1 = Fractional Mode for PLL 1.
                 'Enter "1" only if PLL3 = 2350 or 2353.  Enter "0" for new DDS3/PLL3 combination (SLIM MSA).
    global PLL3phasefreq   'TrkGen PLL3 phase frequency (MHz). If TG is DDS3/PLL3 combo, use same technique as PLL1phasefreq
                     'if Original TG (TGtop=1) then this must be a sub-multiple of both Master Clock and Final Xtal Filter Frequency
    'global sgpreset    'delver114-4h
    global maxpdmout   'bit count for Phase AtoD converter when Phase Det Module output is maximum. ver112-1a
        'For SLIM-ADC-16 = 65535, SLIM-ADC-12 = 4095.  For the Original Control Board and:
        '12 Bit Parallel AtoD = 4095, 16 Bit Serial AtoD = 65535, or 8 Bit Parallel AtoD = 255. These are adjustable during calibration
    global invdeg  'actual phase change when PDM is inverted at a finalfreq of 10.7 MHz. Nominally, 180.  ver116-1b
    global doingPDMCal  '=1 when PDM cal in progress to determine invdeg ver114-5L
    global CalInvDeg    'set to value of invdeg determined by cal ver114-5L
    global cftest       '=1 when doing cavity filter sweep test (Special Tests window)  'ver116-4b
'--SEW End of variables initialized from configuration file

'--SEW2 added the following global declarations to make these available to true subroutines
'del13-7c    global steps    'whole number of steps per sweep. 1 thru 720 is acceptable.  400 is a good number.
'del13-7c    global thisstep 'keeps track of current step number during a sweep
    global globalPort     'Used to pass new port value back from config routine w/o making port global. ver113.7c
                            'ver114-5k deleted globalGlitchtime, which was no longer used
    global globalSteps    'SEWgraph; Number of steps set by user. Set in calcWindoInfo. global version of steps.
    global varwindow, datawindow    '=1 when indicated window is open   'ver115-1b

    global doSpecialGraph '=0 for normal operation; for other values see [doSpecialGraph]
    global doSpecialRandom  'Random number generated at start of each sweep, for doSpecialGraph ver 114-3g
    global doSpecialRLCSpec$ 'RLC spec for doSpecialGraph
    global doSpecialCoaxName$ 'Name of coax last used in RLC spec ver115-4b

    global finalfreq, finalbw       'freq and bandwidth of current final filter
    global LO2          'actual LO2 frequency ver115-1c
    global hasVNA               '=1 if the build includes VNA; otherwise 0
    global hGraphWindow  'Windows handle to the graph window; set when created  ver115-5d
    global hGraphMenuBar 'Windows handle to menu bar in graph window; set when created  ver115-5d
    global suppressPhase    '=1 to force phase to zero without measuring it ver116-1b

    global hFileMenu, hOptionsMenu,hDataMenu,hFunctionsMenu,hOperatingCalMenu, hMultiscanMenu  'Windows handles to some graph window submenus ver115-5d
    global hTwoPortMenu 'ver116-1b
    global menuOperatingCalShowing  '=1 when the operating cal menu is shown; otherwise 0
    global menuMultiscanShowing, menuTwoPortShowing 'same for multiscan and two-port menus ver116-1b

        'zero-based index of menu positions in menu bar
    global menuOptionsPosition, menuDataPosition, menuFunctionsPosition, menuOperatingCalPosition, menuMultiscanPosition
    global menuTwoPortPosition, menuModePosition 'ver116-1b

        'IDs of various listed menu items that may need to be hidden/shown in certain modes
    global menuDataS21ID, menuDataLineCalID, menuDataS11ID, menuDataS11DerivedID  'ver115-5d
    global menuDataLineCalRefID, menuDataLineCalOSLID  'ver115-5d
    global menuOptionsSmithID 'ver115-5d
    global menuFunctionsFilterID, menuFunctionsCrystalID, menuFunctionsMeterID, menuFunctionsRLCID  'ver115-5d
    global menuFunctionsCoaxID,menuFunctionsGenerateS21ID, menuFunctionsGroupDelayID 'ver115-8b
    
    global twoPortWinHndl$  'handle of open Two-Port window, or blank if not open ver116-1b
    global calManWindHndl$   'holds handle of open window for calibration manager, or blank if not running
    global configWindHndl$    'Handle to our main window SEWcal3 moved to here from cal module
    global componentWindHndl$   'Handle to component measurement window
    global crystalListHndl$ 'Handle to crystal list windows; blank if not open   'ver115-5c deleted crystalWindHndl$
    global crystalWindHndl$ 'Handle to crystal window
    global crystalLastUsedID    'Last ID of crystal added to crystal list
    
    global tranRLCLastRLCWasSeries, tranRLCLastNotchWasTop 'most recent settings for tran mode RLC analysis
    global imageSaveLastFolder$  'Folder in which last graph image was saved, regular or Smith chart ver115-2a

    dim Appearances$(10)    'Names of Appearances
    dim customPresetNames$(5)   'User names for custom color presets (1-5) ver115-2a

        'configFilters is a list of final filters; second dimension 0=freq (MHz), 1=BW (KHz);
        'zero entry of first dimension not used
    dim MSAFilters(40,1)
    global MSANumFilters 'Number of filters in list
    dim MSAFiltStrings$(40)  'Same info as MSAFilters(), but freq and bw are combined in a string; zero entry is used

 '------SEWgraph globals for graph params
    global firstScan   'Set to 1 for first scan after background grid for graph is drawn
'ver114-4e deleted global graphAppearance$
    global currGraphBoxHeight, currGraphBoxWidth    'Actual current height, width, adjusted when resizing 'ver115-1c
    global clientHeightOffset, clientWidthOffset    'Difference between window size and client area size; set from test window 'ver115-1c
    global smithLastWindowHeight, smithLastWindowWidth  'Determines size of smith chart; initialized and later adjusted when resizing
    global graphMarLeft, graphMarRight, graphMarTop, graphMarBot    'margins from graph box edge to grid
    global haltAtEnd    'Flag set to 1 to cause a halt at end of current sweep  'SEWgraph
    global hasMarkPeakPos, hasMarkPeakNeg, hasMarkL, hasMarkR, hasAnyMark       'Marker flags
    dim markerIDs$(9)    'IDs of markers, used to fill combo box. marker numbers run from 1 so ID of marker N is markerIDs$(N-1)
    global selMarkerID$  'ID of marker selected by user
    global doGraphMarkers   'Set/cleared by user to show or hide markers on graph
    global doPeaksBounded   '=1 to limit peak search between L and R markers; otherwise 0
    global doLRRelativeTo$  'marker ID of reference marker when L and R are relative to another marker; otherwise blank
    global doLRRelativeAmount   'db offset when L,R are relative to another marker; otherwise 0
    global doLRAbsolute     '=1 when LR are placed around another marker, but at an absolute db level ver115-3f
    global continueCode 'Checked after "scan" command; 0=continue; 1=halt via [Halted]; 2=immediate wait; 3=restart.
    global axisPrefHandle$  'handle variable for axis preference window; non-blank when window is open
    global graphBox$       'handle variable containing handle for current graph box. ver 114-3c
    global TwoPortGraphBox$     'Handle to graph box for two-port graphs, or blank if no window open ver116-1b
    global displaySweepTime '=1 to display sweep time in message area  'ver114-4f
    global interpolateMarkerClicks  '=1 to enable placing marker at exact click point; =0 to round to nearest step. ver115-1a
    dim axisGraphData$(40) : dim axisDataType(40)   'Graph data for dialog selecting graph types ver116-1b

        'ver114-4k variables allowing forward or reverse sweep
    'The following 3 variables are not global, but are used in connection with reverse sweeps
    'sweepDir     '+1 for left-right sweep; -1 for right-left sweep
    'sweepStartStep, sweepEndStep 'start and end steps for current sweep. Note "startfreq" is the x-axis start (left),
    'but sweepStartStep will be at the x-axis right end if we are going in reverse.
    global alternateSweep    '=1 to alternate forward and reverse sweeps; =0 if direction is set by sweepDir ver114-5a
    global componStopAtEnd  '=1 to stop component Measure and end of measure sweep. ver115-1f
    global RefRLCLastNumPoints,RefRLCLastConnect$   'For continuity calling ReflectionRLC
    global analyzeQLastNumPoints        'For continuity in AnalyzeQ
    global GDLastNumPoints  'Number of points last used for group delay analysis

    global refreshEachScan 'Set/cleared by user to control screen refresh. =1 means refresh each scan.
    global refreshForceRefresh   'Forces refresh at end of scan even though refreshEachScan=0
    global refreshOnHalt     'normally =1 to redraw when we halt. Set to 0 internally for some purposes.

    'The following globals determine whether we redraw various components from scratch or by a faster method.
    global refreshGridDirty    'Forces grid (and labels, title) and setup info to redraw from scratch in RefreshGraphs
    global refreshTracesDirty    'Forces traces to be redrawn from raw Y1 and Y2 values in RefreshGraph
    global refreshMarkersDirty    'Forces recalc of marker positions based on their frequency
    global refreshAutoScale     'Do auto scaling on refresh; also implies refreshRedrawFromScratch ver114-7a
    global refreshRedrawFromScratch    'Forces complete redraw from scratch in RefreshGraph

    global doingInitialization  'Set to 1 during startup initialization of context variables; then to 0 ver114-3f
    global haltsweep    'Set to 1 when scan is running
    'global vna          ' =1 when we are in vna mode delver114-5n
            'Y1DisplayMode, Y2DisplayMode: 0=off  1=NormErase  2=NormStick  3=HistoErase  4=HistoStick
    global Y1DisplayMode, Y2DisplayMode 'Type of phase and mag graphing. Made global for sub use.
    global isStickMode      ' =1 when Y1DisplayMode or Y2DisplayMode are in a "stick" mode
    global specialOneSweep  '=1 when [Restart] is called to do one sweep and then return, rather than wait. ver 114-5f
    global returnBeforeFirstStep    '=1 to initialize on Restart and return to caller before taking any data ver115-1d
    global haltedAfterPartialRestart    '=1 after halting as a result of returnBeforeFirstStep. Set to 0 when scan continues. ver1166-1b
    global msaMode$      '=SA, ScalarTrans, VectorTrans or Reflection
    global menuMode$    'msaMode$ to which the graph window menus currently conform
    global autoScaleY1, autoScaleY2 '=1 to autoscale the axes
    global restartTimeStamp$    'Date/time of last restart
    global primaryAxisNum   '1 or 2, to indicate primary Y axis (e.g. where mag dBm defaults to) ver115-3b

    global Y1DataType, Y2DataType    'data component constant to determine Y1 and Y2 graph data

        'The following component constants are used to specify the data component needed to graph
    global constGraphS11DB, constGraphS11Ang, constTheta
    global constMagDBM,constMagWatts,constMagDB,constMagRatio,constMagV
    global constRho,constAngle,constRawAngle,constGD,constSerReact,constParReact 'ver115-1b
    global constSerR,constSerC,constSerL,constParR,constParC,constParL
    global constSWR, constReturnLoss, constInsertionLoss   'ver114-8d
    global constImpedMag, constImpedAng  'ver115-1d
    global constReflectPower, constComponentQ    'ver115-2d
    global constAdmitMag, constAdmitAng, constConductance, constSusceptance  'ver115-4a
    global constAux0, constAux1, constAux2, constAux3, constAux4, constAux5 'for auxGraphData(,)  ver115-4a
    global constNoGraph 'ver115-2c used to suppress data
        'The following identify slots in ReflectArray() but are not used for graphing. They hold
        'S11 prior to the final adjustment for plane extension and graph R0, and are saved to make
        'it easier to recalculate S11 when those values change.
    global constIntermedS11DB, constIntermedS11Ang   'ver115-2d

        'The following constants are for two-port graphs. They do not have values
        'distinct from those for regular graphs ver116-1b
    global constTwoPortS11DB, constTwoPortS21DB, constTwoPortS12DB, constTwoPortS22DB
    global constTwoPortS11Ang, constTwoPortS21Ang, constTwoPortS12Ang, constTwoPortS22Ang
    global constTwoPortMatchedS21DB, constTwoPortMatchedS12DB 'For/Rev gain dB with impedance-matching terminations ver116-2a
    global constTwoPortMatchedS21Ang, constTwoPortMatchedS12Ang   'angle of gain
    global constTwoPortMatchedS11DB, constTwoPortMatchedS22DB       'For/Rev RL dB with specified terminations ver116-2a
    global constTwoPortMatchedS11Ang, constTwoPortMatchedS22Ang   'angle of return loss
    global constTwoPortKStability, constTwoPortMuStability 'ver116-2a 'Amplifier stability factors


        'The following constant values are used to specify graph data,
        'and assigned in such a way that the first ones can be
        'used as the index into ReflectArray(), which holds a bunch of pre-calculated values for reflection mode.
        'Note if these are changed, FilterDataType may have to be changed. 'ver115-1b altered these
        'We start with 1, because entry 0 of ReflectArray is frequency
    constGraphS11DB=1   'S11 db re S11GraphR0
    constGraphS11Ang=2  'S11 angle re S11GraphR0
    constRho=3      'S11 data with mag made linear
    constImpedMag=4  'impedance magnitude (linear, not db)
    constImpedAng=5 'impedance angle
    constSerR=6     'Series Resistance
    constSerReact =7  'Series Reactance
    constParR=8     'Equiv parallel resistance
    constParReact=9 'Equiv parallel reactance
    constSerC=10     'Equiv series capacitance
    constSerL=11     'Equiv series inductance
    constParC=12     'Equiv parallel capacitance
    constParL=13     'Equiv parallel inductance
    constSWR=14     'SWR
    constIntermedS11DB=15   'Not a graph item ver115-2d
    constIntermedS11Ang=16   'Not a graph item ver115-2d

        'The above are in ReflectArray(); the below are either not reflection related
        'or are other names for reflection data, or simple calculations from that data.
    constMagDBM=17   'Magnitude dBM for SA mode
    constMagWatts=18 'Magnitude watts (linear) for SA mode
    constMagV=19     'Magnitude volts (linear) for SA mode
    constMagDB=20    'Magnitude dB for S21
    constMagRatio=21    'Magnitude ratio (linear) for S21
    constAngle=22    'Angle for S21
    constTheta=23   'Theta for reflection
    constGD=24       'Group Delay for S21
    constReturnLoss=25  'Return Loss for reflection   ver114-8d
    constInsertionLoss=26   'Insertion Loss for S21   ver114-8d
    constReflectPower=27        'Percent reflected power ver115-2d
    constComponentQ=28           'component Q (not resonance Q) ver115-2d
    constAdmitMag=29     'ver115-4a
    constAdmitAng=30
    constConductance=31
    constSusceptance=32
    constRawAngle=33        'Transmission angle without cal adjustment
        'constAuxN represent the data in auxGraphData(step, N). They must be consecutively numbered. ver115-4a
    constAux0=34    'N=0 for accessing auxGraphData
    constAux1=35
    constAux2=36
    constAux3=37
    constAux4=38
    constAux5=39
    constNoGraph=40     'Don't graph and don't label the axis

        'The following can be used directly to index TwoPortArray to retrieve the data
        'Note that the constant for an angle is always one more than that for its corresponding DB
    constTwoPortS11DB=1
    constTwoPortS21DB=3
    constTwoPortS12DB=5
    constTwoPortS22DB=7
    constTwoPortS11Ang=2
    constTwoPortS21Ang=4
    constTwoPortS12Ang=6
    constTwoPortS22Ang=8

        'The following can be used directly (after subtracting constTwoPortMatchedS11DB-1) to index TwoPortMatchedSParam to retrieve the data
        'Note that the constant for an angle is always one more than that for its corresponding DB
    constTwoPortMatchedS11DB=9 'ver116-2a
    constTwoPortMatchedS11Ang=10 'ver116-2a
    constTwoPortMatchedS21DB=11 'ver116-2a
    constTwoPortMatchedS21Ang=12 'ver116-2a
    constTwoPortMatchedS12DB =13 'ver116-2a
    constTwoPortMatchedS12Ang =14 'ver116-2a
    constTwoPortMatchedS22DB=15 'ver116-2a
    constTwoPortMatchedS22Ang=16 'ver116-2a

        'The following do not directly index an array; they need to be calculated
    constTwoPortKStability=17    'ver116-2a
    constTwoPortMuStability=18  'ver116-2a

    dim ReflectArray(800,16) 'Actual signal freq (0),GraphS11DB(1), GraphS11Ang(2),linearMag(3),
                            'Impedance Mag(4), Impedance Angle(5),Rs(6), Xs(7), Rp(8),
                            ' Xp(9), Cs(10), Ls(11), Cp(12), Lp(13), SWR(14),
                            'intermedDB(), intermedAng(15) (intermed= w/o R0 transform or plane ext) 'ver115-2d
    dim uWorkReflectData(16) 'Same data as ReflectArray, but for only one entry  'ver115-1b 'ver115-2d
    dim S21DataArray(800,3) 'Frequency (actual input freq) (0), mag(1), phase(2) and intermed phase(3) for VectorTrans and ScalarTrans modes. ver116-1b
                            'Intermed phase is phase before plane extension, saved in case of recalculation
        'The following auxGraphDataXXX arrays have info on auxiliary graphs, which are numbered 0 to 5; the
        'graph number is the first index of the array. If the graphs are specified as data types constAux0, etc., then
        'the graph number is the constant for the desired graph minus constAux0.
    dim auxGraphData(800, 5)  'Used to hold specially calculated data, such as from Q analysis, which can be retrieved and graphed.
        'auxGraphDataInfo$ Info for UpdateGraphDataFormat for auxGraphData. ver115-4a
        '(,0) is graph Name, (,1) is formatting string, (,2) is axis label, (,3) is marker label
    dim auxGraphDataFormatInfo$(5,3)
        'auxGraphDataInfo is numeric info about each item in auxGraphData. (,0) is 1 if the data is an angle
        '(,1) and (,2) are the axis min and max to use as defaults for graphing.
    dim auxGraphDataInfo(5,2) 'ver115-4a

    'SEWgraph; The following hold parameters used to perform filter analysis when requested by the user
    global doFilterAnalysis     '=1 to perform filter analysis; 0 otherwise SEWgraph
    global x1DBDown, x2DBDown   'positive db values for x1 and x2 points SEWgraph
    global filterPeakMarkID$     'Marker that indicates filter peak

    'SEWgraph The following are used to manage a "working array" in the Utilities Module. This array is used for very
    'temporary processing to deal with the fact that in LB you can't pass arrays as arguments. So some routines,
    'such as saving/retrieving data points to/from strings or files, utilize this specific array. The user is
    'responsible for transferring data in and out of this array. Data can't be left in the array for long, because
    'another operation might utilize the array.
    dim uWorkArray(800,8)   'Initially 800 points with up to 9 (0...8) data items per point
    global uWorkMaxPoints, uWorkNumPoints   'max points and actual used points in uWorkArray
    global uWorkMaxPerPoint, uWorkNumPerPoint 'max and actual number of data items per entry of uWorkArray
    dim uWorkFormats$(8)    'Format strings for each data item, in form suitable for "using" function,
                            'or blank to cause str$() function to be used.
    global maxNumSteps      'Absolute max number of steps allowed for sweep or in arrays of points (num points=num steps+1) ver114-3e
    maxNumSteps=40000       'ver114-3e
    global maxPointExtraLines  'max lines in a file or string with point data, not including the numeric point data itself ver 114-3e
    maxPointExtraLines=100  'ver116-4k
    dim uWorkTitle$(6)      'Title info extracted when processing uTextPointArray$; 0 not used   ver114-3e; ver114-5m increased to 6
    dim uTextPointArray$(maxNumSteps+maxPointExtraLines+5)   'Array of points as string; used as short-term
                                'intermediary in saving/retrieving arrays of points ver114-3e

    dim VNAData(1,2)   'For temporary storage of data to be restored to the graph. Resized when needed. ver116-1b
    global VNADataNumSteps 'Number of steps of data in VNAData. Data runs from (0,x) to (VNADataNumSteps, x)  ver116-1b
    global VNADataLinear    '=1 if the data in VNAData has linear spacing, =0 for log. ver116-4a
    global VNADataZ0        'Reference impedance of data in VNAData ver116-4a
    dim VNADataTitle$(4)    'Title of data in VNAData ver116-1b
    global VNARestoreDoR0AndPlaneExt    'Tells VNARestoreData whether to perform R0 conversion and plane extension ver116-4j
    
    dim contextTypes(30)    'Used in connection with save and retrieve context files
                            'indicating which contexts are involved ver114-3a
    'Constants to specify context types in contextTypes()
    global constHardware, constGrid, constTrace, constSweep, constMarker, constBand, constBase, constGraphData, constModeData
    constHardware=0 : constGrid=1 : constTrace=2 : constSweep=3 : constMarker=4 : constBand=5
    constBase=6 : constGraphData=7 : constModeData=8

    'Each entry of calMagCoeffTable() will have 8 numbers. 0-3 are the A,B,C,D
    'coefficients for interpolating the real part; 4-7 are for the imag part.
    dim calMagCoeffTable(100,7)    'SEWcal Cubic coefficients for interpolating in calMagTable moved here by ver115-1d
     'calFreqCoeffTable has the A,B,C,D coefficients for interpolating in the calFreqTable()
    dim calFreqCoeffTable(800,3)    'SEWcal Cubic interpolation coefficients for calFreqTable()  moved here by ver115-1d

        '---Variables relating to touchstone data files--- ver115-5f
    global touchFreq$, touchForm$, touchRef 'Touchstone format specs
    global touchFreqMult     'Frequency multiplier corresponding to touchFreq$. E.g. 1e6 for MHZ.
    global touchBadLine     'Line number of Touchstone file error; zero if no error
    global dataLoadLastFolder$  'Folder from which parameter data was last loaded
    global touchMaxData     'Max data lines in our touchstone files when reading
    touchMaxData=10000  'ver116-1b
    global touchLastFolder$    'Last folder accessed for save or load of parameter files.
    global touchSParamType$   'Type of params in sParam: "", "Reflection", Transmission" or "TwoPort"

    dim touchComments$(10)   'Up to 10 lines of comments from parameter file. Zero entry not used
    global touchCommentCount    'Number of comment lines in touchComments$ 
    '---End variables for data files----

 'Additional variables that need to be available to subroutines
        'Note "startfreq" is the x-axis start (left),
        'but sweepStartStep will be at the x-axis right end if we are going in reverse.
        'baseFrequency  is added when commanding the hardware, but does not affect graph display or file frequencies.
    global startfreq, endfreq, baseFrequency 'ver116-4k
    global wate, planeadj
    global FreqMode '1, 2 or 3, indicating modes 1G, 2G and 3G  ver115-1c
    global sgout   'signal generator output freq when in plain SA mode
    global test     'SEWgraph  Contents get printed to message box on halt
    global spurcheck    '=1 to turn spur test on 'ver114-4f
    global gentrk   '=1 when TG is being used (depends on mode); 0 if SG is used or build does not have TG hardware
    global normrev  '=0 if TG is normal, =1 if TG is in reverse.
    global offset  'TG offset frequency  moved ver115-5f
    global path$    'Currently active filter path (1...); a number as a string in form "Path N" ver114-1d
    global FiltA0, FiltA1   'low and high bits of RBW filter address 'ver116-4j
    dim fileInfo$(1,1)      'array for use with Files() command ver114-3f
    global message$         'Message to print in graph window
    global varwindow        '=1 if variables window is open ver115-1a
    global suppressPDMInversion '=1 to suppress inversion in [ReadStep] ver115-1a
    global leftstep     'Used to hold a marker step number for [preupdatevar] ver115-1a
    global userFreqPref '0 if user last used Center/Span method; 1 if user last used Start/Stop ver115-1d

         'When loading a path calibration file, we determine lowest ADC value for magdata for which phase is considered valid
    global validPhaseThreshold 'ver116-1b

    '------------Items for auto wait time----------'ver116-1b
         'When loading a path calibration file, we divide the response into three sections and calculate approximate slopes:
        'For ADC<calLowADCofCenterSlope the slope is calLowEndSlope, calculated from a segment near the
        '       low end, but not so low as to have tiny slopes compared to the center slope
        'for calLowADCofCenterSlope < ADC < calHighADCofCenterSlope the slope is calCenterSlope, calculated
        '        from a 20 dB or greater segment somewhere in the center area
        'for ADC>calHighADCofCenterSlope the slope is calHighEndSlope calculated from a segment at top end
        'The slope is delta ADC/delta dB
    global calCanUseAutoWait    '=1 if cal table is suitable for auto wait time calculations
    global useAutoWait   '=1 if user specified to use auto wait times
    global autoWaitPrecision$ '"Fast", "Normal" or "Precise"  meaningless if useAutoWait=0
    global calADCofLowFringe    'ADC below which the ADC/dB slope is tiny
    global calLowADCofCenterSlope, calHighADCofCenterSlope
    global calLowEndSlope, calCenterSlope, calHighEndSlope
    global autoWaitTC    'max time constant for video filters of mag and phase (where applicable) ver116-4j
        'The max db error is converted to a max ADC error for the three regions
        'of the path calibration table.
        'With auto wait time, we do repeated readings after waiting a certain time,
        'and determine when the change in readings falls below a maxChange level
        'The maximum allowed change in ADC values, to keep the error at the allowed limit,
    global autoWaitMaxChangeLowEndADC, autoWaitMaxChangeCenterADC, autoWaitMaxChangeHighEndADC
    global autoWaitMaxChangePhaseADC
    '------------End items for auto wait time----------

    '------------Items for transitions between Reflection and VectorTrans modes ver116-1b--------
        'Sweep parameters of last transmission sweep
    global transLastSteps, transLastStartFreq, transLastEndFreq, transLastIsLinear, transLastGraphR0
        'most recent Y axis settings for transmission
    global transLastY1Type, transLastY1Top, transLastY1Bot, transLastY1AutoScale  'data type and axis top and bottom
    global transLastY2Type, transLastY2Top, transLastY2Bot, transLastY2AutoScale
        'Sweep parameters of last reflection sweep
    global refLastSteps, refLastStartFreq, refLastEndFreq, refLastIsLinear, refLastGraphR0
        'most recent Y axis settings for reflection
    global refLastY1Type, refLastY1Top, refLastY1Bot, refLastY1AutoScale
    global refLastY2Type, refLastY2Top, refLastY2Bot, refLastY2AutoScale
        'most recent titles
    dim refLastTitle$(4), transLastTitle$(4)
    '-------------------------------------------------------------------------------------------

'Variables for saving/restoring context files via gosub routines, where these are used as parameters
    global restoreFileName$, restoreFileHndl$, restoreContext$, restoreIsValidation, restoreErr$,restoreLastLineNum 'ver115-8c

        'There can be 1 to four video filter settings, with different capacitor values.   'ver116-1b
        'They each have a name. Mag and phase capacitors can be different, but the names and number
        'of filters are the same. The names must be Wide, Mid, Narrow or XNarrow, but not all those names must be used.
        'Names and capacitor values are set in hardware configuration.
    global videoFilter$     'Selected video filter: Wide, Mid, Narrow or XNarrow  'ver116-1b
    dim videoFilterCaps(4,1)  'Capacitance(uf) for Wide(1), Mid(2), Narrow(3) and XNarrow(4) video filters   'ver116-1b
                               'Second index is 0 for magnitude and 1 for phase filters.
    dim videoFilterNames$(4) 'Names of each video filter, or blank if no filter. Index matches videoFilterCaps   'ver116-1b
    global videoFilterAddress, videoMagCap, videoPhaseCap 'current video filter address (0-3) and cap values (uF) for mag and phase ver116-1b
    global videoMagTC, videoPhaseTC 'Time constants (ms) of video mag and phase filters 'ver116-1b
    global switchFR, switchTR   'current or desired state of forward/reverse (0=forward) and transmission/reflection (0=transmission) switches ver116-1b

    'Globals used to remember state info to allow detection of user changes; added by ver114-6e
    'See RememberState and DetectChanges
    global prevMSAMode$     'msaMode$
    global prevPath$        'Filter path. ver115-1a
    global prevFreqMode     'frequency mode ver115-1c
    global prevStartF, prevEndF, prevBaseF    'ver116-4k
    global prevXIsLinear, prevY1IsLinear, prevY2IsLinear
    global prevSteps, prevSweepDir, prevAlternate
    global prevStartY1, prevEndY1, prevStartY2, prevEndY2
    global prevHorDiv, prevVertDiv
    global prevY1Disp,prevY2Disp
    global prevGenTrk, prevSpurCheck    'ver114-6k
    global prevPath$    'ver116-4j
    global prevTGOff, prevSGFreq    'ver115-1a
    global prevPlaneAdj     'ver114-7f
    global prevY1DataType, prevY2DataType   'ver115-1b deleted source constants
    global prevAutoScaleY1, prevAutoScaleY2 'ver114-7a
    global prevDataChanged      'This must be set to 1 when data is loaded from a context
    global prevS21JigAttach$  '"Series" or "Shunt" to indicate the Transmission jig used ver114-6k
    global prevS21JigR0   'Source and load impedances of Transmission
    global prevS21JigShuntDelay     'Delay of shunt fixture connector
    global prevS11BridgeR0, prevS11GraphR0   'Bridge reference and graph reference for S11   'ver114-6k
    global prevS11JigType$      'Jig previously used for reflection mode  ver115-1b
    global prevSwitchFR     'Previous setting of forward/reverse switch ver116-1b

    'Functions may need to redo the sweep with new parameters. The following are used to save/restore the
    'pre-existing parameters, using SaveAndChangeSweepParameters and [RestoreSweepParameters] ver115-5c
    global functSaveAlternate,functSaveSweepDir,functSavePlaneAdj,functSaveWate, functSaveVideoFilter$  'ver116-4b
    global functSaveAutoWait, functSaveAutoWaitPrecision$  'ver116-1b
    global functSaveSteps, functSaveStartFreq, functSaveEndFreq
    global functSaveAutoY1, functSaveAutoY2
    global functSaveY1Mode, functSaveY1Mode
    global functSaveY1DataType, functSaveY2DataType
    global functSaveXIsLinear, functSaveY1IsLinear, functSaveY2IsLinear
    global functSaveNumHorDiv, functSaveNumVertDiv
    global functSaveDesiredCalLevel

    'ver114-5e added these cal items. Full Line cal is a calibration of through response with the current sweep settings.
    'Baseline cal is a cal of through response with a generic wideband sweep.
    'desiredCalLevel is the user-specified level. applyCalLevel is the level we are actually applying
    global desiredCalLevel, applyCalLevel    '2=band Cal (if exists); 1=Base Cal(if exists); 0=None
            'We keep track of the parameters for the last cal. Number of steps is set to -1 to indicate no valid cal
        'Following are Full Line cal sweep params ver114-5e
    global bandLineStartFreq, bandLineEndFreq, bandLineNumSteps, bandLineLinear, bandLinePath$, bandLineTimeStamp$
        'Following are Base Line cal sweep params ver114-5e
    global baseLineStartFreq, baseLineEndFreq, baseLineNumSteps, baseLineLinear, baseLinePath$, baseLineTimeStamp$

    global baseLineS21JigAttach$, bandLineS21JigAttach$     'S21JigAttach$ for applicable cal 'ver115-1B
    global baseLineS21JigR0, bandLineS21JigR0               'S21 jig R0 for applicable cal ver115-1b

    global installedBandLineTimeStamp$  'Time stamp of the band cal that was installed; not when it was installed ver115-2d
    'Following are sweep params at which Base Line Cal was last installed ver114-5f
    global installedBaseLineStartFreq, installedBaseLineEndFreq, installedBaseLineNumSteps
    global installedBaseLineLinear
    global installedBaseLineTimeStamp$ 'Time stamp of the base cal that was installed; not when it was installed ver115-2d
    dim lineCalArray(800,2)    'calibration data for each step#: (0)freq (tuning),(1)magpower during cal,(2)phaseofpdm during cal
    dim bandLineCal(800,2) 'Bandsweep line calibration data; transferred to lineCalArray when needed ver 114-5f
        'Note baseLineCal is fixed sized
    dim baseLineCal(2000,2) 'Baseline line calibration data; transferred to lineCalArray when needed ver 114-5f

'Impedance can be measured via S21 in a test jig, or via S11 in a reflection bridge
        'We need to know the reference impedance (a resistance) of each, and for the jig
        'we need to know if the DUT is in series, or shunted to ground. The S21 jig is actually used only in reflection mode.
        'ver114-6g added these

    global S21JigAttach$  '"Series" or "Shunt" to indicate the Transmission jig used ver114-6k
    global S21JigR0   'Source and load impedances of Transmission jig
    global S21JigShuntDelay     'One-way delay of connector from shunt fixture through line to the DUT, in ns. ver115-1e
    global S11BridgeR0, S11GraphR0   'Bridge reference and graph reference for S11   'ver114-6k
    global S11JigType$      '="Trans" if S21Jig is being used for reflection mode; "Reflect" if reflection bridge is used ver115-1b

    global lineCalThroughDelay  'Delay (ns) of line cal through connection. Used for continuity between lineCal dialog sessions
                                'and to inform [BandLineCal]. need not be saved as a preference item.

    'ver114-7f added these dialog variables
    global DialogCancelled, DialogRLCConnect$      'Used to pass values to/from some dialogs
    global DialogRValue, DialogLValue, DialogCValue    'Used to pass values to/from some dialogs
    global DialogQLValue, DialogQCValue     'ver115-4b
    global DialogCoaxSpecs$, DialogCoaxName$ 'Used to pass values to/from some dialogs  ver115-4a

    global referenceLineSpec$, referenceLineType    'Spec and type of reference line    ver114-7f
                                                'type: 0=none; 1=use data when ref was selected; 2=use RLC in spec; 3=use fixed value ver115-5d
    global referenceTrace       'which ref lines, based on bits: 1 bit=do trace 1; 2 bit=do trace 2; 4 bit=do Smith trace   'ver115-6b
    global referenceSourceNumPoints     'Number of valid points in referenceSource()
    dim referenceSource(801,2)  'freq(0), db(1) and angle() of source data for reference lines. First entry is 1
    dim referenceTransform(801,2)  'Actual graph data for reference lines  Entries are by point number, 1... (not step num)
    global referenceColor1$, referenceWidth1,referenceColor2$, referenceWidth2    'Reference Trace color and width
    global referenceColorSmith$, referenceWidthSmith     'color/width for smith chart reference line

     'The reference line may be graphed or combined with the scan data.
        'The math transform is in the form A*Ref + B*Data, A and B generally being 1 or -1
        'Math is done if referenceDoMath=1 or 2; otherwise reference is just graphed.
        'If referenceDoMath=1 then math is done on the dBm values that would be graphed with the default graphs.
        'referenceDoMath=1 is allowed only for SA mode; too complicated and not very useful for VNA modes, which use calibration instead.
        'If referenceDoMath=2 then math is done on the current graph values (e.g. capacitance)
    global referenceDoMath, referenceOpA, referenceOpB 'ver114-7b
    dim setupList$(15)  'List of current test setups; used in test setup dialog only
    '-----OSL cal variables-----
        'ver115-1b added the following OSL items
    global OSLdoneO, OSLdoneS, OSLdoneL 'For communicating with [PerformOSLCal]
        'These items are the currently active data for reflection
    dim OSLa(800,1), OSLb(800,1),OSLc(800,1)    'OSL coefficients a, b, c, at same frequencies as ReflectArray() ver115-1b
    global OSLRefType$  'Open, Short or Load ver116-4n
    global OSLcalLastUsedFull   '=1 if cal window last used full OSL; 0 if last used reference.
    'global OSLApplyFull 'delver116-4n We now always apply OSL in reflection mode if applyCalLevel=0.

        'These items are used only during the OSL calibration procedure to calculate coefficients.
    dim OSLstdOpen(800,1)       'Actual refco of open standard at each step, real/imaginary
    dim OSLstdLoad(800,1)       'Actual refco of load standard at each step, real/imaginary
    dim OSLstdShort(800,1)       'Actual refco of short standard at each step, real/imaginary
    dim OSLcalOpen(800,1)       'Measured (during cal) open at each step, db/angle converted to real/imag in ProcessOSLCal
    dim OSLcalLoad(800,1)       'Measured (during cal) load at each step, db/angle converted to real/imag in ProcessOSLCal
    dim OSLcalShort(800,1)      'Measured (during cal) short at each step,  db/angle converted to real/imag in ProcessOSLCal
    global OSLLastSelectedCalSet  'Standard cal set selected last time dialog was open
    global OSLOpenSpec$, OSLShortSpec$, OSLLoadSpec$    'RLC specs for currently selected OSL standards ver116-4i
    global OSLFileOpenSpec$, OSLFileShortSpec$, OSLFileLoadSpec$    'RLC specs for OSL standards, used to transfer to and from files. ver116-4i
    global OSLFileCalSetName$, OSLFileCalSetDescription$   'ver116-4i

    dim OSLCalSetNames$(10)  'List of OSL cal sets; zero entry is used. Entries correspond to OSLCalSetFileNames   ver115-7a
    dim OSLCalSetFileNames$(10)  'List of OSL cal set descriptions (long); zero entry is used. redim'd as necessary  ver115-7a
    global OSLCalSetNumber     'Number of entries in list of cal set names and file names   ver115-7a

        'These items are the result of calibration and allow either base or band cal to be installed
        'The OSLBasex arrays are fixed size; the others are expanded as necessary
    dim OSLBaseA(2000,1), OSLBaseB(2000,1), OSLBaseC(2000,1)       'Base coefficients;  used to fill OSLx()
    dim OSLBandA(800,1), OSLBandB(800,1), OSLBandC(800,1)       'Band coefficients (matches current frequencies)
    dim OSLBaseRef(2000,2), OSLBandRef(800,2)       'Freq(0) and dB(1) and angle(2) update info. Ref is 0 until cal update is run.
    global OSLBandRefType$, OSLBaseRefType$ 'Open, Short or Load ver116-4n

        'We keep track of the parameters for the last cal. Number of steps is set to -1 to indicate no valid cal
    global OSLError '=1 if math error occurred in calculating OSL coeff; used to nullify the cal ver115-4j
    global OSLBaseStartFreq, OSLBaseEndFreq, OSLBaseNumSteps, OSLBaseLinear, OSLBasePath$   'Params for last base OSL calibration
    global OSLBandStartFreq, OSLBandEndFreq, OSLBandNumSteps, OSLBandLinear, OSLBandPath$   'Params for last band OSL calibration
    global OSLBandS11JigType$, OSLBaseS11JigType$   'Jig type--"Reflect" or "Trans"-- for last cal
    global OSLBaseS21JigAttach$, OSLBandS21JigAttach$   'S21 jig for last OSL cal; relevant only if jig type is "Trans"
    global OSLBaseS11BridgeR0, OSLBandS11BridgeR0   'Bridge R0 for last OSL cal; relevant only if jig type is "Reflect"
    global OSLBaseS21JigR0, OSLBandS21JigR0     'S21 jig R0 for last OSL cal; relevant only if jig type is "Trans"
    global OSLBaseTimeStamp$   'Time stamp for last base open, load, short cal.
    global OSLBandTimeStamp$   'Time stamp for last band open, load, short cal.
    global installedOSLBandTimeStamp$   'Time stamp of the installed band cal; not when it was installed ver115-2d

            'Following are sweep params at which base OSL cal coefficients were last installed
            'We don't need these for installed Band cal, because they would all match the sweep params
            'at the time of installation.
    global installedOSLBaseStartFreq, installedOSLBaseEndFreq, installedOSLBaseNumSteps
    global installedOSLBaseLinear, installedOSLBaseRefType$ 'ver116-4n
    global installedOSLBaseTimeStamp$   'Time stamp of the installed base cal; not when it was installed ver115-2d
    '-------End OSL cal variables-------'


    global maxCoaxEntries, numCoaxEntries   'Maximum and actual number of entries in coax data arrays
    maxCoaxEntries=100
    dim coaxNames$(maxCoaxEntries)  'Coax names; 0 entry not used
    dim coaxData(maxCoaxEntries, 4) 'Coax data: R0(,1), VF(,2), K1(,3), K2(,4)
    dim RLCDialogCoaxTypes$(maxCoaxEntries+10)  'For RLC specification dialog ver115-4a

    global calInProgress  'variable set to 1 before starting a cal sweep, then to 0 when done

        'ver114-6b moved the following dim statements here
        'SEWgraph; the following arrays may be expanded in ResizeArrays to accomodate more steps
        'SEWgraph Pixel values are no longer kept in these arrays, so references to thispointx, thispointmag, thispointphase,
        'oldmagpixwl and oldphapixel should be ignored. Eventually, the arrays could be compacted to eliminate those unused slots.
    global constMaxValue
    constMaxValue=1e12   'Max value for RLC components and certain calculations. ver115-1b
        'Data is put into datatable() point by point as it is gathered. Its frequency is the 0-1 GHz "equivalent 1G frequency".
        'In VectorTrans and ScalarTrans modes, the data is also saved to S21DataArray() with the actual sweep frequency,
        'and if reflection mode convert to S11 and save to ReflectArray()
    dim datatable(800,3)   'data from most current sweep, (0)thisstep,(1)thisfreq(hardware freq),(2)processed magpower,(3)processed phase
        'Note: magarray(x,1) is no longer used; magarray(x,2) is never actually used to store magnitude ver116-1b
    dim magarray(800,3)    'magni pixels for each step#: (0)thispointx, (1)oldmagpixel,(2)thispointmag(3)magdata
        'Note: phaarray(x,1) is no longer used; phaarray(x,2) is never actually used to store phase ver116-1b
    dim phaarray(800,4)    '(0)pdmcmd; phase pixels for each step#:(1)oldphapixel,(2)thispointphase,(3)phadata,(4)pdmread ver111-39d

    dim PLL1array(800,48)  '(0-23)N23thruN0,(24-39)notused,(40)pdf1,(43)LO1freq,(45)ncounter,(46)Fcounter,(47)Acounter,(48)Bcounter. ver111-30a
    dim PLL3array(800,48)  '(0-23)N23thruN0,(24-39)notused,(40)pdf3,(43)LO3freq,(45)ncounter,(46)Fcounter,(47)Acounter,(48)Bcounter. ver111-30a
    dim DDS1array(800,46)  '(0-39)sw0-sw39,(40-44)w0-w4,(45)base,(46)actualdds1output
    dim DDS3array(800,46)  '(0-39)sw0-sw39,(40-44)w0-w4,(45)base,(46)actualdds3output
    dim freqCorrection(800) 'freq correction factors for frequency of each step in current sweep; added to raw data
        'frontEndCalData is the raw data, freq and dBm, for the current front end. It is interpolated to frontEndCorrection
        'on Restart to match the current scan points. It is resized if necessary when a front end file is loaded
        'first index is 1-based. For second index: 0=freq; 1=dBm ver115-9d
        'The data is the value to be subtracted from the raw power readings. Subtraction is used so a front end file
        'can be created by a transmission measurement of the front end, after calibrating with a through connection.
    dim frontEndCalData(800,1)
    dim frontEndCorrection(800) 'correction for front end in use for each step in current sweep; subtracted from raw data
    global frontEndCalNumPoints 'Number of valid points in frontEndCalData
    global frontEndActiveFilePath$  'Path name for active front end file. Only relevant in SA modes.ver115-9c
    global frontEndLastFolder$  'path to Last folder from which front end was loaded

            'ver114-2d combined config arrays into one, and deleted configarray
    dim cmdallarray(800,39) '(0-15)DDS1+DDS3, (16-39)PLL1+DDS1+PLL3+DDS3

    global suppressHardware '=1 to suppress hardware operations, otherwise 0 'ver115-6c
    global suppressHardwareInitOnRestart    '=1 to skip hardware re-initialization to speed up Restart or [PartialRestart].
    'ver116-4d deleted wantHardwareInitOnPartialRestart

    '-------Items for multi-scanning----------- ver115-8c
    'In SA mode only, it is possible to rotate through several different scan settings; this
    'is called multi-scanning. One sweep is completed with one setting, then we move to the next.
    'When one is completed, its graphic is displayed in its own window.
    'The settings are saved, as is the datatable info. The datatable info for a scan can be reloaded
    'into the main window when scanning is stopped.
    'Multiscan entries are numbered from 1 through multiscanMaxNum
    'Entry zero is info for the main graph.
    global multiscanCurrNum     'Current entry number
    global multiscanMaxNum      'Maximum entry number
    global multiscanIsOpen   '=1 when multiscan windows are open, even if scan not in progress.
    global multiscanInProgress    '=1 when actually scanning in multiscan mode
    global multiscanHaltAtEnd   'set to non-zero during multiscan to cause halt at end, value depends on window selected
    global multiscanSaveRefreshEachScan 'Saves for restoration when quitting multiscan

    multiscanMaxNum=4   'This is fixed
        'This array holds the context strings for the multiscan entries
        'Blank for grid context means this entry is not to be used
        'Zero entry of first index not used
    dim multiscanContexts$(multiscanMaxNum, 3) 'second index= grid(0), trace(1), sweep(2), marker(3)
        'multiscanDBM holds the magnitude in dBm for each multiscan entry. Main graph is entry 0.
        '2000 steps are allowed for each entry, so the second index for entry N
        'runs from 2001*N to 2001*(N+1)-1
        'multiscanRefData holds reference data arranged similarly
    dim multiscanTitles$(multiscanMaxNum, 4)    'Four title lines for each multiscan entry ver115-9a
    dim multiscanDBM((multiscanMaxNum+1)*2001)
        'Reference data is always to be added/subtracted with graph data, based on the math type
    dim multiscanRefData((multiscanMaxNum+1)*2001)
        'The only reference type allowed in multiscans is reference math with data: the reference
        'data and current data are added or subtracted based on two values OpA and OpB, which are each
        '1 or -1 (see referenceDoMath), but are zero if reference is not to be used.
        'OpA for entry N is at 2*N. OpB is at 2*N+1.
    dim multiscanRefOps((multiscanMaxNum+1)*2)
    dim multiscanGraphHandles$(multiscanMaxNum)     'LB handles to the graphics box for each multiscan; blank if inactive
    dim multiscanWindowHandlesLB$(multiscanMaxNum)    'LB handles to windows for each multiscan; blank if inactive
    dim multiscanWindowHandlesWind(multiscanMaxNum)    'Windows handles to windows for each multiscan
    global multiscanMainWidth, multiscanMainHeight  'Dimensions of main graph box when multiscan was started
    global multiscanWindowsPosition, multiscanSettingsPosition  'positions of multiscan windows that need to be hidden/shown
    global multiscanControlPosition
    multiscanControlPosition=0
    multiscanWindowsPosition=1 'Windows menu is number 1, indexed from 0
    multiscanSettingsPosition=2 'Settings menu is number 2, indexed from 0
    dim multiscanMenuBarHandles(multiscanMaxNum)   'Windows handles to menu bar for each scan
    dim multiscanWindowsMenuHandles(multiscanMaxNum)   'Windows handles to "Windows" menu for each scan
    dim multiscanSettingsMenuHandles(multiscanMaxNum)   'Windows handles to "Settings" menu for each scan
    dim multiscanControlMenuHandles(multiscanMaxNum)    'Windows handles to "Control" menu for each scan
    dim multiscanSkip(multiscanMaxNum)  '=1 to skip this graph when scanning

    '-------End Items for multi-scanning-----------

    '--------Items for auto switches-----------
    'Indicates whether we have software-controlled switches: RBW filter, video filter, xG Band switch, Transmit/Reflect, and Forward/Reverse
    global switchHasRBW, switchHasVideo, switchHasBand, switchHasTR, switchHasFR
    '--------End items for auto switches-----------


    '------Start items for USB interface----------
     ' USB interface dll 'USB:01-08-2010
    ' this is used as a form of handle used by the USB interface
    ' it is necessary because of the limitations of liberty basic
    ' and it is actually a memory pointer to a USB device class object
    ' we set it to zero when the interface is not initialised
    global USBdevice 'USB:01-08-2010
    ' this is a boolean flag used to control when the USB interfcae dll is open (I hate basic)
    global UsbInterfaceOpen 'USB:01-08-2010
    ' set this flag to != 0 for usb interface active
    global bUseUsb 'USB:01-08-2010
    global bUsbAvailable 'USB:01-08-2010
    ' these are string buffers used in USB I/O operations
    global USBwrbuf$ 'USB:01-08-2010
    ' used in ADC input functions; the number of ADC readings made and the results read
    global UsbAdcCount 'USB:01-08-2010
    global UsbAdcResult1 'USB:01-08-2010
    global UsbAdcResult2 'USB:01-08-2010
    USBdevice = 0 'USB:01-08-2010
    UsbInterfaceOpen = 0 'USB:01-08-2010
    bUseUsb = 0 'USB:01-08-2010
    struct USBrBuf, numreads as ulong, magnitude as ulong , phase as ulong 'USB:01-08-2010
    ' the next 3 are used to create a memory buffer for the dll to use to save scan time
    ' it contains a char[][] version of cmdallarray
    global AllArrayBlockSize ' current size of memory block allocated for AllArrays
    global hSAllArray ' handle for memory block
    global ptrSAllArray ' pointer to memory block

    global hSDDS1Array 'USB:05/12/2010
    global ptrSDDS1Array 'USB:05/12/2010
    global hSDDS3Array 'USB:05/12/2010
    global ptrSDDS3Array 'USB:05/12/2010
    global hSPLL1Array 'USB:05/12/2010
    global ptrSPLL1Array 'USB:05/12/2010
    global hSPLL3Array 'USB:05/12/2010
    global ptrSPLL3Array 'USB:05/12/2010
    struct Int64N, msLong as ulong, lsLong as ulong ' USB:15/08/10
    struct Int64SW, msLong as ulong, lsLong as ulong ' USB:15/08/10
    ' This structure is used to minimize time spent forming commands to send to the USB DLL
    ' the values correspond to the parameters of the same name (or similar description) in the parallel function
    ' we aim to try not to set them each time to save processing time in basic code
    struct UsbAllSlimsAndLoadData, thisstep as short, filtbank as short, latches as short, pdmcommand as short, pdmcmdmult as short, pdmcmdadd as short ' USB:15/08/10
    ' This struct is used to control ADC reading configuration. It is set to define the paramaters
    ' and save liberty basic from having to fill it in each time
    ' The parameters map to the parameters in the ADC convert commands in fw-msa
    ' Adcs takes values 1 or 3 normally ( 1 = magnitude ADC only, 3 = both ) but will also do 2 (phase only)
    ' CLocking is the clocking option - 0 for AD7685 and 1 for LT1860
    ' Delay is the ADC Convert clock high time delay - 2 for AD7684, 4 for LT1960
    ' Bits is the number of data bits - 16 for AD7685, 10 for LT1860
    ' Average is normally 1 - set to higher number if you want the interface to do multiple readings and average
    struct UsbAdcControl, Adcs as short, Clocking as short, Delay as short, Bits as short, Average as short ' USB:15/08/10
    '---------End items for USB interface------------

'---------------END OF VARIABLES DECLARATIONS-------

    nomainwin
    Open "OLE32.dll" for dll as #DLL.OLE 'ver116-4q
    CallDLL #DLL.OLE, "CoInitialize", 0 as ULong, res as ULong  'To avoid tooltips crash in file dialog ver116-4q
    
        'Suppress parallel port if we don't have the DLLs
    if uVerifyDLL("ntport") then suppressHardware=0 else suppressHardware=1 'may change when we have cb info

    if uVerifyDLL("msadll") then bUsbAvailable = 1 else bUsbAvailable = 0 'USB:01-08-2010
    ' attempt to initialise the USB interface. Can be called even when open
    if bUsbAvailable then call UsbOpenInterface 'USB:01-08-2010
    if USBdevice <> 0 then 'USB:16-08-2010
        CALLDLL #USB, "UsbMSAGetVersions", USBdevice as long, result as ushort 'USB:16-08-2010
        if int( result / 256) <  2 then 'USB:16-08-2010
            call UsbCloseInterface 'USB:16-08-2010
            notice "The version number of msadll is too old for me to use" 'USB:05/12/2010
        end if 'USB:05/12/2010
        if int( result and 255 ) <  36 then 'USB:16-08-2010
            call UsbCloseInterface 'USB:16-08-2010
            notice "The USB interface is either not plugged in or is too old a version for me" 'USB:05/12/2010
        end if 'USB:05/12/2010
    end if 'USB:16-08-2010

    ' needs to be more complex - this flag uses USB if it is available
    ' do not set it if bUsbAvailable is not set
    ' create a memory block for holding dds data (akin to cmdallarray[][] but accessible within dll)
    AllArrayBlockSize = 801*40 'USB:01-08-2010
    DeviceArrayBlockSize = 801 * 8 'USB:06-08-2010
    hSAllArray = GlobalAlloc( AllArrayBlockSize ) 'USB:01-08-2010
    ptrSAllArray = GlobalLock( hSAllArray ) 'USB:01-08-2010
    hSDDS1Array = GlobalAlloc( DeviceArrayBlockSize ) 'USB:06-08-2010
    ptrSDDS1Array = GlobalLock( hSDDS1Array ) 'USB:06-08-2010
    hSDDS3Array = GlobalAlloc( DeviceArrayBlockSize ) 'USB:06-08-2010
    ptrSDDS3Array = GlobalLock( hSDDS3Array ) 'USB:06-08-2010
    hSPLL1Array = GlobalAlloc( DeviceArrayBlockSize ) 'USB:06-08-2010
    ptrSPLL1Array = GlobalLock( hSPLL1Array ) 'USB:06-08-2010
    hSPLL3Array = GlobalAlloc( DeviceArrayBlockSize ) 'USB:06-08-2010
    ptrSPLL3Array = GlobalLock( hSPLL3Array ) 'USB:06-08-2010
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceSetAllArrayPtr", USBdevice as long, ptrSAllArray as long, 800 as short, 40 as short, result as boolean 'USB:01-08-2010

    interpolateMarkerClicks=0    'Note user has no way to change this    ver115-1a
    steps = 400
    globalSteps=400 'SEWgraph
    call gSetNumDynamicSteps steps      'ver114-1f
    graphBox$=""  'ver115-1a

    markerIDs$(0)="1" : markerIDs$(1)="2" : markerIDs$(2)="3" : markerIDs$(3)="4"  'SEWgraph
    markerIDs$(4)="5" : markerIDs$(5)="6" : markerIDs$(6)="L" : markerIDs$(7)="R"  'SEWgraph
    markerIDs$(8)="P+" : markerIDs$(9)="P-"   'SEWgraph

        '-------Initialize Modules---------
    call uInitFirstUse      'Initialize Utilities Module
    call configInitFirstUse 'Initialize Configuration Module
    'call calInitFirstUse 201, 1001, hasVNA    'delver116-1b This is done a few lines below

    '----------------Load Configuration File---------
    if configFileExists()=0 then
        dum=configRunManager(1)  'Lets user enter configuration data, and saves to file
                            '1 signals we are running on startup so no cancellation is allowed
                            'Will quit program when finished
    else
        errStr$=configLoadData$() 'Initializes the configuration globals from the configuration file
        if errStr$<>"" then     'errStr$ is blank if no error occurred; otherwise it describes the error
            notice "Configuration File Error; "+errStr$;"; Default values used"
            call configInitializeDefaults   'Load default values because of file error
        else
                'Be sure we have the Wide filter. This line is needed for smooth transition from
                'ver115 to ver116. ver1116-4f
            if VideoFilterNames$(1)<>"Wide" then VideoFilterNames$(1)="Wide" : VideoFilterCaps(1,0)=0.002 : VideoFilterCaps(1,1)=0.011
            call configSaveFile 'Save config file in current format ver114-5i
        end if
    end if
        'Convert Configuration Globals into Locals
    port=globalPort 'SEW5 add ver113-7c
    glitchtime=0 'SEW5 add ver113-7c; ver114-5k deleted globalGlitchtime
    status=port+1 'SEW5 add ver113-7c
    control=port+2 'SEW5 add ver113-7c

    if cb=3 then suppressHardware=0  '3 means USB. suppressHardware relates only to parallel port ver116-4b
        'An initial low on the PS bit that controls latched switches may be draining switch capacitors
        'Set it high and allow some recharge time. At this point we don't care what the switches get set to,
        'and if the capacitors are discharging the switches won't change state at all.
    if suppressHardware=0 then call SelectLatchedSwitches   'ver116-4h
    call uSleep 1000    'Wait 1 second for capacitor recharge. There will be additional software delays before PS is used.

        'The following was moved here by ver115-1a so that hasVNA is set first
    call calInitFirstUse 201, 1001, hasVNA    'Initialize Mag/Freq Calibration Module--201 max mag cal points; 1001 max freq cal points ver114-4b
        'ResizeArrays needs TGtop, so we do it after loading config file
    call ResizeArrays 2001    'Make all arrays big enough for 2001 points; also loads BaseLineCal file   'ver114-5m

    '---------Load path and freq calibration info------
    call calInstallFile 0   'Loads frequency calibration file; creates one if necessary
    for i=MSANumFilters to 1 step -1
        'For each filter, create the file if necessary and load it
        'Each one loaded replaces the data from the previous one. We are just
        'trying to be sure they exist and are OK.
        'We do this in reverse order to path 1 will be the last one and stays in place
        'This also sets finalfreq and finalbw
        call calInstallFile i
    next i
    path$="Path 1"
    'Note physical selection of filter 1 is done in step 5 below
    for i=1 to MSANumFilters
        'For each filter, combine freq and bw into nicely aligned string. Used to load #main.FiltList
        MSAFiltStrings$(i-1)="P"+str$(i)+"   "+configFormatFilter$(MSAFilters(i,0), MSAFilters(i,1)) 'ver113-7c
    next i

        'The below are not actually the desired states. See step 3 (initialization) for explanation.
        'The latched filter addresses will be asserted by SelectVideoFilter because the video filter
        'shares the same latch. But the PS bit will not be toggled, so this will not actually affect
        'latched switches that rely on PS, and won't drain capacitors. But this will help initialize
        'latched switches that generate the latching pulse from a change of address.
        'This is done after loading config file so capacitor info is available, and after
        'loading cal files so auto wait info is available.
    videoFilter$="Wide" : FreqMode=2 : switchTR=1 : switchFR=1 : call SelectVideoFilter 'ver116-1b


        '---------Create OperatingCal Folder-------------
    if TGtop>0 then
        isErr=CreateOperatingCalFolder()  'Create OperatingCal folder if it does not exist
        if isErr then notice "Unable to create OperatingCal folder."
    end if

    '-----Load or create coax data file-----
    call CoaxLoadDataFile   'ver115-4a

'2.Establish hard "Global" variables
    'For speed, most of the following are not declared global, and they are not accessible
    'within true subroutines. But a couple have true global versions.
    contclear = 11         'to take all LPT control lines low
    STRB = 10              'to take LPT-pin 1 high. (SLIM Latch 4)
    AUTO = 9               'to take LPT-pin 14 high.  (SLIM Latch 3)
    INIT = 15              'to take LPT-pin 16 high. (SLIM Latch 2)
    SELT = 3               'to take LPT-pin 17 high. (SLIM Latch 1)
    INITSELT =  7          'to take both LPT-pins 16 & 17 high. (INIT,SELT)(was enapt) ver111-22
    STRBAUTO = 8           'to take both LPT-pins 1 & 14 high. (FQUD,WCLK)(was wclkfqud) ver111-22
    global globalSTRB, globalINIT, globalSELT, globalContClear  'ver116-1b
    globalSTRB=STRB : globalINIT=INIT : globalSELT=SELT : globalContClear=contclear   'ver116-1b

    if cb=0 then le1=4:le2=8:le3=16:fqud1=STRB:fqud3=2 'ver111-31b
    if cb=1 then le1=1:le2=1:le3=4:fqud1=2:fqud3=8 'ver111-31b
    if cb=2 then le1=1:le2=16:le3=4:fqud1=2:fqud3=8 'ver111-31b
    if cb=3 then le1=1:le2=16:le3=4:fqud1=2:fqud3=8:bUseUsb = 1  'USB:01-08-2010
    if adconv = 8 then pdmlowlim = 51 : pdmhighlim = 205 'establish boundries for 8 bit parallel A to D ver111-36f
    if adconv = 12 then pdmlowlim = 819 : pdmhighlim = 3277 'establish boundries for 12 bit parallel A to D ver111-36f
    if adconv = 16 then pdmlowlim = 13107 : pdmhighlim = 52429 'establish boundries for 16 bit serial A to D ver111-36f
    if adconv = 22 then pdmlowlim = 819 : pdmhighlim = 3277 'establish boundries for 12 bit serial A to D ver111-37a


'3.Initialize for whatever mode we will start up in
        'Some of these initializations may be changed when the preferences file
        'is loaded in [LoadPreferenceFile]
    doingInitialization=1 'ver114-3f
    suppressPhase=0     'Turns phase on ver116-1b
    suppressHardwareInitOnRestart=0     'Normally we do hardware initialization on each restart.
    refreshOnHalt=1     'We normally redraw the graph when we halt.
    multiscanIsOpen=0
    multiscanInProgress=0
    baseFrequency=0  'ver116-4k
    cftest=0    'cavity filter sweep test off ver116-4b
    message$=""  'ver115-1a
    msaMode$="SA"  'ver115-3b
    planeadj=0  'ver114-4i
    gentrk=0 : normrev=0 : sweepDir=1   'ver 114-4k
    refreshEachScan=1    'ver114-3f
    videoFilter$="Wide"  'ver116-1b
    useAutoWait=0    'ver116-1b
    autoWaitPrecision$="Normal"  'ver116-1b
    call TwoPortInitVariables 'initialize variables including default Y axis ranges for Two-Port ver116-1b

    call SelectVideoFilter 'Sets videoFilterAddress and outputs it ver116-1b
        'Note the latched switches--Band, FR and TR--may actually not be "latching" (because they may just
        'rely on the control board latch, and they may or may not use the Pulse Start (Latch Pulse) to trigger latching.
        'If they generate their own pulse for latching based on a change of required state, we have to be sure
        'that on startup they do so. For example, just setting the frequency band switch to 01 for 1G may not create
        'a latch pulse, if the control board latch already had that value. We need to initialize those switches to
        'two successive values. We have to wait in between to allow for some recharging of the capacitors.
        'The first state was set at the very beginning of the program, to utilize the time delays of initializtion.
        'Video Filter is not a latched switch, but it has to be output along with the other data
        'This should work if the switch capacitors discharge only about 10% and recharge time constant is 1 second or less
        'Note that SelectLatchedSwitches adds some time delay also
    FreqMode=1 : switchTR=0 : switchFR=0 : call SelectLatchedSwitches 'ver116-1b
    call uSleep 500 'wait again because latching will occur again when preferences are loaded.

    returnBeforeFirstStep=0 'ver115-1d
    haltedAfterPartialRestart=0 'ver116-1b
    specialOneSweep=0 'ver115-1d
    crystalLastUsedID=0
    imageSaveLastFolder$=DefaultDir$    'Folder in which image was last saved ver115-2a
    touchLastFolder$=DefaultDir$     'Folder from which param data was last loaded ver115-5f
    doSpecialRLCSpec$="RLC[P, R1000,C1n,L1u]"   'default for doSpecialGraph of simulated RLC
    RefRLCLastNumPoints=0
    RefRLCLastConnect$=""   'For continuity calling ReflectionRLC
    analyzeQLastNumPoints=0   'For continuity in AnalyzeQ
    GDLastNumPoints=0       'Number of points last used for group delay analysis
    frontEndCalNumPoints=0  'No front end adjustment
    frontEndActiveFilePath$=""
    frontEndLastFolder$=DefaultDir$

    'ver114-3f moved the call to gInitFirstUse here from [CreateGraphWindow]

    gosub [FindClientOffsets]   'set clientWidthOffset and clientHeightOffset from test window ver115-1b
    smithLastWindowHeight=430 : smithLastWindowWidth=400    'ver115-5d
    currGraphBoxHeight=600-clientHeightOffset-44   'ver115-1b 'ver115-1c 44 allows for button area below box?
    currGraphBoxWidth=800-clientWidthOffset  'ver115-1b 'ver115-1c
    graphMarLeft=70 : graphMarRight=180 : graphMarTop=55 : graphMarBot=140  'SEWgraph Graph margins from edge of graphicbox
    call gInitFirstUse "#handle.g", currGraphBoxWidth, currGraphBoxHeight, graphMarLeft, graphMarRight, _ 'ver115-1c
                                graphMarTop, graphMarBot 'SEWgraph Initialize graphing module
    gosub [InitGraphParams] 'SEWgraph  'Initialize parameters to set up the graphing module ver114-3f moved
    gosub [ChangeMode] 'create Graph Window in mode of msaMode$ 
    desiredCalLevel=0   'Desire no cal ver114-6b
    call SignalNoCalInstalled
    bandLineNumSteps=-1   'Indicate cal does not exist ver114-5f; baseLine cal was handled above ver114-5mb
    OSLBandNumSteps=-1  'ver115-1b
    OSLBaseNumSteps=-1  'ver115-1b

    OSLLastSelectedCalSet=0  'Indicates there was no prior selection ver115-7a
    OSLOpenSpec$="" : OSLShortSpec$="" : OSLLoadSpec$=""    'ver116-4i
    OSLS11JigType$="Reflect"   'ver115-1b
    'Defaults are now in place. Read the preferences file and save it. If there is no preference file
    'this has the effect of creating one with the default values. If there is one, saving it updates a
    'possibly old preferences file to the current format.
    'ver114-3f added instructions to load and save preference file
    restoreFileName$=DefaultDir$+"\MSA_Info\MSA_Prefs\Prefs.txt"    'tells [LoadPreferenceFile] what to load
    gosub [LoadPreferenceFile]     'ver115-8c
    'call uSleep 500     'Loading Preferences may re-latch switches; allow some recharge time ver116-1b delver116-4d
    if restoreErr$<>"" then 'ver115-1b
        'Error. If error was that file does not exist at all, that is OK; we'll create one
        fileHndl$=OpenContextFile$(restoreFileName$,"IN")
        if fileHndl$<>"" then
            'A file exists and it is bad
            close #fileHndl$
            notice "Error in Preference File: ";restoreErr$  'ver114-7n
        else
            restoreErr$=""  'Clear error if problem was file does not exist. ver115-3c
        end if
    end if
    'Resave preferences only if no serious error
    if restoreErr$="" then call SavePreferenceFile restoreFileName$ 'Save Preference file in current format ver115-4a
    if gGetXIsLinear() then userFreqPref=0 else userFreqPref=1      'Start with Center/Span for linear, Start/Stop for log 'ver115-1d
    call mClearMarkers   'SEWgraph   'Clear all graph markers
    'ver114-3c moved some lines from step 3 into InitGraphParams

'4.measure computer speed and update global, glitchtime
    'Determine speed of computer 'ver111-37c
    if glitchtime = 0 then gosub [AutoGlitchtime] 'ver111-37c
    'return with glitchtime, number approximates 1 millisecond of computer processing speed with Liberty Basic
    'this is a "coarse" calculation.
'5.Command Filter Bank
[InitializeHardware]
    'These hardware initializations are performed on startup and usually repeated on Restart. The reason
    'they are repeated on Restart is to fix any hardware glitches that might occur. Whenever it is known
    'that a hardware change is made, such as filter selection changing, it is best to take action immediately,
    'and not rely on the Restart process. In some cases, Restart skips these initializations for speed.
    if suppressHardware=0 and cb<3 then  'ver115-6c USB:02-08-2010 added cb test
        out port, 0                 'begin with all data lines low
        if cb = 2 then  'ver116-1b
            out control, INITSELT 'latch "0" into SLIM Control Board Buffers 1 and 2
            out control, AUTO 'latch "0" into SLIM Control Board Buffers 3
            'We don't clear SLIM Buffer 4, because it controls among other things the latched switches
            'It was initialized near the beginning to make the PS line high.
        end if
        out control, contclear      'begin with all control lines low
    end if
    if (cb = 3) and (bUseUSB<>0) then 'USB:01-08-2010
        USBwrbuf$ = "A5010000" ' reset all lines low 'USB:01-08-2010
        if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf$ as ptr, 4 as short, result as boolean 'USB:01-08-2010
    end if 'USB:01-08-2010
     'the following are meaningless values to guarantee first time commanding. Used in subroutine, [DetermineModule]
    lastdds1output = appxdds1:lastdds3output = appxdds3:lastpdmstate = 2 'ver111-28
    lastncounter1 = 0 : lastncounter3 = 0 'to guarantee Original MSA will command PLL's after init. ver114-6c
    error$="" : errora$=""  'ver115-1c

    'Initialize Final Filter path. 
    call CommandFilter filtbank   'Commands and sets filtbank. Does nothing if suppressHardware=1. ver115-6c

    call SelectVideoFilter  'reselect video filter in case a glitch got it 'ver116-1b
    'Note we don't reset the latched switches on Restart (for startup, they are set when prefs are loaded),
    'because it can get obnoxius and requires a time delay.
    'Plus, we don't want to set them when the user makes a change, and immediately set again on Restart.
    'These switches are properly set whenever DetectChanges is called, which should take care of them.
    'ver116-4d deleted call to SelectLatchedSwitches

'6.if configured, initialize DDS3 by reseting to serial mode. Frequency is commanded to zero
    if suppressHardware then goto [SkipHardwareInitialization]  'In case there is no hardware ver115-6c
    if TGtop = 0 then goto [endInitializeTrkGen]' there is no Tracking Generator ver111-22
    'Initialize DDS 3
        if cb = 0 and TGtop = 2 then Jcontrol = INIT:swclk = 32:sfqud = 2:gosub [ResetDDS3ser] 'ver111-7
        '[ResetDDS3ser]needs:port,control,Jcontrol,swclk,sfqud,contclear ; resets DDS3 into Serial mode
        if cb = 2 then gosub [ResetDDS3serSLIM] 'ver111-29
        if cb = 3 then gosub [ResetDDS3serUSB]  'USB:01-08-2010
'7.if configured, initialize PLO3. No frequency command yet.
    'Initialize PLL 3. 'CreatePLL3R,CommandPLL3R
        appxpdf=PLL3phasefreq 'ver111-4
        if TGtop = 1 then reference=masterclock 'ver111-4
        if TGtop = 2 then reference=appxdds3 'ver111-4
        gosub [CreateRcounter]'needs:reference,appxpdf ; creates:rcounter 'ver111-14
        rcounter3=rcounter : pdf3=pdf 'ver111-7
    'CommandPLL3R and Init Buffers
        datavalue = 8:levalue = 4 'PLL3 data and le bit values ver111-28
        gosub [CommandPLL3R]'needs:PLL3mode,PLL3phasepolarity,INIT,PLL3 ; Initializes and commands PLL3 R Buffer(s) 'ver111-7
[endInitializeTrkGen]   'skips to here if no TG

'8.initialize and command PLO2 to proper frequency
    'CreatePLL2R
        appxpdf=PLL2phasefreq 'ver111-4
        reference=masterclock 'ver111-4
        gosub [CreateRcounter]'needed:reference,appxpdf ; creates:rcounter,pdf 'ver111-14
        rcounter2 = rcounter 'ver111-7
        pdf2 = pdf    'actual phase detector frequency of PLL 2 'ver111-7
    'CommandPLL2R and Init Buffers
        datavalue = 16: levalue = 16 'PLL2 data and le bit values ver111-28
        gosub [CommandPLL2R]'needs:PLL2phasepolarity,SELT,PLL2 ; Initializes and commands PLL2 R Buffer(s)
    'CreatePLL2N
        appxVCO = appxLO2 : reference = masterclock
        gosub [CreateIntegerNcounter]'needs:appxVCO,reference,rcounter ; creates:ncounter,fcounter(0)
        ncounter2 = ncounter:fcounter2 = fcounter
        gosub [CreatePLL2N]'needs:ncounter,fcounter,PLL2 ; returns with Bcounter,Acounter, and N Bits N0-N23
        Bcounter2=Bcounter: Acounter2=Acounter
        LO2=((Bcounter*preselector)+Acounter+(fcounter/16))*pdf2 'actual LO2 frequency  'ver115-1c LO2 is now global
    'CommandPLL2N
        Jcontrol = SELT : LEPLL = 8
        datavalue = 16: levalue = 16 'PLL2 data and le bit values ver111-28
        gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111-5

'9.Initialize PLO 1. No frequency command yet.
    '[InitializePLL1]'set PLL1 to proper Rcount and initialize
'        appxpdf=PLL1phasefreq 'ver111-4
'        reference=appxdds1 'ver111-4
'        gosub [CreateRcounter]'needed:reference,appxpdf ; creates:rcounter,pdf 'ver111-4
'        rcounter1 = rcounter 'ver111-4
        'Create rcounter1 ver114-2e
    rcounter1=int(appxdds1/PLL1phasefreq)   'ver114-2e
    if (appxdds1/PLL1phasefreq) - rcounter1 >= 0.5 then rcounter1 = rcounter1 + 1   'rounds off rcounter  ver114-2e
    if spurcheck=1 and PLL1mode=0 then rcounter1 = rcounter1 +1 'only do this for IntegerN PLL  ver114-2e

    'CommandPLL1R and Init Buffers
    datavalue = 2: levalue = 1 'PLL1 data and le bit values ver111-28
    gosub [CommandPLL1R]'needs:rcounter1,PLL1mode,PLL1phasepolarity,SELT,PLL1 ; Initializes and commands PLL1 R Buffer(s)

'10.initialize DDS1 by resetting. Frequency is commanded to zero
    'It should power up in parallel mode, but could power up in a bogus condition.
    if cb = 0 and dds1parser = 0 then gosub [ResetDDS1par]'(Orig Control)'needs:control,STRBAUTO,contclear ; resets DDS1 on J5, parallel ver111-21
    if cb = 0 and dds1parser = 1 then gosub [ResetDDS1ser]'(Orig Control)'needed:control,AUTO,STRB,contclear  ; resets DDS1 on J5, into serial mode ver111-21
    if cb = 2 then gosub [ResetDDS1serSLIM]'reset serial DDS1 without disturbing Filter Bank or PDM 'ver111-29
    if cb = 3 then gosub [ResetDDS1serUSB]'reset serial DDS1 without disturbing Filter Bank or PDM  'USB:01-08-2010

[SkipHardwareInitialization]    'Skips to here if there is no hardware (suppressHardware=1) ver115-6c
    cb=cb   'to avoid two labels in a row

'11.[BeginScanSeries] get info from windows and update variables
[BeginScanSeries]   'Start a new series of scans, which requires some initialization

'12.[InitializeGraphModule]
    suppressPDMInversion=0  'ver115-1a
    'ver115-8d moved test for specialOneSweep to step 14
    gosub [UpdateGraphParams] 'SEWgraph Update graph module for any changes made by the user
    firstScan=1     'Signal that the next scan is the first after Restart
    'ver114-5f moved some items to UpdateGraphParams

    'Note x values must be calculated first (in [UpdateGraphParams]) ; modVer115-1c
    'If calInProgress=1, InstallSelectedxxx will just set applyCal=0 and installed base steps=-1    'ver116-4b
    if msaMode$="Reflection" then
        call InstallSelectedOSLCal
    else
        if msaMode$<>"SA" then call InstallSelectedLineCal  'ver115-8c
    end if

    call gInitDynamicDraw   'Set up for first scan of dynamic draw/erase/redraw...
    call ImplementDisplayModes  'Done in [UpdateGraphParams] but gInitDynamicDraw overrode it   'ver115-4e

    'In multiscan, we don't want to update the time stamp on every redraw, which sometimes happens without scanning.
    if multiscanIsOpen=0 or multiscanInProgress=1 then 'ver115-9a
        restartTimeStamp$=date$("mm/dd/yy"); "; ";time$()   'ver115-2c
        call gSetTitleLine 3, restartTimeStamp$   'SEWgraph Put date and time in line 3 of title
        if gGetXIsLinear() then call gSetTitleLine 4, "MSA Linear Sweep ";path$ _
                            else call gSetTitleLine 4, "MSA Log Sweep ";path$      'Save linear/log and path info ver116-1b
    end if

        'For multiscan, the redraw of the background is done prior to scanning via [PartialRestart], and on refresh
    if multiscanInProgress=0 then
     'Redraw background stuff on first scan of a series. ver115-8d
         'ver115-8d deleted calc of centerstep, which is no longer used
        call gDrawGrid      'SEWgraph; Clear graphics area and draw the background grid and labels. Wipes out all prior flushes.
        call DrawSetupInfo   'SEWgraph  Draw info describing the sweep setup
        if smithGraphHndl$()<>"" then  'ver115-1b draw smith chart if we have one ver115-1e
            call smithRedrawChart 'Draw blank chart ver115-2c
        end if
        if referenceLineType<>0 then   'Draw reference lines ver114-8a
            if referenceLineType>1 then call CreateReferenceSource  'RLC or fixed value
            call CreateReferenceTransform   'Generate actual reference graph data
            call gClearAllReferences
            if referenceDoMath=0 then   'don't draw ref if we are using ref for math
                if (referenceTrace and 2) then call gAddReference 1,CreateReferenceTraces$(referenceColor2$,referenceWidth2,2)  'Do Y2 reference
                if (referenceTrace and 1) then _
                        call gAddReference 2,CreateReferenceTraces$(referenceColor1$,referenceWidth1,1) 'Do Y1 reference
                call gDrawReferences
                refHeadingColor1$=referenceColor1$ : refHeadingColor2$=referenceColor2$ 'ver115-5d
            else
                call gGetTraceColors refHeadingColor1$, refHeadingColor2$ 'Use trace colors for "REF" if math is used
            end if
            call PrintReferenceHeading  'Print above axis to indicate which line matches which axis 'ver115-5d
        end if
        #graphBox$, "flush"  'Make the setup info stick
    end if

    useExpeditedDraw=gCanUseExpeditedDraw()   'SEWgraph; For normal SA use, [gDrawSingleTrace] will be used.
    'ver115-1a deleted printing of glitchtime
    doingInitialization=0   'We are done with initialization on startup 'ver114-4g moved
    if calInProgress=1 then 'ver114-5g
        message$="Calibration in progress." : call PrintMessage 'ver114-4g
    else
        message$="" : call PrintMessage    'ver114-4f
    end if
    if msaMode$="SA" and gentrk=0 and multiscanInProgress=0 then    'ver115-4f
        if (endfreq-startfreq)/steps >finalbw/1000 then     'compare as MHz
            message$= "Frequency step size exceeds RBW; signals may be missed."
            call PrintMessage
        end if
    end if

'13.Calculate the command information for first step through last step of the sweep and put in arrays
    cursor hourglass    'ver116-4k
    if suppressHardware then    'ver115-6c
            'These are taken from the routines we are skipping
        for i=0 to steps
            thisfreq=gGetPointXVal(i+1)    'Point number is 1 greater than step number SEWgraph
            if msaMode$<>"SA" then  'Store actual signal freq in VNA arrays ver116-1b
                if msaMode$="Reflection" then ReflectArray(thisstep,0)=thisfreq _
                                else S21DataArray(thisstep,0)=thisfreq
            end if
            if FreqMode<>1 then thisfreq=Equiv1GFreq(thisfreq)  'Convert from display freq to equivalent 1G frequency
            datatable(i,0) = thisstep    'put current step number into the array, row value= thisstep 'moved ver111-18
            datatable(i,1) = thisfreq
            phaarray(i,0) = 0   'pdm state
        next i
    else    'Do these only if we are using the hardware 'ver115-6c
        gosub [CalculateAllStepsForLO1Synth] 'ver111-18
        if TGtop > 0 then gosub [CalculateAllStepsForLO3Synth] 'ver111-18
        gosub [CreateCmdAllArray] 'ver111-31b
    end if
    call CalcFreqCorrection     'Calculate power correction at each frequency SEWgraph1
    cursor normal   'ver116-4k
    if msaMode$="SA" and frontEndActiveFilePath$<>"" then call frontEndInterpolateToScan  'Calculate corrections for front end ver115-9d
    continueCode=0    'SEWgraph Set to other values by subroutines to cause halt, wait or restart

'14.[StartSweep]'Begin sweeping from step 0
'StartSweep begins the outer loop that repeats the entire scan process until halted.
'The scan loop continues until a user action which aborts the scan, or in the case of
'OneStep it continues only for a single point. If specialOneSweep=1 or HaltAtEnd=1, it
'automatically stops at the end of a single sweep.
    scanResumed=0 'used to indicate whether we start with a new scan(0) or resume where we left off(1)SEW
    haltedAfterPartialRestart=0 'May get set to 1 a few lines below. 116-1b
    'ver114-6e Normally, refresh will occur at end of scan only if halted or refreshEachScan=1,
    'and will be done by expedited methods. But if the user makes certain changes, the following
    'variables are used to force more extensive redrawing.
    call mDeleteMarker "Halt"    'ver114-4h moved the -4d version
    suppressSweepTime=1     'to suppress it for the first scan ver114-4h
        'if we just want to go through the initialization procedure we set returnBeforeFirstStep
        'and invoke [Restart] with a gosub; here we return to the caller

        'Save some sweep settings for reflection and transmission for use when changing
        'back to a previously used mode, so we know the nature of the last gathered data
    if msaMode$="Reflection" then   'ver116-1b
        refLastSteps=steps : refLastStartFreq=startfreq : refLastEndFreq=endfreq : refLastIsLinear=gGetXIsLinear()
        refLastGraphR0=S11GraphR0
        refLastY1Type=Y1DataType : refLastY1Top=Y1Top : refLastY1Bot=Y1Bot : refLastY1AutoScale=autoScaleY1
        refLastY2Type=Y2DataType : refLastY2Top=Y2Top : refLastY2Bot=Y2Bot : refLastY2AutoScale=autoScaleY2
        for i=1 to 4 : refLastTitle$(i)=gGetTitleLine$(i) : next i
    else
        if msaMode$="VectorTrans" then
            transLastSteps=steps : transLastStartFreq=startfreq : transLastEndFreq=endfreq : transLastIsLinear=gGetXIsLinear()
            transLastGraphR0=S21JigR0
            transLastY1Type=Y1DataType : transLastY1Top=Y1Top : transLastY1Bot=Y1Bot : transLastY1AutoScale=autoScaleY1
            transLastY2Type=Y2DataType : transLastY2Top=Y2Top : transLastY2Bot=Y2Bot : transLastY2AutoScale=autoScaleY2
            for i=1 to 4 : transLastTitle$(i)=gGetTitleLine$(i) : next i
        end if
    end if

    if returnBeforeFirstStep then   'ver115-2a
        thisstep=sweepStartStep
        returnBeforeFirstStep=0
        haltedAfterPartialRestart=1 'ver116-1b
        gosub [CleanupAfterSweep]
        return    'ver115-1d
    end if
[StartSweep]'enters from above, or [IncrementOneStep]or[FocusKeyBox]([OneStep][Continue])
    ' USB:15/08/10
    ' this code may (or may not) be optimally placed.
    ' it sets the ADC parameters for USB and should only be run
    ' as necessary. Here seemed like a reasonable comproomise.
    ' there is a similar function in the CommandAllSlimsUsb function
    ' but that one uses step 0 to decide (which is also a bit dodgy)
    if cb = 3 then ' USB:15/08/10
        if adconv = 16 then 'USB:05/12/2010
            UsbAdcControl.Clocking.struct = 0 'USB:05/12/2010
            UsbAdcControl.Delay.struct = 0 'USB:05/12/2010
            UsbAdcControl.Bits.struct = 16 'USB:05/12/2010
            UsbAdcControl.Average.struct = 1 'USB:05/12/2010
        else 'USB:05/12/2010
            if adconv = 22 then 'USB:05/12/2010
                UsbAdcControl.Clocking.struct = 1 'USB:05/12/2010
                UsbAdcControl.Delay.struct = 1 'USB:05/12/2010
                UsbAdcControl.Bits.struct = 10 'USB:05/12/2010
                UsbAdcControl.Average.struct = 1 'USB:05/12/2010
            else 'USB:05/12/2010
                ' this is an error - so to be safe cheat and say it is a type 16
                ' ought to return an error but I do not know how to do that in Liberty basic
                UsbAdcControl.Clocking.struct = 0 'USB:05/12/2010
                UsbAdcControl.Delay.struct = 2 'USB:05/12/2010
                UsbAdcControl.Bits.struct = 16 'USB:05/12/2010
                UsbAdcControl.Average.struct = 1 'USB:05/12/2010
            end if 'USB:05/12/2010
        end if 'USB:05/12/2010
    end if  ' USB:15/08/10

    if specialOneSweep then haltAtEnd=1 else haltAtEnd=0   'ver115-8d moved this here
    if haltedAfterPartialRestart=0 and scanResumed=1 then   'ver116-1b
        'For a resumed scan, a halt occurred after the previous step and that step was fully processed.
        'haltsweep will equal 0. If alternateSweep=0 and the halt occurred at the end of a sweep, we need to
        'repeat the last point as the first point of the new sweep. But in the case where we are continuing
        'after a halt resulting from partial restart, we returned before the first step was taken and need to
        'start with that step.
        call mDeleteMarker "Halt"    'ver114-4h moved the -4d version
        if thisstep = sweepStartStep and syncsweep = 1 then gosub [SyncSweep] 'ver112-2b; ver114-4k
        if alternateSweep=0 or haltWasAtEnd=0 then  'ver114-5c Go to next step unless we need to repeat this one
            if sweepDir=1 then
                if thisstep<sweepEndStep then thisstep = thisstep + 1 else thisstep=sweepStartStep
            else
                if thisstep>sweepEndStep then thisstep = thisstep - 1 else thisstep=sweepStartStep
            end if
        end if
    else    'ver114-5c No longer need to retest scanResumed
        thisstep=sweepStartStep  'ver114-4k
    end if
    haltedAfterPartialRestart=0 'Reset. Will stay zero until next partial restart. 116-1b
    scanResumed=0   'Reset flag

'15.[CommandThisStep]. command relevant Control Board and modules
'SEW CommandThisStep begins the inner loop that moves from step to step to complete a single
'SEW scan.This branch label is accessed only from the end of the loop.
[CommandThisStep]'needs:thisstep ; commands PLL1,DDS1,PLL3,DDS3,PDM 'ver111-7
    'a. first, check to see if any or all the 5 module commands are necessary [DetermineModule]
    'b. calculate how much delay is needed for each module[DetermineModule], but use only the largest one[WaitStatement].
    'c. send individual data, clocks, and latch commands that are necessary for[CommandOrigCB]
      'or for SLIM, use [CommandAllSlims] for commanding concurrently 'ver111-31c
    gosub [CommandCurrentStep]  'ver116-4j made this a separate routine

'16.Determine sequence of operations after commanding the modules
    if onestep = 1 then  'in the One Step mode
        glitchhlt = 10 'add extra settling time
        gosub [ReadStep] 'read this step
        gosub [ProcessAndPrint] 'process and print this step
        call DisplayButtonsForHalted    'ver114-4f replaced call to [UpdateBoxes]
        call mAddMarker "Halt", thisstep+1, "1"   'ver114-4d
            'If marker is shown on graph, we need to redraw the whole graph
            'Otherwise just redraw the marker info
        if doGraphMarkers then call RefreshGraph 0 else call mDrawMarkerInfo  'No erasure gap in redraw ver114-5m
        if thisstep=sweepEndStep then
            'Note reversal is after graph is redrawn
            if alternateSweep then gosub [ReverseSweepDirection] 'ver114-4m; ver114-5e
            haltWasAtEnd=1  'ver114-5c
        else
            haltWasAtEnd=0  'ver114-5c
        end if
        wait 'wait here for next button push ver113-6d
    end if

    if haltsweep = 0 then 'in first step after a Halt
        haltsweep = 1 'change flag to say we are not in first step after a Halt, for future steps
        glitchhlt = 10  'add extra settling time
        gosub [ReadStep] 'read this step 'ver113-6d
    else  'if neither, then in middle of sweep. process and print the previous step, then read this step
        gosub [ProcessAndPrintLastStep]
        gosub [ReadStep]'read this step 'ver113-6d
    end if
        'ver116-4k moved sweep time here, so it prints after any refresh action from the prior scan
    if displaySweepTime and thisstep=sweepStartStep then
        currTime=Time$("ms")
        if suppressSweepTime=0 then _
                message$= "Sweep Time=";using("####.##", (currTime-startTime)/1000);" sec." : call PrintMessage   'ver114-4h
        suppressSweepTime=0 'Only suppress on first scan 'ver114-4h
        startTime=currTime       'SEWgraph timer for testing
    end if

'17.[Scan] Check to see if a button has been pushed
'SEWgraph Note that on any user action, if haltsweep=1 the action must have been detected
'during the following "scan". But if haltsweep=0 the action occurred during a wait state.
'Exception: Window resizing is detected when it happens, not during "scan".
[Scan] 'ver113-6d
    scan    'check for any button push and go there. ver111-26
            'otherwise, continue sweeping.
[PostScan]  'SEWgraph. Label is used to return here after a button action handled by a [xyz] routine. 
    'SEWgraph Note: after a button handler in the form of a true subroutine, control will exit sub back
    'to this point. We do not want a "wait" to occur in such a subroutine, because that will suspend
    'control in a non-global namespace, and the user will be unable to take actions that require access
    'to [xyz] routines. To cause a halt, wait or restart in such a subroutine, the subroutine should set
    'the global variable continueCode to 1, 2 or 3.
    if continueCode<>0 then 'SEWgraph created this if... block; =0 means continue normally
        if continueCode=1 then continueCode=0 : goto [Halted]   '=1 means halt immediately
        if continueCode=2 then continueCode=0 : haltsweep=0 : wait     '=2 means wait immediately
        continueCode=0 : haltsweep=0 : goto [Restart]    'Anything else means restart
    end if

'18.[IncrementOneStep]
'SEW IncrementOneStep is the end of both the inner loop over points and the outer loop
'SEW over scans. goto [CommandThisStep] continues the inner loop with the next point.
'SEW goto[StartSweep] continues the outer loop with the next scan.
'SEW [IncrementOneStep] is commented out to be clear it is not used for any goto.
'[IncrementOneStep]
    if thisstep = sweepEndStep and syncsweep = 1 then gosub [SyncSweep] 'ver112-2b 'ver114-4k
    'ver114-5a modified the following
    if sweepDir=1 then  'ver114-4k added this block to handle possible reverse sweeps
        if thisstep<sweepEndStep then thisstep = thisstep + 1 : goto [CommandThisStep]    'ver114-4k
    else
        if thisstep>sweepEndStep then thisstep = thisstep - 1 :goto [CommandThisStep]    'ver114-4k
    end if
    'If we are here, we have just read the final step of a sweep
    if haltAtEnd=0 then
            'Alternate sweep directions if required. When we switch direction, thisstep
            'was the final point of one sweep and becomes the first point of the next.
            'We process and print it  immediately as the last point of this sweep; then reverse
            'direction and start with the same point. To avoid re-processing it at the next step we
            'set haltsweep=0.
        if alternateSweep then  'ver114-5c
            gosub [ProcessAndPrint]
            gosub [ReverseSweepDirection]
            haltsweep=0
        end if
        goto [StartSweep] 'SEWgraph Repeat loop over scans if halt flag not set
    end if

    'SEWgraph We fall out of this loop only when haltAtEnd=1 and we reach thisstep=sweepEndStep
'[EndSweepSeries] 'This label marks the end of the scan loops.

'19.[Halted]
[Halted]    'SEWgraph moved guts of this to FinishSweeping, which can also be called from elsewhere if desired.
    gosub [FinishSweeping]'get raw data, process, print to the computer monitor ver111-22
    if specialOneSweep then
        specialOneSweep=0
        return 'Sweep process was called by gosub; we return to caller.
    end if
    wait 'wait for operator action

      'SEWgraph created FinishSweeping;
      'ver114-6e split the non-graphing cleanup into [CleanupAfterSweep]
[FinishSweeping]    'SEWgraph0 Do cleanup to end sweeping but return for further actions
    'This is a modified version of the former [Halted], without the wait at the end
    gosub [ProcessAndPrint]'process, print to the computer monitor ver111-22
    if haltAtEnd=0 then call mAddMarker "Halt", thisstep+1, "1" 'Add Halt marker ver114-4d
    haltsweep=0 'do now so RefreshGraph will "flush" ver115-1a
    if isStickMode=0 then
        if refreshOnHalt then   'ver115-8c
            refreshGridDirty=1: call RefreshGraph 1  'SEWgraph; redraw and show erasure gap; don't do if stick mode ver114-7d
        else
            'We sometimes don't want to waste time redrawing, such as when we are loading a data file,
            'but we at least need to flush to make the graphics stick.
            #graphBox$, "flush"
        end if
    end if
    if specialOneSweep and thisstep <> sweepEndStep then
        beep: message$="Sweep Aborted" 'ver115-4b
    else
        if calInProgress then beep: message$="Calibration Complete" 'ver115-4b
    end if
        'test is used for troubleshooting. Coder can insert
        'test = (any variable) anywhere in the code, and it will get displayed in the Messages Box during Halt.
    if test<>0 then message$=str$(test)
    if message$<>"" then call PrintMessage 'ver115-4b
        'Alternate sweep directions if required; added by ver114-5a
    if thisstep=sweepEndStep then if alternateSweep then gosub [ReverseSweepDirection] 'ver114-5a
    goto [CleanupAfterSweep]

[CleanupAfterSweep] 'Do cleanup after a sweep to be sure flags are set/reset properly
'Called by [FinishSweeping]. Can also be called by other routines to immediately
'terminate a sweep when they will be Restarting so they don't care about finishing the plotting.
    call DisplayButtonsForHalted    'ver114-4f replaced call to [UpdateBoxes]
    if thisstep=sweepEndStep then haltWasAtEnd=1 else haltWasAtEnd=0  'ver114-5c
    haltAtEnd=0     'SEWgraph In case we got here from auto halt at end of sweep
    calInProgress=0  'ver114-5h
    haltsweep = 0 'this says the sweep has been halted, so don't print the first command of the next sweep step 'ver111-20
return

[ReverseSweepDirection] 'Reverse direction of sweep
'This is called after sweepEndStep has been fully processed, but only if alternateSweep=1
    if sweepDir=1 then
        sweepDir=-1
        sweepStartStep=steps : sweepEndStep=0
    else
        sweepDir=1
        sweepStartStep=0 : sweepEndStep=steps
    end if
    call gSetSweepDir sweepDir 'Notify graph module of new direction
return

'ver116-4j made [CommandCurrentStep] a separate gosub from the old [CommandThisStep] so it can be called not only during regular scanning,
'but on in combination with [ReadStep] to command and read a particular step, once all info is set up.
[CommandCurrentStep]'needs:thisstep ; commands PLL1,DDS1,PLL3,DDS3,PDM 'ver111-7
    'a. first, check to see if any or all the 5 module commands are necessary [DetermineModule]
    'b. calculate how much delay is needed for each module[DetermineModule], but use only the largest one[WaitStatement].
    'c. send individual data, clocks, and latch commands that are necessary for[CommandOrigCB]
      'or for SLIM, use [CommandAllSlims] for commanding concurrently 'ver111-31c
    if suppressHardware=0 then 'ver115-6c
        gosub [DetermineModule] 'determine which, if any, module needs commanding.  ver111-27
        cmdneeded = glitchp1 + glitchd1 + glitchp3 + glitchd3 + glitchpdm 'ver111-38a
        if cmdneeded > 0 and cb = 0 then gosub [CommandOrigCB]'old Control (150 usec, 0 SW) 'ver111-28ver111-38a
        'if cb = 1 then gosub [CommandRevB]'old Control looking like SLIM  'not created yet
        if cmdneeded > 0 and cb = 2 then gosub [CommandAllSlims]'ver111-38a
        if cmdneeded > 0 and cb = 3 then gosub [CommandAllSlimsUSB] 'USB:01-08-2010
        if cftest=1 then gosub [CommandLO2forCavTest] 'cav ver116-4c
    end if
return

[FindClientOffsets]   'set clientWidthOffset and clientHeightOffset from test window ver115-1b
     'Open a small test window so we can find the client area to determine how much
      'smaller it is than the full window size.
    WindowWidth = 150 : WindowHeight = 150
    UpperLeftX = 1 : UpperLeftY = 1
    menu #handle, "File", "Save Image", [SaveImage] 'We need a menu to get the size right
    open "Test" for window as #handle
         'Now that we have a window, find the actual client area--ver114-7o
    hWind = hWnd(#handle)        'Windows handle of graph window
    STRUCT Rect, leftX as long, upperY as long, rightX as long, lowerY as long 'To hold the returned data
    calldll #user32,"GetClientRect", hWind as ulong, Rect as struct, r as long  'Fill Rect with size info
        'The offsets will be the size difference between the full window and the client area
    clientWidthOffset = 150-(Rect.rightX.struct-Rect.leftX.struct)
    clientHeightOffset = 150-(Rect.lowerY.struct-Rect.upperY.struct)
    close #handle   'We don't need the test window anymore
return

[ResizeGraphHandler]  'Called when graph window resizes
    #handle, "hide"     'hide window to avoid multiple system redraws ver115-1b
    #graphBox$ "home"
    #graphBox$ "posxy CenterX CenterY"
    currGraphBoxWidth = CenterX * 2-1   'ver115-1c
    currGraphBoxHeight = CenterY * 2-1  'ver115-1c

    'Note: On resizing, all non-buttons seem to end up a few pixels higher than the original spec,
    'so the Y locations are adjusted accordingly via markTop
    'Note WindowHeight when window is created is entire height; on resizing, it is the client area only
    markTop=currGraphBoxHeight+15 : markSelLeft=5 'ver115-1b   'ver115-1c
    markEditLeft=markSelLeft+55
    markMiscLeft=markEditLeft+185
    configLeft=markMiscLeft+80

    #handle, "refresh"
    #handle.Cover, "!show"      'Cover the crap that can appear from resizing
    #handle.Cover, "!hide"      'Uncover and the crap is gone
    #handle, "show"     'show window ver115-1b

    'The graphicbox auto resizes but we have to update the graph module
    'to let it know the new size

    call gUpdateGraphObject graphBox$, currGraphBoxWidth, currGraphBoxHeight, _    'ver115-1c
                                        graphMarLeft, graphMarRight, graphMarTop, graphMarBot
    call gCalcGraphParams   'Calculate new scaling. May change min or max.
    call gGetXAxisRange  xMin, xMax : if startfreq<>xMin or endfreq<>xMax then call SetStartStopFreq xMin, xMax
    call gGenerateXValues gPointCount() 'recreate x values and x pixel locations; keep same number of points
    call gRecalcPix 0   '0 signals not to recalc x pixel coords, which we just did in gGenerateXValues.
    'If a sweep is in progress, we don't want to redraw from here, because that can cause a crash.
    'So we just clear the graph and signal to wait for the user to redraw. This crash may have to
    'do with the fact that we don't know where we are in the sweep process when resizing is invoked,
    'because it is not synchronous with the scan command. Or it may simply have something to do with
    'the fact that no button has yet been pushed on the graph window, which somehow affects the
    'LB resizing process. The crash still sometimes occurs, so it is best to halt before resizing.
    if haltsweep=1 then
        #graphBox$, "cls"
        notice "Warning: Halt before resizing to avoid LB bug."
        call RequireRestart 'ver115-9e  Otherwise old graph still appears, in wrong place.
    else
        refreshRedrawFromScratch=1 'To redraw from scratch ver115-1b
        call RedrawGraph 0  'Redraw at new size
    end if
    'call RequireRestart    'ver115-9e Not sure why we should require restart
    wait

sub ImplementDisplayModes 'calculate the various items from Y1DisplayMode and Y2DisplayMode
    'Y1DisplayMode, Y2DisplayMode: 0=off  1=NormErase  2=NormStick  3=HistoErase  4=HistoStick
    'ver115-2c added checks for constNoGraph
    if (Y1DataType<>constNoGraph and (Y1DisplayMode=2 or Y1DisplayMode=4)) or _
            (Y2DataType<>constNoGraph and (Y2DisplayMode=2 or Y2DisplayMode=4)) then isStickMode=1 else isStickMode=0
    call gSetDoAxis Y1DataType<>constNoGraph, Y2DataType<>constNoGraph 'Turn graph data on or off ver115-3b
        'Note that gActivateGraphs won't activate a graph if we just set its data existence to zero ver115-3b
    call gActivateGraphs Y1DisplayMode<>0,Y2DisplayMode<>0   'Turn actual graphing on or off ver115-3b
    if (Y1DataType<>constNoGraph and Y1DisplayMode>2) or (Y2DataType<>constNoGraph and Y2DisplayMode>2) then _
                call gSetDoHist 1 else call gSetDoHist 0  'Set histogram or normal trace ver115-3b
    call gGetTraceWidth t1Width, t2Width
        'ver114-4n Erase eraseLead points ahead of drawing. The more steps, the larger eraseLead
    if globalSteps<=50 then 'ver114-4n reduced eraseLead
        eraseLead=1
    else
        eraseLead=2+int(steps/400)
        if ((Y1DataType<>constNoGraph and t1Width>2) or _
                (Y2DataType<>constNoGraph and t2Width>2)) _
                    and globalSteps>200 then eraseLead=eraseLead+1
    end if
    if Y2DisplayMode<>1 and Y2DisplayMode<>3 then doErase2=0 else doErase2=1 'ver114-2f
    if Y1DisplayMode<>1 and Y1DisplayMode<>3 then doErase1=0 else doErase1=1 'ver115-3b
    call gSetErasure doErase1, doErase2, eraseLead
end sub

'SEWgraph added UpdateGraphParams; ver114-4n made it a gosub to allow use of non-globals
[UpdateGraphParams] 'Set up graphs for drawing, but don't draw anything
    if alternateSweep then sweepDir=1 : call gSetSweepDir 1      'Start out forward if alternating ver114-5a
    sweepDir=gGetSweepDir()  'ver114-4k
    if sweepDir=1 then  'ver114-4k added this if... block
        'Forward direction
        sweepStartStep=0 : sweepEndStep=steps
    else
        'Reverse direction
        sweepStartStep=steps : sweepEndStep=0
    end if

    'ver115-3b deleted settings related to Y2DisplayMode and Y1DisplayMode. They were overridden in ImplementMagPhaDisp,
    'which is called by UpdateGraphDataFormat
        'ver115-1d deleted the separate call for linear mode. startfreq and endfreq are now valid in all modes.
    call gInitGraphRange startfreq, endfreq, _
                    Y1Bot, Y1Top, Y2Bot, Y2Top  'min and max values for x, y1 and y2; calls gCalcGraphParams
    call gCalcGraphParams   'Calculate new scaling. May change min or max.
    call gGetXAxisRange  xMin, xMax   'in case gCalcGraphParams changed axis limits ver116-4k
    if startfreq<>xMin or endfreq<>xMax then call SetStartStopFreq xMin, xMax   'ver116-4k
    'ver114-5f moved the following here from step 12
    call gGenerateXValues 0   'Precalculate x values for steps+1 points; reset number of points to 0; ver114-1f deleted parameter
    call UpdateGraphDataFormat 0
return

sub UpdateGraphDataFormat doTwoPort  'Update graph module for the type of data we are graphing, and set data source and component
    'If doTwoPort, we are dealing with two-port graphs ver116-1b
    call gSetGridStyles "EndsAndCenter", "All", "All"
        'For linear sweep we display frequency in MHz; for log we do 1, 1 K, 1 M, or 1 G
    if gGetXIsLinear() then     'ver114-6d modified this block to use startfreq/endfreq for log sweeps
        xForm$= "4,6,9//suffix= M"
    else
        xForm$= "3,4,5//UseMultiplier//DoCompact//Scale=1000000"    'ver115-1e
    end if

        'ver115-2c caused the full procedure to be executed for both dataNum.
        'Also eliminated default setting of yForm$
        'ver115-3a moved the select block to DetermineGraphDataFormat so others can use it
    for dataNum=1 to 2
        if doTwoPort then 'ver116-1b
            if dataNum=1 then componConst=TwoPortGetY1Type() else componConst=TwoPortGetY2Type()
        else
            if dataNum=1 then componConst=Y1DataType else componConst=Y2DataType
        end if

        if componConst=constNoGraph then
            doData=0 'Indicates whether there is a graph ver115-2c
            yAxisLabel$="None"  : yLabel$="None"
            yForm$="####.##"    'Something valid, in case it gets mistakenly used
            if dataNum=1 then
                y1AxisLabel$="None"  : y1Label$="None"
                y1Form$="####.##"    'Something valid, in case it gets mistakenly used
            else
                y2AxisLabel$="None"  : y2Label$="None"
                y2Form$="####.##"    'Something valid, in case it gets mistakenly used
            end if
        else
            doData=1
            if dataNum=1 then
                if doTwoPort then
                    call TwoPortDetermineGraphDataFormat componConst, y1AxisLabel$,y1Label$, y1IsPhase,y1Form$
                else
                    call DetermineGraphDataFormat componConst, y1AxisLabel$,y1Label$, y1IsPhase,y1Form$
                end if
            else
                if doTwoPort then
                    call TwoPortDetermineGraphDataFormat componConst, y2AxisLabel$,y2Label$, y2IsPhase,y2Form$
                else
                    call DetermineGraphDataFormat componConst, y2AxisLabel$,y2Label$, y2IsPhase,y2Form$
                end if
            end if
        end if
        if dataNum=1 then doY1=doData else doY2=doData
    next dataNum

    call gSetIsPhase y1IsPhase, y2IsPhase   'Tell graph module whether data is phase
    call gSetAxisFormats xForm$, y1Form$, y2Form$   'Formats for displaying the data values
    call gSetAxisLabels "", y1AxisLabel$, y2AxisLabel$    'Labels for the axes; No label for freq
    call gSetDataLabels y1Label$, y2Label$      'Shorter labels for marker info
    if doTwoPort then 'ver116-1b
            'gSetDoAxis specifies whether data for the axis even exists. gActivateGraphs specifies whether
            'to actually graph the data, based on display mode, which for two port is always On.
        call gSetDoAxis TwoPortGetY1Type()<>constNoGraph, TwoPortGetY2Type()<>constNoGraph 'Turn graph data on or off ver115-3b
            'Note that gActivateGraphs won't activate a graph if we just set its data existence to zero ver115-3b
        call gActivateGraphs 1, 1   'Turn actual graphing on
    else
        call ImplementDisplayModes  'give effect to Y2DisplayMode and Y1DisplayMode
    end if
end sub

sub DetermineGraphDataFormat componConst, byref yAxisLabel$, byref yLabel$,byref yIsPhase,byref yForm$  'Return format info
    'componConst indicates the data type. We return
    'yAxisLabel$  The label to use at the top of the Y axis
    'yLabel$   A typically shorter label for the marker info table
    'yIsPhase$ =1 if the value represents phase. This indicates whether we have wraparound issues.
    'yForm$    A formatting string to send to uFormatted$() to format the data
    '
    'ver116-1b added code to display S12 or S22 instead of S21 or S11 when DUT is reversed.
    if switchFR=0 then
        Sref$="S11" : Strans$="S21" 'Forward DUT
    else
        Sref$="S22" : Strans$="S12"  'ReverseDUT
    end if
    yIsPhase=0  'Default, since most are not phase
    select case componConst 'ver116-4b shortened some labels
        case constGraphS11DB
            yAxisLabel$=Sref$;" Mag(dB)" : yLabel$=Sref$;" dB"
            yForm$="####.###"   'ver115-5d

        case constRawAngle 'Used for transmission mode only  'added by ver115-1i
            yAxisLabel$="Raw Deg" : yLabel$="Raw Deg"
            yIsPhase=1
            yForm$="#####.##"   'ver115-5d

         case constAngle,constGraphS11Ang,constTheta,constImpedAng
            if componConst=constAngle then yAxisLabel$=Strans$;" Deg" : yLabel$=Strans$;" Deg"
            if componConst=constTheta then yAxisLabel$="Theta" : yLabel$="Theta"
            if componConst=constGraphS11Ang then yAxisLabel$=Sref$;" Deg" : yLabel$=Sref$;" Deg"
            if componConst=constImpedAng then yAxisLabel$="Z Deg" : yLabel$="Z Deg"
            yIsPhase=1
            yForm$="#####.##"     'ver115-5d

        case constGD    'calc group delay
            yAxisLabel$="Grp Delay (sec)" :  yLabel$="G.D."
            yForm$="3,2,4//UseMultiplier//DoCompact"

        case constSerReact
            yAxisLabel$="Xs" : yLabel$="Xs"
            yForm$="3,3,4//UseMultiplier//SuppressMilli//DoCompact" 'ver115-4e

        case constParReact
            yAxisLabel$="Xp" : yLabel$="Xp"
            yForm$="3,3,4//UseMultiplier//SuppressMilli//DoCompact" 'ver115-4e

        case constImpedMag
            yAxisLabel$="Z ohms" : yLabel$="Z ohms"
            yForm$="3,3,4//UseMultiplier//SuppressMilli//DoCompact" 'ver115-4e

        case constSerR
            yAxisLabel$="Rs" : yLabel$="Rs"
            yForm$="3,3,4//UseMultiplier//SuppressMilli//DoCompact" 'ver115-4e

        case constParR
            yAxisLabel$="Rp" : yLabel$="Rp"
            yForm$="3,3,4//UseMultiplier//SuppressMilli//DoCompact" 'ver115-4e

        case constSerC
            yAxisLabel$="Cs" : yLabel$="Cs"
            yForm$="3,3,4//UseMultiplier//DoCompact"

        case constParC
            yAxisLabel$="Cp" : yLabel$="Cp"
            yForm$="3,3,4//UseMultiplier//DoCompact"

        case constSerL
            yAxisLabel$="Ls" : yLabel$="Ls"
            yForm$="3,3,4//UseMultiplier//DoCompact"

        case constParL
            yAxisLabel$="Lp" : yLabel$="Lp"
            yForm$="3,3,4//UseMultiplier//DoCompact"

        case constMagDBM
            if msaMode$="SA" then
                yAxisLabel$="Magnitude (dBm)" : yLabel$="dBm"
            else
                yAxisLabel$="Power (dBm)" : yLabel$="dBm" 'ver115-1i
            end if
            yForm$="####.###"    'ver115-5d

        case constMagWatts
            yAxisLabel$="Magnitude (Watts)" : yLabel$="Watts"
            yForm$="3,3,4//UseMultiplier//DoCompact"

        case constMagDB         'Only done for Transmission
            if msaMode$="ScalarTrans" then  'ver115-1a
                yAxisLabel$="Transmission (dB)"  : yLabel$="dB"
            else
                yAxisLabel$=Strans$;" dB"  : yLabel$=Strans$;" dB"
            end if
           yForm$="####.###"    'ver115-1e

        case constMagRatio  'Only done for TG mode transmission
            if msaMode$="ScalarTrans" then  'ver115-4f
                yAxisLabel$="Trans (Ratio)" : yLabel$="Ratio"
            else
                yAxisLabel$="Mag (Ratio)" : yLabel$="Ratio"
            end if
            yForm$="3,3,4//UseMultiplier//SuppressMilli//DoCompact" 'ver115-4e

        case constMagV
            yAxisLabel$="Mag (Volts)"  : yLabel$="Volts"
            yForm$="3,3,4//UseMultiplier//DoCompact"

        case constRho
            yAxisLabel$="Rho"  : yLabel$="Rho"
            yForm$="#.###"

        case constReturnLoss 'ver114-8d
            yAxisLabel$="RL"  : yLabel$="RL"
            yForm$="###.###"    'ver115-1e

        case constInsertionLoss  'ver114-8d
            yAxisLabel$="Insertion Loss(dB)"  : yLabel$="IL"
            yForm$="###.###"    'ver115-1e

        case constReflectPower  'ver115-2d
            yAxisLabel$="Reflect Pow(%)"  : yLabel$="Ref%"
            yForm$="###.##"

        case constComponentQ     'ver115-2d
            yAxisLabel$="Component Q"  : yLabel$="Q"
            yForm$="#####.#"

        case constSWR  'ver114-8d
            yAxisLabel$="SWR"  : yLabel$="SWR"
            yForm$="####.##"

        case constAdmitMag  'ver115-4a
            yAxisLabel$="Admit. (S)" : yLabel$="Y"
            yForm$="3,3,4//UseMultiplier//DoCompact"

        case constAdmitAng   'ver115-4a
            yAxisLabel$="Admit Deg" : yLabel$="Admit Deg"
            yIsPhase=1
            yForm$="#####.##"

        case constConductance  'ver115-4a
            yAxisLabel$="Conduct. (S)" : yLabel$="Conduct"
            yForm$="3,3,4//UseMultiplier//DoCompact"

        case constSusceptance  'ver115-4a
            yAxisLabel$="Suscep. (S)" : yLabel$="Suscep"
            yForm$="3,3,4//UseMultiplier//DoCompact"

        case constNoGraph   'ver115-2c
            yAxisLabel$="None"  : yLabel$="None"
            yForm$="####.##"    'Something valid, in case it gets mistakenly used

        case constAux0, constAux1, constAux2, constAux3, constAux4, constAux5
            auxNum=componConst-constAux0    'e.g. constAux4 produces 4
            yAxisLabel$=auxGraphDataFormatInfo$(auxNum,2)  : yLabel$=auxGraphDataFormatInfo$(auxNum,3)
            yForm$=auxGraphDataFormatInfo$(auxNum,1)

        case else
            yForm$="###.##"
            yAxisLabel$="Invalid"  : yLabel$="Invalid"
    end select
end sub



[ReadStep]'and put raw data bits into arrays. 'heavily modified ver116-1b
    nonPhaseMode=(msaMode$="SA") or (msaMode$="ScalarTrans")   'ver116-4e
    doingPhase= (suppressPhase=0) and (nonPhaseMode=0)   'ver116-4e
    if useAutoWait then
        wate=int(autoWaitTC+0.5)  'wait this much between measurements ver116-4j
        magIsStable=0
        if doingPhase then phaseIsStable=0 else phaseIsStable=1
        changePhaseADC=0
        changeMagADC=0
        repeatOnceMore=0
    end if
    for readStepCount=1 to 25    'ver116-1b added auto wait time procedures
        'If doing auto wait, repeat up to 25 times until readings become stable, as shown by comparing two
        'successive readings. If not doing auto wait, we bail out in the middle of the first pass.
        'if readStepCount=1 then readTime=uTickCount()  'DEBUG
        gosub [WaitStatement]'needs:wate,glitch variables,glitchtime ;slows program before reading data 'ver111-20b
        prevReadMagData=magdata 'Note if we are starting a new step, but not first of sweep, this is last step final data
        magdata = 0 'reset this variable before reading data
        'Read phase even in non-phase modes unless suppressPhase=1; we just don't process it in non-phase modes
        if suppressPhase=0 then 'ver116-4e
            UsbAdcControl.Adcs.struct = 3 ' USB: 15/08/10
            prevReadPhaseData=phadata
            gosub [ReadPhase]
            'and return with phadata(in bits). Also installed into pharray(thisstep,3).
            ' If serial AtoD, magdata is returned, but not installed in any array
            'if magdata is collected during [ReadPhase], skip Read Magnitude
        else
            phadata=0   'zero phase info if we are suppressing phase
            phaarray(thisstep,3)=0 'phadata
            phaarray(thisstep,4)=0 'pdm Read state
        end if
        'prevReadTime=readTime  'DEBUG
        'readTime=uTickCount()  'DEBUG
        if magdata = 0 then
            UsbAdcControl.Adcs.struct = 1 ' USB: 15/08/10
            gosub [ReadMagnitude]'and return with raw magdata bits 
        end if 'USB:05/12/2010

        magarray(thisstep,3) = magdata 'put raw data into array
        'the phadata could be in dead zone, but magnitude is still valid.
        'if in VNA Mode and PDM is in automatic, check for phasedata (bits) for limits
        'If magnitude is so low that phase is not valid and will be forced to zero, don't do PDM inversion.
        readStepDidInvert=0
        if doingPhase then
            if setpdm = 0 and suppressPDMInversion=0 and _
                            (phadata < pdmlowlim or phadata > pdmhighlim) then
                'Invert PDM and re-read after waiting. But if mag reading is too low for phase
                'to be valid, don't bother.
                if magdata>=validPhaseThreshold then readStepDidInvert=1 : gosub [InvertPDmodule] 'ver116-1b
            end if
        end if
            'Note InvertPDmodule will impose some wait time, but not very much in auto wait mode.
        if useAutoWait=0 then exit for
            'The rest of the loop is just for determining whether to repeat when doing auto wait
        if repeatOnceMore then
'            if thisstep>=45 and thisstep<=55 then    'For DEBUG
'                print "Final Repeat: ms=";magIsStable;" ps=";phaseIsStable;" magChange=";magdata-prevReadMagData;" phaChange=";phadata-prevReadPhaseData
'            end if
            exit for 'we just finished the extra repeat
        end if
            'Decide whether we need to repeat
'        if readStepCount=1 and thisstep>=45 and thisstep<=55 then    'For DEBUG
'            print "-----------";thisstep;"--------------"
'            print "First Read: wait=";wate;" ms=";magIsStable;" ps=";phaseIsStable;" mag=";magdata;" pha=";phadata; " Delay=";readTime-prevReadTime
'        end if

        'We want to keep reading until two successive reads are close to each other, or until the direction reverses.
        'We initially waited one half time constant of the magnitude filter, took a reading, waited,
        'and took a second reading. We determine here whether the changes were less than a predetermined value
        'Once mag or phase is determined to be stable, we flag it as stable so we don't have
        'to re-evaluate after the next read.
        directionReversal=0 'flag for reversal of sign of change
        lowLevelADC=0
        if thisstep<>sweepStartStep then
            'For the first read, we generate the change in readings from the readings left over
            'from the previous step, but only if this is not the first step.
            evaluateThisRead=1
        else
            'Always evaluate if not first read
            if readStepCount>1 then evaluateThisRead=1 else evaluateThisRead=0
        end if

        if evaluateThisRead then
            if magIsStable=0 then   'Evaluate mag change if mag not already deemed stable
                prevMagADCChange=changeMagADC  'save previous change to compare direction
                changeMagADC=magdata-prevReadMagData
                select case
                    case magdata<calADCofLowFringe
                        'For very low level signals, just repeat once more, but not if Wide filter
                        'We can never expect these to be perfectly stable.
                        lowLevelADC=1 : if videoFilter$<>"Wide" then repeatOnceMore=1
                        magIsStable=1 : phaseIsStable=1 'pretend. Note phase will be no good at this level so we don't evaluate it.
                    case magdata<calLowADCofCenterSlope
                        maxADCChange=autoWaitMaxChangeLowEndADC
                    case magdata>calHighADCofCenterSlope
                        maxADCChange=autoWaitMaxChangeHighEndADC
                    case else   'in center region
                        if changeMagADC<0 then
                            'If in center on first read but headed for low end, we may end up
                            'in the low end, so may want to use the minimum allowed change for center and low end
                            'The low end allowed change is likely much smaller than that for center.
                            if abs(changeMagADC)>(magdata-calLowADCofCenterSlope) then    'See whether another change like this gets us to low end
                                maxADCChange=min(autoWaitMaxChangeCenterADC,autoWaitMaxChangeLowEndADC)
                            else
                                maxADCChange=autoWaitMaxChangeCenterADC
                            end if
                        else
                            maxADCChange=autoWaitMaxChangeCenterADC
                        end if
                end select

                if abs(changeMagADC)<=maxADCChange then magIsStable=1
                if magIsStable=0 and lowLevelADC=0 and readStepCount>1 then
                    'Even if we don't have two close readings, we can stop when
                    'the direction of change reverses, which is a sign the change is being dominated by noise.
                    if (prevMagADCChange<0 and changeMagADC>0) _
                            or (prevMagADCChange>0 and changeMagADC<0) then
                        magIsStable=1
                        directionReversal=1
                    end if
                end if
            end if

                'Now do phase
            if phaseIsStable=0 then 'evaluate phase change if phase not already deemed stable
                'Note that if we just inverted phase, it is still valid to compare this phadata to the
                'previous one, because settling time will still be based on that change.
                if magdata<validPhaseThreshold then
                    'In this case we have no valid phase reading. If this is likely a final reading, then just deem
                    'phase to be stable.
                    if magIsStable then phaseIsStable=1 'can't actually evaluate low level phase
                else
                    prevPhaseADCChange=changePhaseADC
                    changePhaseADC=phadata-prevReadPhaseData
                    if abs(changePhaseADC)<=autoWaitMaxChangePhaseADC then
                        phaseIsStable=1
                    else
                        'Even if we don't have two close readings, we can stop when
                        'the direction of change reverses, which is a sign the change is being dominated by noise.
                        'Can't do this if we just inverted the PDM for this read, because that might cause
                        'reversal of sign of the change
                        if readStepDidInvert=0 and readStepCount>1 and _
                                ((prevPhaseADCChange<0 and changePhaseADC>0)  or (prevPhaseADCChange>0 and changePhaseADC<0)) then
                            phaseIsStable=1
                            directionReversal=1
                        end if
                    end if
                end if
            end if
'            if thisstep>=45 and thisstep<=55 then    'For DEBUG
'                print "Analysis "; readStepCount; ": ms=";magIsStable;" ps=";phaseIsStable;" mag=";magdata;" pha=";phadata;" magChange=";changeMagADC;" phaChange=";changePhaseADC;" Delay=";readTime-prevReadTime
'            end if
            if magIsStable and phaseIsStable then
                'If Precise, repeat one more
                'time and stop after that without comparing readings.
                'low level is already set to repeat once more unless Wide filter
                if (autoWaitPrecision$="Precise") then repeatOnceMore=1
                    'For lowLevelADC, we already set repeatOnceMore (and phase is meaningless)
                if lowLevelADC=0 and repeatOnceMore=0 then exit for
                    'Also repeat once more if ending because of direction reversal, other than
                    'direction reversal on second read.
                if directionReversal and readStepCount<3 then repeatOnceMore=1 else exit for
            end if
        end if
        'No point repeating if we aren't actually reading data. But we went through the motions above
        'for possible debugging use.
        if suppressHardware then exit for
    next readStepCount
    return

[ProcessAndPrintLastStep]
    rememberstep = thisstep 'remember where we were when entering this routine 'ver111-19
    'since we are processing and printing the previous step, use raw data in array(thisstep - sweepDir,data)

    if thisstep=sweepStartStep then     'Added by ver114-4m
        thisstep=sweepEndStep   'back up one and wrap around
    else
        thisstep=thisstep-sweepDir  'Back up one step; no wraparound to worry about
    end if
    gosub [ProcessAndPrint]'get raw data, process, print to the computer monitor ver111-22
    thisstep = rememberstep 'ver111-19
    return

[WaitStatement]'needed:wate,glitch()(p1,d1,p3,d3,pdm,hlt),glitchtime ; this slows the program 'ver111-27
    glitch = max(max(max(glitchp1, glitchd1),max(glitchp3, glitchd3)), max(glitchpdm, glitchhlt)) 'ver111-27
        'glitchp1=PLL1:glitchd1=DDS1:glitchp3=PLL3:glitchd3=DDS3:glitchpdm=PDM(10):glitchhlt=halted(10)
    waittime = wate + glitch   'number of ms we need to wait 'ver115-1i
        'in my Toshiba, a waittime count of 80 gives a delay of approx, 1 millisecond
        'therefore, each increment of any "glitchXX" or "wate" (Wait Box) should add 1 ms of delay before a "read"
    if waittime>0 then  'ver115-1i added the use of the system sleep function
        if doingInitialization or waittime<15 then
            'For short wait times we use our own timing loop
            'Also for initialization, when we are measuring glitchtime
            waittime=waittime*glitchtime
            timecounter = 0 'ver111-27
            [TimeLoop] 'ver111-27
            if timecounter < waittime then timecounter = timecounter + 1:goto [TimeLoop] 'ver111-27
        else
            'It is preferable to use the system Sleep function, but it operates in increments of 10-15 ms.
            'So it is only suitable for long wait times
            call uSleep waittime
        end if
    end if
    glitchp1=0:glitchd1=0:glitchp3=0:glitchd3=0:glitchpdm=0:glitchhlt=0 'reset glitch variables back to 0 'ver111-27
    return

[AutoGlitchtime] 'ver111-37c
    glitchtime = 100000
    whatiswate = wate
    wate = 1
    a = time$("ms") 'time of day, in milliseconds. This uses the computer's internal clock
    gosub [WaitStatement]
    b = time$("ms")
    glitchtime = glitchtime/(b-a) 'glitchtime is the value required for a 1 ms wait time
    wate = whatiswate 'change wate back to it's original global value
    return

[ReadMagnitude]'needed: port,status ; creates: magdata (and phadata for serial A/D's)
    if suppressHardware then magdata=0 : return 'ver115-6c
    if adconv =  8 then gosub [Read8Bitmag] 'and return here with magdata
    if adconv = 12 then gosub [Read12Bitmag] 'and return here with magdata
    if adconv = 16 then gosub [ReadAD16Status]:gosub [Process16Mag] 'combined ver111-34a
      'and return here with just magdata 'ver111-33b
    if adconv = 22 then gosub [ReadAD22Status]:gosub [Process22Mag] 'ver111-37a
      'and return here with just magdata 'ver111-37a
    return 'to [ReadStep]

[ReadPhase]'needed: port,status ; creates: phadata (and magdata for serial A/D's)
    if suppressHardware then 'ver115-6c
        phadata=0
    else
        select case adconv  'ver116-1b
            case 8
                gosub [Read8Bitpha] 'and return here with phadata
            case 12
                gosub [Read12Bitpha] 'and return here with phadata
            case else
                if adconv = 16 then gosub [ReadAD16Status]:gosub [Process16MagPha] 'combined ver111-34a
                if adconv = 22 then gosub [ReadAD22Status]:gosub [Process22MagPha] 'ver111-37a
        end select
    end if
    if doSpecialGraph>0 then phadata=maxpdmout/4 + thisstep*30   'SEWgraph Force to a value not requiring constant inversion
      'and return here with phadata (and magdata, if serial AtoD) 'ver111-33b
    'if calibrating the PDM inversion, don't put raw data into arrays, used only in [CalPDMinvdeg]
    if doingPDMCal = 1 then return 'to [CalPDMinvdeg] 'ver111-29 ver114-5L
    phaarray(thisstep,3) = phadata 'put raw data into array 'ver112-2a
    phaarray(thisstep,4) = phaarray(thisstep,0) 'PDM state at which this data is taken. ver112-2a
        'it is only used in Variables Windows to show state of PDM when data was collected.
    return 'to [ReadStep]




[CalcMagpowerPixel]
    power = power + freqerror
    if convdatapwr = 1 then gosub [ConvertDataToPower] 'ver112-2b
    'round off MSA input power to .01 dBm, magpower, no matter which AtoD is used
    magpower = power 'ver115-2d
        'Note if calInProgress=1, applyCalLevel will have been set to 0 by cal installation routine
    if applyCalLevel>0 then if (msaMode$<>"SA") then  magpower = magpower - lineCalArray(thisstep,1)  'ver116-4n  subtract reference.
    if magpower>=0 then magpower=int(magpower*100000+0.5)/100000 else magpower=int(magpower*100000-0.5)/100000 'round to five decimal places ver115-4d
    datatable(thisstep,2) = magpower    'put current power measurement into the array
    return 'to [ProcessAndPrint]

[CreateRcounter]'needed:reference,appxpdf ; creates:rcounter,pdf 'ver111-4
    rcounter = int(reference/appxpdf) 'ver111-4
    if (reference/appxpdf) - rcounter >= .5 then rcounter = rcounter + 1   'rounds off rcounter 'ver111-4
    pdf = reference/rcounter 'ver111-4
    return 'to (Initialize PLL 3),[InitializePLL2],or[InitializePLL1]with rcounter,pdf 'ver111-4

[CommandPLL1R]'needed:rcounter1,PLL1mode,PLL1phasepolarity,SELT,PLL1
    rcounter = rcounter1
    preselector = 32 : if PLL1mode = 1 then preselector = 16
    phasepolarity = PLL1phasepolarity    'inverting op amp is 0, non-inverting loop is 1
    fractional = PLL1mode       '0 for Integer-N; 1 for Fractional-N
    Jcontrol = SELT   'for PLL 1, on Control Board J1, the value is "3"
    LEPLL = 4         'for PLL 1, on Control Board J1, the value is "4"
    PLL = PLL1
    gosub [CommandRBuffer]'needs:rcounter,preselector,phasepolarity,fractional,Jcontrol,LEPLL,PLL
    if len(errora$)>0 then
        error$ = "PLL 1, " + errora$
        message$=error$ : call PrintMessage 'ver114-4e
        call RequireRestart   'ver115-1c
        wait
    end if
    return

[CommandPLL2R]'needed:reference,appxpdf,PLL2phasepolarity,SELT,PLL2
    preselector = 32
    phasepolarity = PLL2phasepolarity    'inverting op amp is 0, non-inverting loop is 1
    fractional = 0    '0 for Integer-N; PLL 2 should not be fractional due to increased noise
    Jcontrol = SELT   'for PLL 2, on Control Board J2, the value is "3"
    LEPLL = 8          'for PLL 2, on Control Board J2, the value is "8"
    PLL = PLL2
    gosub [CommandRBuffer]'needs:rcounter,preselector,phasepolarity,fractional,Jcontrol,LEPLL,PLL
    if len(errora$)>0 then
        error$ = "PLL 2, " + errora$
        message$=error$ : call PrintMessage 'ver114-4e
        call RequireRestart   'ver115-1c
        wait
    end if
    return 'to 'CommandPLL2R and Init Buffers

[CommandPLL3R]'needed:PLL3mode,PLL3phasepolarity,INIT,PLL3
    preselector = 32 : if PLL3mode = 1 then preselector = 16
    phasepolarity = PLL3phasepolarity    'inverting op amp is 0, non-inverting loop is 1
    fractional = PLL3mode       '0 for Integer-N; 1 for Fractional-N
    Jcontrol = INIT   'for Tracking Gen PLL, on Control Board J3, the value is "15"
    LEPLL = 16         'for Tracking Gen PLL, on Control Board J3, the value is "16"
    PLL = PLL3
    gosub [CommandRBuffer]'needs:rcounter,preselector,phasepolarity,fractional,Jcontrol,LEPLL,PLL
    if len(errora$)>0 then
        error$ = "PLL 3, " + errora$
        message$=error$ : call PrintMessage 'ver114-4e
        call RequireRestart   'ver115-1c
        wait
    end if
    return 'to 'CommandPLL3R and Init Buffers

[CommandRBuffer]'needed:rcounter,preselector,phasepolarity,fractional,Jcontrol,LEPLL,PLL
    if PLL = 2325 then gosub [Command2325R]'needs:rcounter,preselector,Jcontrol,port,LEPLL,contclear ; commands LMX2325 rcounter and registers
    if PLL = 2326 then gosub [Command2326R]'needs:rcounter,phasepolarity,Jcontrol,port,LEPLL,contclear ; commands LMX2326 rcounter and registers
    if PLL = 2350 then gosub [Command2350R]'needs:rcounter,phasepolarity,Jcontrol,port,LEPLL,contclear,fractional ; commands LMX2350 rcounter
    if PLL = 2353 then gosub [Command2353R]'needs:rcounter,phasepolarity,Jcontrol,port,LEPLL,contclear,fractional ; commands LMX2353 rcounter
    if PLL = 4112 then gosub [Command4112R]'needs:rcounter,preselector,phasepolarity,Jcontrol,port,LEPLL,contclear ; commands AD4112 rcounter
    return

[CreateIntegerNcounter]'needed:appxVCO,reference,rcounter ; creates:ncount,ncounter,fcounter(0),pdf
    ncount = appxVCO/(reference/rcounter)  'approximates the Ncounter for PLL
    ncounter = int(ncount)     'approximates the ncounter for PLL
    if ncount - ncounter >= .5 then ncounter = ncounter + 1   'rounds off ncounter
    fcounter = 0
    pdf = appxVCO/ncounter        'actual phase freq of PLL
    return  'to 'CreatePLL2N,'[CalculateThisStepPLL1],or '[CalculateThisStepPLL3] with ncount, ncounter and fcounter(=0)

[CreateFractionalNcounter]'needed:appxVCO,reference,rcounter ; creates:ncount,ncounter,fcounter,pdf
    ncount = appxVCO/(reference/rcounter)  'approximates the Ncounter for PLL
    ncounter = int(ncount)    'actual value for PLL Ncounter
    fcount = ncount - ncounter
    fcounter = int(fcount*16) 'ver111
    if (fcount*16) - fcounter >= .5 then fcounter = fcounter + 1 'rounds off fcounter  ver111
    if fcounter = 16 then ncounter = ncounter + 1:fcounter = 0
    pdf = appxVCO/(ncounter + (fcounter/16)) 'actual phase freq for PLL 'ver111-10
    return  'with ncount,ncounter,fcounter,pdf
[AutoSpur]'needed:LO1,LO2,finalfreq,appxdds1,dds1output,rcounter1,finalbw,fcounter,ncounter,spurcheck;changes pdf,dds1output
    '[AutoSpur] is a continuation of [CreateFractionalNcounter], used only in MSA when PLL 1 is Fractional
    spur = 0    'reset spur, and determine if there is potential for a spur
    firstif = LO2 - finalfreq
    fractionalfreq = dds1output/(rcounter1*16)
    harnonicb = int(firstif/fractionalfreq)
    if (firstif/fractionalfreq)-harnonicb >=.5 then harnonicb = harnonicb + 1  'rev108
    harnonica = harnonicb - 1
    harnonicc = harnonicb + 1
    firstiflow = LO2 - (finalfreq + finalbw/1000)
    firstifhigh = LO2 - (finalfreq - finalbw/1000)
    if harnonica*fractionalfreq > firstiflow and harnonica*fractionalfreq < firstifhigh then spur = 1
    if harnonicb*fractionalfreq > firstiflow and harnonicb*fractionalfreq < firstifhigh then spur = 1
    if harnonicc*fractionalfreq > firstiflow and harnonicc*fractionalfreq < firstifhigh then spur = 1
    if spur = 1 and (dds1output<appxdds1) then fcounter = fcounter - 1
    if spur = 1 and (dds1output>appxdds1) then fcounter = fcounter + 1
    if fcounter = 16 then ncounter = ncounter + 1:fcounter = 0  'rev108
    if fcounter <0 then ncounter = ncounter - 1:fcounter = 15  'rev108
    pdf = LO1/(ncounter + (fcounter/16))
    dds1output = pdf * rcounter1    'actual output of DDS1(input Ref to PLL1)
    return 'with possibly new ncounter,fcounter,pdf,dds1output
[ManSpur]'needed:spurcheck,dds1output,appxdds1,fcounter,ncounter
    '[ManSpur] is a continuation of [CreateFractionalNcounter], used only in MSA when PLL 1 is Fractional
    if spurcheck = 1 and (dds1output<appxdds1) then fcounter = fcounter - 1 'causes +shift in pdf1
    if spurcheck = 1 and (dds1output>appxdds1) then fcounter = fcounter + 1 'causes -shift in pdf1
    if fcounter = 16 then ncounter = ncounter + 1:fcounter = 0  'rev108
    if fcounter < 0 then ncounter = ncounter - 1:fcounter = 15  'rev108
    pdf = LO1/(ncounter + (fcounter/16))
    dds1output = pdf * rcounter1    'actual output of DDS1(input Ref to PLL1)
    return 'with possibly new:ncounter,fcounter,pdf,dds1output

[CreatePLL1N]'needed:ncounter,fcounter,PLL1mode,PLL1
    preselector = 32 : if PLL1mode = 1 then preselector = 16
    PLL = PLL1
    gosub [CreateNBuffer]'needs:ncounter,fcounter,PLL,preselector;creates:Bcounter,Acounter, and N Bits N0-Nx
    if len(errora$)>0 then
        error$ = "PLL 1, " + errora$
        message$=error$ : call PrintMessage 'ver114-4e
        call RequireRestart   'ver115-1c
        wait
    end if
    Bcounter1=Bcounter: Acounter1=Acounter
    return 'returns with Bcounter1,Acounter1,N0thruNx

[CreatePLL2N]'needed:ncounter,fcounter,PLL2
    preselector = 32
    PLL = PLL2
    gosub [CreateNBuffer]'needs:ncounter,fcounter,PLL,preselector;creates:Bcounter,Acounter, and N Bits N0-Nx
    if len(errora$)>0 then
        error$ = "PLL 2, " + errora$
        message$=error$ : call PrintMessage 'ver114-4e
        call RequireRestart   'ver115-1c
        wait
    end if
    return 'to 'CreatePLL2N

[CreatePLL3N]'needed:ncounter,fcounter,PLL3mode,PLL3  ver111-14
    preselector = 32 : if PLL3mode = 1 then preselector = 16
    PLL = PLL3
    gosub [CreateNBuffer]'needs:ncounter,fcounter,PLL,preselector;creates:Bcounter,Acounter, and N Bits N0-Nx
    if len(errora$)>0 then
        error$ = "PLL 3, " + errora$
        message$=error$ : call PrintMessage 'ver114-4e
        call RequireRestart   'ver115-1c
        wait
    end if
    Bcounter3=Bcounter: Acounter3=Acounter
    return 'returns with Bcounter3,Acounter3,N0thruNx

[CreateNBuffer]'needed:PLL,ncounter,fcounter,preselector
    if PLL = 2325 then gosub [Create2325N]'needs:ncounter,preselector; creates LMX2325 N Buffer ver111
    if PLL = 2326 then gosub [Create2326N]'needs:ncounter ; creates LMX2326 N Buffer ver111
    if PLL = 2350 then gosub [Create2350N]'needs:ncounter,preselector,fcounter; creates LMX2350 RFN Buffer ver111
    if PLL = 2353 then gosub [Create2353N]'needs: ncounter,preselector,fcounter; creates LMX2353 N Buffer ver111
    if PLL = 4112 then gosub [Create4112N]'needs:ncounter,preselector; creates AD4112 N Buffer ver111
    return 'with Bcounter,Acounter, and N Bits N0-N23

[Command2325R]'needed:rcounter,preselector,control,Jcontrol,port,LEPLL,contclear ; commands LMX2325 rcounter and registers
    if rcounter <3 then beep:errora$ = "2325 Rcounter is < 3":return 'with errora$ ver111-37c
    'ver116-4o deleted "if" block, per Lrev1
    N0 = 1           'address bit, 0 sets the N Buffer, 1 is for R Buffer
    rc1 = int(rcounter/2):N1 = rcounter - 2*rc1 'binary conversion from decimal
    rc2 = int(rc1/2):N2 = rc1 - 2*rc2
    rc3 = int(rc2/2):N3 = rc2 - 2*rc3
    rc4 = int(rc3/2):N4 = rc3 - 2*rc4
    rc5 = int(rc4/2):N5 = rc4 - 2*rc5
    rc6 = int(rc5/2):N6 = rc5 - 2*rc6
    rc7 = int(rc6/2):N7 = rc6 - 2*rc7
    rc8 = int(rc7/2):N8 = rc7 - 2*rc8
    rc9 = int(rc8/2):N9 = rc8 - 2*rc9
    rc10 = int(rc9/2):N10 = rc9 - 2*rc10
    rc11 = int(rc10/2):N11 = rc10 - 2*rc11
    rc12 = int(rc11/2):N12 = rc11 - 2*rc12
    rc13 = int(rc12/2):N13 = rc12 - 2*rc13
    rc14 = int(rc13/2):N14 = rc13 - 2*rc14
    N15 = 1: if preselector = 64 then N15 = 0   'sets preselector divide ratio, 1=32, 0=64
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
    gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111
    return

[Create2325N]'needed:ncounter,preselector; creates LMX2325 n buffer
    Bcounter = int(ncounter/preselector)
    Acounter = ncounter- (Bcounter * preselector)
    if Bcounter<3 then beep:errora$ = "2325 Bcounter < 3":return 'with errora$ ver111-37c
    if Bcounter>2047 then beep:errora$ = "2325 Bcounter > 2047":return 'with errora$ ver111-37c
    if Bcounter<Acounter then beep:errora$ = "2325 Bcounter<Acounter":return 'with errora$ ver111-37c
    'ver116-4o deleted "if" block, per Lrev1
    N0 = 0    'address bit, 0 sets the N Buffer, 1 is for R Buffer
    na1 = int(Acounter/2):N1 = Acounter - 2*na1 'binary conversion from decimal
    na2 = int(na1/2):N2 = na1 - 2*na2
    na3 = int(na2/2):N3 = na2 - 2*na3
    na4 = int(na3/2):N4 = na3 - 2*na4
    na5 = int(na4/2):N5 = na4 - 2*na5
    na6 = int(na5/2):N6 = na5 - 2*na6
    na7 = int(na6/2):N7 = na6 - 2*na7
    nb8 = int(Bcounter/2):N8 = Bcounter - 2*nb8
    nb9 = int(nb8/2):N9 = nb8 - 2*nb9
    nb10 = int(nb9/2):N10 = nb9 - 2*nb10
    nb11 = int(nb10/2):N11 = nb10 - 2*nb11
    nb12 = int(nb11/2):N12 = nb11 - 2*nb12
    nb13 = int(nb12/2):N13 = nb12 - 2*nb13
    nb14 = int(nb13/2):N14 = nb13 - 2*nb14
    nb15 = int(nb14/2):N15 = nb14 - 2*nb15
    nb16 = int(nb15/2):N16 = nb15 - 2*nb16
    nb17 = int(nb16/2):N17 = nb16 - 2*nb17
    nb18 = int(nb17/2):N18 = nb17 - 2*nb18
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
    return

[Command2326R]'needed:rcounter,phasepolarity,control,Jcontrol,port,LEPLL,contclear ; commands LMX2326 rcounter and registers
    '[Create2326InitBuffer]'need phasepolarity
    'ver116-4o deleted "if" block, per Lrev1
    N20=0     'Test, use 0
    N19=0     '1=Power Down Mode, use 0
    N18=0     'Test, use 0
    N17=0     'Test, use 0
    N16=0     'Test, use 0
    N15=0     'Fastlock Time out value, use 0
    N14=0     'Fastlock Time out value, use 0
    N13=0     'Fastlock Time out value, use 0
    N12=0     'Fastlock Time out value, use 0
    N11=0     '1=Time out enable, use 0
    N10=0     'Fastlock control, use 0
    N9=0     '1=Fastlock enable, use 0
    N8=0     '1=Tristate the phase det output, use 0
    N7 = phasepolarity     'Phase det polarity, 1=pos  0=neg
    N6=0        'FoLD control(pin14 output), 0= tristate, 1= R Divider out
    N5=0        '2= N Divider out, 3= Serial Data Output, 4= Digital Lock Detect
    N4=0        '5= Open drain lock detect, 6= High output, 7= Low output
    N3=0        '1= Power Down, use 0
    N2=0        '1= Counter Reset Enable, allows reset of R,N counters,use 0
    N1=1        'F1 address bit 1, must be 1
    N0=1        'F1 address bit 0, must be 1
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 ''ver116-4o per Lrev1
  '[Command2326InitBuffer]'need Jcontrol,LEPLL,contclear
    gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111
  '[Create2326Rbuffer]'need rcounter
    if rcounter <3 then beep:errora$="2326 R counter <3":return 'with errora$ ver111-37c
    if rcounter >16383 then beep:errora$="2326 R counter >16383":return 'with errora$ ver111-37c
    'ver116-4o deleted "if" block, per Lrev1
    N0 = 0                   'R address bit 0, must be 0
    N1 = 0                   'R address vit 1, must be 0
    ra0 = int(rcounter/2):N2 = rcounter- 2*ra0    'LSB
    ra1 = int(ra0/2):N3 = ra0- 2*ra1
    ra2 = int(ra1/2):N4 = ra1- 2*ra2
    ra3 = int(ra2/2):N5 = ra2- 2*ra3
    ra4 = int(ra3/2):N6 = ra3- 2*ra4
    ra5 = int(ra4/2):N7 = ra4- 2*ra5
    ra6 = int(ra5/2):N8 = ra5- 2*ra6
    ra7 = int(ra6/2):N9 = ra6- 2*ra7
    ra8 = int(ra7/2):N10 = ra7- 2*ra8
    ra9 = int(ra8/2):N11 = ra8- 2*ra9
    ra10 = int(ra9/2):N12 = ra9- 2*ra10
    ra11 = int(ra10/2):N13 = ra10- 2*ra11
    ra12 = int(ra11/2):N14 = ra11- 2*ra12
    ra13 = int(ra12/2):N15 = ra12- 2*ra13  'MSB
    N16 = 0     'Test Bit
    N17 = 0     'Test Bit
    N18 = 0     'Test Bit
    N19 = 0     'Test Bit
    N20 = 0     'Lock Detector Mode, 0=3 refcycles, 1=5 cycles
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
  '[Command2326Rbuffer]'need Jcontrol,LEPLL,contclear
    gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111
    return

[Create2326N]'needed:ncounter ; creates LMX2326 n buffer  ver111
    Bcounter = int(ncounter/32)
    Acounter = int(ncounter-(Bcounter*32))
    if Bcounter < 3 then beep:errora$="2326 Bcounter <3":return 'with errora$ ver111-37c
    if Bcounter > 8191 then beep:errora$="2326 Bcounter >8191":return 'with errora$ ver111-37c
    if Bcounter < Acounter then beep:errora$="2326 Bcounter<Acounter":return 'with errora$ ver111-37c
    'ver116-4o deleted "if" block, per Lrev1
    N0 = 1       'n address bit 0, must be 1
    N1 = 0       'n address bit 1, must be 0
    na0 = int(Acounter/2):N2 = Acounter- 2*na0      'Acounter bit 0 LSB
    na1 = int(na0/2):N3 = na0 - 2*na1
    na2 = int(na1/2):N4 = na1 - 2*na2
    na3 = int(na2/2):N5 = na2 - 2*na3
    na4 = int(na3/2):N6 = na3 - 2*na4               'Acounter bit 4 MSB
    nb0 = int(Bcounter/2):N7 = Bcounter- 2*nb0      'Bcounter bit 0 LSB
    nb1 = int(nb0/2):N8 = nb0 - 2*nb1
    nb2 = int(nb1/2):N9 = nb1 - 2*nb2
    nb3 = int(nb2/2):N10 = nb2 - 2*nb3
    nb4 = int(nb3/2):N11 = nb3 - 2*nb4
    nb5 = int(nb4/2):N12 = nb4 - 2*nb5
    nb6 = int(nb5/2):N13 = nb5 - 2*nb6
    nb7 = int(nb6/2):N14 = nb6 - 2*nb7
    nb8 = int(nb7/2):N15 = nb7 - 2*nb8
    nb9 = int(nb8/2):N16 = nb8 - 2*nb9
    nb10 = int(nb9/2):N17 = nb9 - 2*nb10
    nb11 = int(nb10/2):N18 = nb10 - 2*nb11
    nb12 = int(nb11/2):N19 = nb11 - 2*nb12          'Bcounter bit 12 MSB
    N20 = 1    'Phase Det Current, 1= 1 ma, 0= 250 ua
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
    return

[Command2350R]'needed: rcounter,phasepolarity,control,Jcontrol,port,LEPLL,contclear,fractional ; commands LMX2350 rcounter
    '[CreateIFRbuffer2350]'needed:nothing,since IF section is turned off
    'ver116-4o deleted "if" block, per Lrev1
    N23=0     'osc. 0=separate
    N22=1     'Modulo F, 1=16 0=15
    N21=1     'ifr21-ifr19 is FO/LD, 3 Bits (0-7), MSB, 0=IF/RF alogLockDet(open drain)
    N20=1     '1=IF digLockDet, 2=RF digLockDet, 3=IF/RF digLockDet
    N19=1     '4=IF Rcntr, 5=IF Ncntr, 6=RF Rcntr, 7=RF Ncntr, LSB
    N18=0     'IF charge pump, 0=100ua  1=800ua
    N17=1     'IF polarity 1=positive phase action
    N16=0     'IFR counter IF section 15 Bits, MSB 14
    N15=0     'IFRcounter Bit 13
    N14=0     'IFRcounter Bit 12
    N13=0     'IFRcounter Bit 11
    N12=1     'IFRcounter Bit 10
    N11=1     'IFRcounter Bit 9
    N10=1     'IFRcounter Bit 8
    N9=1      'IFRcounter Bit 7
    N8=0      'IFRcounter Bit 6
    N7=1      'IFRcounter Bit 5
    N6=1      'IFRcounter Bit 4
    N5=0      'IFRcounter Bit 3
    N4=0      'IFRcounter Bit 2
    N3=0      'IFRcounter Bit 1
    N2=0      'IFR counter, IF section 15 Bits, LSB 0
    N1=0      '2350 IF_R register, 2 bits, must be 0
    N0=0      '2350 IF_R register, 2 bits, must be 0
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
  '[CommandIFRbuffer2350]
    gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111
  '[CreateIFNbuffer2350]'needed:nothing,since IF section is turned off(N22=1)
    'ver116-4o deleted "if" block, per Lrev1
    N23=0     'IF counter reset, 0=normal operation
    N22=1     'Power down mode for IF section, 1=powered down, 0=powered up
    N21=0     'PWN Mode,  0=async  1=syncro
    N20=0     'Fastlock, 0=CMOS outputs enabled 1= fastlock mode
    N19=0     'test bit, leave at 0
    N18=1     'OUT 0,  1
    N17=0     'OUT 1,  0
    N16=0     'IF N Bcounter 12 Bits MSB bit 11
    N15=0     'IF N Bcounter, bit 10, '512 = 0010 0000 0000
    N14=1     'IF N Bcounter, bit 9
    N13=0     'IF N Bcounter, bit 8
    N12=0     'IF N Bcounter, bit 7
    N11=0     'IF N Bcounter, bit 6
    N10=0     'IF N Bcounter, bit 5
    N9=0      'IF N Bcounter, bit 4
    N8=0      'IF N Bcounter, bit 3
    N7=0      'IF N Bcounter, bit 2
    N6=0      'IF N Bcounter, bit 1
    N5=0      'IF N Bcounter, 12 Bits, LSB bit 0
    N4=0      'bit 2, IF N Acounter 3 Bits MSB
    N3=0      'bit 1, 0 = 000 thru 7 = 111
    N2=0      'bit 0, IF N Acounter 3 Bits LSB
    N1=0      '2350 IF_N register, 2 bits, must be 0
    N0=1      '2350 IF_N register, 2 bits, must be 1
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
  '[CommandIFNbuffer2350]
    gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111
  '[CreateRFRbuffer2350]needed:rcounter,phasepolarity,fractional
    if rcounter < 3 then beep:errora$="2350 Rcounter <3":return 'with errora$ ver111-37c
    if rcounter > 32767 then beep:errora$="2350 Rcounter >32767":return 'with errora$ ver111-37c
    'ver116-4o deleted "if" block, per Lrev1
    N0=0      '2350 RF_R register, 2 bits, must be 0
    N1=1      '2350 RF_R register, 2 bits, must be 1
    rfra2 = int(rcounter/2):N2 = rcounter- 2*rfra2
    rfra3 = int(rfra2/2):N3 = rfra2- 2*rfra3
    rfra4 = int(rfra3/2):N4 = rfra3- 2*rfra4
    rfra5 = int(rfra4/2):N5 = rfra4- 2*rfra5
    rfra6 = int(rfra5/2):N6 = rfra5- 2*rfra6
    rfra7 = int(rfra6/2):N7 = rfra6- 2*rfra7
    rfra8 = int(rfra7/2):N8 = rfra7- 2*rfra8
    rfra9 = int(rfra8/2):N9 = rfra8- 2*rfra9
    rfra10 = int(rfra9/2):N10 = rfra9- 2*rfra10
    rfra11 = int(rfra10/2):N11 = rfra10- 2*rfra11
    rfra12 = int(rfra11/2):N12 = rfra11- 2*rfra12
    rfra13 = int(rfra12/2):N13 = rfra12- 2*rfra13
    rfra14 = int(rfra13/2):N14 = rfra13- 2*rfra14
    rfra15 = int(rfra14/2):N15 = rfra14- 2*rfra15
    rfra16 = int(rfra15/2):N16 = rfra15- 2*rfra16
    N17 = phasepolarity     'RF phase polarity,  1=positive action, 0=inverted action
    N18=1     'LSB of RF charge pump sel, 4 Bits, 16 levels, 100ua/level
    N19=1     'total current = (100ua * bit value)+100ua
    N20=1     '100ua to 1600ua: ie, 800ua = 0111, 1600ua = 1111
    N21=1     'MSB of RF charge pump sel, 4 Bits 100ua/bit
    N22=0     'V2 enable voltage doubler =1   0=norm Vcc
    N23 = fractional   'DLL mode, delay line cal, 0=slow  1=fast,fractional mode
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 ''ver116-4o per Lrev1
    '[CommandRFRbuffer2350]
    gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111
    return

[Create2350N]'needed: ncounter,preselector,fcounter; creates LMX2350 RFN Buffer
    Bcounter = int(ncounter/preselector)
    Acounter = int(ncounter-(Bcounter*preselector))
    if Bcounter < 3 then beep:errora$="2350 Bcounter <3":return 'with errora$ ver111-37c
    if Bcounter > 1023 then beep:errora$="2350 Bcounter >1023":return 'with errora$ ver111-37c
    if Bcounter < Acounter + 2 then beep:errora$="2350 Bcounter<Acounter+2":return 'with errora$ ver111-37c
    'ver116-4o deleted "if" block, per Lrev1
    N0=1      '2350 RF_N register, must be 1
    N1=1      '2350 RF_N register, must be 1 'was N=1 ver113-7a
    f0 = int(fcounter/2):N2 = fcounter - 2*f0      'fcounter bit 0
    f1 = int(f0/2):N3 = f0 - 2*f1       'fcounter bit 1
    f2 = int(f1/2):N4 = f1 - 2*f2       'fcounter bit 2
    f3 = int(f2/2):N5 = f2 - 2*f3       'fcounter bit 3 (0 to 15)
    rfna6 = int(Acounter/2):N6 = Acounter- 2*rfna6
    rfna7 = int(rfna6/2):N7 = rfna6 - 2*rfna7
    rfna8 = int(rfna7/2):N8 = rfna7 - 2*rfna8
    rfna9 = int(rfna8/2):N9 = rfna8 - 2*rfna9
    rfna10 = int(rfna9/2):N10 = rfna9 - 2*rfna10
    rfnb11 = int(Bcounter/2):N11 = Bcounter- 2*rfnb11
    rfnb12 = int(rfnb11/2):N12 = rfnb11 - 2*rfnb12
    rfnb13 = int(rfnb12/2):N13 = rfnb12 - 2*rfnb13
    rfnb14 = int(rfnb13/2):N14 = rfnb13 - 2*rfnb14
    rfnb15 = int(rfnb14/2):N15 = rfnb14 - 2*rfnb15
    rfnb16 = int(rfnb15/2):N16 = rfnb15 - 2*rfnb16
    rfnb17 = int(rfnb16/2):N17 = rfnb16 - 2*rfnb17  'was rgb17 ver113-7a
    rfnb18 = int(rfnb17/2):N18 = rfnb17 - 2*rfnb18
    rfnb19 = int(rfnb18/2):N19 = rfnb18 - 2*rfnb19
    rfnb20 = int(rfnb19/2):N20 = rfnb19 - 2*rfnb20
    N21=0 :if preselector = 32 then N21 = 1  '0=16/17  1=32/33
    N22=0     'Pwr down RF,    0=normal  1=pwr down
    N23=0     'RF cntr reset,  0=normal  1=reset
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
    return

[Command2353R]'needed: rcounter,phasepolarity,control,Jcontrol,port,LEPLL,contclear,fractional ; commands LMX2353 rcounter
  '[Create2353F1Buffer]'globals reqd, none
    'ver116-4o deleted "if" block, per Lrev1
    N23=0
    N22=1     'divider, 1=16 0=15
    N21=0     'FO/LD output selection, 3 Bits 0-7 MSB
    N20=0     '0=alog lock det, 2=dig lock det
    N19=0     '6=Ndivider output, 7=Rdivider output
    N18=0:N17=0:N16=0:N15=0:N14=0
    N13=0:N12=0:N11=0:N10=0:N9=0
    N8=0:N7=0:N6=0:N5=0:N4=0
    N3=0:N2=0
    N1=0        'F1 address bit 1
    N0=0        'F1 address bit 0
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev11
  '[Command2353F1Buffer]
    gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111
  '[Create2353F2Buffer]'globals reqd: none
    'ver116-4o deleted "if" block, per Lrev1
    N23=0:N22=0
    N21=0     'Power Down Mode,  0=async  1=syncro
    N20=0     'Fastlock, 0=CMOS outputs enabled 1= fastlock mode
    N19=0     'test bit, leave at 0
    N18=0     'OUT 1,  0
    N17=0     'OUT 0,  0
    N16=0:N15=0:N14=0:N13=0
    N12=0:N11=0:N10=0:N9=0
    N8=0:N7=0:N6=0:N5=0
    N4=0:N3=0:N2=0
    N1=0        'F2 address bit 1
    N0=1        'F2 address bit 0
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
  '[Command2353F2Buffer]
    gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111
  '[Create2353Rbuffer]'needed:rcounter,phasepolarity,fractional
    if rcounter <3 then beep:errora$ = "2353 Rcounter is < 3":return 'with errora$ ver111-37c
    if rcounter >32767 then beep:errora$ = "2353 Rcounter is > 32767":return 'with errora$ ver111-37c
    'ver116-4o deleted "if" block, per Lrev1
    N0 = 0                   'R address bit 0
    N1 = 1                   'R address bit 1
    ra0 = int(rcounter/2):N2 = rcounter- 2*ra0    'LSB R buffer
    ra1 = int(ra0/2):N3 = ra0- 2*ra1:ra2 = int(ra1/2):N4 = ra1- 2*ra2
    ra3 = int(ra2/2):N5 = ra2- 2*ra3:ra4 = int(ra3/2):N6 = ra3- 2*ra4
    ra5 = int(ra4/2):N7 = ra4- 2*ra5:ra6 = int(ra5/2):N8 = ra5- 2*ra6
    ra7 = int(ra6/2):N9 = ra6- 2*ra7:ra8 = int(ra7/2):N10 = ra7- 2*ra8
    ra9 = int(ra8/2):N11 = ra8- 2*ra9:ra10 = int(ra9/2):N12 = ra9- 2*ra10
    ra11 = int(ra10/2):N13 = ra10- 2*ra11:ra12 = int(ra11/2):N14 = ra11- 2*ra12
    ra13 = int(ra12/2):N15 = ra12- 2*ra13:ra14 = int(ra13/2):N16 = ra13- 2*ra14    'MSB R buffer
    N17 = phasepolarity     'phase detector polarity 1=normal,0=reverse for opamp
    N18 = 1   'LSB of Charge pump control, 100ua x1 +100ua
    N19 = 1          'Charge pump control, 100ua x2 +100ua
    N20 = 1          'Charge pump control, 100ua x4 +100ua
    N21 = 1   'MSB of Charge pump control, 100ua x8 +100ua
    N22 = 0   'Charge Pump Voltage Doubler Enabled when 1
    N23 = fractional 'Delay Line Loop Cal mode, set to 1 for fractional N
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
  '[Cmd2353Rbuffer]
    gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111
    return

[Create2353N]'needed: ncounter,preselector,fcounter; creates LMX2353 N Buffer
    Bcounter = int(ncounter/preselector)
    Acounter = int(ncounter-(Bcounter*preselector))
    if Bcounter < 3 then beep:errora$ = "2353 Bcounter is < 3":return 'with errora$ ver111-37c
    if Bcounter > 1023 then beep:errora$ = "2353 Bcounter is > 1023":return 'with errora$ ver111-37c
    if Bcounter < Acounter + 2 then beep:errora$ = "2353 Bcounter < Acounter+2":return 'with errora$ ver111-37c
    'ver116-4o deleted "if" block, per Lrev1
    N0 = 1       'n address bit 0
    N1 = 1       'n address bit 1
    f0 = int(fcounter/2):N2 = fcounter - 2*f0       'fcounter bit 0
    f1 = int(f0/2):N3 = f0 - 2*f1       'fcounter bit 1
    f2 = int(f1/2):N4 = f1 - 2*f2       'fcounter bit 2
    f3 = int(f2/2):N5 = f2 - 2*f3       'fcounter bit 3 (0 to 15)
    na0 = int(Acounter/2):N6 = Acounter- 2*na0      'Acounter bit 0 LSB
    na1 = int(na0/2):N7 = na0 - 2*na1
    na2 = int(na1/2):N8 = na1 - 2*na2
    na3 = int(na2/2):N9 = na2 - 2*na3
    na4 = int(na3/2):N10 = na3 - 2*na4      'Acounter bit 4 MSB
    nb0 = int(Bcounter/2):N11 = Bcounter- 2*nb0      'Bcounter bit 0 LSB
    nb1 = int(nb0/2):N12 = nb0 - 2*nb1
    nb2 = int(nb1/2):N13 = nb1 - 2*nb2
    nb3 = int(nb2/2):N14 = nb2 - 2*nb3
    nb4 = int(nb3/2):N15 = nb3 - 2*nb4
    nb5 = int(nb4/2):N16 = nb4 - 2*nb5
    nb6 = int(nb5/2):N17 = nb5 - 2*nb6
    nb7 = int(nb6/2):N18 = nb6 - 2*nb7
    nb8 = int(nb7/2):N19 = nb7 - 2*nb8
    nb9 = int(nb8/2):N20 = nb8 - 2*nb9      'Bcounter bit 9 MSB
    N21 = 0 :if preselector = 32 then N21 = 1  '0=16/17  1=32/33
    N22 = 0          'power down if 1
    N23 = 0          'counter reset if 1
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
    return

[Command4112R]'needed: rcounter,preselector,phasepolarity,control,Jcontrol,port,LEPLL,contclear ; commands AD4112 rcounter
  '[Create4112InitBuffer]'needed:preselector,phasepolarity
    'ver116-4o deleted "if" block, per Lrev1
    N23=1     'N23,22 prescaler: 0=8, 1=16, 2=32, 3=64
    N22=0     'preselector defaulted to 32
    if preselector =8 then N23=0:N22=0
    if preselector =16 then N23=0:N22=1
    if preselector =64 then N23=1:N22=1
    N21=0     'Power Down Mode, 0=async, 1=sync  use 0
    N20=0     'N20,19,18 Phase Current for Set 2            '12-3-10
    N19=0     'current= min current + min current*bit value '12-3-10
    N18=1     'use bit value of 1 and 4.7 Kohm for 1.25 ma  '12-3-10
    N17=0     'N17,16,15 Phase Current for Set 1            '12-3-10
    N16=0     'current= min current + min current*bit value '12-3-10
    N15=1     'use bit value of 1 and 4.7 Kohm for 1.25 ma  '12-3-10
    N14=0     'N15,14,13,12 Fastlock Timer cycles
    N13=0     '4 Bits, Cycles= 3 cycles + 4*bit value
    N12=0     'Fastlock Time out value, use 0
    N11=0     'use 4 bit value = 0
    N10=0     '0=Fastlock Mode 1 (command), 1=Mode 2 (automatic)
    N9=0     '1=Fastlock enabled, 0 =Fastlock Disabled
    N8=0      '1=Tristate the phase det output, use 0
    N7 = phasepolarity     'Phase det polarity, 1=pos  0=neg
    N6=0      'FoLD control(pin14 output), 0= tristate, 1= Digital Lock Detect
    N5=0      '2= N Divider out, 3= High output, 4= R Divider output
    N4=0      '5= Open drain lock detect, 6= Serial Data output, 7= Low output
    N3=0      'PD1, Power Down, 0=normal operation, 1=select power down mode
    N2=0      '1= Counter Reset Enable, allows reset of R,N counters,use 0
    N1=1      'F1 address bit 1, must be 1
    N0=1      'F1 address bit 0, must be 1
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
  '[Command4112InitBuffer]
    gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111
  '[Create4112Rbuffer]'needs:rcounter
    if rcounter >16383 then beep:errora$="4112 R counter >16383":return 'with errora$ ver111-37c
    'ver116-4o deleted "if" block, per Lrev1
    N0 = 0                   'R address bit 0, must be 0
    N1 = 0                   'R address vit 1, must be 0
    ra0 = int(rcounter/2):N2 = rcounter- 2*ra0    'LSB R0
    ra1 = int(ra0/2):N3 = ra0- 2*ra1
    ra2 = int(ra1/2):N4 = ra1- 2*ra2
    ra3 = int(ra2/2):N5 = ra2- 2*ra3
    ra4 = int(ra3/2):N6 = ra3- 2*ra4
    ra5 = int(ra4/2):N7 = ra4- 2*ra5
    ra6 = int(ra5/2):N8 = ra5- 2*ra6
    ra7 = int(ra6/2):N9 = ra6- 2*ra7
    ra8 = int(ra7/2):N10 = ra7- 2*ra8
    ra9 = int(ra8/2):N11 = ra8- 2*ra9
    ra10 = int(ra9/2):N12 = ra9- 2*ra10
    ra11 = int(ra10/2):N13 = ra10- 2*ra11
    ra12 = int(ra11/2):N14 = ra11- 2*ra12
    ra13 = int(ra12/2):N15 = ra12- 2*ra13  'MSB
    N16 = 0     'N17,16  Antibacklash width
    N17 = 0     '0=3ns, 1=1.5ns, 2=6ns, 3=3ns
    N18 = 0     'Test Bit, use 0
    N19 = 0     'Test Bit, use 0
    N20 = 0     'Lock Detector Mode, 0=3 refcycles, 1=5 cycles
    N21 = 0     'resyncronization enable 0=normal, 1=resync prescaler
    N22 = 1     '0=resync with nondelayed rf input, 1=resync with delayed rf
    N23 = 0   'reserved, use 0
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
  '[Command4112Rbuffer]
    gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111
    return
  '[endCommand4112R]

[Create4112N]'needed: ncounter,preselector; creates AD4112 N Buffer
    Bcounter = int(ncounter/preselector)
    Acounter = int(ncounter-(Bcounter*preselector))
    if Bcounter < 3 then beep:errora$="4112 N counter <3":return 'with errora$ ver111-37c
    if Bcounter > 8191 then beep:errora$="4112 N counter >8191":return 'with errora$ ver111-37c
    if Bcounter < Acounter then beep:errora$="4112 B counter<Acounter":return 'with errora$ ver111-37c
    'ver116-4o deleted "if" block, per Lrev1
    N0 = 1       'N address bit 0, must be 1
    N1 = 0       'N address bit 1, must be 0
    na0 = int(Acounter/2):N2 = Acounter- 2*na0      'Acounter bit 0 LSB
    na1 = int(na0/2):N3 = na0 - 2*na1
    na2 = int(na1/2):N4 = na1 - 2*na2
    na3 = int(na2/2):N5 = na2 - 2*na3
    na4 = int(na3/2):N6 = na3 - 2*na4
    na5 = int(na4/2):N7 = na4 - 2*na5      'Acounter bit 5 MSB
    nb0 = int(Bcounter/2):N8 = Bcounter- 2*nb0      'Bcounter bit 0 LSB
    nb1 = int(nb0/2):N9 = nb0 - 2*nb1
    nb2 = int(nb1/2):N10 = nb1 - 2*nb2
    nb3 = int(nb2/2):N11 = nb2 - 2*nb3
    nb4 = int(nb3/2):N12 = nb3 - 2*nb4
    nb5 = int(nb4/2):N13 = nb4 - 2*nb5
    nb6 = int(nb5/2):N14 = nb5 - 2*nb6
    nb7 = int(nb6/2):N15 = nb6 - 2*nb7
    nb8 = int(nb7/2):N16 = nb7 - 2*nb8
    nb9 = int(nb8/2):N17 = nb8 - 2*nb9
    nb10 = int(nb9/2):N18 = nb9 - 2*nb10
    nb11 = int(nb10/2):N19 = nb10 - 2*nb11
    nb12 = int(nb11/2):N20 = nb11 - 2*nb12      'Bcounter bit 12 MSB
    N21 = 0    '0=ChargePump setting 1, 1=setting 2
    N22 = 0     'reserved
    N23 = 0     'reserved
    if cb = 3 then Int64N.lsLong.struct = 2^23*N23+ 2^22*N22+ 2^21*N21+ 2^20*N20+ 2^19*N19+ 2^18*N18+ 2^17*N17+ 2^16*N16+ 2^15*N15+_
            2^14*N14+ 2^13*N13+ 2^12*N12+ 2^11*N11+ 2^10*N10+ 2^9*N9+ 2^8*N8+_
            2^7*N7+ 2^6*N6+ 2^5*N5+ 2^4*N4+ 2^3*N3+ 2^2*N2+ 2^1*N1+ 2^0*N0 'ver116-4o per Lrev1
    if cb = 3 then Int64N.msLong.struct = 0 'ver116-4o per Lrev1
    return

[CreateBaseForDDSarray]'needed:ddsoutput,ddsclock ; creates: base,sw0thrusw30,w0thruw4
    'the formula for the frequency output of the DDS(AD9850, 9851, or any 32 bit DDS) is:
    'ddsoutput = base*ddsclock/2^32, where "base" is the decimal equivalent of command words
    'to find "base": first, use: fullbase = (ddsoutput*2^32/ddsclock)
        fullbase=(ddsoutput*2^32/ddsclock) 'decimal number, including fraction
    'then, round it off to the nearest whole bit
            '(the following has a problem) 11-03-08
            'if ddsoutput is greater than ddsclock/2, the program will error out. I don't know why but
                'halt and create an error message
    if ddsoutput >= ddsclock/2 then
        beep:message$="Error, ddsoutput > .5 ddsclock" : call PrintMessage :goto [Halted] 'ver114-4e
    end if
        base = int(fullbase) 'rounded down to whole number
        if fullbase - base >= .5 then base = base + 1 'rounded to nearest whole number
    'now, the actual ddsoutput can be determined by: ddsoutput = base*ddsclock/2^32
  'Create Parallel Words 'needed:base
        w0= 0 'a "1" here will activate the x4 internal multiplier, but not recommended
        w1= int(base/2^24)  'w1 thru w4 converts decimal base code to 4 words, each are 8 bit binary
        w2= int((base-(w1*2^24))/2^16)
        w3= int((base-(w1*2^24)-(w2*2^16))/2^8)
        w4= int(base-(w1*2^24)-(w2*2^16)-(w3*2^8))
    if cb = 3 then 'USB:05/12/2010
        Int64SW.msLong.struct = 0 'USB:05/12/2010
        Int64SW.lsLong.struct = int( base ) 'USB:05/12/2010
    else 'USB:05/12/2010
        'Create Serial Bits'needed:base ; creates serial word bits; sw0 thru sw39
        b0 = int(base/2):sw0 = base - 2*b0  'LSB, Freq-b0.  sw is serial word bit
        b1 = int(b0/2):sw1 = b0 - 2*b1:b2 = int(b1/2):sw2 = b1 - 2*b2
        b3 = int(b2/2):sw3 = b2 - 2*b3:b4 = int(b3/2):sw4 = b3 - 2*b4
        b5 = int(b4/2):sw5 = b4 - 2*b5:b6 = int(b5/2):sw6 = b5 - 2*b6
        b7 = int(b6/2):sw7 = b6 - 2*b7:b8 = int(b7/2):sw8 = b7 - 2*b8
        b9 = int(b8/2):sw9 = b8 - 2*b9:b10 = int(b9/2):sw10 = b9 - 2*b10
        b11 = int(b10/2):sw11 = b10 - 2*b11:b12 = int(b11/2):sw12 = b11 - 2*b12
        b13 = int(b12/2):sw13 = b12 - 2*b13:b14 = int(b13/2):sw14 = b13 - 2*b14
        b15 = int(b14/2):sw15 = b14 - 2*b15:b16 = int(b15/2):sw16 = b15 - 2*b16
        b17 = int(b16/2):sw17 = b16 - 2*b17:b18 = int(b17/2):sw18 = b17 - 2*b18
        b19 = int(b18/2):sw19 = b18 - 2*b19:b20 = int(b19/2):sw20 = b19 - 2*b20
        b21 = int(b20/2):sw21 = b20 - 2*b21:b22 = int(b21/2):sw22 = b21 - 2*b22
        b23 = int(b22/2):sw23 = b22 - 2*b23:b24 = int(b23/2):sw24 = b23 - 2*b24
        b25 = int(b24/2):sw25 = b24 - 2*b25:b26 = int(b25/2):sw26 = b25 - 2*b26
        b27 = int(b26/2):sw27 = b26 - 2*b27:b28 = int(b27/2):sw28 = b27 - 2*b28
        b29 = int(b28/2):sw29 = b28 - 2*b29:b30 = int(b29/2):sw30 = b29 - 2*b30
        b31 = int(b30/2):sw31 = b30 - 2*b31  'MSB, Freq-b31
        sw32 = 0 'x4 multiplier, 1=enable, but not recommended
        sw33 = 0 'control bit
        sw34 = 0 'power down bit
        sw35 = 0 'phase data
        sw36 = 0 'phase data
        sw37 = 0 'phase data
        sw38 = 0 'phase data
        sw39 = 0 'phase data
    end if 'USB:05/12/2010
    return
'[endCreateBaseForDDSarray]

[ResetDDS1par]'needed:control,STRBAUTO,contclear ; resets DDS1 on J5(OrigControlBd), into parallel mode
    out control, STRBAUTO        'wclk and fqud lines high, causing DDS "Reset" line to go high
    out control, contclear     'wclk and fqud lines low (all control lines low)
    return

[ResetDDS1ser]'OrigControlBoard.needed:AUTO,STRB,STRBAUTO ; set DDS1(J5)to serial mode. ver113-2c
    'DDS (AD9850/9851) can be hard wired. pin2=D2=0, pin3=D1=1,pin4=D0=1, D3-D7 are don't care.
    'this will reset DDS into parallel, involk serial mode, then command zero output
    out port, 3 'data=0000 0011, if the DDS is not already hard wired
    '(reset DDS1 to parallel)Data,WCLK up,WCLK up and FQUD up,WCLK up and FQUD down,WCLK down
    out control, AUTO       'WCLK up, FQUD=0
    out control, STRBAUTO   'WCLK=1 and FQUD up
    out control, AUTO       'WCLK=1, FQUD down
    out control, contclear  'WCLK down, FQUD=0
    '(end reset DDS1 to parallel)
    '(involk serial mode DDS1)WCLK up, WCLK down, FQUD up, FQUD down
    out control, AUTO:out control, contclear 'WCLK up, WCLK down
    out control, STRB:out control, contclear 'FQUD up, FQUD down
    'even if the DDS1, D0-D2 is not hard wired, it will be in Serial Mode
    '(end involk serial mode DDS1)
    '(command DDS1 to flush registers)D7=0,WCLK up,WCLK down,(repeat39more),FQUD up,FQUD down
    out port, 0  'D7=0
    for thisloop = 0 to 39
    out control, AUTO:out control, contclear  'D7=0,WCLK up,WCLK down
    next thisloop
    out control, STRB:out control, contclear 'FQUD up, FQUD down
    '(end command DDS1 flush)DDS will output a DC signal
    return

[ResetDDS3ser]'OrigControlBoard.needed:AUTO,STRB,STRBAUTO ; set DDS3(J4)to serial mode. ver113-2c
    'DDS3 (AD9850/9851) must be hard wired. pin2=D2=0, pin3=D1=1,pin4=D0=1, D3-D7 are don't care.
    out control, Jcontrol  'enable Control Board J connector
    '(reset DDS3 to parallel)WCLK up and FQUD up,WCLK down and FQUD down
    out port, 34 'WCLK up and FQUD up.  DDS register pointer will reset
    '(end reset DDS1 to parallel)



    out port, sfqud:out port, 0 ' DDSpin8, FQUD up, FQUD down.  DDS register pointer will reset
    out port, swclk:out port, 0 ' DDSpin9, WCLK up, DDS WCLK down
    out port, sfqud:out port, 0 ' DDSpin8, FQUD up, FQUD down.  DDS will go to 0 Hz.
    out control, contclear  'disable Control Board J connector
    return

[ResetDDS1serSLIM]'reset serial DDS1 without disturbing Filter Bank or PDM. ver113-2c
    'must have DDS (AD9850/9851) hard wired. pin2=D2=0, pin3=D1=1,pin4=D0=1, D3-D7 are don't care.
    'this will reset DDS into parallel, involk serial mode, then command to 0 Hz.
    pdmcmd = phaarray(thisstep,0) 'ver111-39d

    '(reset DDS1 to parallel)WCLK up,WCLK up and FQUD up,WCLK up and FQUD down,WCLK down
    out port, filtbank + 1     'apply last known filter path and WCLK=D0=1 to buffer
    out control, SELT          'DDSpin9, WCLK up to DDS
    out control, contclear     'disable buffer,leaving filtbank, and WCLK=high to DDS
    out port, pdmcmd*64 + 2    'apply last known pdmcmd and FQUD=D3=1 to buffer
    out control, INIT          'DDSpin8, FQUD up,DDS resets to parallel,register pointer will reset
    out port, pdmcmd*64        'DDSpin8, FQUD down
    out control, contclear     'disable buffer, leaving last known PDM state latched
    out port, filtbank         'apply last known filter path and WCLK=D0=0 to buffer
    out control, SELT          'DDSpin9, WCLK down
    out control, contclear     'disable buffer,leaving filtbank
    '(end reset DDS1 to parallel)
    '(involk serial mode DDS1)WCLK up, WCLK down, FQUD up, FQUD down
    out port, filtbank + 1     'apply last known filter path and WCLK=D0=1 to buffer
    out control, SELT          'DDSpin9, WCLK up to DDS
    out port, filtbank         'apply last known filter path and WCLK=D0=0 to DDS
    out control, contclear     'disable buffer,leaving filtbank
    out port, pdmcmd*64 + 2    'apply last known pdmcmd and FQUD=D3=1 to buffer
    out control, INIT          'DDSpin8, FQUD up,DDS resets to parallel,register pointer will reset
    out port, pdmcmd*64        'DDSpin8, FQUD down
    out control, contclear     'disable buffer, leaving last known PDM state latched
    '(end involk serial mode DDS1)
    '(flush and command DDS1)D7,WCLK up,WCLK down,(repeat39more),FQUD up,FQUD down
    'present data to buffer,latch buffer,disable buffer,present data+clk to buffer,latch buffer,disable buffer
    a=filtbank
    for thisloop = 0 to 39
    out port, a:out control, SELT:out control, contclear: out port, a+1:out control, SELT:out control, contclear
    next thisloop
    out port, a:out control, SELT:out control, contclear 'leaving filtbank latched
    out port, pdmcmd*64 + 2    'apply last known pdmcmd and FQUD=D3=1 to buffer
    out control, INIT          'DDSpin8, FQUD up,DDS resets to parallel,register pointer will reset
    out port, pdmcmd*64        'DDSpin8, FQUD down
    out control, contclear     'disable buffer, leaving last known PDM state latched
    '(end flush command DDS1)
    return 'to '[InitializeDDS1]

[ResetDDS1serUSB] 'USB:01-08-2010
    pdmcmd = phaarray(thisstep,0) 'ver111-39d

'(reset DDS3 to parallel)WCLK up,WCLK up and FQUD up,WCLK up and FQUD down,WCLK down
    USBwrbuf$ = "A10100"+ToHex$(filtbank + 1)
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf$ as ptr, 4 as short, result as boolean
    if result <>  TRUE then
        if USBdevice <> 0 then CALLDLL #USB, "UsbMSAInit", USBdevice as long, result as boolean
        if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf$ as ptr, 4 as short, result as boolean
    end if
    USBwrbuf2$ = "A30200"+ToHex$(pdmcmd*64 + 2)+ToHex$(pdmcmd*64)
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf2$ as ptr, 5 as short, result as boolean
    USBwrbuf$ = "A10300"+ToHex$(filtbank)+ToHex$(filtbank + 1)+ToHex$(filtbank)
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf$ as ptr, 6 as short, result as boolean
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf2$ as ptr, 5 as short, result as boolean
'(end involk serial mode DDS3)
'(flush and command DDS3)D7,WCLK up,WCLK down,(repeat39more),FQUD up,FQUD down
'present data to buffer,latch buffer,disable buffer,present data+clk to buffer,latch buffer,disable buffer
    USBwrbuf$ = "A12801"
    USBwrbuf3$ = ToHex$( filtbank )
    for thisloop = 0 to 39
        USBwrbuf$ = USBwrbuf$ + USBwrbuf3$
    next thisloop
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf$ as ptr, 43 as short, result as boolean
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf2$ as ptr, 5 as short, result as boolean
'(end flush command DDS3)

    return

[ResetDDS3serSLIM]'reset serial DDS3 without disturbing Filter Bank or PDM. ver113-2c
    'must have DDS (AD9850/9851) hard wired. pin2=D2=0, pin3=D1=1,pin4=D0=1, D3-D7 are don't care.
    'this will reset DDS into parallel, involk serial mode, then command to 0 Hz.
    pdmcmd = phaarray(thisstep,0) 'ver111-39d

    '(reset DDS3 to parallel)WCLK up,WCLK up and FQUD up,WCLK up and FQUD down,WCLK down
    out port, filtbank + 1     'apply last known filter path and WCLK=D0=1 to buffer
    out control, SELT          'DDSpin9, WCLK up to DDS
    out control, contclear     'disable buffer,leaving filtbank, and WCLK=high to DDS
    out port, pdmcmd*64 + 8    'apply last known pdmcmd and FQUD=D3=1 to buffer
    out control, INIT          'DDSpin8, FQUD up,DDS resets to parallel,register pointer will reset
    out port, pdmcmd*64        'DDSpin8, FQUD down
    out control, contclear     'disable buffer, leaving last known PDM state latched
    out port, filtbank         'apply last known filter path and WCLK=D0=0 to buffer
    out control, SELT          'DDSpin9, WCLK down
    out control, contclear     'disable buffer,leaving filtbank
    '(end reset DDS3 to parallel)
    '(involk serial mode DDS3)WCLK up, WCLK down, FQUD up, FQUD down
    out port, filtbank + 1     'apply last known filter path and WCLK=D0=1 to buffer
    out control, SELT          'DDSpin9, WCLK up to DDS
    out port, filtbank         'apply last known filter path and WCLK=D0=0 to DDS
    out control, contclear     'disable buffer,leaving filtbank
    out port, pdmcmd*64 + 8    'apply last known pdmcmd and FQUD=D3=1 to buffer
    out control, INIT          'DDSpin8, FQUD up,DDS resets to parallel,register pointer will reset
    out port, pdmcmd*64        'DDSpin8, FQUD down
    out control, contclear     'disable buffer, leaving last known PDM state latched
    '(end involk serial mode DDS3)
    '(flush and command DDS3)D7,WCLK up,WCLK down,(repeat39more),FQUD up,FQUD down
    'present data to buffer,latch buffer,disable buffer,present data+clk to buffer,latch buffer,disable buffer
    a=filtbank
    for thisloop = 0 to 39
    out port, a:out control, SELT:out control, contclear: out port, a+1:out control, SELT:out control, contclear
    next thisloop
    out port, a:out control, SELT:out control, contclear 'leaving filtbank latched
    out port, pdmcmd*64 + 8    'apply last known pdmcmd and FQUD=D3=1 to buffer
    out control, INIT          'DDSpin8, FQUD up,DDS resets to parallel,register pointer will reset
    out port, pdmcmd*64        'DDSpin8, FQUD down
    out control, contclear     'disable buffer, leaving last known PDM state latched
    '(end flush command DDS3)
    return 'to '(InitializeDDS 3)

[ResetDDS3serUSB] 'USB:01-08-2010
'reset serial DDS3 without disturbing Filter Bank or PDM. usb v1.0
'must have DDS (AD9850/9851) hard wired. pin2=D2=0, pin3=D1=1,pin4=D0=1, D3-D7 are don't care.
'this will reset DDS into parallel, involk serial mode, then command to 0 Hz.
    pdmcmd = phaarray(thisstep,0) 'ver111-39d

    '(reset DDS3 to parallel)WCLK up,WCLK up and FQUD up,WCLK up and FQUD down,WCLK down
    USBwrbuf$ = "A10100"+ToHex$(filtbank + 1)
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf$ as ptr, 4 as short, result as boolean
    USBwrbuf2$ = "A30200"+ToHex$(pdmcmd*64 + 8)+ToHex$(pdmcmd*64)
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf2$ as ptr, 5 as short, result as boolean
    USBwrbuf$ = "A10300"+ToHex$(filtbank)+ToHex$(filtbank + 1)+ToHex$(filtbank)
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf$ as ptr, 6 as short, result as boolean
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf2$ as ptr, 5 as short, result as boolean
    '(end involk serial mode DDS3)
    '(flush and command DDS3)D7,WCLK up,WCLK down,(repeat39more),FQUD up,FQUD down
    'present data to buffer,latch buffer,disable buffer,present data+clk to buffer,latch buffer,disable buffer
    USBwrbuf$ = "A12801"
    USBwrbuf3$ = ToHex$( filtbank )
    for thisloop = 0 to 39
        USBwrbuf$ = USBwrbuf$ + USBwrbuf3$
    next thisloop
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf$ as ptr, 43 as short, result as boolean
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf2$ as ptr, 5 as short, result as boolean
    '(end flush command DDS3)
    return 'to '(InitializeDDS 3)

[CommandDDS1]'ver111-36b. ver113-4a
    'this will recalculate DDS1, using the values in the Command DDS 1 Box, and "with DDS Clock at" Box.
    'it will insert the new DDS 1 frequency into the command arrays for all steps, leaving others alone
    'it will initiate a re-command at thisstep (where the sweep was halted)
      'if Original Control Board is used, only the DDS 1 is re-commanded. ver113-4a
      'if SLIM Control Board is used, all 4 modules will be re-commanded. ver113-4a
    'using One Step or Continue will retain the new DDS1 frequency.
    'PLO1 will be non-functional until [Restart] button is clicked. PLL1 will break lock and "slam" to extreme.
    '[Restart] will reset arrays and begin sweeping at step 0. Special Tests Window will not be updated.
    'Signal Generator or Tracking Generator output will not be effected.
    'caution, do not enter a frequency that is higher than 1/2 the masterclock frequency (ddsclock)
    print #special.dds1out, "!contents? dds1out$";   'grab contents of Command DDS 1 Box
    ddsoutput = val(dds1out$) 'intended output frequency of DDS 1
    print #special.masclkf, "!contents? msclk$";   'grab contents of "with DDS Clock at" box
    msclk = val(msclk$) 'if "with DDS Clock at" box was not changed, this is the real MasterClock frequency
    ddsclock = msclk
    'caution: if ddsoutput >= to .5 ddsclock, the program will error out
    gosub [CreateBaseForDDSarray]'needed:ddsoutput,ddsclock ; creates: base,sw0thrusw39,w0thruw4
    remember = thisstep 'remember where we were when entering this subroutine
    for thisstep = 0 to steps 'ver112-2a
    gosub [FillDDS1array]'need thisstep,sw0-sw39,w0-w4,base,ddsclock
    next thisstep 'ver112-2a
    thisstep = remember 'ver112-2a
    gosub [CreateCmdAllArray] 'ver112-2a
    if cb = 0 then gosub [CommandDDS1OrigCB]'will command DDS 1, only
'delver113-4a    if cb = 2 then gosub [CommandDDS1SlimCB]'will command DDS 1, only
    if cb = 2 then gosub [CommandAllSlims]'will command all 4 modules. ver113-4a
    if cb = 3 then gosub [CommandAllSlimsUSB]'will command all 4 modules. ver113-4a 'USB:01-08-2010
    wait

[CommandDDS3]'ver111-38a
    'this will recalculate DDS3, using the values in the Command DDS 3 Box, and "with DDS Clock at" Box.
    'it will insert the new DDS 3 frequency into the command arrays for all steps, leaving others alone
    'it will initiate a re-command at thisstep (where the sweep was halted)
      'only the DDS 3 is re-commanded
    'using One Step or Continue will retain the new DDS3 frequency.
    'PLO3 will be non-functional until [Restart] button is clicked. PLL3 will break lock and "slam" to extreme.
    '[Restart] will reset arrays and begin sweeping at step 0. Special Tests Window will not be updated.
    'Signal Generator or Tracking Generator output will be non functional.
    'Spectrum Analyzer function is not effected
    'caution, do not enter a frequency that is higher than 1/2 the masterclock frequency (ddsclock)
    print #special.dds3out, "!contents? dds3out$";   'grab contents of Command DDS 3 Box
    ddsoutput = val(dds3out$) 'intended output frequency of DDS 3
    print #special.masclkf, "!contents? msclk$";   'grab contents of "with DDS Clock at" box
    msclk = val(msclk$) 'if "with DDS Clock at" box was not changed, this is the real MasterClock frequency
    ddsclock = msclk
    'caution: if ddsoutput >= to .5 ddsclock, the program will error out
    gosub [CreateBaseForDDSarray]'needed:ddsoutput,ddsclock ; creates: base,sw0thrusw39,w0thruw4
    remember = thisstep 'remember where we were when entering this subroutine
    for thisstep = 0 to steps
    gosub [FillDDS3array]'need thisstep,sw0-sw39,w0-w4,base,ddsclock
    next thisstep
    thisstep = remember
    gosub [CreateCmdAllArray]
    if cb = 0 then gosub [CommandDDS3OrigCB]'will command DDS 3, only
'delver113-4a    if cb = 2 then gosub [CommandDDS3SlimCB]'will command DDS 3, only
    if cb = 2 then gosub [CommandAllSlims]'will command all 4 modules. ver113-4a
    if cb = 3 then gosub [CommandAllSlimsUSB]'will command all 4 modules. ver113-4a 'USB:01-08-2010
    wait

[DDS3Track]'ver111-39d
    'This uses DDS3 as a Tracking Generator, but is limited to 0 to 32 MHz, when MasterClock is 64 MHz
    'DDS3 spare output is rich in harmonics and aliases.
    'Tracks the values in Working Window, Center Frequency and Sweep Width (already in the command arrays)
    'The Spectrum Analyzer function is not effected.
    'PLO3, Normal Tracking Generator, and Phase portion of VNA will be non-functional
    'Operation:
    'In Working Window, enter Center Frequency to be within 0 to 32 (MHz), or less than 1/2 the MasterClock
    'In Working Window, enter Sweep Width (in MHz). But, do not allow sweep to go below 0 or abov 1/2 MasterClock
    'Click [Restart], then halt.
    'In Special Tests Window, click [DDS 3 Track].  DDS 3 will, immediately, re-command to new frequency.
    'Click [Continue]. Sweep will resume, but with DDS 3 tracking the Spectrum Analalyzer
    '[One Step] and [Continue] and halting operates normally until [Restart] button is pressed.
    '[Restart] will reset arrays, and leave the DDS 3 Track Mode. ie, normal sweeping.
    ddsclock = masterclock
    remember = thisstep
    for thisstep = 0 to steps
    ddsoutput = datatable(thisstep,1)
    'caution: if ddsoutput >= to .5 ddsclock, the program will error out
    gosub [CreateBaseForDDSarray]'needed:ddsoutput,ddsclock ; creates: base,sw0thrusw39,w0thruw4
    gosub [FillDDS3array]'need thisstep,sw0-sw39,w0-w4,base,ddsclock
    next thisstep
    thisstep = remember
    gosub [CreateCmdAllArray]
    if cb = 0 then gosub [CommandDDS3OrigCB]'will command DDS 3, only
    if cb = 2 then gosub [CommandAllSlims]'will command all 4 modules. ver113-4a
    if cb = 3 then gosub [CommandAllSlimsUSB]'will command all 4 modules. 'USB:01-08-2010
    wait

[DDS1Sweep]'ver112-2c
    'This forces the DDS 1 to the values in Working Window: Center Frequency and Sweep Width (already in the command arrays)
    'DDS1 spare output is rich in harmonics and aliases.
    'PLO1, and thus, the Spectrum Analyzer will be non-functional in this mode.
    'Signal Generator or Tracking Generator output will not be affected.
    'Operation:
    'In Working Window, enter Center Frequency to be within 0 to 32 (MHz), or less than 1/2 the MasterClock
    'In Working Window, enter Sweep Width (in MHz). But, do not allow sweep to go below 0 or abov 1/2 MasterClock
    'Click [Restart], then halt.
    'In Special Tests Window, click [DDS 1 Sweep].  DDS 1 will, immediately, re-command to new frequency.
    'Click [Continue]. Sweep will resume, but with DDS 1 sweeping.
    '[One Step] and [Continue] and halting operates normally until [Restart] button is pressed.
    '[Restart] will reset arrays, and will leave the DDS 1 Sweep Mode. ie, normal sweeping.
    ddsclock = masterclock
    remember = thisstep
    for thisstep = 0 to steps
    ddsoutput = datatable(thisstep,1)
    'caution: if ddsoutput >= to .5 ddsclock, the program will error out
    gosub [CreateBaseForDDSarray]'needed:ddsoutput,ddsclock ; creates: base,sw0thrusw39,w0thruw4
    gosub [FillDDS1array]'need thisstep,sw0-sw39,w0-w4,base,ddsclock
    next thisstep
    thisstep = remember
    gosub [CreateCmdAllArray]
    if cb = 0 then gosub [CommandDDS1OrigCB]'will command DDS 1, only
    if cb = 2 then gosub [CommandAllSlims]'will command all 4 modules. ver113-4a
    if cb = 3 then gosub [CommandAllSlimsUSB]'will command all 4 modules.  'USB:01-08-2010 moved ver116-4f
    wait

[CalculateAllStepsForLO1Synth]
    haltstep = thisstep 'remember where we were in the sweep when halted
    for thisstep = 0 to steps
              'SEWgraph deleted the following calc of pixels, since we have already calculated them and don't need them in magarray()
        'magarray(thisstep,0) = thisstep * x/steps  'establishes "thispointx" (on x-axis of graph) (0 to 720)ver111-25
            'SEWgraph Frequencies have been pre-calculated in the graphing module via gGenerateXValues
        thisfreq=gGetPointXVal(thisstep+1)    'Point number is 1 greater than step number SEWgraph
        if msaMode$<>"SA" then  'Store actual signal freq in VNA arrays ver116-1b
            if msaMode$="Reflection" then ReflectArray(thisstep,0)=thisfreq _
                            else S21DataArray(thisstep,0)=thisfreq
        end if
        if FreqMode<>1 then thisfreq=Equiv1GFreq(thisfreq)  'Convert from display freq to equivalent 1G frequency
            'ver116-4k added baseFrequency, which gets added when commanding but does not affect the stored frequencies.
        LO1 = baseFrequency + thisfreq + LO2 - finalfreq    'calculates the actual LO1 frequency:thisfreq,LO2,finalfreq are actuals.
        datatable(thisstep,0) = thisstep    'put current step number into the array, row value= thisstep 'moved ver111-18
        datatable(thisstep,1) = thisfreq    'put current frequency into the array, row value= thisstep 'moved ver111-18

        '[CalculateThisStepPLL1]        
        appxVCO=LO1 : reference=appxdds1 : rcounter=rcounter1
        if PLL1mode = 0 then gosub [CreateIntegerNcounter]'needed:appxVCO,reference,rcounter ; creates:ncount,ncounter,fcounter(0)
            'returns with ncount,ncounter,fcounter(0),pdf
        if PLL1mode = 1 then gosub [CreateFractionalNcounter]'needed:appxVCO,reference,rcounter ; creates:ncount,ncounter,fcounter,pdf
            'returns with ncount,ncounter,fcounter,pdf
        dds1output = pdf * rcounter    'actual output of DDS1(input Ref to PLL1)
        if PLL1mode = 1 then gosub [AutoSpur]'needed:LO2,finalfreq,dds1output,rcounter1,finalbw,appxdds1,fcounter,ncounter ver111-8
            '[AutoSpur] is a continuation of [CreateFractionalNcounter], used only in MSA when PLL 1 is Fractional
            'returns with possibly new: ncounter,fcounter,pdf,dds1output
        if PLL1mode = 1 then gosub [ManSpur]'ver111-10
        '[ManSpur] is a continuation of [CreateFractionalNcounter], used only in MSA when PLL 1 is Fractional
        'if Spur Test Button On, will return with new ncounter,fcounter,pdf,dds1output
        gosub [CreatePLL1N]'needs:ncounter,fcounter,PLL1mode,PLL1 ; creates PLL NBuffer N0-Nx
        gosub [FillPLL1array]'need:N0-Nx,pdf,dds1output,LO1,ncount,ncounter,Fcounter,Acounter,Bcounter;creates samePLL1
        '[endCalculateThisStepPLL1]
        '[CalculateThisStepDDS1]'need:dds1output,masterclock,appxdds1,dds1filbw
          ddsoutput = dds1output : ddsclock = masterclock
        if dds1output-appxdds1>dds1filbw/2 then  'ver114-4e
            beep:error$="DDS1output too high for filter"
            message$=error$ : call PrintMessage  'ver114-4e
            call RequireRestart   'ver115-1c
            wait
        end if
        if appxdds1-dds1output>dds1filbw/2 then  'ver114-4e
            beep:error$="DDS1output too low for filter"
            message$=error$ : call PrintMessage 'ver114-4e
            call RequireRestart   'ver115-1c
            wait
        end if
            gosub [CreateBaseForDDSarray]'needed:ddsoutput,ddsclock ; creates: base,sw0thrusw39,w0thruw4
            gosub [FillDDS1array]'need thisstep,sw0-sw39,w0-w4,base,ddsclock
        '[endCalculateThisStepDDS1]
    next thisstep
    thisstep = haltstep 'return to the step in the sweep, where we halted, if needed
    return
  '[endCalculateAllStepsForLO1Synth]

[CalculateAllStepsForLO3Synth]'for hybrid, and orig (fixed freq) TG
    'if TGtop = 0 then skip all this (return), actually we should not have even entered this subroutine.
    haltstep = thisstep 'remember where we were in the sweep when halted
    for thisstep = 0 to steps  'ver111-17
        'SEWgraph Frequencies have been pre-calculated in the graphing module via gGenerateXValues
    TrueFreq=gGetPointXVal(thisstep+1)+baseFrequency    'Point number is 1 greater than step number SEWgraph ver115-1c ver116-4L
        'ver116-4k baseFrequency gets added when commanding but does not affect the stored frequencies.
    if FreqMode=1 then thisfreq=TrueFreq else thisfreq=Equiv1GFreq(TrueFreq)  'ver115-1c get equivalent 1G frequency ver115-1d
    
    if TGtop = 1 then LO3 = LO2 - finalfreq - offset 'for orig, fixed freq TG  ver111-15a
                  'or LO3 = LO1 - thisfreq - offset
        'ver115-1c rearranged the following if... block and added the FreqMode=3 test
    if TGtop = 2 and gentrk = 1 then
        if normrev = 0 then
            if FreqMode=3 then
                LO3 = TrueFreq + offset - LO2  'ver116-4k  'Mode 3G sets LO3 differently 
            else
                LO3 = LO2 + thisfreq + offset  'ver116-4k 'for new TG, Trk Gen mode, normal
            end if
        end if
        if normrev = 1 then
            'SEWgraph Frequencies have been pre-calculated in the graphing module via gGenerateXValues
            'We can just retrieve them in reverse order.
            TrueFreq=gGetPointXVal(steps-thisstep+1)+baseFrequency    'Point number is 1 greater than step number ver116-4L
            if FreqMode=1 then revfreq=TrueFreq else revfreq=Equiv1GFreq(TrueFreq)  'ver115-1d get equiv 1G freq
            if FreqMode=3 then  'ver115-1d added this if... block
                LO3 = TrueFreq + offset - LO2  'Mode 3G sets LO3 differently
            else
                LO3 = LO2 + revfreq + offset 'for new TG, Trk Gen mode, normal
            end if
        end if
    end if

    if TGtop = 2 and gentrk = 0 then  
        'for new TG, Sig Gen mode ver116-4p
        'We will try to produce sgout, either by LO3-LO2(for 0 to LO2), LO3 (LO2 to 2*LO2) or LO3+LO2 (above 2*LO2)
        select case   'ver116-4p
            case sgout<LO2 : LO3=sgout+LO2
            case sgout<2*LO2 : LO3=sgout
            case else : LO3=sgout-LO2
        end select
    end if
    '[CalculateThisStepPLL3]
        appxVCO=LO3 : reference=appxdds3 : rcounter=rcounter3
        if appxdds3 = 0 then reference=masterclock 'for orig, fixed freq TG with no DDS3 steering. ver111-17
        if PLL3mode = 0 then gosub [CreateIntegerNcounter]'needed:appxVCO,reference,rcounter ; creates:ncount,ncounter,fcounter(0)
            'returns with ncount,ncounter,fcounter(0),pdf
        if PLL3mode = 1 then gosub [CreateFractionalNcounter]'needed:appxVCO,reference,rcounter ; creates:ncount,ncounter,fcounter,pdf
            'returns with ncount,ncounter,fcounter,pdf
            dds3output = pdf * rcounter    'actual output of DDS3(input Ref to PLL3)
            gosub [CreatePLL3N]'needs:ncounter,fcounter,PLL3mode,PLL3 ; creates PLL NBuffer N0-Nx
            gosub [FillPLL3array]'need thisstep,N0thruN23,pdf3(40),dds3output(41),samePLL3(42)see dim PLL3array for slot info 'ver111-14
    '[endCalculateThisStepPLL3]
    '[CalculateThisStepDDS3]'need:dds3output,masterclock,appxdds3,dds3filbw
      if appxdds3 = 0 then goto [endCalculateThisStepDDS3] 'there is no DDS, skip this section ver111-17
        ddsoutput = dds3output : ddsclock = masterclock
        if dds3output-appxdds3>dds3filbw/2 then    'ver114-4e
            beep:error$="DDS3 output too high for filter"
            message$=error$ : call PrintMessage
            call RequireRestart   'ver115-1c
            wait
        end if

        if appxdds3-dds3output>dds3filbw/2 then  'ver114-4e
            beep:error$="DDS3output too low for filter"
            message$=error$ : call PrintMessage
            call RequireRestart   'ver115-1c
            wait
        end if
        gosub [CreateBaseForDDSarray]'needed:ddsoutput,ddsclock ; creates: base,sw0thrusw39,w0thruw4
        gosub [FillDDS3array]'need thisstep,sw0-sw39,w0-w4,base,ddsclock  ver111-15
    [endCalculateThisStepDDS3]
    'del.ver112-2b phaarray(thisstep,0) = 0 'this will set all pdmstates, to 0 'ver112-1a
    phaarray(thisstep,0) = 0 'this will set all pdmstates, to 0 'undeleted, ver113-7e
    next thisstep
    thisstep = haltstep 'return to the step in the sweep, where we halted, if needed
    lastpdmstate = 2 'this will guarantee that the PDM will get commanded 'ver112-1a
    return
'[endCalculateAllStepsForLO3Synth]

[FillPLL1array]'need thisstep,N0thruN23,pdf1(40),dds1output(41),samePLL1(42)see dim PLL1array for slot info 'ver111-1
    if cb = 3 then 'USB:11-08-2010
        if USBdevice <> 0 then CALLDLL #USB, "UsbMSADevicePopulateDDSArrayBitReverse", USBdevice as long, ptrSPLL1Array as ulong, Int64N as ptr, thisstep as short, 40 as short, result as boolean 'USB:11-08-2010
    else 'USB:05/12/2010

        'reversed sequence for N23 to be first. ver111-31a
        PLL1array(thisstep,23) = N0:PLL1array(thisstep,22) = N1
        PLL1array(thisstep,21) = N2:PLL1array(thisstep,20) = N3
        PLL1array(thisstep,19) = N4:PLL1array(thisstep,18) = N5
        PLL1array(thisstep,17) = N6:PLL1array(thisstep,16) = N7
        PLL1array(thisstep,15) = N8:PLL1array(thisstep,14) = N9
        PLL1array(thisstep,13) = N10:PLL1array(thisstep,12) = N11
        PLL1array(thisstep,11) = N12:PLL1array(thisstep,10) = N13
        PLL1array(thisstep,9) = N14:PLL1array(thisstep,8) = N15
        PLL1array(thisstep,7) = N16:PLL1array(thisstep,6) = N17
        PLL1array(thisstep,5) = N18:PLL1array(thisstep,4) = N19
        PLL1array(thisstep,3) = N20:PLL1array(thisstep,2) = N21
        PLL1array(thisstep,1) = N22:PLL1array(thisstep,0) = N23
    end if 'USB:05/12/2010
    PLL1array(thisstep,40) = pdf
    PLL1array(thisstep,43) = LO1
    PLL1array(thisstep,45) = ncounter
    PLL1array(thisstep,46) = fcounter
    PLL1array(thisstep,47) = Acounter
    PLL1array(thisstep,48) = Bcounter
    return

[FillPLL3array]'need thisstep,N0thruN23,pdf3(40),dds3output(41),samePLL3(42)see dim PLL3array for slot info 'ver111-14
    if cb = 3 then'USB:11-08-2010
        if USBdevice <> 0 then CALLDLL #USB, "UsbMSADevicePopulateDDSArrayBitReverse", USBdevice as long, ptrSPLL3Array as ulong, Int64N as ptr, thisstep as short, 40 as short, result as boolean 'USB:11-08-2010
    else 'USB:05/12/2010
        'reversed sequence for N23 to be first. ver111-31a
        PLL3array(thisstep,23) = N0:PLL3array(thisstep,22) = N1
        PLL3array(thisstep,21) = N2:PLL3array(thisstep,20) = N3
        PLL3array(thisstep,19) = N4:PLL3array(thisstep,18) = N5
        PLL3array(thisstep,17) = N6:PLL3array(thisstep,16) = N7
        PLL3array(thisstep,15) = N8:PLL3array(thisstep,14) = N9
        PLL3array(thisstep,13) = N10:PLL3array(thisstep,12) = N11
        PLL3array(thisstep,11) = N12:PLL3array(thisstep,10) = N13
        PLL3array(thisstep,9) = N14:PLL3array(thisstep,8) = N15
        PLL3array(thisstep,7) = N16:PLL3array(thisstep,6) = N17
        PLL3array(thisstep,5) = N18:PLL3array(thisstep,4) = N19
        PLL3array(thisstep,3) = N20:PLL3array(thisstep,2) = N21
        PLL3array(thisstep,1) = N22:PLL3array(thisstep,0) = N23
    end if 'USB:05/12/2010
    PLL3array(thisstep,40) = pdf
    PLL3array(thisstep,43) = LO3
    PLL3array(thisstep,45) = ncounter
    PLL3array(thisstep,46) = fcounter
    PLL3array(thisstep,47) = Acounter
    PLL3array(thisstep,48) = Bcounter
    return

[FillDDS1array]'need thisstep,sw0-sw39,w0-w4,base,ddsclock 'ver111-12
    if cb = 3 then'USB:11-08-2010
        if USBdevice <> 0 then CALLDLL #USB, "UsbMSADevicePopulateDDSArray", USBdevice as long, ptrSDDS1Array as ulong, Int64SW as ptr, thisstep as short, result as boolean 'USB:11-08-2010
    else 'USB:05/12/2010
        DDS1array(thisstep,0) = sw0:DDS1array(thisstep,1) = sw1
        DDS1array(thisstep,2) = sw2:DDS1array(thisstep,3) = sw3
        DDS1array(thisstep,4) = sw4:DDS1array(thisstep,5) = sw5
        DDS1array(thisstep,6) = sw6:DDS1array(thisstep,7) = sw7
        DDS1array(thisstep,8) = sw8:DDS1array(thisstep,9) = sw9
        DDS1array(thisstep,10) = sw10:DDS1array(thisstep,11) = sw11
        DDS1array(thisstep,12) = sw12:DDS1array(thisstep,13) = sw13
        DDS1array(thisstep,14) = sw14:DDS1array(thisstep,15) = sw15
        DDS1array(thisstep,16) = sw16:DDS1array(thisstep,17) = sw17
        DDS1array(thisstep,18) = sw18:DDS1array(thisstep,19) = sw19
        DDS1array(thisstep,20) = sw20:DDS1array(thisstep,21) = sw21
        DDS1array(thisstep,22) = sw22:DDS1array(thisstep,23) = sw23
        DDS1array(thisstep,24) = sw24:DDS1array(thisstep,25) = sw25
        DDS1array(thisstep,26) = sw26:DDS1array(thisstep,27) = sw27
        DDS1array(thisstep,28) = sw28:DDS1array(thisstep,29) = sw29
        DDS1array(thisstep,30) = sw30:DDS1array(thisstep,31) = sw31
        DDS1array(thisstep,32) = sw32:DDS1array(thisstep,33) = sw33
        DDS1array(thisstep,34) = sw34:DDS1array(thisstep,35) = sw35
        DDS1array(thisstep,36) = sw36:DDS1array(thisstep,37) = sw37
        DDS1array(thisstep,38) = sw38:DDS1array(thisstep,39) = sw39
    end if 'USB:05/12/2010
    DDS1array(thisstep,40) = w0
    DDS1array(thisstep,41) = w1
    DDS1array(thisstep,42) = w2
    DDS1array(thisstep,43) = w3
    DDS1array(thisstep,44) = w4
    DDS1array(thisstep,45) = base 'base is decimal command
    DDS1array(thisstep,46) = base*ddsclock/2^32 'actual dds 1 output freq
    return

[FillDDS3array]'need thisstep,sw0-sw39,w0-w4,base,ddsclock 'ver111-15
    if cb = 3 then'USB:11-08-2010
        if USBdevice <> 0 then CALLDLL #USB, "UsbMSADevicePopulateDDSArray", USBdevice as long, ptrSDDS3Array as ulong, Int64SW as ptr, thisstep as short, result as boolean 'USB:11-08-2010
    else 'USB:05/12/2010
        DDS3array(thisstep,0) = sw0:DDS3array(thisstep,1) = sw1
        DDS3array(thisstep,2) = sw2:DDS3array(thisstep,3) = sw3
        DDS3array(thisstep,4) = sw4:DDS3array(thisstep,5) = sw5
        DDS3array(thisstep,6) = sw6:DDS3array(thisstep,7) = sw7
        DDS3array(thisstep,8) = sw8:DDS3array(thisstep,9) = sw9
        DDS3array(thisstep,10) = sw10:DDS3array(thisstep,11) = sw11
        DDS3array(thisstep,12) = sw12:DDS3array(thisstep,13) = sw13
        DDS3array(thisstep,14) = sw14:DDS3array(thisstep,15) = sw15
        DDS3array(thisstep,16) = sw16:DDS3array(thisstep,17) = sw17
        DDS3array(thisstep,18) = sw18:DDS3array(thisstep,19) = sw19
        DDS3array(thisstep,20) = sw20:DDS3array(thisstep,21) = sw21
        DDS3array(thisstep,22) = sw22:DDS3array(thisstep,23) = sw23
        DDS3array(thisstep,24) = sw24:DDS3array(thisstep,25) = sw25
        DDS3array(thisstep,26) = sw26:DDS3array(thisstep,27) = sw27
        DDS3array(thisstep,28) = sw28:DDS3array(thisstep,29) = sw29
        DDS3array(thisstep,30) = sw30:DDS3array(thisstep,31) = sw31
        DDS3array(thisstep,32) = sw32 'x4 multiplier
        DDS3array(thisstep,33) = sw33 'control bit
        DDS3array(thisstep,34) = sw34 'power down bit
        DDS3array(thisstep,35) = sw35 '35-39 are Phase
        DDS3array(thisstep,36) = sw36:DDS3array(thisstep,37) = sw37
        DDS3array(thisstep,38) = sw38:DDS3array(thisstep,39) = sw39
    end if 'USB:05/12/2010
    DDS3array(thisstep,40) = w0 'word 0, 8 bits, mult, control and phase
    DDS3array(thisstep,41) = w1 'word 1, 8 bits
    DDS3array(thisstep,42) = w2 'word 2, 8 bits
    DDS3array(thisstep,43) = w3 'word 3, 8 bits
    DDS3array(thisstep,44) = w4 'word 4, 8 bits
    DDS3array(thisstep,45) = base 'base is decimal command
    DDS3array(thisstep,46) = base*ddsclock/2^32 'actual dds 3 output freq
    return

[CreateCmdAllArray] 'for SLIM CB only 'ver-31b
    'a DDS serial command, will begin with LSB (W0), thru MSB (W31), ending with Phase bit 4 (W39)
    'a PLL serial command, will begin with MSB (N23), thru LSB (N0, the address bit)
    rememberthisstep = thisstep 'remember where we were when entering this subroutine
    if cb <> 3 then 'USB:05/12/2010
        for thisstep = 0 to steps
            for clmn = 0 to 15
                cmdallarray(thisstep,clmn) = DDS1array(thisstep,clmn)*4 + DDS3array(thisstep,clmn)*16
            next clmn
            for clmn = 16 to 39
                cmdallarray(thisstep,clmn) = PLL1array(thisstep,clmn-16)*2 + DDS1array(thisstep,clmn)*4 + PLL3array(thisstep,clmn-16)*8 + DDS3array(thisstep,clmn)*16
            next clmn
        next thisstep
    else 'USB:05/12/2010
        if USBdevice <> 0 then CALLDLL #USB, "UsbMSADevicePopulateAllArray", USBdevice as long, steps as short, 40 as short, _
                            0 as long, ptrSPLL1Array as long, ptrSDDS1Array as long, ptrSPLL3Array as long, _
                            ptrSDDS3Array as long, 0 as long, 0 as long, 0 as long, _
                            result as boolean 'USB:11-08-2010
    end if 'USB:05/12/2010
    thisstep = rememberthisstep 'ver116-4k
    return

[CommandPLL]' comes here during PLL R Initializations and PLL 2 N command ver111-28
    if cb = 0 then gosub [CommandPLLorig] 'ver111-28
    if cb = 2 then gosub [CommandPLLslim] 'ver111-28
    if cb = 3 then gosub [CommandPLLslimUSB]  'USB:01-08-2010
    return 'to [InitializePLL2]or[CommandXPllRbuffer]

[CommandPLLorig]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111-28
    'used during initialization of PLL1, PLL2, and PLL3.  PDM will get set to "0".
    'when PLL1 or PLL2 then Jcontrol=SELT. when PLL3 the Jcontrol=INIT
    out control, Jcontrol   'enable Control Board J connector
    out port, N23:out port, N23 + 2
    out port, N22:out port, N22 + 2:out port, N21:out port, N21 + 2
    out port, N20:out port, N20 + 2:out port, N19:out port, N19 + 2
    out port, N18:out port, N18 + 2:out port, N17:out port, N17 + 2
    out port, N16:out port, N16 + 2:out port, N15:out port, N15 + 2
    out port, N14:out port, N14 + 2:out port, N13:out port, N13 + 2
    out port, N12:out port, N12 + 2:out port, N11:out port, N11 + 2
    out port, N10:out port, N10 + 2:out port, N9:out port, N9 + 2
    out port, N8:out port, N8 + 2:out port, N7:out port, N7 + 2
    out port, N6:out port, N6 + 2:out port, N5:out port, N5 + 2
    out port, N4:out port, N4 + 2:out port, N3:out port, N3 + 2
    out port, N2:out port, N2 + 2:out port, N1:out port, N1 + 2
    out port, N0:out port, N0 + 2:out port, LEPLL:out port, 0     'Latch buffer
    out control, contclear       'Disable the Control Board J connector
    return 'to [CommandPLL]

[CommandPLLslimUSB] 'USB:01-08-2010
    if USBdevice = 0 then return 'USB:05/12/2010
    CALLDLL #USB, "UsbMSADeviceWriteInt64MsbFirst", USBdevice as long, 161 as short, Int64N as ptr, 24 as short, 1 as short, filtbank as short, datavalue as short, result as boolean  'USB:11-08-2010
    pdmcommand = phaarray(thisstep,0)*64 'do not disturb PDM state, this may be used during Spur Test
    USBwrbuf$ = "A30200"+ToHex$(pdmcommand + levalue)+ToHex$(pdmcommand)
    CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf$ as ptr, 5 as short, result as boolean 'USB:05/12/2010
    return

[CommandPLLslim]'needs:datavalue,levalue,N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,SLIM ControlBoard ver111-28
    'used during initialization of PLL1, PLL2, and PLL3.  PDM will get set to "0" during Initializations
    'selt word = 1 common clock, 4 datas, plus 3 (filtbank). entering this sub, selt word should = filtbank only
    'init word = 5 latch lines plus 2 pdm commands. entering this sub, init word should = pdmcmd + pdmclk only.ver111-39d
    'two steps to do: command data and clock without disturbing Filter Bank, then send LE without disturbing PDM
  'step 1. Command the PLL without changing the filter bank.
    'For PLL1,datavalue=2, for PLL2,datavalue=16, for PLL3,datavalue=8
    'following code lines changed in ver113-3c
    a=filtbank + N23*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N22*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N21*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N20*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N19*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N18*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N17*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N16*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N15*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N14*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N13*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N12*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N11*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N10*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N9*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N8*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N7*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N6*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N5*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N4*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N3*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N2*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N1*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    a=filtbank + N0*datavalue:out port, a:out control, SELT:out control, contclear:out port, a+1:out control, SELT:out control, contclear
    out port, filtbank:out control, SELT:out control, contclear 'leaving lines latched to filter bank
    out port, 0
  'step 2. Command the PLL without changing the PDM
    pdmcommand = phaarray(thisstep,0)*64 'do not disturb PDM state, this may be used during Spur Test
    out port, pdmcommand + levalue  'levalues: PLL1=1, PLL2=16, PLL3=4
    out control, INIT
    out port, pdmcommand
    out control, contclear  'leaving lines latched, and unchanged, to PDM
    out port, 0
    return 'to [CommandPLL]

[DetermineModule] 'ver111-28
    'All "glitchXX's" are "0" when entering this subroutine. Either from "fresh RUN" or [WaitStatement]
    'if a module is not present, or if it doesn't need commanding, return with it's "glitchXX = 0"
  '[DDS1]
    dds1output = DDS1array(thisstep,46) 'ver111-16
    if dds1output = lastdds1output then goto [PLL1] 'dds 1 is same, don't waste time commanding 'ver111-28
    glitchd1 = 1 'ver111-36h
    lastdds1output = dds1output
  [PLL1]
    ncounter1=PLL1array(thisstep,45):fcounter1=PLL1array(thisstep,46) 'ver111-16
    if ncounter1=lastncounter1 and fcounter1=lastfcounter1 then goto [PLL3] 'don't waste time commanding 'ver111-28
    glitchp1 = 1 'add 1 msec delay. ver111-28
    lastncounter1=ncounter1:lastfcounter1=fcounter1 'ver111-16
  [PLL3]
    if  TGtop = 0 then return 'there is no PLL 3, no DDS 3,and no PDM for VNA
    ncounter3=PLL3array(thisstep,45):fcounter3=PLL3array(thisstep,46)
    if ncounter3=lastncounter3 and fcounter3=lastfcounter3 then goto [DDS3] 'don't waste time commanding 'ver111-28
    glitchp3 = 1 'add 1 msec delay. ver111-28
    lastncounter3=ncounter3:lastfcounter3=fcounter3
  [DDS3]
    if appxdds3 = 0 then goto [PDM] 'if 0, there is no DDS3, but, there can be VNA ver111-28
    dds3output = DDS3array(thisstep,46)
    if dds3output = lastdds3output then goto [PDM] 'dds 3 is same, don't waste time commanding 'ver111-29
    glitchd3 = 1 'ver111-36h
    lastdds3output = dds3output
  [PDM]
    if suppressPhase or msaMode$="SA" or msaMode$="ScalarTrans" then return ' not in VNA mode, skip the PDM 'ver116-1b
    pdmcmd = phaarray(thisstep,0) 'ver111-39d
    if pdmcmd = lastpdmstate then return 'don't waste time commanding
    gosub [VideoGlitchPDM]  'ver114-6c
    'lastpdmstate = pdmcmd  'delver114-6c we won't update "lastpdmstate" until it is actually commanded
    return 'to [CommandThisStep]

[CommandOrigCB]' correct modules have been determined in [DetermineModule]
    'Command necessary modules, independently, from Original Control Board
    if glitchd1 > 0 then gosub [CommandDDS1OrigCB]
    if glitchp1 > 0 then gosub [CommandPLL1OrigCB]
    if glitchd3 > 0 and TGtop > 0 then gosub [CommandDDS3OrigCB]
    if glitchp3 > 0 and TGtop > 0 then gosub [CommandPLL3OrigCB]
    if glitchpdm > 0  and msaMode$<>"SA" and msaMode$<>"ScalarTrans" then gosub [CommandPDMOrigCB] 'ver114-5n
    return 'to [CommandThisStep]

[CommandDDS1OrigCB] 'needed:DDS1array 'ver111-21
    if dds1parser = 1 then goto [CommandDDS1OrigCBserial] 'ver111-21
   '(CommandDDS1OrigCBparallel)'needed:DDS1array(w0-w4),port,control,AUTO,STRB,contclear ; commands DDS1 on J5, parallel. ver111-21
        'note, a DDS commanded parallel, will begin with Control Word (W0), then MSB Word (W1), ending with LSB Word (W4)
          'set word 0        'set 8 bit word, W0 (0), phase info
        out port,DDS1array(thisstep,40) ' a "1" here would activate the x4 internal multiplier, but not recommended
        out control, AUTO       'wclk line goes high
        out control, contclear      'wclk line goes low
          'set word 1
        out port,DDS1array(thisstep,41) 'set 8 bit word, W1, MSB freq
        out control,AUTO:out control, contclear
          'set word 2
        out port,DDS1array(thisstep,42) 'W2
        out control,AUTO:out control, contclear
          'set word 3
        out port,DDS1array(thisstep,43) 'W3
        out control,AUTO:out control, contclear
         'set word 4
        out port,DDS1array(thisstep,44) 'set 8 bit word, W4, LSB freq
        out control,AUTO:out control, contclear
        out port, 0            'return the output port data lines to 0
         'send fqud
        out control, STRB          'set fqud to 1, freq changes now
        out control, contclear         'set fqud to 0 and all others to 0
    return 'to [CommandOrigCB]

    [CommandDDS1OrigCBserial]'needed:DDS1array(sw0-sw39),control,AUTO,STRB,contclear ; commands DDS1 on J5, serially ver111-21
        'note: once the DDS1 has been reset into serial mode, the D0 thru D6 data lines are "don't care".
        'note, a DDS serial command, will begin with LSB (W0), thru MSB (W31), ending with Phase bit 4 (W39)
        for clmn = 0 to 39 'ver111-21
        out port, DDS1array(thisstep,clmn)*128 'apply data bit to DDS1pin25, D7 data line
        out control, AUTO:out control, contclear  'retain data bit while wclk up, then down
        next clmn 'next bit in 40 bit serial data transfer
        out port, 0
        out control, STRB:out control, contclear 'fqud up, fqud down
    return 'to [CommandOrigCB]
'[endCommandDDS1OldRevA]

[CommandPLL1OrigCB]'needed:PLL1array(N23-N0),SELT,lastncounter1,lastfcounter1 'ver111-21
    'ver111-28a makes the SELT buffer "see" the pdm state before commanding PLL1, to prevent orig PDM from changing states.
    Jcontrol = SELT : LEPLL = 4 'ver111-21
    'Command PLL1,oldControl using N23-N0,control,Jcontrol,port,contclear,LEPLL ver111-21
    'note, a PLL will serially command beginning with N23 and end with N0 (address bit)
    pdmcmd = phaarray(thisstep,0) 'ver111-39d
    out port, pdmcmd*128 'ver111-28a
    out control, Jcontrol   'enable Control Board J connector
    for clmn = 0 to 23  'reversed order 'ver111-31a
    out port, pdmcmd*128 + PLL1array(thisstep,clmn):out port, pdmcmd*128 + PLL1array(thisstep,clmn) + 2 'ver111-21 'ver111-28a
    next clmn 'ver111-21
    out port, pdmcmd*128 + LEPLL:out port, pdmcmd*128     'Latch buffer 'ver111-28a
    out control, contclear       'Disable the Control Board J connector
    out port, 0 'ver111-28a
    lastpdmstate=phaarray(thisstep,0)   'ver114-6c
    return 'to [CommandOrigCB]

[CommandDDS3OrigCB]'needed:DDS3array,lastdds3output,INIT 'ver111-18
    Jcontrol = INIT:swclk = 32:sfqud = 2 'for Orig Control Bd,J4,DDS3 ver111-16
    'Command DDS3,serially,oldControl using sw0-sw39,swclk,sfqud,control,Jcontrol,port,contclear,LEPLL ver111-21
    'note, a DDS commanded serially, will begin with LSB, continue to MSB, and end with Control Word MSB Phase Bit
    'present filter bank data while commanding DDS3, so as not to change filter bank ver111-29
    out port, filtbank 'ver111-29
    out control, Jcontrol  'enable Control Board J connector
    for clmn = 0 to 39 'ver111-21
    out port, filtbank + DDS3array(thisstep,clmn) 'apply data bit to DDS, and also filter lines 'ver111-29
    out port, filtbank + DDS3array(thisstep,clmn) + swclk  'apply data bit and wclk 'ver111-29
    next clmn
    out port, filtbank:out port, filtbank + sfqud:out port, filtbank 'last sw down and swclk down, sfqud up, sfqud down 'ver111-29
    out control, contclear  'disable J connector
    out port, 0 'ver111-29
    return 'to [CommandOrigCB]

[CommandPLL3OrigCB]'needed:PLL3array(N23-N0),INIT,lastncounter3,lastfcounter3 ver111-18
    Jcontrol = INIT : LEPLL = 16 'ver111-21
    'Command PLL3,Orig Control using N23-N0,control,Jcontrol,port,contclear,LEPLL ver111-21
    'note, a PLL will serially command beginning with N23 and end with N0 (address bit)
    'present filter bank data while commanding PLL3, so as not to change filter bank ver111-29
    out port, filtbank 'ver111-29
    out control, Jcontrol   'enable Control Board J connector
    for clmn = 0 to 23  'reversed order 'ver111-31a
    out port, filtbank + PLL3array(thisstep,clmn) 'ver111-29
    out port, filtbank + PLL3array(thisstep,clmn) + 2 'ver111-21 'ver111-29
    next clmn 'ver111-21
    out port, filtbank + LEPLL:out port, filtbank     'Latch buffer 'ver111-29
    out control, contclear       'Disable the Control Board J connector
    out port, 0 'ver111-29
    return 'to [CommandOrigCB]

[CommandPDMonly] 'ver111-28
    if cb = 0 then goto [CommandPDMOrigCB] 'ver111-28
    if cb = 2 then goto [CommandPDMSlimCB] 'ver111-28
    if cb = 3 then goto [CommandPDMSlimUSB]  'USB:01-08-2010
    return 'to [InvertPDmodule]

[CommandPDMOrigCB]'Set original PDM phase for last known mode, since a PLL1 or PLL2 command will reset the PDM to Norm.
    out port, phaarray(thisstep,0)*128: out control, SELT: out control, contclear: out port, 0  'pdmcmd is determined in [InvertPDmodule] 'ver111-20
    lastpdmstate=phaarray(thisstep,0)   'ver114-6c
    return 'to [CommandOrigCB]or[CommandPDMonly]

[CommandPDMSlimCB]'also sending a "latch signal", used by orig PDM module
    out port, phaarray(thisstep,0)*64
    out control, INIT
    out port, phaarray(thisstep,0)*64 + 32
    out port, phaarray(thisstep,0)*64
    out control, contclear
    out port, 0
    lastpdmstate=phaarray(thisstep,0)   'ver114-6c
    return 'to [CommandPDMonly]

[CommandPDMSlimUSB] 'USB:01-08-2010
    i = phaarray(thisstep,0)*64
    USBwrbuf$ = "A30300"+ToHex$(i)+ToHex$(i+32)+ToHex$(i)
    if USBdevice <> 0 then CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf$ as ptr, 6 as short, result as boolean
    lastpdmstate=phaarray(thisstep,0)   'ver114-6c
    return

[CommandAllSlims]'for SLIM Control and SLIM modules. Old PDM and old Filt Bank can be used 'ver111-31c
   '(send data and clocks without changing Filter Bank)
    '0-15 is DDS1bit*4 + DDS3bit*16, data = 0 to PLL 1 and PLL 3. see[CreateCmdAllArray].
    'present new Data with no clock,latch high,latch low,present new data with clock,latch high,latch low. ver113-2a
    'repeat for each bit. (40 data bits and 40 clocks for each module, even if they don't need that many)
    'this format guarantees that the common clock will not transition with a data transition, preventing crosstalk in LPT cable. ver111-32c
    for clmn = 0 to 39  'ver113-3c
        a= cmdallarray(thisstep,clmn)+ filtbank
        out port, a : out control, SELT:out control, contclear 'a is the data, without clock
        out port, a+1:out control, SELT:out control, contclear 'a+1 is data, plus clock
    next clmn
    out port, filtbank 'remove data, leaving filtbank data to filter bank.
    out control, SELT:out control, contclear 'disable buffer. filtbank signals will be latched to filter bank assembly

  'send LE's to PLL1, PLL3, FQUD's to DDS1, DDS3, and command PDM
    'begin by setting up init word=LE's and Fquds + PDM state for thisstep
    pdmcmd = phaarray(thisstep,0)*64 'ver111-39d
    out port, le1 + fqud1 + le3 + fqud3 + pdmcmd 'present data to buffer input'ver111-39d
    out control, INIT: out control, contclear  'latch the buffer, moving the signals to the 5 modules'ver113-2a
    out port, pdmcmd + 32 'remove LEs and Fquds, leaving PDM data, but add a latch signal P2D5 for old PDM if used.'ver111-39d
    out control, INIT: out control, contclear  'sends latch signal to old PDM'ver113-2a
    out port, pdmcmd  'remove the added latch signal to PDM, leaving just the PDM's static data'ver111-39d
    out control, INIT: out control, contclear  'ver113-2a
    out port, 0    'bring all Data lines low. PDM data remains static
    lastpdmstate=phaarray(thisstep,0)   'ver114-6c
    return 'to [CommandThisStep]

' ****************
'USB:05/12/2010
' following code changed from previous USB code
' USB: 15/08/10
' all three of the following work. SLowest at the top, fastest at the bottom

'[CommandAllSlimsUSB] 'USB:01-08-2010
'    '(send data and clocks without changing Filter Bank)
'    if USBdevice = 0 then return
'    CALLDLL #USB, "UsbMSADeviceAllSlims", USBdevice as long, thisstep as short, filtbank as short, result as boolean 'USB:11-08-2010
'    pdmcmd = phaarray(thisstep,0)*64 'ver111-39d
'    USBwrbuf$ = "A30300"+ToHex$(le1 + fqud1 + le3 + fqud3 + pdmcmd)+ToHex$(pdmcmd + 32)+ToHex$(pdmcmd)
'    CALLDLL #USB, "UsbMSADeviceWriteString", USBdevice as long, USBwrbuf$ as ptr, 6 as short, result as boolean
'    lastpdmstate=phaarray(thisstep,0)   'ver114-6c
'    return

'[CommandAllSlimsUSB] 'USB:01-08-2010
'    '(send data and clocks without changing Filter Bank)
'    if USBdevice = 0 then return
'    pdmcmd = phaarray(thisstep,0)*64
'    i = le1 + fqud1 + le3 + fqud3
'    CALLDLL #USB, "UsbMSADeviceAllSlimsAndLoad", USBdevice as long, thisstep as short, filtbank as short, i as short, pdmcmd as short, 32 as short, result as boolean 'USB:14-08-2010
'    lastpdmstate=phaarray(thisstep,0)   'ver114-6c
'    return

[CommandAllSlimsUSB] ' USB: 15/08/10
    '(send data and clocks without changing Filter Bank)
    if USBdevice = 0 then return 'USB:05/12/2010
    if thisstep = 0 then 'USB:05/12/2010
        UsbAllSlimsAndLoadData.filtbank.struct = filtbank 'USB:05/12/2010
        UsbAllSlimsAndLoadData.latches.struct = le1 + fqud1 + le3 + fqud3 'USB:05/12/2010
        UsbAllSlimsAndLoadData.pdmcmdmult.struct = 64 'USB:05/12/2010
        UsbAllSlimsAndLoadData.pdmcmdadd.struct = 32 'USB:05/12/2010
    end if 'USB:05/12/2010
    UsbAllSlimsAndLoadData.pdmcommand.struct = phaarray(thisstep,0) 'USB:05/12/2010
    UsbAllSlimsAndLoadData.thisstep.struct = thisstep 'USB:05/12/2010
    CALLDLL #USB, "UsbMSADeviceAllSlimsAndLoadStruct", USBdevice as long, UsbAllSlimsAndLoadData as struct, result as boolean ' USB: 15/08/10
    lastpdmstate=phaarray(thisstep,0)   'ver114-6c
    return



'cav :Cavity filter test routines added by ver116-4c
'   Scans around zero in 0.1 MHz increments--e.g. span=10, steps=100. span=10 to 50, steps=span/0.1
'   User sets up scan, restarts, halts, opens Special Tests Window, clicks cavity filter test button
'   The software will command the PLO2 to maintain an offset from PLO 1 by exactly the amount of the
'   Final I.F., that is, PLO2 will always be equal to PLO1+FIF
'   PLL2 Rcounter buffer is commanded one time, to assure pdf will be 100 KHz. Done during Init after "Restart"
'   PLO2 N counter buffer is commanded at each step in the sweep
'   The actual frequency that is passed through the Cavity Filter is the displayed frequency plus 1024 MHz.
'   The Cavity Filter sweep limitations are:
'   the lowest frequency possible is where PLO 1 cannot legally command (Bcounter=31, appx 964 MHz)
'       (PLO1 or PLO2 bottoming out at 0V is also limit, but likely below 964 MHz)
'   the highest frequency possible s where PLO 2 tops out (vco volts near 5v, somewhere between 1050 to 1073 MHz)
'   Sweep can be halted at any time and Sweep Parameters can be changed, then Continue or Restart
'   The Special Tests window must be closed before MSA will return to normal. Must click "Restart".

[CavityFilterTest]'comes here when Cavity Filter Test button is clicked
    if cftest = 1 then wait
    cftest = 1
    enterPLL2phasefreq = PLL2phasefreq
    PLL2phasefreq = .1
goto [Restart]

[CloseCavityFilterTest]'will come here when Special Tests Window is closed
    cftest = 0
    PLL2phasefreq = enterPLL2phasefreq
return

[CommandLO2forCavTest]
    appxVCO = finalfreq + PLL1array(thisstep,43)
    reference = masterclock
    rcounter = rcounter2
    gosub [CreateIntegerNcounter]'needs:appxVCO,reference,rcounter ; creates:ncounter,fcounter(0)
    ncounter2 = ncounter:fcounter2 = fcounter
    gosub [CreatePLL2N]'needs:ncounter,fcounter,PLL2 ; returns with Bcounter,Acounter, and N Bits N0-N23
    Bcounter2=Bcounter: Acounter2=Acounter
    LO2=((Bcounter*preselector)+Acounter+(fcounter/16))*pdf2 'actual LO2 frequency  'ver115-1c LO2 is now global
    'CommandPLL2N
    Jcontrol = SELT : LEPLL = 8
    datavalue = 16: levalue = 16 'PLL2 data and le bit values ver111-28
    gosub [CommandPLL]'needs:N23-N0,control,Jcontrol,port,contclear,LEPLL ; commands N23-N0,old ControlBoard ver111-5
return      

[OpenDataWindow]'ver113-5a
if haltsweep = 1 then gosub [FinishSweeping] 'ver114-6f
    'if the "Array Data Window" is already open, close it.
        if datawindow = 1 then close #datawin:datawindow = 0
    'create window called, Data Window, to display all data for each step

    WindowWidth = 425   'ver115-4h
    WindowHeight = 300
    UpperLeftX = DisplayWidth-WindowWidth-20    'ver114-6f
    UpperLeftY = 20    'ver114-6f
    BackgroundColor$ = "white"
    ForegroundColor$ = "black"
    open "Data Window" for text as #datawin
    datawindow = 1
    #datawin, "!font Courier_New 9"  'ver115-2d
    return

sub CloseDataWindow hndl$    'ver115-1b changed to sub. Note this is never used anyway
    close #datawin:datawindow = 0
end sub

'
'ver114-5p added Interpolation Module; ver114-5q moved it to preced config module
'====================START INTERPOLATION ROUTINES==========================
'The interpolation module handles linear and cubic interpolation, and deals with the fact
'that LB does not allow arrays as arguments. Three arrays are created: intSrc() contains the
'original data, which we assume for the moment is freq, mag and phase, in ascending order of freq.
'intDest() contains only frequency, and needs its mag and phase determined by interpolation. If
'cubic interpolation is used, the user calls intCreateCubicCoeffTable which fills the third array,
'intSrcCoeff(), with the eight coefficients (4 for mag, 4 for phase) that are needed to apply cubic
'interpolation.
'The user fills intSrc and intDest, calls intCreateCubicCoeffTable if necessary, and then calls
'intSrcToDest, which performs the interpolation. The user then copies the data (rounding as desired)
'from intDest() to the array where the data really belongs. The data in these arrays should not be
'relied on long-term, because another routine may make use of them.
'Angles must be in the range -180 to 180 degrees to interpolate properly. Returned angles may be
'outside that range and must be normalized as desired.

   'Variables for interpolation routines
    dim intSrc(1000, 2), intDest(1000,2) 'Data for InterpolateTableToTable (freq, real, imag); first index runs from 1
    dim intSrcCoeff(1000,7)   'Cubic coefficents (A,B,C,D) for interpolating real and imag parts from intSrc()
    global intSrcPoints, intDestPoints, intMaxPoints

'------------------Data access routines--------------------
'Even though our data is global, the user should access data only through these routines
sub intSetMaxNumPoints maxPoints
    if maxPoints+5<=intMaxPoints then exit sub  'ver115-2d  'so we never shrink or waste time
    intMaxPoints=maxPoints+5 'ver115-2d
    redim intSrc(intMaxPoints,2) : redim intDest(intMaxPoints,2)
    redim intSrcCoeff(intMaxPoints,7)
end sub

sub intReset    'Reset arrays and variables
    for i=0 to intMaxPoints 'we don't use zero for first index, but clear it anyway
        intSrc(i,0)=0:intSrc(i,1)=0:intSrc(i,2)=0
        intDest(i,0)=0:intDest(i,1)=0:intDest(i,2)=0
    next i
    intSrcPoints=0
    intDestPoints=0
end sub

sub intClearSrc     'Set source table to zero entries
    intSrcPoints=0
end sub

sub intClearDest     'Set destination table to zero entries
    intDestPoints=0
end sub

sub intAddSrcEntry f, r, im     'Add entry to end of source table
    'f=frequency, r=real part, im=imaginary part
    intSrcPoints=intSrcPoints+1
    intSrc(intSrcPoints,0)=f
    intSrc(intSrcPoints,1)=r
    intSrc(intSrcPoints,2)=im
end sub

sub intAddDestFreq f        'Add new frequency to end of destination table
    intDestPoints=intDestPoints+1
    intDest(intDestPoints,0)=f
end sub

sub intGetSrc num, byRef f, byRef r, byRef im     'Get values for source entry number num (1...)
    'f=frequency, r=real part, im=imaginary part
    f=intSrc(num,0)
    r=intSrc(num,1)
    im=intSrc(num,2)
end sub

function intSrcFreq(num)     'Get frequency for source entry number num (1...)
    intSrcFreq=intSrc(num,0)
end function

sub intGetDest num, byRef f, byRef r, byRef im     'Get values for dest entry number num (1...)
    'f=frequency, r=real part, im=imaginary part
    f=intDest(num,0)
    r=intDest(num,1)
    im=intDest(num,2)
end sub

function intDestFreq(num)     'Get frequency for dest entry number num (1...)
    intDestFreq=intDest(num,0)
end function

function intMaxEntries()  'Get maximum number of entries
    intMaxEntries=intMaxPoints
end function

function intSrcEntries()  'Get number of steps in source table
    intSrcEntries=intSrcPoints
end function

function intDestEntries()  'Get number of steps in destination table
    intDestEntries=intDestPoints
end function

'------------------End Data access routines--------------------

'------------------Start Interpolation routines--------------------

function intLinearInterpolateDegrees(fract, v1, v2, angleMin, angleMax)   'linearly interpolate between phase v1 and v2 based on fract
    'fract is a proportion (0...1) representing how far from v1 to v2 we want to go. 
    'angleMin and angleMax are the min and max allowable values for angles; if the total range is at
    'least 360 degrees, angleMin is actually not allowed.
    'The complication with phase is that it wraps around. If the distance from v1 to v2 is shorter via wrap-around,
    'we will assume that wrap-around occurred.
    'This works when the real phase shift from point to point is normally small.
    dif=v2-v1 : if dif<0 then absDif=0-dif else absDif=dif
    if absDif>180 then 'change must exceed 180 degrees before we consider wrap-around
        'If allowed range is 360 or less, we have wrap, which changes the angle difference by 360
        range=angleMax-angleMin
        if range<=360 then 
            if dif>0 then dif=dif-360 else dif=dif+360
        else    'Unusual case with allowed range over 360 degrees
            range=360*int((angleMax-angleMin)/360)  'make range a multiple of 360
            if absDif>range/2 then  'Assume wrap if change exceeds half the range
                if dif>0 then dif=dif-range else dif=dif+range  'change difference by range, which makes dif magnitude smaller
            end if
        end if      
    end if
    res=v1+fract*dif
        'At this point, interp may be outside the bounds of angleMin and angleMax
    if res>angleMin and res<=angleMax then intLinearInterpolateDegrees=res : exit function
    if res=angleMin then    'disallow angleMin value if range is at least 360 degrees
        if angleMax-angleMin>=360 then res=res+360
        intLinearInterpolateDegrees=res : exit function
    end if
    'Here res is below angleMin or above angleMax, so it should wrap around
    if range<=360 then delta=360 else delta=range
    if res<angleMin then wrappedRes=res+delta else wrappedRes=res-delta   'do wrap
    if wrappedRes>=angleMin and wrappedRes<=angleMax then intLinearInterpolateDegrees=wrappedRes _
                    else intLinearInterpolateDegrees=res    'Use wrapped value if wrap put it in allowed range (which it will if range>=360)
end function

sub intLinearInterpolation freq, isPolar, f1,R1,I1,f2,R2,I2,byref p1, byref p2
' linearly interpolate between points 1 and 2 (each a frequency with complex value)
' freq=frequency for which value is to be determined
' This function will return the interpolated value for the real and imaginary parts
' in p1 and p2, respectively
'Interpolated angles may end up outside the original bounds of the angles, if wrap-around occurred,
'so the user should put them back in bounds. They will be at most 360 degress out-out-bounds.

    fSpan=f2-f1
            'If we are interpolating between identical entries, just use data from the first
    if fSpan=0 then p1=R1:p2=I1:exit sub
    ratio=(freq-f1)/fSpan  'Interpolation ratio
    p1=R1+ratio*(R2-R1)
    dif=I2-I1
    if isPolar=1 then
        'We are interpolating polar data, so what we are labeling as the imaginary part
        'is really the angle in degrees.
        if dif<0 then absDif=0-dif else absDif=dif
        if absDif<=180 then
            p2=I1+ratio*dif
        else
            'Apparent shift is more than 180 degrees, which means wrap occurred between v1 and v2.
            'We shift v1 360 degrees towards v2 which reduces the magnitude of dif by 360.
            if dif>0 then dif=dif-360 else dif=dif+360
            p2=I1+ratio*dif
            'At this point, p2 may be outside the bounds of the scaling of v1 and v2. The user must
            'adjust the result
        end if
    else
        p2=I1+ratio*dif
    end if
end sub

sub intCreateCubicCoeffTable doPart1, doPart2, isAngle, favorFlat, doingPhaseCorrection    'Create table of cubic coefficients intSrc() in intCubicCoeff ver116-1b
    'We separately calculate for the specified parts of srcInt. If isAngle=1 then part2 is an angle.
    'We pass favorFlat on to intCalcCubicCoeff, except if part2 is an angle we pass on favorFlat=1 for that part.
    'Each entry of the cubic coefficient table will have 4 numbers. 0-3 are the A,B,C,D
    'coefficients for interpolating the frequency power correction, which is a scalar.
    'Angles must be in the range -180<=angle<=180
    'If doingPhaseCorrection is 1, we are interpolating the phase correction factor,
    'and want to force the correction to 180 if this point phase correction is >179 or <-179. Extreme
    'correction values are used to indicate that phase at that level is unreliable, so the table of
    'phase corrections may have values of 180 for all ADC readings at and below a certain level. ver116-1b
    for i=1 to intSrcPoints
        if doPart1=0 then
            A=0 : B=0 : C=0 : D=0
        else
            partNum=1 : partIsAngle=0
            call intCalcCubicCoeff i,partNum,partIsAngle,favorFlat,A,B,C,D
        end if
        intSrcCoeff(i,0)=A :intSrcCoeff(i,1)=B
        intSrcCoeff(i,2)=C :intSrcCoeff(i,3)=D
        forcedTo180=0
        if doingPhaseCorrection=1 then  'ver116-1b
            'If this point is 180, force calculation to 180
            checkPhase=intSrc(i,2)
            if checkPhase>179 or checkPhase<-179 then
                'This point is 180; set coefficients so phase correction will calculate to 180
                A=180 : B=0 : C=0 : D=0
                forcedTo180=1   'Flag that we need no more calculation at this point
            end if
            if forcedTo180=0 then
                'this point is not 180 but one of prior two points is, then treat the 180 value
                'as being the same as this point. Note we may alter intSrc, but that is just a temporary
                'array used only for these interpolations.
                if i>1 then
                    checkPhase=intSrc(i-1,2)
                    if checkPhase>179 or checkPhase<-179 then intSrc(i-1,2)=intSrc(i,2)
                end if
                if i>2 then
                    checkPhase=intSrc(i-2,2)
                    if checkPhase>179 or checkPhase<-179 then intSrc(i-2,2)=intSrc(i,2)
                end if
            end if
        end if

        if doPart2=0 then
            A=0 : B=0 : C=0 : D=0
        else
            if forcedTo180=0 then    'ver116-1b
                partNum=2 : partIsAngle=isAngle
                'For an angle, specify to "favor flat", because
                'we expect the phase not to approach vertical
                if isAngle then doFlat=1 else doFlat=favorFlat
                call intCalcCubicCoeff i,partNum,partIsAngle,doFlat,A,B,C,D
            end if
        end if
        intSrcCoeff(i,4)=A :intSrcCoeff(i,5)=B
        intSrcCoeff(i,6)=C :intSrcCoeff(i,7)=D
    next i
end sub

sub intCalcCubicCoeff pointNum, partNum, isAngle, favorFlat,byref A,byref B,byref C,byref D
    'Calculate the cubic interpolation coefficients to apply to a point
    'lying between (possibly including) points pointNum-1 and pointNum of intSrc().
    'partNum=1 to process real part and =2 to process imag part.
    'If isAngle=1 then we are interpolating an angle, otherwise not.

'The coefficients will approximate y values in the interval from pointNum-1 to pointNum
'as a cubic equation, such that it passes through the endpoints with the desired slope.
'To determine the desired slope, we use the points to the left and right of the interval.
'This gives us four points, 0-3, of which pointNum is number 2. The interval in which data
'will be interpolated with these coefficients is from point 1 to point 2. We determine what
' slopes we want at the endpoints and then fit a cubic equation to the interval. In
' general, at each point we want the curve on each side to have a common slope equal
' to some sort of average of the interval slopes on each side. We do a straight
' arithmetic average, except that if favorFlat=1 then we average the inverses of the
' slopes and invert that average. The latter tends to make the averaged slope flatter,
' which is useful to avoid overshoot/undershoot.
' Assume the cubic function
'       y=A + B(x-x2) + C(x-x2)^2 + D(x-x2)^3
' passes through points (x1, y1) and (x2, y2), with f1<f2, and with slopes of
' m1 and m2 respectively. Then
'   Let K=x1-x2
'   A=y2
'   B=m2
'   C=(m1-m2)/2 - (3/2)*D*K
'   D=[K(m1+m2)+2(y2-y1)] / K^3

    '(0)If data is less than point 1, we are going to use the point 1 value
    if pointNum=1 then
        A=intSrc(1,partNum)
        B=0 : C=0 :D=0
        exit sub
    end if

    '(1) Get data for four points, 0-3, to be used in interpolation. pointNum is
    'point 2 of these four, so the data to be interpolated will lie between points
    '1 and 2.If we are near one end, point 0 or point 3 will be missing.
    x1=intSrc(pointNum-1,0) : y1=intSrc(pointNum-1,partNum)
    x2=intSrc(pointNum,0) : y2=intSrc(pointNum,partNum)
    if pointNum>2 then x0=intSrc(pointNum-2,0) : y0=intSrc(pointNum-2,partNum)
    if pointNum<intSrcPoints then x3=intSrc(pointNum+1,0) : y3=intSrc(pointNum+1,partNum)
    dif01=y1-y0     'y change from point 0 to point 1 (not used if there is no point 1)
    dif12=y2-y1     'y change from point 1 to point 2
    dif23=y3-y2     'y change from point 2 to point 3 (not used if there is no point 3)

'(2) if points 1 and 2 are the same, then use Y1 value
    if x1=x2 then A=y1 : B=0 : C=0 :D=0 : exit sub

'(3) Deal with angles wrapping at +/-180
    if isAngle=1 then    'deal with angles wrapping at +/-180
        'For any pair of adjacent points, the absolute phase difference exceeds 180 degrees,
        'we assume wrap-around occurred so that the real difference is no more than 180 degrees.
        'If this occurs we reduce the magnitude of the difference by 360 degrees.
        if dif12<0 then absDif12=0-dif12 else absDif12=dif12
        if pointNum<intSrcPoints then   'true if we have point 3
            if dif23<0 then absDif23=0-dif23 else absDif23=dif23
            if absDif23>180 then   'large difference between points 2 and 3
                if dif23>0 then dif23=dif23-360 else dif23=dif23+360  'reduce magnitude of difference by 360
            end if
        end if
        if absDif12>180 then   'large difference between points 1 and 2
            if dif12>0 then dif12=dif12-360 else dif12=dif12+360  'reduce magnitude of difference by 360
        end if
        if pointNum>2 then  'true if we have point 0
            if dif01<0 then absDif01=0-dif01 else absDif01=dif01
            if absDif01>180 then   'large difference between points 0 and 1
                if dif01>0 then dif01=dif01-360 else dif01=dif01+360  'reduce magnitude of difference by 360
            end if
        end if
    end if
'(4) Find m1, the desired slope of the cubic at point 1
    if pointNum>2 then
        if x0=x1 then
            inSlope=0      'should not happen
        else
            inSlope=dif01/(x1-x0)    'slope from point 0 to point 1
        end if
    else
        inSlope=dif12/(x2-x1)    'slope from point 1 to point 2
    end if
    outSlope=dif12/(x2-x1)  'slope from point 1 to point 2
    prod=inSlope*outSlope
    if prod <=0 then
            'if slope on either side is zero, or is positive on one side
            'and negative on the other, we want a zero slope
            m1= 0
    else
        'Calculate an average slope for point 1 based on connecting lines
        'We average the inverses of the slopes, then invert
        m1= 2*prod/(inSlope+outSlope)
    end if

'(5) Find m2, the desired slope of the cubic at point 2
    inSlope=outSlope    'slope from point 1 to point 2
    if pointNum<intSrcPoints then
        if x2=x3 then
            outSlope=0      'should not happen
        else
            outSlope=dif23/(x3-x2)   'slope from point 2 to point 3
        end if
    end if
    'if pointNum is the final point, outSlope will remain the slope from point 1 to point 2
    prod=inSlope*outSlope
    if prod <=0 then
            'if slope on either side is zero, or is positive on one side
            'and negative on the other, we want a zero slope
            m2= 0
    else
        'Calculate an average slope for point 1 based on connecting lines
        if favorFlat then
            m2= 2*prod/(inSlope+outSlope) 'average the inverses of the slopes, then invert
        else
            m2=(inSlope+outSlope)/2
        end if
    end if

'(6) Calc constants for cubic. We are returning A,B,C and D.
    K=x1-x2
    A=y2
    B=m2
    D=(K*(m1+m2)+2*(dif12))/K^3
    C=(m1-m2)/(2*K)-1.5*D*K
end sub

sub intCubicInterpolation targData, ceil, wantV2, byref v1, byref v2
    'This function returns the interpolated values v1 and v2 (but v2 only if wantV2=1),
    'based on the value targData and its position in intSrc(). ceil specifies the ceiling
    'entry for targData in intSrc()--i.e. the first entry whose x value is  >= targData.
    'if ceil=-1, we will look up that position with binary search. Arrays must be in
    'ascending order of x values (usually frequency).
    'We use cubic interpolation using the cubic coefficients which must have been
    'precalculated in intSrcCoeff()

    if ceil=-1 then ceil=intBinarySearch(targData)   'search intSrc() to get ceil
    'ceil now is the first entry >= magdata, except that if no entry meets that test,
    'ceil will be one past the end.
    v1=0 :v2=0
    if ceil>intSrcPoints then
        'Off top end;use values for final intSrc() entry
        v1=intSrc(intSrcPoints,1)
        v2=intSrc(intSrcPoints,2)
        exit sub
    end if
    if ceil=1 then
        'Off bottom end;use mag and phase correction for smallest ADC entry
        v1=intSrc(1,1)
        v2=intSrc(1,2)
        exit sub
    end if

    'Evaluate cubic at x=targData
    dif=targData-intSrc(ceil,0)
    A=intSrcCoeff(ceil,0) : B=intSrcCoeff(ceil,1)
    C=intSrcCoeff(ceil,2) : D=intSrcCoeff(ceil,3)
    v1 = A+dif*(B+dif*(C+dif*D))

    if wantV2=1 then
        A=intSrcCoeff(ceil,4) : B=intSrcCoeff(ceil,5)
        C=intSrcCoeff(ceil,6) : D=intSrcCoeff(ceil,7)
        v2 = A+dif*(B+dif*(C+dif*D))
        'TO DO--caller must put phase in proper range
    end if
end sub

function intBinarySearch(searchVal)
'Perform search of intSrc() to find the lowest entry whose lookup value is >=searchVal
'If dataType=0 then the lookup is for ADC value in calMagTable; otherwise it is for freq in in calFreqTable
'If searchVal is beyond the highest entry, we will return intSrcPoints+1
    top=intSrcPoints
    bot=1
    span=top-bot+1
    while span>4  'Do preliminary search to narrow the search area
        halfSpan=int(span/2)
        mid=bot+halfSpan
        thisVal=intSrc(mid,0)
        if thisVal=searchVal then intBinarySearch=mid : exit function   'exact hit
        if thisVal<searchVal then bot=mid+1 else top=mid  'Narrow to whichever half thisVal is in
        span=top-bot+1  'Repeat with either bot or top endpoints changed
    wend
    'Here thisVal is >= entry bot-1 but <= entry top, and we have a span of less than 4
    'Start with bot entry and find first entry >= searchVal
    ceil=bot
    thisVal=intSrc(ceil,0)
    if thisVal>=searchVal then intBinarySearch=ceil : exit function
    ceil=ceil+1 : if ceil>intSrcPoints then intBinarySearch=ceil : exit function
    thisVal=intSrc(ceil,0)
    if thisVal>=searchVal then intBinarySearch=ceil : exit function
    ceil=ceil+1 : if ceil>intSrcPoints then intBinarySearch=ceil : exit function
    thisVal=intSrc(ceil,0)
    if thisVal>=searchVal then intBinarySearch=ceil : exit function
    intBinarySearch=ceil+1     'searchVal is off the top of the table
end function

sub intSrcToDest isPolar, interpMode, params
'Creates a table of frequency vs. complex numbers by interpolating from another table.
'isPolar=1 if data is in polar form; otherwise 0
'interpMode=0 for linear interpolation; interpMode=1 for cubic interpolation
'params=1 means process real part only; params=2 means imag only; params=3 means both parts
'This is used to create the tables used for calibration. We may have measured data for the
'calibration standards, which may not be at the specific frequencies needed. We may also have
'already calculated the OSL coefficients, but not at the current frequency points. This is
'used to convert such data into data at the exact frequencies of the current scan.
'Since LB does not allow arrays as arguments, this has to be done with two fixed arrays.
'The original array is intSrc(point, data); the final is intDest(point, data),
'where point is the index for the entry number (0 is not used) and data is the index for
'the table data, which consists of frequency (0), real (1) and imaginary (2). (The actual
'labels "point" and "data" are not used.)
'intDest must be pre-filled with the desired frequencies, and intDestSteps must contain
'the number of frequency entries in intDest.
'intSrcSteps must contain the number of entries in intSrc (it's actually the number of points, not steps).
'Everything is written in terms of the data being real and imaginary, but it could instead
'be magnitude and angle. Different results are obtained using the different formats; the
'difference is small if the changes in the data from step to step are small, as they normally
'should be. The complication of interpolating in polar format is that the angle wraps around
'at +/- 180 degrees. If interpolating half way between -178 and +178, we don't want to get 0.
'Therefore, if isPolar=1, if the angle difference between two successive points exceeds 180, we
'assume a wrap-around occurred.
'For cubic interpolation, the cubic coefficient table must be up to date.
'Any angles we actually interpolate will be put into in the range -180<angle<=180
    if intSrcPoints<2 then notice "Interpolation error: insufficient points"
    currDestPoint=0 'Keeps track of last intDest point successfully processed
    srcF1=intSrc(1,0) 'First frequency in intSrc()
    'Use the cal data for the first intSrc() point so long as our target
    'frequency is less than or equal to the intSrc() frequency.
    for i=1 to intDestPoints
        destF=intDest(i,0) 'current step frequency
        if destF>srcF1 then exit for
        currDestPoint=i
        intDest(i,1)=intSrc(1,1)    'Directly transfer value from first entry
        intDest(i,2)=intSrc(1,2)
    next i
    if currDestPoint=intDestPoints then exit sub 'We have processed all points in intDest() ver115-4i
    currDestPoint=currDestPoint+1 'The next intDest() point to process
    'As long as we are within the bounds of intSrc,
    'interpolate to get the cal data for each step frequency.
    ceil=1 'This will become the point number of the first intSrc() freq
                    'greater than or equal to the then current intDest() frequency.
    for i=currDestPoint to intDestPoints
        destF=intDest(i,0)  'current point frequency
        'Move through intSrc() until we hit a frequency >= this intDest() frequency.
        while destF>intSrc(ceil,0)
            ceil=ceil+1   'move to next point in intSrc()
            if ceil>intSrcPoints then exit while   'ran off end of source table
        wend
        if ceil>intSrcPoints then exit for 'ran off end
        'Here ceil is the first baseLine step with a frequency>= our target step freq.
        'So we interpolate between points ceil-1 and ceil.
        srcF1=intSrc(ceil-1,0)   'low entry in src table : freq, mag and phase
        srcMag1=intSrc(ceil-1,1)
        srcPhase1=intSrc(ceil-1,2)  'Note this may actually be imaginary part, not phase
        srcF2=intSrc(ceil,0)   'high entry in src table : freq, mag and phase
        srcMag2=intSrc(ceil,1)
        srcPhase2=intSrc(ceil,2)
        if srcF2=destF then 'If exact frequency match, just copy source data ver115-2d
            v1=srcMag2 : v2=srcPhase2
        else
                'Do the interpolation
            if interpMode=0 then
                call intLinearInterpolation destF, isPolar, srcF1, srcMag1, srcPhase1, srcF2, srcMag2, srcPhase2,v1, v2
            else
                wantV2=1
                call intCubicInterpolation destF, ceil, wantV2, v1, v2 'find v1, v2 by cubic interpolation
            end if
        end if

        currDestPoint=i 'Record that we have this step taken care of
        intDest(i,1)=v1  'Enter results in intDest; x value is already there
        if isPolar then
            if v2>180 then v2=v2-360 else if v2<=-180 then v2=v2+360   'put into range -180<v2<=180
        end if
        intDest(i,2)=v2
    next i
    if currDestPoint=intDestPoints then exit sub 'We have processed all dest points
    currDestPoint=currDestPoint+1 'First unprocessed step
    'For any remaining points, use the values for the final entry
    'in intSrc
    for i=currDestPoint to intDestPoints
        intDest(i,1)=intSrc(intSrcPoints,1) 'Enter results in intDest; x value is already there
        intDest(i,2)=intSrc(intSrcPoints,2)
    next i
    'TO DO--caller must round results if desired
end sub

'====================END INTERPOLATION MODULE==========================
'


'SEW Added the following Calibration Module to load magnitude and frequency calibration from files
'
'===================================================================
'@==============Calibration Manager Module==========================
'@ Version 1.03
'This module displays a window for viewing and editing files for calibration-over-signal-level
'("mag cal") and calibration-over-frequency ("freq cal"). The former records actual power levels
'corresponding to given ADC readings. For VNA use, it also includes errors in phase
'measurement over frequency.
'
'To perform calibration call calManRunManager. While it is running, calManWindHndl$ will be non-blank.
'To check whether a freq and mag path 1 file exist, use "calFileExists()".
'Before calling anything other than calManRunManager, call InitFirstUse maxMagPoint, maxFreqPoints
'to initialize and to set the max number of arrays. Call calCloseWindows if the calling program quits when
'there may still be a cal manager window open.
'
'To apply the calibration tables to a measurement, use calConvertMagData or calConvertFreqError.

    global calEditorPathNum  '0 for freq; 1-N for magnitude
    dim calManFileList$(40)    'List of active paths; zero entry is used
    global calManFiles$  'Indicates which files exist (marked by "1")
                            'First char is freq cal file, Nth (1...41) is mag cal file N-1
    dim calFileInfo$(10,3)  'Used internally to request file info

    global calManEntryIsRef  'Used to keep track of whether data entry is on first (reference) point
    global calManRefPhase, calManRefFreq, calManRefPower     'Used during entry of data
    global calManOldText$   'original text or text at time of last Save.
    global calManLastAutoPoint  'Previous point number entered for freq cal from the user sweep 'SEW8
    global calManEnterError   'Set to 1 if error occurs in calManEnter; otherwise 0 ver114-3d

'This is a gosub subroutine rather than a true subroutine so that it can call [Measure] which
'in turn can call the non-subroutines that currently run the MSA hardware.
[calManRunManager]    'Open window and let user proceed.
    WindowWidth = 575 : WindowHeight =620
    UpperLeftX = 200 : UpperLeftY = 50
    BackgroundColor$ = "lightgray" : ForegroundColor$ = "black"
    TextboxColor$="white" : ComboboxColor$="white"
                'Text Editor
    texteditor #calman.te 45, 45,220,220
statictext #calman.teLabel,"Frequency Calibration Table", 68, 30, 150,15 'ver113-7e
'delver113-7e    statictext #calman.xLabel,"Frequency", 68, 30, 50,15
'delver113-7e    statictext #calman.yLabel,"  Power", 135, 30, 50,15
'delver113-7e    statictext #calman.pLabel,"Phase", 195, 30, 50,15
                'File buttons
    calTEbot=260 : calButtonLeft=300 :calTop=30
    button #calman.Reload, "Re-Load File", calManBtnFile, UL, calButtonLeft, calTop+15, 100, 25
    button #calman.Save, "Save File", calManBtnFile, UL, calButtonLeft, calTop+45, 100, 25
    button #calman.Return, "Return to MSA", [calManBtnReturn], UL, calButtonLeft, calTop+75, 100, 25
        'The following checkbox can be implemented for testing
    'checkbox #calman.phase, "Include Phase", [calManSetPhase], [calManResetPhase], calButtonLeft+120, calTop+15, 100, 25
                'File List
    statictext #calman.fileLabel,"Available Files", calButtonLeft+25, calTop+115, 80,13
    calManFileList$(0)="0 (Frequency)"
    listbox #calman.pathList calManFileList$(), [calManSelectPath], calButtonLeft, calTop+130, 120, 100 'ver116-1b
                'Data Entry items
    calTextLeft=10 : calTextTop=calTEbot+60
    button #calman.Clean, "Clean Up", calManClean, UL, 50, calTEbot+10, 90, 25
        'SEW8 changed .Default to .DispDefault so Return key doesn't activate it
    button #calman.DispDefault, "Display Defaults", calManDisplayDefault, UL, 150, calTEbot+10, 90, 25

    TextboxColor$="cyan"
    textbox #calman.data1, calTextLeft,calTextTop+15, 75,20 'create label and box, will also be "Input (dBm)" ver113-7d
    statictext #calman.Ldata1, "Point Num",calTextLeft+5,calTextTop+3, 65,13  'ver114-3b
    textbox #calman.data2, calTextLeft+90,calTextTop+15, 75,20 'create label and box, will also be "ADC value" ver113-7d
    statictext #calman.Ldata2, "Freq(MHz)",calTextLeft+95,calTextTop+3, 75,13
    textbox #calman.data3, calTextLeft+180,calTextTop+15, 75,20 'create label and box, will also be "Phase (degrees)" ver113-7d
    stylebits #handle.Ldata3, _BS_MULTILINE, 0, 0, 0   'Prints label over two lines 'ver114-3b
    statictext #calman.Ldata3, "Measured Power (dBm)",calTextLeft+185,calTextTop-11, 75,26 'ver116-1b
    textbox #calman.ref, calTextLeft+430,calTextTop+15, 75,20 'create label and box, will also be "Ref Freq (MHz)" ver113-7d
    statictext #calman.Lref, "Ref ADC",calTextLeft+430,calTextTop+3, 100,13 'will also be "True Power(dbm)"
    textbox #calman.ref2, calTextLeft+430,calTextTop-20, 75,20 'create label and box ver113-7d
    statictext #calman.Lref2, "Ref Phase(deg)",calTextLeft+430,calTextTop-32, 90,13

    checkbox #calman.Force, "Force Phase To", [calManForceOn], [calManForceOff], calTextLeft+265, calTextTop-15, 95, 20   'ver116-4j
    textbox #calman.ForceVal,  calTextLeft+365,calTextTop-15, 45,20   'ver116-4j
    button #calman.Measure, "Measure", [calManMenuMeasure], UL, calTextLeft+265,calTextTop+10, 75,25
             'SEW8 Added NextPoint and PrevPoint to replace Measure for freq cal.
    button #calman.NextPoint, "Next Point", [calManBtnNextFreqPoint], UL, calTextLeft+265,calTextTop-5, 75,20
    button #calman.PrevPoint, "Prev Point", [calManBtnPrevFreqPoint], UL, calTextLeft+265,calTextTop+15, 75,20
    button #calman.EnterAll, "Enter All", [calManEnterAll], UL, calTextLeft+345,calTextTop-5, 75,20   'SEW8 For freq cal
    button #calman.Enter, "Enter", calManEnter, UL, calTextLeft+345,calTextTop+15, 75,20
    button #calman.StartEntry, "Start Data Entry", calManEnterInstructions, UL, calTextLeft+150,calTextTop+10, 150,25

    graphicbox #calman.enterInstruct, calTextLeft,calTextTop+40, 550,210    'ver116-1b

    'call calInitFirstUse 100,800  'delver114-4b Done elsewhere (with more points)
                'Open Window  SEW8 changed to use window rather than modal dialog
    open  "Calibration File Manager ver. "+using("##.##", calVersion()) for window_nf as #calman

    calManWindHndl$="#calman"    'So MSA can close it if it is left open
    #calman, "trapclose [calManFinished]"
    #calman.te, "!cls"
    #calman.te, "!font courier_new 9"
    #calman.ref, "!disable"  'So user doesn't change it
    #calman.enterInstruct, "font Times_New_Roman 11"
    #calman.pathList, "singleclickselect"
    #calman.Force, "hide" : #calman.ForceVal, "!hide"  'ver116-4j
        'Set to handle phase based on MSA global hasVNA. When we are not in the MSA,
        'this is undefined and we implement checkbox control, so set the checkbox
    call calSetDoPhase hasVNA
    'if hasVNA=0 then #calman.phase "set" else #calman.phase "reset"

    call calManEnterAvailablePaths  'Loads list of filter files
    #calman.pathList, "selectindex 1"    'Select frequency file
    calEditorPathNum=0    'SEW6 ver113-7c
    dum=calManFile("Load")    'Loads frequency file (creates if necessary)
    wait    'wait for user action in cal window

[calManForceOn] 'ver116-4j
    #calman.ForceVal, "!show" : #calman.ForceVal, "0" : wait
[calManForceOff] 'ver116-4j
    #calman.ForceVal, "!hide" : wait

[calManSetPhase]        'Only relevant when checkbox is implemented
    call calSetDoPhase 1
    if calEditorPathNum<>0 then call calManClean "" : call calManPrepareEntry
    wait

[calManResetPhase]  'Only relevant when checkbox is implemented
    call calSetDoPhase 0
    if calEditorPathNum<>0 then call calManClean ""  : call calManPrepareEntry
    wait

[calManBtnReturn]
    doCan=calManWarnToSave(1) 'Warn to save if data changed. Allow cancel.
    if doCan=1 then return  'cancelled  Exit calManRunManager
    call calInstallFile 0   'Leave with freq file installed
    'call calInstallFile 1   'Leave with path 1 installed  del ver116-1b
    call calCloseWindows
    haltsweep=0     'SEW6
    call RequireRestart 'ver116-4j
    return      'Returns to caller of calManRunManager

[calManFinished]
    goto [calManBtnReturn]  'ver116-1b

sub calCloseWindows     'Close any windows that are open
    if calManWindHndl$<>"" then close #calman : calManWindHndl$=""
end sub

[calManEnterAll]   'button to Enter all sweep points for freq cal
        for i=0 to steps  'ver113-7e
        call calManGetFreqInput i   'Put data into boxes
        if i=steps then call calManEnter "xx" else call calManEnter "" 'Enter data into table; clean up on last one
        if calManEnterError then exit for 'ver114-3d Stop if error
    next i
    wait

sub calManEnter btn$    'Handle Enter button
    'If called by button click, btn$ will not be blank and we do clean up at the end.
    'If btn$ is blank, that is a signal to skip the clean up because we are entering
    'a mass of points.
    calManEnterError=0
    #calman.data1, "!contents? s1$" : #calman.data2, "!contents? s2$" : #calman.data3, "!contents? s3$"
    doForce$="reset"    'so it is reset for freq cal ver116-4j
    if calEditorPathNum=0 then
        'If doing freq cal, we need to update the entry for
        'freq and measured power, because user may have changed the scan. Note that this
        'means the user cannot manually enter data into the freq and power boxes, which are
        'are now disabled.
        'ver114-3d changed to update the boxes
        #calman.data1, "!contents? measPoint$"  'get point number
        measPoint=val(measPoint$)
        call calManGetFreqInput measPoint   'Get point data into boxes and wait
        #calman.data2, "!contents? s2$" : #calman.data3, "!contents? s3$"

        'Frequency cal. data2 is freq, data3 is power;
        'if first point, need to put freq in ref
        freq=val(s2$)
        if freq<0 or freq>10000 then notice "Invalid frequency." : calManEnterError=1 :exit sub 'ver114-3d
        pow=val(s3$)
        f$=using("####.######",freq)
        p$=using("####.###",pow)  'ver115-1e
        if pow<-200 or pow>100 then notice "Invalid Measured Power." : calManEnterError=1 :exit sub  'ver114-3b; ver114-3d
        if calManEntryIsRef then  'Get ref freq and power first time only
            calManRefFreq=freq
            calManRefPower=pow
            truePower=pow       'ver114-3b
            print #calman.ref, p$
            print #calman.ref2,f$
        else    'after first time, get true power from .ref; ver114-3b created this else... block
            #calman.ref, "!contents? s4$"
            truePower=val(s4$)
            if truePower<-200 or truePower>100 then notice "Invalid True Power value." : calManEnterError=1 : exit sub 'ver114-3d ver115-1a
        end if
        'error factor is true power minus measured power; this is added to measured power when freq cal is applied ver 114-3b
        p$=using("####.###",truePower-pow)  'SEW9 reversed the subtraction. ver114-3b used truePower  'ver115-1e
        print #calman.te, f$;" ";p$
    else
        'Mag cal. data 1 is power, data2 is ADC, data3 is phase (if we are doing phase);
        'if first point, need to be sure freq is in ref and put phase in ref2
        pow=val(s1$)
        if pow<-200 or pow>100 then notice "Invalid dBm value." : calManEnterError=1 : exit sub 'ver116-1b
        adc=val(s2$)
        if adc<0 then notice "Invalid adc value." : calManEnterError=1 : exit sub 'ver114-3d
        if calGetDoPhase()=0 then phase=0 else phase=val(s3$)
        if phase<-180  or phase>180 then notice "Invalid phase value." : calManEnterError=1 : exit sub 'ver114-3d
        adc$=using ("#######", adc)
        v$=using ("####.###", pow)
        p$=using ("####.##", phase)
        forceVal=0 : #calman.Force, "value? doForce$"   'ver116-4j added forced phase items
        if doForce$="set" then #calman.ForceVal, "!contents? forceVal$" : forceVal=val(uCompact$(forceVal$))
        if forceVal<-180  or forceVal>180 then notice "Invalid forced phase value. Zero used." : forceVal=0
        if calManEntryIsRef then  'Get freq first time only; after that it is fixed
            #calman.ref, "!contents? freq$"
            freq=val(freq$)
            if freq<=0 or freq>10000 then notice "Invalid frequency." : calManEnterError=1 : exit sub 'ver114-3d
            print #calman.ref2, using("####.##", phase)  'Fix reference phase first time only
            calManRefPhase=phase : calManRefFreq=freq : calManRefPower=pow
        end if
            'The power correction factor is phase-calManRefPhase, put into normal range
        if calGetDoPhase()=1 then
            if doForce$="set" then pCor=forceVal else pCor=phase-calManRefPhase    'ver116-4j
            while pCor<=-180 : pCor=pCor+360: wend
            while pCor>180 : pCor=pCor-360: wend
            p$=using ("####.##", pCor)
                'Enter the mag cal data, with phase
            print #calman.te, adc$;"  ";v$;"  ";p$ 
        else
                'Enter the mag cal data, without phase
            print #calman.te, adc$;"  ";v$
        end if
    end if

        'We have now entered the data, whether mag or freq. Now clean things up.
    if btn$<>"" or calManEntryIsRef then
        fErr$=calLoadFromEditor$("#calman.te", calEditorPathNum) 'Load variables from editor
        if fErr$<>"" then
            'Error in existing data, so we won't mess with formatting or fixing the comments
            notice "Error in data: " + fErr$
            if calManEntryIsRef then call calManEnterInstructions ""
            if doForce$="set" then #calman.Force, "set" : #calman.ForceVal, forceVal$ : #calman.ForceVal, "!show"    'gets hidden by calManEnterInstructions ver116-4j
            calManEntryIsRef=0
            exit sub
        end if
    end if

    if calManEntryIsRef=1 then
        'First point--Enter info into comments
        c2$="Calibrated " + Date$("mm/dd/yy") + " at "
        if calEditorPathNum=0 then
            'Freq
            c1$="Calibration over frequency"
            c2$=c2$ + Trim$(using ("####.###", calManRefPower)) + " dBm."  'ver116-1b
        else
            'Mag
            filtFreq=MSAFilters(calEditorPathNum,0) : filtBW=MSAFilters(calEditorPathNum,1)
            c1$="Filter Path " + str$(calEditorPathNum) + ": CenterFreq=" _
                            + Trim$(using("###.######",filtFreq)) +" MHz; Bandwidth=" _
                            + Trim$(using("####.######",filtBW)) +" kHz"    'ver116-1b
            c2$=c2$ + Trim$(using("####.######", calManRefFreq)) + " MHz."
        end if
        calFileComments$(1)=c1$
        calFileComments$(2)=c2$
        calManEntryIsRef=0
        call calManEnterInstructions ""
        if doForce$="set" then #calman.Force, "set" : #calman.ForceVal, forceVal$ : #calman.ForceVal, "!show"    'gets hidden by calManEnterInstructions ver116-4j
    end if
        'SEW8 made the following conditional
    if btn$<>"" or calManEntryIsRef then
        #calman.te, "!cls"
        call calSaveToEditor "#calman.te", calEditorPathNum   'Restore data, now with comments
    end if

    if calEditorPathNum=0 then  'SEW8 created this if... block
        call calManGetFreqInput calManLastAutoPoint 'Display last retrieved point
    else    'Path calibration
        #calman.data1, "" : #calman.data2, "" : #calman.data3, ""   'SEW6
    end if

end sub

function calFileExists()    'Return 1 if both the freq file and path 1 mag file exist
    On Error goto [calNoFile]
    open calFilePath$()+calFileName$(0) for input as #calIn 'Frequency cal file
    close #calIn
    open calFilePath$()+calFileName$(1) for input as #calIn 'Path 1 mag cal
    close #calIn
    calFileExists=1
    exit function
[calNoFile]      'Error means it isn't there
    calFileExists=0
end function

[calManBtnNextFreqPoint]    'Handles button to retrieve next point in automated freq cal.
    #calman.data1, "!contents? measPoint$"  'get point number
    measPoint=val(measPoint$)
    if measPoint<0 or measPoint>steps-1 then notice "Invalid point number" : wait
    measPoint=measPoint+1   'Next point
    #calman.data1, str$(measPoint)  'Display point number
    call calManGetFreqInput measPoint   'Get point data into boxes and wait
    wait

 'SEW8 added calManBtnPrevFreqPoint
[calManBtnPrevFreqPoint] 'Handles button to retrieve previous point in automated freq cal.
    #calman.data1, "!contents? measPoint$"  'get point number
    measPoint=val(measPoint$)
    if measPoint<1 or measPoint>steps then notice "Invalid point number" : wait
    measPoint=measPoint-1   'Next point
    #calman.data1, str$(measPoint)  'Display point number
    call calManGetFreqInput measPoint   'Get point data into boxes and wait
    wait

sub calManGetFreqInput  measStep 'Enter user scanned data for point measStep (0...steps)
        'The user has performed a scan covering the desired range with the desired number of points.
        'We enter the points from the MSA arrays into the calibration table.
        'This is valid if the Path calibration for the currently installed filter is current.
        'We retrieve freq and power (dbm) for each point and put it into calman.data2 (freq)
        'and calman.data3 (power). We put point number into calman.data1.
    calManLastAutoPoint=measStep       'SEW8 Record number of point of the user-established sweep
    #calman.data1, using("####", measStep)  'Enter step number into box
    #calman.data2, using("####.######", datatable(measStep,1))    'Enter frequency into box
    #calman.data3, using("####.###", datatable(measStep, 2))     'Enter power into box  'ver115-1e
end sub

[calManMenuMeasure] 'enters here when the "Measure" button is clicked. ver113-7d
    gosub [calManMeasure]
    wait

[calManMeasure]
        'Measure a point for calibration over signal strength (mag cal)
        'data 1 is power, data2 is ADC, data3 is phase (if we are doing phase);
        'freq is in ref. We measure power and phase at this frequency and enter into boxes.
        'It is assumed that the user has set up the scan at the desired frequency.Scotty-yes, in zero sweep width
    #calman.ref, centfreq  'ver113-7d
    if centfreq<0 or centfreq>1000 then notice "Invalid frequency." : return  'ver113-7d
    if centfreq=0 then notice Chr$(13);"Center Frequency of 0 is not allowed.";Chr$(13);"Return to MSA and Change Center Frequency" : return  'ver113-7d
    if sweepwidth>0 then notice Chr$(13);"Sweep Width is not 0.";Chr$(13);"Return to MSA and Change Sweep Width to 0" : return  'ver113-7d
        'We have to average 30 readings.  For mag, we just average ADC, since we are
        'just now determining the calibration factors needed to convert ADC to power.
        'For phase, if needed, we convert from ADC bits to phase and average that.
        'Averaging phase is a little tricky.
        'A -179 degree angle and a 179 degree angle should average to 180, not 0 degrees.
        'We expect only a few degrees variation in phase. We save the first phase reading and
        'average the difference (this reading-first reading). If a difference exceeds 180 degrees,
        'we decrease it by 360 because we know we must take the difference in the opposite
        'direction. E.g. 179 minus -179 gives a difference of 358; we subtract 360 to get the true difference.
        'of -2 degrees. Likewise, if the difference is less than -180 degrees, we add 360.
    measADCSum=0 : measPhaseDifSum=0
            'Average the mag ADC readings and the phase degrees.
    saveMode$=msaMode$ 'Save MSA current setting
    saveThreshold=validPhaseThreshold :validPhaseThreshold=0    'So PDM inversions will be done even at low levels ver116-1b
    if calGetDoPhase()=1 then msaMode$="VectorTrans" else msaMode$="SA"  'Include phase if possible 'ver114-5n
    suppressPDMInversion=0  'ver115-1a
    suppressPhase=0 'If user has suppressed phase, undo it ver116-4j
    #calman.data2, "" : #calman.data3, ""   'Blank out ADC and phase so user sees Measure is in progress ver115-1a
    'Wait time is accomplished primarily by the time between the user setting the input level and
    'our first read, but just in case we delay the first read. We set the wait time based on the time constants.
    'Note that once we take the first read, there are no more glitches from changing PLLs, etc. since freq is fixed.
    saveWate=wate : saveUseAutoWait=useAutoWait   'ver116-4j
    useAutoWait=0 : if calGetDoPhase() then wate=max(videoMagTC, videoPhaseTC) else wate=videoMagTC 'ver116-4j
    if wate<150 then wate=150 'ver116-4j
    thisstep=0 : gosub [CommandCurrentStep] 'ver116-4j  to be sure we are set to proper freq, etc.; assumes [PartialRestart] has been done
    glitchhlt = 8*wate  'add extra settling time before first read. [ReadStep] uses this.
    for measStep=0 to 29  'ver116-4j
            'Read phase and mag. 
        gosub [ReadStep]    'Read phase and mag
            'magdata now has ADC bits for magnitude; if MSA can measure phase phadata has ADC bits for phase.
        measADCSum=measADCSum+magdata 
        if calGetDoPhase()=1 then   
                'phaarray(thisstep,0) is 1 if inverted otherwise 0; phadata is phase ADC reading
            measPhase=360*phadata/maxpdmout-invdeg*phaarray(thisstep,0) 'ver115-1a
            if measStep=0 then measFirstPhase=measPhase 'Save first phase reading
            measPhaseDif=measPhase-measFirstPhase   'phase dif
            if (phadata < pdmlowlim or phadata > pdmhighlim) then   'Phase in crap zone even with inversion ver114-7k
                suppressPDMInversion=1  'Don't waste time inverting for remaining points
            end if
            if measPhaseDif<-180 then measPhaseDif=measPhaseDif+360
            if measPhaseDif>180 then measPhaseDif=measPhaseDif-360
            measPhaseDifSum=measPhaseDifSum+measPhaseDif    'Accumulate phase difference
        end if
    next measStep
    #calman.data2, using("######", measADCSum/30)   'Enter averaged ADC into box ver116-4j
    if calGetDoPhase()=1 then   'this block rewritten by ver115-1a
        if suppressPDMInversion then
            measPhase=calManRefPhase+180 'Choose value so phase correction comes out to +/-180
            if measPhase>180 then measPhase=measPhase-360
        else
            measPhase=measFirstPhase + measPhaseDifSum/30   'First phase measurement plus average difference ver116-4j
            while measPhase<=-180 : measPhase=measPhase+360 : wend  'Put into range -180 to 180
            while measPhase>180 : measPhase=measPhase-360 : wend
        end if
    else
        measPhase=0
    end if
    suppressPDMInversion=0  'ver115-1a
    #calman.data3, using("####.##", measPhase)     'Enter averaged phase into box
    msaMode$=saveMode$ 'Restore settings
    wate=saveWate : useAutoWait=saveUseAutoWait   'ver116-4j
    validPhaseThreshold=saveThreshold   'ver116-1b
return

sub calManClean  btn$     'Format and Sort list
    'We just get the data and put it back
    fErr$=calLoadFromEditor$("#calman.te", calEditorPathNum) 'This also sorts
    if fErr$<>"" then notice "Bad Data. "+fErr$ : exit sub
    if calEditorPathNum=0 then nPoints=calNumFreqPoints() else nPoints=calNumMagPoints()
    #calman.te, "!cls"
    call calSaveToEditor  "#calman.te", calEditorPathNum
    if nPoints=0 then
        'If there are no data points, start over with a clean header
        call calCreateDefaults calEditorPathNum,"#calman.te", 0 'zero signals no points
    end if
end sub

sub calManDisplayDefault btn$
   if calManWarnToSave(1)=1 then exit sub
   call calCreateDefaults calEditorPathNum,"#calman.te", 1 '1 signals to do points
   call calManClean ""
   call calManPrepareEntry
end sub

[calManSelectPath]    'Handle user selection of calibration file ver116-1b made this a branch label menu item
    'Can be called only from menu
    'Set current path number. Will be negative if there is no selection
    if calManWarnToSave(1)=1 then wait   'cancelled ver116-1b
    prevPathNum=calEditorPathNum
    #calman.pathList, "selectionindex? sel"
    if sel=0 then notice "No File Selected." : wait  'Must be error ver116-1b
    calEditorPathNum=sel-1    'sel starts at 1 if valid selection; calEditorPathNum starts at 0
    if calManFile("Reload")=1 then  'Load new file ver116-1b
        'User cancelled the switch
        #calman.pathList, "selectindex ";prevPathNum+1  'index is one more than path number
        calEditorPathNum=prevPathNum
    end if

    if calEditorPathNum=0 then
        #calman.teLabel, "Frequency Calibration Table"
    else
        #calman.teLabel, "Path Calibration Table"
        path$="Path ";calEditorPathNum   'set main MSA path variable ver116-1b
        call SelectFilter filtbank : gosub [PartialRestart]  'ver116-4j
    end if
    calManRefPhase=0    'ver116-4j
wait

sub calManEnterAvailablePaths    'Create list of available paths
    'The list is based on the MSA info in MSAFilters(), which lists
    'freq and bw for each calibration path, starting at path 1=index 1
    'We create a list where index 0 is the frequency cal file, and index N (1...)
    'is the Nth mag cal path.

    for i=0 to 40: calManFileList$(i)="": next i   'Clear list
    calManFileList$(0)="0 (Frequency)"
    for i=1 to MSANumFilters
        freq$=str$(MSAFilters(i,0)) : bw=MSAFilters(i,1) : bw$=str$(bw)
        freq$=freq$+Space$(max(0,7-len(freq$)))   'makes freq$ a fixed width
        if bw=0 then bwPad$="   " else _
                bwPad$=Space$(5-int(Log(bw)/Log(10))) 'Aligns bw$ at decimal;
        bw$=bwPad$+str$(bw)
        configFormatFilter$=freq$+bw$ 
        thisFilt$=freq$+bw$
        calManFileList$(i)=str$(i) + " (" + Trim$(thisFilt$) + ")"   'Path num plus freq and bw
    next
    #calman.pathList, "reload"  'Loads new strings into listbox
end sub

sub calManPrepareEntry
        'hide all entry boxes, labels and buttons
    #calman.Lref, "!hide" : #calman.ref, "!hide" : #calman.ref2, "!hide"
    #calman.Lref2, "!hide" : #calman.data1, "!hide" : #calman.Ldata1, "!hide"
    #calman.data2, "!hide" : #calman.Ldata2, "!hide" : #calman.data3, "!hide"
    #calman.Ldata3, "!hide" : #calman.Measure, "!hide" : #calman.Enter, "!hide"
    #calman.NextPoint, "!hide" : #calman.PrevPoint, "!hide" : #calman.EnterAll, "!hide"
    #calman.StartEntry, "!show"
    #calman.Force, "hide" : #calman.ForceVal, "!hide"  'ver116-4j
    #calman.enterInstruct, "cls"
    #calman.enterInstruct, "place 10 15"
    #calman.enterInstruct, "\To begin entry of calibration data, click Start Entry."
    #calman.enterInstruct, "\Alternatively, you may enter, alter and delete data in the text editor."
    calManEntryIsRef=1    'We are set for the first point
end sub

sub calManEnterInstructions btn$ 'Display instructions for entering data
    'calManEntryIsRef=1 for entering the first (reference) point; =0 for other points
    #calman.enterInstruct, "cls"
    #calman.enterInstruct, "place 10 15"
    #calman.StartEntry, "!hide"
    #calman.Enter, "!show"
    #calman.Lref, "!hide"   'Reference and its label hide
    #calman.ref, "!hide"
    #calman.data1, "!show"   'data 1 and its label show
    #calman.Ldata1, "!show"
    #calman.data2, "!show"    'data2 and its label show
     #calman.Ldata2, "!show"
    #calman.data2, "!enable"  'data2 enable  ver 114-3d
    #calman.data3, "!show"   'data3 and its label show
    #calman.Ldata3, "!show"
    #calman.data3, "!enable"  'data3 enable  ver 114-3d
    #calman.ref2, "!hide"  'ref2 and its label hide
    #calman.Lref2, "!hide"
    #calman.data2, ""   'Clear data boxes
    #calman.data3, ""

    if calEditorPathNum=0 then
        'Here we are doing freq cal. We need frequency and measured power for each point.
        'These are entered from sweep data, so the boxes are disabled to prevent the user
        'from directly entering data.
        #calman.data2, "!disable"  'disable freq    'ver 114-3d
        #calman.data3, "!disable"  'disable power   'ver 114-3d
        #calman.Lref, "True Power(dBm)" 'ver116-1b
        #calman.Ldata1, "Point Number"  'SEW8
        #calman.Ldata2, "Freq (MHz)"
        #calman.Ldata3, "Measured Power (dBm)" 'ver116-1b
        #calman.ref, "!enable" 'ver114-3b
        #calman.ref2, "!disable"
        #calman.Measure, "!hide"      'SEW8 Hide Measure and show NextPoint, PrevPoint instead
        #calman.NextPoint, "!show"
        #calman.PrevPoint, "!show"
        #calman.EnterAll, "!show"
        #calman.Force, "hide" : #calman.ForceVal, "!hide"  'ver116-4j
        
            'ver113-7e altered the following instructions for each "#calman.enterInstruct" modver115-1a
        #calman.enterInstruct, "\Frequency Calibration should be performed with an input power of -20 dBm to -40 dBm."
        #calman.enterInstruct, "\A calibration list has been created from the previous sweep, with the number of Points"
        #calman.enterInstruct, "\equaling the number of steps in the sweep, +1. Data of each Point can be accessed by"
        #calman.enterInstruct, "\clicking the Next Point button.  You must change the value in the True Power box to"
        #calman.enterInstruct, "\the Known Input Power. You may modify any data by highlighting it and retyping it."
        #calman.enterInstruct, "\You may insert data from a single Point by clicking the Enter button. You may insert"
        #calman.enterInstruct, "\the data of every Point in the sweep by clicking the Enter All button. You may insert"
        #calman.enterInstruct, "\data by typing the data directly into the Frequency Calibration Table, after the last"
        #calman.enterInstruct, "\data Point. Data may be entered in any order. Clean Up to sort the data at any time."
        #calman.enterInstruct, "\The Magnitude Error vs Frequency Correction Factor = True Power - Measured Power."

        if calManEntryIsRef=1 then
                'first point for frequency cal
                'SEW8 Rewrote the following message
            notice "Notice" +chr$(13)+"Automatic entry of points assumes you have performed " _
                                +chr$(13)+"a sweep over the desired frequency range before entering the" _
                                +chr$(13)+"Calibration Mangager. Numbers (0...) refer to points of that sweep."
            #calman.ref, ""
            #calman.ref2, ""
            call calManGetFreqInput 0   'SEW8 Display data for point 0
        else
            '2nd+ points for frequency cal
            #calman.ref, "!show"       'show power ref
            #calman.Lref, "!show"
        end if
    else
        'Here we are doing mag cal. We need Input power, ADC and possibly phase for each point
        'For first point we need frequency.
        #calman.data1, ""   'Clear data box
        #calman.Lref2, "Ref Phase"
        #calman.Ldata2, "ADC value"
        #calman.Ldata3, "Phase (degrees)" 'ver114-3b Note label is two lines high
        #calman.Ldata1, "Input (dBm)"   'ver116-1b
        #calman.ref2, "!disable"     'Don't allow reference phase to be changed
        #calman.Force, "show" : #calman.Force, "reset" : #calman.ForceVal, "0"  : #calman.ForceVal, "!hide"  'ver116-4j
        if calGetDoPhase() =0 then
            #calman.data3, "!hide"  'data3 and its label hide (not doing phase)
            #calman.Ldata3, "!hide"
            #calman.Force, "hide" : #calman.ForceVal, "!hide"  'ver116-4j
        end if
        #calman.ref, "!show"           'ref will hold reference frequency
        #calman.Lref, "!show"
        #calman.Lref, "Ref Freq (MHz)"
        #calman.Measure, "!show"      'SEW8 Hide Measure and show NextPoint, PrevPoint instead
        #calman.Measure, "!enable"
        #calman.NextPoint, "!hide"
        #calman.PrevPoint, "!hide"
        #calman.EnterAll, "!hide"
        if calGetDoPhase() =0 then
            #calman.enterInstruct, "\For each point, enter the input power level and the ADC reading, then click Enter."
            #calman.enterInstruct, "\For the first point also enter the reference frequency at which the measurements"
            #calman.enterInstruct, "\made. The power level may be entered manually or automatically measured with Measure."
            #calman.enterInstruct, "\or automatically measured with Measure."
            #calman.enterInstruct, "\Data may be entered in any order; to sort the data click Clean Up"
            #calman.enterInstruct, "\You may add, alter or delete lines directly in the text editor if desired."

        else  'ver113-7 changed the following instructions:
            #calman.enterInstruct, "\The Spectrum Analyzer must be configured for zero sweep width. Center Frequency must"
            #calman.enterInstruct, "\be higher than 0 MHz. The first data Point will become the Reference data for all"
            #calman.enterInstruct, "\other data Points. Click Measure button to display the data measurements for ADC"
            #calman.enterInstruct, "\value and Phase. Manually, enter the Known Power Level into the Input (dBm) box."
            #calman.enterInstruct, "\Click the Enter button to insert the data into the Path Calibration Table."
            #calman.enterInstruct, "\Subsequent Data may be entered in any order, and sorted by clicking Clean Up. ADC"
            #calman.enterInstruct, "\bits MUST increase in order and no two can be the same. You may alter the Data in"
            #calman.enterInstruct, "\the table, or boxes, by highlighting and retyping. The Phase Data (Phase Error vs"
            #calman.enterInstruct, "\Input Power) = Measured Phase - Ref Phase, is Correction Factor used in VNA."
            #calman.enterInstruct, "\Phase is meaningless for the Basic MSA, or MSA with TG."

        end if
        if calManEntryIsRef=1 then
            'first point for mag cal
            #calman.ref, "!enable"    'Allow ref freq to be changed on first point only
            #calman.ref, ""
            #calman.ref2, ""
                'Measure can be used only if this mag file matches active filter path
            f=MSAFilters(calEditorPathNum,0) : bw=MSAFilters(calEditorPathNum,1)
            if f=finalfreq and bw=finalbw then haveProperFilt=1 _
                                else haveProperFilt=0
            if haveProperFilt=1 then
                #calman.Measure, "!enable"
            else
                #calman.Measure, "!disable"
                notice "Notice"+chr$(13)+"Measurements made at this time will not be accurate, because " _
                                +chr$(13)+"the selected filter is not the currently installed filter."
            end if
        else
            '2nd+ points for mag cal
            #calman.ref, "!disable"    'Allow ref freq to be changed on first point only
            if calGetDoPhase() =1 then
                #calman.ref2, "!show"
                #calman.Lref2, "!show"   'ref2 and its label show (phase ref)
                #calman.Lref2, "Ref Phase(deg)"
            end if
        end if
    end if
    #calman.enterInstruct, "flush"  'Make graphics stick
end sub

function calManWarnToSave(allowCancel)   'Give warning to save if data has changed. Return 1 to cancel
    'calManOldText$ contains the old data to compare against. This method is used
    'because the !modified? command cannot be reset at the time of a Save.
    'If allowCancel=1, the message box will allow the user to cancel
    can=0
    #calman.te, "!contents? curr$"  'get current data
    if calManOldText$=curr$ then calManWarnToSave=0: exit function

        'Put up a message box.
    msg$="Unsaved calibration changes will be lost. Do you want to SAVE first?"
    ans$=uPrompt$("Warning", msg$,1,allowCancel)    'yes, no, possibly cancel
    if ans$="no" then calManWarnToSave=0: exit function    'Proceed, no save
    if ans$="cancel" then calManWarnToSave=1 : exit function   'Cancel
    errMsg$=calLoadFromEditor$("#calman.te", calEditorPathNum)   'Get current text editor data ver116-1b
    if errMsg$<>"" then   'ver116-1b
        notice "Data Error: "+errMsg$+ "; file not saved."  'ver116-1b
        calManWarnToSave=allowCancel        'cancel if possible ver116-1b
        exit function  'ver116-1b
    end if
    call calSaveToFile calEditorPathNum     'User wants to save so do it
    call calInstallFile calEditorPathNum     'Re-read to check for errors SEWcal2,ver113-7g
    #calman.te, "!contents? calManOldText$"  'Save for future compare
    calManWarnToSave=0  'No cancel
end function

sub calManBtnFile btn$   'Handle file buttons
    pos=instr(btn$,".")
    if pos>0 then btn$=Mid$(btn$, pos+1)    'Drop part before period
    r=calManFile(btn$)  'r indicates whether cancelled; we don't care.
end sub

function calManFile(btn$)   'Load/save/create files returns 0, except 1 if user cancels
    'All file loads must be done through this function
    calManFile=0
    select case btn$
        case  "Load", "Reload"
            if calEditorPathNum<0 then notice "You must select a calibration file to save."
            if calEditorPathNum=0 then
                        'Before loading the frequency file we must be sure that the mag cal
                    'data is valid for the MSA's active filter
                match=0
                for i=1 to MSANumFilters
                    f=MSAFilters(i, 0) : bw=MSAFilters(i,1)
                    if f=finalfreq and bw=finalbw then match=i : exit for
                next i
                if match=0 then notice "Can't find active filter; don't use Measure."    'only possible in debugging
                call calInstallFile match  'Loads mag cal file
            end if
            call calInstallFile calEditorPathNum
            #calman.te, "!cls"
            call calSaveToEditor "#calman.te", calEditorPathNum 'display data
            call calManPrepareEntry 'Prepare buttons and instructions for data entry
        case "Save"
            if calEditorPathNum<0 then notice "You must select a calibration file to save."
            errMsg$=calLoadFromEditor$("#calman.te", calEditorPathNum)
            if errMsg$<>"" then notice "Data Error: "+errMsg$+ "; file not saved." : exit function
            call calSaveToFile, calEditorPathNum
            dum=calManFile("Load")    'Reload to update editor and verify data
        case "Default"
            if calManWarnToSave(1)=1 then calManFile=1 : exit function   'Cancelled  moved here by ver116-1b
            if calEditorPathNum<0 then notice "You must first select a calibration file."
            call calCreateDefaults calEditorPathNum, "", 1
            call calManPrepareEntry
        case else
    end select
    #calman.te, "!contents? calManOldText$"  'Save for future compare
end function

'@====================Original Calibration Manager Ended Here=====================

'=========================================================================================
'@==========================Start of Mag/Freq Calibration Module==========================
'Mag/Freq Calibration module version 1.03
'
'This module is now a subset of the Configuration Manager Module. This portion
'manages calibration files and applies them upon request to raw data.
'calInitFirstUse must be called before any other routines, but that is
'taken care of if calManRunManager has already been run.
'The principal routines of interest to the user are:
'
'function calConvertMagError(magdata) 'Returns dbm level given raw ADC reading
'function calConvertFreqError(freq) 'Returns correction factor for this freq,
                                    'to be added to raw power to get final

'

'---------------Start Global Variables for Mag/Freq Calibration Module---------
        'calMagTable contains calibration data for the current signal path
        'The zero entry for the first dimension is not used.
        'magCalTable(i, v) gives the ith entry; v=0 gives ADC reading; v=1 gives actual db value.
    global calMaxMagPoints, calMaxFreqPoints
    dim calMagTable(100,2)  'Size can be changed dynamically
        'calMagTable contains calibration data for the current signal path
        'The zero entry for the first dimension is not used.
        'calMagTable(i, v) gives the ith entry; v=0 gives freq (in MHz); v=1 gives dbm magnitude;
        'v=2 gives phase correction to be subtracted from raw phase readings, used only for VNA.
    'Each entry of calMagCoeffTable() will have 8 numbers. 0-3 are the A,B,C,D
    'coefficients for interpolating the real part; 4-7 are for the imag part.
    'dim calMagCoeffTable(100,7)    'SEWcal Cubic coefficients for interpolating in calMagTable 'ver115-1d moved to top of program

    global calMagPoints  'Number of actual points in calMagTable
    global calMagFileHadPhase 'Set to 1 if file data had third data element (for VNA)
                            'Otherwise, 0.
    global calDoPhase     'Set to 1 to treat mag cal data as including phase; otherwise phase ignored.

         'calFreqTable contains calibration data for response over frequency
        'The zero entry for the first dimension is not used.
        'calFreqTable(i, v) gives the ith entry; v=0 gives freq (in MHz);
        'v=1 gives magnitude correction.
    dim calFreqTable(800,1)     'Size can be changed dynamically
     'calFreqCoeffTable has the A,B,C,D coefficients for interpolating in the calFreqTable()
    'dim calFreqCoeffTable(800,3)    'SEWcal Cubic interpolation coefficients for calFreqTable()  'ver115-1d moved to top of program
    global calFreqPoints  'Number of actual points in calFreqTable

    dim calFileComments$(2)    'Up to 2 lines of comments from files
    global calFileCommentChar$     'Character marking comments in data files
    global calModuleVersion 'Version of this module; set in calInitFirstUse
    global calFileVersion   'Version of last opened file

    dim calFileInfo$(10,3)  'Used internally to request file info

'------------Routines------------------

sub calInitFirstUse maxMagPoints, maxFreqPoints, doPhase    'Call to initialize before using other routines
    if maxMagPoints<=800 then calMaxMagPoints=801 else calMaxMagPoints=maxMagPoints+1 'ver114-4b
    if maxFreqPoints<=800 then calMaxFreqPoints=801 else calMaxFreqPoints=maxFreqPoints+1 'ver114-4b
    calFileCommentChar$="*"        'Use asterisk as comment char
    redim calMagTable(calMaxMagPoints,2)
    redim calFreqTable(calMaxFreqPoints,1)
    redim calMagCoeffTable(calMaxMagPoints,7)   'SEWcal
    redim calFreqCoeffTable(calMaxFreqPoints,3)   'SEWcal   'ver114-4b
    calMagFileHadPhase=0
    calDoPhase=doPhase  'ver114-5q
    calModuleVersion=1.03         'Change this code to update version
end sub

function calVersion()    'Version of this module
    calVersion=calModuleVersion
end function

sub calSetMaxPoints maxMagPoints, maxFreqPoints    'Call to change max points; data is lost
    calMaxMagPoints=max(801, maxMagPoints+1) 'ver114-4b
    calMaxFreqPoints=max(801, maxFreqPoints+1) 'ver114-4b
    redim calMagTable(calMaxMagPoints,2)
    redim calMagCoeffTable(calMaxMagPoints,7)   'SEWcal
    redim calFreqTable(calMaxFreqPoints,1)
    redim calFreqCoeffTable(calMaxFreqPoints,3)   'SEWcal
end sub

sub calClearMagPoints       'Clear calMagTable() to zero points
    'It's not necessary to zero the points, but it keeps things clean
    for i=1 to calMaxMagPoints
        calMagTable(i,0)=0 :calMagTable(i,1)=0:calMagTable(i,2)=0
    next i
    calMagPoints=0
end sub

sub calClearFreqPoints      'Clear calFreqTable() to zero points
     'It's not necessary to zero the points, but it keeps things clean
    for i=1 to calMaxFreqPoints
        calFreqTable(i,0)=0 :calFreqTable(i,1)=0
    next i
    calFreqPoints=0
end sub

function calDataHadPhase()      'Returns 1 if last file or editor data had phase correction data
    calDataHadPhase=calMagFileHadPhase
end function

sub calSetDoPhase doPhase   'Set calDoPhase to 1 or 0, indicating whether output should include phase
    calDoPhase=doPhase
end sub

function calGetDoPhase()   'Get calDoPhase:or 0, indicating whether output should include phase
    calGetDoPhase=calDoPhase
end function

sub calClearComments
    calFileComments$(1)=""
    calFileComments$(2)=""
end sub

function calNumMagPoints()          'Get number of points in calMagTable
    calNumMagPoints=calMagPoints
end function

function calNumFreqPoints()          'Get number of points in calFreqTable
    calNumFreqPoints=calFreqPoints
end function

sub calGetMagPoint N, byref adc, byref db, byref phase       'Get mag data for point N
    if N<=0 or N>calMagPoints then notice "Invalid mag cal point number."
    adc=calMagTable(N,0)
    db=calMagTable(N,1)
    phase=calMagTable(N,2)
end sub

sub calGetFreqPoint N, byref f, byref db       'Get mag data for point N
    if N<=0 or N>calFreqPoints then notice "Invalid frequency cal point number."
    f=calFreqTable(N,0)
    db=calFreqTable(N,1)
end sub

function calAddMagPoint(adc, db, phase)    'Add mag cal point; return 1 if error
    if calMagPoints>=calMaxMagPoints then calAddFreqPoint=1: exit function
    calMagPoints=calMagPoints+1
    calMagTable(calMagPoints,0)=adc
    calMagTable(calMagPoints,1)=db
    calMagTable(calMagPoints,2)=phase
    calAddMagPoint=0        'No error
end function

function calAddFreqPoint(f, db)         'Add freq point; return 1 if error
    if calFreqPoints>=calMaxFreqPoints then calAddFreqPoint=1: exit function
    calFreqPoints=calFreqPoints+1
    calFreqTable(calFreqPoints,0)=f
    calFreqTable(calFreqPoints,1)=db
    calAddFreqPoint=0        'No error
end function

sub calSortMag    'Sort mag cal table in ascending order of ADC reading
    sort calMagTable(),1, calMagPoints,0
end sub

sub calSortFreq   'Sort freq cal table in ascending order of frequency.
    sort calFreqTable(),1, calFreqPoints,0
end sub

sub calCreateMagCubicCoeff     'Create table of cubic coefficients for mag cal table and set validPhaseThreshold
    'Each entry of the mag coefficient table will have 8 numbers. 0-3 are the A,B,C,D
    'coefficients for interpolating the real part; 4-7 are for the phase correction.

    call intSetMaxNumPoints calMagPoints  'Be sure we have room ver115-9d
    call intClearSrc
    validPhaseThreshold=0  'ver116-1b
    leftADC=-1 : leftDB=-150  'ver116-1b
    centerSlopeFound=0
    for i=1 to calMagPoints 'copy cal table to intSrc
        'Phase corrections of 180 or -180 indicate phase is invalid at this ADC value for magnitude
        'We set validPhaseThreshold to 1 greater than the highest ADC value with such a correction.
        thisPhaseCorrection=calMagTable(i,2) 'ver116-1b
        thisADC=calMagTable(i,0)
        thisDB=calMagTable(i,1)
        if thisPhaseCorrection=180 or thisPhaseCorrection=180 then validPhaseThreshold=calMagTable(i,0)+1 'ver116-1b
        call intAddSrcEntry thisADC, thisDB ,thisPhaseCorrection 'ver116-1b
        'Save first point at -60 dBm or higher
        if leftADC=-1 and thisDB>=-60 then leftADC=thisADC : leftDB=thisDB  'ver116-1b
        'Save slope of ADC vs. dB line for a centrally located slope for at least a 20 dB segment.
        'This is used for adaptive wait times
        if leftADC<>-1 and centerSlopeFound=0 and (thisDB-leftDB) >=20 then   'ver116-1b
            calCenterSlope=(thisADC-leftADC)/(thisDB-leftDB)
            highPointNumOfCenterSlope=i 'This will be refined below ver116-1b
            centerSlopeFound=1
        end if
    next i

        'Signal that our "imaginary" part is an angle, and not to "favor flat", because
        'we expect the db as a function of ADC to become vertical near the ends.
        'Do phase only if calGetDoPhase()=1
    favorFlat=0 : isAngle=1
    doPhase=calGetDoPhase() 'ver116-1b
    'First "1" means do magnitude; second 1 means we are doing phase correction table 'ver116-1b
    call intCreateCubicCoeffTable 1,doPhase,isAngle, favorFlat, 1   'ver116-1b
    for i=1 to calMagPoints  'put the data where we want it
        calMagCoeffTable(i,0)=intSrcCoeff(i,0) :calMagCoeffTable(i,1)=intSrcCoeff(i,1)
        calMagCoeffTable(i,2)=intSrcCoeff(i,2) :calMagCoeffTable(i,3)=intSrcCoeff(i,3)
        calMagCoeffTable(i,4)=intSrcCoeff(i,4) :calMagCoeffTable(i,5)=intSrcCoeff(i,5)
        calMagCoeffTable(i,6)=intSrcCoeff(i,6) :calMagCoeffTable(i,7)=intSrcCoeff(i,7)
    next i


        'We re-iterate the calibration table to determine a couple other items. For calculating
        'adaptive wait times, we need to know the approximate slope (delta ADC)/(delta dB).
        'We potentially divide the response into three sections:
        'For ADC<calLowADCofCenterSlope the slope is calLowEndSlope, calculated from at least a 5 dB segment near the
        '       low end, but not so low as to have slopes less than 7% of the center slope
        'for calLowADCofCenterSlope < ADC < calHighADCofCenterSlope the slope is calCenterSlope, calculated
        '        from a 20 dB or greater segment somewhere in the center area
        'for ADC>calHighADCofCenterSlope the slope is calHighEndSlope calculated from a 7 dB or greater segment at top end
        'We need a reasonable range of values to be able to do this.  ver116-1b
    calCanUseAutoWait=1
    slopeErr=(calMagPoints<5)
    if slopeErr=0 then if calMagTable(1,1)>-60 or calMagTable(calMagPoints, 1)<-40 then slopeErr=1
    if slopeErr=1 then
        'ver116-4b deleted irritating error message
        calCanUseAutoWait=0   
        exit sub
    end if
    foundBoundary=0
    for i=highPointNumOfCenterSlope to calMagPoints-1 'ver116-1b
        'find point where slope after that point drops below 3/4 of the center slope
        thisSlope=(calMagTable(i+1,0)-calMagTable(i,0))/(calMagTable(i+1,1)-calMagTable(i,1))   'delta ADC over delta dB
        if thisSlope<0.75*calCenterSlope then highPointNumOfCenterSlope=i : foundBoundary=1 : exit for
    next  i
    if foundBoundary=0 then
        'Slope never dropped off much so we don't need separate high end slope
        calHighADCofCenterSlope=calMagTable(calMagPoints,0) : calHighEndSlope=calCenterSlope
    else
        'highPointNumOfCenterSlope is the point that divides the use of the center slope and use of the high end slope
        calHighADCofCenterSlope=calMagTable(highPointNumOfCenterSlope,0)
        rightADC=calMagTable(calMagPoints,0) : rightDB=calMagTable(calMagPoints,1)  'final ADC and dB entries  'ver116-1b
        'find the high end slope
        for i=calMagPoints-1 to highPointNumOfCenterSlope step -1   'ver116-1b
            thisADC=calMagTable(i,0)
            thisDB=calMagTable(i,1)
            if i=highPointNumOfCenterSlope or (rightDB-thisDB>=7) then
                slopeDenom=rightDB-thisDB
                if slopeDenom=0 then
                    Notice "Error in determining high end slope for Auto Wait."
                    calHighADCofCenterSlope=calMagTable(calMagPoints,0) : calHighEndSlope=calCenterSlope
                    calCanUseAutoWait=0
                else
                    calHighEndSlope=(rightADC-thisADC)/slopeDenom   'Slope of top 7 dB
                end if
            end if
        next i
    end if
    if calCanUseAutoWait=0 then useAutoWait=0   'ver116-4e
    'We now have to find the lowest level at which calCenterSlope can be used. This is the point where it drops off
    'at least 60%
    foundBoundary=0
    for i=highPointNumOfCenterSlope to 2 step -1 'ver116-1b
        'find point where slope after that point drops below 3/4 of the center slope
        'But don't mess with DB values above -60 dB.
        thisADC=calMagTable(i,0)
        thisDB=calMagTable(i,1)
        slopeDenom=thisDB-calMagTable(i-1,1)
        if slopeDenom=0 then thisSlope=0 else thisSlope=(thisADC-calMagTable(i-1,0))/slopeDenom   'delta ADC over delta dB
        if thisSlope<0.60*calCenterSlope then lowPointNumOfCenterSlope=i : foundBoundary=1 : exit for
    next  i
    if foundBoundary=0 then
        'Slope never dropped off much so we don't need separate low end slope
        calLowADCofCenterSlope=0 : calLowEndSlope=calCenterSlope
    else
        'lowPointNumOfCenterSlope is the point that divides the use of the center slope and use of the low end slope
        calLowADCofCenterSlope=calMagTable(lowPointNumOfCenterSlope,0)
            'Now find the low end slope and boundary
        firstSlopePointNum=0
        calLowEndSlope=calCenterSlope    'In case it can't be determined, use center slope value
        for i=1 to calMagPoints 'ver116-1b
            'find point where slope after that is at least 1/10 of the center slope
            thisADC=calMagTable(i,0)
            thisDB=calMagTable(i,1)
            if thisDB>-60 then
                Notice "Error in determining low end slope for Auto Wait."    'Cal is messed up.
                canUseAutoWait=0
                exit for
            end if
            slopeDenom=calMagTable(i+1,1)-thisDB
            if slopeDenom=0 then thisSlope=calCenterSlope else thisSlope=(calMagTable(i+1,0)-thisADC)/slopeDenom   'delta ADC over delta dB
            if thisSlope>=0.07*calCenterSlope then firstSlopePointNum=i : exit for
        next  i
        if firstSlopePointNum=0 then
            Notice "Error in determining low end slope for Auto Wait--all slopes too small"
            canUseAutoWait=0
            calADCofLowFringe=0
        else
            leftADC=calMagTable(firstSlopePointNum,0) : leftDB=calMagTable(firstSlopePointNum,1)
            calADCofLowFringe=leftADC   'below this the slopes are tiny.
            'Find low end slope
            for i=firstSlopePointNum+1 to calMagPoints   'ver116-1b
                thisADC=calMagTable(i,0)
                thisDB=calMagTable(i,1)
                if thisDB>-60 then
                    Notice "Error in determining low end slope for Auto Wait."    'Cal is messed up. Use center slope to very bottom.
                    canUseAutoWait=0
                    exit for
                end if
                if (thisDB-leftDB)>=5 then 'Segment at least 5 dB long
                    calLowEndSlope=(thisADC-leftADC)/(thisDB-leftDB) 'Slope of low end
                    exit for
                end if
            next i
        end if
    end if
    a1=calLowEndSlope
    a2=calCenterSlope
    a3=calHighEndSlope
    a4=calLowADCofCenterSlope
    a5=calHighADCofCenterSlope
end sub

sub calCreateFreqCubicCoeff     'Create table of cubic coefficients for mag cal table
    'Each entry of the freq coefficient table will have 4 numbers. 0-3 are the A,B,C,D
    'coefficients for interpolating the frequency power correction, which is a scalar.
    call intSetMaxNumPoints calFreqPoints  'Be sure we have room ver114-5q
    call intClearSrc
    for i=1 to calFreqPoints 'copy cal table to intSrc
        call intAddSrcEntry calFreqTable(i,0),calFreqTable(i,1),0
    next i
        'Signal that this is not an angle, and  to "favor flat", because
        'we expect the freq response not to become near-vertical
    favorFlat=1 : isAngle=0
    call intCreateCubicCoeffTable 1,0,isAngle, favorFlat, 0 'Final 0 means we are not doing phase correction ver116-1b
    for i=1 to calFreqPoints  'put the data where we want it
        calFreqCoeffTable(i,0)=intSrcCoeff(i,0) :calFreqCoeffTable(i,1)=intSrcCoeff(i,1)
        calFreqCoeffTable(i,2)=intSrcCoeff(i,2) :calFreqCoeffTable(i,3)=intSrcCoeff(i,3)
    next i
end sub

function calBinarySearch(dataType, searchVal)
'Perform search to find the lowest entry whose lookup value is >=searchVal
'Entries are numbered 1...
'If dataType=0 then the lookup is for ADC value in calMagTable; otherwise it is for freq in in calFreqTable
'Uses primarily a binary search with recursive calls.
'If searchVal is beyond the highest entry, we will return top+1
    if dataType=0 then 'ver115-9b
        nPoints=calMagPoints
        thisVal=calMagTable(1,0)
        if searchVal<=thisVal then calBinarySearch=1 : exit function    'off bottom
        thisVal=calMagTable(nPoints,0)
        if searchVal>thisVal then calBinarySearch=nPoints+1 : exit function    'off top
    else
        nPoints=calFreqPoints
        thisVal=calFreqTable(1,0)
        if searchVal<=thisVal then calBinarySearch=1 : exit function    'off bottom
        thisVal=calFreqTable(nPoints,0)
        if searchVal>thisVal then calBinarySearch=nPoints+1 : exit function    'off top
    end if
    'Here we know searchVal is >first entry and <=final entry
    top=nPoints
    bot=1
    span=top-bot+1
    while span>4  'Do preliminary search to narrow the search area to no more than 4 entries
        halfSpan=int(span/2)
        mid=bot+halfSpan
        if dataType=0 then thisVal=calMagTable(mid,0) else thisVal=calFreqTable(mid,0)
        if thisVal=searchVal then calBinarySearch=mid : exit function   'exact hit
        if thisVal<searchVal then bot=mid+1 else top=mid  'Narrow to whichever half thisVal is in
        span=top-bot+1  'Repeat with either bot or top endpoints changed
    wend
    'Here searchVal is > entry bot, and <= entry top, and there are at most 4 entries to search, starting with bot
    'Start with bot entry and find first entry >= searchVal. Even if there are fewer than 4 entries left, it
    'is guaranteed we will not keep searching past the top of the table.
    ceil=bot
    if dataType=0 then thisVal=calMagTable(ceil,0) else thisVal=calFreqTable(ceil,0)
    if thisVal>=searchVal then calBinarySearch=ceil : exit function 'compare first of possible 4
    ceil=ceil+1     'ver115-9b deleted check for ceil>nPoints, which is dealt with above
    if dataType=0 then thisVal=calMagTable(ceil,0) else thisVal=calFreqTable(ceil,0)
    if thisVal>=searchVal then calBinarySearch=ceil : exit function 'compare second of possible 4
    ceil=ceil+1
    if dataType=0 then thisVal=calMagTable(ceil,0) else thisVal=calFreqTable(ceil,0)
    if thisVal>=searchVal then calBinarySearch=ceil : exit function 'compare third of possible 4
    calBinarySearch=ceil+1     'return index of fourth, which is the only remaining possibility.
end function

sub calConvertMagPhase magdata, wantPhase, byref magDB, byref phaseCorrect
    'This function returns the magnitude (magDB) and the phase correction (phaseCorrect)
    'based on the ADC signal strength reading magdata. The phase correction is returned
    'only if wantPhase=1; otherwise zero is returned.

    'It uses cubic interpolation on the calibration entries of calMagTable.
    'Note: calMagTable is organized with the smallest ADC values first
        'SEW8 revised call to calBinarySearch
    index=calBinarySearch(0,magdata)   'search calMagTable
    'index now is the first entry >= magdata, except that if no entry meets that test,
    'index will be one past the end.
    phaseCorrect=0
    if index>calMagPoints then
        'Off top end;use largest mag and phase correction for largest ADC entry
        magDB=calMagTable(calMagPoints,1)
        if wantPhase then phaseCorrect=calMagTable(calMagPoints,2)  'ver115-1a
        exit sub
    end if
    if index=1 then
        'Off bottom end;use mag and phase correction for smallest ADC entry
        magDB=calMagTable(1,1)
        if wantPhase then phaseCorrect=calMagTable(1,2) 'ver115-1a
        exit sub
    end if

    'Evaluate cubic at x=magdata
    dif=magdata-calMagTable(index,0)
    A=calMagCoeffTable(index,0) : B=calMagCoeffTable(index,1)
    C=calMagCoeffTable(index,2) : D=calMagCoeffTable(index,3)
    magDB = A+dif*(B+dif*(C+dif*D))

    if wantPhase=1 then
        A=calMagCoeffTable(index,4) : B=calMagCoeffTable(index,5)
        C=calMagCoeffTable(index,6) : D=calMagCoeffTable(index,7)
        phaseCorrect = A+dif*(B+dif*(C+dif*D))
        if phaseCorrect<=-180 then phaseCorrect=phaseCorrect+360 'ver113-7e
        if phaseCorrect>180 then phaseCorrect=phaseCorrect-360 'ver113-7e
    end if
end sub

function calConvertFreqError(freq)
    'Find power correction for frequency error contribution.
    'needed: datatable, frequency error calibration table; creates average freqerror

    'This routine uses cubic interpolation on calFreqTable to determine the frequency
    'error for thisfreq. The calibration table consists of entries numbered from 1, each
    'containing a frequency and the error at that frequency. The entries must be in order of
    'increasing, with the first entry being the lowest frequency. If the frequency is less than the
    'first entry, or greater than the final entry,the error for that entry will be used.
    'The final entry need not be for the maximum possible frequency.
            'SEW8 revised call to calBinarySearch
    index=calBinarySearch(1,freq)    'Search calFreqTable
    'index now is the first entry >= freq, except that if no entry meets that test,
    'index will be one past the end.
    if index>calFreqPoints then _
            calConvertFreqError=calFreqTable(calFreqPoints,1): exit function  'use largest value
    if index=1 then _
            calConvertFreqError=calFreqTable(1,1): exit function    'Use smallest value

        '(5)Evaluate cubic at x=magdata
    dif=freq-calFreqTable(index,0)
    A=calFreqCoeffTable(index,0) : B=calFreqCoeffTable(index,1)
    C=calFreqCoeffTable(index,2) : D=calFreqCoeffTable(index,3)
    calConvertFreqError = A+dif*(B+dif*(C+dif*D))
end function

sub calCreateDefaults pathNum, editor$, doPoints 'Create default data in file or text editor.
    'if editor$<>"", it contains the destination text editor; otherwise we do a file,
    'which replaces any existing file. For a file, we assume any necessary folders exist.
    'If pathNum=0 we create a freq cal file; othwerwise a mag cal file for path pathNum
    'if doPoints=1 we create two data points; otherwise no points
    if editor$="" then
        doFile=1
        open calFilePath$()+calFileName$(pathNum) for output as #calOut
        editor$="#calOut"
    else
        doFile=0
        #editor$, "!cls"    'Clear existing data in text editor
    end if

        'Start with comments
    if pathNum=0 then
            'Freq
        c1$="Calibration over frequency"
        c2$="Calibrated " + Date$("mm/dd/yy") + " at -zz dBm." 'ver116-1b
    else
            'Mag
        c1$="Path"+str$(pathNum)+" CenterFreq= xxx; Bandwidth=yyy"
        c2$="Calibrated " + Date$("mm/dd/yy") + " at zz MHz."
    end if
    calFileComments$(1)=c1$  : calFileComments$(2)=c2$   'Save comments in our array
    print #editor$, calFileCommentChar$+c1$
    print #editor$, calFileCommentChar$+c2$   'print comments
    print #editor$, "CalVersion=";using("##.##", calModuleVersion)
    if pathNum=0 then
        print #editor$, "FreqTable="
        print #editor$, calFileCommentChar$+"    MHz        dB   in increasing order of MHz"    'ver116-4L
        if doPoints=1 then
            print #editor$, "0           0"
            print #editor$, "1000        0"
        end if
    else
        print #editor$, "MagTable="
        print #editor$, calFileCommentChar$+"  ADC      dBm  in increasing order of ADC"    'ver116-1b
        if calDoPhase=1 then p$="      0" else p$=""
        if doPoints=1 then
            'ver115_1a changed defaults to be based on ADC type  ver115-1b made max half of maxbits
            'AtoD topology."8" for original 8 bit,"12" for optional 12 bit ladder,"16" for serial 16 bit AtoD, or "22" for serial 12 bit AtoD
            maxADC=int(65535/2) : maxADCval=0          '16 bit ADC ver115-4e
            if adconv=8 then maxADC=int(255/2)            '8 bit ADC
            if adconv=12 or adconv=22 then maxADC=4095 : maxADCval=10   '12 bit ADC ver115-4e
            print #editor$, "0        -120"
            print #editor$, maxADC; "      ";maxADCval 'ver115-4e
        end if
    end if
    'For now, we don't use the EndCalFile tag, and don't require it on input either
    'print #editor$, "EndCalFile="
    if doFile then close #calOut
end sub

function calAvailableFiles$()    'Return string indicating which cal files exist
    'Returned string will have only as many characters as needed to flag existing files.
    'The first is for the frequency cal file.
    'The Nth character, for N=1 to 41, relates to the (N-1)th mag cal file. If the file exists,
    'the character is 1; otherwise the character is 0.
    p$=calFilePath$()
    sLen=0  'This will be length needed to handle all the files we have
    files p$, "MSA_CalFreq.txt", calFileInfo$()
    if calFileInfo$(0,0)<>"0" then
        s$="1"  'enter 1 if freq file exists; else 0
        sLen=1
    else
        s$="0"
    end if

    files p$, "MSA_CalPath*.txt", calFileInfo$()    'search for mag cal files
    numFiles=val(calFileInfo$(0,0))
    for j=1 to 40   'Test for 40 possible file numbers
        j$=str$(j)+".txt"
        thisFlag$="0"
        for i=1 to numFiles
            n$=calFileInfo$(i,0)
            n$=Mid$(n$,12)   'get name without "MSA_CalPath"
            if j$=n$ then thisFlag$="1": sLen=j+1: exit for
        next i
        s$=s$+thisFlag$     'append a 1 if this file exists; 0 otherwise
    next j
        'We have determined that we need no more than sLen characters; all
        'the extras would be 0.
    calAvailableFiles$=Left$(s$,sLen)
end function

sub calInstallFile pathNum  'Open cal file and load data. Create folder and file if necessary
    'If pathNum=0 we want a freq cal file; otherwise a mag cal file for path pathNum
    'Find out whether the necessary folders exist for config and calibration
    'information. If no proper cal file exists, then create one with
    'default values, but ask user before replacing an existing file
    'SEWcal Also calculates cubic interpolation coefficients for the newly loaded file

    'Create the name of the desired file
    fileName$=calFileName$(pathNum)
    'See if we have the folder MSA_Info
    files DefaultDir$, "", calFileInfo$()
    numFolders=val(calFileInfo$(0,1))
    haveFolder=0
    for i=1 to numFolders
        if calFileInfo$(i,1)="MSA_Info" then haveFolder=1: exit for
    next  i
    'Create folders and file if necessary, or just open file
    if haveFolder=0 then
            'Create MSA_Info folder
        if 0<>mkDir("MSA_Info") then notice "Cannot access files."
    end if
        'We have the MSA_Info folder. See if we have MSA_Cal folder
    files DefaultDir$+"\MSA_Info", "", calFileInfo$()
    numFolders=val(calFileInfo$(0,1))
    haveFolder=0
    for i=1 to numFolders
        if calFileInfo$(i,1)="MSA_Cal" then haveFolder=1: exit for
    next  i
    if haveFolder=0 then
            'Create MSA_Cal folder
        if 0<>mkDir("MSA_Info\MSA_Cal") then notice "Cannot access files."
    end if
        'We have the MSA_Info\MSA_CAL folder. See if it has proper file
    files calFilePath$(), fileName$, calFileInfo$()
    if calFileInfo$(0,0)="0" then
        call calCreateDefaults pathNum, "",1  'No file. Create one.
    end if

        'Get data from file
    fErr$=calLoadFromFile$(pathNum)
    if fErr$<>""then
        'File is there, but has a problem
        Confirm "File error: "+fErr$+"; Replace with default file?"; response$
        if response$="no" then
            notice "File not replaced. Internal cal data cleared."
            call calClearComments
            if pathNum=0 then call calClearFreqPoints else call calClearMagPoints
            exit sub
        end if
        call calCreateDefaults pathNum, "", 1
            'Try once more. Any error now is a system error
        if ""<>calLoadFromFile$(pathNum) then notice "Cannot access file" : exit sub
    end if

    if calFileVersion>calfigModuleVersion then _
                notice "Warning: Calibration file format is later version than the sofware"
    'Here we should update the file if calFileVersion is old. But at this
    'time we have no old file versions

            'SEWcal Calculate interpolation coefficients for this file, and for path file determine validPhaseThreshold
    if pathNum=0 then call calCreateFreqCubicCoeff else call calCreateMagCubicCoeff
end sub

sub calSaveToFile pathNum   'Save to file
    'Save a cal file for path number pathNum. We assume the necessary
    'folders are already in place
    'If pathNum=0 we are dealing with a freq cal file.
    'File output replaces any existing file of the same name.
    'calFileComments$() are placed at the beginning of the file, marked
    'with calFileCommentChar$

    calFile$=calFilePath$() + calFileName$(pathNum)
    open calFile$ for output as #outFile
    startLine=1
    call calOutputData "#outFile", 1 ,pathNum,startLine
    close #outFile
end sub

sub calSaveToEditor editor$, pathNum   'Save to text editor
    'Save variables and point array into text editor, in same format as a file
    'calFileComments$() are placed at the beginning of the data, marked
    'with calFileCommentChar$
    startLine=1
    call calOutputData editor$, 0 ,pathNum,startLine
    #editor$, "!select 1,1"
    #editor$, "!copy" : #editor$, "!paste"  'This scrolls the window to the top
end sub

sub calOutputData calFile$, isFile, pathNum, byRef startLine     'Output data to file or text editor
    'Save data as a frequency cal file. Replaces any existing file.
    'The data is sent to  a file already opened, whose handle is in calFile$
    'or from a text editor whose handle is in calFile$. The type of source is
    'indicated by isFile (1 for file, 0 for textEditor).
    'calFileComments$() are placed at the beginning of the file, marked
    'with calFileCommentChar$
    'We update startLine to one past the last line we write

    'Start with comments
    for i=1 to 2
        com$=calFileComments$(i)
        if com$<>"" then print #calFile$, calFileCommentChar$+com$
    next i

    print #calFile$, "CalVersion=";using("##.##", calModuleVersion)
    if pathNum=0 then
        print #calFile$, "FreqTable="
        print #calFile$, calFileCommentChar$+"    MHz        db   in increasing order of MHz"
        for i=1 to calFreqPoints
            f=calFreqTable(i,0)
            f$=using ("####.######", calFreqTable(i,0))
            v$=using ("####.###", calFreqTable(i,1))  'ver115-1e
            print #calFile$, f$;"  ";v$  'output data line
        next i
    else
        print #calFile$, "MagTable="
        if calDoPhase=0 then
            print #calFile$, calFileCommentChar$+"  ADC      dBm   in increasing order of ADC"  'ver116-1b
        else
            print #calFile$, calFileCommentChar$+"  ADC      dBm      Phase   in increasing order of ADC"   'ver116-1b
        end if
        for i=1 to calMagPoints
            adc$=using ("#######", calMagTable(i,0))
            v$=using ("####.###", calMagTable(i,1))  'ver115-1e
            p$=using ("####.##", calMagTable(i,2))
            if calDoPhase=0 then
                print #calFile$, adc$;"  ";v$  'output data line
            else
                print #calFile$, adc$;"  ";v$;"  ";p$ 'output data line
            end if
        next i
    end if
    'For now, we don't print the end tag, nor do we require it on input
    'print #calFile$, "EndCalFile="
end sub

function calOpenFile$(pathNum)
    'Open calibration file for path number pathNum; return its handle
    'if pathNum=0 then we want freq cal file
    'If file does not exist, return "".
    fName$=calFileName$(pathNum)
    On Error goto [noFile]
    open fName$ for input as #calFile
    calOpenFile$="#calFile"
    exit function
[noFile]
    fName$=""
    calOpenFile$="" 'ver114-2f
end function

function calFileName$(pathNum)   'Return file name for the specified pathNum
    'If pathNum=0 we want the freq cal file; otherwise mag cal for path number pathNum
     if pathNum=0 then
        calFileName$="MSA_CalFreq.txt"
    else
        calFileName$="MSA_CalPath" + str$(pathNum) + ".txt"
    end if
end function

function calFilePath$()   'Return path name for cal files, ending in "\"
     calFilePath$=DefaultDir$ + "\MSA_Info\MSA_Cal\"
end function

function calLoadFromEditor$(editor$,pathNum) 'Read data from text editor. Return error message
    'editor$ is the handle of the text editor. We also sort.
    startLine=1 : isFile=0
    fErr$=calReadFile$(editor$,isFile,pathNum,startLine)
    if pathNum=0 then call calSortFreq else call calSortMag
    calLoadFromEditor$=fErr$
    if pathNum<>0 then finalfreq=MSAFilters(pathNum, 0) : finalbw=MSAFilters(pathNum, 1)   'SEW7 made conditional
end function

function calLoadFromFile$(pathNum) 'Open and read file. Return error message
    'This is called only after we know the file exists
    fileName$=calFilePath$() + calFileName$(pathNum)
    open fileName$ for input as #calIn
    startLine=1 : isFile=1
    calLoadFromFile$=calReadFile$("#calIn",isFile,pathNum,startLine)
    close #calIn
    if pathNum<>0 then finalfreq=MSAFilters(pathNum, 0) : finalbw=MSAFilters(pathNum, 1)   'SEW7 made conditional
end function

function calReadFile$(calFile$, isFile, pathNum, byRef startLine)
    'Reads calibration data, starting at line startLine and continuing until
    'we find the tag "EndCalFile=" (case insensitive). We update startLine
    'to one past the line containing that tag.
    'If pathNum=1... we read magnitude calibration data for a
    'particular signal path. There is one path for each RBW filter. if pathNum=0 then
    'we read frequency calibration data, for which there is only one data set.
    'The data is read from a file already opened, whose handle is in calFile$
    'or from a text editor whose handle is in calFile$. The type of source is
    'indicated by isFile (1 for file, 0 for textEditor).
    'The data is in the following format; comments are allowed and are marked
    'by calFileCommentChar$. First two comment lines preceding tags are put into calFileComments$
    'Each data tag ends with "=", and is case insensitive.
    '---Magnitude calibration file format---
    'VERSION=xxx    'Version of the module that wrote the file
    'MagTable=      'Start of magnitude cal data; no data on this line
    'xxxx xxxx      'Up to calMaxMagPoints lines of data pairs: ADC reading(increasing) and mag(db)
    'xxxx xxxx      'Table is put into calMagTable, with ADC reading as x value
    'ENDCALFILE=    'marks end of data
    '
    '---Frequency calibration file format---
    'VERSION=xxx    'Version of the module that wrote the file
    'FreqTable=      'Start of frequency cal data; no data on this line
    'xxxx xxxx      'Up to to calMaxPoints lines of data pairs:frequency(MHz)(increasing) and magitude(db)
    'xxxx xxxx      'Table is put into calFreqTable with frequency as x value
    'ENDCALFILE=    'marks end of data
    '---end of file formats
    '
    'The number of entries in calMagTable is put into calMagPoints
    'The number of entries in calFreqTable is put into calFreqPoints
    'If error, we return a string describing the error
    'If no error, we return ""
    doMag=0     '=1 while we are processing mag cal data
    doFreq=0     '=1 while we are processing freq cal data
    if pathNum>0 then     'Reset number of points
        call calClearMagPoints
    else
        call calClearFreqPoints
    end if
    ncalFileComments=0:accumComments=1
    calFileComments$(1)="": calFileComments$(2)=""
    if isFile=0 then #calFile$, "!lines nLines" else nLines=100000
    prevXVal=-100000      'SEW7; Used to check for increasing values
    p$="Path "+str$(pathNum)+": "   'For error messages
    for fileLine=startLine to nLines
            'Loop until we reach nLines (if text editor) or end of file (if file)
        if isFile then
            if EOF(#calFile$)<>0 then exit for
            Line Input #calFile$, tLine$   'Read one line
        else
            print #calFile$, "!line ";fileLine;" tLine$" 'Get line number fileLine
        end if
        startLine=startLine+1   'one past where we are
        tLine$=Trim$(tLine$)    'drop extra blanks
        startChar$=Left$(tLine$,1)  'first real character of line
        if startChar$=calFileCommentChar$ then isComment=1 else isComment=0
        if isComment or startChar$="" then
            'Save first two comment lines, w/0 the comment character
            if accumComments and isComment and ncalFileComments<2 then
                ncalFileComments=ncalFileComments+1
                calFileComments$(ncalFileComments)=Mid$(tLine$,2)
            end if
        else    'valid non-comment line
            accumComments=0     'no more comments put into calFileComments$$
            equalPos=instr(tLine$,"=")     'equal sign marks end of tag
            if equalPos<>0 then 'Equal sign found
                tag$=Upper$(Left$(tLine$, equalPos-1))
                item$=Mid$(tLine$, equalPos+1)  'stuff after equal sign
                commentPos=instr(item$, calFileCommentChar$)
                if commentPos>0 then item$=Left$(item$, commentPos-1)   'drop comments
                item$=Trim$(item$)
                doMag=0: doFreq=0:isErr=0
                select case tag$
                 case "CALVERSION"
                    if val(item$)>calModuleVersion then _
                            calReadFile$="Warning: "+p$+"File format is later version than the sofware"
                    case "MAGTABLE"
                        if pathNum=0 then isErr=1 else doMag=1
                    case "FREQTABLE"
                        if pathNum>0 then isErr=1 else doFreq=1
                    case "ENDCALFILE"
                        'This marks the end of the file
                        'SEW7 deleted check to see that data points existed
                        calReadFile$=""   'no error
                        exit function
                    case else
                        isErr=true
                end select
                if isErr then calReadFile$=p$+"Line "+str$(fileLine):exit function 'invalid tag
            else    'No equal sign; must be data
                'Now retrieve the numbers from this data line. There will always
                'be at least two numbers; for mag calibration of VNA there may also be phase.
                 'Numbers are separated by comma, tab or space
                delims$=" ," + chr$(9)    'space, comma and tab are delimiters
                isErr=uExtractNumericItems(2, tLine$, delims$,data1, data2, data3)
                    'If not numeric, signal error
                if isErr=1 then calReadFile$=p$+"Line "+str$(fileLine): exit function
                if doMag then
                    calMagDataHadPhase=0
                    if tLine$<>"" then  'If there is data left, we must have a third number
                        isErr=uExtractNumericItems(1, tLine$, delims$,data3, data4, data5)
                        if isErr=1 then calReadFile$=p$+"Line "+str$(fileLine): exit function
                        calMagDataHadPhase=1
                    else
                        data3=0
                    end if
                end if
                'SEWcal2 modified the next approx 17 lines to eliminate
                'duplicate entries in some cases. ver113-7g
                skipData=0
                if data1<=prevXVal then 'SEWcal2 created this if...else block
                    'Data in file must have increasing x values; in text editor we let it slide
                    'If we have absolutely identical entries, we skip the duplicates.
                    if data1=prevXVal and data2=prevY1 and data3=prevY2 then
                        skipData=1
                    else
                        if isFile=1 then calReadFile$=p$+"Line "+str$(fileLine): exit function 'not increasing
                    end if
                end if
                prevXVal=data1 : prevY1=data2 : prevY2=data3 'SEWcal2
                tooMany=0
                if doMag then   'Add mag point; quit if error (i.e. too many points)
                    if skipData=0 then tooMany=calAddMagPoint(data1, data2, data3)   'SEWcal2
                else    'Add frequency point
                    if skipData=0 then tooMany=calAddFreqPoint(data1, data2)   'SEWcal2
                end if
                if tooMany=1 then calReadFile$=p$+"Line "+str$(fileLine)+"; Too many points.": exit function
            end if
        end if  'end of processing non-comment line
    next fileLine

    'We can only get here if we did not find the EndCalFile tag, which is really
    'only needed if this file is embedded in another. So for now we let it go.
    'calReadFile$="Proper end of data not found."
    calReadFile$=""
end function
'@=================End Mag/Freq  Calibration Module===================
'@=================End Calibration Manager Module===================
'

'
'--SEW Added the following module to provide utility routines used by other modules
'

'=====================Start Complex Functions Module====================

sub cxRectToPolarRad R,I, byref m, byref ang  'Convert rectangular coordinates to polar(rad)
    'R+jI to mag m at angle ang (radians, -pi to +pi)
   ang=uATan2(R, I)*0.0174532925199433
   m=sqr(R^2+I^2)   'mag
end sub

sub cxPolarRadToRect  m,ang, byref R, byref I   'Convert polar(rad) coordinates to rectangular
    'mag m at angle ang (radians) to R+jI
    if m<0 then m=0-m : ang=ang+uPi()   'Negative mag; point in opposite direction
    R=m*cos(ang)   'Trig is done in radians
    I=m*sin(ang)
end sub

sub cxInvert R, I, byref Rres, byref Ires     'Invert complex number R + jI; put into Rres, Ires
    '1/(R+jI)=(R-jI)/(R^2+I^2)
    D=R^2+I^2
    if D=0 then Rres=constMaxValue : Ires=0 : exit sub
    Rres=R/D : Ires=0-I/D
end sub

sub cxDivide Rnum, Inum, Rden, Iden, byref Rres, byref Ires     'Divide (Rnum + jInum)/(Rden + jIden); put into Rres, Ires
    if Rnum=0 and Inum=0 then Rres=0: Ires=0 : exit sub   '0 numerator means zero result; we do this even if denominator=0
    on error goto [MathErr] 'ver115-4d
    'First invert the denominator
    D=Rden^2+Iden^2
    if D=0 then Rres=constMaxValue : Ires=0: exit sub
    Rinv=Rden/D : Iinv=0-Iden/D
    'Now multiply Rnum+jInum times Rinv+jIinv
    Rres=Rnum*Rinv-Inum*Iinv
    Ires=Rnum*Iinv+Inum*Rinv
    exit sub
[MathErr]
    notice "Division Error"
    Rres=constMaxValue : Ires=0
end sub

'ver115-4e renamed cxMultiply
sub cxMultiply R1, I1, R2, I2, byref Rres, byref Ires     'Multiply (R1 + jI1)*(R2 + jI2); put into Rres, Ires
    on error goto [MathErr]  'ver115-4d
    Rres=R1*R2-I1*I2
    Ires=R1*I2+I1*R2
    exit sub
[MathErr]
    notice "Multiplication Error"
    Rres=constMaxValue : Ires=0
end sub

sub cxSqrt R,I, byref resR, byref resI 'Square root of R+jI is resR+jresI (non-neg real part)
    'Square root of X; branch cut on negative X-axis
    magR = abs(R):magI = abs(I)
    if magR=0 and magI=0 then resR=0 : resI=0 : exit sub  'square root of zero equals zero
    If (magR >= magI) Then
        t = magI / magR
        w = sqr(magR) * sqr(0.5 * (1 + sqr(1 + t * t)))
    Else
        t = magR / magI
        w = sqr(magI) * sqr(0.5 * (t + sqr(1 + t * t)))
    End If

    if w=0 then resR=0 : resI=0 : exit sub

    if (R >= 0) then
        resR=w
        resI=I / (w + w)
    else
        'Note that resR will always be non-negative, since w is non-negative
        resR = magI/ (w + w)
        if I >= 0 then
            resI = w
        else
            resI = 0-w
        end if
    end if
end sub

sub cxNatLog R,I, byref resR, byref resI 'Nat log of R+jI is resR+jresI
    call cxRectToPolarRad R,I, m, ang  'Convert to mag, radian format
        'Take the log of the magnitude and carry over the same angle (radians) as imaginary part
    resR=log(m)
    resI=ang
end sub

sub cxEPower R,I, byref resR, byref resI 'e to powere of R+jI is resR+jresI
    if R=0 and I=0 then resR=1 : resI=0 : exit sub  'e^0 is 1
    ex=exp(R)
    if I=0 then resR=ex : resI=0 : exit sub    'For real X, e^X is simple
    resR=ex*cos(I) : resI=ex*sin(I)      'e^x=e^R*(cos(I)+j*sin(I))
end sub

sub cxCos R,I, byref resR, byref resI 'Cosine of X=R+jI is resR+jresI
    '[e^(X*j)+ e^(-X*j)]/2
    call cxEPower 0-I,R,R1,I1    'e^(X*j)
    call cxEPower I,0-R,R2,I2    'e^(-X*j)
    resR=(R1+R2)/2
    resI=(I1+I2)/2
end sub

sub cxSin R,I, byref resR, byref resI 'Sine of X=R+jI is resR+jresI
    '[e^(X*j)- e^(-X*j)]/(2*j)  Note 1/(2*j)=-j/2
    call cxEPower 0-I,R,R1,I1    'e^(X*j)
    call cxEPower I,0-R,R2,I2    'e^(-X*j)
    resR=(I1-I2)/2      '(a+jb)/(2j)=-j(a+jb)/2=(b-ja)/2
    resI=(R2-R1)/2
end sub

sub cxTan R,I, byref resR, byref resI 'Tan of X=R+jI is resR+jresI
    'sin(X)/cos(X)
    if R=0 and I=0 then resR=0 : resI=0 :  exit sub 'tan(0) is 0
    call cxSin R,I,sR,sI      'sin(X)
    call cxCos R,I,cR,cI      'cos(X)
    call cxDivide sR,sI,cR,cI,resR,resI  'sin(X)/cos(X)
end sub

sub cxCot R,I, byref resR, byref resI 'Cotan of X=R+jI is resR+jresI
    'cot(X)=tan(pi/2-X)
    halfPi=uPi()/2
    call cxTan halfPi-R, 0-I, resR, resI  'Tan(pi/2-X)
end sub

sub cxASin R,I, byref resR, byref resI 'Arcsine of X=R+jI is resR+jresI
    'arcsin(x)=-j*log(jX+sqrt(1-X^2))
    call cxMultiply R,I,0-R,0-I,xsqR, xsqI    '-X^2
    xsqR=xsqR+1         '1-X^2
    call cxSqrt xsqR,xsqI, sqrtR, sqrtI           'sqrt(1-X^2)
    call cxNatLog sqrtR-I, sqrtI+R, logR, logI    'log[jX+sqrt(1-X^2)]
    resR=logI            '-j*log[jX+sqrt(1-X^2)]
    resI=0-logR
 end sub

sub cxACos R,I, byref resR, byref resI 'Arccos of X=R+jI is resR+jresI
    'arccos(x)=-j*log(X+sqrt(X^2-1))=pi/2 - arcsin(X)
    call cxASin R,I,aR,aI                    'Arcsin(X)
    halfPi=uPi()/2
    resR=halfPi-aR           'pi/2 - arcsin(X)
    resI=0-aI
end sub

sub cxCosh R,I, byref resR, byref resI 'Cosh of X=R+jI is resR+jresI
        'hyperbolic cosine of X
        ' Cosh(z) = (e^z + e^(-z))/2 = cos(j*z)
    call cxCos 0-I, R,resR, resI          'cos(j*X)
end sub

sub cxSinh R,I, byref resR, byref resI 'Sinh of X=R+jI is resR+jresI
        'hyperbolic sine of X
        ' Sinh(z) = (e^z - e^(-z))/2 = -j*sin(j*z)
    call cxSin 0-I, R, sR, sI            'sin(j*X)
    resR=sI   '-j*sin(j*x)
    resI=0-sR
end sub

sub cxTanh R,I, byref resR, byref resI 'Tanh of X=R+jI is resR+jresI
 'hyperbolic tangent of X
       ' Tanh(z) = (sinh(z)/cosh(z))
    call cxSinh R,I,sR,sI        'sinh(X)
    call cxCosh R,I,cR,cI        'cosh(X)
    call cxDivide sR, sI, cR, cI, resR, resI    'sinh(X)/cosh(X)
end sub

'=====================Start Complex Functions Module====================

