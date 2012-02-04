//
//  ModuleDelegate.h
//  Spectrum Analyzer
//
//  Created by William Dillon on 2/3/12.
//  Copyright (c) 2012. All rights reserved.
//

#import <Foundation/Foundation.h>

@protocol ModuleDelegate <NSObject>

@optional
- (void)moduleDidBecomeStale:(id)sender;
- (void)moduleWillUpdateHardware:(id)sender;

@end
