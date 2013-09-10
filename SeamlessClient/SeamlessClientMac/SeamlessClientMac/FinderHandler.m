//
//  FinderHandler.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 8/13/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "FinderHandler.h"
#import <ScriptingBridge/ScriptingBridge.h>
#import "CGSPrivate.h"

@implementation FinderHandler

-(id) init {
    self.finder = [SBApplication applicationWithBundleIdentifier:@"com.apple.Finder"];
    return self;
}

-(NSString*) getRecentlyChanged {
    NSURL *URL = [NSURL fileURLWithPath:[@"~/Desktop" stringByStandardizingPath]];
    if (URL) {
        finderFolder *folder = [[self.finder folders] objectAtLocation:URL];
        NSLog(@"folder == %@", folder);
    }
    
    SBElementArray *trashItems = [[self.finder trash] items];
    if ([trashItems count] > 0) {
        NSLog(@"Count is %d", [trashItems count]);
        for (finderItem *item in trashItems) {
            if ([item locked]==YES)
                [item setLocked:NO];
        }
    }
    


    return [self getCurrentApplication];

    //[ws openFile:@"/Users/sumitpasupalak/report.pdf"];
    //NSLog (@"%@", apps);
}

-(NSString*) getCurrentApplication {
    NSWorkspace * ws = [NSWorkspace sharedWorkspace];
    NSArray * apps = [ws runningApplications];
    
    NSRunningApplication* currentApp = [ws frontmostApplication];
    
  //  NSLog(@"Apps are %@", apps);
    for (NSRunningApplication* app in apps) {
        
    }
    
    return [currentApp localizedName];
    
}

@end
