//
//  OpenApplicationWindows.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/9/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "OpenApplicationWindows.h"

@implementation OpenApplicationWindows

+(NSMutableArray*) getOpenWindows {
    NSLog(@"Secong row of titles");
    CFArrayRef windowList = CGWindowListCopyWindowInfo(kCGWindowListOptionOnScreenOnly | kCGWindowListExcludeDesktopElements, kCGNullWindowID);
    NSMutableArray *data = [(__bridge NSArray *) windowList mutableCopy];
    
    NSMutableArray *filteredData = [[NSMutableArray alloc] initWithCapacity:10];
    
    for (NSMutableDictionary *theDict in data) {
        id layer = [theDict objectForKey:(id)kCGWindowLayer];
        
        if ([layer intValue] == 0) {
            [filteredData addObject:theDict];
        }
    }
    
    NSLog(@"%@", filteredData);
    
    return filteredData;
    // NSLog(@"window: %@", filteredData);
}


@end
