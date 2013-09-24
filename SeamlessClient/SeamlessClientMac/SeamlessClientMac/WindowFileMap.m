//
//  WindowFileMap.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/15/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "WindowFileMap.h"

@implementation WindowFileMap

-(id)init{
    
    if((self=[super init])){
        self.recentItem = [[RecentItem alloc] init];
        self.windowInfo = [[NSMutableDictionary alloc] init];
    }
    return self;
}


@end
