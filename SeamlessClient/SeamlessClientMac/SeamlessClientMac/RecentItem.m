//
//  RecentItem.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/9/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "RecentItem.h"

@implementation RecentItem

-(id) initWithName:(NSString*)newName andPath:(NSString*) newPath {
    if(self = [super init]) {
        self.name = newName;
        self.path = newPath;
        return self;
    }
    return nil;
}
- (id)init {
    if (self = [super init]) {
        return self;
    }
    return nil;
}

@end
