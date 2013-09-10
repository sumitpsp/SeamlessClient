//
//  RecentItemsDataController.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/9/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "RecentItemsDataController.h"

@interface RecentItemsDataController ()
-(void) initializeDefaultRecentItemsList;
@end
@implementation RecentItemsDataController

-(void) initializeDefaultRecentItemsList {
    NSMutableArray* recentItemsList = [[NSMutableArray alloc] init];
    self.recentMasterItemsList = recentItemsList;
}

-(void) setRecentMasterItemsList:(NSMutableArray *)newList {
    if (_recentMasterItemsList != newList) {
        _recentMasterItemsList = [newList mutableCopy];
    }
}

- (id)init {
    if (self = [super init]) {
        [self initializeDefaultRecentItemsList];
        return self;
    }
    return nil;
}

-(NSUInteger) countOfRecentMasterItemsList {
    return [self.recentMasterItemsList count];
}

-(RecentItem*) objectInListAtIndex:(NSUInteger)theIndex {
    return [self.recentMasterItemsList objectAtIndex:theIndex];
}

- (void)addItem:(RecentItem *)item {
    [self.recentMasterItemsList addObject:item];
}

@end


