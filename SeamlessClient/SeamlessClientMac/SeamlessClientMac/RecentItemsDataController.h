//
//  RecentItemsDataController.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/9/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "RecentItem.h"

@interface RecentItemsDataController : NSObject
@property (strong, nonatomic) NSMutableArray* recentMasterItemsList;
- (NSUInteger) countOfRecentMasterItemsList;
- (RecentItem*) objectInListAtIndex:(NSUInteger)theIndex;
- (void)addItem:(RecentItem*) item;
@end
