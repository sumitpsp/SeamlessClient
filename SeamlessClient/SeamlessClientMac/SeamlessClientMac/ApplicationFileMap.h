//
//  ApplicationFileMap.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/15/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>
#include "RecentItem.h"

@interface ApplicationFileMap : NSObject

@property (strong, nonatomic) NSRunningApplication* application;
@property (strong, nonatomic) RecentItem* recentItem;

@end
