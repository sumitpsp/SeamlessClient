//
//  WindowFileMap.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/15/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "RecentItem.h"

@interface WindowFileMap : NSObject

@property (strong, nonatomic) NSMutableDictionary* windowInfo;
@property (strong, nonatomic) RecentItem* recentItem;

@end
