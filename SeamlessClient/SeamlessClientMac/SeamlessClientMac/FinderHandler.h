//
//  FinderHandler.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 8/13/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "finder.h"
#import "RecentItemsDataController.h"

@interface FinderHandler : NSObject
-(NSString*) getRecentlyOpened;
-(NSString*) getRecentlyChanged;

@property (nonatomic, strong) finderApplication* finder;
@end
