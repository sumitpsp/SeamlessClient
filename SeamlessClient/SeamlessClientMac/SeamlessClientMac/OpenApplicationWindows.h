//
//  OpenApplicationWindows.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/9/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface OpenApplicationWindows : NSObject

@property (strong, nonatomic) NSDictionary* openWindows;
+(NSMutableArray*) getOpenWindows;
@end
