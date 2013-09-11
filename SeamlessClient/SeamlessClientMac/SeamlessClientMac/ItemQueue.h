//
//  ItemQueue.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/10/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "MutableArrayQueue.h"

@interface ItemQueue : NSObject
@property (strong, nonatomic) NSMutableArray* queue;
@end
