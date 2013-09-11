//
//  MutableArrayQueue.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/10/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface NSMutableArray (QueueAdditions)
- (id) dequeue;
- (void) enqueue:(id)obj;
@end
