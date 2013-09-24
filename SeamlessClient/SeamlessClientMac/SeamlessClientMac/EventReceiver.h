//
//  EventReceiver.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/20/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface EventReceiver : NSObject
@property (strong, nonatomic) NSData* data;

-(BOOL) parseIncomingData:(NSData*) data;
@end
