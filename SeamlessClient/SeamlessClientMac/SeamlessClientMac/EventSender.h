//
//  EventSender.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/20/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface EventSender : NSObject

@property(nonatomic, strong) NSString* host;
@property(nonatomic, assign) NSInteger port;

+(BOOL) sendData:(NSData*) data withHost:(NSString*) host andPort:(NSInteger) port;

@end
