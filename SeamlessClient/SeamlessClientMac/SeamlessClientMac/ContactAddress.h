//
//  ContactAddress.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/23/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface ContactAddress : NSObject
@property (nonatomic, strong) NSString* host;
@property (nonatomic, assign) NSInteger port;
@end
