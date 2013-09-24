//
//  Contact.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/23/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "ContactAddress.h"

@interface Contact : NSObject
@property (strong, nonatomic) NSString* name;
@property (strong, nonatomic) NSString* description;
@property (strong, nonatomic) ContactAddress* address;

@end
