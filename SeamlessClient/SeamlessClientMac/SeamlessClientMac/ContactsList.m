//
//  ContactsList.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/23/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "ContactsList.h"

@implementation ContactsList
static ContactsList* _sharedContactsList = nil;

+(ContactsList*)sharedContactsList {
    @synchronized([ContactsList class]) {
        if (!_sharedContactsList)
            [[self alloc] init];
        return _sharedContactsList;
    }
    return nil;
}

+(id)alloc {
        @synchronized([ContactsList class])
        {
            NSAssert(_sharedContactsList == nil, @"Attempted to allocate a second instance of a singleton.");
            _sharedContactsList = [super alloc];
            return _sharedContactsList;
        }
    return nil;
}

-(id)init {
    self = [super init];
    if (self != nil) {
        self.contacts = [[NSMutableArray alloc] init];
    }
    return self;
}
-(void)sayHello {
    NSLog(@"Hello World!");
}

-(void) addContact:(Contact*) contact {
    [self.contacts addObject:contact];
}

@end
