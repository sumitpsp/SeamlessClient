//
//  ContactsList.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/23/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "Contact.h"

@interface ContactsList : NSObject {

}
+(ContactsList*) sharedContactsList;
-(void) addContact:(Contact*) contact;
-(void)sayHello;

@property (strong, nonatomic) NSMutableArray* contacts;

@end
