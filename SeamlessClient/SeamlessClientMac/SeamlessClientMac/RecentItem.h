//
//  RecentItem.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/9/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface RecentItem : NSObject

-(id) initWithName:(NSString*)newName andPath:(NSString*) newPath;

@property (nonatomic, strong) NSString* name;
@property (nonatomic, strong) NSString *path;
@property (nonatomic, strong) NSFileHandle *type;
@property (nonatomic, strong) NSImage* icon;
@end
