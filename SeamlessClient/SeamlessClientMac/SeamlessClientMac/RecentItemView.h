//
//  RecentItemView.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/13/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Cocoa/Cocoa.h>
#import "ShareWindowController.h"
#import "RecentItem.h"
#import "ContactsList.h"

@interface RecentItemView : NSView

@property (strong, nonatomic) NSImageView *icon;
@property (strong, nonatomic) NSTextField*  name;
@property (strong, nonatomic) NSButton* shareButton;
@property (strong, nonatomic) NSTextField* path;
@property (strong, nonatomic) ShareWindowController* windowController;
@property (strong, nonatomic) RecentItem* item;
@property (strong, nonatomic) NSWindow* shareWindow;

- (id)initWithFrame:(NSRect)frame andItem:(RecentItem*) item;

@end
