//
//  ShareWindowController.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/15/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Cocoa/Cocoa.h>
#import "RecentItem.h"
#import "ShareWindow.h"

@interface ShareWindowController : NSWindowController <NSWindowDelegate, NSTableViewDataSource, NSTableViewDelegate>
@property (strong, nonatomic) RecentItem* item;
@property (weak) IBOutlet NSButton *optionButton;
@property (weak) IBOutlet NSButton *typeButton;
@property (weak) IBOutlet NSButton *shareButton;
@property (weak) IBOutlet NSTableView *contacts;

@end
