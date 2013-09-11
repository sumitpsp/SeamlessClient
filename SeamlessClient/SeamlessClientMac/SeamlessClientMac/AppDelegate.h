//
//  AppDelegate.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 7/26/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Cocoa/Cocoa.h>
#include "Testcpp.h"
#import "SCEventListenerProtocol.h"
#import "RecentItemsDataController.h"
#import "ItemQueue.h"

@interface AppDelegate : NSObject <NSApplicationDelegate, NSMenuDelegate, SCEventListenerProtocol> {
    NSStatusItem *statusItem;
    __weak NSMenu *_menu;
    __weak NSView *_itemView;
    SCEvents *_events;
}

- (void)setupEventListener;


@property (assign) IBOutlet NSWindow *window;
@property (weak) IBOutlet NSMenu *menu;
@property (weak) IBOutlet NSView *itemView;
@property (strong, nonatomic) RecentItemsDataController* recentItemsDataController;
@property (strong, nonatomic) ItemQueue* itemQueue;
@end
