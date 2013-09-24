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
#import "RecentItemView.h"
#import "ShareViewController.h"
#import "MutableArrayQueue.h"
#import "ServerBrowser.h"
#import "Contact.h"

@class HTTPServer;

@interface AppDelegate : NSObject <NSApplicationDelegate, NSMenuDelegate, SCEventListenerProtocol> {
    SCEvents *_events;
    HTTPServer *httpServer;
}

- (void)setupEventListener;


@property (strong, nonatomic) NSStatusItem *statusItem;
@property (strong, nonatomic) ShareViewController* shareViewController;
@property (strong, nonatomic) ServerBrowser* serverBrowser;
@property (weak) IBOutlet NSMenu *menu;
@property (strong, nonatomic) RecentItemsDataController* recentItemsDataController;
@property (strong, nonatomic) NSMutableArray* windowFileMap;
@property (strong, nonatomic) NSMutableArray* applicationFileMap;
@property (strong, nonatomic) NSMutableArray* windows;
@property (strong, nonatomic) NSMutableArray* recentFiles;
@property (strong, nonatomic) NSMutableArray* itemQueue;
@property (strong, nonatomic) NSMutableDictionary* activeTab;
@property (nonatomic, strong) NSString* host;
@property (nonatomic, assign) NSInteger port;
@property (nonatomic, strong) NSMutableArray* contacts;
@end
