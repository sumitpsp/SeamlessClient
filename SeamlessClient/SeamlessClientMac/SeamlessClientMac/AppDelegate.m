//
//  AppDelegate.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 7/26/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "AppDelegate.h"
#import <ScriptingBridge/ScriptingBridge.h>
#import "FinderHandler.h"
#import "SCEvent.h"
#import "SCEvents.h"
#import "ItemViewController.h"
#import "OpenApplicationWindows.h"
#import "RecentlyChangedFiles.h"
#import "HTTPServer.h"
#import "SeamlessHTTPServer.h"
#import "DDLog.h"
#import "DDTTYLogger.h"
#import "RecentItemView.h"
#import "WindowFileMap.h"
#import "ApplicationFileMap.h"
#import "iTunes.h"
#import "ContactsList.h"

// Log levels: off, error, warn, info, verbose
static const int ddLogLevel = LOG_LEVEL_VERBOSE;

@implementation AppDelegate

- (void) serverEvent:(NSNotification *)notification {
    NSLog(@"Files View received New Server Info");
    NSDictionary* userInfo = [notification userInfo];
    NSNumber* temp = [userInfo objectForKey:@"port"];
    self.host = [userInfo objectForKey:@"host"];
    self.port = [temp integerValue];
    Contact* newContact = [[Contact alloc] init];
    newContact.address = [[ContactAddress alloc] init];
    newContact.name = @"Ari Gold";
    newContact.address.host = [userInfo objectForKey:@"host"];
    newContact.address.port = [temp integerValue];
    [self.contacts addObject:newContact];
    
    NSLog(@"Host is %@ and port is %ld", self.host, (long)self.port);
    [[ContactsList sharedContactsList] addContact:newContact];
}

- (void) activeURLEvent:(NSNotification *)notification {
    NSLog(@"Active URL received New Server Info");
    NSDictionary* userInfo = [notification userInfo];
    NSString* url = [userInfo objectForKey:@"url"];
    NSString* title = [userInfo objectForKey:@"title"];
    [self.activeTab setObject:title
                       forKey:@"title"];
    [self.activeTab setObject:url forKey:@"url"];
    //self.activeTab[@"title"] = title;
    //self.activeTab[@"url"] = url;
    NSLog(@"Title is %@ and URL is %@", title, url);
    
    
    
}


- (void)applicationDidFinishLaunching:(NSNotification *)aNotification
{
   // self.menu = [[NSMenu alloc] initWithTitle:@""];
    self.statusItem = [[NSStatusBar systemStatusBar] statusItemWithLength:NSVariableStatusItemLength] ;
    self.serverBrowser = [[ServerBrowser alloc] init];
     NSImage* icon = [NSImage imageNamed:@"podcast.png"];
    
    [self.statusItem setImage:icon];
    [self.statusItem setMenu:self.menu];
  //  [self.statusItem setHighlightMode:YES];
    [self.menu setDelegate:self];
    
    [self setupEventListener];
    
    self.recentItemsDataController = [[RecentItemsDataController alloc] init];
    self.itemQueue = [[NSMutableArray alloc] init];
    self.applicationFileMap = [[NSMutableArray alloc] init];
    self.windowFileMap = [[NSMutableArray alloc] init];
    self.activeTab = [[NSMutableDictionary alloc] initWithObjectsAndKeys:@"", @"title", @"", @"url", nil];
    NSLog(@"Share requested");
    [[NSNotificationCenter defaultCenter] addObserver:self selector:@selector(activeURLEvent:) name:@"URLNotification" object:nil];
    [[NSNotificationCenter defaultCenter] addObserver:self selector:@selector(serverEvent:) name:@"ServerNotification" object:nil];
    [self createHTTPServer];
    [self.serverBrowser start];
    self.contacts = [[NSMutableArray alloc] init];
    
    NSApplication* app = [NSApplication sharedApplication];
    
    if( [app respondsToSelector: @selector(setActivationPolicy:)] ) {
        
        NSMethodSignature* method = [[app class] instanceMethodSignatureForSelector: @selector(setActivationPolicy:)];
        NSInvocation* invocation = [NSInvocation invocationWithMethodSignature: method];
        [invocation setTarget: app];
        [invocation setSelector: @selector(setActivationPolicy:)];
        NSInteger myNSApplicationActivationPolicyAccessory = 0;
        [invocation setArgument: &myNSApplicationActivationPolicyAccessory atIndex: 2];
        [invocation invoke];
        
    }
    
    
    // Insert code here to initialize your application
}



-(void)awakeFromNib{


}

-(BOOL) addIfExists:(NSString*) name {
    RecentItem* item = [self.recentItemsDataController fileWithNameExists:name];
    if(item != nil) {
        [self.itemQueue enqueue:item];
        NSLog(@"Added item %@", [item name]);
        return TRUE;
    }
    else {
        return FALSE;
    }
}


-(void) createHTTPServer {
	// Configure our logging framework.
	// To keep things simple and fast, we're just going to log to the Xcode console.
	[DDLog addLogger:[DDTTYLogger sharedInstance]];
	
	// Initalize our http server
	httpServer = [[HTTPServer alloc] init];
	
	// Tell the server to broadcast its presence via Bonjour.
	// This allows browsers such as Safari to automatically discover our service.
	[httpServer setType:@"_http._tcp."];
	
	// Normally there's no need to run our server on any specific port.
	// Technologies like Bonjour allow clients to dynamically discover the server's port at runtime.
	// However, for easy testing you may want force a certain port so you can just hit the refresh button.
	[httpServer setPort:11111];
	
	// We're going to extend the base HTTPConnection class with our MyHTTPConnection class.
	// This allows us to do all kinds of customizations.
	[httpServer setConnectionClass:[SeamlessHTTPServer class]];
	
	// Serve files from our embedded Web folder
	/**NSString *webPath = [[[NSBundle mainBundle] resourcePath] stringByAppendingPathComponent:@"Web"];
     DDLogInfo(@"Setting document root: %@", webPath);
     
     [httpServer setDocumentRoot:webPath];
     */
	
	NSError *error = nil;
	if(![httpServer start:&error])
	{
		DDLogError(@"Error starting HTTP Server: %@", error);
	}


}

-(void) getRecentItems:(NSMutableArray*)recentFiles recentWindows:(NSMutableArray*)windows {
    NSWorkspace * ws = [NSWorkspace sharedWorkspace];
    NSArray * apps = [ws runningApplications];
    for (NSRunningApplication* app in apps) {
        pid_t pid = [app processIdentifier];
        for (NSMutableDictionary* windowInfo in windows) {
            pid_t wpid = (int)[[windowInfo valueForKey:@"kCGWindowOwnerPID"] integerValue];
            if(pid == wpid) {
                NSString* probableName = [windowInfo valueForKey:@"kCGWindowName"];
                NSLog(@"PID match for %ld. Probable name is %@", (long)pid, probableName);
                if(![self addIfExists:probableName]) {
                    [self addIfExists:[probableName stringByDeletingPathExtension]];
                };
            }
        }
    }
}

-(void)addRecentFiles {
    for (id url in self.recentFiles) {
        CFURLRef cfurl = (__bridge CFURLRef)(url);
        NSURL* url = (__bridge NSURL *)(cfurl);
        NSString* path = url.path;
        NSString* name = [path lastPathComponent];
        NSLog(@"Adding name %@", name);
        
        RecentItem* item = [[RecentItem alloc] initWithName:name andPath:path];
        [self.recentItemsDataController addItem:item];
    }
}

-(void) matchFilesToWindows {
    for (NSMutableDictionary* windowInfo in self.windows) {
        NSString* name = [windowInfo valueForKey:@"kCGWindowName"];
        NSString* probableName = [name stringByAddingPercentEscapesUsingEncoding:NSUTF8StringEncoding];        
        RecentItem* item = [self.recentItemsDataController fileWithNameExists:probableName];
        if(item != nil) {
            WindowFileMap* map = [[WindowFileMap alloc] init];
            map.windowInfo = windowInfo;
            map.recentItem = item;
            [self.windowFileMap addObject:map];
        }
        else {
            NSLog(@"Cant find file for window name %@", probableName);  
        }
    }

}

-(void) matchApplicationToFile {
    NSWorkspace * ws = [NSWorkspace sharedWorkspace];
    NSArray * apps = [ws runningApplications];
    for (NSRunningApplication* app in apps) {
        pid_t pid = [app processIdentifier];
        for (id element in self.windowFileMap) {
            WindowFileMap *mapping = (WindowFileMap*) element;
            NSMutableDictionary* windowInfo = mapping.windowInfo;
            pid_t wpid = (int)[[windowInfo valueForKey:@"kCGWindowOwnerPID"] integerValue];
            if(pid == wpid) {
                ApplicationFileMap* appFileMap = [[ApplicationFileMap alloc] init];
                appFileMap.application = app;
                appFileMap.recentItem = mapping.recentItem;
                appFileMap.recentItem.icon = app.icon;
                [self.applicationFileMap addObject:appFileMap];
            }
        }
    }
}

-(RecentItem*) currentItunesItem {
    RecentItem* iTunesItem = [[RecentItem alloc] init];
    iTunesApplication* itunes = [SBApplication applicationWithBundleIdentifier:@"com.apple.iTunes"];
    iTunesItem.name = itunes.currentTrack.name;
    iTunesItem.path= itunes.currentTrack.album;
    return iTunesItem;
    
}

-(RecentItem*) getItemForActiveApp {
    NSWorkspace * ws = [NSWorkspace sharedWorkspace];
    NSRunningApplication* activeApp = [ws frontmostApplication];
    NSLog(@"Active Application %@ Bundle Identifier %@", [activeApp localizedName], [activeApp bundleIdentifier] );
    if ([activeApp.localizedName isEqualToString:@"Google Chrome"]) {
        RecentItem* browserItem = [[RecentItem alloc] init];
        browserItem.name = [self.activeTab objectForKey:@"title"];
        browserItem.path = [self.activeTab objectForKey:@"url"];
       // browserItem.name = self.activeTab[@"title"];
       // browserItem.path= self.activeTab[@"url"];
        browserItem.icon = [activeApp icon];
        return browserItem;
    }
    
    if([activeApp.localizedName isEqualToString:@"iTunes"]) {
        RecentItem* itunesItem = [self currentItunesItem];
        itunesItem.icon = [activeApp icon];
        return itunesItem;
    }
    for (ApplicationFileMap* map in self.applicationFileMap) {
        if(map.application.processIdentifier == activeApp.processIdentifier) {
            RecentItem* fileitem = map.recentItem;
            [self.applicationFileMap removeObject:map];
            return fileitem;
        }
    }
    return nil;
}

-(void) addRemainingitems {
    for (ApplicationFileMap* map in self.applicationFileMap) {
            RecentItem* fileitem = map.recentItem;
            [self.itemQueue enqueue:fileitem];
    }
    [self.applicationFileMap removeAllObjects];
    
}
-(void) addStatusItems {
    [self matchFilesToWindows];
    [self matchApplicationToFile];
    for (id obj in self.applicationFileMap) {
        ApplicationFileMap* appFileMap = (ApplicationFileMap*) obj;
        NSLog(@"App file map %@", appFileMap.recentItem.name);
    }
    
    RecentItem* item = [self getItemForActiveApp];
    if(item!=nil) {
        [self.itemQueue enqueue:item];
    }
    [self addRemainingitems];

}

-(void) showStatusItems {
    NSUInteger count = 0;
    while( count < 10 && ([self.itemQueue count] != 0)) {
        count++;
        RecentItem* item = [self.itemQueue dequeue];
        RecentItemView *myview =  [[RecentItemView alloc] initWithFrame:NSMakeRect(0, 0, 250, 40) andItem:item];
        NSMenuItem* entry = [[NSMenuItem alloc] init];
        [entry setView:myview];
        [entry view];
        [[self.statusItem menu] addItem:entry];
        NSLog(@"Added entry for Item %@", [item name]);
    }
}

-(void)menuWillOpen:(NSMenu *)menu{
    
    [[self.statusItem menu] removeAllItems];
    
    self.windows = [OpenApplicationWindows getOpenWindows];
    self.recentFiles = [RecentlyChangedFiles getFilesList];

    [self addRecentFiles];
    [self addStatusItems];
    [self showStatusItems];
    /**
    RecentItemView *myview =  [[RecentItemView alloc] initWithFrame:NSMakeRect(0, 0, 250, 40)];
    NSMenuItem* currApp3 = [[NSMenuItem alloc] init];
    [currApp3 setView:myview];
    [currApp3 view];
    [[statusItem menu] addItem:currApp3];
    
    RecentItemView *myview2 =  [[RecentItemView alloc] initWithFrame:NSMakeRect(0, 0, 250, 40)];
    NSMenuItem* currApp4 = [[NSMenuItem alloc] init];
    [currApp4 setView:myview2];
    [currApp4 view];
    [[statusItem menu] addItem:currApp4];*/
    
    
    
}

-(void)menuDidClose:(NSMenu *)menu{

//    [[self.statusItem menu] removeAllItems];
    [self.windowFileMap removeAllObjects];
    [self.applicationFileMap removeAllObjects];
    
}

- (void)setupEventListener
{
	if (_events) return;
	
    _events = [[SCEvents alloc] init];
    
    [_events setDelegate:self];
    
    NSMutableArray *paths = [NSMutableArray arrayWithObject:NSHomeDirectory()];
    //NSMutableArray *excludePaths = [NSMutableArray arrayWithObject:[NSHomeDirectory() stringByAppendingPathComponent:SCEventsDownloadsDirectory]];
    
	// Set the paths to be excluded
	//[_events setExcludedPaths:excludePaths];
	
	// Start receiving events
	 [_events startWatchingPaths:paths];
    
	// Display a description of the stream
	//NSLog(@"%@", [_events streamDescription]);
}

-(BOOL) checkIfHiddenDirectory: (NSString*) path {
    NSString* homeDirectory = NSHomeDirectory();
    if([path isEqualToString:homeDirectory] || [path isEqualToString:@""]) {
        return FALSE;
    }
    else {
        NSString* directoryName = [path lastPathComponent];
        if([directoryName hasPrefix:@"."]) {
       //     NSLog(@"Hidden Directory");
            return TRUE;
        }
        else {
            return [self checkIfHiddenDirectory:[path stringByDeletingLastPathComponent]];
        }
    }
}

-(BOOL) checkIfSystemFolder: (NSString*) path {
    NSString* libFolder = [NSHomeDirectory() stringByAppendingPathComponent:@"Library"];
    NSString* apdbFolder = [[[[NSHomeDirectory() stringByAppendingPathComponent:@"Pictures"] stringByAppendingPathComponent:@"iPhoto Library.photolibrary"] stringByAppendingPathComponent:@"Database"] stringByAppendingPathComponent:@"apdb"];
    NSMutableArray* listOfSystemDirectories = [[NSMutableArray alloc] init];
    [listOfSystemDirectories addObject:libFolder];
    [listOfSystemDirectories addObject:apdbFolder];
    if([path isEqualToString:NSHomeDirectory()] || [path isEqualToString:@""]) {
        return FALSE;
    }
    else {
        for(id object in listOfSystemDirectories) {
            NSString* dirPath = (NSString*) object;
            if([dirPath isEqualToString:path]) {
          //      NSLog(@"Lib Folder found");
                return TRUE;
            }
        }
        return [self checkIfSystemFolder:[path stringByDeletingLastPathComponent]];
    }
}

/**
 * This is the only method to be implemented to conform to the SCEventListenerProtocol.
 * As this is only an example the event received is simply printed to the console.
 *
 * @param pathwatcher The SCEvents instance that received the event
 * @param event       The actual event
 */
- (void)pathWatcher:(SCEvents *)pathWatcher eventOccurred:(SCEvent *)event
{ 
    NSString* path = [event eventPath];
  //   NSLog(@"Id is %ld", (unsigned long)[event eventId]);
  //  NSLog(@"Path is %@", [event eventPath]);
    NSFileManager *fm = [NSFileManager defaultManager];
    NSURL *directoryURL = [NSURL fileURLWithPath:[event eventPath]];
    NSArray *urls = [fm contentsOfDirectoryAtURL:directoryURL
                      includingPropertiesForKeys:[NSArray arrayWithObjects:NSURLNameKey, NSURLIsDirectoryKey, NSURLContentModificationDateKey, nil]
                                         options:NSDirectoryEnumerationSkipsHiddenFiles | NSDirectoryEnumerationSkipsPackageDescendants | NSDirectoryEnumerationSkipsSubdirectoryDescendants
                                           error:nil];
    NSError *error = nil;
    NSString *name;
    NSDate *modificationDate;
    NSNumber *isDirectory;
    
    if([self checkIfHiddenDirectory:[event eventPath]] || [self checkIfSystemFolder:[event eventPath]] ) {
       // NSLog(@"Directory is No good");
        return;
    }
    else {
     //   NSLog(@"Changed Directory");
    }
    
    for (NSURL *url in urls) {
        if (![url getResourceValue:&name forKey:NSURLNameKey error:&error]) {
            NSLog(@"name: %@", [error localizedDescription]);error = nil;
        }
        if (![url getResourceValue:&modificationDate forKey:NSURLContentModificationDateKey error:&error]) {
            NSLog(@"modificationDate: %@", [error localizedDescription]); error = nil;
        }
        if (![url getResourceValue:&isDirectory forKey:NSURLIsDirectoryKey error:&error]) {
            NSLog(@"isDirectory: %@", [error localizedDescription]); error = nil;
        }
        
        NSDate* now = [NSDate date];
        NSTimeInterval timeSince = [now timeIntervalSinceDate:modificationDate];
        if(timeSince < 10) {
           // NSLog(@"%@ %@ %@", name, modificationDate, isDirectory);
            NSLog(@"Changed File is %@", name);
            RecentItem* item = [[RecentItem alloc] initWithName:name andPath:[url path]];
            [self.recentItemsDataController addItem:item];
            
        }
        else {
           // NSLog(@"NOT SAME DAY");
           // NSLog(@"Today: %@ Modified: %@ timeSince: %f", now, modificationDate, timeSince);
            // NSLog(@"%@ %@ %@", name, modificationDate, isDirectory);
        }
    }
    
    
    
}

@end
