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
#import "ItemView.h"
#import "ItemViewController.h"
#import "OpenApplicationWindows.h"
#import "RecentlyChangedFiles.h"

@implementation AppDelegate

- (void)applicationDidFinishLaunching:(NSNotification *)aNotification
{
    // Insert code here to initialize your application
}

-(void)awakeFromNib{
    
    statusItem = [[NSStatusBar systemStatusBar] statusItemWithLength:NSVariableStatusItemLength] ;
    NSImage* icon = [NSImage imageNamed:@"podcast.png"];

    [statusItem setImage:icon];
    [statusItem setMenu:self.menu];
    [statusItem setHighlightMode:YES];
    [self.menu setDelegate:self];
    [self setupEventListener];
}

-(void) getRecentItems:(NSMutableArray*)recentFiles recentWindows:(NSMutableArray*)windows {
    NSWorkspace * ws = [NSWorkspace sharedWorkspace];
    NSArray * apps = [ws runningApplications];
    for (NSRunningApplication* app in apps) {
        pid_t pid = [app processIdentifier];
        for (NSMutableDictionary* windowInfo in windows) {
            pid_t wpid = (int)[[windowInfo valueForKey:@"kCGWindowOwnerPID"] integerValue];
            if(pid == wpid) {
                NSLog(@"PID match for %ld", (long)pid);
            }
        }
    }
}

-(void)addRecentFiles:(NSMutableArray*) files {
    for (NSURL* url in files) {
        NSString* name = [url lastPathComponent];
        NSString* path = [url absoluteString];
        RecentItem* item = [[RecentItem alloc] initWithName:name andPath:path];
        [self.recentItemsDataController addItem:item];
    }
}

-(void)menuWillOpen:(NSMenu *)menu{
    ItemView* itemView = [[ItemView alloc] initWithFrame:NSMakeRect(0, 0, 200, 50)];
    ItemView* itemView2 = [[ItemView alloc] initWithFrame:NSMakeRect(0, 0, 200, 50)];
    NSMutableArray* windows = [OpenApplicationWindows getOpenWindows];
    NSMutableArray* recentFiles = [RecentlyChangedFiles getFilesList];
    [self getRecentItems:recentFiles recentWindows:windows];
    NSMenuItem* currApp = [[NSMenuItem alloc] init];
    [currApp setView:itemView];
    [currApp view];
    [[statusItem menu] addItem:currApp];
    
    NSMenuItem* currApp2 = [[NSMenuItem alloc] init];
    [currApp2 setView:itemView2];
    [currApp2 view];
    [[statusItem menu] addItem:currApp2];
    
}

-(void)menuDidClose:(NSMenu *)menu{
    [[statusItem menu] removeAllItems];
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
	// [_events startWatchingPaths:paths];
    
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
            NSLog(@"Hidden Directory");
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
                NSLog(@"Lib Folder found");
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
    NSLog(@"Path is %@", [event eventPath]);
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
        NSLog(@"Directory is No good");
        return;
    }
    else {
        NSLog(@"Changed Directory");
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
            NSLog(@"File is %@", name);
            RecentItem* item = [[RecentItem alloc] initWithName:name andPath:[url absoluteString]];
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
