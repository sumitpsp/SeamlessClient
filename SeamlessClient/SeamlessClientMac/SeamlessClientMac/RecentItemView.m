//
//  RecentItemView.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/13/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "RecentItemView.h"
#import "ShareViewController.h"
#import "Contact.h"
#define menuItem ([self enclosingMenuItem])

@implementation RecentItemView

- (id)initWithFrame:(NSRect)frame
{
    self = [super initWithFrame:frame];
    if (self) {
        self.item = [[RecentItem alloc] init];
        [self setNeedsDisplay:YES];
    }
    
    return self;
}

- (id)initWithFrame:(NSRect)frame andItem:(RecentItem*) item{
    self = [super initWithFrame:frame];
    if (self) {
        self.item = [[RecentItem alloc] init];
        self.item = item;
        self.windowController = [[ShareWindowController alloc] initWithWindowNibName:@"Window"];
        self.windowController.item = self.item;
        [self setNeedsDisplay:YES];
    }
    return self;
}

-(void)awakeFromNib {
    [super awakeFromNib];
}

- (void)openWindow:(id)sender {

    
    NSWindow *window = [self window]; // Get the window to open
    [window makeKeyAndOrderFront:nil];
}

- (IBAction)shareRequestedByWindow:(id)sender {
    NSLog(@"Sharing item %@", self.item.name);
}

- (IBAction) shareRequested:(id)sender {
    NSLog(@"Share requested");
    Contact* c = [[ContactsList sharedContactsList].contacts objectAtIndex:0];
    NSLog(@"Name is %@", c.name);
    
    CGRect eventFrame = [[[NSApp currentEvent] window] frame];
    CGPoint eventOrigin = eventFrame.origin;
    CGSize eventSize = eventFrame.size;
    
    if(self.windowController == nil) {
        self.windowController = [[ShareWindowController alloc] initWithWindowNibName:@"Window"];
        self.windowController.item = self.item;
    }
    self.shareWindow = [self.windowController window];
    
    
    // Calculate the position of the window to
    // place it centered below of the status item
    CGRect windowFrame = self.shareWindow.frame;
    CGSize windowSize = windowFrame.size;
    CGPoint windowTopLeftPosition = CGPointMake(eventOrigin.x + eventSize.width/2.f - windowSize.width/2.f, eventOrigin.y - 20);
    [NSApp activateIgnoringOtherApps:YES];
    // Set position of the window and display it
    [self.shareWindow setFrameTopLeftPoint:windowTopLeftPosition];
    [self.shareWindow makeKeyAndOrderFront:self];
    [self.shareWindow setIsVisible:YES];
    [self.shareWindow setLevel:NSStatusWindowLevel];
    [self.shareWindow setAcceptsMouseMovedEvents:YES];
    [self.shareWindow setIgnoresMouseEvents:NO];
    NSLog(@"Window Level is %ld", [self.shareWindow level]);
    //[window set
    
    // Show your window in front of all other apps
    
}


- (void)drawRect:(NSRect)rect
{
    
    NSRect fullBounds = [self bounds];
    fullBounds.size.height += 4;
    [[NSBezierPath bezierPathWithRect:fullBounds] setClip];
    
    [[NSColor whiteColor] set];
    NSRect frame = NSMakeRect(5, 20, 16, 16);
    self.icon = [[NSImageView alloc] initWithFrame:frame];
    [self.icon setImage:self.item.icon];
    [self addSubview:self.icon];
	
    frame = NSMakeRect(58, 20, 100, 20);
    self.name = [[NSTextField alloc] initWithFrame:frame];
    [self.name setStringValue:self.item.name];
    [self.name setFont:[NSFont fontWithName:@"Helvetica Neue" size:13]];
    [self.name setBezeled:NO];
    [self.name setDrawsBackground:NO];
    [self.name setEditable:NO];
    [self.name setSelectable:NO];
    [self addSubview:self.name];
    
    frame = NSMakeRect(164, 17, 80, 20);
    self.shareButton = [[NSButton alloc] initWithFrame:frame];
    [self.shareButton setTitle:@"Share"];
    [self.shareButton setBordered:FALSE];
    [self.shareButton setBezelStyle:NSRoundRectBezelStyle];
    [self.shareButton setButtonType:NSPushOnPushOffButton];
    [self.shareButton setAction:@selector(shareRequested:)];
    [self.shareButton setTarget:self];
    
    frame = NSMakeRect(58, 0, 200, 15);
    self.path = [[NSTextField alloc] initWithFrame:frame];
    [self.path setStringValue:self.item.path];
    [self addSubview:self.shareButton];
    [self.path setTextColor:[NSColor grayColor]];
    [self.path setFont:[NSFont fontWithName:@"Arial" size:10]];
    [self.path setBezeled:NO];
    [self.path setDrawsBackground:NO];
    [self.path setEditable:NO];
    [self.path setSelectable:NO];
    //[self addSubview:self.path];
    
    BOOL isHighlighted = [menuItem isHighlighted];
    if (isHighlighted) {
        [[NSColor blueColor] set];
        [[NSColor selectedMenuItemColor] set];
        [NSBezierPath fillRect:rect];
    } else {
        [super drawRect: rect];
    }
    NSRectFill(rect);
}

@end
