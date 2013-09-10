//
//  ItemView.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 8/25/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "ItemView.h"
#define menuItem ([self enclosingMenuItem])

@implementation ItemView

- (id)initWithFrame:(NSRect)frame
{
    self = [super initWithFrame:frame];
    if (self) {
    
    }
    
    return self;
}

- (id)initWithFrame:(NSRect)frame fileName:(NSString*)name fileType:(NSString*)type filePath:(NSString*) path
{
    fileName = name;
    fileType = type;
    filePath = path;
    self = [super initWithFrame:frame];
    if (self) {
        
    }
    
    return self;
}

- (IBAction) shareRequested:(id)sender {
    NSLog(@"Share requested");
}

- (void)drawRect:(NSRect)rect
{
    NSRect fullBounds = [self bounds];
    fullBounds.size.height += 4;
    //[[NSBezierPath bezierPathWithRect:fullBounds] setClip];
    
    [[NSColor whiteColor] set];
	
    NSRect frame = NSMakeRect(75, 0, 100, 50);
    NSButton* pushButton = [[NSButton alloc] initWithFrame: frame];
    [pushButton setTitle:@"Share"];
    [pushButton setTarget:self];
    [pushButton setAction:@selector(shareRequested:)];

    pushButton.bezelStyle = NSRoundedBezelStyle;
    
    [self addSubview:pushButton];
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
