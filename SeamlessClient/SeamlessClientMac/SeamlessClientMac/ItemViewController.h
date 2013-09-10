//
//  ItemViewController.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 8/29/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Cocoa/Cocoa.h>

@interface ItemViewController : NSViewController
@property (weak) IBOutlet NSImageView *icon;
@property (weak) IBOutlet NSTextField *name;

@end
