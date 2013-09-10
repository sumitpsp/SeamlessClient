//
//  ItemView.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 8/25/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Cocoa/Cocoa.h>

@interface ItemView : NSView {
    NSString* fileName;
    NSString* fileType;
    NSString* filePath;
}

@property (nonatomic, strong) NSString* fileName;
@property (nonatomic, strong) NSString* fileType;
@property (nonatomic, strong) NSString* filePath;

@end
