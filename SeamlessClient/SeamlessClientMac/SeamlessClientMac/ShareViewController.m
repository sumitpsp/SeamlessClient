//
//  ShareViewController.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/14/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "ShareViewController.h"

@interface ShareViewController ()

@end

@implementation ShareViewController

- (id)initWithNibName:(NSString *)nibNameOrNil bundle:(NSBundle *)nibBundleOrNil
{
    self = [super initWithNibName:nibNameOrNil bundle:nibBundleOrNil];
    if (self) {
        // Initialization code here.
    }
    
    return self;
}

-(void)awakeFromNib{
    [super awakeFromNib];
    NSLog(@"HERE");
    [self view];
    
}

- (void)viewDidMoveToSuperview {
    NSLog(@"HERE 2");
}

@end
