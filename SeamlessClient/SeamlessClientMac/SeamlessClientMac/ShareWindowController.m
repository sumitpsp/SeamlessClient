//
//  ShareWindowController.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/15/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "ShareWindowController.h"
#import "ContactsList.h"
#import "Contact.h"
#import "AFHTTPClient.h"
#import "AFHTTPRequestOperation.h"
#import "MultipartFormDataParser.h"

@interface ShareWindowController ()

@end

@implementation ShareWindowController

- (id)initWithWindow:(NSWindow *)window
{
    self = [super initWithWindow:window];
    if (self) {
        self.item = [[RecentItem alloc] init];
    }
    
    return self;
}

- (BOOL) canBecomeKeyWindow { return YES; }
- (BOOL) canBecomeMainWindow { return YES; }
- (BOOL) acceptsFirstResponder { return YES; }
- (BOOL) becomeFirstResponder { return YES; }
- (BOOL) resignFirstResponder { return YES; }

- (IBAction)shareRequestedByWindow:(id)sender {
    NSLog(@"Sharing item %@", self.item.name);
    Contact* c = [[ContactsList sharedContactsList].contacts objectAtIndex:0];
    if(c.address.host == nil)  {
        NSLog(@"NIL HOST");
        [self.window close];
        return;
    }
    NSString* url = [NSString stringWithFormat:@"http://%@:%ld", c.address.host, (long)c.address.port];
    NSURL *urlPath = [NSURL URLWithString:url];
    
    NSLog(@"Trying to transfer %@ to %@", self.item.name, url);
    
    NSData* fileData = nil;
    
    if([[NSFileManager defaultManager] fileExistsAtPath:self.item.path])
    {
        fileData = [[NSFileManager defaultManager] contentsAtPath:self.item.path];
    }
        else
    {
        NSLog(@"File does not exist");
        [self.window close];
        return;
    }
    NSError* error;


    AFHTTPClient *httpClient = [[AFHTTPClient alloc] initWithBaseURL:urlPath];
    NSMutableURLRequest *request = [httpClient multipartFormRequestWithMethod:@"POST" path:@"/connection" parameters:nil constructingBodyWithBlock: ^(id <AFMultipartFormData>formData) {
        [formData appendPartWithFileData:fileData name: [[self.item.name lastPathComponent] stringByDeletingPathExtension] fileName:self.item.name mimeType:@"application/octet-stream"];
    }];
    
    AFHTTPRequestOperation *operation = [[AFHTTPRequestOperation alloc] initWithRequest:request];
    [operation setUploadProgressBlock:^(NSUInteger bytesWritten, long long totalBytesWritten, long long totalBytesExpectedToWrite) {
        NSLog(@"Sent %lld of %lld bytes", totalBytesWritten, totalBytesExpectedToWrite);
    }];
    [httpClient enqueueHTTPRequestOperation:operation];
    [self.window close];
    
}

- (void)windowDidLoad
{
    [super windowDidLoad];
    [self.shareButton setAction:@selector(shareRequestedByWindow:)];
    [self.shareButton setTarget:self];
    [self.contacts setDataSource:self];
    [self.contacts setDelegate:self];
    
    // Implement this method to handle any initialization after your window controller's window has been loaded from its nib file.
    
}

- (NSInteger)numberOfRowsInTableView:(NSTableView *)tableView {
    return [[ContactsList sharedContactsList].contacts count];
}

- (id)          tableView:(NSTableView *)tableView
objectValueForTableColumn:(NSTableColumn *)column
                      row:(NSInteger)rowIndex {
    
    NSString *identifier = [column identifier];
    if ([identifier length] == 0)
        return nil;
    
    NSButtonCell* data = [column dataCellForRow:rowIndex];
    [data setButtonType:NSSwitchButton];
    Contact* c = [[ContactsList sharedContactsList].contacts objectAtIndex:rowIndex];
    NSLog(@"%@", c.name);
    [data setTitle:c.name];
    return data;
}

- (BOOL)tableView:(NSTableView *)tableView shouldSelectRow:(NSInteger)rowIndex {
    NSLog(@"%li tapped!", (long)rowIndex);
    return YES;
}


- (void)tableView:(NSTableView *)aTableView setObjectValue:(id)anObject forTableColumn:(NSTableColumn *)aTableColumn row:(NSInteger)rowIndex {
        NSButtonCell* data = [aTableColumn dataCellForRow:rowIndex];
    if(data.state == YES) {
        data.state = NO;
    }
    else {
        data.state = YES;
    }
}


@end
