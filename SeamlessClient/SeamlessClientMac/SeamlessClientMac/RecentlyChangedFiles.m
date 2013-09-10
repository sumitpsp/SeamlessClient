//
//  RecentlyChangedFiles.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/9/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "RecentlyChangedFiles.h"

@implementation RecentlyChangedFiles
+(NSMutableArray*) getFilesList {
    LSSharedFileListRef recentDocsFileList;
    NSMutableArray *recentDocsFiles;
    NSMutableArray *recentDocsURLs = nil;
    UInt32 seed;
    
    recentDocsFileList = LSSharedFileListCreate(NULL,
                                                kLSSharedFileListRecentDocumentItems    , NULL);
    if (! recentDocsFileList) return NULL;
    
    recentDocsFiles = (NSArray *)CFBridgingRelease(LSSharedFileListCopySnapshot(recentDocsFileList,
                                                                                &seed));
    UInt32 resolveFlags = kLSSharedFileListNoUserInteraction;
    if (recentDocsFiles) {
        recentDocsURLs = [NSMutableArray array];
        
        for (id file in recentDocsFiles) {
            CFURLRef fileURL = NULL;
            NSURL* nsURL = NULL;
            LSSharedFileListItemResolve((__bridge LSSharedFileListItemRef)file, 0,
                                        (CFURLRef *)&fileURL, NULL);
            nsURL = (__bridge NSURL *) fileURL;
            if (fileURL) {
                [recentDocsURLs addObject:nsURL];
                NSString* path = [NSString stringWithFormat:@"%@", fileURL];
                //NSString* path = [event eventPath];
                NSString* name = [path lastPathComponent];
                NSLog(@"%@", name);
                NSLog(@"%@", nsURL);
            }
        }
    }
    
    NSLog(@"%@", NSHomeDirectory());
    
    CFRelease(recentDocsFileList);
    return recentDocsFiles;
}

@end
