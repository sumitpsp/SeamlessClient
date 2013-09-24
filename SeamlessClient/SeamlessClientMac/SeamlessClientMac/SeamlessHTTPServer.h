//
//  HTTPServer.h
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/11/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "HTTPConnection.h"
#import "MultipartFormDataParser.h"

@interface SeamlessHTTPServer : HTTPConnection {
    MultipartFormDataParser*        parser;
	NSFileHandle*					storeFile;
	
	NSMutableArray*					uploadedFiles;
}

@end
