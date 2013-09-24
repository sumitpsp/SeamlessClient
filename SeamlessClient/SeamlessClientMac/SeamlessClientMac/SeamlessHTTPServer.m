//
//  HTTPServer.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/11/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "SeamlessHTTPServer.h"
#import "HTTPMessage.h"
#import "HTTPDataResponse.h"
#import "DDNumber.h"
#import "HTTPLogging.h"
#import "MultipartMessageHeader.h"
#import "MultipartMessageHeaderField.h"

// Log levels : off, error, warn, info, verbose
// Other flags: trace
static const int httpLogLevel = HTTP_LOG_LEVEL_VERBOSE; // | HTTP_LOG_FLAG_TRACE;


/**
 * All we have to do is override appropriate methods in HTTPConnection.
 **/

@implementation SeamlessHTTPServer

- (void)handleUnknownMethod:(NSString *)method
{
    // Override me for custom error handling of 405 method not allowed responses.
    // If you simply want to add a few extra header fields, see the preprocessErrorResponse: method.
    // You can also use preprocessErrorResponse: to add an optional HTML body.
    //
    // See also: supportsMethod:atPath:
    
    HTTPLogVerbose(@"HTTP Server: Error 405 - Method Not Allowed: %@ (%@)", method, [self requestURI]);
    
    // Status code 405 - Method Not Allowed
    HTTPMessage *response = [[HTTPMessage alloc] initResponseWithStatusCode:405 description:nil version:HTTPVersion1_1];
    [response setHeaderField:@"Content-Length" value:@"0"];
    [response setHeaderField:@"Connection" value:@"close"];
    
    NSData *responseData = [self preprocessErrorResponse:response];
    
    // Note: We used the HTTP_FINAL_RESPONSE tag to disconnect after the response is sent.
    // We do this because the method may include an http body.
    // Since we can't be sure, we should close the connection.
}

- (BOOL)supportsMethod:(NSString *)method atPath:(NSString *)path
{
	HTTPLogTrace();
	
	// Add support for POST
    
   // HTTPLogVerbose(@"%@[%p]: Received PATH is: %@", THIS_FILE, self, path);
	
	if ([method isEqualToString:@"POST"] || [method isEqualToString:@"OPTIONS"])
	{
		if ([path isEqualToString:@"/activeurl"] || [path isEqualToString:@"/connection"] || [path isEqualToString:@"/connectionurl"])
		{
			// Let's be extra cautious, and make sure the upload isn't 5 gigs
			
			return TRUE;
		}
        
	}
	
	return [super supportsMethod:method atPath:path];
}

- (BOOL)expectsRequestBodyFromMethod:(NSString *)method atPath:(NSString *)path
{
	HTTPLogTrace();
	
	// Inform HTTP server that we expect a body to accompany a POST request
	
	if([method isEqualToString:@"POST"] && [path isEqualToString:@"/activeurl"])
		return YES;
    
    if([method isEqualToString:@"POST"] && [path isEqualToString:@"/connectionurl"]) {
        return YES;
    }
    if([method isEqualToString:@"POST"] && [path isEqualToString:@"/connection"]) {
        // here we need to make sure, boundary is set in header
        NSString* contentType = [request headerField:@"Content-Type"];
        NSUInteger paramsSeparator = [contentType rangeOfString:@";"].location;
        if( NSNotFound == paramsSeparator ) {
            return NO;
        }
        if( paramsSeparator >= contentType.length - 1 ) {
            return NO;
        }
        NSString* type = [contentType substringToIndex:paramsSeparator];
        if( ![type isEqualToString:@"multipart/form-data"] ) {
            // we expect multipart/form-data content type
            return NO;
        }
        
		// enumerate all params in content-type, and find boundary there
        NSArray* params = [[contentType substringFromIndex:paramsSeparator + 1] componentsSeparatedByString:@";"];
        for( NSString* param in params ) {
            paramsSeparator = [param rangeOfString:@"="].location;
            if( (NSNotFound == paramsSeparator) || paramsSeparator >= param.length - 1 ) {
                continue;
            }
            NSString* paramName = [param substringWithRange:NSMakeRange(1, paramsSeparator-1)];
            NSString* paramValue = [param substringFromIndex:paramsSeparator+1];
            
            if( [paramName isEqualToString: @"boundary"] ) {
                // let's separate the boundary from content-type, to make it more handy to handle
                [request setHeaderField:@"boundary" value:paramValue];
            }
        }
        // check if boundary specified
        NSLog(@"HERE returning NO");
        if( nil == [request headerField:@"boundary"] )  {
            return NO;
        }
        NSLog(@"HERE returning YES");
        return YES;
    }
	
	return [super expectsRequestBodyFromMethod:method atPath:path];
    
}

- (NSObject<HTTPResponse> *)httpResponseForMethod:(NSString *)method URI:(NSString *)path
{
	HTTPLogTrace();
	
	if ([method isEqualToString:@"POST"] && [path isEqualToString:@"/activeurl"])
	{
		HTTPLogVerbose(@"%@[%p]: postContentLength: %qu", THIS_FILE, self, requestContentLength);
		
		NSString *postStr = nil;
		
		NSData *postData = [request body];
		if (postData)
		{
			postStr = [[NSString alloc] initWithData:postData encoding:NSUTF8StringEncoding];
		}
		
		HTTPLogVerbose(@"%@[%p]: Received URL is: %@", THIS_FILE, self, postStr);
		
        NSError *e = nil;
        NSDictionary* urlInfo = [NSJSONSerialization JSONObjectWithData:postData options:NSJSONReadingMutableContainers error:&e];
        
        
        
        NSMutableDictionary* userInfo = [NSMutableDictionary dictionaryWithCapacity:2];
        
       //[userInfo setObject:<#(id)#> forKey:<#(id)#>]
        
        [userInfo setObject:[urlInfo objectForKey:@"url"] forKey:@"url"];
        [userInfo setObject:[urlInfo objectForKey:@"title"] forKey:@"title"];
        
        [[NSNotificationCenter defaultCenter]
         postNotificationName:@"URLNotification"
         object:self
         userInfo:userInfo];
        
		// Result will be of the form "answer=..."
		
		//int answer = [[postStr substringFromIndex:7] intValue];
		
		NSData *response = nil;
		response = [@"<html><body><body></html>" dataUsingEncoding:NSUTF8StringEncoding];
		return [[HTTPDataResponse alloc] initWithData:response];
	}
    else if([method isEqualToString:@"OPTIONS"] && [path isEqualToString:@"/activeurl"]) {
        NSData *response = nil;
        response = [@"<html><body><body></html>" dataUsingEncoding:NSUTF8StringEncoding];
        HTTPDataResponse* httpResponseNow = [[HTTPDataResponse alloc] initWithData:response];
        [httpResponseNow setOptionsHeader];
        return httpResponseNow;
	}
    else if([method isEqualToString:@"POST"] && [path isEqualToString:@"/connection"]) {
        HTTPLogVerbose(@"%@[%p]: postContentLength: %qu", THIS_FILE, self, requestContentLength);
		
		NSString *postStr = nil;
		
		NSData *postData = [request body];
		if (postData)
		{
			postStr = [[NSString alloc] initWithData:postData encoding:NSUTF8StringEncoding];
		}
		
		HTTPLogVerbose(@"%@[%p]: Received data is: %@", THIS_FILE, self, postData);
        
        NSData *response = nil;
		response = [@"<html><body><body></html>" dataUsingEncoding:NSUTF8StringEncoding];
		return [[HTTPDataResponse alloc] initWithData:response];
        
    }
    else if ([method isEqualToString:@"POST"] && [path isEqualToString:@"/connectionurl"])
	{
		HTTPLogVerbose(@"%@[%p]: postContentLength: %qu", THIS_FILE, self, requestContentLength);
		
		NSString *postStr = nil;
		
		NSData *postData = [request body];
		if (postData)
		{
			postStr = [[NSString alloc] initWithData:postData encoding:NSUTF8StringEncoding];
		}
		
		HTTPLogVerbose(@"%@[%p]: Received URL is: %@", THIS_FILE, self, postStr);
		
        NSError *e = nil;
        NSDictionary* urlInfo = [NSJSONSerialization JSONObjectWithData:postData options:NSJSONReadingMutableContainers error:&e];
        
        
        
        NSMutableDictionary* userInfo = [NSMutableDictionary dictionaryWithCapacity:2];
        
        NSLog(@"Received URL is %@", );
        
        [userInfo setObject:[urlInfo objectForKey:@"url"] forKey:@"url"];
        [userInfo setObject:[urlInfo objectForKey:@"title"] forKey:@"title"];
        

		
		NSData *response = nil;
		response = [@"<html><body><body></html>" dataUsingEncoding:NSUTF8StringEncoding];
		return [[HTTPDataResponse alloc] initWithData:response];
	}
	return [super httpResponseForMethod:method URI:path];
}

- (void)prepareForBodyWithSize:(UInt64)contentLength
{
	HTTPLogTrace();
	
	// If we supported large uploads,
	// we might use this method to create/open files, allocate memory, etc.
    NSString* boundary = [request headerField:@"boundary"];
    if(!(boundary==nil)) {
        parser = [[MultipartFormDataParser alloc] initWithBoundary:boundary formEncoding:NSUTF8StringEncoding];
        parser.delegate = self;
        
        uploadedFiles = [[NSMutableArray alloc] init];
    }
    
}

- (void)processBodyData:(NSData *)postDataChunk
{
	HTTPLogTrace();
	
	// Remember: In order to support LARGE POST uploads, the data is read in chunks.
	// This prevents a 50 MB upload from being stored in RAM.
	// The size of the chunks are limited by the POST_CHUNKSIZE definition.
	// Therefore, this method may be called multiple times for the same POST request.
    BOOL result;
    NSString* boundary = [request headerField:@"boundary"];
    if(!(boundary==nil)) {
        result = [parser appendData:postDataChunk];
    }
    else {
        result = [request appendData:postDataChunk];
    }
    
    
	if (!result)
	{
		HTTPLogError(@"%@[%p]: %@ - Couldn't append bytes!", THIS_FILE, self, THIS_METHOD);
	}
}

//-----------------------------------------------------------------
#pragma mark multipart form data parser delegate


- (void) processStartOfPartWithHeader:(MultipartMessageHeader*) header {
	// in this sample, we are not interested in parts, other then file parts.
	// check content disposition to find out filename
    
    MultipartMessageHeaderField* disposition = [header.fields objectForKey:@"Content-Disposition"];
	NSString* filename = [[disposition.params objectForKey:@"filename"] lastPathComponent];
    
    if ( (nil == filename) || [filename isEqualToString: @""] ) {
        // it's either not a file part, or
		// an empty form sent. we won't handle it.
		return;
	}
	NSString* uploadDirPath = [NSHomeDirectory() stringByAppendingPathComponent:@"upload"];
    
	BOOL isDir = YES;
	if (![[NSFileManager defaultManager]fileExistsAtPath:uploadDirPath isDirectory:&isDir ]) {
		[[NSFileManager defaultManager]createDirectoryAtPath:uploadDirPath withIntermediateDirectories:YES attributes:nil error:nil];
	}
	
    NSString* filePath = [uploadDirPath stringByAppendingPathComponent: filename];
    if( [[NSFileManager defaultManager] fileExistsAtPath:filePath] ) {
        storeFile = nil;
    }
    else {
		HTTPLogVerbose(@"Saving file to %@", filePath);
		if(![[NSFileManager defaultManager] createDirectoryAtPath:uploadDirPath withIntermediateDirectories:true attributes:nil error:nil]) {
			HTTPLogError(@"Could not create directory at path: %@", filePath);
		}
		if(![[NSFileManager defaultManager] createFileAtPath:filePath contents:nil attributes:nil]) {
			HTTPLogError(@"Could not create file at path: %@", filePath);
		}
		storeFile = [NSFileHandle fileHandleForWritingAtPath:filePath];
		[uploadedFiles addObject: [NSString stringWithFormat:@"/upload/%@", filename]];
        [[NSWorkspace sharedWorkspace] openFile:filePath];

    }
}


- (void) processContent:(NSData*) data WithHeader:(MultipartMessageHeader*) header
{
	// here we just write the output from parser to the file.
	if( storeFile ) {
		[storeFile writeData:data];
	}
}

- (void) processEndOfPartWithHeader:(MultipartMessageHeader*) header
{
	// as the file part is over, we close the file.
	[storeFile closeFile];
	storeFile = nil;
}

- (void) processPreambleData:(NSData*) data
{
    // if we are interested in preamble data, we could process it here.
    
}

- (void) processEpilogueData:(NSData*) data
{
    // if we are interested in epilogue data, we could process it here.
    
}


@end
