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
		if ([path isEqualToString:@"/activeurl"])
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
	
	if([method isEqualToString:@"POST"])
		return YES;
	
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
	return [super httpResponseForMethod:method URI:path];
}

- (void)prepareForBodyWithSize:(UInt64)contentLength
{
	HTTPLogTrace();
	
	// If we supported large uploads,
	// we might use this method to create/open files, allocate memory, etc.
}

- (void)processBodyData:(NSData *)postDataChunk
{
	HTTPLogTrace();
	
	// Remember: In order to support LARGE POST uploads, the data is read in chunks.
	// This prevents a 50 MB upload from being stored in RAM.
	// The size of the chunks are limited by the POST_CHUNKSIZE definition.
	// Therefore, this method may be called multiple times for the same POST request.
	
	BOOL result = [request appendData:postDataChunk];
	if (!result)
	{
		HTTPLogError(@"%@[%p]: %@ - Couldn't append bytes!", THIS_FILE, self, THIS_METHOD);
	}
}

@end
