//
//  RecentItemsDataController.m
//  SeamlessClientMac
//
//  Created by Sumit Pasupalak on 9/9/13.
//  Copyright (c) 2013 Sumit Pasupalak. All rights reserved.
//

#import "RecentItemsDataController.h"

@interface RecentItemsDataController ()
-(void) initializeDefaultRecentItemsList;
@end
@implementation RecentItemsDataController

-(void) initializeDefaultRecentItemsList {
    NSMutableArray* recentItemsList = [[NSMutableArray alloc] init];
    self.recentMasterItemsList = recentItemsList;
}

-(void) setRecentMasterItemsList:(NSMutableArray *)newList {
    if (_recentMasterItemsList != newList) {
        _recentMasterItemsList = [newList mutableCopy];
    }
}

- (id)init {
    if (self = [super init]) {
        [self initializeDefaultRecentItemsList];
        return self;
    }
    return nil;
}

-(NSUInteger) countOfRecentMasterItemsList {
    return [self.recentMasterItemsList count];
}

-(RecentItem*) objectInListAtIndex:(NSUInteger)theIndex {
    return [self.recentMasterItemsList objectAtIndex:theIndex];
}

- (void)addItem:(RecentItem *)item {
    [self.recentMasterItemsList addObject:item];
}

- (RecentItem*) fileWithNameExists:(NSString*) name {
    if(name == nil) {
        return nil;
    }
    NSError* error = nil;
    NSRegularExpression *regex = [NSRegularExpression
                                  regularExpressionWithPattern:@"(.+)\\%20\\(.*page.*\\)"
                                  options:NSRegularExpressionCaseInsensitive
                                  error:&error];
    
    NSTextCheckingResult *textCheckingResult = [regex firstMatchInString:name options:0 range:NSMakeRange(0, name.length)];
    
    NSRange matchRange = [textCheckingResult rangeAtIndex:1];
    NSString *match = [name substringWithRange:matchRange];
    if(![match isEqualTo:@""]) {
        NSLog(@"Found string '%@'", match);
        name = match;
    }
    
    
    for (RecentItem* item in self.recentMasterItemsList) {
       
        if( ([[item name] isEqualToString:name]) || ([name isEqualToString:[[item name] stringByDeletingPathExtension]])) {
            return item;
        }
    }
    return nil;
}

@end


