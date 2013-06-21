//
//  OCScriptingBridgeExcelController.m
//  Automation Framework
//
//  Created by Philip Regan on 12/15/10.
//  Copyright 2010 Philip Regan. All rights reserved.
//

#import "OCScriptingBridgeExcelController.h"

@interface OCExcelSBHelperObj ()
@property (readwrite, nonatomic, strong) excel2011Application *excelApp;
@end

@implementation OCExcelSBHelperObj

@synthesize excelApp = _excelApp;

#pragma mark -
#pragma mark Class Management Methods
#pragma mark -

- (id) init
{
	self = [super init];
	if (self != nil) {
		self.excelApp = [SBApplication applicationWithBundleIdentifier:@"com.microsoft.Excel"];
		if (!self.excelApp) {
			// TO DO: Handle the error
		}
	}
	return self;
}

- (NSDictionary *)testClass {
	NSMutableDictionary *testResults = [NSDictionary dictionary];
	
	if (!self.excelApp) {
		[testResults setObject:kOC_Error_SrptgBrdg_ExcelAppNotFoundValue forKey:kOC_Error_SrptgBrdg_ExcelAppNotFoundKey];
	}
	
	return testResults;
}

#pragma mark -
#pragma mark Automation Methods
#pragma mark -

#pragma mark File Management

- (excel2011Workbook *) openWorkbookAtPath:(NSString *)hfsPath {
		// filepath is an HFS file path (colon-based), not a POSIX path.
	excel2011Workbook *workbook = [self.excelApp openWorkbookWorkbookFileName:hfsPath 
															 updateLinks:excel2011E294DoNotUpdateLinks 
																readOnly:NO 
																  format:excel2011E295NoDelimiter 
																password:[NSString stringWithFormat:@""] 
												   writeReservedPassword:[NSString stringWithFormat:@""] 
											   ignoreReadOnlyRecommended:NO 
																  origin:excel2011E211Macintosh 
															   delimiter:nil
																editable:YES 
																  notify:NO 
															   converter:0
																addToMru:NO];
	return workbook;
}

- (void) closeWorkbook:(excel2011Workbook *)workbook {
        
    [workbook closeSaving:NO savingIn:excel2011XLfdPosixPath];
     
}

- (excel2011Sheet *) getSheetInWorkbook:(excel2011Workbook *)workbook withName:(NSString *)sheetName {
	if ([[[workbook sheets] objectWithName:sheetName] exists]) {
		return [[workbook sheets] objectWithName:sheetName];
	}
	return nil;
}

- (excel2011Sheet *) getSheetInWorkbook:(excel2011Workbook *)workbook atIndex:(int)sheetIndex {
	NSUInteger sheetCount = [[workbook sheets] count];
	if (sheetIndex <= sheetCount - 1) {
		return [[workbook sheets] objectAtLocation:[NSNumber numberWithInt:sheetIndex]];
	}
	return nil;
}

#pragma mark Content Management

- (int) getItemCountInColumn:(NSString *)keyColumn removeHeaderRow:(BOOL)headerRowExists inSheet:(excel2011Sheet *)sheet {
	int rowNumber = 0;
	NSString *theText = @"cSyoyodylg";
	while (![theText isEqualToString:@""]) {
		rowNumber++;
		
		// create the range
		NSString *rangeString = [NSString stringWithFormat:@"%@%i:%@%i", keyColumn, rowNumber, keyColumn, rowNumber];
		theText = [self getStringValueInCell:rangeString inSheet:sheet];
	}
	if (headerRowExists) {
		return rowNumber - 1;
	}
	return rowNumber;
}

- (NSArray *) getValuesInColumn:(NSString *)columnLetter fromFirstRow:(NSNumber *)firstRow toLastRow:(NSNumber *)lastRow inSheet:(excel2011Sheet *)sheet {
	NSString *rangeString = [NSString stringWithFormat:@"%@%i:%@%i", columnLetter, [firstRow intValue], columnLetter, [lastRow intValue]];
	excel2011Range *targetRange = [self createRange:rangeString inSheet:sheet inExcel:self.excelApp];
	if (!targetRange) {
		return nil;
	}
	NSArray *cellValues = [targetRange.value get];
    return [self deitemizeColumnArray:cellValues];
}

- (BOOL) setString:(NSString *)value inCell:(NSString *)cellColumnRow inSheet:(excel2011Sheet *)sheet {
    
    // get the range for the coordinates
    excel2011Range *range = [self createRange:[NSString stringWithFormat:@"%@:%@", cellColumnRow, cellColumnRow] 
                                      inSheet:sheet 
                                      inExcel:self.excelApp];
    
    if (!range) {
        return NO;
    }
    // set the value for that range
    range.value = value;
    
    return YES;
}

- (NSArray *)deitemizeColumnArray:(NSArray *)array {
    NSMutableArray *bufferArray = [NSMutableArray array];
    for ( NSUInteger thisItem = 0, lastItem = [array count]; thisItem < lastItem ; thisItem++ ) {
        [bufferArray addObject:[(NSArray *)[array objectAtIndex:thisItem] objectAtIndex:0]];
    }
    return bufferArray;
}

#pragma mark Object Creation

- (excel2011Range *) createRange:(NSString *)aRange inSheet:(excel2011Sheet *)sheet inExcel:(excel2011Application *)excel {
	excel2011Range *newRange = [[[excel classForScriptingClass:@"range"] alloc] initWithProperties:[NSDictionary dictionaryWithObjectsAndKeys:aRange, @"name", nil]];
	[[sheet ranges] addObject:newRange];
	excel2011Range *currentRange;
	if ([[[sheet ranges] objectWithName:aRange] exists]) {
		currentRange = [[sheet ranges] objectWithName:aRange];
	}
	return currentRange;
}

@end
