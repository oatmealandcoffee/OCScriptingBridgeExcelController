//
//  OCScriptingBridgeExcelController.h
//  Automation Framework
//
//  Created by Philip Regan on 12/15/10.
//  Copyright 2010 Philip Regan. All rights reserved.
//

#import <Cocoa/Cocoa.h>
#import <ScriptingBridge/ScriptingBridge.h>

#import "excel2011.h"


@interface OCExcelSBHelperObj : NSObject 

#pragma mark -
#pragma mark Class Management Methods
#pragma mark -

- (NSDictionary *)testClass;

#pragma mark -
#pragma mark Automation Methods
#pragma mark -

#pragma mark File Management

/*
 Useful snippets (assuming this class is instantiated as 'excel')
 
 // get the first sheet in a selected workbook
 excel2011Workbook *workbook = [excel openWorkbookAtPath:path];
 excel2011Sheet *worksheet = [excel getSheetInWorkbook:workbook atIndex:0];
 */

- (excel2011Workbook *) openWorkbookAtPath:(NSString *)hfsPath;
- (void) closeWorkbook:(excel2011Workbook *)workbook;
- (excel2011Sheet *) getSheetInWorkbook:(excel2011Workbook *)workbook withName:(NSString *)sheetName;
- (excel2011Sheet *) getSheetInWorkbook:(excel2011Workbook *)workbook atIndex:(int)sheetIndex;

#pragma mark Content Management

/* GET */

- (int) getItemCountInColumn:(NSString *)keyColumn removeHeaderRow:(BOOL)headerRowExists inSheet:(excel2011Sheet *)sheet;
- (NSArray *) getValuesInColumn:(NSString *)columnLetter fromFirstRow:(NSNumber *)firstRow toLastRow:(NSNumber *)lastRow inSheet:(excel2011Sheet *)sheet;
- (NSString *) getStringValueInCell:(NSString *)cellColumnRow inSheet:(excel2011Sheet *)sheet;

/* SET */

- (BOOL) setString:(NSString *)value inCell:(NSString *)cellColumnRow inSheet:(excel2011Sheet *)sheet;

#pragma mark Object Creation

- (excel2011Range *) createRange:(NSString *)aRange inSheet:(excel2011Sheet *)sheet inExcel:(excel2011Application *)excel;

@end
