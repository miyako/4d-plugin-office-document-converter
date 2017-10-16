#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>

@class ExcelApplication, ExcelWorkbook;

typedef enum ExcelSavo {
    ExcelSavoYes = 'yes ',
    ExcelSavoNo = 'no  ',
    ExcelSavoAsk = 'ask '
} ExcelSavo;

typedef enum
{
    ExcelMAScMsoAutomationSecurityLow = '\000\243\000\001',
    ExcelMAScMsoAutomationSecurityByUI = '\000\243\000\002',
    ExcelMAScMsoAutomationSecurityForceDisable = '\000\243\000\003'
} ExcelMASc;

typedef enum
{
    ExcelE188CSVFileFormat = '\002\274\000\006',
    ExcelE188CSVMacFileFormat = '\002\274\000\026',
    ExcelE188CSVMSDosFileFormat = '\002\274\000\030',
    ExcelE188CSVWindowsFileFormat = '\002\274\000\027',
    ExcelE188DBF3FileFormat = '\002\274\000\010',
    ExcelE188DBF4FileFormat = '\002\274\000\013',
    ExcelE188DIFFileFormat = '\002\274\000\011',
    ExcelE188Excel2FileFormat = '\002\274\000\020',
    ExcelE188Excel2EastAsianFileFormat = '\002\274\000\033',
    ExcelE188Excel3FileFormat = '\002\274\000\035',
    ExcelE188Excel4FileFormat = '\002\274\000!',
    ExcelE188Excel5FileFormat = '\002\274\000\'',
    ExcelE188Excel7FileFormat = '\002\274\000\'',
    ExcelE188Excel4WorkbookFileFormat = '\002\274\000#',
    ExcelE188InternationalAddInFileFormat = '\002\274\000\032',
    ExcelE188InternationalMacroFileFormat = '\002\274\000\031',
    ExcelE188WorkbookNormalFileFormat = '\002\273\357\321',
    ExcelE188SYLKFileFormat = '\002\274\000\002',
    ExcelE188CurrentPlatformTextFileFormat = '\002\273\357\302',
    ExcelE188TextMacFileFormat = '\002\274\000\023',
    ExcelE188TextMSDosFileFormat = '\002\274\000\025',
    ExcelE188TextPrinterFileFormat = '\002\274\000$',
    ExcelE188TextWindowsFileFormat = '\002\274\000\024',
    ExcelE188HTMLFileFormat = '\002\274\000,',
    ExcelE188XMLSpreadsheetFileFormat = '\002\274\000-',
    ExcelE188PDFFileFormat = '\002\274\000.',
    ExcelE188ExcelBinaryFileFormat = '\002\274\0003',
    ExcelE188ExcelXMLFileFormat = '\002\274\0004',
    ExcelE188MacroEnabledXMLFileFormat = '\002\274\0005',
    ExcelE188MacroEnabledTemplateFileFormat = '\002\274\0006',
    ExcelE188TemplateFileFormat = '\002\274\0007',
    ExcelE188AddInFileFormat = '\002\274\0008',
    ExcelE188Excel98to2004FileFormat = '\002\274\0009',
    ExcelE188Excel98to2004TemplateFileFormat = '\002\274\000\021',
    ExcelE188Excel98to2004AddInFileFormat = '\002\274\000\022'
} ExcelE188;

typedef enum
{
    ExcelE211Macintosh = '\002c\000\001',
    ExcelE211MSDos = '\002c\000\003',
    ExcelE211MSWindows = '\002c\000\002'
} ExcelE211;

typedef enum
{
    ExcelE221Exclusive = '\002m\000\003',
    ExcelE221NoChange = '\002m\000\001',
    ExcelE221Shared = '\002m\000\002'
} ExcelE221;

typedef enum
{
    ExcelE222LocalSessionChanges = '\002n\000\002',
    ExcelE222OtherSessionChanges = '\002n\000\003',
    ExcelE222UserResolution = '\002n\000\001'
} ExcelE222;

typedef enum
{
    ExcelE294DoNotUpdateLinks = '\002\266\000\000',
    ExcelE294UpdateExternalLinksOnly = '\002\266\000\001',
    ExcelE294UpdateRemoteLinksOnly = '\002\266\000\002',
    ExcelE294UpdateRemoteAndExternalLinks = '\002\266\000\003'
} ExcelE294;

typedef enum
{
    ExcelE295TabDelimiter = '\002\267\000\001',
    ExcelE295CommasDelimiter = '\002\267\000\002',
    ExcelE295SpacesDelimiter = '\002\267\000\003',
    ExcelE295SemicolonDelimiter = '\002\267\000\004',
    ExcelE295NoDelimiter = '\002\267\000\005',
    ExcelE295CustomCharacterDelimiter = '\002\267\000\006'
} ExcelE295;

@interface ExcelApplication : SBApplication

- (SBElementArray *) workbooks;

@property ExcelMASc automationSecurity;
@property BOOL startupDialog;
@property BOOL displayAlerts;
@property BOOL alertBeforeOverwriting;

@property (readonly) NSInteger calculationVersion;

- (ExcelWorkbook *) openWorkbookWorkbookFileName:(NSString *)workbookFileName
                                     updateLinks:(ExcelE294)updateLinks
                                        readOnly:(BOOL)readOnly
                                          format:(ExcelE295)format
                                        password:(NSString *)password
                           writeReservedPassword:(NSString *)writeReservedPassword
                       ignoreReadOnlyRecommended:(BOOL)ignoreReadOnlyRecommended
                                          origin:(ExcelE211)origin
                                       delimiter:(NSString *)delimiter
                                        editable:(BOOL)editable
                                          notify:(BOOL)notify
                                       converter:(NSInteger)converter
                                        addToMru:(BOOL)addToMru;

@end

@interface ExcelWorkbook : SBObject

@property (copy, readonly) NSString *fullName;

- (void) saveWorkbookAsFilename:(NSString *)filename
                     fileFormat:(ExcelE188)fileFormat
                       password:(NSString *)password
       writeReservationPassword:(NSString *)writeReservationPassword
            readOnlyRecommended:(BOOL)readOnlyRecommended
                   createBackup:(BOOL)createBackup
                     accessMode:(ExcelE221)accessMode
             conflictResolution:(ExcelE222)conflictResolution
      addToMostRecentlyUsedList:(BOOL)addToMostRecentlyUsedList
                      overwrite:(BOOL)overwrite;

- (void) saveIn:(NSURL *)path as:(ExcelE188)as;

- (void) closeSaving:(ExcelSavo)saving
            savingIn:(NSURL *)savingIn;

@end
