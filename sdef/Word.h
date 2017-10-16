#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>

@class WordApplication, WordDocument;

typedef enum
{
    WordE161FormatDocument97 = '\0021\000\000',
    WordE161FormatTemplate97 = '\0021\000\001',
    WordE161FormatText = '\0021\000\002',
    WordE161FormatTextLineBreaks = '\0021\000\003',
    WordE161FormatDostext = '\0021\000\004',
    WordE161FormatDostextLineBreaks = '\0021\000\005',
    WordE161FormatRtf = '\0021\000\006',
    WordE161FormatUnicodeText = '\0021\000\007',
    WordE161FormatHTML = '\0021\000\010',
    WordE161FormatWebArchive = '\0021\000\011',
    WordE161FormatStationery = '\0021\000\012',
    WordE161FormatXml = '\0021\000\013',
    WordE161FormatDocument = '\0021\000\014',
    WordE161FormatDocumentME = '\0021\000\015',
    WordE161FormatTemplate = '\0021\000\016',
    WordE161FormatTemplateME = '\0021\000\017',
    WordE161FormatPDF = '\0021\000\020',
    WordE161FormatFlatDocument = '\0021\000\021',
    WordE161FormatFlatDocumentME = '\0021\000\022',
    WordE161FormatFlatTemplate = '\0021\000\023',
    WordE161FormatFlatTemplateME = '\0021\000\024',
    WordE161FormatCustomDictionary = '\0021\000\025',
    WordE161FormatExcludeDictionary = '\0021\000\026',
    WordE161FormatDocumentAuto = '\0021\015\014',
    WordE161FormatTemplateAuto = '\0021\015\007'
} WordE161;

typedef enum
{
    WordE162OpenFormatAuto = '\0022\000\000',
    WordE162OpenFormatDocument = '\0022\000\001',
    WordE162OpenFormatTemplate = '\0022\000\002',
    WordE162OpenFormatRtf = '\0022\000\003',
    WordE162OpenFormatText = '\0022\000\004',
    WordE162OpenFormatUnicodeText = '\0022\000\005',
    WordE162OpenFormatEncodedText = '\0022\000\005',
    WordE162OpenFormatMacReadable = '\0022\000\006',
    WordE162OpenFormatWebPages = '\0022\000\007',
    WordE162OpenFormatXml = '\0022\000\010',
    WordE162OpenFormatDocument97 = '\0022\000\011',
    WordE162OpenFormatTemplate97 = '\0022\000\012',
    WordE162OpenFormatOffice = '\0022\000\013'
} WordE162;

typedef enum
{
	WordSavoYes = 'yes ',
	WordSavoNo = 'no  ',
	WordSavoAsk = 'ask '
} WordSavo;

@interface WordApplication : SBApplication

@property BOOL startupDialog;
@property BOOL displayAlerts;

@property (copy, readonly) NSString *build;
@property (copy, readonly) NSString *applicationVersion;

- (SBElementArray *) documents;

- (void) openFileName:(NSString *)fileName
   confirmConversions:(BOOL)confirmConversions
             readOnly:(BOOL)readOnly
     addToRecentFiles:(BOOL)addToRecentFiles
     passwordDocument:(NSString *)passwordDocument
     passwordTemplate:(NSString *)passwordTemplate
               Revert:(BOOL)Revert
        writePassword:(NSString *)writePassword
writePasswordTemplate:(NSString *)writePasswordTemplate
        fileConverter:(WordE162)fileConverter;

@end

@interface WordDocument : SBObject

@property (copy, readonly) NSString *fullName;

- (void) saveIn:(NSURL *)path as:(WordE161)as;
- (void) closeSaving:(WordSavo)saving savingIn:(NSURL *)savingIn;

@end
