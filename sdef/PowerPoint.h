#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>

@class PowerPointApplication, PowerPointPresentation;

typedef enum
{
    PowerPointPPffSaveAsPresentation = '\000\314\000\001',
    PowerPointPPffSaveAsTemplate = '\000\314\000\005',
    PowerPointPPffSaveAsRTF = '\000\314\000\006',
    PowerPointPPffSaveAsShow = '\000\314\000\007',
    PowerPointPPffSaveAsDefault = '\000\314\000\012',
    PowerPointPPffSaveAsHTML = '\000\314\000\013',
    PowerPointPPffSaveAsMovie = '\000\314\000\014',
    PowerPointPPffSaveAsPackage = '\000\314\000\015',
    PowerPointPPffSaveAsPDF = '\000\314\000\016',
    PowerPointPPffSaveAsOpenXMLPresentation = '\000\314\000\017',
    PowerPointPPffSaveAsOpenXMLPresentationMacroEnabled = '\000\314\000\020',
    PowerPointPPffSaveAsOpenXMLShow = '\000\314\000\021',
    PowerPointPPffSaveAsOpenXMLShowMacroEnabled = '\000\314\000\022',
    PowerPointPPffSaveAsOpenXMLTemplate = '\000\314\000\023',
    PowerPointPPffSaveAsOpenXMLTemplateMacroEnabled = '\000\314\000\024',
    PowerPointPPffSaveAsOpenXMLTheme = '\000\314\000\025',
    PowerPointPPffSaveAsGIF = '\000\314\000\026',
    PowerPointPPffSaveAsJPG = '\000\314\000\027',
    PowerPointPPffSaveAsPNG = '\000\314\000\030',
    PowerPointPPffSaveAsBMP = '\000\314\000\031',
    PowerPointPPffSaveAsTIF = '\000\314\000\032'
} PowerPointPPff;

typedef enum PowerPointSavo
{
    PowerPointSavoYes = 'yes ',
    PowerPointSavoNo = 'no  ',
    PowerPointSavoAsk = 'ask '
} PowerPointSavo;

@interface PowerPointApplication : SBApplication

@property BOOL startUpDialog;

@property (copy, readonly) NSString *Version;

- (SBElementArray *) presentations;

- (void) open:(NSString *)path
     password:(NSString *)password
writePassword:(NSString *)writePassword;

@end

@interface PowerPointPresentation : SBObject

@property (copy, readonly) NSString *fullName;

- (void) saveIn:(NSURL *)path as:(PowerPointPPff)as;
- (void) closeSaving:(PowerPointSavo)saving savingIn:(NSURL *)path;

@end
