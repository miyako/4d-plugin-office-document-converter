#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>

#pragma mark enum

@class PowerPointBaseObject, PowerPointBaseApplication, PowerPointBaseDocument, PowerPointBasicWindow, PowerPointPrintSettings, PowerPointCommandBarControl, PowerPointCommandBarButton, PowerPointCommandBarCombobox, PowerPointCommandBarPopup, PowerPointCommandBar, PowerPointDocumentProperty, PowerPointCustomDocumentProperty, PowerPointWebPageFont, PowerPointActionSetting, PowerPointAddIn, PowerPointAnimationBehavior, PowerPointAnimationPoint, PowerPointAnimationSettings, PowerPointApplication, PowerPointAutocorrectEntry, PowerPointAutocorrect, PowerPointBroadcast, PowerPointBulletFormat, PowerPointColorScheme, PowerPointColorsEffect, PowerPointCommandEffect, PowerPointCustomLayout, PowerPointDefaultWebOptions, PowerPointDesign, PowerPointDocumentWindow, PowerPointEffectInformation, PowerPointEffectParameters, PowerPointEffect, PowerPointFilterEffect, PowerPointFirstLetterException, PowerPointFont, PowerPointHeaderOrFooter, PowerPointHeadersAndFooters, PowerPointHyperlink, PowerPointMaster, PowerPointMotionEffect, PowerPointNamedSlideShow, PowerPointPageSetup, PowerPointPane, PowerPointParagraphFormat, PowerPointPlaySettings, PowerPointPlayer, PowerPointPreferences, PowerPointPresentation, PowerPointPresenterTool, PowerPointPresenterViewWindow, PowerPointPrintOptions, PowerPointPrintRange, PowerPointPropertyEffect, PowerPointRotatingEffect, PowerPointRulerLevel, PowerPointRuler, PowerPointSaveAsMovieSettings, PowerPointScaleEffect, PowerPointSectionProperties, PowerPointSelection, PowerPointSequence, PowerPointSetEffect, PowerPointSlideRange, PowerPointSlideShowSettings, PowerPointSlideShowTransition, PowerPointSlideShowView, PowerPointSlideShowWindow, PowerPointSlide, PowerPointSoundEffect, PowerPointTabStop, PowerPointTextStyleLevel, PowerPointTextStyle, PowerPointTimeline, PowerPointTiming, PowerPointTwoInitialCapsException, PowerPointView, PowerPointWebOptions, PowerPointAdjustment, PowerPointCalloutFormat, PowerPointConnectorFormat, PowerPointFillFormat, PowerPointGlowFormat, PowerPointGradientStop, PowerPointLineFormat, PowerPointLinkFormat, PowerPointOfficeTheme, PowerPointPictureFormat, PowerPointPlaceholderFormat, PowerPointReflectionFormat, PowerPointShadowFormat, PowerPointShapeRange, PowerPointShape, PowerPointCallout, PowerPointComment, PowerPointConnector, PowerPointLineShape, PowerPointMediaObject, PowerPointMedia2Object, PowerPointPicture, PowerPointPlaceHolder, PowerPointShapeTable, PowerPointSoftEdgeFormat, PowerPointTextBox, PowerPointTextColumn, PowerPointTextFrame, PowerPointThemeColorScheme, PowerPointThemeColor, PowerPointThemeEffectScheme, PowerPointThemeFontScheme, PowerPointThemeFont, PowerPointMajorThemeFont, PowerPointMinorThemeFont, PowerPointThreeDFormat, PowerPointWordArtFormat, PowerPointTextRange, PowerPointCharacter, PowerPointLine, PowerPointParagraph, PowerPointSentence, PowerPointTextFlow, PowerPointWord, PowerPointCell, PowerPointColumn, PowerPointRow, PowerPointTable;

enum PowerPointSaveOption {
	PowerPointSaveOptionYes = 'yes ' /* Save objects now */,
	PowerPointSaveOptionNo = 'no  ' /* Do not save objects */,
	PowerPointSaveOptionAsk = 'ask ' /* Ask the user whether to save */
};
typedef enum PowerPointSaveOption PowerPointSaveOption;

enum WordSaveOption {
	WordSaveOptionYes = 'yes ' /* Save objects now */,
	WordSaveOptionNo = 'no  ' /* Do not save objects */,
	WordSaveOptionAsk = 'ask ' /* Ask the user whether to save */
};
typedef enum WordSaveOption WordSaveOption;

enum ExcelSaveOption {
	ExcelSaveOptionYes = 'yes ' /* Save objects now */,
	ExcelSaveOptionNo = 'no  ' /* Do not save objects */,
	ExcelSaveOptionAsk = 'ask ' /* Ask the user whether to save */
};
typedef enum ExcelSaveOption ExcelSaveOption;

#pragma mark -

enum PowerPointFilePath {
	PowerPointFilePathMacintoshPath = 'utxt' /* Maintosh path i.e. Foo:bar.txt */,
	PowerPointFilePathPosixPath = 'file' /* Posix path i.e. file:://foo/bar.txt */
};
typedef enum PowerPointFilePath PowerPointFilePath;

enum WordFilePath {
	WordFilePathMacintoshPath = 'utxt' /* Maintosh path i.e. Foo:bar.txt */,
	WordFilePathPosixPath = 'file' /* Posix path i.e. file:://foo/bar.txt */
};
typedef enum WordFilePath WordFilePath;


enum ExcelFilePath {
	ExcelFilePathMacintoshPath = 'mPth' /* Maintosh path i.e. Foo:bar.txt */,
	ExcelFilePathPosixPath = 'file' /* Posix path i.e. file:://foo/bar.txt */
};
typedef enum ExcelFilePath ExcelFilePath;

#pragma mark -

enum PowerPointFileFormat {
	PowerPointFileFormatSaveAsPresentation = '\000\314\000\001',
	PowerPointFileFormatSaveAsTemplate = '\000\314\000\005',
	PowerPointFileFormatSaveAsRTF = '\000\314\000\006',
	PowerPointFileFormatSaveAsShow = '\000\314\000\007',
	PowerPointFileFormatSaveAsDefault = '\000\314\000\012',
	PowerPointFileFormatSaveAsHTML = '\000\314\000\013',
	PowerPointFileFormatSaveAsMovie = '\000\314\000\014',
	PowerPointFileFormatSaveAsPackage = '\000\314\000\015',
	PowerPointFileFormatSaveAsPDF = '\000\314\000\016',
	PowerPointFileFormatSaveAsOpenXMLPresentation = '\000\314\000\017',
	PowerPointFileFormatSaveAsOpenXMLPresentationMacroEnabled = '\000\314\000\020',
	PowerPointFileFormatSaveAsOpenXMLShow = '\000\314\000\021',
	PowerPointFileFormatSaveAsOpenXMLShowMacroEnabled = '\000\314\000\022',
	PowerPointFileFormatSaveAsOpenXMLTemplate = '\000\314\000\023',
	PowerPointFileFormatSaveAsOpenXMLTemplateMacroEnabled = '\000\314\000\024',
	PowerPointFileFormatSaveAsOpenXMLTheme = '\000\314\000\025',
	PowerPointFileFormatSaveAsGIF = '\000\314\000\026',
	PowerPointFileFormatSaveAsJPG = '\000\314\000\027',
	PowerPointFileFormatSaveAsPNG = '\000\314\000\030',
	PowerPointFileFormatSaveAsBMP = '\000\314\000\031',
	PowerPointFileFormatSaveAsTIF = '\000\314\000\032'
};
typedef enum PowerPointFileFormat PowerPointFileFormat;

enum WordFileFormat {
	WordFileFormatFormatDocument97 = '\0021\000\000',
	WordFileFormatFormatTemplate97 = '\0021\000\001',
	WordFileFormatFormatText = '\0021\000\002',
	WordFileFormatFormatTextLineBreaks = '\0021\000\003',
	WordFileFormatFormatDostext = '\0021\000\004',
	WordFileFormatFormatDostextLineBreaks = '\0021\000\005',
	WordFileFormatFormatRtf = '\0021\000\006',
	WordFileFormatFormatUnicodeText = '\0021\000\007',
	WordFileFormatFormatHTML = '\0021\000\010',
	WordFileFormatFormatWebArchive = '\0021\000\011',
	WordFileFormatFormatStationery = '\0021\000\012',
	WordFileFormatFormatXml = '\0021\000\013',
	WordFileFormatFormatDocument = '\0021\000\014',
	WordFileFormatFormatDocumentME = '\0021\000\015',
	WordFileFormatFormatTemplate = '\0021\000\016',
	WordFileFormatFormatTemplateME = '\0021\000\017',
	WordFileFormatFormatPDF = '\0021\000\020',
	WordFileFormatFormatFlatDocument = '\0021\000\021',
	WordFileFormatFormatFlatDocumentME = '\0021\000\022',
	WordFileFormatFormatFlatTemplate = '\0021\000\023',
	WordFileFormatFormatFlatTemplateME = '\0021\000\024',
	WordFileFormatFormatCustomDictionary = '\0021\000\025',
	WordFileFormatFormatExcludeDictionary = '\0021\000\026',
	WordFileFormatFormatDocumentAuto = '\0021\015\014',
	WordFileFormatFormatTemplateAuto = '\0021\015\007'
};
typedef enum WordFileFormat WordFileFormat;

enum ExcelFileFormat {
	ExcelFileFormatCSVFileFormat = '\002\274\000\006',
	ExcelFileFormatCSVMacFileFormat = '\002\274\000\026',
	ExcelFileFormatCSVMSDosFileFormat = '\002\274\000\030',
	ExcelFileFormatCSVWindowsFileFormat = '\002\274\000\027',
	ExcelFileFormatDBF3FileFormat = '\002\274\000\010',
	ExcelFileFormatDBF4FileFormat = '\002\274\000\013',
	ExcelFileFormatDIFFileFormat = '\002\274\000\011',
	ExcelFileFormatExcel2FileFormat = '\002\274\000\020',
	ExcelFileFormatExcel2EastAsianFileFormat = '\002\274\000\033',
	ExcelFileFormatExcel3FileFormat = '\002\274\000\035',
	ExcelFileFormatExcel4FileFormat = '\002\274\000!',
	ExcelFileFormatExcel5FileFormat = '\002\274\000\'',
	ExcelFileFormatExcel7FileFormat = '\002\274\000\'',
	ExcelFileFormatExcel4WorkbookFileFormat = '\002\274\000#',
	ExcelFileFormatInternationalAddInFileFormat = '\002\274\000\032',
	ExcelFileFormatInternationalMacroFileFormat = '\002\274\000\031',
	ExcelFileFormatWorkbookNormalFileFormat = '\002\273\357\321',
	ExcelFileFormatSYLKFileFormat = '\002\274\000\002',
	ExcelFileFormatCurrentPlatformTextFileFormat = '\002\273\357\302',
	ExcelFileFormatTextMacFileFormat = '\002\274\000\023',
	ExcelFileFormatTextMSDosFileFormat = '\002\274\000\025',
	ExcelFileFormatTextPrinterFileFormat = '\002\274\000$',
	ExcelFileFormatTextWindowsFileFormat = '\002\274\000\024',
	ExcelFileFormatHTMLFileFormat = '\002\274\000,',
	ExcelFileFormatXMLSpreadsheetFileFormat = '\002\274\000-',
	ExcelFileFormatPDFFileFormat = '\002\274\000.',
	ExcelFileFormatExcelBinaryFileFormat = '\002\274\0003',
	ExcelFileFormatExcelXMLFileFormat = '\002\274\0004',
	ExcelFileFormatMacroEnabledXMLFileFormat = '\002\274\0005',
	ExcelFileFormatMacroEnabledTemplateFileFormat = '\002\274\0006',
	ExcelFileFormatTemplateFileFormat = '\002\274\0007',
	ExcelFileFormatAddInFileFormat = '\002\274\0008',
	ExcelFileFormatExcel98to2004FileFormat = '\002\274\0009',
	ExcelFileFormatExcel98to2004TemplateFileFormat = '\002\274\000\021',
	ExcelFileFormatExcel98to2004AddInFileFormat = '\002\274\000\022'
};
typedef enum ExcelFileFormat ExcelFileFormat;

#pragma mark -

enum PowerPointAutomationSecurity {
	PowerPointAutomationSecurityMsoAutomationSecurityLow = '\000\243\000\001',
	PowerPointAutomationSecurityMsoAutomationSecurityByUI = '\000\243\000\002',
	PowerPointAutomationSecurityMsoAutomationSecurityForceDisable = '\000\243\000\003'
};
typedef enum PowerPointAutomationSecurity PowerPointAutomationSecurity;

enum WordAutomationSecurity {
	WordAutomationSecurityAlertsNone = '\002\032\000\000',
	WordAutomationSecurityAlertsMessageBox = '\002\031\377\376',
	WordAutomationSecurityAlertsAll = '\002\031\377\377'
};
typedef enum WordAutomationSecurity WordAutomationSecurity;

enum ExcelAutomationSecurity {
	ExcelMAScMsoAutomationSecurityLow = '\000\243\000\001',
	ExcelMAScMsoAutomationSecurityByUI = '\000\243\000\002',
	ExcelMAScMsoAutomationSecurityForceDisable = '\000\243\000\003'
};
typedef enum ExcelAutomationSecurity ExcelAutomationSecurity;

#pragma mark -

enum WordE162 {
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
};
typedef enum WordE162 WordE162;

enum WordMtEn {
	WordMtEnEncodingThai = '\000\213\003j',
	WordMtEnEncodingJapaneseShiftJIS = '\000\213\003\244',
	WordMtEnEncodingSimplifiedChinese = '\000\213\003\250',
	WordMtEnEncodingKorean = '\000\213\003\265',
	WordMtEnEncodingBig5TraditionalChinese = '\000\213\003\266',
	WordMtEnEncodingLittleEndian = '\000\213\004\260',
	WordMtEnEncodingBigEndian = '\000\213\004\261',
	WordMtEnEncodingCentralEuropean = '\000\213\004\342',
	WordMtEnEncodingCyrillic = '\000\213\004\343',
	WordMtEnEncodingWestern = '\000\213\004\344',
	WordMtEnEncodingGreek = '\000\213\004\345',
	WordMtEnEncodingTurkish = '\000\213\004\346',
	WordMtEnEncodingHebrew = '\000\213\004\347',
	WordMtEnEncodingArabic = '\000\213\004\350',
	WordMtEnEncodingBaltic = '\000\213\004\351',
	WordMtEnEncodingVietnamese = '\000\213\004\352',
	WordMtEnEncodingISO88591Latin1 = '\000\213o\257',
	WordMtEnEncodingISO88592CentralEurope = '\000\213o\260',
	WordMtEnEncodingISO88593Latin3 = '\000\213o\261',
	WordMtEnEncodingISO88594Baltic = '\000\213o\262',
	WordMtEnEncodingISO88595Cyrillic = '\000\213o\263',
	WordMtEnEncodingISO88596Arabic = '\000\213o\264',
	WordMtEnEncodingISO88597Greek = '\000\213o\265',
	WordMtEnEncodingISO88598Hebrew = '\000\213o\266',
	WordMtEnEncodingISO88599Turkish = '\000\213o\267',
	WordMtEnEncodingISO885915Latin9 = '\000\213o\275',
	WordMtEnEncodingISO2022JapaneseNoHalfWidthKatakana = '\000\213\304,',
	WordMtEnEncodingISO2022JapaneseJISX02021984 = '\000\213\304-',
	WordMtEnEncodingISO2022JapaneseJISX02011989 = '\000\213\304.',
	WordMtEnEncodingISO2022KR = '\000\213\3041',
	WordMtEnEncodingISO2022CNTraditionalChinese = '\000\213\3043',
	WordMtEnEncodingISO2022CNSimplifiedChinese = '\000\213\3045',
	WordMtEnEncodingMacRoman = '\000\213\'\020',
	WordMtEnEncodingMacJapanese = '\000\213\'\021',
	WordMtEnEncodingMacTraditionalChinese = '\000\213\'\022',
	WordMtEnEncodingMacKorean = '\000\213\'\023',
	WordMtEnEncodingMacArabic = '\000\213\'\024',
	WordMtEnEncodingMacHebrew = '\000\213\'\025',
	WordMtEnEncodingMacGreek1 = '\000\213\'\026',
	WordMtEnEncodingMacCyrillic = '\000\213\'\027',
	WordMtEnEncodingMacSimplifiedChineseGB2312 = '\000\213\'\030',
	WordMtEnEncodingMacRomania = '\000\213\'\032',
	WordMtEnEncodingMacUkraine = '\000\213\'!',
	WordMtEnEncodingMacLatin2 = '\000\213\'-',
	WordMtEnEncodingMacIcelandic = '\000\213\'_',
	WordMtEnEncodingMacTurkish = '\000\213\'a',
	WordMtEnEncodingMacCroatia = '\000\213\'b',
	WordMtEnEncodingEBCDICUSCanada = '\000\213\000%',
	WordMtEnEncodingEBCDICInternational = '\000\213\001\364',
	WordMtEnEncodingEBCDICMultilingualROECELatin2 = '\000\213\003f',
	WordMtEnEncodingEBCDICGreekModern = '\000\213\003k',
	WordMtEnEncodingEBCDICTurkishLatin5 = '\000\213\004\002',
	WordMtEnEncodingEBCDICGermany = '\000\213O1',
	WordMtEnEncodingEBCDICDenmarkNorway = '\000\213O5',
	WordMtEnEncodingEBCDICFinlandSweden = '\000\213O6',
	WordMtEnEncodingEBCDICItaly = '\000\213O8',
	WordMtEnEncodingEBCDICLatinAmericaSpain = '\000\213O<',
	WordMtEnEncodingEBCDICUnitedKingdom = '\000\213O=',
	WordMtEnEncodingEBCDICJapaneseKatakanaExtended = '\000\213OB',
	WordMtEnEncodingEBCDICFrance = '\000\213OI',
	WordMtEnEncodingEBCDICArabic = '\000\213O\304',
	WordMtEnEncodingEBCDICGreek = '\000\213O\307',
	WordMtEnEncodingEBCDICHebrew = '\000\213O\310',
	WordMtEnEncodingEBCDICKoreanExtended = '\000\213Qa',
	WordMtEnEncodingEBCDICThai = '\000\213Qf',
	WordMtEnEncodingEBCDICIcelandic = '\000\213Q\207',
	WordMtEnEncodingEBCDICTurkish = '\000\213Q\251',
	WordMtEnEncodingEBCDICRussian = '\000\213Q\220',
	WordMtEnEncodingEBCDICSerbianBulgarian = '\000\213R!',
	WordMtEnEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = '\000\213\306\362',
	WordMtEnEncodingEBCDICUSCanadaAndJapanese = '\000\213\306\363',
	WordMtEnEncodingEBCDICExtendedAndKorean = '\000\213\306\365',
	WordMtEnEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = '\000\213\306\367',
	WordMtEnEncodingEBCDICUSCanadaAndTraditionalChinese = '\000\213\306\371',
	WordMtEnEncodingEBCDICJapaneseLatinExtendedAndJapanese = '\000\213\306\373',
	WordMtEnEncodingOEMUnitedStates = '\000\213\001\265',
	WordMtEnEncodingOEMGreek = '\000\213\002\341',
	WordMtEnEncodingOEMBaltic = '\000\213\003\007',
	WordMtEnEncodingOEMMultilingualLatinI = '\000\213\003R',
	WordMtEnEncodingOEMMultilingualLatinII = '\000\213\003T',
	WordMtEnEncodingOEMCyrillic = '\000\213\003W',
	WordMtEnEncodingOEMTurkish = '\000\213\003Y',
	WordMtEnEncodingOEMPortuguese = '\000\213\003\\',
	WordMtEnEncodingOEMIcelandic = '\000\213\003]',
	WordMtEnEncodingOEMHebrew = '\000\213\003^',
	WordMtEnEncodingOEMCanadianFrench = '\000\213\003_',
	WordMtEnEncodingOEMArabic = '\000\213\003`',
	WordMtEnEncodingOEMNordic = '\000\213\003a',
	WordMtEnEncodingOEMCyrillicII = '\000\213\003b',
	WordMtEnEncodingOEMModernGreek = '\000\213\003e',
	WordMtEnEncodingEUCJapanese = '\000\213\312\334',
	WordMtEnEncodingEUCChineseSimplifiedChinese = '\000\213\312\340',
	WordMtEnEncodingEUCKorean = '\000\213\312\355',
	WordMtEnEncodingEUCTaiwaneseTraditionalChinese = '\000\213\312\356',
	WordMtEnEncodingDevanagari = '\000\213\336\252',
	WordMtEnEncodingBengali = '\000\213\336\253',
	WordMtEnEncodingTamil = '\000\213\336\254',
	WordMtEnEncodingTelugu = '\000\213\336\255',
	WordMtEnEncodingAssamese = '\000\213\336\256',
	WordMtEnEncodingOriya = '\000\213\336\257',
	WordMtEnEncodingKannada = '\000\213\336\260',
	WordMtEnEncodingMalayalam = '\000\213\336\261',
	WordMtEnEncodingGujarati = '\000\213\336\262',
	WordMtEnEncodingPunjabi = '\000\213\336\263',
	WordMtEnEncodingArabicASMO = '\000\213\002\304',
	WordMtEnEncodingArabicTransparentASMO = '\000\213\002\320',
	WordMtEnEncodingKoreanJohab = '\000\213\005Q',
	WordMtEnEncodingTaiwanCNS = '\000\213N ',
	WordMtEnEncodingTaiwanTCA = '\000\213N!',
	WordMtEnEncodingTaiwanEten = '\000\213N\"',
	WordMtEnEncodingTaiwanIBM5550 = '\000\213N#',
	WordMtEnEncodingTaiwanTeletext = '\000\213N$',
	WordMtEnEncodingTaiwanWang = '\000\213N%',
	WordMtEnEncodingIA5IRV = '\000\213N\211',
	WordMtEnEncodingIA5German = '\000\213N\212',
	WordMtEnEncodingIA5Swedish = '\000\213N\213',
	WordMtEnEncodingIA5Norwegian = '\000\213N\214',
	WordMtEnEncodingUSASCII = '\000\213N\237',
	WordMtEnEncodingT61 = '\000\213O%',
	WordMtEnEncodingISO6937NonspacingAccent = '\000\213O-',
	WordMtEnEncodingKOI8R = '\000\213Q\202',
	WordMtEnEncodingExtAlphaLowercase = '\000\213R#',
	WordMtEnEncodingKOI8U = '\000\213Uj',
	WordMtEnEncodingEuropa3 = '\000\213qI',
	WordMtEnEncodingHZGBSimplifiedChinese = '\000\213\316\310',
	WordMtEnEncodingUTF7 = '\000\213\375\350',
	WordMtEnEncodingUTF8 = '\000\213\375\351'
};
typedef enum WordMtEn WordMtEn;

enum WordLineEndingType {
	WordLineEndingTypeLineEndingCrLf = '\002\307\000\000',
	WordLineEndingTypeLineEndingCrOnly = '\002\307\000\001',
	WordLineEndingTypeLineEndingLfOnly = '\002\307\000\002',
	WordLineEndingTypeLineEndingLfCr = '\002\307\000\003',
	WordLineEndingTypeLineEndingLsPs = '\002\307\000\004'
};

typedef enum WordLineEndingType WordLineEndingType;

enum ExcelUpdateLinksOption {
	ExcelUpdateLinksOptionDoNotUpdateLinks = '\002\266\000\000',
	ExcelUpdateLinksOptionUpdateExternalLinksOnly = '\002\266\000\001',
	ExcelUpdateLinksOptionUpdateRemoteLinksOnly = '\002\266\000\002',
	ExcelUpdateLinksOptionUpdateRemoteAndExternalLinks = '\002\266\000\003'
};
typedef enum ExcelUpdateLinksOption ExcelUpdateLinksOption;

enum ExcelFormat {
	ExcelFormatTabDelimiter = '\002\267\000\001',
	ExcelFormatCommasDelimiter = '\002\267\000\002',
	ExcelFormatSpacesDelimiter = '\002\267\000\003',
	ExcelFormatSemicolonDelimiter = '\002\267\000\004',
	ExcelFormatNoDelimiter = '\002\267\000\005',
	ExcelFormatCustomCharacterDelimiter = '\002\267\000\006'
};
typedef enum ExcelFormat ExcelFormat;

enum ExcelOrigin {
	ExcelOriginMacintosh = '\002c\000\001',
	ExcelOriginMSDos = '\002c\000\003',
	ExcelOriginMSWindows = '\002c\000\002'
};
typedef enum ExcelOrigin ExcelOrigin;

enum ExcelAccessMode {
	ExcelAccessModeExclusive = '\002m\000\003',
	ExcelAccessModeNoChange = '\002m\000\001',
	ExcelAccessModeShared = '\002m\000\002'
};
typedef enum ExcelAccessMode ExcelAccessMode;

enum ExcelConflictResolution {
	ExcelConflictResolutionLocalSessionChanges = '\002n\000\002',
	ExcelConflictResolutionOtherSessionChanges = '\002n\000\003',
	ExcelConflictResolutionUserResolution = '\002n\000\001'
};
typedef enum ExcelConflictResolution ExcelConflictResolution;

#pragma mark -

@protocol PowerPointGenericMethods

- (void) open:(NSString *)path password:(NSString *)password writePassword:(NSString *)writePassword;  // Open the specified object(s)
- (void) saveIn:(PowerPointFilePath)in_ as:(PowerPointFileFormat)as;  // Save an object
- (void) closeSaving:(PowerPointSaveOption)saving savingIn:(NSURL *)savingIn;  // Close an object

@end

#pragma mark -

@protocol WordGenericMethods

- (void) openFileName:(NSString *)fileName confirmConversions:(BOOL)confirmConversions readOnly:(BOOL)readOnly addToRecentFiles:(BOOL)addToRecentFiles passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate Revert:(BOOL)Revert writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate fileConverter:(WordE162)fileConverter;  // Open the specified object(s). Returns a reference to the opened file when using the file name parameter.
- (void) saveIn:(WordFilePath)in_ as:(WordFileFormat)as;  // Save an object
- (void) closeSaving:(WordSaveOption)saving savingIn:(NSURL *)savingIn;  // Close an object

@end

#pragma mark -

@protocol ExcelGenericMethods

- (void) open:(NSString *)path;  // Open the specified object(s)
- (void) saveIn:(ExcelFilePath)in_ as:(ExcelFileFormat)as;  // Save an object
- (void) closeSaving:(ExcelSaveOption)saving savingIn:(NSURL *)savingIn;  // Close an object

@end

#pragma mark -

#pragma mark PowerPoint

@interface PowerPointApplication : SBApplication

- (SBElementArray *) presentations;

- (PowerPointPresentation *) open:(NSString *)path
												 password:(NSString *)password
										writePassword:(NSString *)writePassword;  // Open the specified object(s)

#pragma mark -

@property (copy, readonly) NSString *Version;
@property (copy, readonly) NSString *build;
@property PowerPointAutomationSecurity automationSecurity;
@property BOOL startUpDialog;

@end

@interface PowerPointBaseObject : SBObject <PowerPointGenericMethods>

@property (copy) NSDictionary *properties;  // All of the object's properties

@end

@interface PowerPointPresentation : PowerPointBaseObject

@property (copy, readonly) NSString *path;
@property (copy, readonly) NSString *fullName;//full path

- (void) saveIn:(NSString *)fileName as:(PowerPointFileFormat)as; 

@end

#pragma mark -

#pragma mark Word

@class WordBaseObject, WordBaseApplication, WordBaseDocument, WordBasicWindow, WordPrintSettings, WordCommandBarControl, WordCommandBarButton, WordCommandBarCombobox, WordCommandBarPopup, WordCommandBar, WordDocumentProperty, WordCustomDocumentProperty, WordWebPageFont, WordWordComment, WordWordList, WordWordOptions, WordAddIn, WordApplication, WordAutoTextEntry, WordBookmark, WordBorderOptions, WordBorder, WordBrowser, WordBuildingBlockCategory, WordBuildingBlockType, WordBuildingBlock, WordCaptionLabel, WordCheckBox, WordCoauthLock, WordCoauthUpdate, WordCoauthor, WordCoauthoring, WordConflict, WordCustomLabel, WordDataMergeDataField, WordDataMergeDataSource, WordDataMergeFieldName, WordDataMergeField, WordDataMerge, WordDefaultWebOptions, WordDialog, WordDocumentVersion, WordDocument, WordDropCap, WordDropDown, WordEndnoteOptions, WordEndnote, WordEnvelope, WordFieldOptions, WordField, WordFileConverter, WordFind, WordFont, WordFootnoteOptions, WordFootnote, WordFormField, WordFrame, WordHeaderFooter, WordHeadingStyle, WordHyperlinkObject, WordIndex, WordKeyBinding, WordLetterContent, WordLineNumbering, WordLinkFormat, WordListEntry, WordListFormat, WordListGallery, WordListLevel, WordListTemplate, WordMailingLabel, WordMathAccent, WordMathAutocorrectEntry, WordMathAutocorrect, WordMathBar, WordMathBorderBox, WordMathBox, WordMathBreak, WordMathDelimiter, WordMathEquationArray, WordMathFraction, WordMathFunc, WordMathFunction, WordMathGroupChar, WordMathLeftScripts, WordMathLowerLimit, WordMathMatrixColumn, WordMathMatrixRow, WordMathMatrix, WordMathNary, WordMathObject, WordMathPhantom, WordMathRadical, WordMathRecognizedFunction, WordMathSubAndSuperScript, WordMathSubscript, WordMathSuperscript, WordMathUpperLimit, WordPageNumberOptions, WordPageNumber, WordPageSetup, WordPane, WordRangeEndnoteOptions, WordRangeFootnoteOptions, WordRecentFile, WordReplacement, WordReviewer, WordRevision, WordSelectionObject, WordSubdocument, WordSystemObject, WordTabStop, WordTableOfAuthorities, WordTableOfContents, WordTableOfFigures, WordTemplate, WordTextColumn, WordTextInput, WordTextRetrievalMode, WordVariable, WordView, WordWebOptions, WordWindow, WordZoom, WordAdjustment, WordCalloutFormat, WordFillFormat, WordGlowFormat, WordHorizontalLineFormat, WordInlineShape, WordInlineHorizontalLine, WordInlinePictureBullet, WordInlinePicture, WordLineFormat, WordOfficeTheme, WordPictureFormat, WordReflectionFormat, WordShadowFormat, WordShape, WordCallout, WordLineShape, WordPicture, WordSoftEdgeFormat, WordStandardInlineHorizontalLine, WordTextBox, WordTextFrame, WordThemeColorScheme, WordThemeColor, WordThemeEffectScheme, WordThemeFontScheme, WordThemeFont, WordMajorThemeFont, WordMinorThemeFont, WordThemeFonts, WordThreeDFormat, WordWordArtFormat, WordWordArt, WordWrapFormat, WordWordStyle, WordParagraphFormat, WordParagraph, WordSection, WordShading, WordTextRange, WordCharacter, WordGrammaticalError, WordSentence, WordSpellingError, WordStoryRange, WordWord, WordAutocorrectEntry, WordAutocorrect, WordDictionary, WordFirstLetterException, WordLanguage, WordOtherCorrectionsException, WordReadabilityStatistic, WordSynonymInfo, WordTwoInitialCapsException, WordCell, WordColumnOptions, WordColumn, WordCondition, WordRowOptions, WordRow, WordTableStyle, WordTable;

@interface WordApplication : SBApplication

- (SBElementArray *) documents;

#pragma mark -

@property (copy, readonly) NSString *applicationVersion;  // Returns the Microsoft Word version number as a string.
@property (copy, readonly) NSString *build;  // Returns the version and build number of the Microsoft Word application.
@property WordAutomationSecurity automationSecurity;
@property BOOL startupDialog;  // Returns of sets if the project gallery dialog will be shown when starting Microsoft Word.
@property BOOL displayAlerts;  // Returns or sets the way certain alerts or messages are handled while an AppleScript is running.

- (WordDocument *) openFileName:(NSString *)fileName
						 confirmConversions:(BOOL)confirmConversions
											 readOnly:(BOOL)readOnly
							 addToRecentFiles:(BOOL)addToRecentFiles
							 passwordDocument:(NSString *)passwordDocument
							 passwordTemplate:(NSString *)passwordTemplate
												 Revert:(BOOL)Revert
									writePassword:(NSString *)writePassword
					writePasswordTemplate:(NSString *)writePasswordTemplate
									fileConverter:(WordE162)fileConverter;  // Open the specified object(s). Returns a reference to the opened file when using the file name parameter.
@end

@interface WordBaseObject : SBObject <WordGenericMethods>

@property (copy) NSDictionary *properties;  // All of the object's properties

@end

@interface WordDocument : WordBaseObject

@property (copy, readonly) NSString *path;
@property (copy, readonly) NSString *fullName;  // Returns the full name of the document.

- (void) saveAsFileName:(NSString *)fileName
						 fileFormat:(WordFileFormat)fileFormat
					 lockComments:(BOOL)lockComments
							 password:(NSString *)password
			 addToRecentFiles:(BOOL)addToRecentFiles
					writePassword:(NSString *)writePassword
		readOnlyRecommended:(BOOL)readOnlyRecommended
		 embedTruetypeFonts:(BOOL)embedTruetypeFonts
saveNativePictureFormat:(BOOL)saveNativePictureFormat
					saveFormsData:(BOOL)saveFormsData
					 textEncoding:(WordMtEn)textEncoding
			 insertLineBreaks:(BOOL)insertLineBreaks
		 allowSubstitutions:(BOOL)allowSubstitutions
				 lineEndingType:(WordLineEndingType)lineEndingType
	HTMLDisplayOnlyOutput:(BOOL)HTMLDisplayOnlyOutput
	maintainCompatibility:(BOOL)maintainCompatibility;  // Saves the document with a new name or format.

@end

#pragma mark -

#pragma mark Excel

@class ExcelBaseObject, ExcelBaseApplication, ExcelBaseDocument, ExcelBasicWindow, ExcelPrintSettings, ExcelCommandBarControl, ExcelCommandBarButton, ExcelCommandBarCombobox, ExcelCommandBarPopup, ExcelCommandBar, ExcelDocumentProperty, ExcelCustomDocumentProperty, ExcelWebPageFont, ExcelExcelComment, ExcelODBCError, ExcelProtection, ExcelAboveAverageFormatCondition, ExcelAddIn, ExcelApplication, ExcelAutofilter, ExcelBorder, ExcelButton, ExcelCalculatedMember, ExcelCheckbox, ExcelColorScaleCriteria, ExcelColorScaleCriterion, ExcelColorScaleFormatCondition, ExcelColorstop, ExcelColorstops, ExcelConditionValue, ExcelCubeField, ExcelCustomView, ExcelDatabarBorder, ExcelDatabarFormatCondition, ExcelDefaultWebOptions, ExcelDialog, ExcelDisplayFormat, ExcelDropdown, ExcelFilter, ExcelFormatColor, ExcelFormatConditionIconObject, ExcelFormatConditionIconSet, ExcelFormatConditionIconSets, ExcelFormatCondition, ExcelGraphic, ExcelGroupbox, ExcelHorizontalPageBreak, ExcelHyperlink, ExcelIconCriteria, ExcelIconCriterion, ExcelIconSetFormatCondition, ExcelLabel, ExcelLinearGradient, ExcelListColumn, ExcelListObject, ExcelListRow, ExcelListbox, ExcelNamedItem, ExcelNegativeBarFormat, ExcelOptionButton, ExcelOutline, ExcelPageSetup, ExcelPane, ExcelPhonetic, ExcelPivotAxis, ExcelPivotCache, ExcelPivotCell, ExcelPivotField, ExcelCalculatedField, ExcelColumnField, ExcelDataField, ExcelHiddenField, ExcelPageField, ExcelPivotFilter, ExcelActiveFilter, ExcelPivotFormula, ExcelPivotItem, ExcelCalculatedItem, ExcelChildItem, ExcelColumnItem, ExcelHiddenItem, ExcelParentItem, ExcelPivotLine, ExcelPivotTable, ExcelQueryTable, ExcelRecentFile, ExcelRectangularGradient, ExcelRowField, ExcelRowItem, ExcelScenario, ExcelScrollbar, ExcelSheet, ExcelInternationalMacroSheet, ExcelMacroSheet, ExcelSlicer, ExcelSort, ExcelSortfield, ExcelSortfields, ExcelSpinner, ExcelTab, ExcelTableStyleElement, ExcelTableStyle, ExcelTextbox, ExcelTop10FormatCondition, ExcelTreeviewControl, ExcelUniqueValuesFormatCondition, ExcelValidation, ExcelValueChange, ExcelVerticalPageBreak, ExcelWebOptions, ExcelWindow, ExcelWorkbookConnection, ExcelWorkbook, ExcelDocument, ExcelWorksheet, ExcelXlspellingOptions, ExcelAdjustment, ExcelArc, ExcelBulletFormat, ExcelCalloutFormat, ExcelConnectorFormat, ExcelFillFormat, ExcelGlowFormat, ExcelGradientStop, ExcelLineFormat, ExcelLine, ExcelOfficeTheme, ExcelOval, ExcelParagraphFormat, ExcelPictureFormat, ExcelRectangle, ExcelReflectionFormat, ExcelRulerLevel, ExcelRuler, ExcelShadowFormat, ExcelShapeFont, ExcelShapeTextFrame, ExcelShape, ExcelCallout, ExcelPicture, ExcelShapeConnector, ExcelShapeLine, ExcelShapeTextbox, ExcelSoftEdgeFormat, ExcelTabStop, ExcelTextColumn, ExcelTextFrame, ExcelThemeColorScheme, ExcelThemeColor, ExcelThemeEffectScheme, ExcelThemeFontScheme, ExcelThemeFont, ExcelMajorThemeFont, ExcelMinorThemeFont, ExcelThreeDFormat, ExcelWordArtFormat, ExcelCharacter, ExcelFont, ExcelStyle, ExcelTextRange, ExcelParagraph, ExcelSentence, ExcelTextFlow, ExcelTextRangeCharacter, ExcelTextRangeLine, ExcelWord, ExcelRange, ExcelCell, ExcelColumn, ExcelRow, ExcelAutocorrect, ExcelAxisTitle, ExcelAxis, ExcelChartArea, ExcelChartFillFormat, ExcelChartFormat, ExcelChartGroup, ExcelAreaGroup, ExcelBarGroup, ExcelChartObject, ExcelChartTitle, ExcelChart, ExcelChartSheet, ExcelColumnGroup, ExcelCorners, ExcelDataLabel, ExcelDataTable, ExcelDisplayUnitLabel, ExcelDoughnutGroup, ExcelDownBars, ExcelDropLines, ExcelErrorBars, ExcelFloor, ExcelGridlines, ExcelHiloLines, ExcelInterior, ExcelLeaderLines, ExcelLegendEntry, ExcelLegendKey, ExcelLegend, ExcelLineGroup, ExcelPieGroup, ExcelPlotArea, ExcelRadarGroup, ExcelSeriesLines, ExcelSeriesPoint, ExcelSeries, ExcelTickLabels, ExcelTrendline, ExcelUpBars, ExcelWalls, ExcelXyGroup;

@interface ExcelApplication : SBApplication

- (SBElementArray *) workbooks;

#pragma mark -

@property (readonly) NSInteger calculationVersion;  // Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits, on the left, are the major version of Microsoft Excel.
@property (readonly) NSInteger build;  // Returns the Microsoft Excel build number.
@property ExcelAutomationSecurity automationSecurity;
@property BOOL startupDialog;  // Returns or sets if the startup dialog should be shown when starting up the application.
@property BOOL displayAlerts;  // Returns or sets if Microsoft Excel displays certain alerts and messages while handling events from AppleScript.
@property BOOL alertBeforeOverwriting;  // Returns or sets if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation.

- (ExcelWorkbook *) openWorkbookWorkbookFileName:(NSString *)workbookFileName
																		 updateLinks:(ExcelUpdateLinksOption)updateLinks
																				readOnly:(BOOL)readOnly
																					format:(ExcelFormat)format
																				password:(NSString *)password
													 writeReservedPassword:(NSString *)writeReservedPassword
											 ignoreReadOnlyRecommended:(BOOL)ignoreReadOnlyRecommended
																					origin:(ExcelOrigin)origin
																			 delimiter:(NSString *)delimiter
																				editable:(BOOL)editable
																					notify:(BOOL)notify
																			 converter:(NSInteger)converter
																				addToMru:(BOOL)addToMru;  // Opens a workbook.
@end

@interface ExcelBaseObject : SBObject <ExcelGenericMethods>

@property (copy) NSDictionary *properties;  // All of the object's properties

@end

@interface ExcelWorkbook : ExcelBaseObject

@property (copy, readonly) NSString *path;  // Returns the complete path of the object, excluding the final separator and name of the object.
@property (copy, readonly) NSString *fullName;  // Returns the name of the workbook, including its path on disk, as a string.

- (void) saveWorkbookAsFilename:(NSString *)filename
										 fileFormat:(ExcelFileFormat)fileFormat
											 password:(NSString *)password
			 writeReservationPassword:(NSString *)writeReservationPassword
						readOnlyRecommended:(BOOL)readOnlyRecommended
									 createBackup:(BOOL)createBackup
										 accessMode:(ExcelAccessMode)accessMode
						 conflictResolution:(ExcelConflictResolution)conflictResolution
			addToMostRecentlyUsedList:(BOOL)addToMostRecentlyUsedList
											overwrite:(BOOL)overwrite;  // Saves changes into a different file.
@end
