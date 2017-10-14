/*
 * PowerPoint.h
 */

#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>


@class PowerPointBaseObject, PowerPointBaseApplication, PowerPointBaseDocument, PowerPointBasicWindow, PowerPointPrintSettings, PowerPointCommandBarControl, PowerPointCommandBarButton, PowerPointCommandBarCombobox, PowerPointCommandBarPopup, PowerPointCommandBar, PowerPointDocumentProperty, PowerPointCustomDocumentProperty, PowerPointWebPageFont, PowerPointActionSetting, PowerPointAddIn, PowerPointAnimationBehavior, PowerPointAnimationPoint, PowerPointAnimationSettings, PowerPointApplication, PowerPointAutocorrectEntry, PowerPointAutocorrect, PowerPointBroadcast, PowerPointBulletFormat, PowerPointColorScheme, PowerPointColorsEffect, PowerPointCommandEffect, PowerPointCustomLayout, PowerPointDefaultWebOptions, PowerPointDesign, PowerPointDocumentWindow, PowerPointEffectInformation, PowerPointEffectParameters, PowerPointEffect, PowerPointFilterEffect, PowerPointFirstLetterException, PowerPointFont, PowerPointHeaderOrFooter, PowerPointHeadersAndFooters, PowerPointHyperlink, PowerPointMaster, PowerPointMotionEffect, PowerPointNamedSlideShow, PowerPointPageSetup, PowerPointPane, PowerPointParagraphFormat, PowerPointPlaySettings, PowerPointPlayer, PowerPointPreferences, PowerPointPresentation, PowerPointPresenterTool, PowerPointPresenterViewWindow, PowerPointPrintOptions, PowerPointPrintRange, PowerPointPropertyEffect, PowerPointRotatingEffect, PowerPointRulerLevel, PowerPointRuler, PowerPointSaveAsMovieSettings, PowerPointScaleEffect, PowerPointSectionProperties, PowerPointSelection, PowerPointSequence, PowerPointSetEffect, PowerPointSlideRange, PowerPointSlideShowSettings, PowerPointSlideShowTransition, PowerPointSlideShowView, PowerPointSlideShowWindow, PowerPointSlide, PowerPointSoundEffect, PowerPointTabStop, PowerPointTextStyleLevel, PowerPointTextStyle, PowerPointTimeline, PowerPointTiming, PowerPointTwoInitialCapsException, PowerPointView, PowerPointWebOptions, PowerPointAdjustment, PowerPointCalloutFormat, PowerPointConnectorFormat, PowerPointFillFormat, PowerPointGlowFormat, PowerPointGradientStop, PowerPointLineFormat, PowerPointLinkFormat, PowerPointOfficeTheme, PowerPointPictureFormat, PowerPointPlaceholderFormat, PowerPointReflectionFormat, PowerPointShadowFormat, PowerPointShapeRange, PowerPointShape, PowerPointCallout, PowerPointComment, PowerPointConnector, PowerPointLineShape, PowerPointMediaObject, PowerPointMedia2Object, PowerPointPicture, PowerPointPlaceHolder, PowerPointShapeTable, PowerPointSoftEdgeFormat, PowerPointTextBox, PowerPointTextColumn, PowerPointTextFrame, PowerPointThemeColorScheme, PowerPointThemeColor, PowerPointThemeEffectScheme, PowerPointThemeFontScheme, PowerPointThemeFont, PowerPointMajorThemeFont, PowerPointMinorThemeFont, PowerPointThreeDFormat, PowerPointWordArtFormat, PowerPointTextRange, PowerPointCharacter, PowerPointLine, PowerPointParagraph, PowerPointSentence, PowerPointTextFlow, PowerPointWord, PowerPointCell, PowerPointColumn, PowerPointRow, PowerPointTable;

enum PowerPointSavo {
	PowerPointSavoYes = 'yes ' /* Save objects now */,
	PowerPointSavoNo = 'no  ' /* Do not save objects */,
	PowerPointSavoAsk = 'ask ' /* Ask the user whether to save */
};
typedef enum PowerPointSavo PowerPointSavo;

enum PowerPointKfrm {
	PowerPointKfrmIndex = 'indx' /* keyform designating indexed access */,
	PowerPointKfrmNamed = 'name' /* keyform designating named access */,
	PowerPointKfrmId = 'ID  ' /* keyform designating access by unique identifier */
};
typedef enum PowerPointKfrm PowerPointKfrm;

enum PowerPointPPfd {
	PowerPointPPfdMacintoshPath = 'utxt' /* Maintosh path i.e. Foo:bar.txt */,
	PowerPointPPfdPosixPath = 'file' /* Posix path i.e. file:://foo/bar.txt */
};
typedef enum PowerPointPPfd PowerPointPPfd;

enum PowerPointPPff {
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
};
typedef enum PowerPointPPff PowerPointPPff;

enum PowerPointMlDs {
	PowerPointMlDsLineDashStyleUnset = '\000\222\377\376',
	PowerPointMlDsLineDashStyleSolid = '\000\223\000\001',
	PowerPointMlDsLineDashStyleSquareDot = '\000\223\000\002',
	PowerPointMlDsLineDashStyleRoundDot = '\000\223\000\003',
	PowerPointMlDsLineDashStyleDash = '\000\223\000\004',
	PowerPointMlDsLineDashStyleDashDot = '\000\223\000\005',
	PowerPointMlDsLineDashStyleDashDotDot = '\000\223\000\006',
	PowerPointMlDsLineDashStyleLongDash = '\000\223\000\007',
	PowerPointMlDsLineDashStyleLongDashDot = '\000\223\000\010',
	PowerPointMlDsLineDashStyleLongDashDotDot = '\000\223\000\011',
	PowerPointMlDsLineDashStyleSystemDash = '\000\223\000\012',
	PowerPointMlDsLineDashStyleSystemDot = '\000\223\000\013',
	PowerPointMlDsLineDashStyleSystemDashDot = '\000\223\000\014'
};
typedef enum PowerPointMlDs PowerPointMlDs;

enum PowerPointMLnS {
	PowerPointMLnSLineStyleUnset = '\000\224\377\376',
	PowerPointMLnSSingleLine = '\000\225\000\001',
	PowerPointMLnSThinThinLine = '\000\225\000\002',
	PowerPointMLnSThinThickLine = '\000\225\000\003',
	PowerPointMLnSThickThinLine = '\000\225\000\004',
	PowerPointMLnSThickBetweenThinLine = '\000\225\000\005'
};
typedef enum PowerPointMLnS PowerPointMLnS;

enum PowerPointMAhS {
	PowerPointMAhSArrowheadStyleUnset = '\000\221\377\376',
	PowerPointMAhSNoArrowhead = '\000\222\000\001',
	PowerPointMAhSTriangleArrowhead = '\000\222\000\002',
	PowerPointMAhSOpen_arrowhead = '\000\222\000\003',
	PowerPointMAhSStealthArrowhead = '\000\222\000\004',
	PowerPointMAhSDiamondArrowhead = '\000\222\000\005',
	PowerPointMAhSOvalArrowhead = '\000\222\000\006'
};
typedef enum PowerPointMAhS PowerPointMAhS;

enum PowerPointMAhW {
	PowerPointMAhWArrowheadWidthUnset = '\000\220\377\376',
	PowerPointMAhWNarrowWidthArrowhead = '\000\221\000\001',
	PowerPointMAhWMediumWidthArrowhead = '\000\221\000\002',
	PowerPointMAhWWideArrowhead = '\000\221\000\003'
};
typedef enum PowerPointMAhW PowerPointMAhW;

enum PowerPointMAhL {
	PowerPointMAhLArrowheadLengthUnset = '\000\223\377\376',
	PowerPointMAhLShortArrowhead = '\000\224\000\001',
	PowerPointMAhLMediumArrowhead = '\000\224\000\002',
	PowerPointMAhLLongArrowhead = '\000\224\000\003'
};
typedef enum PowerPointMAhL PowerPointMAhL;

enum PowerPointMFdT {
	PowerPointMFdTFillUnset = '\000c\377\376',
	PowerPointMFdTFillSolid = '\000d\000\001',
	PowerPointMFdTFillPatterned = '\000d\000\002',
	PowerPointMFdTFillGradient = '\000d\000\003',
	PowerPointMFdTFillTextured = '\000d\000\004',
	PowerPointMFdTFillBackground = '\000d\000\005',
	PowerPointMFdTFillPicture = '\000d\000\006'
};
typedef enum PowerPointMFdT PowerPointMFdT;

enum PowerPointMGdS {
	PowerPointMGdSGradientUnset = '\000d\377\376',
	PowerPointMGdSHorizontalGradient = '\000e\000\001',
	PowerPointMGdSVerticalGradient = '\000e\000\002',
	PowerPointMGdSDiagonalUpGradient = '\000e\000\003',
	PowerPointMGdSDiagonalDownGradient = '\000e\000\004',
	PowerPointMGdSFromCornerGradient = '\000e\000\005',
	PowerPointMGdSFromTitleGradient = '\000e\000\006',
	PowerPointMGdSFromCenterGradient = '\000e\000\007'
};
typedef enum PowerPointMGdS PowerPointMGdS;

enum PowerPointMGCt {
	PowerPointMGCtGradientTypeUnset = '\003\357\377\376',
	PowerPointMGCtSingleShadeGradientType = '\003\360\000\001',
	PowerPointMGCtTwoColorsGradientType = '\003\360\000\002',
	PowerPointMGCtPresetColorsGradientType = '\003\360\000\003',
	PowerPointMGCtMultiColorsGradientType = '\003\360\000\004'
};
typedef enum PowerPointMGCt PowerPointMGCt;

enum PowerPointMxtT {
	PowerPointMxtTTextureTypeTextureTypeUnset = '\003\360\377\376',
	PowerPointMxtTTextureTypePresetTexture = '\003\361\000\001',
	PowerPointMxtTTextureTypeUserDefinedTexture = '\003\361\000\002'
};
typedef enum PowerPointMxtT PowerPointMxtT;

enum PowerPointMPzT {
	PowerPointMPzTPresetTextureUnset = '\000e\377\376',
	PowerPointMPzTTexturePapyrus = '\000f\000\001',
	PowerPointMPzTTextureCanvas = '\000f\000\002',
	PowerPointMPzTTextureDenim = '\000f\000\003',
	PowerPointMPzTTextureWovenMat = '\000f\000\004',
	PowerPointMPzTTextureWaterDroplets = '\000f\000\005',
	PowerPointMPzTTexturePaperBag = '\000f\000\006',
	PowerPointMPzTTextureFishFossil = '\000f\000\007',
	PowerPointMPzTTextureSand = '\000f\000\010',
	PowerPointMPzTTextureGreenMarble = '\000f\000\011',
	PowerPointMPzTTextureWhiteMarble = '\000f\000\012',
	PowerPointMPzTTextureBrownMarble = '\000f\000\013',
	PowerPointMPzTTextureGranite = '\000f\000\014',
	PowerPointMPzTTextureNewsprint = '\000f\000\015',
	PowerPointMPzTTextureRecycledPaper = '\000f\000\016',
	PowerPointMPzTTextureParchment = '\000f\000\017',
	PowerPointMPzTTextureStationery = '\000f\000\020',
	PowerPointMPzTTextureBlueTissuePaper = '\000f\000\021',
	PowerPointMPzTTexturePinkTissuePaper = '\000f\000\022',
	PowerPointMPzTTexturePurpleMesh = '\000f\000\023',
	PowerPointMPzTTextureBouquet = '\000f\000\024',
	PowerPointMPzTTextureCork = '\000f\000\025',
	PowerPointMPzTTextureWalnut = '\000f\000\026',
	PowerPointMPzTTextureOak = '\000f\000\027',
	PowerPointMPzTTextureMediumWood = '\000f\000\030'
};
typedef enum PowerPointMPzT PowerPointMPzT;

enum PowerPointPpTy {
	PowerPointPpTyPatternUnset = '\000f\377\376',
	PowerPointPpTyFivePercentPattern = '\000g\000\001',
	PowerPointPpTyTenPercentPattern = '\000g\000\002',
	PowerPointPpTyTwentyPercentPattern = '\000g\000\003',
	PowerPointPpTyTwentyFivePercentPattern = '\000g\000\004',
	PowerPointPpTyThirtyPercentPattern = '\000g\000\005',
	PowerPointPpTyFortyPercentPattern = '\000g\000\006',
	PowerPointPpTyFiftyPercentPattern = '\000g\000\007',
	PowerPointPpTySixtyPercentPattern = '\000g\000\010',
	PowerPointPpTySeventyPercentPattern = '\000g\000\011',
	PowerPointPpTySeventyFivePercentPattern = '\000g\000\012',
	PowerPointPpTyEightyPercentPattern = '\000g\000\013',
	PowerPointPpTyNinetyPercentPattern = '\000g\000\014',
	PowerPointPpTyDarkHorizontalPattern = '\000g\000\015',
	PowerPointPpTyDarkVerticalPattern = '\000g\000\016',
	PowerPointPpTyDarkDownwardDiagonalPattern = '\000g\000\017',
	PowerPointPpTyDarkUpwardDiagonalPattern = '\000g\000\020',
	PowerPointPpTySmallCheckerBoardPattern = '\000g\000\021',
	PowerPointPpTyTrellisPattern = '\000g\000\022',
	PowerPointPpTyLightHorizontalPattern = '\000g\000\023',
	PowerPointPpTyLightVerticalPattern = '\000g\000\024',
	PowerPointPpTyLightDownwardDiagonalPattern = '\000g\000\025',
	PowerPointPpTyLightUpwardDiagonalPattern = '\000g\000\026',
	PowerPointPpTySmallGridPattern = '\000g\000\027',
	PowerPointPpTyDottedDiamondPattern = '\000g\000\030',
	PowerPointPpTyWideDownwardDiagonal = '\000g\000\031',
	PowerPointPpTyWideUpwardDiagonalPattern = '\000g\000\032',
	PowerPointPpTyDashedUpwardDiagonalPattern = '\000g\000\033',
	PowerPointPpTyDashedDownwardDiagonalPattern = '\000g\000\034',
	PowerPointPpTyNarrowVerticalPattern = '\000g\000\035',
	PowerPointPpTyNarrowHorizontalPattern = '\000g\000\036',
	PowerPointPpTyDashedVerticalPattern = '\000g\000\037',
	PowerPointPpTyDashedHorizontalPattern = '\000g\000 ',
	PowerPointPpTyLargeConfettiPattern = '\000g\000!',
	PowerPointPpTyLargeGridPattern = '\000g\000\"',
	PowerPointPpTyHorizontalBrickPattern = '\000g\000#',
	PowerPointPpTyLargeCheckerBoardPattern = '\000g\000$',
	PowerPointPpTySmallConfettiPattern = '\000g\000%',
	PowerPointPpTyZigZagPattern = '\000g\000&',
	PowerPointPpTySolidDiamondPattern = '\000g\000\'',
	PowerPointPpTyDiagonalBrickPattern = '\000g\000(',
	PowerPointPpTyOutlinedDiamondPattern = '\000g\000)',
	PowerPointPpTyPlaidPattern = '\000g\000*',
	PowerPointPpTySpherePattern = '\000g\000+',
	PowerPointPpTyWeavePattern = '\000g\000,',
	PowerPointPpTyDottedGridPattern = '\000g\000-',
	PowerPointPpTyDivotPattern = '\000g\000.',
	PowerPointPpTyShinglePattern = '\000g\000/',
	PowerPointPpTyWavePattern = '\000g\0000',
	PowerPointPpTyHorizontalPattern = '\000g\0001',
	PowerPointPpTyVerticalPattern = '\000g\0002',
	PowerPointPpTyCrossPattern = '\000g\0003',
	PowerPointPpTyDownwardDiagonalPattern = '\000g\0004',
	PowerPointPpTyUpwardDiagonalPattern = '\000g\0005',
	PowerPointPpTyDiagonalCrossPattern = '\000g\0005'
};
typedef enum PowerPointPpTy PowerPointPpTy;

enum PowerPointMPGb {
	PowerPointMPGbPresetGradientUnset = '\000g\377\376',
	PowerPointMPGbGradientEarlySunset = '\000h\000\001',
	PowerPointMPGbGradientLateSunset = '\000h\000\002',
	PowerPointMPGbGradientNightfall = '\000h\000\003',
	PowerPointMPGbGradientDaybreak = '\000h\000\004',
	PowerPointMPGbGradientHorizon = '\000h\000\005',
	PowerPointMPGbGradientDesert = '\000h\000\006',
	PowerPointMPGbGradientOcean = '\000h\000\007',
	PowerPointMPGbGradientCalmWater = '\000h\000\010',
	PowerPointMPGbGradientFire = '\000h\000\011',
	PowerPointMPGbGradientFog = '\000h\000\012',
	PowerPointMPGbGradientMoss = '\000h\000\013',
	PowerPointMPGbGradientPeacock = '\000h\000\014',
	PowerPointMPGbGradientWheat = '\000h\000\015',
	PowerPointMPGbGradientParchment = '\000h\000\016',
	PowerPointMPGbGradientMahogany = '\000h\000\017',
	PowerPointMPGbGradientRainbow = '\000h\000\020',
	PowerPointMPGbGradientRainbow2 = '\000h\000\021',
	PowerPointMPGbGradientGold = '\000h\000\022',
	PowerPointMPGbGradientGold2 = '\000h\000\023',
	PowerPointMPGbGradientBrass = '\000h\000\024',
	PowerPointMPGbGradientChrome = '\000h\000\025',
	PowerPointMPGbGradientChrome2 = '\000h\000\026',
	PowerPointMPGbGradientSilver = '\000h\000\027',
	PowerPointMPGbGradientSapphire = '\000h\000\030'
};
typedef enum PowerPointMPGb PowerPointMPGb;

enum PowerPointMSdT {
	PowerPointMSdTShadowUnset = '\003_\377\376',
	PowerPointMSdTShadow1 = '\003`\000\001',
	PowerPointMSdTShadow2 = '\003`\000\002',
	PowerPointMSdTShadow3 = '\003`\000\003',
	PowerPointMSdTShadow4 = '\003`\000\004',
	PowerPointMSdTShadow5 = '\003`\000\005',
	PowerPointMSdTShadow6 = '\003`\000\006',
	PowerPointMSdTShadow7 = '\003`\000\007',
	PowerPointMSdTShadow8 = '\003`\000\010',
	PowerPointMSdTShadow9 = '\003`\000\011',
	PowerPointMSdTShadow10 = '\003`\000\012',
	PowerPointMSdTShadow11 = '\003`\000\013',
	PowerPointMSdTShadow12 = '\003`\000\014',
	PowerPointMSdTShadow13 = '\003`\000\015',
	PowerPointMSdTShadow14 = '\003`\000\016',
	PowerPointMSdTShadow15 = '\003`\000\017',
	PowerPointMSdTShadow16 = '\003`\000\020',
	PowerPointMSdTShadow17 = '\003`\000\021',
	PowerPointMSdTShadow18 = '\003`\000\022',
	PowerPointMSdTShadow19 = '\003`\000\023',
	PowerPointMSdTShadow20 = '\003`\000\024',
	PowerPointMSdTShadow21 = '\003`\000\025',
	PowerPointMSdTShadow22 = '\003`\000\026',
	PowerPointMSdTShadow23 = '\003`\000\027',
	PowerPointMSdTShadow24 = '\003`\000\030',
	PowerPointMSdTShadow25 = '\003`\000\031',
	PowerPointMSdTShadow26 = '\003`\000\032',
	PowerPointMSdTShadow27 = '\003`\000\033',
	PowerPointMSdTShadow28 = '\003`\000\034',
	PowerPointMSdTShadow29 = '\003`\000\035',
	PowerPointMSdTShadow30 = '\003`\000\036',
	PowerPointMSdTShadow31 = '\003`\000\037',
	PowerPointMSdTShadow32 = '\003`\000 ',
	PowerPointMSdTShadow33 = '\003`\000!',
	PowerPointMSdTShadow34 = '\003`\000\"',
	PowerPointMSdTShadow35 = '\003`\000#',
	PowerPointMSdTShadow36 = '\003`\000$',
	PowerPointMSdTShadow37 = '\003`\000%',
	PowerPointMSdTShadow38 = '\003`\000&',
	PowerPointMSdTShadow39 = '\003`\000\'',
	PowerPointMSdTShadow40 = '\003`\000(',
	PowerPointMSdTShadow41 = '\003`\000)',
	PowerPointMSdTShadow42 = '\003`\000*',
	PowerPointMSdTShadow43 = '\003`\000+'
};
typedef enum PowerPointMSdT PowerPointMSdT;

enum PowerPointMPXF {
	PowerPointMPXFWordartFormatUnset = '\003\361\377\376',
	PowerPointMPXFWordartFormat1 = '\003\362\000\000',
	PowerPointMPXFWordartFormat2 = '\003\362\000\001',
	PowerPointMPXFWordartFormat3 = '\003\362\000\002',
	PowerPointMPXFWordartFormat4 = '\003\362\000\003',
	PowerPointMPXFWordartFormat5 = '\003\362\000\004',
	PowerPointMPXFWordartFormat6 = '\003\362\000\005',
	PowerPointMPXFWordartFormat7 = '\003\362\000\006',
	PowerPointMPXFWordartFormat8 = '\003\362\000\007',
	PowerPointMPXFWordartFormat9 = '\003\362\000\010',
	PowerPointMPXFWordartFormat10 = '\003\362\000\011',
	PowerPointMPXFWordartFormat11 = '\003\362\000\012',
	PowerPointMPXFWordartFormat12 = '\003\362\000\013',
	PowerPointMPXFWordartFormat13 = '\003\362\000\014',
	PowerPointMPXFWordartFormat14 = '\003\362\000\015',
	PowerPointMPXFWordartFormat15 = '\003\362\000\016',
	PowerPointMPXFWordartFormat16 = '\003\362\000\017',
	PowerPointMPXFWordartFormat17 = '\003\362\000\020',
	PowerPointMPXFWordartFormat18 = '\003\362\000\021',
	PowerPointMPXFWordartFormat19 = '\003\362\000\022',
	PowerPointMPXFWordartFormat20 = '\003\362\000\023',
	PowerPointMPXFWordartFormat21 = '\003\362\000\024',
	PowerPointMPXFWordartFormat22 = '\003\362\000\025',
	PowerPointMPXFWordartFormat23 = '\003\362\000\026',
	PowerPointMPXFWordartFormat24 = '\003\362\000\027',
	PowerPointMPXFWordartFormat25 = '\003\362\000\030',
	PowerPointMPXFWordartFormat26 = '\003\362\000\031',
	PowerPointMPXFWordartFormat27 = '\003\362\000\032',
	PowerPointMPXFWordartFormat28 = '\003\362\000\033',
	PowerPointMPXFWordartFormat29 = '\003\362\000\034',
	PowerPointMPXFWordartFormat30 = '\003\362\000\035'
};
typedef enum PowerPointMPXF PowerPointMPXF;

enum PowerPointMPTs {
	PowerPointMPTsTextEffectShapeUnset = '\000\227\377\376',
	PowerPointMPTsPlainText = '\000\230\000\001',
	PowerPointMPTsStop = '\000\230\000\002',
	PowerPointMPTsTriangleUp = '\000\230\000\003',
	PowerPointMPTsTriangleDown = '\000\230\000\004',
	PowerPointMPTsChevronUp = '\000\230\000\005',
	PowerPointMPTsChevronDown = '\000\230\000\006',
	PowerPointMPTsRingInside = '\000\230\000\007',
	PowerPointMPTsRingOutside = '\000\230\000\010',
	PowerPointMPTsArchUpCurve = '\000\230\000\011',
	PowerPointMPTsArchDownCurve = '\000\230\000\012',
	PowerPointMPTsCircleCurve = '\000\230\000\013',
	PowerPointMPTsButtonCurve = '\000\230\000\014',
	PowerPointMPTsArchUpPour = '\000\230\000\015',
	PowerPointMPTsArchDownPour = '\000\230\000\016',
	PowerPointMPTsCirclePour = '\000\230\000\017',
	PowerPointMPTsButtonPour = '\000\230\000\020',
	PowerPointMPTsCurveUp = '\000\230\000\021',
	PowerPointMPTsCurveDown = '\000\230\000\022',
	PowerPointMPTsCanUp = '\000\230\000\023',
	PowerPointMPTsCanDown = '\000\230\000\024',
	PowerPointMPTsWave1 = '\000\230\000\025',
	PowerPointMPTsWave2 = '\000\230\000\026',
	PowerPointMPTsDoubleWave1 = '\000\230\000\027',
	PowerPointMPTsDoubleWave2 = '\000\230\000\030',
	PowerPointMPTsInflate = '\000\230\000\031',
	PowerPointMPTsDeflate = '\000\230\000\032',
	PowerPointMPTsInflateBottom = '\000\230\000\033',
	PowerPointMPTsDeflateBottom = '\000\230\000\034',
	PowerPointMPTsInflateTop = '\000\230\000\035',
	PowerPointMPTsDeflateTop = '\000\230\000\036',
	PowerPointMPTsDeflateInflate = '\000\230\000\037',
	PowerPointMPTsDeflateInflateDeflate = '\000\230\000 ',
	PowerPointMPTsFadeRight = '\000\230\000!',
	PowerPointMPTsFadeLeft = '\000\230\000\"',
	PowerPointMPTsFadeUp = '\000\230\000#',
	PowerPointMPTsFadeDown = '\000\230\000$',
	PowerPointMPTsSlantUp = '\000\230\000%',
	PowerPointMPTsSlantDown = '\000\230\000&',
	PowerPointMPTsCascadeUp = '\000\230\000\'',
	PowerPointMPTsCascadeDown = '\000\230\000('
};
typedef enum PowerPointMPTs PowerPointMPTs;

enum PowerPointMTxA {
	PowerPointMTxATextEffectAlignmentUnset = '\000\226\377\376',
	PowerPointMTxALeftTextEffectAlignment = '\000\227\000\001',
	PowerPointMTxACenteredTextEffectAlignment = '\000\227\000\002',
	PowerPointMTxARightTextEffectAlignment = '\000\227\000\003',
	PowerPointMTxAJustifyTextEffectAlignment = '\000\227\000\004',
	PowerPointMTxAWordJustifyTextEffectAlignment = '\000\227\000\005',
	PowerPointMTxAStretchJustifyTextEffectAlignment = '\000\227\000\006'
};
typedef enum PowerPointMTxA PowerPointMTxA;

enum PowerPointMPLd {
	PowerPointMPLdPresetLightingDirectionUnset = '\000\233\377\376',
	PowerPointMPLdLightFromTopLeft = '\000\234\000\001',
	PowerPointMPLdLightFromTop = '\000\234\000\002',
	PowerPointMPLdLightFromTopRight = '\000\234\000\003',
	PowerPointMPLdLightFromLeft = '\000\234\000\004',
	PowerPointMPLdLightFromNone = '\000\234\000\005',
	PowerPointMPLdLightFromRight = '\000\234\000\006',
	PowerPointMPLdLightFromBottomLeft = '\000\234\000\007',
	PowerPointMPLdLightFromBottom = '\000\234\000\010',
	PowerPointMPLdLightFromBottomRight = '\000\234\000\011'
};
typedef enum PowerPointMPLd PowerPointMPLd;

enum PowerPointMlSf {
	PowerPointMlSfLightingSoftnessUnset = '\000\234\377\376',
	PowerPointMlSfLightingDim = '\000\235\000\001',
	PowerPointMlSfLightingNormal = '\000\235\000\002',
	PowerPointMlSfLightingBright = '\000\235\000\003'
};
typedef enum PowerPointMlSf PowerPointMlSf;

enum PowerPointMPMt {
	PowerPointMPMtPresetMaterialUnset = '\000\235\377\376',
	PowerPointMPMtMatte = '\000\236\000\001',
	PowerPointMPMtPlastic = '\000\236\000\002',
	PowerPointMPMtMetal = '\000\236\000\003',
	PowerPointMPMtWireframe = '\000\236\000\004',
	PowerPointMPMtMatte2 = '\000\236\000\005',
	PowerPointMPMtPlastic2 = '\000\236\000\006',
	PowerPointMPMtMetal2 = '\000\236\000\007',
	PowerPointMPMtWarmMatte = '\000\236\000\010',
	PowerPointMPMtTranslucentPowder = '\000\236\000\011',
	PowerPointMPMtPowder = '\000\236\000\012',
	PowerPointMPMtDarkEdge = '\000\236\000\013',
	PowerPointMPMtSoftEdge = '\000\236\000\014',
	PowerPointMPMtMaterialClear = '\000\236\000\015',
	PowerPointMPMtFlat = '\000\236\000\016',
	PowerPointMPMtSoftMetal = '\000\236\000\017'
};
typedef enum PowerPointMPMt PowerPointMPMt;

enum PowerPointMExD {
	PowerPointMExDPresetExtrusionDirectionUnset = '\000\231\377\376',
	PowerPointMExDExtrudeBottomRight = '\000\232\000\001',
	PowerPointMExDExtrudeBottom = '\000\232\000\002',
	PowerPointMExDExtrudeBottomLeft = '\000\232\000\003',
	PowerPointMExDExtrudeRight = '\000\232\000\004',
	PowerPointMExDExtrudeNone = '\000\232\000\005',
	PowerPointMExDExtrudeLeft = '\000\232\000\006',
	PowerPointMExDExtrudeTopRight = '\000\232\000\007',
	PowerPointMExDExtrudeTop = '\000\232\000\010',
	PowerPointMExDExtrudeTopLeft = '\000\232\000\011'
};
typedef enum PowerPointMExD PowerPointMExD;

enum PowerPointM3DF {
	PowerPointM3DFPresetThreeDFormatUnset = '\000\230\377\376',
	PowerPointM3DFFormat1 = '\000\231\000\001',
	PowerPointM3DFFormat2 = '\000\231\000\002',
	PowerPointM3DFFormat3 = '\000\231\000\003',
	PowerPointM3DFFormat4 = '\000\231\000\004',
	PowerPointM3DFFormat5 = '\000\231\000\005',
	PowerPointM3DFFormat6 = '\000\231\000\006',
	PowerPointM3DFFormat7 = '\000\231\000\007',
	PowerPointM3DFFormat8 = '\000\231\000\010',
	PowerPointM3DFFormat9 = '\000\231\000\011',
	PowerPointM3DFFormat10 = '\000\231\000\012',
	PowerPointM3DFFormat11 = '\000\231\000\013',
	PowerPointM3DFFormat12 = '\000\231\000\014',
	PowerPointM3DFFormat13 = '\000\231\000\015',
	PowerPointM3DFFormat14 = '\000\231\000\016',
	PowerPointM3DFFormat15 = '\000\231\000\017',
	PowerPointM3DFFormat16 = '\000\231\000\020',
	PowerPointM3DFFormat17 = '\000\231\000\021',
	PowerPointM3DFFormat18 = '\000\231\000\022',
	PowerPointM3DFFormat19 = '\000\231\000\023',
	PowerPointM3DFFormat20 = '\000\231\000\024'
};
typedef enum PowerPointM3DF PowerPointM3DF;

enum PowerPointMExC {
	PowerPointMExCExtrusionColorUnset = '\000\232\377\376',
	PowerPointMExCExtrusionColorAutomatic = '\000\233\000\001',
	PowerPointMExCExtrusionColorCustom = '\000\233\000\002'
};
typedef enum PowerPointMExC PowerPointMExC;

enum PowerPointMCtT {
	PowerPointMCtTConnectorTypeUnset = '\000h\377\376',
	PowerPointMCtTStraight = '\000i\000\001',
	PowerPointMCtTElbow = '\000i\000\002',
	PowerPointMCtTCurve = '\000i\000\003'
};
typedef enum PowerPointMCtT PowerPointMCtT;

enum PowerPointMHzA {
	PowerPointMHzAHorizontalAnchorUnset = '\000\236\377\376',
	PowerPointMHzAHorizontalAnchorNone = '\000\237\000\001',
	PowerPointMHzAHorizontalAnchorCenter = '\000\237\000\002'
};
typedef enum PowerPointMHzA PowerPointMHzA;

enum PowerPointMVtA {
	PowerPointMVtAVerticalAnchorUnset = '\000\237\377\376',
	PowerPointMVtAAnchorTop = '\000\240\000\001',
	PowerPointMVtAAnchorTopBaseline = '\000\240\000\002',
	PowerPointMVtAAnchorMiddle = '\000\240\000\003',
	PowerPointMVtAAnchorBottom = '\000\240\000\004',
	PowerPointMVtAAnchorBottomBaseline = '\000\240\000\005'
};
typedef enum PowerPointMVtA PowerPointMVtA;

enum PowerPointMAsT {
	PowerPointMAsTAutoshapeShapeTypeUnset = '\000i\377\376',
	PowerPointMAsTAutoshapeRectangle = '\000j\000\001',
	PowerPointMAsTAutoshapeParallelogram = '\000j\000\002',
	PowerPointMAsTAutoshapeTrapezoid = '\000j\000\003',
	PowerPointMAsTAutoshapeDiamond = '\000j\000\004',
	PowerPointMAsTAutoshapeRoundedRectangle = '\000j\000\005',
	PowerPointMAsTAutoshapeOctagon = '\000j\000\006',
	PowerPointMAsTAutoshapeIsoscelesTriangle = '\000j\000\007',
	PowerPointMAsTAutoshapeRightTriangle = '\000j\000\010',
	PowerPointMAsTAutoshapeOval = '\000j\000\011',
	PowerPointMAsTAutoshapeHexagon = '\000j\000\012',
	PowerPointMAsTAutoshapeCross = '\000j\000\013',
	PowerPointMAsTAutoshapeRegularPentagon = '\000j\000\014',
	PowerPointMAsTAutoshapeCan = '\000j\000\015',
	PowerPointMAsTAutoshapeCube = '\000j\000\016',
	PowerPointMAsTAutoshapeBevel = '\000j\000\017',
	PowerPointMAsTAutoshapeFoldedCorner = '\000j\000\020',
	PowerPointMAsTAutoshapeSmileyFace = '\000j\000\021',
	PowerPointMAsTAutoshapeDonut = '\000j\000\022',
	PowerPointMAsTAutoshapeNoSymbol = '\000j\000\023',
	PowerPointMAsTAutoshapeBlockArc = '\000j\000\024',
	PowerPointMAsTAutoshapeHeart = '\000j\000\025',
	PowerPointMAsTAutoshapeLightningBolt = '\000j\000\026',
	PowerPointMAsTAutoshapeSun = '\000j\000\027',
	PowerPointMAsTAutoshapeMoon = '\000j\000\030',
	PowerPointMAsTAutoshapeArc = '\000j\000\031',
	PowerPointMAsTAutoshapeDoubleBracket = '\000j\000\032',
	PowerPointMAsTAutoshapeDoubleBrace = '\000j\000\033',
	PowerPointMAsTAutoshapePlaque = '\000j\000\034',
	PowerPointMAsTAutoshapeLeftBracket = '\000j\000\035',
	PowerPointMAsTAutoshapeRightBracket = '\000j\000\036',
	PowerPointMAsTAutoshapeLeftBrace = '\000j\000\037',
	PowerPointMAsTAutoshapeRightBrace = '\000j\000 ',
	PowerPointMAsTAutoshapeRightArrow = '\000j\000!',
	PowerPointMAsTAutoshapeLeftArrow = '\000j\000\"',
	PowerPointMAsTAutoshapeUpArrow = '\000j\000#',
	PowerPointMAsTAutoshapeDownArrow = '\000j\000$',
	PowerPointMAsTAutoshapeLeftRightArrow = '\000j\000%',
	PowerPointMAsTAutoshapeUpDownArrow = '\000j\000&',
	PowerPointMAsTAutoshapeQuadArrow = '\000j\000\'',
	PowerPointMAsTAutoshapeLeftRightUpArrow = '\000j\000(',
	PowerPointMAsTAutoshapeBentArrow = '\000j\000)',
	PowerPointMAsTAutoshapeUTurnArrow = '\000j\000*',
	PowerPointMAsTAutoshapeLeftUpArrow = '\000j\000+',
	PowerPointMAsTAutoshapeBentUpArrow = '\000j\000,',
	PowerPointMAsTAutoshapeCurvedRightArrow = '\000j\000-',
	PowerPointMAsTAutoshapeCurvedLeftArrow = '\000j\000.',
	PowerPointMAsTAutoshapeCurvedUpArrow = '\000j\000/',
	PowerPointMAsTAutoshapeCurvedDownArrow = '\000j\0000',
	PowerPointMAsTAutoshapeStripedRightArrow = '\000j\0001',
	PowerPointMAsTAutoshapeNotchedRightArrow = '\000j\0002',
	PowerPointMAsTAutoshapePentagon = '\000j\0003',
	PowerPointMAsTAutoshapeChevron = '\000j\0004',
	PowerPointMAsTAutoshapeRightArrowCallout = '\000j\0005',
	PowerPointMAsTAutoshapeLeftArrowCallout = '\000j\0006',
	PowerPointMAsTAutoshapeUpArrowCallout = '\000j\0007',
	PowerPointMAsTAutoshapeDownArrowCallout = '\000j\0008',
	PowerPointMAsTAutoshapeLeftRightArrowCallout = '\000j\0009',
	PowerPointMAsTAutoshapeUpDownArrowCallout = '\000j\000:',
	PowerPointMAsTAutoshapeQuadArrowCallout = '\000j\000;',
	PowerPointMAsTAutoshapeCircularArrow = '\000j\000<',
	PowerPointMAsTAutoshapeFlowchartProcess = '\000j\000=',
	PowerPointMAsTAutoshapeFlowchartAlternateProcess = '\000j\000>',
	PowerPointMAsTAutoshapeFlowchartDecision = '\000j\000\?',
	PowerPointMAsTAutoshapeFlowchartData = '\000j\000@',
	PowerPointMAsTAutoshapeFlowchartPredefinedProcess = '\000j\000A',
	PowerPointMAsTAutoshapeFlowchartInternalStorage = '\000j\000B',
	PowerPointMAsTAutoshapeFlowchartDocument = '\000j\000C',
	PowerPointMAsTAutoshapeFlowchartMultiDocument = '\000j\000D',
	PowerPointMAsTAutoshapeFlowchartTerminator = '\000j\000E',
	PowerPointMAsTAutoshapeFlowchartPreparation = '\000j\000F',
	PowerPointMAsTAutoshapeFlowchartManualInput = '\000j\000G',
	PowerPointMAsTAutoshapeFlowchartManualOperation = '\000j\000H',
	PowerPointMAsTAutoshapeFlowchartConnector = '\000j\000I',
	PowerPointMAsTAutoshapeFlowchartOffpageConnector = '\000j\000J',
	PowerPointMAsTAutoshapeFlowchartCard = '\000j\000K',
	PowerPointMAsTAutoshapeFlowchartPunchedTape = '\000j\000L',
	PowerPointMAsTAutoshapeFlowchartSummingJunction = '\000j\000M',
	PowerPointMAsTAutoshapeFlowchartOr = '\000j\000N',
	PowerPointMAsTAutoshapeFlowchartCollate = '\000j\000O',
	PowerPointMAsTAutoshapeFlowchartSort = '\000j\000P',
	PowerPointMAsTAutoshapeFlowchartExtract = '\000j\000Q',
	PowerPointMAsTAutoshapeFlowchartMerge = '\000j\000R',
	PowerPointMAsTAutoshapeFlowchartStoredData = '\000j\000S',
	PowerPointMAsTAutoshapeFlowchartDelay = '\000j\000T',
	PowerPointMAsTAutoshapeFlowchartSequentialAccessStorage = '\000j\000U',
	PowerPointMAsTAutoshapeFlowchartMagneticDisk = '\000j\000V',
	PowerPointMAsTAutoshapeFlowchartDirectAccessStorage = '\000j\000W',
	PowerPointMAsTAutoshapeFlowchartDisplay = '\000j\000X',
	PowerPointMAsTAutoshapeExplosionOne = '\000j\000Y',
	PowerPointMAsTAutoshapeExplosionTwo = '\000j\000Z',
	PowerPointMAsTAutoshapeFourPointStar = '\000j\000[',
	PowerPointMAsTAutoshapeFivePointStar = '\000j\000\\',
	PowerPointMAsTAutoshapeEightPointStar = '\000j\000]',
	PowerPointMAsTAutoshapeSixteenPointStar = '\000j\000^',
	PowerPointMAsTAutoshapeTwentyFourPointStar = '\000j\000_',
	PowerPointMAsTAutoshapeThirtyTwoPointStar = '\000j\000`',
	PowerPointMAsTAutoshapeUpRibbon = '\000j\000a',
	PowerPointMAsTAutoshapeDownRibbon = '\000j\000b',
	PowerPointMAsTAutoshapeCurvedUpRibbon = '\000j\000c',
	PowerPointMAsTAutoshapeCurvedDownRibbon = '\000j\000d',
	PowerPointMAsTAutoshapeVerticalScroll = '\000j\000e',
	PowerPointMAsTAutoshapeHorizontalScroll = '\000j\000f',
	PowerPointMAsTAutoshapeWave = '\000j\000g',
	PowerPointMAsTAutoshapeDoubleWave = '\000j\000h',
	PowerPointMAsTAutoshapeRectangularCallout = '\000j\000i',
	PowerPointMAsTAutoshapeRoundedRectangularCallout = '\000j\000j',
	PowerPointMAsTAutoshapeOvalCallout = '\000j\000k',
	PowerPointMAsTAutoshapeCloudCallout = '\000j\000l',
	PowerPointMAsTAutoshapeLineCalloutOne = '\000j\000m',
	PowerPointMAsTAutoshapeLineCalloutTwo = '\000j\000n',
	PowerPointMAsTAutoshapeLineCalloutThree = '\000j\000o',
	PowerPointMAsTAutoshapeLineCalloutFour = '\000j\000p',
	PowerPointMAsTAutoshapeLineCalloutOneAccentBar = '\000j\000q',
	PowerPointMAsTAutoshapeLineCalloutTwoAccentBar = '\000j\000r',
	PowerPointMAsTAutoshapeLineCalloutThreeAccentBar = '\000j\000s',
	PowerPointMAsTAutoshapeLineCalloutFourAccentBar = '\000j\000t',
	PowerPointMAsTAutoshapeLineCalloutOneNoBorder = '\000j\000u',
	PowerPointMAsTAutoshapeLineCalloutTwoNoBorder = '\000j\000v',
	PowerPointMAsTAutoshapeLineCalloutThreeNoBorder = '\000j\000w',
	PowerPointMAsTAutoshapeLineCalloutFourNoBorder = '\000j\000x',
	PowerPointMAsTAutoshapeCalloutOneBorderAndAccentBar = '\000j\000y',
	PowerPointMAsTAutoshapeCalloutTwoBorderAndAccentBar = '\000j\000z',
	PowerPointMAsTAutoshapeCalloutThreeBorderAndAccentBar = '\000j\000{',
	PowerPointMAsTAutoshapeCalloutFourBorderAndAccentBar = '\000j\000|',
	PowerPointMAsTAutoshapeActionButtonCustom = '\000j\000}',
	PowerPointMAsTAutoshapeActionButtonHome = '\000j\000~',
	PowerPointMAsTAutoshapeActionButtonHelp = '\000j\000\177',
	PowerPointMAsTAutoshapeActionButtonInformation = '\000j\000\200',
	PowerPointMAsTAutoshapeActionButtonBackOrPrevious = '\000j\000\201',
	PowerPointMAsTAutoshapeActionButtonForwardOrNext = '\000j\000\202',
	PowerPointMAsTAutoshapeActionButtonBeginning = '\000j\000\203',
	PowerPointMAsTAutoshapeActionButtonEnd = '\000j\000\204',
	PowerPointMAsTAutoshapeActionButtonReturn = '\000j\000\205',
	PowerPointMAsTAutoshapeActionButtonDocument = '\000j\000\206',
	PowerPointMAsTAutoshapeActionButtonSound = '\000j\000\207',
	PowerPointMAsTAutoshapeActionButtonMovie = '\000j\000\210',
	PowerPointMAsTAutoshapeBalloon = '\000j\000\211',
	PowerPointMAsTAutoshapeNotPrimitive = '\000j\000\212',
	PowerPointMAsTAutoshapeFlowchartOfflineStorage = '\000j\000\213',
	PowerPointMAsTAutoshapeLeftRightRibbon = '\000j\000\214',
	PowerPointMAsTAutoshapeDiagonalStripe = '\000j\000\215',
	PowerPointMAsTAutoshapePie = '\000j\000\216',
	PowerPointMAsTAutoshapeNonIsoscelesTrapezoid = '\000j\000\217',
	PowerPointMAsTAutoshapeDecagon = '\000j\000\220',
	PowerPointMAsTAutoshapeHeptagon = '\000j\000\221',
	PowerPointMAsTAutoshapeDodecagon = '\000j\000\222',
	PowerPointMAsTAutoshapeSixPointsStar = '\000j\000\223',
	PowerPointMAsTAutoshapeSevenPointsStar = '\000j\000\224',
	PowerPointMAsTAutoshapeTenPointsStar = '\000j\000\225',
	PowerPointMAsTAutoshapeTwelvePointsStar = '\000j\000\226',
	PowerPointMAsTAutoshapeRoundOneRectangle = '\000j\000\227',
	PowerPointMAsTAutoshapeRoundTwoSameRectangle = '\000j\000\230',
	PowerPointMAsTAutoshapeRoundTwoDiagonalRectangle = '\000j\000\231',
	PowerPointMAsTAutoshapeSnipRoundRectangle = '\000j\000\232',
	PowerPointMAsTAutoshapeSnipOneRectangle = '\000j\000\233',
	PowerPointMAsTAutoshapeSnipTwoSameRectangle = '\000j\000\234',
	PowerPointMAsTAutoshapeSnipTwoDiagonalRectangle = '\000j\000\235',
	PowerPointMAsTAutoshapeFrame = '\000j\000\236',
	PowerPointMAsTAutoshapeHalfFrame = '\000j\000\237',
	PowerPointMAsTAutoshapeTear = '\000j\000\240',
	PowerPointMAsTAutoshapeChord = '\000j\000\241',
	PowerPointMAsTAutoshapeCorner = '\000j\000\242',
	PowerPointMAsTAutoshapeMathPlus = '\000j\000\243',
	PowerPointMAsTAutoshapeMathMinus = '\000j\000\244',
	PowerPointMAsTAutoshapeMathMultiply = '\000j\000\245',
	PowerPointMAsTAutoshapeMathDivide = '\000j\000\246',
	PowerPointMAsTAutoshapeMathEqual = '\000j\000\247',
	PowerPointMAsTAutoshapeMathNotEqual = '\000j\000\250',
	PowerPointMAsTAutoshapeCornerTabs = '\000j\000\251',
	PowerPointMAsTAutoshapeSquareTabs = '\000j\000\252',
	PowerPointMAsTAutoshapePlaqueTabs = '\000j\000\253',
	PowerPointMAsTAutoshapeGearSix = '\000j\000\254',
	PowerPointMAsTAutoshapeGearNine = '\000j\000\255',
	PowerPointMAsTAutoshapeFunnel = '\000j\000\256',
	PowerPointMAsTAutoshapePieWedge = '\000j\000\257',
	PowerPointMAsTAutoshapeLeftCircularArrow = '\000j\000\260',
	PowerPointMAsTAutoshapeLeftRightCircularArrow = '\000j\000\261',
	PowerPointMAsTAutoshapeSwooshArrow = '\000j\000\262',
	PowerPointMAsTAutoshapeCloud = '\000j\000\263',
	PowerPointMAsTAutoshapeChartX = '\000j\000\264',
	PowerPointMAsTAutoshapeChartStar = '\000j\000\265',
	PowerPointMAsTAutoshapeChartPlus = '\000j\000\266',
	PowerPointMAsTAutoshapeLineInverse = '\000j\000\267'
};
typedef enum PowerPointMAsT PowerPointMAsT;

enum PowerPointMShp {
	PowerPointMShpShapeTypeUnset = '\000\213\377\376',
	PowerPointMShpShapeTypeAuto = '\000\214\000\001',
	PowerPointMShpShapeTypeCallout = '\000\214\000\002',
	PowerPointMShpShapeTypeChart = '\000\214\000\003',
	PowerPointMShpShapeTypeComment = '\000\214\000\004',
	PowerPointMShpShapeTypeFreeForm = '\000\214\000\005',
	PowerPointMShpShapeTypeGroup = '\000\214\000\006',
	PowerPointMShpShapeTypeEmbeddedOLEControl = '\000\214\000\007',
	PowerPointMShpShapeTypeFormControl = '\000\214\000\010',
	PowerPointMShpShapeTypeLine = '\000\214\000\011',
	PowerPointMShpShapeTypeLinkedOLEObject = '\000\214\000\012',
	PowerPointMShpShapeTypeLinkedPicture = '\000\214\000\013',
	PowerPointMShpShapeTypeOLEControl = '\000\214\000\014',
	PowerPointMShpShapeTypePicture = '\000\214\000\015',
	PowerPointMShpShapeTypePlaceHolder = '\000\214\000\016',
	PowerPointMShpShapeTypeWordArt = '\000\214\000\017',
	PowerPointMShpShapeTypeMedia = '\000\214\000\020',
	PowerPointMShpShapeTypeTextBox = '\000\214\000\021',
	PowerPointMShpShapeTypeTable = '\000\214\000\022',
	PowerPointMShpShapeTypeCanvas = '\000\214\000\023',
	PowerPointMShpShapeTypeDiagram = '\000\214\000\024',
	PowerPointMShpShapeTypeInk = '\000\214\000\025',
	PowerPointMShpShapeTypeInkComment = '\000\214\000\026',
	PowerPointMShpShapeTypeSmartartGraphic = '\000\214\000\027',
	PowerPointMShpShapeTypeSlicer = '\000\214\000\030'
};
typedef enum PowerPointMShp PowerPointMShp;

enum PowerPointMCrT {
	PowerPointMCrTColorTypeUnset = '\000j\377\376',
	PowerPointMCrTRGB = '\000k\000\001',
	PowerPointMCrTScheme = '\000k\000\002'
};
typedef enum PowerPointMCrT PowerPointMCrT;

enum PowerPointMPc {
	PowerPointMPcPictureColorTypeUnset = '\000\265\377\376',
	PowerPointMPcPictureColorAutomatic = '\000\266\000\001',
	PowerPointMPcPictureColorGrayScale = '\000\266\000\002',
	PowerPointMPcPictureColorBlackAndWhite = '\000\266\000\003',
	PowerPointMPcPictureColorWatermark = '\000\266\000\004'
};
typedef enum PowerPointMPc PowerPointMPc;

enum PowerPointMCAt {
	PowerPointMCAtAngleUnset = '\000k\377\376',
	PowerPointMCAtAngleAutomatic = '\000l\000\001',
	PowerPointMCAtAngle30 = '\000l\000\002',
	PowerPointMCAtAngle45 = '\000l\000\003',
	PowerPointMCAtAngle60 = '\000l\000\004',
	PowerPointMCAtAngle90 = '\000l\000\005'
};
typedef enum PowerPointMCAt PowerPointMCAt;

enum PowerPointMCDt {
	PowerPointMCDtDropUnset = '\000l\377\376',
	PowerPointMCDtDropCustom = '\000m\000\001',
	PowerPointMCDtDropTop = '\000m\000\002',
	PowerPointMCDtDropCenter = '\000m\000\003',
	PowerPointMCDtDropBottom = '\000m\000\004'
};
typedef enum PowerPointMCDt PowerPointMCDt;

enum PowerPointMCot {
	PowerPointMCotCalloutUnset = '\000m\377\376',
	PowerPointMCotCalloutOne = '\000n\000\001',
	PowerPointMCotCalloutTwo = '\000n\000\002',
	PowerPointMCotCalloutThree = '\000n\000\003',
	PowerPointMCotCalloutFour = '\000n\000\004'
};
typedef enum PowerPointMCot PowerPointMCot;

enum PowerPointTxOr {
	PowerPointTxOrTextOrientationUnset = '\000\215\377\376',
	PowerPointTxOrHorizontal = '\000\216\000\001',
	PowerPointTxOrUpward = '\000\216\000\002',
	PowerPointTxOrDownward = '\000\216\000\003',
	PowerPointTxOrVerticalEastAsian = '\000\216\000\004',
	PowerPointTxOrVertical = '\000\216\000\005',
	PowerPointTxOrHorizontalRotatedEastAsian = '\000\216\000\006'
};
typedef enum PowerPointTxOr PowerPointTxOr;

enum PowerPointMSFr {
	PowerPointMSFrScaleFromTopLeft = '\000o\000\000',
	PowerPointMSFrScaleFromMiddle = '\000o\000\001',
	PowerPointMSFrScaleFromBottomRight = '\000o\000\002'
};
typedef enum PowerPointMSFr PowerPointMSFr;

enum PowerPointMPzC {
	PowerPointMPzCPresetCameraUnset = '\000\256\377\376',
	PowerPointMPzCCameraLegacyObliqueFromTopLeft = '\000\257\000\001',
	PowerPointMPzCCameraLegacyObliqueFromTop = '\000\257\000\002',
	PowerPointMPzCCameraLegacyObliqueFromTopright = '\000\257\000\003',
	PowerPointMPzCCameraLegacyObliqueFromLeft = '\000\257\000\004',
	PowerPointMPzCCameraLegacyObliqueFromFront = '\000\257\000\005',
	PowerPointMPzCCameraLegacyObliqueFromRight = '\000\257\000\006',
	PowerPointMPzCCameraLegacyObliqueFromBottomLeft = '\000\257\000\007',
	PowerPointMPzCCameraLegacyObliqueFromBottom = '\000\257\000\010',
	PowerPointMPzCCameraLegacyObliqueFromBottomRight = '\000\257\000\011',
	PowerPointMPzCCameraLegacyPerspectiveFromTopLeft = '\000\257\000\012',
	PowerPointMPzCCameraLegacyPerspectiveFromTop = '\000\257\000\013',
	PowerPointMPzCCameraLegacyPerspectiveFromTopRight = '\000\257\000\014',
	PowerPointMPzCCameraLegacyPerspectiveFromLeft = '\000\257\000\015',
	PowerPointMPzCCameraLegacyPerspectiveFromFront = '\000\257\000\016',
	PowerPointMPzCCameraLegacyPerspectiveFromRight = '\000\257\000\017',
	PowerPointMPzCCameraLegacyPerspectiveFromBottomLeft = '\000\257\000\020',
	PowerPointMPzCCameraLegacyPerspectiveFromBottom = '\000\257\000\021',
	PowerPointMPzCCameraLegacyPerspectiveFromBottomRight = '\000\257\000\022',
	PowerPointMPzCCameraOrthographic = '\000\257\000\023',
	PowerPointMPzCCameraIsometricFromTopUp = '\000\257\000\024',
	PowerPointMPzCCameraIsometricFromTopDown = '\000\257\000\025',
	PowerPointMPzCCameraIsometricFromBottomUp = '\000\257\000\026',
	PowerPointMPzCCameraIsometricFromBottomDown = '\000\257\000\027',
	PowerPointMPzCCameraIsometricFromLeftUp = '\000\257\000\030',
	PowerPointMPzCCameraIsometricFromLeftDown = '\000\257\000\031',
	PowerPointMPzCCameraIsometricFromRightUp = '\000\257\000\032',
	PowerPointMPzCCameraIsometricFromRightDown = '\000\257\000\033',
	PowerPointMPzCCameraIsometricOffAxis1FromLeft = '\000\257\000\034',
	PowerPointMPzCCameraIsometricOffAxis1FromRight = '\000\257\000\035',
	PowerPointMPzCCameraIsometricOffAxis1FromTop = '\000\257\000\036',
	PowerPointMPzCCameraIsometricOffAxis2FromLeft = '\000\257\000\037',
	PowerPointMPzCCameraIsometricOffAxis2FromRight = '\000\257\000 ',
	PowerPointMPzCCameraIsometricOffAxis2FromTop = '\000\257\000!',
	PowerPointMPzCCameraIsometricOffAxis3FromLeft = '\000\257\000\"',
	PowerPointMPzCCameraIsometricOffAxis3FromRight = '\000\257\000#',
	PowerPointMPzCCameraIsometricOffAxis3FromBottom = '\000\257\000$',
	PowerPointMPzCCameraIsometricOffAxis4FromLeft = '\000\257\000%',
	PowerPointMPzCCameraIsometricOffAxis4FromRight = '\000\257\000&',
	PowerPointMPzCCameraIsometricOffAxis4FromBottom = '\000\257\000\'',
	PowerPointMPzCCameraObliqueFromTopLeft = '\000\257\000(',
	PowerPointMPzCCameraObliqueFromTop = '\000\257\000)',
	PowerPointMPzCCameraObliqueFromTopRight = '\000\257\000*',
	PowerPointMPzCCameraObliqueFromLeft = '\000\257\000+',
	PowerPointMPzCCameraObliqueFromRight = '\000\257\000,',
	PowerPointMPzCCameraObliqueFromBottomLeft = '\000\257\000-',
	PowerPointMPzCCameraObliqueFromBottom = '\000\257\000.',
	PowerPointMPzCCameraObliqueFromBottomRight = '\000\257\000/',
	PowerPointMPzCCameraPerspectiveFromFront = '\000\257\0000',
	PowerPointMPzCCameraPerspectiveFromLeft = '\000\257\0001',
	PowerPointMPzCCameraPerspectiveFromRight = '\000\257\0002',
	PowerPointMPzCCameraPerspectiveFromAbove = '\000\257\0003',
	PowerPointMPzCCameraPerspectiveFromBelow = '\000\257\0004',
	PowerPointMPzCCameraPerspectiveFromAboveFacingLeft = '\000\257\0005',
	PowerPointMPzCCameraPerspectiveFromAboveFacingRight = '\000\257\0006',
	PowerPointMPzCCameraPerspectiveContrastingFacingLeft = '\000\257\0007',
	PowerPointMPzCCameraPerspectiveContrastingFacingRight = '\000\257\0008',
	PowerPointMPzCCameraPerspectiveHeroicFacingLeft = '\000\257\0009',
	PowerPointMPzCCameraPerspectiveHeroicFacingRight = '\000\257\000:',
	PowerPointMPzCCameraPerspectiveHeroicExtremeFacingLeft = '\000\257\000;',
	PowerPointMPzCCameraPerspectiveHeroicExtremeFacingRight = '\000\257\000<',
	PowerPointMPzCCameraPerspectiveRelaxed = '\000\257\000=',
	PowerPointMPzCCameraPerspectiveRelaxedModerately = '\000\257\000>'
};
typedef enum PowerPointMPzC PowerPointMPzC;

enum PowerPointMLtT {
	PowerPointMLtTLightRigUnset = '\000\257\377\376',
	PowerPointMLtTLightRigFlat1 = '\000\260\000\001',
	PowerPointMLtTLightRigFlat2 = '\000\260\000\002',
	PowerPointMLtTLightRigFlat3 = '\000\260\000\003',
	PowerPointMLtTLightRigFlat4 = '\000\260\000\004',
	PowerPointMLtTLightRigNormal1 = '\000\260\000\005',
	PowerPointMLtTLightRigNormal2 = '\000\260\000\006',
	PowerPointMLtTLightRigNormal3 = '\000\260\000\007',
	PowerPointMLtTLightRigNormal4 = '\000\260\000\010',
	PowerPointMLtTLightRigHarsh1 = '\000\260\000\011',
	PowerPointMLtTLightRigHarsh2 = '\000\260\000\012',
	PowerPointMLtTLightRigHarsh3 = '\000\260\000\013',
	PowerPointMLtTLightRigHarsh4 = '\000\260\000\014',
	PowerPointMLtTLightRigThreePoint = '\000\260\000\015',
	PowerPointMLtTLightRigBalanced = '\000\260\000\016',
	PowerPointMLtTLightRigSoft = '\000\260\000\017',
	PowerPointMLtTLightRigHarsh = '\000\260\000\020',
	PowerPointMLtTLightRigFlood = '\000\260\000\021',
	PowerPointMLtTLightRigContrasting = '\000\260\000\022',
	PowerPointMLtTLightRigMorning = '\000\260\000\023',
	PowerPointMLtTLightRigSunrise = '\000\260\000\024',
	PowerPointMLtTLightRigSunset = '\000\260\000\025',
	PowerPointMLtTLightRigChilly = '\000\260\000\026',
	PowerPointMLtTLightRigFreezing = '\000\260\000\027',
	PowerPointMLtTLightRigFlat = '\000\260\000\030',
	PowerPointMLtTLightRigTwoPoint = '\000\260\000\031',
	PowerPointMLtTLightRigGlow = '\000\260\000\032',
	PowerPointMLtTLightRigBrightRoom = '\000\260\000\033'
};
typedef enum PowerPointMLtT PowerPointMLtT;

enum PowerPointMBlT {
	PowerPointMBlTBevelTypeUnset = '\000\260\377\376',
	PowerPointMBlTBevelNone = '\000\261\000\001',
	PowerPointMBlTBevelRelaxedInset = '\000\261\000\002',
	PowerPointMBlTBevelCircle = '\000\261\000\003',
	PowerPointMBlTBevelSlope = '\000\261\000\004',
	PowerPointMBlTBevelCross = '\000\261\000\005',
	PowerPointMBlTBevelAngle = '\000\261\000\006',
	PowerPointMBlTBevelSoftRound = '\000\261\000\007',
	PowerPointMBlTBevelConvex = '\000\261\000\010',
	PowerPointMBlTBevelCoolSlant = '\000\261\000\011',
	PowerPointMBlTBevelDivot = '\000\261\000\012',
	PowerPointMBlTBevelRiblet = '\000\261\000\013',
	PowerPointMBlTBevelHardEdge = '\000\261\000\014',
	PowerPointMBlTBevelArtDeco = '\000\261\000\015'
};
typedef enum PowerPointMBlT PowerPointMBlT;

enum PowerPointMSSt {
	PowerPointMSStShadowStyleUnset = '\000\261\377\376',
	PowerPointMSStShadowStyleInner = '\000\262\000\001',
	PowerPointMSStShadowStyleOuter = '\000\262\000\002'
};
typedef enum PowerPointMSSt PowerPointMSSt;

enum PowerPointPpgA {
	PowerPointPpgAParagraphAlignmentUnset = '\000\346\377\376',
	PowerPointPpgAParagraphAlignLeft = '\000\347\000\000',
	PowerPointPpgAParagraphAlignCenter = '\000\347\000\001',
	PowerPointPpgAParagraphAlignRight = '\000\347\000\002',
	PowerPointPpgAParagraphAlignJustify = '\000\347\000\003',
	PowerPointPpgAParagraphAlignDistribute = '\000\347\000\004',
	PowerPointPpgAParagraphAlignThai = '\000\347\000\005',
	PowerPointPpgAParagraphAlignJustifyLow = '\000\347\000\006'
};
typedef enum PowerPointPpgA PowerPointPpgA;

enum PowerPointMTSt {
	PowerPointMTStStrikeUnset = '\000\263\377\376',
	PowerPointMTStNoStrike = '\000\264\000\000',
	PowerPointMTStSingleStrike = '\000\264\000\001',
	PowerPointMTStDoubleStrike = '\000\264\000\002'
};
typedef enum PowerPointMTSt PowerPointMTSt;

enum PowerPointMTCa {
	PowerPointMTCaCapsUnset = '\000\264\377\376',
	PowerPointMTCaNoCaps = '\000\265\000\000',
	PowerPointMTCaSmallCaps = '\000\265\000\001',
	PowerPointMTCaAllCaps = '\000\265\000\002'
};
typedef enum PowerPointMTCa PowerPointMTCa;

enum PowerPointMTUl {
	PowerPointMTUlUnderlineUnset = '\003\356\377\376',
	PowerPointMTUlNoUnderline = '\003\357\000\000',
	PowerPointMTUlUnderlineWordsOnly = '\003\357\000\001',
	PowerPointMTUlUnderlineSingleLine = '\003\357\000\002',
	PowerPointMTUlUnderlineDoubleLine = '\003\357\000\003',
	PowerPointMTUlUnderlineHeavyLine = '\003\357\000\004',
	PowerPointMTUlUnderlineDottedLine = '\003\357\000\005',
	PowerPointMTUlUnderlineHeavyDottedLine = '\003\357\000\006',
	PowerPointMTUlUnderlineDashLine = '\003\357\000\007',
	PowerPointMTUlUnderlineHeavyDashLine = '\003\357\000\010',
	PowerPointMTUlUnderlineLongDashLine = '\003\357\000\011',
	PowerPointMTUlUnderlineHeavyLongDashLine = '\003\357\000\012',
	PowerPointMTUlUnderlineDotDashLine = '\003\357\000\013',
	PowerPointMTUlUnderlineHeavyDotDashLine = '\003\357\000\014',
	PowerPointMTUlUnderlineDotDotDashLine = '\003\357\000\015',
	PowerPointMTUlUnderlineHeavyDotDotDashLine = '\003\357\000\016',
	PowerPointMTUlUnderlineWavyLine = '\003\357\000\017',
	PowerPointMTUlUnderlineHeavyWavyLine = '\003\357\000\020',
	PowerPointMTUlUnderlineWavyDoubleLine = '\003\357\000\021'
};
typedef enum PowerPointMTUl PowerPointMTUl;

enum PowerPointMTTA {
	PowerPointMTTATabUnset = '\000\266\377\376',
	PowerPointMTTALeftTab = '\000\267\000\000',
	PowerPointMTTACenterTab = '\000\267\000\001',
	PowerPointMTTARightTab = '\000\267\000\002',
	PowerPointMTTADecimalTab = '\000\267\000\003'
};
typedef enum PowerPointMTTA PowerPointMTTA;

enum PowerPointMTCW {
	PowerPointMTCWCharacterWrapUnset = '\000\267\377\376',
	PowerPointMTCWNoCharacterWrap = '\000\270\000\000',
	PowerPointMTCWStandardCharacterWrap = '\000\270\000\001',
	PowerPointMTCWStrictCharacterWrap = '\000\270\000\002',
	PowerPointMTCWCustomCharacterWrap = '\000\270\000\003'
};
typedef enum PowerPointMTCW PowerPointMTCW;

enum PowerPointMTFA {
	PowerPointMTFAFontAlignmentUnset = '\000\270\377\376',
	PowerPointMTFAAutomaticAlignment = '\000\271\000\000',
	PowerPointMTFATopAlignment = '\000\271\000\001',
	PowerPointMTFACenterAlignment = '\000\271\000\002',
	PowerPointMTFABaselineAlignment = '\000\271\000\003',
	PowerPointMTFABottomAlignment = '\000\271\000\004'
};
typedef enum PowerPointMTFA PowerPointMTFA;

enum PowerPointPAtS {
	PowerPointPAtSAutoSizeUnset = '\000\344\377\376',
	PowerPointPAtSAutoSizeNone = '\000\345\000\000',
	PowerPointPAtSShapeToFitText = '\000\345\000\001',
	PowerPointPAtSTextToFitShape = '\000\345\000\002'
};
typedef enum PowerPointPAtS PowerPointPAtS;

enum PowerPointMPFo {
	PowerPointMPFoPathTypeUnset = '\000\272\377\376',
	PowerPointMPFoNoPathType = '\000\273\000\000',
	PowerPointMPFoPathType1 = '\000\273\000\001',
	PowerPointMPFoPathType2 = '\000\273\000\002',
	PowerPointMPFoPathType3 = '\000\273\000\003',
	PowerPointMPFoPathType4 = '\000\273\000\004'
};
typedef enum PowerPointMPFo PowerPointMPFo;

enum PowerPointMWFo {
	PowerPointMWFoWarpFormatUnset = '\000\273\377\376',
	PowerPointMWFoWarpFormat1 = '\000\274\000\000',
	PowerPointMWFoWarpFormat2 = '\000\274\000\001',
	PowerPointMWFoWarpFormat3 = '\000\274\000\002',
	PowerPointMWFoWarpFormat4 = '\000\274\000\003',
	PowerPointMWFoWarpFormat5 = '\000\274\000\004',
	PowerPointMWFoWarpFormat6 = '\000\274\000\005',
	PowerPointMWFoWarpFormat7 = '\000\274\000\006',
	PowerPointMWFoWarpFormat8 = '\000\274\000\007',
	PowerPointMWFoWarpFormat9 = '\000\274\000\010',
	PowerPointMWFoWarpFormat10 = '\000\274\000\011',
	PowerPointMWFoWarpFormat11 = '\000\274\000\012',
	PowerPointMWFoWarpFormat12 = '\000\274\000\013',
	PowerPointMWFoWarpFormat13 = '\000\274\000\014',
	PowerPointMWFoWarpFormat14 = '\000\274\000\015',
	PowerPointMWFoWarpFormat15 = '\000\274\000\016',
	PowerPointMWFoWarpFormat16 = '\000\274\000\017',
	PowerPointMWFoWarpFormat17 = '\000\274\000\020',
	PowerPointMWFoWarpFormat18 = '\000\274\000\021',
	PowerPointMWFoWarpFormat19 = '\000\274\000\022',
	PowerPointMWFoWarpFormat20 = '\000\274\000\023',
	PowerPointMWFoWarpFormat21 = '\000\274\000\024',
	PowerPointMWFoWarpFormat22 = '\000\274\000\025',
	PowerPointMWFoWarpFormat23 = '\000\274\000\026',
	PowerPointMWFoWarpFormat24 = '\000\274\000\027',
	PowerPointMWFoWarpFormat25 = '\000\274\000\030',
	PowerPointMWFoWarpFormat26 = '\000\274\000\031',
	PowerPointMWFoWarpFormat27 = '\000\274\000\032',
	PowerPointMWFoWarpFormat28 = '\000\274\000\033',
	PowerPointMWFoWarpFormat29 = '\000\274\000\034',
	PowerPointMWFoWarpFormat30 = '\000\274\000\035',
	PowerPointMWFoWarpFormat31 = '\000\274\000\036',
	PowerPointMWFoWarpFormat32 = '\000\274\000\037',
	PowerPointMWFoWarpFormat33 = '\000\274\000 ',
	PowerPointMWFoWarpFormat34 = '\000\274\000!',
	PowerPointMWFoWarpFormat35 = '\000\274\000\"',
	PowerPointMWFoWarpFormat36 = '\000\274\000#'
};
typedef enum PowerPointMWFo PowerPointMWFo;

enum PowerPointPCgC {
	PowerPointPCgCCaseSentence = '\000\344\000\001',
	PowerPointPCgCCaseLower = '\000\344\000\002',
	PowerPointPCgCCaseUpper = '\000\344\000\003',
	PowerPointPCgCCaseTitle = '\000\344\000\004',
	PowerPointPCgCCaseToggle = '\000\344\000\005'
};
typedef enum PowerPointPCgC PowerPointPCgC;

enum PowerPointMDTF {
	PowerPointMDTFDateTimeFormatUnset = '\000\275\377\376',
	PowerPointMDTFDateTimeFormatMdyy = '\000\276\000\001',
	PowerPointMDTFDateTimeFormatDdddMMMMddyyyy = '\000\276\000\002',
	PowerPointMDTFDateTimeFormatDMMMMyyyy = '\000\276\000\003',
	PowerPointMDTFDateTimeFormatMMMMdyyyy = '\000\276\000\004',
	PowerPointMDTFDateTimeFormatDMMMyy = '\000\276\000\005',
	PowerPointMDTFDateTimeFormatMMMMyy = '\000\276\000\006',
	PowerPointMDTFDateTimeFormatMMyy = '\000\276\000\007',
	PowerPointMDTFDateTimeFormatMMddyyHmm = '\000\276\000\010',
	PowerPointMDTFDateTimeFormatMMddyyhmmAMPM = '\000\276\000\011',
	PowerPointMDTFDateTimeFormatHmm = '\000\276\000\012',
	PowerPointMDTFDateTimeFormatHmmss = '\000\276\000\013',
	PowerPointMDTFDateTimeFormatHmmAMPM = '\000\276\000\014',
	PowerPointMDTFDateTimeFormatHmmssAMPM = '\000\276\000\015',
	PowerPointMDTFDateTimeFormatFigureOut = '\000\276\000\016'
};
typedef enum PowerPointMDTF PowerPointMDTF;

enum PowerPointMSET {
	PowerPointMSETSoftEdgeUnset = '\000\277\377\376',
	PowerPointMSETNoSoftEdge = '\000\300\000\000',
	PowerPointMSETSoftEdgeType1 = '\000\300\000\001',
	PowerPointMSETSoftEdgeType2 = '\000\300\000\002',
	PowerPointMSETSoftEdgeType3 = '\000\300\000\003',
	PowerPointMSETSoftEdgeType4 = '\000\300\000\004',
	PowerPointMSETSoftEdgeType5 = '\000\300\000\005',
	PowerPointMSETSoftEdgeType6 = '\000\300\000\006'
};
typedef enum PowerPointMSET PowerPointMSET;

enum PowerPointMCSI {
	PowerPointMCSIFirstDarkSchemeColor = '\000\301\000\001',
	PowerPointMCSIFirstLightSchemeColor = '\000\301\000\002',
	PowerPointMCSISecondDarkSchemeColor = '\000\301\000\003',
	PowerPointMCSISecondLightSchemeColor = '\000\301\000\004',
	PowerPointMCSIFirstAccentSchemeColor = '\000\301\000\005',
	PowerPointMCSISecondAccentSchemeColor = '\000\301\000\006',
	PowerPointMCSIThirdAccentSchemeColor = '\000\301\000\007',
	PowerPointMCSIFourthAccentSchemeColor = '\000\301\000\010',
	PowerPointMCSIFifthAccentSchemeColor = '\000\301\000\011',
	PowerPointMCSISixthAccentSchemeColor = '\000\301\000\012',
	PowerPointMCSIHyperlinkSchemeColor = '\000\301\000\013',
	PowerPointMCSIFollowedHyperlinkSchemeColor = '\000\301\000\014'
};
typedef enum PowerPointMCSI PowerPointMCSI;

enum PowerPointMCoI {
	PowerPointMCoIThemeColorUnset = '\000\301\377\376',
	PowerPointMCoINoThemeColor = '\000\302\000\000',
	PowerPointMCoIFirstDarkThemeColor = '\000\302\000\001',
	PowerPointMCoIFirstLightThemeColor = '\000\302\000\002',
	PowerPointMCoISecondDarkThemeColor = '\000\302\000\003',
	PowerPointMCoISecondLightThemeColor = '\000\302\000\004',
	PowerPointMCoIFirstAccentThemeColor = '\000\302\000\005',
	PowerPointMCoISecondAccentThemeColor = '\000\302\000\006',
	PowerPointMCoIThirdAccentThemeColor = '\000\302\000\007',
	PowerPointMCoIFourthAccentThemeColor = '\000\302\000\010',
	PowerPointMCoIFifthAccentThemeColor = '\000\302\000\011',
	PowerPointMCoISixthAccentThemeColor = '\000\302\000\012',
	PowerPointMCoIHyperlinkThemeColor = '\000\302\000\013',
	PowerPointMCoIFollowedHyperlinkThemeColor = '\000\302\000\014',
	PowerPointMCoIFirstTextThemeColor = '\000\302\000\015',
	PowerPointMCoIFirstBackgroundThemeColor = '\000\302\000\016',
	PowerPointMCoISecondTextThemeColor = '\000\302\000\017',
	PowerPointMCoISecondBackgroundThemeColor = '\000\302\000\020'
};
typedef enum PowerPointMCoI PowerPointMCoI;

enum PowerPointMFLI {
	PowerPointMFLIThemeFontLatin = '\000\303\000\001',
	PowerPointMFLIThemeFontComplexScript = '\000\303\000\002',
	PowerPointMFLIThemeFontHighAnsi = '\000\303\000\003',
	PowerPointMFLIThemeFontEastAsian = '\000\303\000\004'
};
typedef enum PowerPointMFLI PowerPointMFLI;

enum PowerPointMSSI {
	PowerPointMSSIShapeStyleUnset = '\000\303\377\376',
	PowerPointMSSIShapeNotAPreset = '\000\304\000\000',
	PowerPointMSSIShapePreset1 = '\000\304\000\001',
	PowerPointMSSIShapePreset2 = '\000\304\000\002',
	PowerPointMSSIShapePreset3 = '\000\304\000\003',
	PowerPointMSSIShapePreset4 = '\000\304\000\004',
	PowerPointMSSIShapePreset5 = '\000\304\000\005',
	PowerPointMSSIShapePreset6 = '\000\304\000\006',
	PowerPointMSSIShapePreset7 = '\000\304\000\007',
	PowerPointMSSIShapePreset8 = '\000\304\000\010',
	PowerPointMSSIShapePreset9 = '\000\304\000\011',
	PowerPointMSSIShapePreset10 = '\000\304\000\012',
	PowerPointMSSIShapePreset11 = '\000\304\000\013',
	PowerPointMSSIShapePreset12 = '\000\304\000\014',
	PowerPointMSSIShapePreset13 = '\000\304\000\015',
	PowerPointMSSIShapePreset14 = '\000\304\000\016',
	PowerPointMSSIShapePreset15 = '\000\304\000\017',
	PowerPointMSSIShapePreset16 = '\000\304\000\020',
	PowerPointMSSIShapePreset17 = '\000\304\000\021',
	PowerPointMSSIShapePreset18 = '\000\304\000\022',
	PowerPointMSSIShapePreset19 = '\000\304\000\023',
	PowerPointMSSIShapePreset20 = '\000\304\000\024',
	PowerPointMSSIShapePreset21 = '\000\304\000\025',
	PowerPointMSSIShapePreset22 = '\000\304\000\026',
	PowerPointMSSIShapePreset23 = '\000\304\000\027',
	PowerPointMSSIShapePreset24 = '\000\304\000\030',
	PowerPointMSSIShapePreset25 = '\000\304\000\031',
	PowerPointMSSIShapePreset26 = '\000\304\000\032',
	PowerPointMSSIShapePreset27 = '\000\304\000\033',
	PowerPointMSSIShapePreset28 = '\000\304\000\034',
	PowerPointMSSIShapePreset29 = '\000\304\000\035',
	PowerPointMSSIShapePreset30 = '\000\304\000\036',
	PowerPointMSSIShapePreset31 = '\000\304\000\037',
	PowerPointMSSIShapePreset32 = '\000\304\000 ',
	PowerPointMSSIShapePreset33 = '\000\304\000!',
	PowerPointMSSIShapePreset34 = '\000\304\000\"',
	PowerPointMSSIShapePreset35 = '\000\304\000#',
	PowerPointMSSIShapePreset36 = '\000\304\000$',
	PowerPointMSSIShapePreset37 = '\000\304\000%',
	PowerPointMSSIShapePreset38 = '\000\304\000&',
	PowerPointMSSIShapePreset39 = '\000\304\000\'',
	PowerPointMSSIShapePreset40 = '\000\304\000(',
	PowerPointMSSIShapePreset41 = '\000\304\000)',
	PowerPointMSSIShapePreset42 = '\000\304\000*',
	PowerPointMSSILinePreset1 = '\000\304\'\021',
	PowerPointMSSILinePreset2 = '\000\304\'\022',
	PowerPointMSSILinePreset3 = '\000\304\'\023',
	PowerPointMSSILinePreset4 = '\000\304\'\024',
	PowerPointMSSILinePreset5 = '\000\304\'\025',
	PowerPointMSSILinePreset6 = '\000\304\'\026',
	PowerPointMSSILinePreset7 = '\000\304\'\027',
	PowerPointMSSILinePreset8 = '\000\304\'\030',
	PowerPointMSSILinePreset9 = '\000\304\'\031',
	PowerPointMSSILinePreset10 = '\000\304\'\032',
	PowerPointMSSILinePreset11 = '\000\304\'\033',
	PowerPointMSSILinePreset12 = '\000\304\'\034',
	PowerPointMSSILinePreset13 = '\000\304\'\035',
	PowerPointMSSILinePreset14 = '\000\304\'\036',
	PowerPointMSSILinePreset15 = '\000\304\'\037',
	PowerPointMSSILinePreset16 = '\000\304\' ',
	PowerPointMSSILinePreset17 = '\000\304\'!',
	PowerPointMSSILinePreset18 = '\000\304\'\"',
	PowerPointMSSILinePreset19 = '\000\304\'#',
	PowerPointMSSILinePreset20 = '\000\304\'$',
	PowerPointMSSILinePreset21 = '\000\304\'%'
};
typedef enum PowerPointMSSI PowerPointMSSI;

enum PowerPointMBSI {
	PowerPointMBSIBackgroundUnset = '\000\304\377\376',
	PowerPointMBSIBackgroundNotAPreset = '\000\305\000\000',
	PowerPointMBSIBackgroundPreset1 = '\000\305\000\001',
	PowerPointMBSIBackgroundPreset2 = '\000\305\000\002',
	PowerPointMBSIBackgroundPreset3 = '\000\305\000\003',
	PowerPointMBSIBackgroundPreset4 = '\000\305\000\004',
	PowerPointMBSIBackgroundPreset5 = '\000\305\000\005',
	PowerPointMBSIBackgroundPreset6 = '\000\305\000\006',
	PowerPointMBSIBackgroundPreset7 = '\000\305\000\007',
	PowerPointMBSIBackgroundPreset8 = '\000\305\000\010',
	PowerPointMBSIBackgroundPreset9 = '\000\305\000\011',
	PowerPointMBSIBackgroundPreset10 = '\000\305\000\012',
	PowerPointMBSIBackgroundPreset11 = '\000\305\000\013',
	PowerPointMBSIBackgroundPreset12 = '\000\305\000\014'
};
typedef enum PowerPointMBSI PowerPointMBSI;

enum PowerPointPDrT {
	PowerPointPDrTTextDirectionUnset = '\000\352\377\376',
	PowerPointPDrTLeftToRight = '\000\353\000\001',
	PowerPointPDrTRightToLeft = '\000\353\000\002'
};
typedef enum PowerPointPDrT PowerPointPDrT;

enum PowerPointPBtT {
	PowerPointPBtTBulletTypeUnset = '\000\347\377\376',
	PowerPointPBtTBulletTypeNone = '\000\350\000\000',
	PowerPointPBtTBulletTypeUnnumbered = '\000\350\000\001',
	PowerPointPBtTBulletTypeNumbered = '\000\350\000\002',
	PowerPointPBtTPictureBulletType = '\000\350\000\003'
};
typedef enum PowerPointPBtT PowerPointPBtT;

enum PowerPointPBtS {
	PowerPointPBtSNumberedBulletStyleUnset = '\000\350\377\376',
	PowerPointPBtSNumberedBulletStyleAlphaLowercasePeriod = '\000\351\000\000',
	PowerPointPBtSNumberedBulletStyleAlphaUppercasePeriod = '\000\351\000\001',
	PowerPointPBtSNumberedBulletStyleArabicRightParen = '\000\351\000\002',
	PowerPointPBtSNumberedBulletStyleArabicPeriod = '\000\351\000\003',
	PowerPointPBtSNumberedBulletStyleRomanLowercaseParenBoth = '\000\351\000\004',
	PowerPointPBtSNumberedBulletStyleRomanLowercaseParenRight = '\000\351\000\005',
	PowerPointPBtSNumberedBulletStyleRomanLowercasePeriod = '\000\351\000\006',
	PowerPointPBtSNumberedBulletStyleRomanUppercasePeriod = '\000\351\000\007',
	PowerPointPBtSNumberedBulletStyleAlphaLowercaseParenBoth = '\000\351\000\010',
	PowerPointPBtSNumberedBulletStyleAlphaLowercaseParenRight = '\000\351\000\011',
	PowerPointPBtSNumberedBulletStyleAlphaUppercaseParenBoth = '\000\351\000\012',
	PowerPointPBtSNumberedBulletStyleAlphaUppercaseParenRight = '\000\351\000\013',
	PowerPointPBtSNumberedBulletStyleArabicParenBoth = '\000\351\000\014',
	PowerPointPBtSNumberedBulletStyleArabicPlain = '\000\351\000\015',
	PowerPointPBtSNumberedBulletStyleRomanUppercaseParenBoth = '\000\351\000\016',
	PowerPointPBtSNumberedBulletStyleRomanUppercaseParenRight = '\000\351\000\017',
	PowerPointPBtSNumberedBulletStyleSimplifiedChinesePlain = '\000\351\000\020',
	PowerPointPBtSNumberedBulletStyleSimplifiedChinesePeriod = '\000\351\000\021',
	PowerPointPBtSNumberedBulletStyleCircleNumberPlain = '\000\351\000\022',
	PowerPointPBtSNumberedBulletStyleCircleNumberWhitePlain = '\000\351\000\023',
	PowerPointPBtSNumberedBulletStyleCircleNumberBlackPlain = '\000\351\000\024',
	PowerPointPBtSNumberedBulletStyleTraditionalChinesePlain = '\000\351\000\025',
	PowerPointPBtSNumberedBulletStyleTraditionalChinesePeriod = '\000\351\000\026',
	PowerPointPBtSNumberedBulletStyleArabicAlphaDash = '\000\351\000\027',
	PowerPointPBtSNumberedBulletStyleArabicAbjadDash = '\000\351\000\030',
	PowerPointPBtSNumberedBulletStyleHebrewAlphaDash = '\000\351\000\031',
	PowerPointPBtSNumberedBulletStyleKanjiKoreanPlain = '\000\351\000\032',
	PowerPointPBtSNumberedBulletStyleKanjiKoreanPeriod = '\000\351\000\033',
	PowerPointPBtSNumberedBulletStyleArabicDBPlain = '\000\351\000\034',
	PowerPointPBtSNumberedBulletStyleArabicDBPeriod = '\000\351\000\035',
	PowerPointPBtSNumberedBulletStyleThaiAlphaPeriod = '\000\351\000\036',
	PowerPointPBtSNumberedBulletStyleThaiAlphaParenRight = '\000\351\000\037',
	PowerPointPBtSNumberedBulletStyleThaiAlphaParenBoth = '\000\351\000 ',
	PowerPointPBtSNumberedBulletStyleThaiNumberPeriod = '\000\351\000!',
	PowerPointPBtSNumberedBulletStyleThaiNumberParenRight = '\000\351\000\"',
	PowerPointPBtSNumberedBulletStyleThaiParenBoth = '\000\351\000#',
	PowerPointPBtSNumberedBulletStyleHindiAlphaPeriod = '\000\351\000$',
	PowerPointPBtSNumberedBulletStyleHindiNumberPeriod = '\000\351\000%',
	PowerPointPBtSNumberedBulletStyleKanjiSimpifiedChineseDBPeriod = '\000\351\000&',
	PowerPointPBtSNumberedBulletStyleHindiNumberParenRight = '\000\351\000\'',
	PowerPointPBtSNumberedBulletStyleHindiAlpha1Period = '\000\351\000('
};
typedef enum PowerPointPBtS PowerPointPBtS;

enum PowerPointPTSp {
	PowerPointPTSpTabstopUnset = '\000\364\377\376',
	PowerPointPTSpTabstopLeft = '\000\365\000\001',
	PowerPointPTSpTabstopCenter = '\000\365\000\002',
	PowerPointPTSpTabstopRight = '\000\365\000\003',
	PowerPointPTSpTabstopDecimal = '\000\365\000\004'
};
typedef enum PowerPointPTSp PowerPointPTSp;

enum PowerPointMRfT {
	PowerPointMRfTReflectionUnset = '\003\351\377\376',
	PowerPointMRfTReflectionTypeNone = '\003\352\000\000',
	PowerPointMRfTReflectionType1 = '\003\352\000\001',
	PowerPointMRfTReflectionType2 = '\003\352\000\002',
	PowerPointMRfTReflectionType3 = '\003\352\000\003',
	PowerPointMRfTReflectionType4 = '\003\352\000\004',
	PowerPointMRfTReflectionType5 = '\003\352\000\005',
	PowerPointMRfTReflectionType6 = '\003\352\000\006',
	PowerPointMRfTReflectionType7 = '\003\352\000\007',
	PowerPointMRfTReflectionType8 = '\003\352\000\010',
	PowerPointMRfTReflectionType9 = '\003\352\000\011'
};
typedef enum PowerPointMRfT PowerPointMRfT;

enum PowerPointMTtA {
	PowerPointMTtATextureUnset = '\003\352\377\376',
	PowerPointMTtATextureTopLeft = '\003\353\000\000',
	PowerPointMTtATextureTop = '\003\353\000\001',
	PowerPointMTtATextureTopRight = '\003\353\000\002',
	PowerPointMTtATextureLeft = '\003\353\000\003',
	PowerPointMTtATextureCenter = '\003\353\000\004',
	PowerPointMTtATextureRight = '\003\353\000\005',
	PowerPointMTtATextureBottomLeft = '\003\353\000\006',
	PowerPointMTtATextureBotton = '\003\353\000\007',
	PowerPointMTtATextureBottomRight = '\003\353\000\010'
};
typedef enum PowerPointMTtA PowerPointMTtA;

enum PowerPointPBlA {
	PowerPointPBlATextBaselineAlignmentUnset = '\003\353\377\376',
	PowerPointPBlATextBaselineAlignBaseline = '\003\354\000\001',
	PowerPointPBlATextBaselineAlignTop = '\003\354\000\002',
	PowerPointPBlATextBaselineAlignCenter = '\003\354\000\003',
	PowerPointPBlATextBaselineAlignEastAsian50 = '\003\354\000\004',
	PowerPointPBlATextBaselineAlignAutomatic = '\003\354\000\005'
};
typedef enum PowerPointPBlA PowerPointPBlA;

enum PowerPointMCbF {
	PowerPointMCbFClipboardFormatUnset = '\003\354\377\376',
	PowerPointMCbFNativeClipboardFormat = '\003\355\000\001',
	PowerPointMCbFHTMlClipboardFormat = '\003\355\000\002',
	PowerPointMCbFRTFClipboardFormat = '\003\355\000\003',
	PowerPointMCbFPlainTextClipboardFormat = '\003\355\000\004'
};
typedef enum PowerPointMCbF PowerPointMCbF;

enum PowerPointMTiP {
	PowerPointMTiPInsertBefore = '\003\356\000\000',
	PowerPointMTiPInsertAfter = '\003\356\000\001'
};
typedef enum PowerPointMTiP PowerPointMTiP;

enum PowerPointMPiT {
	PowerPointMPiTSaveAsDefault = '\003\362\377\376',
	PowerPointMPiTSaveAsPNGFile = '\003\363\000\000',
	PowerPointMPiTSaveAsBMPFile = '\003\363\000\001',
	PowerPointMPiTSaveAsGIFFile = '\003\363\000\002',
	PowerPointMPiTSaveAsJPGFile = '\003\363\000\003',
	PowerPointMPiTSaveAsPDFFile = '\003\363\000\004'
};
typedef enum PowerPointMPiT PowerPointMPiT;

enum PowerPointMPeT {
	PowerPointMPeTNoEffect = '\003\364\000\000',
	PowerPointMPeTEffectBackgroundRemoval = '\003\364\000\001',
	PowerPointMPeTEffectBlur = '\003\364\000\002',
	PowerPointMPeTEffectBrightnessContrast = '\003\364\000\003',
	PowerPointMPeTEffectCement = '\003\364\000\004',
	PowerPointMPeTEffectCrisscrossEtching = '\003\364\000\005',
	PowerPointMPeTEffectChalkSketch = '\003\364\000\006',
	PowerPointMPeTEffectColorTemperature = '\003\364\000\007',
	PowerPointMPeTEffectCutout = '\003\364\000\010',
	PowerPointMPeTEffectFilmGrain = '\003\364\000\011',
	PowerPointMPeTEffectGlass = '\003\364\000\012',
	PowerPointMPeTEffectGlowDiffused = '\003\364\000\013',
	PowerPointMPeTEffectGlowEdges = '\003\364\000\014',
	PowerPointMPeTEffectLightScreen = '\003\364\000\015',
	PowerPointMPeTEffectLineDrawing = '\003\364\000\016',
	PowerPointMPeTEffectMarker = '\003\364\000\017',
	PowerPointMPeTEffectMosiaicBubbles = '\003\364\000\020',
	PowerPointMPeTEffectPaintBrush = '\003\364\000\021',
	PowerPointMPeTEffectPaintStrokes = '\003\364\000\022',
	PowerPointMPeTEffectPastelsSmooth = '\003\364\000\023',
	PowerPointMPeTEffectPencilGrayscale = '\003\364\000\024',
	PowerPointMPeTEffectPencilSketch = '\003\364\000\025',
	PowerPointMPeTEffectPhotocopy = '\003\364\000\026',
	PowerPointMPeTEffectPlasticWrap = '\003\364\000\027',
	PowerPointMPeTEffectSaturation = '\003\364\000\030',
	PowerPointMPeTEffectSharpenSoften = '\003\364\000\031',
	PowerPointMPeTEffectTexturizer = '\003\364\000\032',
	PowerPointMPeTEffectWatercolorSponge = '\003\364\000\033'
};
typedef enum PowerPointMPeT PowerPointMPeT;

enum PowerPointMSgT {
	PowerPointMSgTLine = '\000\217\000\000',
	PowerPointMSgTCurve = '\000\217\000\001'
};
typedef enum PowerPointMSgT PowerPointMSgT;

enum PowerPointMEdT {
	PowerPointMEdTAuto = '\000\220\000\000',
	PowerPointMEdTCorner = '\000\220\000\001',
	PowerPointMEdTSmooth = '\000\220\000\002',
	PowerPointMEdTSymmetric = '\000\220\000\003'
};
typedef enum PowerPointMEdT PowerPointMEdT;

enum PowerPointSANP {
	PowerPointSANPDefaultNodePosition = '\003\365\000\001',
	PowerPointSANPAfterNode = '\003\365\000\002',
	PowerPointSANPBeforeNode = '\003\365\000\003',
	PowerPointSANPAboveNode = '\003\365\000\004',
	PowerPointSANPBelowNode = '\003\365\000\005'
};
typedef enum PowerPointSANP PowerPointSANP;

enum PowerPointSANT {
	PowerPointSANTDefaultNode = '\003\366\000\001',
	PowerPointSANTAssistantNode = '\003\366\000\002'
};
typedef enum PowerPointSANT PowerPointSANT;

enum PowerPointMOCT {
	PowerPointMOCTOrgChartLayoutUnset = '\003\366\377\376',
	PowerPointMOCTOrgChartLayoutStandard = '\003\367\000\001',
	PowerPointMOCTOrgChartLayoutBothHanging = '\003\367\000\002',
	PowerPointMOCTOrgChartLayoutLeftHanging = '\003\367\000\003',
	PowerPointMOCTOrgChartLayoutRightHanging = '\003\367\000\004',
	PowerPointMOCTOrgChartLayoutDefault = '\003\367\000\005'
};
typedef enum PowerPointMOCT PowerPointMOCT;

enum PowerPointMAlC {
	PowerPointMAlCAlignLefts = '\000\000\000\000',
	PowerPointMAlCAlignCenters = '\000\000\000\001',
	PowerPointMAlCAlignRights = '\000\000\000\002',
	PowerPointMAlCAlignTops = '\000\000\000\003',
	PowerPointMAlCAlignMiddles = '\000\000\000\004',
	PowerPointMAlCAlignBottoms = '\000\000\000\005'
};
typedef enum PowerPointMAlC PowerPointMAlC;

enum PowerPointMDsC {
	PowerPointMDsCDistributeHorizontally = '\000\000\000\000',
	PowerPointMDsCDistributeVertically = '\000\000\000\001'
};
typedef enum PowerPointMDsC PowerPointMDsC;

enum PowerPointMOrT {
	PowerPointMOrTOrientationUnset = '\000\214\377\376',
	PowerPointMOrTHorizontalOrientation = '\000\215\000\001',
	PowerPointMOrTVerticalOrientation = '\000\215\000\002'
};
typedef enum PowerPointMOrT PowerPointMOrT;

enum PowerPointMZoC {
	PowerPointMZoCBringShapeToFront = '\000p\000\000',
	PowerPointMZoCSendShapeToBack = '\000p\000\001',
	PowerPointMZoCBringShapeForward = '\000p\000\002',
	PowerPointMZoCSendShapeBackward = '\000p\000\003',
	PowerPointMZoCBringShapeInFrontOfText = '\000p\000\004',
	PowerPointMZoCSendShapeBehindText = '\000p\000\005'
};
typedef enum PowerPointMZoC PowerPointMZoC;

enum PowerPointMFlC {
	PowerPointMFlCFlipHorizontal = '\000q\000\000',
	PowerPointMFlCFlipVertical = '\000q\000\001'
};
typedef enum PowerPointMFlC PowerPointMFlC;

enum PowerPointMTri {
	PowerPointMTriTrue = '\000\240\377\377',
	PowerPointMTriFalse = '\000\241\000\000',
	PowerPointMTriCTrue = '\000\241\000\001',
	PowerPointMTriToggle = '\000\240\377\375',
	PowerPointMTriTriStateUnset = '\000\240\377\376'
};
typedef enum PowerPointMTri PowerPointMTri;

enum PowerPointMBW {
	PowerPointMBWBlackAndWhiteUnset = '\000\254\377\376',
	PowerPointMBWBlackAndWhiteModeAutomatic = '\000\255\000\001',
	PowerPointMBWBlackAndWhiteModeGrayScale = '\000\255\000\002',
	PowerPointMBWBlackAndWhiteModeLightGrayScale = '\000\255\000\003',
	PowerPointMBWBlackAndWhiteModeInverseGrayScale = '\000\255\000\004',
	PowerPointMBWBlackAndWhiteModeGrayOutline = '\000\255\000\005',
	PowerPointMBWBlackAndWhiteModeBlackTextAndLine = '\000\255\000\006',
	PowerPointMBWBlackAndWhiteModeHighContrast = '\000\255\000\007',
	PowerPointMBWBlackAndWhiteModeBlack = '\000\255\000\010',
	PowerPointMBWBlackAndWhiteModeWhite = '\000\255\000\011',
	PowerPointMBWBlackAndWhiteModeDontShow = '\000\255\000\012'
};
typedef enum PowerPointMBW PowerPointMBW;

enum PowerPointMBPS {
	PowerPointMBPSBarLeft = '\000r\000\000',
	PowerPointMBPSBarTop = '\000r\000\001',
	PowerPointMBPSBarRight = '\000r\000\002',
	PowerPointMBPSBarBottom = '\000r\000\003',
	PowerPointMBPSBarFloating = '\000r\000\004',
	PowerPointMBPSBarPopUp = '\000r\000\005',
	PowerPointMBPSBarMenu = '\000r\000\006'
};
typedef enum PowerPointMBPS PowerPointMBPS;

enum PowerPointMBPt {
	PowerPointMBPtNoProtection = '\000s\000\000',
	PowerPointMBPtNoCustomize = '\000s\000\001',
	PowerPointMBPtNoResize = '\000s\000\002',
	PowerPointMBPtNoMove = '\000s\000\004',
	PowerPointMBPtNoChangeVisible = '\000s\000\010',
	PowerPointMBPtNoChangeDock = '\000s\000\020',
	PowerPointMBPtNoVerticalDock = '\000s\000 ',
	PowerPointMBPtNoHorizontalDock = '\000s\000@'
};
typedef enum PowerPointMBPt PowerPointMBPt;

enum PowerPointMBTY {
	PowerPointMBTYNormalCommandBar = '\000t\000\000',
	PowerPointMBTYMenubarCommandBar = '\000t\000\001',
	PowerPointMBTYPopupCommandBar = '\000t\000\002'
};
typedef enum PowerPointMBTY PowerPointMBTY;

enum PowerPointMCLT {
	PowerPointMCLTControlCustom = '\000u\000\000',
	PowerPointMCLTControlButton = '\000u\000\001',
	PowerPointMCLTControlEdit = '\000u\000\002',
	PowerPointMCLTControlDropDown = '\000u\000\003',
	PowerPointMCLTControlCombobox = '\000u\000\004',
	PowerPointMCLTButtonDropDown = '\000u\000\005',
	PowerPointMCLTSplitDropDown = '\000u\000\006',
	PowerPointMCLTOCXDropDown = '\000u\000\007',
	PowerPointMCLTGenericDropDown = '\000u\000\010',
	PowerPointMCLTGraphicDropDown = '\000u\000\011',
	PowerPointMCLTControlPopup = '\000u\000\012',
	PowerPointMCLTGraphicPopup = '\000u\000\013',
	PowerPointMCLTButtonPopup = '\000u\000\014',
	PowerPointMCLTSplitButtonPopup = '\000u\000\015',
	PowerPointMCLTSplitButtonMRUPopup = '\000u\000\016',
	PowerPointMCLTControlLabel = '\000u\000\017',
	PowerPointMCLTExpandingGrid = '\000u\000\020',
	PowerPointMCLTSplitExpandingGrid = '\000u\000\021',
	PowerPointMCLTControlGrid = '\000u\000\022',
	PowerPointMCLTControlGauge = '\000u\000\023',
	PowerPointMCLTGraphicCombobox = '\000u\000\024',
	PowerPointMCLTControlPane = '\000u\000\025',
	PowerPointMCLTActiveX = '\000u\000\026',
	PowerPointMCLTControlGroup = '\000u\000\027',
	PowerPointMCLTControlTab = '\000u\000\030',
	PowerPointMCLTControlSpinner = '\000u\000\031'
};
typedef enum PowerPointMCLT PowerPointMCLT;

enum PowerPointMBns {
	PowerPointMBnsButtonStateUp = '\000v\000\000',
	PowerPointMBnsButtonStateDown = '\000u\377\377',
	PowerPointMBnsButtonStateUnset = '\000v\000\002'
};
typedef enum PowerPointMBns PowerPointMBns;

enum PowerPointMcOu {
	PowerPointMcOuNeither = '\000w\000\000',
	PowerPointMcOuServer = '\000w\000\001',
	PowerPointMcOuClient = '\000w\000\002',
	PowerPointMcOuBoth = '\000w\000\003'
};
typedef enum PowerPointMcOu PowerPointMcOu;

enum PowerPointMBTs {
	PowerPointMBTsButtonAutomatic = '\000x\000\000',
	PowerPointMBTsButtonIcon = '\000x\000\001',
	PowerPointMBTsButtonCaption = '\000x\000\002',
	PowerPointMBTsButtonIconAndCaption = '\000x\000\003'
};
typedef enum PowerPointMBTs PowerPointMBTs;

enum PowerPointMXcb {
	PowerPointMXcbComboboxStyleNormal = '\000y\000\000',
	PowerPointMXcbComboboxStyleLabel = '\000y\000\001'
};
typedef enum PowerPointMXcb PowerPointMXcb;

enum PowerPointMMNA {
	PowerPointMMNANone = '\000{\000\000',
	PowerPointMMNARandom = '\000{\000\001',
	PowerPointMMNAUnfold = '\000{\000\002',
	PowerPointMMNASlide = '\000{\000\003'
};
typedef enum PowerPointMMNA PowerPointMMNA;

enum PowerPointMHlT {
	PowerPointMHlTHyperlinkTypeTextRange = '\000\226\000\000',
	PowerPointMHlTHyperlinkTypeShape = '\000\226\000\001',
	PowerPointMHlTHyperlinkTypeInlineShape = '\000\226\000\002'
};
typedef enum PowerPointMHlT PowerPointMHlT;

enum PowerPointMXiM {
	PowerPointMXiMAppendString = '\000\256\000\000',
	PowerPointMXiMPostString = '\000\256\000\001'
};
typedef enum PowerPointMXiM PowerPointMXiM;

enum PowerPointMANT {
	PowerPointMANTIdle = '\000|\000\001',
	PowerPointMANTGreeting = '\000|\000\002',
	PowerPointMANTGoodbye = '\000|\000\003',
	PowerPointMANTBeginSpeaking = '\000|\000\004',
	PowerPointMANTCharacterSuccessMajor = '\000|\000\006',
	PowerPointMANTGetAttentionMajor = '\000|\000\013',
	PowerPointMANTGetAttentionMinor = '\000|\000\014',
	PowerPointMANTSearching = '\000|\000\015',
	PowerPointMANTPrinting = '\000|\000\022',
	PowerPointMANTGestureRight = '\000|\000\023',
	PowerPointMANTWritingNotingSomething = '\000|\000\026',
	PowerPointMANTWorkingAtSomething = '\000|\000\027',
	PowerPointMANTThinking = '\000|\000\030',
	PowerPointMANTSendingMail = '\000|\000\031',
	PowerPointMANTListensToComputer = '\000|\000\032',
	PowerPointMANTDisappear = '\000|\000\037',
	PowerPointMANTAppear = '\000|\000 ',
	PowerPointMANTGetArtsy = '\000|\000d',
	PowerPointMANTGetTechy = '\000|\000e',
	PowerPointMANTGetWizardy = '\000|\000f',
	PowerPointMANTCheckingSomething = '\000|\000g',
	PowerPointMANTLookDown = '\000|\000h',
	PowerPointMANTLookDownLeft = '\000|\000i',
	PowerPointMANTLookDownRight = '\000|\000j',
	PowerPointMANTLookLeft = '\000|\000k',
	PowerPointMANTLookRight = '\000|\000l',
	PowerPointMANTLookUp = '\000|\000m',
	PowerPointMANTLookUpLeft = '\000|\000n',
	PowerPointMANTLookUpRight = '\000|\000o',
	PowerPointMANTSaving = '\000|\000p',
	PowerPointMANTGestureDown = '\000|\000q',
	PowerPointMANTGestureLeft = '\000|\000r',
	PowerPointMANTGestureUp = '\000|\000s',
	PowerPointMANTEmptyTrash = '\000|\000t'
};
typedef enum PowerPointMANT PowerPointMANT;

enum PowerPointMBSt {
	PowerPointMBStButtonNone = '\000}\000\000',
	PowerPointMBStButtonOk = '\000}\000\001',
	PowerPointMBStButtonCancel = '\000}\000\002',
	PowerPointMBStButtonsOkCancel = '\000}\000\003',
	PowerPointMBStButtonsYesNo = '\000}\000\004',
	PowerPointMBStButtonsYesNoCancel = '\000}\000\005',
	PowerPointMBStButtonsBackClose = '\000}\000\006',
	PowerPointMBStButtonsNextClose = '\000}\000\007',
	PowerPointMBStButtonsBackNextClose = '\000}\000\010',
	PowerPointMBStButtonsRetryCancel = '\000}\000\011',
	PowerPointMBStButtonsAbortRetryIgnore = '\000}\000\012',
	PowerPointMBStButtonsSearchClose = '\000}\000\013',
	PowerPointMBStButtonsBackNextSnooze = '\000}\000\014',
	PowerPointMBStButtonsTipsOptionsClose = '\000}\000\015',
	PowerPointMBStButtonsYesAllNoCancel = '\000}\000\016'
};
typedef enum PowerPointMBSt PowerPointMBSt;

enum PowerPointMIct {
	PowerPointMIctIconNone = '\000~\000\000',
	PowerPointMIctIconApplication = '\000~\000\001',
	PowerPointMIctIconAlert = '\000~\000\002',
	PowerPointMIctIconTip = '\000~\000\003',
	PowerPointMIctIconAlertCritical = '\000~\000e',
	PowerPointMIctIconAlertWarning = '\000~\000g',
	PowerPointMIctIconAlertInfo = '\000~\000h'
};
typedef enum PowerPointMIct PowerPointMIct;

enum PowerPointMWAt {
	PowerPointMWAtInactive = '\000\202\000\000',
	PowerPointMWAtActive = '\000\202\000\001',
	PowerPointMWAtSuspend = '\000\202\000\002',
	PowerPointMWAtResume = '\000\202\000\003'
};
typedef enum PowerPointMWAt PowerPointMWAt;

enum PowerPointMeDP {
	PowerPointMeDPPropertyTypeNumber = '\000\242\000\001',
	PowerPointMeDPPropertyTypeBoolean = '\000\242\000\002',
	PowerPointMeDPPropertyTypeDate = '\000\242\000\003',
	PowerPointMeDPPropertyTypeString = '\000\242\000\004',
	PowerPointMeDPPropertyTypeFloat = '\000\242\000\005'
};
typedef enum PowerPointMeDP PowerPointMeDP;

enum PowerPointMASc {
	PowerPointMAScMsoAutomationSecurityLow = '\000\243\000\001',
	PowerPointMAScMsoAutomationSecurityByUI = '\000\243\000\002',
	PowerPointMAScMsoAutomationSecurityForceDisable = '\000\243\000\003'
};
typedef enum PowerPointMASc PowerPointMASc;

enum PowerPointMSsz {
	PowerPointMSszResolution544x376 = '\000\204\000\000',
	PowerPointMSszResolution640x480 = '\000\204\000\001',
	PowerPointMSszResolution720x512 = '\000\204\000\002',
	PowerPointMSszResolution800x600 = '\000\204\000\003',
	PowerPointMSszResolution1024x768 = '\000\204\000\004',
	PowerPointMSszResolution1152x882 = '\000\204\000\005',
	PowerPointMSszResolution1152x900 = '\000\204\000\006',
	PowerPointMSszResolution1280x1024 = '\000\204\000\007',
	PowerPointMSszResolution1600x1200 = '\000\204\000\010',
	PowerPointMSszResolution1800x1440 = '\000\204\000\011',
	PowerPointMSszResolution1920x1200 = '\000\204\000\012'
};
typedef enum PowerPointMSsz PowerPointMSsz;

enum PowerPointMChS {
	PowerPointMChSArabicCharacterSet = '\000\205\000\001',
	PowerPointMChSCyrillicCharacterSet = '\000\205\000\002',
	PowerPointMChSEnglishCharacterSet = '\000\205\000\003',
	PowerPointMChSGreekCharacterSet = '\000\205\000\004',
	PowerPointMChSHebrewCharacterSet = '\000\205\000\005',
	PowerPointMChSJapaneseCharacterSet = '\000\205\000\006',
	PowerPointMChSKoreanCharacterSet = '\000\205\000\007',
	PowerPointMChSMultilingualUnicodeCharacterSet = '\000\205\000\010',
	PowerPointMChSSimplifiedChineseCharacterSet = '\000\205\000\011',
	PowerPointMChSThaiCharacterSet = '\000\205\000\012',
	PowerPointMChSTraditionalChineseCharacterSet = '\000\205\000\013',
	PowerPointMChSVietnameseCharacterSet = '\000\205\000\014'
};
typedef enum PowerPointMChS PowerPointMChS;

enum PowerPointMtEn {
	PowerPointMtEnEncodingThai = '\000\213\003j',
	PowerPointMtEnEncodingJapaneseShiftJIS = '\000\213\003\244',
	PowerPointMtEnEncodingSimplifiedChinese = '\000\213\003\250',
	PowerPointMtEnEncodingKorean = '\000\213\003\265',
	PowerPointMtEnEncodingBig5TraditionalChinese = '\000\213\003\266',
	PowerPointMtEnEncodingLittleEndian = '\000\213\004\260',
	PowerPointMtEnEncodingBigEndian = '\000\213\004\261',
	PowerPointMtEnEncodingCentralEuropean = '\000\213\004\342',
	PowerPointMtEnEncodingCyrillic = '\000\213\004\343',
	PowerPointMtEnEncodingWestern = '\000\213\004\344',
	PowerPointMtEnEncodingGreek = '\000\213\004\345',
	PowerPointMtEnEncodingTurkish = '\000\213\004\346',
	PowerPointMtEnEncodingHebrew = '\000\213\004\347',
	PowerPointMtEnEncodingArabic = '\000\213\004\350',
	PowerPointMtEnEncodingBaltic = '\000\213\004\351',
	PowerPointMtEnEncodingVietnamese = '\000\213\004\352',
	PowerPointMtEnEncodingISO88591Latin1 = '\000\213o\257',
	PowerPointMtEnEncodingISO88592CentralEurope = '\000\213o\260',
	PowerPointMtEnEncodingISO88593Latin3 = '\000\213o\261',
	PowerPointMtEnEncodingISO88594Baltic = '\000\213o\262',
	PowerPointMtEnEncodingISO88595Cyrillic = '\000\213o\263',
	PowerPointMtEnEncodingISO88596Arabic = '\000\213o\264',
	PowerPointMtEnEncodingISO88597Greek = '\000\213o\265',
	PowerPointMtEnEncodingISO88598Hebrew = '\000\213o\266',
	PowerPointMtEnEncodingISO88599Turkish = '\000\213o\267',
	PowerPointMtEnEncodingISO885915Latin9 = '\000\213o\275',
	PowerPointMtEnEncodingISO2022JapaneseNoHalfWidthKatakana = '\000\213\304,',
	PowerPointMtEnEncodingISO2022JapaneseJISX02021984 = '\000\213\304-',
	PowerPointMtEnEncodingISO2022JapaneseJISX02011989 = '\000\213\304.',
	PowerPointMtEnEncodingISO2022KR = '\000\213\3041',
	PowerPointMtEnEncodingISO2022CNTraditionalChinese = '\000\213\3043',
	PowerPointMtEnEncodingISO2022CNSimplifiedChinese = '\000\213\3045',
	PowerPointMtEnEncodingMacRoman = '\000\213\'\020',
	PowerPointMtEnEncodingMacJapanese = '\000\213\'\021',
	PowerPointMtEnEncodingMacTraditionalChinese = '\000\213\'\022',
	PowerPointMtEnEncodingMacKorean = '\000\213\'\023',
	PowerPointMtEnEncodingMacArabic = '\000\213\'\024',
	PowerPointMtEnEncodingMacHebrew = '\000\213\'\025',
	PowerPointMtEnEncodingMacGreek1 = '\000\213\'\026',
	PowerPointMtEnEncodingMacCyrillic = '\000\213\'\027',
	PowerPointMtEnEncodingMacSimplifiedChineseGB2312 = '\000\213\'\030',
	PowerPointMtEnEncodingMacRomania = '\000\213\'\032',
	PowerPointMtEnEncodingMacUkraine = '\000\213\'!',
	PowerPointMtEnEncodingMacLatin2 = '\000\213\'-',
	PowerPointMtEnEncodingMacIcelandic = '\000\213\'_',
	PowerPointMtEnEncodingMacTurkish = '\000\213\'a',
	PowerPointMtEnEncodingMacCroatia = '\000\213\'b',
	PowerPointMtEnEncodingEBCDICUSCanada = '\000\213\000%',
	PowerPointMtEnEncodingEBCDICInternational = '\000\213\001\364',
	PowerPointMtEnEncodingEBCDICMultilingualROECELatin2 = '\000\213\003f',
	PowerPointMtEnEncodingEBCDICGreekModern = '\000\213\003k',
	PowerPointMtEnEncodingEBCDICTurkishLatin5 = '\000\213\004\002',
	PowerPointMtEnEncodingEBCDICGermany = '\000\213O1',
	PowerPointMtEnEncodingEBCDICDenmarkNorway = '\000\213O5',
	PowerPointMtEnEncodingEBCDICFinlandSweden = '\000\213O6',
	PowerPointMtEnEncodingEBCDICItaly = '\000\213O8',
	PowerPointMtEnEncodingEBCDICLatinAmericaSpain = '\000\213O<',
	PowerPointMtEnEncodingEBCDICUnitedKingdom = '\000\213O=',
	PowerPointMtEnEncodingEBCDICJapaneseKatakanaExtended = '\000\213OB',
	PowerPointMtEnEncodingEBCDICFrance = '\000\213OI',
	PowerPointMtEnEncodingEBCDICArabic = '\000\213O\304',
	PowerPointMtEnEncodingEBCDICGreek = '\000\213O\307',
	PowerPointMtEnEncodingEBCDICHebrew = '\000\213O\310',
	PowerPointMtEnEncodingEBCDICKoreanExtended = '\000\213Qa',
	PowerPointMtEnEncodingEBCDICThai = '\000\213Qf',
	PowerPointMtEnEncodingEBCDICIcelandic = '\000\213Q\207',
	PowerPointMtEnEncodingEBCDICTurkish = '\000\213Q\251',
	PowerPointMtEnEncodingEBCDICRussian = '\000\213Q\220',
	PowerPointMtEnEncodingEBCDICSerbianBulgarian = '\000\213R!',
	PowerPointMtEnEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = '\000\213\306\362',
	PowerPointMtEnEncodingEBCDICUSCanadaAndJapanese = '\000\213\306\363',
	PowerPointMtEnEncodingEBCDICExtendedAndKorean = '\000\213\306\365',
	PowerPointMtEnEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = '\000\213\306\367',
	PowerPointMtEnEncodingEBCDICUSCanadaAndTraditionalChinese = '\000\213\306\371',
	PowerPointMtEnEncodingEBCDICJapaneseLatinExtendedAndJapanese = '\000\213\306\373',
	PowerPointMtEnEncodingOEMUnitedStates = '\000\213\001\265',
	PowerPointMtEnEncodingOEMGreek = '\000\213\002\341',
	PowerPointMtEnEncodingOEMBaltic = '\000\213\003\007',
	PowerPointMtEnEncodingOEMMultilingualLatinI = '\000\213\003R',
	PowerPointMtEnEncodingOEMMultilingualLatinII = '\000\213\003T',
	PowerPointMtEnEncodingOEMCyrillic = '\000\213\003W',
	PowerPointMtEnEncodingOEMTurkish = '\000\213\003Y',
	PowerPointMtEnEncodingOEMPortuguese = '\000\213\003\\',
	PowerPointMtEnEncodingOEMIcelandic = '\000\213\003]',
	PowerPointMtEnEncodingOEMHebrew = '\000\213\003^',
	PowerPointMtEnEncodingOEMCanadianFrench = '\000\213\003_',
	PowerPointMtEnEncodingOEMArabic = '\000\213\003`',
	PowerPointMtEnEncodingOEMNordic = '\000\213\003a',
	PowerPointMtEnEncodingOEMCyrillicII = '\000\213\003b',
	PowerPointMtEnEncodingOEMModernGreek = '\000\213\003e',
	PowerPointMtEnEncodingEUCJapanese = '\000\213\312\334',
	PowerPointMtEnEncodingEUCChineseSimplifiedChinese = '\000\213\312\340',
	PowerPointMtEnEncodingEUCKorean = '\000\213\312\355',
	PowerPointMtEnEncodingEUCTaiwaneseTraditionalChinese = '\000\213\312\356',
	PowerPointMtEnEncodingDevanagari = '\000\213\336\252',
	PowerPointMtEnEncodingBengali = '\000\213\336\253',
	PowerPointMtEnEncodingTamil = '\000\213\336\254',
	PowerPointMtEnEncodingTelugu = '\000\213\336\255',
	PowerPointMtEnEncodingAssamese = '\000\213\336\256',
	PowerPointMtEnEncodingOriya = '\000\213\336\257',
	PowerPointMtEnEncodingKannada = '\000\213\336\260',
	PowerPointMtEnEncodingMalayalam = '\000\213\336\261',
	PowerPointMtEnEncodingGujarati = '\000\213\336\262',
	PowerPointMtEnEncodingPunjabi = '\000\213\336\263',
	PowerPointMtEnEncodingArabicASMO = '\000\213\002\304',
	PowerPointMtEnEncodingArabicTransparentASMO = '\000\213\002\320',
	PowerPointMtEnEncodingKoreanJohab = '\000\213\005Q',
	PowerPointMtEnEncodingTaiwanCNS = '\000\213N ',
	PowerPointMtEnEncodingTaiwanTCA = '\000\213N!',
	PowerPointMtEnEncodingTaiwanEten = '\000\213N\"',
	PowerPointMtEnEncodingTaiwanIBM5550 = '\000\213N#',
	PowerPointMtEnEncodingTaiwanTeletext = '\000\213N$',
	PowerPointMtEnEncodingTaiwanWang = '\000\213N%',
	PowerPointMtEnEncodingIA5IRV = '\000\213N\211',
	PowerPointMtEnEncodingIA5German = '\000\213N\212',
	PowerPointMtEnEncodingIA5Swedish = '\000\213N\213',
	PowerPointMtEnEncodingIA5Norwegian = '\000\213N\214',
	PowerPointMtEnEncodingUSASCII = '\000\213N\237',
	PowerPointMtEnEncodingT61 = '\000\213O%',
	PowerPointMtEnEncodingISO6937NonspacingAccent = '\000\213O-',
	PowerPointMtEnEncodingKOI8R = '\000\213Q\202',
	PowerPointMtEnEncodingExtAlphaLowercase = '\000\213R#',
	PowerPointMtEnEncodingKOI8U = '\000\213Uj',
	PowerPointMtEnEncodingEuropa3 = '\000\213qI',
	PowerPointMtEnEncodingHZGBSimplifiedChinese = '\000\213\316\310',
	PowerPointMtEnEncodingUTF7 = '\000\213\375\350',
	PowerPointMtEnEncodingUTF8 = '\000\213\375\351'
};
typedef enum PowerPointMtEn PowerPointMtEn;

enum PowerPoint4000 {
	PowerPoint4000CommandBar = 'msCB',
	PowerPoint4000CommandBarControl = 'mCBC'
};
typedef enum PowerPoint4000 PowerPoint4000;

enum PowerPointMHyT {
	PowerPointMHyTHyperlinkRange = '\000\310\000\000',
	PowerPointMHyTHyperlinkShape = '\000\310\000\001',
	PowerPointMHyTHyperlinkInlineShape = '\000\310\000\002'
};
typedef enum PowerPointMHyT PowerPointMHyT;

enum PowerPointPWnS {
	PowerPointPWnSWindowNormal = '\000\311\000\001',
	PowerPointPWnSWindowMinimized = '\000\311\000\002'
};
typedef enum PowerPointPWnS PowerPointPWnS;

enum PowerPointPArS {
	PowerPointPArSArrangeTiled = '\000\321\000\001',
	PowerPointPArSArrangeCascade = '\000\321\000\002'
};
typedef enum PowerPointPArS PowerPointPArS;

enum PowerPointPVTy {
	PowerPointPVTySlideView = '\000\312\000\001',
	PowerPointPVTyMasterView = '\000\312\000\002',
	PowerPointPVTyPageView = '\000\312\000\003',
	PowerPointPVTyHandoutMasterView = '\000\312\000\004',
	PowerPointPVTyNotesMasterView = '\000\312\000\005',
	PowerPointPVTyOutlineView = '\000\312\000\006',
	PowerPointPVTySlideSorterView = '\000\312\000\007',
	PowerPointPVTyTitleMasterView = '\000\312\000\010',
	PowerPointPVTyNormalView = '\000\312\000\011',
	PowerPointPVTyThumbnailView = '\000\312\000\012',
	PowerPointPVTyThumbnailMasterView = '\000\312\000\013'
};
typedef enum PowerPointPVTy PowerPointPVTy;

enum PowerPointPCSi {
	PowerPointPCSiSchemeColorUnset = '\000\362\377\376',
	PowerPointPCSiNotASchemeColor = '\000\363\000\000',
	PowerPointPCSiBackgroundScheme = '\000\363\000\001',
	PowerPointPCSiForegroundScheme = '\000\363\000\002',
	PowerPointPCSiShadowScheme = '\000\363\000\003',
	PowerPointPCSiTitleScheme = '\000\363\000\004',
	PowerPointPCSiFillScheme = '\000\363\000\005',
	PowerPointPCSiAccent1Scheme = '\000\363\000\006',
	PowerPointPCSiAccent2Scheme = '\000\363\000\007',
	PowerPointPCSiAccent3Scheme = '\000\363\000\010'
};
typedef enum PowerPointPCSi PowerPointPCSi;

enum PowerPointSSzT {
	PowerPointSSzTSlideSizeOnScreen = '\000\313\000\001',
	PowerPointSSzTSlideSizeLetterPaper = '\000\313\000\002',
	PowerPointSSzTSlideSizeA4Paper = '\000\313\000\003',
	PowerPointSSzTSlideSize35MM = '\000\313\000\004',
	PowerPointSSzTSlideSizeOverhead = '\000\313\000\005',
	PowerPointSSzTSlideSizeBanner = '\000\313\000\006',
	PowerPointSSzTSlideSizeCustom = '\000\313\000\007'
};
typedef enum PowerPointSSzT PowerPointSSzT;

enum PowerPointPSAT {
	PowerPointPSATSaveAsPresentation = '\000\314\000\001',
	PowerPointPSATSaveAsTemplate = '\000\314\000\005',
	PowerPointPSATSaveAsRTF = '\000\314\000\006',
	PowerPointPSATSaveAsShow = '\000\314\000\007',
	PowerPointPSATSaveAsAddIn = '\000\314\000\010',
	PowerPointPSATSaveAsDefault = '\000\314\000\012',
	PowerPointPSATSaveAsHTML = '\000\314\000\013',
	PowerPointPSATSaveAsMovie = '\000\314\000\014',
	PowerPointPSATSaveAsPackage = '\000\314\000\015',
	PowerPointPSATSaveAsPDF = '\000\314\000\016',
	PowerPointPSATSaveAsOpenXMLPresentation = '\000\314\000\017',
	PowerPointPSATSaveAsOpenXMLPresentationMacroEnabled = '\000\314\000\020',
	PowerPointPSATSaveAsOpenXMLShow = '\000\314\000\021',
	PowerPointPSATSaveAsOpenXMLShowMacroEnabled = '\000\314\000\022',
	PowerPointPSATSaveAsOpenXMLTemplate = '\000\314\000\023',
	PowerPointPSATSaveAsOpenXMLTemplateMacroEnabled = '\000\314\000\024',
	PowerPointPSATSaveAsOpenXMLTheme = '\000\314\000\025',
	PowerPointPSATSaveAsGIF = '\000\314\000\026',
	PowerPointPSATSaveAsJPG = '\000\314\000\027',
	PowerPointPSATSaveAsPNG = '\000\314\000\030',
	PowerPointPSATSaveAsBMP = '\000\314\000\031',
	PowerPointPSATSaveAsTIF = '\000\314\000\032'
};
typedef enum PowerPointPSAT PowerPointPSAT;

enum PowerPointPTst {
	PowerPointPTstTextStyleDefault = '\001*\000\001',
	PowerPointPTstTextStyleTitle = '\001*\000\002',
	PowerPointPTstTextStyleBody = '\001*\000\003'
};
typedef enum PowerPointPTst PowerPointPTst;

enum PowerPointPSLo {
	PowerPointPSLoSlideLayoutUnset = '\000\314\377\376',
	PowerPointPSLoSlideLayoutTitleSlide = '\000\315\000\001',
	PowerPointPSLoSlideLayoutTextSlide = '\000\315\000\002',
	PowerPointPSLoSlideLayoutTwoColumnText = '\000\315\000\003',
	PowerPointPSLoSlideLayoutTable = '\000\315\000\004',
	PowerPointPSLoSlideLayoutTextAndChart = '\000\315\000\005',
	PowerPointPSLoSlideLayoutChartAndText = '\000\315\000\006',
	PowerPointPSLoSlideLayoutOrgchart = '\000\315\000\007',
	PowerPointPSLoSlideLayoutChart = '\000\315\000\010',
	PowerPointPSLoSlideLayoutTextAndClipart = '\000\315\000\011',
	PowerPointPSLoSlideLayoutClipartAndText = '\000\315\000\012',
	PowerPointPSLoSlideLayoutTitleOnly = '\000\315\000\013',
	PowerPointPSLoSlideLayoutBlank = '\000\315\000\014',
	PowerPointPSLoSlideLayoutTextAndObject = '\000\315\000\015',
	PowerPointPSLoSlideLayoutObjectAndText = '\000\315\000\016',
	PowerPointPSLoSlideLayoutLargeObject = '\000\315\000\017',
	PowerPointPSLoSlideLayoutObject = '\000\315\000\020',
	PowerPointPSLoSlideLayoutMediaClip = '\000\315\000\021',
	PowerPointPSLoSlideLayoutMediaClipAndText = '\000\315\000\022',
	PowerPointPSLoSlideLayoutObjectOverText = '\000\315\000\023',
	PowerPointPSLoSlideLayoutTextOverObject = '\000\315\000\024',
	PowerPointPSLoSlideLayoutTextAndTwoObjects = '\000\315\000\025',
	PowerPointPSLoSlideLayoutTwoObjectsAndText = '\000\315\000\026',
	PowerPointPSLoSlideLayoutTwoObjectsOverText = '\000\315\000\027',
	PowerPointPSLoSlideLayoutFourObjects = '\000\315\000\030',
	PowerPointPSLoSlideLayoutVerticalText = '\000\315\000\031',
	PowerPointPSLoSlideLayoutClipartAndVerticalText = '\000\315\000\032',
	PowerPointPSLoSlideLayoutVerticalTitleAndText = '\000\315\000\033',
	PowerPointPSLoSlideLayoutVerticalTitleAndTextOverChart = '\000\315\000\034',
	PowerPointPSLoSlideLayoutTwoObjects = '\000\315\000\035',
	PowerPointPSLoSlideLayoutObjectAndTwoObjects = '\000\315\000\036',
	PowerPointPSLoSlideLayoutTwoObjectsAndObject = '\000\315\000\037',
	PowerPointPSLoSlideLayoutCustom = '\000\315\000 ',
	PowerPointPSLoSlideLayoutSectionHeader = '\000\315\000!',
	PowerPointPSLoSlideLayoutComparison = '\000\315\000\"',
	PowerPointPSLoSlideLayoutContentWithCaption = '\000\315\000#',
	PowerPointPSLoSlideLayoutPictureWithCaption = '\000\315\000$'
};
typedef enum PowerPointPSLo PowerPointPSLo;

enum PowerPointPEeF {
	PowerPointPEeFEntryEffectAppear = '\000\366\017\004',
	PowerPointPEeFEntryEffectBlindsHorizontal = '\000\366\003\001',
	PowerPointPEeFEntryEffectBlindsVertical = '\000\366\003\002',
	PowerPointPEeFEntryEffectBoxDown = '\000\366\017U',
	PowerPointPEeFEntryEffectBoxIn = '\000\366\014\002',
	PowerPointPEeFEntryEffectBoxLeft = '\000\366\017R',
	PowerPointPEeFEntryEffectBoxOut = '\000\366\014\001',
	PowerPointPEeFEntryEffectBoxRight = '\000\366\017T',
	PowerPointPEeFEntryEffectBoxUp = '\000\366\017S',
	PowerPointPEeFEntryEffectCheckerboardAcross = '\000\366\004\001',
	PowerPointPEeFEntryEffectCheckerboardDown = '\000\366\004\002',
	PowerPointPEeFEntryEffectCircle = '\000\366\017\005',
	PowerPointPEeFEntryEffectCollapseAcross = '\000\366\015\027',
	PowerPointPEeFEntryEffectCollapseBottom = '\000\366\015\033',
	PowerPointPEeFEntryEffectCollapseLeft = '\000\366\015\030',
	PowerPointPEeFEntryEffectCollapseRight = '\000\366\015\032',
	PowerPointPEeFEntryEffectCollapseUp = '\000\366\015\031',
	PowerPointPEeFEntryEffectCombHorizontal = '\000\366\017\007',
	PowerPointPEeFEntryEffectCombVertical = '\000\366\017\010',
	PowerPointPEeFEntryEffectConveyorLeft = '\000\366\017*',
	PowerPointPEeFEntryEffectConveyorRight = '\000\366\017+',
	PowerPointPEeFEntryEffectCoverDown = '\000\366\005\004',
	PowerPointPEeFEntryEffectCoverLeftDown = '\000\366\005\007',
	PowerPointPEeFEntryEffectCoverLeftUp = '\000\366\005\005',
	PowerPointPEeFEntryEffectCoverLeft = '\000\366\005\001',
	PowerPointPEeFEntryEffectCoverRightDown = '\000\366\005\010',
	PowerPointPEeFEntryEffectCoverRightUp = '\000\366\005\006',
	PowerPointPEeFEntryEffectCoverRight = '\000\366\005\003',
	PowerPointPEeFEntryEffectCoverUp = '\000\366\005\002',
	PowerPointPEeFEntryEffectCrawlFromDown = '\000\366\015\020',
	PowerPointPEeFEntryEffectCrawlFromLeft = '\000\366\015\015',
	PowerPointPEeFEntryEffectCrawlFromRight = '\000\366\015\017',
	PowerPointPEeFEntryEffectCrawlFromUp = '\000\366\015\016',
	PowerPointPEeFEntryEffectCubeDown = '\000\366\017M',
	PowerPointPEeFEntryEffectCubeLeft = '\000\366\017J',
	PowerPointPEeFEntryEffectCubeRight = '\000\366\017K',
	PowerPointPEeFEntryEffectCubeUp = '\000\366\017L',
	PowerPointPEeFEntryEffectCutBlack = '\000\366\001\002',
	PowerPointPEeFEntryEffectCut = '\000\366\001\001',
	PowerPointPEeFEntryEffectDiamond = '\000\366\017\006',
	PowerPointPEeFEntryEffectDissolve = '\000\366\006\001',
	PowerPointPEeFEntryEffectDoorsHorizontal = '\000\366\017-',
	PowerPointPEeFEntryEffectDoorsVertical = '\000\366\017,',
	PowerPointPEeFEntryEffectFadeBlack = '\000\366\007\001',
	PowerPointPEeFEntryEffectFadeFlyFromBottomLeft = '\000\366\015$',
	PowerPointPEeFEntryEffectFadeFlyFromBottomRight = '\000\366\015%',
	PowerPointPEeFEntryEffectFadeFlyFromBottom = '\000\366\015!',
	PowerPointPEeFEntryEffectFadeFlyFromLeft = '\000\366\015\036',
	PowerPointPEeFEntryEffectFadeFlyFromRight = '\000\366\015 ',
	PowerPointPEeFEntryEffectFadeFlyFromTopLeft = '\000\366\015\"',
	PowerPointPEeFEntryEffectFadeFlyFromTopRight = '\000\366\015#',
	PowerPointPEeFEntryEffectFadeFlyFromTop = '\000\366\015\037',
	PowerPointPEeFEntryEffectFadeSmoothly = '\000\366\017\011',
	PowerPointPEeFEntryEffectFade = '\000\366\017\011',
	PowerPointPEeFEntryEffectFerrisWheelLeft = '\000\366\017;',
	PowerPointPEeFEntryEffectFerrisWheelRight = '\000\366\017<',
	PowerPointPEeFEntryEffectFlashOnceFast = '\000\366\017\001',
	PowerPointPEeFEntryEffectFlashOnceMedium = '\000\366\017\002',
	PowerPointPEeFEntryEffectFlashOnceSlow = '\000\366\017\003',
	PowerPointPEeFEntryEffectFlashbulb = '\000\366\017E',
	PowerPointPEeFEntryEffectFlipDown = '\000\366\017D',
	PowerPointPEeFEntryEffectFlipLeft = '\000\366\017A',
	PowerPointPEeFEntryEffectFlipRight = '\000\366\017B',
	PowerPointPEeFEntryEffectFlipUp = '\000\366\017C',
	PowerPointPEeFEntryEffectFlyFromBottomLeft = '\000\366\015\007',
	PowerPointPEeFEntryEffectFlyFromBottomRight = '\000\366\015\010',
	PowerPointPEeFEntryEffectFlyFromBottom = '\000\366\015\004',
	PowerPointPEeFEntryEffectFlyFromLeft = '\000\366\015\001',
	PowerPointPEeFEntryEffectFlyFromRight = '\000\366\015\003',
	PowerPointPEeFEntryEffectFlyFromTopLeft = '\000\366\015\005',
	PowerPointPEeFEntryEffectFlyFromTopRight = '\000\366\015\006',
	PowerPointPEeFEntryEffectFlyFromTop = '\000\366\015\002',
	PowerPointPEeFEntryEffectFlyThroughInBounce = '\000\366\0174',
	PowerPointPEeFEntryEffectFlyThroughIn = '\000\366\0172',
	PowerPointPEeFEntryEffectFlyThroughOutBounce = '\000\366\0175',
	PowerPointPEeFEntryEffectFlyThroughOut = '\000\366\0173',
	PowerPointPEeFEntryEffectGalleryLeft = '\000\366\017(',
	PowerPointPEeFEntryEffectGalleryRight = '\000\366\017)',
	PowerPointPEeFEntryEffectGlitterDiamondDown = '\000\366\017#',
	PowerPointPEeFEntryEffectGlitterDiamondLeft = '\000\366\017 ',
	PowerPointPEeFEntryEffectGlitterDiamondRight = '\000\366\017\"',
	PowerPointPEeFEntryEffectGlitterDiamondUp = '\000\366\017!',
	PowerPointPEeFEntryEffectGlitterHexagonDown = '\000\366\017\'',
	PowerPointPEeFEntryEffectGlitterHexagonLeft = '\000\366\017$',
	PowerPointPEeFEntryEffectGlitterHexagonRight = '\000\366\017&',
	PowerPointPEeFEntryEffectGlitterHexagonUp = '\000\366\017%',
	PowerPointPEeFEntryEffectHoneycomb = '\000\366\017:',
	PowerPointPEeFEntryEffectNewsFlash = '\000\366\017\012',
	PowerPointPEeFEntryEffectNone = '\000\366\000\000',
	PowerPointPEeFEntryEffectOrbitDown = '\000\366\017Y',
	PowerPointPEeFEntryEffectOrbitLeft = '\000\366\017V',
	PowerPointPEeFEntryEffectOrbitRight = '\000\366\017X',
	PowerPointPEeFEntryEffectOrbitUp = '\000\366\017W',
	PowerPointPEeFEntryEffectPanDown = '\000\366\017]',
	PowerPointPEeFEntryEffectPanLeft = '\000\366\017Z',
	PowerPointPEeFEntryEffectPanRight = '\000\366\017\\',
	PowerPointPEeFEntryEffectPanUp = '\000\366\017[',
	PowerPointPEeFEntryEffectPeekFromDown = '\000\366\015\012',
	PowerPointPEeFEntryEffectPeekFromLeft = '\000\366\015\011',
	PowerPointPEeFEntryEffectPeekFromRight = '\000\366\015\013',
	PowerPointPEeFEntryEffectPeekFromUp = '\000\366\015\014',
	PowerPointPEeFEntryEffectPlus = '\000\366\017\013',
	PowerPointPEeFEntryEffectPushDown = '\000\366\017\014',
	PowerPointPEeFEntryEffectPushLeft = '\000\366\017\015',
	PowerPointPEeFEntryEffectPushRight = '\000\366\017\016',
	PowerPointPEeFEntryEffectPushUp = '\000\366\017\017',
	PowerPointPEeFEntryEffectRandomBarsHorizontal = '\000\366\011\001',
	PowerPointPEeFEntryEffectRandomBarsVertical = '\000\366\011\002',
	PowerPointPEeFEntryEffectRandom = '\000\366\002\001',
	PowerPointPEeFEntryEffectRevealBlackLeft = '\000\366\0178',
	PowerPointPEeFEntryEffectRevealBlackRight = '\000\366\0179',
	PowerPointPEeFEntryEffectRevealSmoothLeft = '\000\366\0176',
	PowerPointPEeFEntryEffectRevealSmoothRight = '\000\366\0177',
	PowerPointPEeFEntryEffectRippleCenter = '\000\366\017\033',
	PowerPointPEeFEntryEffectRippleLeftDown = '\000\366\017\036',
	PowerPointPEeFEntryEffectRippleLeftUp = '\000\366\017\035',
	PowerPointPEeFEntryEffectRippleRightDown = '\000\366\017\037',
	PowerPointPEeFEntryEffectRippleRightUp = '\000\366\017\034',
	PowerPointPEeFEntryEffectRotateDown = '\000\366\017Q',
	PowerPointPEeFEntryEffectRotateLeft = '\000\366\017N',
	PowerPointPEeFEntryEffectRotateRight = '\000\366\017P',
	PowerPointPEeFEntryEffectRotateUp = '\000\366\017O',
	PowerPointPEeFEntryEffectShredRectangleIn = '\000\366\017H',
	PowerPointPEeFEntryEffectShredRectangleOut = '\000\366\017I',
	PowerPointPEeFEntryEffectShredStripsIn = '\000\366\017F',
	PowerPointPEeFEntryEffectShredStripsOut = '\000\366\017G',
	PowerPointPEeFEntryEffectSpinner = '\000\366\017\012',
	PowerPointPEeFEntryEffectSpiral = '\000\366\015\035',
	PowerPointPEeFEntryEffectSplitHorizontalIn = '\000\366\016\002',
	PowerPointPEeFEntryEffectSplitHorizontalOut = '\000\366\016\001',
	PowerPointPEeFEntryEffectSplitVerticalIn = '\000\366\016\004',
	PowerPointPEeFEntryEffectSplitVerticalOut = '\000\366\016\003',
	PowerPointPEeFEntryEffectStretchAcross = '\000\366\015\027',
	PowerPointPEeFEntryEffectStretchDown = '\000\366\015\033',
	PowerPointPEeFEntryEffectStretchLeft = '\000\366\015\030',
	PowerPointPEeFEntryEffectStretchRight = '\000\366\015\032',
	PowerPointPEeFEntryEffectStretchUp = '\000\366\015\031',
	PowerPointPEeFEntryEffectStripsDownLeft = '\000\366\012\003',
	PowerPointPEeFEntryEffectStripsDownRight = '\000\366\012\004',
	PowerPointPEeFEntryEffectStripsLeftDown = '\000\366\012\007',
	PowerPointPEeFEntryEffectStripsLeftUp = '\000\366\012\005',
	PowerPointPEeFEntryEffectStripsRightDown = '\000\366\012\010',
	PowerPointPEeFEntryEffectStripsRightUp = '\000\366\012\006',
	PowerPointPEeFEntryEffectStripsUpLeft = '\000\366\012\001',
	PowerPointPEeFEntryEffectStripsUpRight = '\000\366\012\002',
	PowerPointPEeFEntryEffectSwitchDown = '\000\366\017@',
	PowerPointPEeFEntryEffectSwitchLeft = '\000\366\017=',
	PowerPointPEeFEntryEffectSwitchRight = '\000\366\017\?',
	PowerPointPEeFEntryEffectSwitchUp = '\000\366\017>',
	PowerPointPEeFEntryEffectSwivel = '\000\366\015\034',
	PowerPointPEeFEntryEffectUncoverDown = '\000\366\010\004',
	PowerPointPEeFEntryEffectUncoverLeftDown = '\000\366\010\007',
	PowerPointPEeFEntryEffectUncoverLeftUp = '\000\366\010\005',
	PowerPointPEeFEntryEffectUncoverLeft = '\000\366\010\001',
	PowerPointPEeFEntryEffectUncoverRightDown = '\000\366\010\010',
	PowerPointPEeFEntryEffectUncoverRightUp = '\000\366\010\006',
	PowerPointPEeFEntryEffectUncoverRight = '\000\366\010\003',
	PowerPointPEeFEntryEffectUncoverUp = '\000\366\010\002',
	PowerPointPEeFEntryEffectUnset = '\000\365\377\376',
	PowerPointPEeFEntryEffectVortexDown = '\000\366\017\032',
	PowerPointPEeFEntryEffectVortexLeft = '\000\366\017\027',
	PowerPointPEeFEntryEffectVortexRight = '\000\366\017\031',
	PowerPointPEeFEntryEffectVortexUp = '\000\366\017\030',
	PowerPointPEeFEntryEffectWarpIn = '\000\366\0170',
	PowerPointPEeFEntryEffectWarpOut = '\000\366\0171',
	PowerPointPEeFEntryEffectWedge = '\000\366\017\020',
	PowerPointPEeFEntryEffectWheel1Spoke = '\000\366\017\021',
	PowerPointPEeFEntryEffectWheel2Spokes = '\000\366\017\022',
	PowerPointPEeFEntryEffectWheel3Spokes = '\000\366\017\023',
	PowerPointPEeFEntryEffectWheel4Spokes = '\000\366\017\024',
	PowerPointPEeFEntryEffectWheel8Spokes = '\000\366\017\025',
	PowerPointPEeFEntryEffectWheelReverse1Spoke = '\000\366\017\026',
	PowerPointPEeFEntryEffectWindowHorizontal = '\000\366\017/',
	PowerPointPEeFEntryEffectWindowVertical = '\000\366\017.',
	PowerPointPEeFEntryEffectWipeDown = '\000\366\013\004',
	PowerPointPEeFEntryEffectWipeLeft = '\000\366\013\001',
	PowerPointPEeFEntryEffectWipeRight = '\000\366\013\003',
	PowerPointPEeFEntryEffectWipeUp = '\000\366\013\002',
	PowerPointPEeFEntryEffectZoomBottom = '\000\366\015\026',
	PowerPointPEeFEntryEffectZoomCenter = '\000\366\015\025',
	PowerPointPEeFEntryEffectZoomInSlightly = '\000\366\015\022',
	PowerPointPEeFEntryEffectZoomIn = '\000\366\015\021',
	PowerPointPEeFEntryEffectZoomOutSlightly = '\000\366\015\024',
	PowerPointPEeFEntryEffectZoomOut = '\000\366\015\023'
};
typedef enum PowerPointPEeF PowerPointPEeF;

enum PowerPointPTlE {
	PowerPointPTlEAnimationLevelUnset = '\000\337\377\376',
	PowerPointPTlEAnimateLevelNone = '\000\340\000\000',
	PowerPointPTlEAnimateLevelFirstLevel = '\000\340\000\001',
	PowerPointPTlEAnimateLevelSecondLevel = '\000\340\000\002',
	PowerPointPTlEAnimateLevelThirdLevel = '\000\340\000\003',
	PowerPointPTlEAnimateLevelFourthLevel = '\000\340\000\004',
	PowerPointPTlEAnimateLevelFifthLevel = '\000\340\000\005',
	PowerPointPTlEAnimateLevelAllLevels = '\000\340\000\020'
};
typedef enum PowerPointPTlE PowerPointPTlE;

enum PowerPointPTuE {
	PowerPointPTuEAnimationUnitUnset = '\000\340\377\376',
	PowerPointPTuETextUnitEffectByParagraph = '\000\341\000\000',
	PowerPointPTuETextUnitEffectByWord = '\000\341\000\001',
	PowerPointPTuETextUnitEffectByCharacter = '\000\341\000\002'
};
typedef enum PowerPointPTuE PowerPointPTuE;

enum PowerPointPCuE {
	PowerPointPCuEAnimationChartUnset = '\000\341\377\376',
	PowerPointPCuEChartUnitEffectBySeries = '\000\342\000\001',
	PowerPointPCuEChartUnitEffectByCategory = '\000\342\000\002',
	PowerPointPCuEChartUnitEffectBySeriesElement = '\000\342\000\003'
};
typedef enum PowerPointPCuE PowerPointPCuE;

enum PowerPointAsAe {
	PowerPointAsAeAfterEffectUnset = '\000\363\377\376',
	PowerPointAsAeAfterEffectNone = '\000\364\000\000',
	PowerPointAsAeAfterEffectHide = '\000\364\000\001',
	PowerPointAsAeAfterEffectDim = '\000\364\000\002'
};
typedef enum PowerPointAsAe PowerPointAsAe;

enum PowerPointAdMd {
	PowerPointAdMdAdvanceModeUnset = '\000\361\377\376',
	PowerPointAdMdAdvanceModeOnClick = '\000\362\000\001'
};
typedef enum PowerPointAdMd PowerPointAdMd;

enum PowerPointPSnX {
	PowerPointPSnXSoundEffectUnset = '\000\331\377\376',
	PowerPointPSnXSoundEffectNone = '\000\332\000\000',
	PowerPointPSnXSoundEffectStopPrevious = '\000\332\000\001',
	PowerPointPSnXSoundEffectFile = '\000\332\000\002'
};
typedef enum PowerPointPSnX PowerPointPSnX;

enum PowerPointPUdO {
	PowerPointPUdOUpdateOptionUnset = '\000\336\377\376',
	PowerPointPUdOUpdateOptionManual = '\000\337\000\001'
};
typedef enum PowerPointPUdO PowerPointPUdO;

enum PowerPointPDgM {
	PowerPointPDgMDialogModeUnset = '\000\357\377\376',
	PowerPointPDgMDialogModeModless = '\000\360\000\000',
	PowerPointPDgMDialogModeModal = '\000\360\000\001'
};
typedef enum PowerPointPDgM PowerPointPDgM;

enum PowerPointPDgS {
	PowerPointPDgSDialogStyleUnset = '\000\360\377\376',
	PowerPointPDgSDialogStyleStandard = '\000\361\000\001'
};
typedef enum PowerPointPDgS PowerPointPDgS;

enum PowerPointPSsP {
	PowerPointPSsPSlideShowPointerNone = '\000\322\000\000',
	PowerPointPSsPSlideShowPointerArrow = '\000\322\000\001',
	PowerPointPSsPSlideShowPointerPen = '\000\322\000\002',
	PowerPointPSsPSlideShowPointerAlwaysHidden = '\000\322\000\003'
};
typedef enum PowerPointPSsP PowerPointPSsP;

enum PowerPointPShS {
	PowerPointPShSSlideShowStateRunning = '\000\323\000\001',
	PowerPointPShSSlideShowStatePaused = '\000\323\000\002',
	PowerPointPShSSlideShowStateBlackScreen = '\000\323\000\003',
	PowerPointPShSSlideShowStateWhiteScreen = '\000\323\000\004',
	PowerPointPShSSlideShowStateDone = '\000\323\000\005'
};
typedef enum PowerPointPShS PowerPointPShS;

enum PowerPointPSta {
	PowerPointPStaPlayerStatePlaying = '\000\365\000\000',
	PowerPointPStaPlayerStatePaused = '\000\365\000\001',
	PowerPointPStaPlayerStateStopped = '\000\365\000\002',
	PowerPointPStaPlayerStateNotReady = '\000\365\000\003'
};
typedef enum PowerPointPSta PowerPointPSta;

enum PowerPointPSaM {
	PowerPointPSaMSlideShowAdvanceManualAdvance = '\000\324\000\001',
	PowerPointPSaMSlideShowAdvanceUseSlideTimings = '\000\324\000\002'
};
typedef enum PowerPointPSaM PowerPointPSaM;

enum PowerPointPtOt {
	PowerPointPtOtPrintSlides = '\000\330\000\001',
	PowerPointPtOtPrintTwoSlideHandouts = '\000\330\000\002',
	PowerPointPtOtPrintThreeSlideHandouts = '\000\330\000\003',
	PowerPointPtOtPrintSixSlideHandouts = '\000\330\000\004',
	PowerPointPtOtPrintNotesPages = '\000\330\000\005',
	PowerPointPtOtPrintOutline = '\000\330\000\006',
	PowerPointPtOtPrintFourSlideHandouts = '\000\330\000\007',
	PowerPointPtOtPrintNineSlideHandouts = '\000\330\000\010'
};
typedef enum PowerPointPtOt PowerPointPtOt;

enum PowerPointPrCt {
	PowerPointPrCtPrintColor = '\000\327\000\001',
	PowerPointPrCtPrintBlackAndWhite = '\000\327\000\002'
};
typedef enum PowerPointPrCt PowerPointPrCt;

enum PowerPointPSEL {
	PowerPointPSELSelectionTypeNone = '\000\316\000\000',
	PowerPointPSELSelectionTypeSlides = '\000\316\000\001',
	PowerPointPSELSelectionTypeShapes = '\000\316\000\002',
	PowerPointPSELSelectionTypeText = '\000\316\000\003'
};
typedef enum PowerPointPSEL PowerPointPSEL;

enum PowerPointPDtF {
	PowerPointPDtFUnset = '\000\342\377\376',
	PowerPointPDtFMdyy = '\000\343\000\001',
	PowerPointPDtFDdddMMMMddyyyy = '\000\343\000\002',
	PowerPointPDtFMMMMyyyy = '\000\343\000\003',
	PowerPointPDtFMMMMdyyyy = '\000\343\000\004',
	PowerPointPDtFMMMyy = '\000\343\000\005',
	PowerPointPDtFMMMMyy = '\000\343\000\006',
	PowerPointPDtFMMyy = '\000\343\000\007',
	PowerPointPDtFMMddyyHmm = '\000\343\000\010',
	PowerPointPDtFMddyyhmmAMPM = '\000\343\000\011',
	PowerPointPDtFHmm = '\000\343\000\012',
	PowerPointPDtFHmmss = '\000\343\000\013',
	PowerPointPDtFHmmAMPM = '\000\343\000\014',
	PowerPointPDtFHmmssAMPM = '\000\343\000\015'
};
typedef enum PowerPointPDtF PowerPointPDtF;

enum PowerPointPTnS {
	PowerPointPTnSTransitionSpeedUnset = '\000\330\377\376',
	PowerPointPTnSTransistionSpeedSlow = '\000\331\000\001',
	PowerPointPTnSTransistionSpeedMedium = '\000\331\000\002'
};
typedef enum PowerPointPTnS PowerPointPTnS;

enum PowerPointPMtv {
	PowerPointPMtvMouseActivationMouseClick = '\000\372\000\001',
	PowerPointPMtvMouseActivationMouseOver = '\000\372\000\002'
};
typedef enum PowerPointPMtv PowerPointPMtv;

enum PowerPointPAxT {
	PowerPointPAxTActionTypeUnset = '\000\345\377\376',
	PowerPointPAxTActionTypeNone = '\000\346\000\000',
	PowerPointPAxTActionTypeNextSlide = '\000\346\000\001',
	PowerPointPAxTActionTypePreviousSlide = '\000\346\000\002',
	PowerPointPAxTActionTypeFirstSlide = '\000\346\000\003',
	PowerPointPAxTActionTypeLastSlide = '\000\346\000\004',
	PowerPointPAxTActionTypeLastSlideViewed = '\000\346\000\005',
	PowerPointPAxTActionTypeEndShow = '\000\346\000\006',
	PowerPointPAxTActionTypeHyperlinkAction = '\000\346\000\007',
	PowerPointPAxTActionTypeRunMacro = '\000\346\000\010',
	PowerPointPAxTActionTypeRunProgram = '\000\346\000\011',
	PowerPointPAxTActionTypeNamedSlideshowAction = '\000\346\000\012',
	PowerPointPAxTActionTypeOLEVerb = '\000\346\000\013'
};
typedef enum PowerPointPAxT PowerPointPAxT;

enum PowerPointPPhd {
	PowerPointPPhdPlaceholderTypeUnset = '\000\332\377\376',
	PowerPointPPhdPlaceholderTypeTitlePlaceholder = '\000\333\000\001',
	PowerPointPPhdPlaceholderTypeBodyPlaceholder = '\000\333\000\002',
	PowerPointPPhdPlaceholderTypeCenterTitlePlaceholder = '\000\333\000\003',
	PowerPointPPhdPlaceholderTypeSubtitlePlaceholder = '\000\333\000\004',
	PowerPointPPhdPlaceholderTypeVerticalTitlePlaceholder = '\000\333\000\005',
	PowerPointPPhdPlaceholderTypeVerticalBodyPlaceholder = '\000\333\000\006',
	PowerPointPPhdPlaceholderTypeObjectPlaceholder = '\000\333\000\007',
	PowerPointPPhdPlaceholderTypeChartPlaceholder = '\000\333\000\010',
	PowerPointPPhdPlaceholderTypeBitmapPlaceholder = '\000\333\000\011',
	PowerPointPPhdPlaceholderTypeMediaClipPlaceholder = '\000\333\000\012',
	PowerPointPPhdPlaceholderTypeOrgChartPlaceholder = '\000\333\000\013',
	PowerPointPPhdPlaceholderTypeTablePlaceholder = '\000\333\000\014',
	PowerPointPPhdPlaceholderTypeSlideNumberPlaceholder = '\000\333\000\015',
	PowerPointPPhdPlaceholderTypeHeaderPlaceholder = '\000\333\000\016',
	PowerPointPPhdPlaceholderTypeFooterPlaceholder = '\000\333\000\017',
	PowerPointPPhdPlaceholderTypeDatePlaceholder = '\000\333\000\020',
	PowerPointPPhdPlaceholderTypeVerticalObjectPlaceholder = '\000\333\000\021',
	PowerPointPPhdPlaceholderTypePicturePlaceholder = '\000\333\000\022'
};
typedef enum PowerPointPPhd PowerPointPPhd;

enum PowerPointPSSt {
	PowerPointPSStSlideShowTypeSpeaker = '\000\325\000\001',
	PowerPointPSStSlideShowTypeWindow = '\000\325\000\002',
	PowerPointPSStSlideShowTypeKiosk = '\000\325\000\003',
	PowerPointPSStSlideShowTypePresenter = '\000\325\000\005'
};
typedef enum PowerPointPSSt PowerPointPSSt;

enum PowerPointRgTy {
	PowerPointRgTyPrintRangeAll = '\000\367\000\001',
	PowerPointRgTyPrintRangeSelection = '\000\367\000\002',
	PowerPointRgTyPrintRangeCurrent = '\000\367\000\003',
	PowerPointRgTyPrintRangeSlideRange = '\000\367\000\004',
	PowerPointRgTyPrintSection = '\000\367\000\005'
};
typedef enum PowerPointRgTy PowerPointRgTy;

enum PowerPointMedT {
	PowerPointMedTMediaTypeUnset = '\000\333\377\376',
	PowerPointMedTMediaTypeOther = '\000\334\000\001',
	PowerPointMedTMediaTypeSound = '\000\334\000\002',
	PowerPointMedTMediaTypeMovie = '\000\334\000\003'
};
typedef enum PowerPointMedT PowerPointMedT;

enum PowerPointPSFy {
	PowerPointPSFySoundFormatUnset = '\000\367\377\376',
	PowerPointPSFySoundFormatNone = '\000\370\000\000',
	PowerPointPSFySoundFormatWAV = '\000\370\000\001',
	PowerPointPSFySoundFormatMIDI = '\000\370\000\002'
};
typedef enum PowerPointPSFy PowerPointPSFy;

enum PowerPointPeBl {
	PowerPointPeBlEastAsianLineBreakNormal = '\000\354\000\001',
	PowerPointPeBlEastAsianLineBreakStrict = '\000\354\000\002',
	PowerPointPeBlEastAsianLineBreakCustom = '\000\354\000\003'
};
typedef enum PowerPointPeBl PowerPointPeBl;

enum PowerPointSRgT {
	PowerPointSRgTSlideShowRangeShowAll = '\000\326\000\001',
	PowerPointSRgTSlideShowRange = '\000\326\000\002',
	PowerPointSRgTSlideShowRangeNamedSlideshow = '\000\326\000\003'
};
typedef enum PowerPointSRgT PowerPointSRgT;

enum PowerPointFClr {
	PowerPointFClrFrameColorsBrowserColors = '\000\317\000\001',
	PowerPointFClrFrameColorsPresentationSchemeTextColor = '\000\317\000\002',
	PowerPointFClrFrameColorsPresentationSchemeAccentColor = '\000\317\000\003',
	PowerPointFClrFrameColorsWhiteTextOnBlack = '\000\317\000\004',
	PowerPointFClrFrameColorsBlackTextOnWhite = '\000\317\000\005'
};
typedef enum PowerPointFClr PowerPointFClr;

enum PowerPointPMOt {
	PowerPointPMOtMovieOptimizationNormal = '\000\317\377\376',
	PowerPointPMOtMovieOptimizationSize = '\000\320\000\001',
	PowerPointPMOtMovieOptimizationSpeed = '\000\320\000\002',
	PowerPointPMOtMovieOptimizationQuality = '\000\320\000\003'
};
typedef enum PowerPointPMOt PowerPointPMOt;

enum PowerPointPShF {
	PowerPointPShFShapeFormatGIF = '\000\335\000\000',
	PowerPointPShFShapeFormatJPEG = '\000\335\000\001',
	PowerPointPShFShapeFormatPNG = '\000\335\000\002',
	PowerPointPShFShapeFormatPICT = '\000\335\000\003'
};
typedef enum PowerPointPShF PowerPointPShF;

enum PowerPointPBrT {
	PowerPointPBrTTopBorder = '\001\032\000\001',
	PowerPointPBrTLeftBorder = '\001\032\000\002',
	PowerPointPBrTBottomBorder = '\001\032\000\003',
	PowerPointPBrTRightBorder = '\001\032\000\004',
	PowerPointPBrTDiagonalDownBorder = '\001\032\000\005',
	PowerPointPBrTDiagonalUpBorder = '\001\032\000\006'
};
typedef enum PowerPointPBrT PowerPointPBrT;

enum PowerPointCivt {
	PowerPointCivtMinorVersion = '\002\304\000\000',
	PowerPointCivtMajorVersion = '\002\304\000\001',
	PowerPointCivtOverwriteCurrentVersion = '\002\304\000\002'
};
typedef enum PowerPointCivt PowerPointCivt;

enum PowerPointPALO {
	PowerPointPALOPageLayoutNormal = '\000\355\000\000',
	PowerPointPALOPageLayoutFullScreen = '\000\355\000\001'
};
typedef enum PowerPointPALO PowerPointPALO;

enum PowerPointPBuT {
	PowerPointPBuTRegular = '\000\356\000\001',
	PowerPointPBuTFancy = '\000\356\000\002',
	PowerPointPBuTTextOnly = '\000\356\000\003'
};
typedef enum PowerPointPBuT PowerPointPBuT;

enum PowerPointPNBp {
	PowerPointPNBpBarPlacementTop = '\000\357\000\001',
	PowerPointPNBpBarPlacementBottom = '\000\357\000\002'
};
typedef enum PowerPointPNBp PowerPointPNBp;

enum PowerPointAnFX {
	PowerPointAnFXAnimationTypeCustom = '\001\002\000\000',
	PowerPointAnFXAnimationTypeAppear = '\001\002\000\001',
	PowerPointAnFXAnimationTypeFly = '\001\002\000\002',
	PowerPointAnFXAnimationTypeBlinds = '\001\002\000\003',
	PowerPointAnFXAnimationTypeBox = '\001\002\000\004',
	PowerPointAnFXAnimationTypeCheckerboard = '\001\002\000\005',
	PowerPointAnFXAnimationTypeCircle = '\001\002\000\006',
	PowerPointAnFXAnimationTypeCrawl = '\001\002\000\007',
	PowerPointAnFXAnimationTypeDiamond = '\001\002\000\010',
	PowerPointAnFXAnimationTypeDissolve = '\001\002\000\011',
	PowerPointAnFXAnimationTypeFade = '\001\002\000\012',
	PowerPointAnFXAnimationTypeFlashOnce = '\001\002\000\013',
	PowerPointAnFXAnimationTypePeek = '\001\002\000\014',
	PowerPointAnFXAnimationTypePlus = '\001\002\000\015',
	PowerPointAnFXAnimationTypeRandomBars = '\001\002\000\016',
	PowerPointAnFXAnimationTypeSpiral = '\001\002\000\017',
	PowerPointAnFXAnimationTypeSplit = '\001\002\000\020',
	PowerPointAnFXAnimationTypeStretch = '\001\002\000\021',
	PowerPointAnFXAnimationTypeStrips = '\001\002\000\022',
	PowerPointAnFXAnimationTypeSwivel = '\001\002\000\023',
	PowerPointAnFXAnimationTypeWedge = '\001\002\000\024',
	PowerPointAnFXAnimationTypeWheel = '\001\002\000\025',
	PowerPointAnFXAnimationTypeWipe = '\001\002\000\026',
	PowerPointAnFXAnimationTypeZoom = '\001\002\000\027',
	PowerPointAnFXAnimationTypeRandomEffect = '\001\002\000\030',
	PowerPointAnFXAnimationTypeBoomerang = '\001\002\000\031',
	PowerPointAnFXAnimationTypeBounce = '\001\002\000\032',
	PowerPointAnFXAnimationTypeColorReveal = '\001\002\000\033',
	PowerPointAnFXAnimationTypeCredits = '\001\002\000\034',
	PowerPointAnFXAnimationTypeEaseIn = '\001\002\000\035',
	PowerPointAnFXAnimationTypeFloat = '\001\002\000\036',
	PowerPointAnFXAnimationTypeGrowAndTurn = '\001\002\000\037',
	PowerPointAnFXAnimationTypeLightSpeed = '\001\002\000 ',
	PowerPointAnFXAnimationTypePinwheel = '\001\002\000!',
	PowerPointAnFXAnimationTypeRiseUp = '\001\002\000\"',
	PowerPointAnFXAnimationTypeSwish = '\001\002\000#',
	PowerPointAnFXAnimationTypeThinLine = '\001\002\000$',
	PowerPointAnFXAnimationTypeUnfold = '\001\002\000%',
	PowerPointAnFXAnimationTypeWhip = '\001\002\000&',
	PowerPointAnFXAnimationTypeAscend = '\001\002\000\'',
	PowerPointAnFXAnimationTypeCenterRevolve = '\001\002\000(',
	PowerPointAnFXAnimationTypeFadedSwivel = '\001\002\000)',
	PowerPointAnFXAnimationTypeDescend = '\001\002\000*',
	PowerPointAnFXAnimationTypeSling = '\001\002\000+',
	PowerPointAnFXAnimationTypeSpinner = '\001\002\000,',
	PowerPointAnFXAnimationTypeStretchy = '\001\002\000-',
	PowerPointAnFXAnimationTypeZip = '\001\002\000.',
	PowerPointAnFXAnimationTypeArcUp = '\001\002\000/',
	PowerPointAnFXAnimationTypeFadeZoom = '\001\002\0000',
	PowerPointAnFXAnimationTypeGlide = '\001\002\0001',
	PowerPointAnFXAnimationTypeExpand = '\001\002\0002',
	PowerPointAnFXAnimationTypeFlip = '\001\002\0003',
	PowerPointAnFXAnimationTypeShimmer = '\001\002\0004',
	PowerPointAnFXAnimationTypeFold = '\001\002\0005',
	PowerPointAnFXAnimationTypeChangeFillColor = '\001\002\0006',
	PowerPointAnFXAnimationTypeChangeFont = '\001\002\0007',
	PowerPointAnFXAnimationTypeChangeFontColor = '\001\002\0008',
	PowerPointAnFXAnimationTypeChangeFontSize = '\001\002\0009',
	PowerPointAnFXAnimationTypeChangeFontStyle = '\001\002\000:',
	PowerPointAnFXAnimationTypeGrowShrink = '\001\002\000;',
	PowerPointAnFXAnimationTypeChangeLineColor = '\001\002\000<',
	PowerPointAnFXAnimationTypeSpin = '\001\002\000=',
	PowerPointAnFXAnimationTypeTransparency = '\001\002\000>',
	PowerPointAnFXAnimationTypeBoldFlash = '\001\002\000\?',
	PowerPointAnFXAnimationTypeBlast = '\001\002\000@',
	PowerPointAnFXAnimationTypeBoldReveal = '\001\002\000A',
	PowerPointAnFXAnimationTypeBrushOnColor = '\001\002\000B',
	PowerPointAnFXAnimationTypeBrushOnUnderline = '\001\002\000C',
	PowerPointAnFXAnimationTypeColorBlend = '\001\002\000D',
	PowerPointAnFXAnimationTypeColorWave = '\001\002\000E',
	PowerPointAnFXAnimationTypeComplementaryColor = '\001\002\000F',
	PowerPointAnFXAnimationTypeComplementaryColor2 = '\001\002\000G',
	PowerPointAnFXAnimationTypeConstrastingColor = '\001\002\000H',
	PowerPointAnFXAnimationTypeDarken = '\001\002\000I',
	PowerPointAnFXAnimationTypeDesaturate = '\001\002\000J',
	PowerPointAnFXAnimationTypeFlashBulb = '\001\002\000K',
	PowerPointAnFXAnimationTypeFlicker = '\001\002\000L',
	PowerPointAnFXAnimationTypeGrowWithColor = '\001\002\000M',
	PowerPointAnFXAnimationTypeLighten = '\001\002\000N',
	PowerPointAnFXAnimationTypeStyleEmphasis = '\001\002\000O',
	PowerPointAnFXAnimationTypeTeeter = '\001\002\000P',
	PowerPointAnFXAnimationTypeVerticalGrow = '\001\002\000Q',
	PowerPointAnFXAnimationTypeWave = '\001\002\000R',
	PowerPointAnFXAnimationTypeMediaPlay = '\001\002\000S',
	PowerPointAnFXAnimationTypeMediaPause = '\001\002\000T',
	PowerPointAnFXAnimationTypeMediaStop = '\001\002\000U',
	PowerPointAnFXAnimationTypeCirclePath = '\001\002\000V',
	PowerPointAnFXAnimationTypeRightTrianglePath = '\001\002\000W',
	PowerPointAnFXAnimationTypeDiamondPath = '\001\002\000X',
	PowerPointAnFXAnimationTypeHexagonPath = '\001\002\000Y',
	PowerPointAnFXAnimationType5PointStarPath = '\001\002\000Z',
	PowerPointAnFXAnimationTypeCrescentMoonPath = '\001\002\000[',
	PowerPointAnFXAnimationTypeSquarePath = '\001\002\000\\',
	PowerPointAnFXAnimationTypeTrapezoidPath = '\001\002\000]',
	PowerPointAnFXAnimationTypeHeartPath = '\001\002\000^',
	PowerPointAnFXAnimationTypeOctagonPath = '\001\002\000_',
	PowerPointAnFXAnimationType6PointStarPath = '\001\002\000`',
	PowerPointAnFXAnimationTypeFootballPath = '\001\002\000a',
	PowerPointAnFXAnimationTypeEqualTrianglePath = '\001\002\000b',
	PowerPointAnFXAnimationTypeParallelogramPath = '\001\002\000c',
	PowerPointAnFXAnimationTypePentagonPath = '\001\002\000d',
	PowerPointAnFXAnimationType4PointStarPath = '\001\002\000e',
	PowerPointAnFXAnimationType8PointStarPath = '\001\002\000f',
	PowerPointAnFXAnimationTypeTeardropPath = '\001\002\000g',
	PowerPointAnFXAnimationTypePointyStarPath = '\001\002\000h',
	PowerPointAnFXAnimationTypeCurvedSquarePath = '\001\002\000i',
	PowerPointAnFXAnimationTypeCurvedXPath = '\001\002\000j',
	PowerPointAnFXAnimationTypeVerticalFigure8Path = '\001\002\000k',
	PowerPointAnFXAnimationTypeCurvyStarPath = '\001\002\000l',
	PowerPointAnFXAnimationTypeLoopDeLoopPath = '\001\002\000m',
	PowerPointAnFXAnimationTypeBuzzsawPath = '\001\002\000n',
	PowerPointAnFXAnimationTypeHorizontalFigure8Path = '\001\002\000o',
	PowerPointAnFXAnimationTypePeanutPath = '\001\002\000p',
	PowerPointAnFXAnimationTypeFigure8FourPath = '\001\002\000q',
	PowerPointAnFXAnimationTypeNeutronPath = '\001\002\000r',
	PowerPointAnFXAnimationTypeSwooshPath = '\001\002\000s',
	PowerPointAnFXAnimationTypeBeanPath = '\001\002\000t',
	PowerPointAnFXAnimationTypePlusPath = '\001\002\000u',
	PowerPointAnFXAnimationTypeInvertedTrianglePath = '\001\002\000v',
	PowerPointAnFXAnimationTypeInvertedSquarePath = '\001\002\000w',
	PowerPointAnFXAnimationTypeLeftPath = '\001\002\000x',
	PowerPointAnFXAnimationTypeTurnRightPath = '\001\002\000y',
	PowerPointAnFXAnimationTypeArcDownPath = '\001\002\000z',
	PowerPointAnFXAnimationTypeZigzagPath = '\001\002\000{',
	PowerPointAnFXAnimationTypeSCurve2Path = '\001\002\000|',
	PowerPointAnFXAnimationTypeSineWavePath = '\001\002\000}',
	PowerPointAnFXAnimationTypeBounceLeftPath = '\001\002\000~',
	PowerPointAnFXAnimationTypeDownPath = '\001\002\000\177',
	PowerPointAnFXAnimationTypeTurnUpPath = '\001\002\000\200',
	PowerPointAnFXAnimationTypeArcUpPath = '\001\002\000\201',
	PowerPointAnFXAnimationTypeHeartbeatPath = '\001\002\000\202',
	PowerPointAnFXAnimationTypeSpiralRightPath = '\001\002\000\203',
	PowerPointAnFXAnimationTypeWavePath = '\001\002\000\204',
	PowerPointAnFXAnimationTypeCurvyLeftPath = '\001\002\000\205',
	PowerPointAnFXAnimationTypeDiagonalDownRightPath = '\001\002\000\206',
	PowerPointAnFXAnimationTypeTurnDownPath = '\001\002\000\207',
	PowerPointAnFXAnimationTypeArcLeftPath = '\001\002\000\210',
	PowerPointAnFXAnimationTypeFunnelPath = '\001\002\000\211',
	PowerPointAnFXAnimationTypeSpringPath = '\001\002\000\212',
	PowerPointAnFXAnimationTypeBounceRightPath = '\001\002\000\213',
	PowerPointAnFXAnimationTypeSpiralLeftPath = '\001\002\000\214',
	PowerPointAnFXAnimationTypeDiagonalUpRightPath = '\001\002\000\215',
	PowerPointAnFXAnimationTypeTurnUpRightPath = '\001\002\000\216',
	PowerPointAnFXAnimationTypeArcRightPath = '\001\002\000\217',
	PowerPointAnFXAnimationTypeSCurve1Path = '\001\002\000\220',
	PowerPointAnFXAnimationTypeDecayingWavePath = '\001\002\000\221',
	PowerPointAnFXAnimationTypeCurvyRightPath = '\001\002\000\222',
	PowerPointAnFXAnimationTypeStairsDownPath = '\001\002\000\223',
	PowerPointAnFXAnimationTypeUpPath = '\001\002\000\224',
	PowerPointAnFXAnimationTypeRightPath = '\001\002\000\225'
};
typedef enum PowerPointAnFX PowerPointAnFX;

enum PowerPointAnLv {
	PowerPointAnLvTextByNoLevels = '\001\001\000\000',
	PowerPointAnLvTextByAllLevels = '\001\001\000\001',
	PowerPointAnLvTextByFirstLevel = '\001\001\000\002',
	PowerPointAnLvTextBySecondLevel = '\001\001\000\003',
	PowerPointAnLvTextByThirdLevel = '\001\001\000\004',
	PowerPointAnLvTextByFourthLevel = '\001\001\000\005',
	PowerPointAnLvTextByFifthLevel = '\001\001\000\006',
	PowerPointAnLvChartAllAtOnce = '\001\001\000\007',
	PowerPointAnLvChartByCategory = '\001\001\000\010',
	PowerPointAnLvChartByCtageoryElements = '\001\001\000\011',
	PowerPointAnLvChartBySeries = '\001\001\000\012',
	PowerPointAnLvChartBySeriesElements = '\001\001\000\013'
};
typedef enum PowerPointAnLv PowerPointAnLv;

enum PowerPointAnTr {
	PowerPointAnTrNoTrigger = '\001\000\000\000',
	PowerPointAnTrOnPageClick = '\001\000\000\001',
	PowerPointAnTrWithPrevious = '\001\000\000\002',
	PowerPointAnTrAfterPrevious = '\001\000\000\003',
	PowerPointAnTrOnShapeClick = '\001\000\000\004'
};
typedef enum PowerPointAnTr PowerPointAnTr;

enum PowerPointAnAE {
	PowerPointAnAENoAfterEffect = '\000\377\000\000',
	PowerPointAnAEDim = '\000\377\000\001',
	PowerPointAnAEHide = '\000\377\000\002',
	PowerPointAnAEHideOnNextClick = '\000\377\000\003'
};
typedef enum PowerPointAnAE PowerPointAnAE;

enum PowerPointAnTU {
	PowerPointAnTUByParagraph = '\000\376\000\000',
	PowerPointAnTUByCharacter = '\000\376\000\001',
	PowerPointAnTUByWord = '\000\376\000\002'
};
typedef enum PowerPointAnTU PowerPointAnTU;

enum PowerPointAnRs {
	PowerPointAnRsRestartAlways = '\000\375\000\001',
	PowerPointAnRsRestartWhenOff = '\000\375\000\002',
	PowerPointAnRsNeverRestart = '\000\375\000\003'
};
typedef enum PowerPointAnRs PowerPointAnRs;

enum PowerPointAnEA {
	PowerPointAnEAAfterFreeze = '\000\374\000\001',
	PowerPointAnEAAfterRemove = '\000\374\000\002',
	PowerPointAnEAAfterHold = '\000\374\000\003',
	PowerPointAnEAAfterTransition = '\000\374\000\004'
};
typedef enum PowerPointAnEA PowerPointAnEA;

enum PowerPointAnDi {
	PowerPointAnDiNoDirection = '\000\371\000\000',
	PowerPointAnDiUp = '\000\371\000\001',
	PowerPointAnDiRight = '\000\371\000\002',
	PowerPointAnDiDown = '\000\371\000\003',
	PowerPointAnDiLeft = '\000\371\000\004',
	PowerPointAnDiOrdinalMask = '\000\371\000\005',
	PowerPointAnDiUpLeft = '\000\371\000\006',
	PowerPointAnDiUpRight = '\000\371\000\007',
	PowerPointAnDiDownRight = '\000\371\000\010',
	PowerPointAnDiDownLeft = '\000\371\000\011',
	PowerPointAnDiTop = '\000\371\000\012',
	PowerPointAnDiBottom = '\000\371\000\013',
	PowerPointAnDiTopLeft = '\000\371\000\014',
	PowerPointAnDiTopRight = '\000\371\000\015',
	PowerPointAnDiBottomRight = '\000\371\000\016',
	PowerPointAnDiBottomLeft = '\000\371\000\017',
	PowerPointAnDiHorizontal = '\000\371\000\020',
	PowerPointAnDiVertical = '\000\371\000\021',
	PowerPointAnDiAcross = '\000\371\000\022',
	PowerPointAnDiInward = '\000\371\000\023',
	PowerPointAnDiOut = '\000\371\000\024',
	PowerPointAnDiClockwise = '\000\371\000\025',
	PowerPointAnDiCounterclockwise = '\000\371\000\026',
	PowerPointAnDiHorizontalIn = '\000\371\000\027',
	PowerPointAnDiHorizontalOut = '\000\371\000\030',
	PowerPointAnDiVerticalIn = '\000\371\000\031',
	PowerPointAnDiVerticalOut = '\000\371\000\032',
	PowerPointAnDiSlightly = '\000\371\000\033',
	PowerPointAnDiCenter = '\000\371\000\034',
	PowerPointAnDiInSlightly = '\000\371\000\035',
	PowerPointAnDiInCenter = '\000\371\000\036',
	PowerPointAnDiInBottom = '\000\371\000\037',
	PowerPointAnDiOutSlightly = '\000\371\000 ',
	PowerPointAnDiOutCenter = '\000\371\000!',
	PowerPointAnDiOutBottom = '\000\371\000\"',
	PowerPointAnDiFontBold = '\000\371\000#',
	PowerPointAnDiFontItalic = '\000\371\000$',
	PowerPointAnDiFontUnderline = '\000\371\000%',
	PowerPointAnDiFontStrikethrough = '\000\371\000&',
	PowerPointAnDiFontShadow = '\000\371\000\'',
	PowerPointAnDiFontAllCaps = '\000\371\000(',
	PowerPointAnDiInstant = '\000\371\000)',
	PowerPointAnDiGradual = '\000\371\000*',
	PowerPointAnDiCycleClockwise = '\000\371\000+',
	PowerPointAnDiCycleCounterclockwise = '\000\371\000,'
};
typedef enum PowerPointAnDi PowerPointAnDi;

enum PowerPointAnTy {
	PowerPointAnTyAnimationTypeNone = '\001\003\000\000',
	PowerPointAnTyAnimationTypeMotion = '\001\003\000\001',
	PowerPointAnTyAnimationTypeColor = '\001\003\000\002',
	PowerPointAnTyAnimationTypeScale = '\001\003\000\003',
	PowerPointAnTyAnimationTypeRotation = '\001\003\000\004',
	PowerPointAnTyAnimationTypeProperty = '\001\003\000\005',
	PowerPointAnTyAnimationTypeCommand = '\001\003\000\006',
	PowerPointAnTyAnimationTypeFilter = '\001\003\000\007',
	PowerPointAnTyAnimationTypeSet = '\001\003\000\010'
};
typedef enum PowerPointAnTy PowerPointAnTy;

enum PowerPointAnpp {
	PowerPointAnppNoAdditive = '\001\007\000\001',
	PowerPointAnppMotion = '\001\007\000\002'
};
typedef enum PowerPointAnpp PowerPointAnpp;

enum PowerPointAnSm {
	PowerPointAnSmNoAccumulate = '\001\004\000\001',
	PowerPointAnSmAlways = '\001\004\000\002'
};
typedef enum PowerPointAnSm PowerPointAnSm;

enum PowerPointAnPr {
	PowerPointAnPrNoProperty = '\001\005\000\000',
	PowerPointAnPrX = '\001\005\000\001',
	PowerPointAnPrY = '\001\005\000\002',
	PowerPointAnPrWidth = '\001\005\000\003',
	PowerPointAnPrHeight = '\001\005\000\004',
	PowerPointAnPrOpacity = '\001\005\000\005',
	PowerPointAnPrRotation = '\001\005\000\006',
	PowerPointAnPrColors = '\001\005\000\007',
	PowerPointAnPrVisibility = '\001\005\000\010',
	PowerPointAnPrTextFontBold = '\001\005\000d',
	PowerPointAnPrTextFontColor = '\001\005\000e',
	PowerPointAnPrTextFontEmboss = '\001\005\000f',
	PowerPointAnPrTextFontItalic = '\001\005\000g',
	PowerPointAnPrTextFontName = '\001\005\000h',
	PowerPointAnPrTextFontShadow = '\001\005\000i',
	PowerPointAnPrTextFontSize = '\001\005\000j',
	PowerPointAnPrTextFontSubscript = '\001\005\000k',
	PowerPointAnPrTextFontSuperscript = '\001\005\000l',
	PowerPointAnPrTextFontUnderline = '\001\005\000m',
	PowerPointAnPrTextFontStrikethrough = '\001\005\000n',
	PowerPointAnPrTextBulletCharacter = '\001\005\000o',
	PowerPointAnPrTextBulletFontName = '\001\005\000p',
	PowerPointAnPrTextBulletNumber = '\001\005\000q',
	PowerPointAnPrTextBulletColor = '\001\005\000r',
	PowerPointAnPrTextBulletRelativeSize = '\001\005\000s',
	PowerPointAnPrTextBulletStyle = '\001\005\000t',
	PowerPointAnPrTextBulletType = '\001\005\000u',
	PowerPointAnPrShapePictureContrast = '\001\005\003\350',
	PowerPointAnPrShapePictureBrightness = '\001\005\003\351',
	PowerPointAnPrShapePictureGamma = '\001\005\003\352',
	PowerPointAnPrShapePictureGrayscale = '\001\005\003\353',
	PowerPointAnPrShapeFillOn = '\001\005\003\354',
	PowerPointAnPrShapeFillColor = '\001\005\003\355',
	PowerPointAnPrShapeFillOpacity = '\001\005\003\356',
	PowerPointAnPrShapeFillBackColor = '\001\005\003\357',
	PowerPointAnPrShapeLineOn = '\001\005\003\360',
	PowerPointAnPrShapeLineColor = '\001\005\003\361',
	PowerPointAnPrShapeShadowOn = '\001\005\003\362',
	PowerPointAnPrShapeShadowType = '\001\005\003\363',
	PowerPointAnPrShapeShadowColor = '\001\005\003\364',
	PowerPointAnPrShapeShadowOpacity = '\001\005\003\365',
	PowerPointAnPrShapeShadowOffsetX = '\001\005\003\366',
	PowerPointAnPrShapeShadowOffsetY = '\001\005\003\367'
};
typedef enum PowerPointAnPr PowerPointAnPr;

enum PowerPointAnCT {
	PowerPointAnCTEvent = '\001\006\000\000',
	PowerPointAnCTCall = '\001\006\000\001',
	PowerPointAnCTVerb = '\001\006\000\002'
};
typedef enum PowerPointAnCT PowerPointAnCT;

enum PowerPointAfet {
	PowerPointAfetNoFilterEffectType = '\001\010\000\000',
	PowerPointAfetFilterEffectTypeBarn = '\001\010\000\001',
	PowerPointAfetFilterEffectTypeBlinds = '\001\010\000\002',
	PowerPointAfetFilterEffectTypeBox = '\001\010\000\003',
	PowerPointAfetFilterEffectTypeCheckerboard = '\001\010\000\004',
	PowerPointAfetFilterEffectTypeCircle = '\001\010\000\005',
	PowerPointAfetFilterEffectTypeDiamond = '\001\010\000\006',
	PowerPointAfetFilterEffectTypeDissolve = '\001\010\000\007',
	PowerPointAfetFilterEffectTypeFade = '\001\010\000\010',
	PowerPointAfetFilterEffectTypeImage = '\001\010\000\011',
	PowerPointAfetFilterEffectTypePixelate = '\001\010\000\012',
	PowerPointAfetFilterEffectTypePlus = '\001\010\000\013',
	PowerPointAfetFilterEffectTypeRandomBar = '\001\010\000\014',
	PowerPointAfetFilterEffectTypeSlide = '\001\010\000\015',
	PowerPointAfetFilterEffectTypeStretch = '\001\010\000\016',
	PowerPointAfetFilterEffectTypeStrips = '\001\010\000\017',
	PowerPointAfetFilterEffectTypeWedge = '\001\010\000\020',
	PowerPointAfetFilterEffectTypeWheel = '\001\010\000\021',
	PowerPointAfetFilterEffectTypeWipe = '\001\010\000\022'
};
typedef enum PowerPointAfet PowerPointAfet;

enum PowerPointAfes {
	PowerPointAfesNoEffectSubtype = '\001\011\000\000',
	PowerPointAfesFilterEffectSubtypeInVertical = '\001\011\000\001',
	PowerPointAfesFilterEffectSubtypeOutVertical = '\001\011\000\002',
	PowerPointAfesFilterEffectSubtypeInHorizontal = '\001\011\000\003',
	PowerPointAfesFilterEffectSubtypeOutHorizontal = '\001\011\000\004',
	PowerPointAfesFilterEffectSubtypeHorizontal = '\001\011\000\005',
	PowerPointAfesFilterEffectSubtypeVertical = '\001\011\000\006',
	PowerPointAfesFilterEffectSubtypeInward = '\001\011\000\007',
	PowerPointAfesFilterEffectSubtypeOut = '\001\011\000\010',
	PowerPointAfesFilterEffectSubtypeAcross = '\001\011\000\011',
	PowerPointAfesFilterEffectSubtypeFromLeft = '\001\011\000\012',
	PowerPointAfesFilterEffectSubtypeFromRight = '\001\011\000\013',
	PowerPointAfesFilterEffectSubtypeFromTop = '\001\011\000\014',
	PowerPointAfesFilterEffectSubtypeFromBottom = '\001\011\000\015',
	PowerPointAfesFilterEffectSubtypeDownLeft = '\001\011\000\016',
	PowerPointAfesFilterEffectSubtypeUpLeft = '\001\011\000\017',
	PowerPointAfesFilterEffectSubtypeDownRight = '\001\011\000\020',
	PowerPointAfesFilterEffectSubtypeUpRight = '\001\011\000\021',
	PowerPointAfesFilterEffectSubtypeSpoke1 = '\001\011\000\022',
	PowerPointAfesFilterEffectSubtypeSpokes2 = '\001\011\000\023',
	PowerPointAfesFilterEffectSubtypeSpokes3 = '\001\011\000\024',
	PowerPointAfesFilterEffectSubtypeSpokes4 = '\001\011\000\025',
	PowerPointAfesFilterEffectSubtypeSpokes8 = '\001\011\000\026',
	PowerPointAfesFilterEffectSubtypeLeft = '\001\011\000\027',
	PowerPointAfesFilterEffectSubtypeRight = '\001\011\000\030',
	PowerPointAfesFilterEffectSubtypeDown = '\001\011\000\031',
	PowerPointAfesFilterEffectSubtypeUp = '\001\011\000\032'
};
typedef enum PowerPointAfes PowerPointAfes;

enum PowerPoint4001 {
	PowerPoint4001View = 'pVEW',
	PowerPoint4001SlideShowView = 'PSSv'
};
typedef enum PowerPoint4001 PowerPoint4001;

enum PowerPoint4003 {
	PowerPoint4003View = 'pVEW',
	PowerPoint4003Presentation = 'pptP'
};
typedef enum PowerPoint4003 PowerPoint4003;

enum PowerPoint4002 {
	PowerPoint4002Slide = 'pSLD',
	PowerPoint4002Master = 'pMtr',
	PowerPoint4002Presentation = 'pptP'
};
typedef enum PowerPoint4002 PowerPoint4002;

enum PowerPoint4011 {
	PowerPoint4011Shape = 'pShp',
	PowerPoint4011FillFormat = 'pFFm'
};
typedef enum PowerPoint4011 PowerPoint4011;

enum PowerPoint4016 {
	PowerPoint4016Shape = 'pShp',
	PowerPoint4016FillFormat = 'pFFm'
};
typedef enum PowerPoint4016 PowerPoint4016;

enum PowerPoint4024 {
	PowerPoint4024Callout = 'cD00',
	PowerPoint4024CalloutFormat = 'cCoF'
};
typedef enum PowerPoint4024 PowerPoint4024;

enum PowerPoint4019 {
	PowerPoint4019Connector = 'cD01',
	PowerPoint4019ConnectorFormat = 'pCxF'
};
typedef enum PowerPoint4019 PowerPoint4019;

enum PowerPoint4020 {
	PowerPoint4020Connector = 'cD01',
	PowerPoint4020ConnectorFormat = 'pCxF'
};
typedef enum PowerPoint4020 PowerPoint4020;

enum PowerPoint4026 {
	PowerPoint4026Callout = 'cD00',
	PowerPoint4026CalloutFormat = 'cCoF'
};
typedef enum PowerPoint4026 PowerPoint4026;

enum PowerPoint4025 {
	PowerPoint4025Callout = 'cD00',
	PowerPoint4025CalloutFormat = 'cCoF'
};
typedef enum PowerPoint4025 PowerPoint4025;

enum PowerPoint4021 {
	PowerPoint4021Connector = 'cD01',
	PowerPoint4021ConnectorFormat = 'pCxF'
};
typedef enum PowerPoint4021 PowerPoint4021;

enum PowerPoint4022 {
	PowerPoint4022Connector = 'cD01',
	PowerPoint4022ConnectorFormat = 'pCxF'
};
typedef enum PowerPoint4022 PowerPoint4022;

enum PowerPoint4012 {
	PowerPoint4012Shape = 'pShp',
	PowerPoint4012FillFormat = 'pFFm'
};
typedef enum PowerPoint4012 PowerPoint4012;

enum PowerPoint4010 {
	PowerPoint4010Shape = 'pShp',
	PowerPoint4010ShapeRange = 'ShpR'
};
typedef enum PowerPoint4010 PowerPoint4010;

enum PowerPoint4015 {
	PowerPoint4015Shape = 'pShp',
	PowerPoint4015FillFormat = 'pFFm'
};
typedef enum PowerPoint4015 PowerPoint4015;

enum PowerPoint4023 {
	PowerPoint4023Shape = 'pShp',
	PowerPoint4023ThreeDFormat = 'D3Df'
};
typedef enum PowerPoint4023 PowerPoint4023;

enum PowerPoint4017 {
	PowerPoint4017Shape = 'pShp',
	PowerPoint4017FillFormat = 'pFFm'
};
typedef enum PowerPoint4017 PowerPoint4017;

enum PowerPoint4018 {
	PowerPoint4018Shape = 'pShp',
	PowerPoint4018FillFormat = 'pFFm'
};
typedef enum PowerPoint4018 PowerPoint4018;

enum PowerPoint4009 {
	PowerPoint4009Shape = 'pShp',
	PowerPoint4009ShapeRange = 'ShpR'
};
typedef enum PowerPoint4009 PowerPoint4009;

enum PowerPoint4014 {
	PowerPoint4014Shape = 'pShp',
	PowerPoint4014FillFormat = 'pFFm'
};
typedef enum PowerPoint4014 PowerPoint4014;

enum PowerPoint4013 {
	PowerPoint4013Shape = 'pShp',
	PowerPoint4013FillFormat = 'pFFm'
};
typedef enum PowerPoint4013 PowerPoint4013;

enum PowerPoint4004 {
	PowerPoint4004Shape = 'pShp',
	PowerPoint4004ShapeRange = 'ShpR'
};
typedef enum PowerPoint4004 PowerPoint4004;

enum PowerPoint4008 {
	PowerPoint4008Shape = 'pShp',
	PowerPoint4008ShapeRange = 'ShpR'
};
typedef enum PowerPoint4008 PowerPoint4008;

enum PowerPoint4005 {
	PowerPoint4005Shape = 'pShp',
	PowerPoint4005ShapeRange = 'ShpR'
};
typedef enum PowerPoint4005 PowerPoint4005;

enum PowerPoint4006 {
	PowerPoint4006Shape = 'pShp',
	PowerPoint4006ShapeRange = 'ShpR'
};
typedef enum PowerPoint4006 PowerPoint4006;

enum PowerPoint4007 {
	PowerPoint4007Shape = 'pShp',
	PowerPoint4007ShapeRange = 'ShpR',
	PowerPoint4007ShapeRange = 'ShpR'
};
typedef enum PowerPoint4007 PowerPoint4007;

@protocol PowerPointGenericMethods

- (void) closeSaving:(PowerPointSavo)saving savingIn:(PowerPointPPfd)savingIn;  // Close an object
- (NSInteger) dataSizeAs:(NSNumber *)as;  // Return the size in bytes of an object
- (void) delete;  // Delete an element from an object
- (SBObject *) duplicateTo:(SBObject *)to;  // Duplicate object(s)
- (BOOL) exists;  // Verify if an object exists
- (SBObject *) moveTo:(SBObject *)to;  // Move object(s) to a new location
- (void) openPassword:(NSString *)password writePassword:(NSString *)writePassword;  // Open the specified object(s)
- (void) saveIn:(PowerPointPPfd)in_ as:(PowerPointPPff)as;  // Save an object
- (void) select;  // Make a selection
- (BOOL) canCheckOutFileName:(NSString *)fileName;  // Returns True if PowerPoint can check out a specified presentation from a server.
- (void) checkOutFileName:(NSString *)fileName;  // Copies a specified presentation from a server to a local computer for editing. Returns a String that represents the local path and file name of the presentation checked out.
- (void) quit;
- (void) setPrinterSettingsPrinterSettings:(NSInteger)printerSettings;

@end



/*
 * Standard Suite
 */

// A scriptable object
@interface PowerPointBaseObject : SBObject <PowerPointGenericMethods>

@property (copy) NSDictionary *properties;  // All of the object's properties


@end

// A basic application program
@interface PowerPointBaseApplication : PowerPointBaseObject

@property (readonly) BOOL frontmost;  // Is this the frontmost application?
@property (copy, readonly) NSString *name;  // the name
@property (copy, readonly) NSString *version;  // the version of the application


@end

// A basic document
@interface PowerPointBaseDocument : PowerPointBaseObject

@property (readonly) BOOL modified;  // Has the document been modified since the last save?
@property (copy, readonly) NSString *name;  // the name


@end

// A basic window
@interface PowerPointBasicWindow : PowerPointBaseObject

@property NSRect bounds;  // the boundary rectangle for the window
@property (readonly) BOOL closeable;  // Does the window have a close box?
@property (readonly) BOOL titled;  // Does the window have a title bar?
@property (readonly) NSInteger entryIndex;  // the number of the window
@property (readonly) BOOL floating;  // Does the window float?
@property (readonly) BOOL modal;  // Is the window modal?
@property NSPoint position;  // upper left coordinates of the window
@property (readonly) BOOL resizable;  // Is the window resizable?
@property (readonly) BOOL zoomable;  // Is the window zoomable?
@property BOOL zoomed;  // Is the window zoomed?
@property (copy, readonly) NSString *name;  // the title of the window
@property (readonly) BOOL visible;  // Is the window visible?
@property (readonly) BOOL collapsable;  // Is the window collapasable?
@property BOOL collapsed;  // Is the window collapsed?
@property (readonly) BOOL sheet;  // Is this window a sheet window?


@end

@interface PowerPointPrintSettings : SBObject <PowerPointGenericMethods>

@property NSInteger copies;  // the number of copies of a document to be printed 
@property BOOL collating;  // Should printed copies be collated?
@property NSInteger startingPage;  // the first page of the document to be printed
@property NSInteger endingPage;  // the last page of the document to be printed
@property NSInteger pagesAcross;  // number of logical pages laid across a physical page
@property NSInteger pagesDown;  // number of logical pages laid out down a physical page
@property (copy) NSDate *requestedPrintTime;  // the time at which the desktop printer should print the document...
@property (copy) id errorHandling;  // how errors are handled
@property (copy) NSString *faxNumber;  // for fax number
@property (copy) NSString *targetPrinter;  // the queue name of the target printer


@end



/*
 * Microsoft Office Suite
 */

// A control within a command bar.
@interface PowerPointCommandBarControl : PowerPointBaseObject

@property BOOL beginGroup;  // Returns or sets if the command bar control appears at the beginning of a group of controls on the command bar.
@property (readonly) BOOL builtIn;  // Returns true if the command bar control is a built-in command bar control.
@property (readonly) PowerPointMCLT controlType;  // Returns the type of the command bar control.
@property (copy) NSString *descriptionText;  // Returns or sets the description for a command bar control.  The description is not displayed to the user, but it can be useful for documenting the behavior of a control.
@property BOOL enabled;  // Returns or sets if the command bar control is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number for this command bar control.
@property NSInteger height;  // Returns or sets the height of a command bar control.
@property NSInteger helpContextID;  // Returns or sets the help context ID number for the Help topic attached to the command bar control.
@property (copy) NSString *helpFile;  // Returns or sets the file name for the help topic attached to the command bar.  To use this property, you must also set the help context ID property.
- (NSInteger) id;  // Returns the id for a built-in command bar control.
@property (readonly) NSInteger leftPosition;  // Returns the left position of the command bar control.
@property (copy) NSString *name;  // Returns or sets the caption text for a command bar control.
@property (copy) NSString *parameter;  // Returns or sets a string that is used to execute a command.
@property NSInteger priority;  // Returns or sets the priority of a command bar control. A controls priority determines whether the control can be dropped from a docked command bar if the command bar controls can not fit in a single row.  Valid priority number are 0 through 7.
@property (copy) NSString *tag;  // Returns or sets information about the command bar control, such as data that can be used as an argument in procedures, or information that identifies the control.
@property (copy) NSString *tooltipText;  // Returns or sets the text displayed in a command bar controls tooltip.
@property (readonly) NSInteger top;  // Returns the top position of a command bar control.
@property BOOL visible;  // Returns or sets if the command bar control is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar control.

- (void) execute;  // Runs the procedure or built-in command assigned to the specified command bar control.

@end

// A button control within a command bar.
@interface PowerPointCommandBarButton : PowerPointCommandBarControl

@property (readonly) BOOL buttonFaceIsDefault;  // Returns if the face of a command bar button control is the original built-in face.
@property PowerPointMBns buttonState;  // Returns or set the appearance of a command bar button control.  The property is read-only for built-in command bar buttons.
@property PowerPointMBTs buttonStyle;  // Returns or sets the way a command button control is displayed.
@property NSInteger faceId;  // Returns or sets the Id number for the face of the command bar button control.


@end

// A combobox menu control within a command bar.
@interface PowerPointCommandBarCombobox : PowerPointCommandBarControl

@property PowerPointMXcb comboboxStyle;  // Returns or sets the way a command bar combobox control is displayed.
@property (copy) NSString *comboboxText;  // Returns or sets the text in the display or edit portion of the command bar combobox control.
@property NSInteger dropDownLines;  // Returns or sets the number of lines in a command bar control combobox control.  The combobox control must be a custom control.
@property NSInteger dropDownWidth;  // Returns or sets the width in pixels of the list for the specified command bar combobox control.  An error occurs if you attempt to set this property for a built-in combobox control.
@property NSInteger listIndex;

- (void) addItemToComboboxComboboxItem:(NSString *)comboboxItem entry_index:(NSInteger)entry_index;  // Add a new string to a custom combobox control.
- (void) clearCombobox;  // Clear all of the strings form a custom combobox.
- (NSString *) getComboboxItemEntry_index:(NSInteger)entry_index;  // Return the string at the given index within a combobox.
- (NSInteger) getCountOfComboboxItems;  // Return the number of strings within a combobox.
- (void) removeAnItemFromComboboxEntry_index:(NSInteger)entry_index;  // Remove a string from a custom combobox.
- (void) setComboboxItemEntry_index:(NSInteger)entry_index comboboxItem:(NSString *)comboboxItem;  // Set the string an a given index for a custom combobox.

@end

// A popup menu control within a command bar.
@interface PowerPointCommandBarPopup : PowerPointCommandBarControl

- (SBElementArray<PowerPointCommandBarControl *> *) commandBarControls;


@end

// Toolbars used in all of the Office applications.
@interface PowerPointCommandBar : PowerPointBaseObject

- (SBElementArray<PowerPointCommandBarControl *> *) commandBarControls;

@property PowerPointMBPS barPosition;  // Returns or sets the position of the command bar.
@property (readonly) PowerPointMBTY barType;  // Returns the type of this command bar.
@property (readonly) BOOL builtIn;  // True if the command bar is built-in.
@property (copy, readonly) NSString *context;  // Returns or sets a string that determines where a command bar will be saved.
@property (readonly) BOOL embeddable;  // Returns if the command bar can be displayed inside the document window.
@property BOOL embedded;  // Returns or sets if the command bar will be displayed inside the document window.
@property BOOL enabled;  // Returns or set if the command bar is enabled.
@property (readonly) NSInteger entry_index;  // The index of the command bar.
@property NSInteger height;  // Returns or sets the height of the command bar.
@property NSInteger leftPosition;  // Returns or sets the left position of the command bar.
@property (copy) NSString *localName;  // Returns or sets the name of the command bar in the localized language of the application.  An error is returned when trying to set the local name of a built-in command bar.
@property (copy) NSString *name;  // Returns or sets the name of the command bar.
@property (copy) NSArray *protection;  // Returns or sets the way a command bar is protected from user customization.  It accepts a list of the following items: no protection, no customize, no resize, no move, no change visible, no change dock, no vertical dock, no horizontal dock.
@property NSInteger rowIndex;  // Returns or sets the docking order of a command bar in relation to other command bars in the same docking area.  Can be an integer greater than zero.
@property NSInteger top;  // Returns or sets the top position of a command bar.
@property BOOL visible;  // Returns or sets if the command bar is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar.


@end

@interface PowerPointDocumentProperty : PowerPointBaseObject

@property (copy) NSNumber *documentPropertyType;  // Returns or sets the document property type.
@property (copy) NSString *linkSource;  // Returns or sets the source of a lined custom document property.
@property BOOL linkToContent;  // True if the value of the document property is lined to the content of the container document.  False if the value is static.  This only applies to custom document properties.  For built-in properties this is always false.
@property (copy) NSString *name;  // Returns or sets the name of the document property.
@property (copy) NSString *value;  // Returns or sets the value of the document property.


@end

@interface PowerPointCustomDocumentProperty : PowerPointDocumentProperty


@end

@interface PowerPointWebPageFont : PowerPointBaseObject

@property (copy) NSString *fixedWidthFont;  // Returns or sets the fixed-width font setting.
@property double fixedWidthFontSize;  // Returns or sets the fixed-width font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.
@property (copy) NSString *proportionalFont;  // Returns or sets the proportional font setting.
@property double proportionalFontSize;  // Returns or sets the proportional font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.


@end



/*
 * Microsoft PowerPoint Suite
 */

@interface PowerPointActionSetting : PowerPointBaseObject

@property PowerPointPAxT action;
@property (copy) NSString *actionSettingToRun;
@property (copy, readonly) PowerPointSoundEffect *actionSoundEffect;
@property (copy) NSString *actionVerb;
@property BOOL animateAction;
@property (copy, readonly) PowerPointHyperlink *hyperlink;
@property (copy) NSString *slideShowName;


@end

@interface PowerPointAddIn : PowerPointBaseObject

@property BOOL autoLoad;
@property (copy, readonly) NSString *fullName;
@property BOOL loaded;
@property (copy, readonly) NSString *name;
@property (copy, readonly) NSString *path;
@property BOOL registered;
@property (readonly) BOOL registeredInHKLM;


@end

@interface PowerPointAnimationBehavior : PowerPointBaseObject

@property PowerPointAnSm accumulate;
@property PowerPointAnpp additive;
@property PowerPointAnTy animationBehaviorType;
@property (copy, readonly) PowerPointColorsEffect *colorsEffect;
@property (copy, readonly) PowerPointCommandEffect *commandEffect;
@property (copy, readonly) PowerPointFilterEffect *filterEffect;
@property (copy, readonly) PowerPointMotionEffect *motionEffect;
@property (copy, readonly) PowerPointPropertyEffect *propertyEffect;
@property (copy, readonly) PowerPointRotatingEffect *rotatingEffect;
@property (copy, readonly) PowerPointScaleEffect *scaleEffect;
@property (copy, readonly) PowerPointSetEffect *setEffect;
@property (copy, readonly) PowerPointTiming *timing;


@end

@interface PowerPointAnimationPoint : PowerPointBaseObject

@property (copy) NSString *formula;
@property double time;
@property (copy) SBObject *value;


@end

@interface PowerPointAnimationSettings : PowerPointBaseObject

@property double advanceTime;
@property PowerPointAsAe afterEffect;
@property BOOL animate;
@property BOOL animateBackground;
@property BOOL animateTextInReverse;
@property NSInteger animationOrder;
@property (copy, readonly) PowerPointPlaySettings *animationPlaySettings;
@property (copy, readonly) PowerPointSoundEffect *animationSoundEffect;
@property PowerPointPCuE chartUnitEffect;
@property (copy) NSColor *dimColor;
@property PowerPointMCoI dimColorThemeIndex;
@property PowerPointPEeF entryEffect;
@property PowerPointPTlE textLevelEffect;
@property PowerPointPTuE textUnitEffect;


@end

// An AppleScript object representing the Microsoft POWERPOINT application.
@interface PowerPointApplication : SBApplication

- (SBElementArray<PowerPointPresentation *> *) presentations;
- (SBElementArray<PowerPointDocumentWindow *> *) documentWindows;
- (SBElementArray<PowerPointSlideShowWindow *> *) slideShowWindows;
- (SBElementArray<PowerPointCommandBar *> *) commandBars;
- (SBElementArray<PowerPointAddIn *> *) addIns;

@property (copy, readonly) NSString *Version;
@property (copy, readonly) PowerPointPresentation *activePresentation;
@property (copy, readonly) NSString *activePrinter;
@property (copy, readonly) PowerPointDocumentWindow *activeWindow;
@property (copy, readonly) PowerPointAutocorrect *autocorrectObject;  // Returns the autocorrect object
@property (copy, readonly) NSString *build;
@property (copy, readonly) NSString *caption;
@property PowerPointPSAT defaultSaveFormat;
@property (copy, readonly) PowerPointDefaultWebOptions *defaultWebOptionsObject;
@property NSInteger keyboardScript;  // Returns the current keyboard script
@property (copy, readonly) NSString *name;
@property (copy, readonly) NSString *operatingSystem;
@property (copy, readonly) NSString *path;
@property (copy, readonly) PowerPointPreferences *preferenceObject;
@property (copy, readonly) PowerPointSaveAsMovieSettings *saveAsMovieSettingsObject;
@property BOOL startUpDialog;

- (void) print:(NSArray<SBObject *> *)x withProperties:(PowerPointPrintSettings *)withProperties;  // Print the specified object(s)
- (void) quitSaving:(PowerPointSavo)saving;  // Quit an application program
- (void) select;  // Make a selection
- (void) reset:(PowerPoint4000)x;  // Resets a built-in command bar or command bar control to its default configuration.
- (void) applyTheme:(PowerPoint4002)x fileName:(NSString *)fileName;  // Applies a theme or design template to the specified slide, master or presentation
- (void) arrangeWindows:(PowerPointPArS)x;  // Arrange Document Windows
- (PowerPointPlayer *) getPlayerFrom:(PowerPoint4001)x for:(PowerPointShape *)for_;  // get a player from a shape inside a slide show view
- (void) insertTheText:(NSString *)theText at:(SBObject *)at;
- (void) pasteObject:(PowerPoint4003)x;
- (PowerPointAddIn *) registerAddIn:(NSString *)x;
- (NSInteger) runVBMacroMacroName:(NSString *)macroName listOfParameters:(NSArray *)listOfParameters;  // Runs a Visual Basic macro.
- (void) apply:(PowerPoint4004)x;
- (void) automaticLength:(PowerPoint4024)x;
- (void) beginConnect:(PowerPoint4019)x connectedShape:(PowerPointShape *)connectedShape connectionSite:(NSInteger)connectionSite;
- (void) beginDisconnect:(PowerPoint4020)x;
- (void) customDrop:(PowerPoint4025)x dropAmount:(double)dropAmount;
- (void) customLength:(PowerPoint4026)x length:(double)length;
- (void) endConnect:(PowerPoint4021)x connectedShape:(PowerPointShape *)connectedShape connectionSite:(NSInteger)connectionSite;
- (void) endDisconnect:(PowerPoint4022)x;
- (void) flip:(PowerPoint4008)x direction:(PowerPointMFlC)direction;
- (PowerPointActionSetting *) getActionSettingFor:(PowerPoint4010)x event:(PowerPointPMtv)event;
- (void) oneColorGradient:(PowerPoint4011)x style:(PowerPointMGdS)style variant:(NSInteger)variant degree:(double)degree;
- (void) patterned:(PowerPoint4012)x pattern:(PowerPointPpTy)pattern;
- (void) pickUp:(PowerPoint4005)x;
- (void) presetGradient:(PowerPoint4013)x style:(PowerPointMGdS)style variant:(NSInteger)variant gradientType:(PowerPointMPGb)gradientType;
- (void) presetTextured:(PowerPoint4014)x texture:(PowerPointMPzT)texture;
- (void) rerouteConnections:(PowerPoint4006)x;
- (void) resetRotation:(PowerPoint4023)x;  // Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.
- (void) setShapesDefaultProperties:(PowerPoint4007)x;
- (void) solid:(PowerPoint4015)x;
- (void) twoColorGradient:(PowerPoint4016)x style:(PowerPointMGdS)style variant:(NSInteger)variant;
- (void) userPicture:(PowerPoint4017)x pictureFile:(NSString *)pictureFile;
- (void) userTextured:(PowerPoint4018)x textureFile:(NSString *)textureFile;
- (void) zOrder:(PowerPoint4009)x zOrderPosition:(PowerPointMZoC)zOrderPosition;

@end

// Represents a single autocorrect entry.
@interface PowerPointAutocorrectEntry : PowerPointBaseObject

@property (copy, readonly) NSString *autocorrectValue;  // Returns the value of the auto correct entry.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of the auto correct entry.


@end

// Represents the autocorrect functionality in PowerPoint.
@interface PowerPointAutocorrect : PowerPointBaseObject

- (SBElementArray<PowerPointAutocorrectEntry *> *) autocorrectEntries;
- (SBElementArray<PowerPointFirstLetterException *> *) firstLetterExceptions;
- (SBElementArray<PowerPointTwoInitialCapsException *> *) twoInitialCapsExceptions;

@property BOOL correctDays;  // Returns if PowerPoint automatically capitalizes the first letter of days of the week.
@property BOOL correctInitialCaps;  // Returns if PowerPoint automatically makes the second letter lowercase if the first two letters of a word are typed in uppercase. For example, POwerPoint is corrected to PowerPoint.
@property BOOL correctSentenceCaps;  // Returns if PowerPoint automatically capitalizes the first letter in each sentence.
@property BOOL replaceText;  // Returns if Microsoft PowerPoint automatically replaces specified text with entries from the autocorrect list.


@end

// Represents an interface for broadcasting presentations
@interface PowerPointBroadcast : PowerPointBaseObject

- (void) endSession;  // End a running broadcast session.
- (NSString *) getAttendeeURL;
- (BOOL) getIsBroadcasting;  // Returns true if the current presentation is being broadcast.
- (void) startServerUrl:(NSString *)serverUrl;  // Starts broadcasting to the given URL.

@end

@interface PowerPointBulletFormat : PowerPointBaseObject

@property (copy) NSString *bulletCharacter;  // Returns or sets the Unicode character value that is used for bullets in the specified text.
@property (copy, readonly) PowerPointFont *bulletFont;  // Returns a font object that represents character formatting for a bullet format object.
@property (readonly) NSInteger bulletNumber;  // Returns the bullet number of a paragraph.
@property NSInteger bulletStartValue;
@property PowerPointPBtS bulletStyle;  // Returns or sets a constant that represents the style of a bullet.
@property PowerPointPBtT bulletType;  // Returns or sets a constant that represents the type of bullet.
@property double relativeSize;  // Returns or sets the bullet size relative to the size of the first text character in the paragraph.
@property BOOL useTextColor;  // Determines whether the specified bullets are set to the color of the first text character in the paragraph.
@property BOOL useTextFont;  // Determines whether the specified bullets are set to the font of the first text character in the paragraph.
@property BOOL visible;  // Returns or sets a value that specifies whether the bullet is visible.

- (void) setBulletPicturePictureFile:(NSString *)pictureFile;  // Sets the graphics file to be used for bullets in a bulleted list.

@end

@interface PowerPointColorScheme : PowerPointBaseObject

- (NSColor *) getColorFromAt:(PowerPointPCSi)at;
- (void) setColorForAt:(PowerPointPCSi)at toColor:(NSColor *)toColor;

@end

@interface PowerPointColorsEffect : PowerPointBaseObject

@property (copy) NSColor *by_color;
@property PowerPointMCoI by_colorThemeIndex;  // Returns an object that represents a change to the color of the object by the specified number, expressed in RGB format.
@property (copy) NSColor *from_color;
@property PowerPointMCoI from_colorThemeIndex;  // Returns or sets an object that represents the starting RGB color value of an animation behavior.
@property (copy) NSColor *to_color;
@property PowerPointMCoI to_colorThemeIndex;  // Returns or sets an object that represents the RGB color value of an animation behavior.


@end

@interface PowerPointCommandEffect : PowerPointBaseObject

@property (copy) NSString *command;
@property PowerPointAnCT type;


@end

@interface PowerPointCustomLayout : PowerPointBaseObject

- (SBElementArray<PowerPointShape *> *) shapes;

@property (copy, readonly) PowerPointShape *background;
@property (copy, readonly) PowerPointDesign *design;
@property BOOL displayMasterShapes;
@property BOOL followMasterBackground;
@property (copy, readonly) PowerPointHeadersAndFooters *headersAndFooters;
@property (readonly) double height;
@property (copy, readonly) PowerPointThemeColorScheme *themeColorScheme;  // Returns the color scheme of a custom layout.
@property (copy, readonly) PowerPointTimeline *timeline;
@property (readonly) double width;


@end

@interface PowerPointDefaultWebOptions : PowerPointBaseObject

@property BOOL allowPNG;
@property BOOL alwaysSaveInDefaultEncoding;
@property BOOL checkIfOfficeIsHTMLEditor;
@property PowerPointMtEn encoding;
@property BOOL updateLinksOnSave;


@end

@interface PowerPointDesign : PowerPointBaseObject

@property (copy, readonly) PowerPointMaster *slideMaster;


@end

@interface PowerPointDocumentWindow : PowerPointBasicWindow

- (SBElementArray<PowerPointPane *> *) panes;

@property (readonly) BOOL active;
@property (copy, readonly) PowerPointPane *activePane;
@property BOOL blackAndWhite;
@property (copy, readonly) NSString *caption;
@property (readonly) NSInteger entry_index;
@property double height;
@property double leftPosition;
@property (copy, readonly) PowerPointPresentation *presentation;
@property (copy, readonly) PowerPointSelection *selection;  // Represents the selection in the specified document window.
@property NSInteger splitHorizontal;
@property NSInteger splitVertical;
@property double top;
@property (copy, readonly) PowerPointView *view;
@property PowerPointPVTy viewType;
@property double width;

- (void) collapseSectionAtPosition:(NSInteger)atPosition;
- (void) expandSectionAtPosition:(NSInteger)atPosition;
- (BOOL) getIsExpandedOfSectionAtPosition:(NSInteger)atPosition;
- (void) launchSpellerOn;

@end

@interface PowerPointEffectInformation : PowerPointBaseObject

@property (readonly) PowerPointAnAE afterEffectInformation;
@property (readonly) BOOL animateBackgroundInformation;
@property (readonly) BOOL animateTextInReverseInformation;
@property (readonly) PowerPointAnLv buildByLevel;
@property (copy, readonly) NSColor *dim;
@property (copy, readonly) PowerPointPlaySettings *playSettingsInformation;
@property (copy, readonly) PowerPointSoundEffect *soundEffectInformation;
@property (readonly) PowerPointAnTU textUnitEffectInformation;


@end

@interface PowerPointEffectParameters : PowerPointBaseObject

@property double amount;
@property (copy, readonly) NSColor *color2;
@property (readonly) PowerPointMCoI color2ColorThemeIndex;  // Returns an object that represents the color on which to end a color-cycle animation.
@property PowerPointAnDi direction;
@property (copy) NSString *font2;
@property BOOL relative;
@property double size2;


@end

@interface PowerPointEffect : PowerPointBaseObject

- (SBElementArray<PowerPointAnimationBehavior *> *) animationBehaviors;

@property PowerPointAnFX animationEffectType;
@property (copy, readonly) PowerPointEffectInformation *effectInformation;
@property (copy, readonly) PowerPointEffectParameters *effectParameters;
@property BOOL exitAnimation;
@property (copy, readonly) NSString *name;
@property (copy) PowerPointShape *shape;
@property NSInteger targetParagraph;
@property (readonly) NSInteger textRangeLength;
@property (readonly) NSInteger textRangeStart;
@property (copy, readonly) PowerPointTiming *timing;

- (PowerPointAnimationBehavior *) addBehaviorType:(PowerPointAnTy)type;  // add an animation behavior

@end

@interface PowerPointFilterEffect : PowerPointBaseObject

@property PowerPointAfet filterType;
@property BOOL reveal;
@property PowerPointAfes subtype;


@end

// Represents an abbreviation excluded from automatic correction.
@interface PowerPointFirstLetterException : PowerPointBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of the FirstLetterException.


@end

// Contains font attributes, such as font name, size, and color, for an object.
@interface PowerPointFont : PowerPointBaseObject

@property (copy) NSString *ASCIIName;  // Returns or sets the font used for Latin text; characters with character codes from 0 through 127.
@property BOOL autoRotateNumbers;  // Returns or sets a value that specifies whether the numbers in a numbered list should be rotated when the text is rotated.
@property double baseLineOffset;  // Returns or sets a value specifying the horizontaol offset of the selected font.
@property BOOL bold;  // Returns or sets a value specifying whether the font should be bold.
@property PowerPointMTCa capsType;  // Returns or sets a value specifying how the text should be capitalized.
@property (copy) NSString *eastAsianName;  // Returns or sets the font name used for Asian text.
@property (readonly) BOOL embedable;  // Returns a value indicating whether the font can be embedded in a page.
@property (readonly) BOOL embedded;  // Returns a value specifying whether the font is embedded in a page.
@property BOOL emboss;
@property BOOL equalizeCharacterHeight;  // Returns or sets a value specifying whether the text should have the same horizontal height.
@property (copy, readonly) PowerPointFillFormat *fillFormat;  // Returns a value specifying the fill formatting for text.
@property (copy) NSColor *fontColor;
@property PowerPointMCoI fontColorThemeIndex;  // Returns or sets the color for the specified font.
@property (copy) NSString *fontName;  // Returns or sets a value specifying the font to use for a selection.
@property (copy) NSString *fontNameOther;  // Returns or sets the font used for characters whose character set numbers are greater than 127.
@property double fontSize;
@property (copy, readonly) PowerPointGlowFormat *glowFormat;  // Returns a value specifying the glow formatting of the text.
@property (copy) NSColor *highlightColor;  // Returns or sets the text highlight color for object.
@property PowerPointMCoI highlightColorThemeIndex;  // Returns or sets the specified text highlight color for object.
@property BOOL italic;
@property double kerning;  // Returns or sets a value specifying the amount of spacing between text characters.
@property (copy, readonly) PowerPointLineFormat *lineFormat;  // Returns a value specifiying the line formatting of the text.
@property (copy, readonly) PowerPointReflectionFormat *reflectionFormat;  // Returns a value specifying the reflection formatting of the text.
@property (copy, readonly) PowerPointShadowFormat *shadowFormat;  // Returns the value specifying the type of shadow effect for the selection of text.
@property PowerPointMSET softEdgeFormat;  // Returns or sets a value specifying the soft edge fromatting of the text.
@property double spacing;  // Returns or sets a value specifying the spacing between characters in a selection of text.
@property PowerPointMTSt strikeType;  // Returns or sets a value specifying the strike format used for a selection of text.
@property BOOL subscript;  // Returns or sets a value specifying that the selected text should be displayed a subscript.
@property BOOL superscript;  // Returns or sets a value specifying that the selected text should be displayed a superscript.
@property double transparency;
@property BOOL underline;
@property (copy) NSColor *underlineColor;  // Returns or sets the 24-bit color of the underline for the specified font object.
@property PowerPointMCoI underlineColorThemeIndex;  // Returns a value specifying the color of the underline for the selected text.
@property PowerPointMTUl underlineStyle;  // Returns or sets a value specifying the underline style for the selected text.
@property PowerPointMPXF wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.


@end

@interface PowerPointHeaderOrFooter : PowerPointBaseObject

@property PowerPointPDtF dateFormat;
@property (copy) NSString *headerFooterText;
@property BOOL useDateFormat;
@property BOOL visible;


@end

@interface PowerPointHeadersAndFooters : PowerPointBaseObject

@property (copy, readonly) PowerPointHeaderOrFooter *dateAndTime;
@property BOOL displayHeadersFootersOnTitleSlide;
@property (copy, readonly) PowerPointHeaderOrFooter *footer;
@property (copy, readonly) PowerPointHeaderOrFooter *header;
@property (copy, readonly) PowerPointHeaderOrFooter *slideNumber;


@end

@interface PowerPointHyperlink : PowerPointBaseObject

@property (copy) NSString *hyperlinkAddress;
@property (copy) NSString *hyperlinkSubAddress;
@property (readonly) PowerPointMHyT hyperlinkType;


@end

@interface PowerPointMaster : PowerPointBaseObject

- (SBElementArray<PowerPointShape *> *) shapes;
- (SBElementArray<PowerPointHyperlink *> *) hyperlinks;
- (SBElementArray<PowerPointCustomLayout *> *) customLayouts;

@property (copy, readonly) PowerPointShape *background;
@property (copy, readonly) PowerPointColorScheme *colorScheme;
@property (copy, readonly) PowerPointDesign *design;
@property (copy, readonly) PowerPointHeadersAndFooters *headersAndFooters;
@property (readonly) double height;
@property (copy, readonly) PowerPointOfficeTheme *theme;
@property (copy, readonly) PowerPointTimeline *timeline;
@property (readonly) double width;

- (PowerPointTextStyle *) getTextStyleFromAt:(PowerPointPTst)at;

@end

@interface PowerPointMotionEffect : PowerPointBaseObject

@property double byX;
@property double byY;
@property double fromX;
@property double fromY;
@property (copy) NSString *path;
@property double toX;
@property double toY;


@end

@interface PowerPointNamedSlideShow : PowerPointBaseObject

@property (copy, readonly) NSString *name;
@property (readonly) NSInteger numberOfSlides;
@property (copy, readonly) NSArray *slideIDs;


@end

@interface PowerPointPageSetup : PowerPointBaseObject

@property NSInteger firstSlideNumber;
@property PowerPointMOrT notesOrientation;
@property PowerPointMOrT slideOrientation;
@property PowerPointSSzT slideSize;
@property double slideWidth;


@end

@interface PowerPointPane : PowerPointBaseObject

@property (readonly) BOOL active;
@property (readonly) PowerPointPVTy paneViewType;


@end

@interface PowerPointParagraphFormat : PowerPointBaseObject

- (SBElementArray<PowerPointTabStop *> *) tabStops;

@property PowerPointPpgA alignment;
@property PowerPointPBlA baselineAlignment;  // Returns or sets a constant that represents the vertical position of fonts in a paragraph.
@property (copy, readonly) PowerPointBulletFormat *bulletFormat;
@property BOOL eastAsianLineBreakControl;
@property double firstLineIndent;  // Returns or sets the value, in points, for a first line or hanging indent.
@property BOOL hangingPunctuation;  // Determines whether hanging punctuation is enabled for the specified paragraphs.
@property NSInteger indentLevel;  // Returns or sets a value representing the indent level assigned to text in the selected paragraph.
@property double leftIndent;  // Returns or sets a value that represents the left indent value, in points, for the specified paragraphs.
@property BOOL lineRuleAfter;  // Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines.
@property BOOL lineRuleBefore;  // Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines.
@property BOOL lineRuleWithin;  // Determines whether line spacing between base lines is set to a specific number of points or lines.
@property double rightIndent;  // Returns or sets the right indent, in points, for the specified paragraphs.
@property double spaceAfter;  // Returns or sets the spacing, in points, after the specified paragraphs.
@property double spaceBefore;  // Returns or sets the spacing, in points, before the specified paragraphs.
@property double spaceWithin;  // Returns or sets the amount of space between base lines in the specified paragraph, in points or lines.
@property PowerPointPDrT textDirection;  // Returns or sets the text direction for the specified paragraph.
@property BOOL wordWrap;  // Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs.


@end

@interface PowerPointPlaySettings : PowerPointBaseObject

@property BOOL hideWhileNotPlaying;
@property BOOL loopUntilStopped;
@property BOOL pauseAnimation;
@property BOOL playOnEntry;
@property BOOL rewindMove;
@property NSInteger stopAfterSlides;


@end

// Represents an interface for playing movies
@interface PowerPointPlayer : PowerPointBaseObject

@property NSInteger currentPosition;
@property (readonly) PowerPointPSta playerState;

- (void) goToNextBookmarkForPlayer;  // Advance the player to the next bookmark.
- (void) goToPreviousBookmarkForPlayer;  // Rewind the player to the previous bookmark.
- (void) pausePlayer;  // Pause playback.
- (void) playPlayer;  // Begin or resume playback.
- (void) stopPlayer;  // Stop playback.

@end

@interface PowerPointPreferences : PowerPointBaseObject

@property NSInteger alwaysSuggestCorrections;
@property NSInteger appendDOSExtentions;
@property NSInteger autoFit;
@property NSInteger autoRecoveryFrequency;
@property NSInteger backgroundSpelling;
@property NSInteger compressFile;
@property NSInteger defaultView;
@property NSInteger dragDrop;
@property NSInteger endBlankSlide;
@property NSInteger filePropertiesPrompt;
@property NSInteger fullTextSearch;
@property NSInteger hideSpellingErrors;
@property NSInteger ignoreNumbersInWords;
@property NSInteger ignoreUppercase;
@property NSInteger mruListActive;
@property NSInteger mruListSize;
@property NSInteger optionBitmap;
@property NSInteger rulerUnits;
@property NSInteger saveAutoRecovery;
@property NSInteger showVerticalRuler;
@property NSInteger slideShowControl;
@property NSInteger slideShowRightMouse;
@property NSInteger smartCutPaste;
@property NSInteger smartQuotes;
@property NSInteger undoLevelCount;
@property (copy) NSString *userInitials;
@property (copy) NSString *userName;
@property NSInteger wordSelection;


@end

@interface PowerPointPresentation : PowerPointBaseObject

- (SBElementArray<PowerPointSlide *> *) slides;
- (SBElementArray<PowerPointColorScheme *> *) colorSchemes;
- (SBElementArray<PowerPointFont *> *) fonts;
- (SBElementArray<PowerPointDocumentWindow *> *) documentWindows;
- (SBElementArray<PowerPointDocumentProperty *> *) documentProperties;
- (SBElementArray<PowerPointCustomDocumentProperty *> *) customDocumentProperties;
- (SBElementArray<PowerPointDesign *> *) designs;

@property (copy, readonly) PowerPointBroadcast *broadcast;
@property (copy, readonly) PowerPointShape *defaultShape;
@property PowerPointPeBl eastAsianLineBreakLevel;  // Returns or sets the East Asian line break control level for the specified paragraph.
@property (copy, readonly) NSString *fullName;
@property (copy, readonly) PowerPointMaster *handoutMaster;
@property (readonly) BOOL hasTitleMaster;
@property (readonly) BOOL inMergeMode;
@property (readonly) BOOL isProtected;  // Returns true if the presentation is protected by Information Rights Management.
@property PowerPointPDrT layoutDirection;
@property (copy, readonly) NSString *name;
@property (copy) NSString *noLineBreakAfter;
@property (copy) NSString *noLineBreakBefore;
@property (copy, readonly) PowerPointMaster *notesMaster;
@property (copy, readonly) PowerPointPageSetup *pageSetup;
@property (copy) NSString *password;  // The password for opening the presentation
@property (copy, readonly) NSString *path;
@property (copy, readonly) PowerPointPrintOptions *printOptions;
@property (readonly) BOOL readOnly;
@property (copy, readonly) PowerPointSaveAsMovieSettings *saveAsMovieSettings;
@property BOOL saved;
@property (copy, readonly) PowerPointSectionProperties *sectionProperties;
@property (copy, readonly) PowerPointMaster *slideMaster;
@property (copy, readonly) PowerPointSlideShowSettings *slideShowSettings;
@property (copy, readonly) PowerPointSlideShowWindow *slideShowWindow;
@property (copy, readonly) NSString *templateName;
@property (copy, readonly) PowerPointMaster *titleMaster;
@property (copy, readonly) PowerPointWebOptions *webOptions;
@property (copy) NSString *writePassword;  // The password for saving changes to the presentation

- (void) acceptAll;  // Accept all of the Merge Compare differences between two Presentations.
- (PowerPointDesign *) addDesignDesignName:(NSString *)DesignName index:(NSInteger)index;  // add a new design
- (void) applyTemplateFileName:(NSString *)fileName;  // Applies a theme or design template to the specified slide, master or presentation
- (BOOL) canCheckIn;  // Returns True if PowerPoint can check in a specified presentation to a server.
- (void) checkInSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic;  // Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) checkInWithVersionSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic versionType:(PowerPointCivt)versionType;  // Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) endReview;  // End the Merge Compare Process between two Presentations
- (void) mergePresentationWithRevisionPath:(NSString *)withRevisionPath withBaselinePath:(id)withBaselinePath;  // Compare the active presentation between one or two presentations. If two files are to be compared to the active presentation, then the optional second parameter is the Path to the baseline.
- (void) mergeWithBaselineWithRevisionPath:(NSString *)withRevisionPath withBaselinePath:(NSString *)withBaselinePath;  // Compare the active presentation with two presentations. The two files are to be compared to the active presentation, then the second parameter is the Path to the baseline.
- (void) printOutFrom:(NSInteger)from to:(NSInteger)to printToFile:(NSString *)printToFile copies:(NSInteger)copies collate:(BOOL)collate showDialog:(BOOL)showDialog;
- (void) redoTimes:(NSInteger)times;
- (void) rejectAll;  // Reject all of the Merge Compare differences between two Presentations.
- (void) undoTimes:(NSInteger)times;
- (void) updateLinks;

@end

@interface PowerPointPresenterTool : PowerPointBaseObject

@property (copy, readonly) PowerPointSlide *currentPresenterSlide;
@property (readonly) NSInteger currentSlideInShow;
@property (readonly) double elapsedTime;
@property (readonly) NSInteger maxSlidesInShow;
@property (copy) NSString *notesText;
@property NSInteger notesZoom;
@property (readonly) BOOL slideMiniature;

- (void) endShow;
- (void) next;
- (void) pauseTimer;
- (void) previous;
- (void) resetTimer;
- (void) startTimer;
- (void) swapDisplays;
- (void) toggleSlideMiniature;

@end

@interface PowerPointPresenterViewWindow : PowerPointBasicWindow

@property (readonly) BOOL active;
@property (readonly) double height;
@property (copy, readonly) PowerPointPresentation *presentation;
@property (copy, readonly) PowerPointPresenterTool *presenterTool;
@property (readonly) double width;


@end

@interface PowerPointPrintOptions : PowerPointBaseObject

- (SBElementArray<PowerPointPrintRange *> *) printRanges;

@property (copy, readonly) NSString *activePrinter;
@property BOOL collate;
@property BOOL fitToPage;
@property BOOL frameSlides;
@property NSInteger numberOfCopies;
@property PowerPointPtOt outputType;
@property PowerPointPrCt printColorType;
@property BOOL printFontsAsGraphics;
@property BOOL printHiddenSlides;
@property PowerPointRgTy rangeType;
@property NSInteger sectionIndex;
@property (copy) NSString *slideShowName;


@end

@interface PowerPointPrintRange : PowerPointBaseObject

@property (readonly) NSInteger rangeEnd;
@property (readonly) NSInteger rangeStart;


@end

@interface PowerPointPropertyEffect : PowerPointBaseObject

- (SBElementArray<PowerPointAnimationPoint *> *) animationPoints;

@property (copy, readonly) id ending;
@property PowerPointAnPr propertySetEffect;
@property (copy, readonly) id starting;


@end

@interface PowerPointRotatingEffect : PowerPointBaseObject

@property double rotating;


@end

@interface PowerPointRulerLevel : PowerPointBaseObject

@property double firstMargin;  // Returns or sets the first-line indent for the specified outline level, in points.
@property double leftMargin;  // Returns or sets the left indent for the specified outline level, in points.


@end

// Represents the ruler for the text in the specified shape or for all text in the specified text style. Contains tab stops and the indentation settings for text outline levels.
@interface PowerPointRuler : PowerPointBaseObject

- (SBElementArray<PowerPointTabStop *> *) tabStops;
- (SBElementArray<PowerPointRulerLevel *> *) rulerLevels;


@end

@interface PowerPointSaveAsMovieSettings : PowerPointBaseObject

@property BOOL autoLoopEnabled;
@property (copy) NSString *backgroundSoundTrackFile;
@property NSInteger backgroundTrackSegmentEnd;
@property NSInteger backgroundTrackSegmentStart;
@property NSInteger backgroundTrackStart;
@property BOOL createMoviePreview;
@property BOOL forceAllInline;
@property BOOL includeNarrationAndSounds;
@property (copy) NSString *movieActors;
@property (copy) NSString *movieAuthor;
@property (copy) NSString *movieCopyright;
@property NSInteger movieFrameHeight;
@property NSInteger movieFrameWidth;
@property (copy) NSString *movieProducer;
@property PowerPointPMOt optimization;
@property BOOL showMovieController;
@property (copy) NSString *transitionDescription;
@property BOOL useSingleTransition;


@end

@interface PowerPointScaleEffect : PowerPointBaseObject

@property double byX;
@property double byY;
@property double fromX;
@property double fromY;
@property double toX;
@property double toY;


@end

@interface PowerPointSectionProperties : PowerPointBaseObject

- (void) deleteSectionAtPosition:(NSInteger)atPosition deletingSlides:(BOOL)deletingSlides;
- (NSInteger) getCountOfSections;
- (NSInteger) getFirstSlideOfSectionAtPosition:(NSInteger)atPosition;
- (NSString *) getIdOfSectionAtPosition:(NSInteger)atPosition;
- (NSString *) getNameOfSectionAtPosition:(NSInteger)atPosition;
- (NSInteger) getSlideCountOfSectionAtPosition:(NSInteger)atPosition;
- (NSInteger) insertSectionBeforeSection:(NSInteger)beforeSection beforeSlide:(NSInteger)beforeSlide titled:(NSString *)titled;
- (void) moveSectionAtPosition:(NSInteger)atPosition toPosition:(NSInteger)toPosition;
- (void) renameSectionAtPosition:(NSInteger)atPosition to:(NSString *)to;

@end

// Represents the selection in the specified document window
@interface PowerPointSelection : PowerPointBaseObject

@property (copy, readonly) PowerPointShapeRange *childShapeRange;  // Returns the child shapes of a selection.
@property (readonly) BOOL hasChildShapeRange;  // Returns whether the selection contains child shapes.
@property (readonly) PowerPointPSEL selectionType;  // Represents the type of objects in a selection.
@property (copy, readonly) PowerPointShapeRange *shapeRange;  // Returns a collection of shapes selected on the specified slide.
@property (copy, readonly) PowerPointSlideRange *slideRange;  // Returns a collection of selected slides.
@property (copy, readonly) PowerPointTextRange *textRange;  // Returns the text and text properties of the selected text.

- (void) unselect;  // Cancels the current selection.

@end

@interface PowerPointSequence : PowerPointBaseObject

- (SBElementArray<PowerPointEffect *> *) effects;

- (PowerPointEffect *) addEffectFor:(PowerPointShape *)for_ fx:(PowerPointAnFX)fx level:(PowerPointAnLv)level trigger:(PowerPointAnTr)trigger index:(NSInteger)index;  // add an effect for a shape
- (PowerPointEffect *) convertToTextUnitEffectEffect:(PowerPointEffect *)Effect unit:(PowerPointAnTU)unit;  // convert an effect to a text unit effect

@end

@interface PowerPointSetEffect : PowerPointBaseObject

@property (copy) id ending;
@property PowerPointAnPr propertySetEffect;


@end

// A collection that represents a notes page or a slide range, which is a set of slides that can contain a single slide to as many as all slides in a presentation.
@interface PowerPointSlideRange : PowerPointBaseObject

- (SBElementArray<PowerPointSlide *> *) slides;
- (SBElementArray<PowerPointShape *> *) shapes;
- (SBElementArray<PowerPointHyperlink *> *) hyperlinks;

@property (copy) PowerPointColorScheme *colorScheme;  // Returns or sets the color scheme for the specified slide, slide range, or slide master.
@property (copy) PowerPointCustomLayout *customLayout;  // Returns the custom layout associated with the specified range of slides.
@property (copy) PowerPointDesign *design;
@property BOOL displayMasterShapes;  // Determines whether the specified range of slides displays the background objects on the slide master.
@property BOOL followMasterBackground;  // Determines whether the range of slides follows the slide master background.
@property (copy, readonly) PowerPointHeadersAndFooters *headersAndFooters;  // Returns a collection that represents the header, footer, date and time, and slide number associated with the slide, slide master, or range of slides.
@property PowerPointPSLo layout;  // Returns or sets the slide layout.
@property (copy, readonly) PowerPointSlideRange *notesPage;  // Returns a slide range object that represents the notes pages for the specified slide or range of slides.
@property (readonly) NSInteger printSteps;
@property (readonly) NSInteger slideID;
@property (readonly) NSInteger slideIndex;
@property (copy, readonly) PowerPointMaster *slideMaster;  // Returns the slide master.
@property (readonly) NSInteger slideNumber;  // Returns the slide number.
@property (copy, readonly) PowerPointSlideShowTransition *slideShowTransitions;  // Returns the special effects for the specified slide transition.
@property (copy, readonly) PowerPointTimeline *timeline;  // Returns the animation timeline for the slide.


@end

@interface PowerPointSlideShowSettings : PowerPointBaseObject

- (SBElementArray<PowerPointNamedSlideShow *> *) namedSlideShows;

@property PowerPointPSaM advanceMode;
@property NSInteger endingSlide;
@property BOOL loopUntilStopped;
@property (copy) NSColor *pointerColor;
@property PowerPointMCoI pointerColorThemeIndex;  // Returns or sets the color of  default pointer color for a presentation.
@property PowerPointSRgT rangeType;
@property PowerPointPSSt showType;
@property (readonly) BOOL showWithAnimation;
@property BOOL showWithNarration;
@property BOOL showWithPresenter;
@property (copy) NSString *slideShowName;
@property NSInteger startingSlide;

- (PowerPointSlideShowWindow *) runSlideShow;

@end

@interface PowerPointSlideShowTransition : PowerPointBaseObject

@property BOOL advanceOnClick;
@property BOOL advanceOnTime;
@property double advanceTime;
@property PowerPointPEeF entryEffect;
@property BOOL hidden;
@property BOOL loopSoundUntilNext;
@property (copy, readonly) PowerPointSoundEffect *soundEffectTransition;
@property double transitionDuration;


@end

@interface PowerPointSlideShowView : PowerPointBaseObject

@property BOOL accelerationsEnabled;
@property (readonly) NSInteger currentShowPosition;
@property (readonly) BOOL isNamedShow;
@property (copy, readonly) PowerPointSlide *lastSlideViewed;
@property (copy) NSColor *pointerColor;
@property PowerPointMCoI pointerColorThemeIndex;  // Returns or sets the color of temporary pointer color for a view of a slide show.
@property PowerPointPSsP pointerType;
@property (readonly) double presentationElapsedTime;
@property (copy, readonly) PowerPointSlide *slide;
@property double slideElapsedTime;
@property (copy, readonly) NSString *slideShowName;
@property PowerPointPShS slideState;
@property (readonly) NSInteger zoom;

- (void) exitSlideShow;
- (void) goToFirstSlide;
- (void) goToLastSlide;
- (void) goToNextSlide;
- (void) goToPreviousSlide;
- (void) resetSlideTime;

@end

@interface PowerPointSlideShowWindow : PowerPointBasicWindow

@property (readonly) BOOL active;
@property NSRect bounds;
@property double height;
@property (readonly) BOOL isFullScreen;
@property double leftPosition;
@property (copy, readonly) PowerPointPresentation *presentation;
@property (copy, readonly) PowerPointSlideShowView *slideshowView;
@property double top;
@property double width;


@end

@interface PowerPointSlide : PowerPointBaseObject

- (SBElementArray<PowerPointShape *> *) shapes;
- (SBElementArray<PowerPointHyperlink *> *) hyperlinks;

@property (copy, readonly) PowerPointShape *background;
@property (copy) PowerPointColorScheme *colorScheme;
@property (copy) PowerPointCustomLayout *customLayout;
@property (copy) PowerPointDesign *design;
@property BOOL displayMasterShapes;
@property BOOL followMasterBackground;
@property (copy, readonly) PowerPointHeadersAndFooters *headersAndFooters;
@property PowerPointPSLo layout;
@property (copy, readonly) PowerPointSlide *notesPage;
@property (readonly) NSInteger printSteps;
@property (readonly) NSInteger sectionIndex;
@property (readonly) NSInteger sectionNumber;
@property (readonly) NSInteger slideID;
@property (readonly) NSInteger slideIndex;
@property (copy, readonly) PowerPointMaster *slideMaster;
@property (readonly) NSInteger slideNumber;
@property (copy, readonly) PowerPointSlideShowTransition *slideShowTransition;
@property (copy, readonly) PowerPointTimeline *timeline;

- (void) applyThemeColorSchemeFileName:(NSString *)fileName;  // Applies a theme color to the specified slide.
- (void) copyObject NS_RETURNS_NOT_RETAINED;
- (void) cutObject;
- (void) moveToStartOfSectionAtPosition:(NSInteger)atPosition;

@end

@interface PowerPointSoundEffect : PowerPointBaseObject

@property (copy, readonly) NSString *name;
@property PowerPointPSnX soundType;

- (void) importSoundFileSoundFileName:(NSString *)soundFileName;
- (void) playSoundEffect;

@end

// Represents a single tab stop.
@interface PowerPointTabStop : PowerPointBaseObject

@property double tabPosition;  // Returns or sets the position of a tab stop relative to the left margin.
@property PowerPointPTSp tabStopType;  // Returns or sets the type of the tab stop object.


@end

@interface PowerPointTextStyleLevel : PowerPointBaseObject

@property (copy, readonly) PowerPointFont *font;
@property (copy, readonly) PowerPointParagraphFormat *paragraphFormat;


@end

@interface PowerPointTextStyle : PowerPointBaseObject

- (SBElementArray<PowerPointTextStyleLevel *> *) textStyleLevels;

@property (copy, readonly) PowerPointRuler *ruler;
@property (copy, readonly) PowerPointTextFrame *textFrame;


@end

@interface PowerPointTimeline : PowerPointBaseObject

- (SBElementArray<PowerPointSequence *> *) sequences;

@property (copy, readonly) PowerPointSequence *mainSequence;

- (PowerPointSequence *) addSequenceIndex:(NSInteger)index;  // add an interactive sequence

@end

@interface PowerPointTiming : PowerPointBaseObject

@property double acceleration;
@property BOOL autoreverse;
@property double deceleration;
@property double duration;
@property NSInteger repeatCount;
@property double repeatDuration;
@property PowerPointAnRs restart;
@property BOOL rewind;
@property BOOL smoothEnd;
@property BOOL smoothStart;
@property double speed;


@end

// Represents a single initial-capital autocorrect exception.
@interface PowerPointTwoInitialCapsException : PowerPointBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of the TwoInitialCapsException.


@end

@interface PowerPointView : PowerPointBaseObject

@property BOOL displaySlideMiniature;
@property (copy) PowerPointSlide *slide;
@property (readonly) PowerPointPVTy viewType;
@property NSInteger zoom;
@property BOOL zoomToFit;

- (void) goToSlideNumber:(NSInteger)number;

@end

@interface PowerPointWebOptions : PowerPointBaseObject

@property BOOL allowPNG;
@property PowerPointMtEn encoding;


@end



/*
 * Drawing Suite
 */

@interface PowerPointAdjustment : PowerPointBaseObject

@property double adjustment_value;


@end

@interface PowerPointCalloutFormat : PowerPointBaseObject

@property BOOL accent;
@property PowerPointMCAt angle;
@property BOOL autoAttach;
@property (readonly) BOOL autoLength;
@property BOOL border;
@property (readonly) double calloutFormatLength;
@property PowerPointMCot calloutType;
@property (readonly) double drop;
@property (readonly) PowerPointMCDt dropType;
@property double gap;

- (void) presetDropDropType:(PowerPointMCDt)dropType;

@end

@interface PowerPointConnectorFormat : PowerPointBaseObject

@property (readonly) BOOL beginConnected;
@property (copy, readonly) PowerPointShape *beginConnectedShape;
@property (readonly) NSInteger beginConnectionSite;
@property PowerPointMCtT connectorType;
@property (readonly) BOOL endConnected;
@property (copy, readonly) PowerPointShape *endConnectedShape;
@property (readonly) NSInteger endConnectionSite;


@end

// Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.
@interface PowerPointFillFormat : PowerPointBaseObject

- (SBElementArray<PowerPointGradientStop *> *) gradientStops;

@property (copy) NSColor *backColor;
@property PowerPointMCoI backColorThemeIndex;  // Returns or sets the specified fill background color.
@property (readonly) PowerPointMFdT fillFormatType;
@property (copy) NSColor *foreColor;
@property PowerPointMCoI foreColorThemeIndex;  // Returns or sets the specified foreground fill or solid color.
@property (readonly) PowerPointMGCt gradientColorType;
@property (readonly) double gradientDegree;
@property (readonly) PowerPointMGdS gradientStyle;
@property (readonly) NSInteger gradientVariant;
@property (readonly) PowerPointPpTy pattern;
@property (readonly) PowerPointMPGb presetGradientType;
@property (readonly) PowerPointMPzT presetTexture;
@property BOOL rotateWithObject;  // Returns or sets whether the fill rotates with the specified shape.
@property PowerPointMTtA textureAlignment;  // Returns or sets the texture alignment for the specified object.
@property double textureHorizontalScale;  // Returns or sets the texture alignment for the specified object.
@property (copy, readonly) NSString *textureName;
@property double textureOffsetX;  // Returns or sets the texture alignment for the specified object.
@property double textureOffsetY;  // Returns or sets the texture alignment for the specified object.
@property BOOL textureTile;  // Returns the texture tile style for the specified fill.
@property double textureVerticalScale;  // Returns or sets the texture alignment for the specified object.
@property double transparency;
@property BOOL visible;

- (void) deleteGradientStopStopIndex:(NSInteger)stopIndex;  // Removes a gradient stop.
- (void) insertGradientStopCustomColor:(NSColor *)customColor position:(double)position transparency:(double)transparency stopIndex:(NSInteger)stopIndex;  // Adds a stop to a gradient.

@end

// Represents the glow formatting for a shape or range of shapes
@interface PowerPointGlowFormat : PowerPointBaseObject

@property (copy) NSColor *color;  // Returns or sets the color for the specified glow format.
@property PowerPointMCoI colorThemeIndex;  // Returns or sets the color for the specified glow format.
@property double radius;  // Returns or sets the length of the radius for the specified glow format.


@end

// Represents one gradient stop.
@interface PowerPointGradientStop : PowerPointBaseObject

@property (copy) NSColor *color;  // Returns or sets the color for the specified the gradient stop.
@property PowerPointMCoI colorThemeIndex;  // Returns or sets the color for the specified gradient stop.
@property double position;  // Returns or sets the position for the specified gradient stop expressed as a percent.
@property double transparency;  // Returns or sets a value representing the transparency of the gradient fill expressed as a percent.


@end

@interface PowerPointLineFormat : PowerPointBaseObject

@property (copy) NSColor *backColor;
@property PowerPointMCoI backColorThemeIndex;  // Returns or sets the background color for a patterned line.
@property PowerPointMAhL beginArrowHeadLength;
@property PowerPointMAhS beginArrowheadStyle;
@property PowerPointMAhW beginArrowheadWidth;
@property PowerPointMlDs dashStyle;
@property PowerPointMAhL endArrowheadLength;
@property PowerPointMAhS endArrowheadStyle;
@property PowerPointMAhW endArrowheadWidth;
@property (copy) NSColor *foreColor;
@property PowerPointMCoI foreColorThemeIndex;  // Returns or sets the foreground color for the line.
@property PowerPointPpTy lineFormatPatterned;
@property PowerPointMLnS lineStyle;
@property double lineWeight;
@property double transparency;


@end

@interface PowerPointLinkFormat : PowerPointBaseObject

@property PowerPointPUdO autoUpdate;
@property (copy) NSString *sourceFullName;


@end

// Represents a Microsoft Office theme.
@interface PowerPointOfficeTheme : PowerPointBaseObject

@property (copy, readonly) PowerPointThemeColorScheme *themeColorScheme;  // Returns the color scheme of a Microsoft Office theme.
@property (copy, readonly) PowerPointThemeEffectScheme *themeEffectScheme;  // Returns the effects scheme of a Microsoft Office theme.
@property (copy, readonly) PowerPointThemeFontScheme *themeFontScheme;  // Returns the font scheme of a Microsoft Office theme.


@end

@interface PowerPointPictureFormat : PowerPointBaseObject

@property double brightness;  // Returns or sets the brightness of the specified picture. The value for this property must be a number from 0.0, dimmest to 1.0, brightest.
@property PowerPointMPc colorType;  // Returns or sets the type of color transformation applied to the specified picture.
@property double contrast;  // Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast to 1.0, the greatest contrast.
@property double cropBottom;  // Returns or sets the number of points that are cropped off the bottom of the specified picture.
@property double cropLeft;  // Returns or sets the number of points that are cropped off the left side of the specified picture.
@property double cropRight;  // Returns or sets the number of points that are cropped off the right side of the specified picture.
@property double cropTop;  // Returns or sets the number of points that are cropped off the top of the specified picture.
@property (copy) NSColor *transparencyColor;  // Returns or sets the transparent color for the specified picture as aRGB color. For this property to take effect, the transparent background property must be set to true.
@property BOOL transparentBackground;  // Returns or sets if the parts of the picture that are defined with a transparent color actually appear transparent.


@end

@interface PowerPointPlaceholderFormat : PowerPointBaseObject

@property (readonly) PowerPointMShp containedType;
@property (copy) NSString *placeholderName;
@property (readonly) PowerPointPPhd placeholderType;


@end

// Represents the reflection effect in Office graphics.
@interface PowerPointReflectionFormat : PowerPointBaseObject

@property PowerPointMRfT reflectionType;  // Returns or sets the type of the reflection format object.


@end

@interface PowerPointShadowFormat : PowerPointBaseObject

@property double blur;
@property (copy) NSColor *foreColor;
@property PowerPointMCoI foreColorThemeIndex;  // Returns or sets the foreground color for the shadow format.
@property BOOL obscured;
@property double offsetX;
@property double offsetY;
@property BOOL rotateWithShape;  // Returns or sets whether to rotate the shadow when rotating the shape.
@property PowerPointMSSt shadowStyle;  // Returns or sets the style of shadow formatting to apply to a shape.
@property PowerPointMSdT shadowType;
@property double size;  // Returns or sets the width of the shadow.
@property double transparency;
@property BOOL visible;


@end

// A collection that represents a set of shapes on a document.
@interface PowerPointShapeRange : PowerPointBaseObject

- (SBElementArray<PowerPointAdjustment *> *) adjustments;
- (SBElementArray<PowerPointShape *> *) shapes;

@property (copy, readonly) PowerPointAnimationSettings *animationSettings;  // Returns all the special effects that you can apply to the animation of the specified shape.
@property PowerPointMAsT autoShapeType;  // Returns or sets the shape type for AutoShapes other than a line, freeform drawing, or connector.
@property PowerPointMBSI backgroundStyle;  // Returns or sets the background style for the specified object.
@property PowerPointMBW blackAndWhiteMode;  // Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode.
@property (copy, readonly) PowerPointCalloutFormat *calloutFormat;  // Returns callout formatting properties for the specified line callout shapes.
@property (readonly) BOOL child;  // Returns whether all shapes in a shape range are child shapes of the same parent.
@property (readonly) NSInteger connectionSiteCount;  // Returns the number of connection sites on the specified shape.
@property (copy, readonly) PowerPointFillFormat *fillFormat;  // Returns the fill format properties for the specified object.
@property (copy, readonly) PowerPointGlowFormat *glowFormat;  // Returns the glow format for the specified range of shapes.
@property (readonly) BOOL hasChild;
@property (readonly) BOOL hasTable;  // Returns whether the specified shape is a table.
@property (readonly) BOOL hasTextFrame;  // Returns whether the specified shape has a text frame.
@property double height;  // Returns or sets the height of the specified object.
@property (readonly) BOOL horizontalFlip;  // Returns whether the specified shape is flipped around the horizontal axis.
@property (readonly) BOOL isConnector;  // Determines whether the specified shape is a connector.
@property double leftPosition;  // Returns or sets a value equal to the distance from the left edge of the leftmost shape in the shape range to the left edge of the slide.
@property (copy, readonly) PowerPointLineFormat *lineFormat;  // Returns line format properties for the specified line or shape border.
@property (copy, readonly) PowerPointLinkFormat *linkFormat;  // Returns the properties for linked OLE objects.
@property BOOL lockAspectRatio;  // Determines whether the specified shape retains its original proportions when you resize it.
@property (readonly) PowerPointMedT mediaType;  // Returns the OLE media type.
@property (copy) NSString *name;  // Identifies the shape on a slide.
@property (copy, readonly) PowerPointReflectionFormat *reflectionFormat;  // Returns the reflection format for the specified range of shapes.
@property double rotation;  // Returns or sets the number of degrees that the specified shape is rotated around the z-axis.
@property (copy, readonly) PowerPointShadowFormat *shadowFormat;  // Returns shadow formatting for specified shapes.
@property PowerPointMSSI shapeStyle;  // Returns or sets the shape style index for the specified object.
@property (readonly) PowerPointMShp shapeType;  // Represents the type of shape or shapes in a range of shapes.
@property (copy, readonly) PowerPointSoftEdgeFormat *softEdgeFormat;  // Returns the soft edge format for the specified range of shapes.
@property (copy, readonly) PowerPointTextFrame *textFrame;  // Returns the alignment and anchoring properties for the specified shape or master text style.
@property (copy, readonly) PowerPointThreeDFormat *threeDFormat;  // Returns 3-D effect formatting properties for the specified shape.
@property double top;  // Returns or sets a value equal to the distance from the top edge of the topmost shape in the shape range to the top edge of the document.
@property (readonly) BOOL verticalFlip;  // Determines whether the specified shape is flipped around the vertical axis.
@property BOOL visible;  // Returns or sets the visibility of the specified object or the formatting applied to the specified object.
@property double width;  // Returns or sets the width of the specified object.
@property (readonly) NSInteger zOrderPosition;  // Returns the position of the specified shape in the z-order.

- (void) alignAlignType:(PowerPointMAlC)alignType relativeToSlide:(BOOL)relativeToSlide;  // Aligns the shapes in the specified range of shapes.
- (void) copyShapeRange NS_RETURNS_NOT_RETAINED;
- (void) cutShapeRange;
- (void) distributeDistributeType:(PowerPointMDsC)distributeType relativeToSlide:(BOOL)relativeToSlide;  // Evenly distributes the shapes in the specified range of shapes. You can specify whether you want to distribute the shapes horizontally or vertically and whether you want to distribute them over the entire slide or over the space they originally occupy.
- (PowerPointShape *) group;  // Groups the shapes in the specified shape range.
- (PowerPointShape *) regroup;  // Regroups the group that the specified shape range belonged to previously.
- (PowerPointShapeRange *) ungroup;  // Ungroups any grouped shapes in the specified shape or range of shapes.

@end

@interface PowerPointShape : PowerPointBaseObject

- (SBElementArray<PowerPointShape *> *) shapes;
- (SBElementArray<PowerPointCallout *> *) callouts;
- (SBElementArray<PowerPointConnector *> *) connectors;
- (SBElementArray<PowerPointPicture *> *) pictures;
- (SBElementArray<PowerPointLineShape *> *) lineShapes;
- (SBElementArray<PowerPointPlaceHolder *> *) placeHolders;
- (SBElementArray<PowerPointTextBox *> *) textBoxes;
- (SBElementArray<PowerPointComment *> *) comments;
- (SBElementArray<PowerPointShapeTable *> *) shapeTables;
- (SBElementArray<PowerPointAdjustment *> *) adjustments;

@property (copy, readonly) PowerPointAnimationSettings *animationSettings;
@property PowerPointMAsT autoShapeType;
@property PowerPointMBSI backgroundStyle;  // Returns or sets the background style.
@property PowerPointMBW blackAndWhiteMode;
@property (readonly) BOOL child;  // True if the shape is a child shape.
@property (readonly) NSInteger connectionSiteCount;
@property (copy, readonly) PowerPointFillFormat *fillFormat;  // Returns a fill format object that contains fill formatting properties for the specified shape.
@property (copy, readonly) PowerPointGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property (readonly) BOOL hasTable;
@property (readonly) BOOL hasTextFrame;
@property double height;
@property (readonly) BOOL horizontalFlip;
@property (readonly) BOOL isConnector;
@property double leftPosition;
@property (copy, readonly) PowerPointLineFormat *lineFormat;
@property (copy, readonly) PowerPointLinkFormat *linkFormat;
@property BOOL lockAspectRatio;
@property (readonly) PowerPointMedT mediaType;
@property (copy) NSString *name;
@property (copy, readonly) PowerPointReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property double rotation;
@property (copy, readonly) PowerPointShadowFormat *shadowFormat;
@property PowerPointMSSI shapeStyle;  // Returns or sets the shape style corresponding to the Quick Styles.
@property (readonly) PowerPointMShp shapeType;
@property (copy, readonly) PowerPointSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) PowerPointTextFrame *textFrame;
@property (copy, readonly) PowerPointThreeDFormat *threeDFormat;  // Returns a threeD format object that contains 3-D-effect formatting properties for the specified shape.
@property double top;
@property (readonly) BOOL verticalFlip;
@property BOOL visible;
@property double width;
@property (copy, readonly) PowerPointWordArtFormat *wordArtFormat;  // Returns the formatting properties for a word art effect.
@property (readonly) NSInteger zOrderPosition;

- (void) copyShape NS_RETURNS_NOT_RETAINED;
- (void) cutShape;
- (void) saveAsPicturePictureType:(PowerPointMPiT)pictureType fileName:(NSString *)fileName;  // Saves the shape in the requested file using the stated graphic format

@end

@interface PowerPointCallout : PowerPointShape

@property (readonly) PowerPointMCot calloutType;
@property (copy, readonly) PowerPointCalloutFormat *calloutFormat;


@end

@interface PowerPointComment : PowerPointShape


@end

@interface PowerPointConnector : PowerPointShape

@property (copy, readonly) PowerPointConnectorFormat *connectorFormat;
@property (readonly) PowerPointMCtT connectorType;


@end

// The line shape uses begin line X, begin line Y, end line X, and end line Y when created
@interface PowerPointLineShape : PowerPointShape

@property double beginLineX;
@property double beginLineY;
@property double endLineX;
@property double endLineY;


@end

@interface PowerPointMediaObject : PowerPointShape

@property (copy, readonly) NSString *fileName;


@end

@interface PowerPointMedia2Object : PowerPointShape

@property (copy, readonly) NSString *fileName;
@property (readonly) BOOL linkToFile;
@property (readonly) BOOL saveWithDocument;


@end

@interface PowerPointPicture : PowerPointShape

@property (copy, readonly) NSString *fileName;
@property (readonly) BOOL linkToFile;
@property (copy, readonly) PowerPointPictureFormat *pictureFormat;
@property (readonly) BOOL saveWithDocument;

- (void) scaleHeightFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(PowerPointMSFr)scale;
- (void) scaleWidthFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(PowerPointMSFr)scale;

@end

@interface PowerPointPlaceHolder : PowerPointShape

@property (copy, readonly) PowerPointPlaceholderFormat *placeHolderFormat;
@property (readonly) PowerPointPPhd placeholderType;


@end

@interface PowerPointShapeTable : PowerPointShape

@property (readonly) NSInteger numberOfColumns;
@property (readonly) NSInteger numberOfRows;
@property (copy, readonly) PowerPointTable *tableObject;


@end

// Represents the soft edge formatting for a shape or range of shapes
@interface PowerPointSoftEdgeFormat : PowerPointBaseObject

@property PowerPointMSET softEdgeType;  // Returns or sets the type soft edge format object.


@end

@interface PowerPointTextBox : PowerPointShape

@property (readonly) PowerPointTxOr textOrientation;


@end

// Represents a single text column.
@interface PowerPointTextColumn : PowerPointBaseObject

@property NSInteger columnNumber;  // Returns or sets the index of the text column object.
@property double spacing;  // Returns or sets the spacing between text columns in a text column object.
@property PowerPointPDrT textDirection;  // Returns or sets the direction of text in the text column object.


@end

@interface PowerPointTextFrame : PowerPointBaseObject

@property PowerPointPAtS autoSize;
@property (readonly) BOOL hasText;  // Returns whether the specified text frame has text.
@property PowerPointMHzA horizontalAnchor;
@property double marginBottom;
@property double marginLeft;
@property double marginRight;
@property double marginTop;
@property (readonly) PowerPointTxOr orientation;
@property PowerPointMPFo pathFormat;  // Returns or sets the path type for the specified text frame.
@property (copy, readonly) PowerPointRuler *ruler;
@property (copy, readonly) PowerPointTextColumn *textColumn;  // Returns the text column object that represents the columns within the text frame.
@property PowerPointTxOr textOrientation;
@property (copy, readonly) PowerPointTextRange *textRange;
@property (copy, readonly) PowerPointThreeDFormat *threeDFormat;  // Returns the 3-D effect formatting properties for the specified text.
@property PowerPointMVtA verticalAnchor;
@property PowerPointMWFo warpFormat;  // Returns or sets the warp type for the specified text frame.
@property (readonly) PowerPointMPXF wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.
@property BOOL wordWrap;  // Returns or sets text break lines within or past the boundaries of the shape.
@property PowerPointMPXF wordartFormat;  // Returns or sets the WordArt type for the specified text frame.


@end

// Represents the color scheme of an Office Theme
@interface PowerPointThemeColorScheme : PowerPointBaseObject

- (SBElementArray<PowerPointThemeColor *> *) themeColors;

- (NSColor *) getCustomColorName:(NSString *)name;  // Returns the custom color for the specified Microsoft Office theme.
- (void) loadThemeColorSchemeFileName:(NSString *)fileName;  // Loads the color scheme of a Microsoft Office theme from a file.
- (void) saveThemeColorSchemeFileName:(NSString *)fileName;  // Saves the color scheme of a Microsoft Office theme to a file.

@end

// Represents a color in the color scheme of a Microsoft Office 2007 theme.
@interface PowerPointThemeColor : PowerPointBaseObject

@property (copy) NSColor *RGB;  // Returns or sets a value of a color in the color scheme of a Microsoft Office theme.
@property (readonly) PowerPointMCSI themeColorSchemeIndex;  // Returns the index value a color scheme of a Microsoft Office theme.


@end

// Represents the effect scheme of a Microsoft Office theme.
@interface PowerPointThemeEffectScheme : PowerPointBaseObject

- (void) loadThemeEffectSchemeFileName:(NSString *)fileName;  // Loads the effects scheme of a Microsoft Office theme from a file.

@end

// Represents the font scheme of a Microsoft Office theme.
@interface PowerPointThemeFontScheme : PowerPointBaseObject

- (SBElementArray<PowerPointMinorThemeFont *> *) minorThemeFonts;
- (SBElementArray<PowerPointMajorThemeFont *> *) majorThemeFonts;

- (void) loadThemeFontSchemeFileName:(NSString *)fileName;  // Loads the font scheme of a Microsoft Office theme from a file.
- (void) saveThemeFontSchemeFileName:(NSString *)fileName;  // Saves the font scheme of a Microsoft Office theme to a file.

@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface PowerPointThemeFont : PowerPointBaseObject

@property (copy) NSString *name;  // Returns or sets a value specifying the font to use for a selection.


@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface PowerPointMajorThemeFont : PowerPointThemeFont


@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface PowerPointMinorThemeFont : PowerPointThemeFont


@end

// Represents a shape's three-dimensional formatting.
@interface PowerPointThreeDFormat : PowerPointBaseObject

@property double ZDistance;  // Returns or sets a Single that represents the distance from the center of an object or text.
@property double bevelBottomDepth;  // Returns or sets the depth/height of the bottom bevel.
@property double bevelBottomInset;  // Returns or sets the inset size/width for the bottom bevel.
@property PowerPointMBlT bevelBottomType;  // Returns or sets the bevel type for the bottom bevel.
@property double bevelTopDepth;  // Returns or sets the depth/height of the top bevel.
@property double bevelTopInset;  // Returns or sets the inset size/width for the top bevel.
@property PowerPointMBlT bevelTopType;  // Returns or sets the bevel type for the top bevel.
@property (copy, readonly) NSColor *contourColor;  // Returns or sets the color of the contour of an object or text.
@property PowerPointMCoI contourColorThemeIndex;  // Returns or sets the color for the specified contour.
@property double contourWidth;  // Returns or sets the width of the contour of an object or text.
@property double depth;  // Returns or sets the depth of the shape's extrusion.
@property (copy, readonly) NSColor *extrusionColor;  // Returns or sets a RGB color that represents the color of the shape's extrusion.
@property PowerPointMCoI extrusionColorThemeIndex;  // Returns or sets the color for the specified extrusion.
@property PowerPointMExC extrusionColorType;  // Returns or sets a value that indicates what will determine the extrusion color.
@property double fieldOfView;  // Returns or sets the amount of perspective for an object or text.
@property double lightAngle;  // Returns or sets the angle of the lighting.
@property BOOL perspective;  // Returns or sets if the extrusion appears in perspective that is, if the walls of the extrusion narrow toward a vanishing point. If false, the extrusion is a parallel, or orthographic, projection that is, if the walls don't narrow toward a vanishing point.
@property (readonly) PowerPointMPzC presetCamera;  // Returns a constant that represents the camera preset.
@property (readonly) PowerPointMExD presetExtrusionDirection;  // Returns or sets the direction taken by the extrusion's sweep path leading away from the extruded shape, the front face of the extrusion.
@property PowerPointMPLd presetLightingDirection;  // Returns or sets the position of the light source relative to the extrusion.
@property PowerPointMLtT presetLightingRig;  // Returns a constant that represents the lighting preset.
@property PowerPointMlSf presetLightingSoftness;  // Returns or sets the intensity of the extrusion lighting.
@property PowerPointMPMt presetMaterial;  // Returns or sets the extrusion surface material.
@property (readonly) PowerPointM3DF presetThreeDFormat;  // Returns or sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion.
@property BOOL projectText;  // Returns or sets whether text on a shape rotates with shape.
@property double rotationX;  // Returns or sets the rotation of the extruded shape around the x-axis in degrees. A positive value indicates upward rotation; a negative value indicates downward rotation.
@property double rotationY;  // Returns or sets the rotation of the extruded shape around the y-axis, in degrees. A positive value indicates rotation to the left; a negative value indicates rotation to the right.
@property double rotationZ;  // Returns or sets the rotation of the extruded shape around the z-axis, in degrees. A positive value indicates clockwise rotation; a negative value indicates anti-clockwise rotation.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

@interface PowerPointWordArtFormat : PowerPointBaseObject

@property BOOL fontBold;
@property BOOL fontItalic;
@property (copy) NSString *fontName;
@property double fontSize;
@property BOOL kernedPairs;
@property BOOL normalizedHeight;
@property PowerPointMPTs presetShape;
@property BOOL rotatedChars;
@property PowerPointMTxA textAlignment;
@property double tracking;
@property PowerPointMPXF wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.
@property (copy) NSString *wordArtText;

- (void) toggleVerticalText;

@end



/*
 * Text Suite
 */

@interface PowerPointTextRange : PowerPointBaseObject

- (SBElementArray<PowerPointCharacter *> *) characters;
- (SBElementArray<PowerPointWord *> *) words;
- (SBElementArray<PowerPointSentence *> *) sentences;
- (SBElementArray<PowerPointLine *> *) lines;
- (SBElementArray<PowerPointParagraph *> *) paragraphs;
- (SBElementArray<PowerPointTextFlow *> *) textFlows;

@property (readonly) double boundsHeight;
@property (readonly) double boundsWidth;
@property (copy) NSString *content;
@property (copy, readonly) PowerPointFont *font;
@property NSInteger indentLevel;
@property (readonly) double leftBounds;
@property (readonly) NSInteger offset;
@property (copy, readonly) PowerPointParagraphFormat *paragraphFormat;
@property (readonly) NSInteger textLength;
@property (readonly) double topBounds;

- (void) addPeriodsTo;
- (void) changeCaseTo:(PowerPointPCgC)to;
- (void) copyTextRange NS_RETURNS_NOT_RETAINED;
- (void) cutTextRange;
- (NSArray *) getRotatedTextBounds;  // Returns back a list containing the four point of the text bounds as follows  {x1, y1}, {x2, y2}, {x3, y3}, {x4, y4} }
- (PowerPointActionSetting *) getTextActionSettingResult:(PowerPointPMtv)result;
- (void) insertTextTextRangeInsertWhere:(PowerPointMTiP)insertWhere newText:(NSString *)newText;  // Adds text to existing text.
- (void) pasteTextRange;
- (void) removePeriodsFrom;

@end

@interface PowerPointCharacter : PowerPointTextRange


@end

@interface PowerPointLine : PowerPointTextRange


@end

@interface PowerPointParagraph : PowerPointTextRange


@end

@interface PowerPointSentence : PowerPointTextRange


@end

@interface PowerPointTextFlow : PowerPointTextRange


@end

@interface PowerPointWord : PowerPointTextRange


@end



/*
 * Table Suite
 */

@interface PowerPointCell : PowerPointBaseObject

@property (readonly) BOOL selected;
@property (copy, readonly) PowerPointShape *shape;

- (PowerPointLineFormat *) getBorderEdge:(PowerPointPBrT)edge;
- (void) mergeMergeWith:(PowerPointCell *)mergeWith;
- (void) splitNumberOfRows:(NSInteger)numberOfRows numberOfColumns:(NSInteger)numberOfColumns;

@end

@interface PowerPointColumn : PowerPointBaseObject

- (SBElementArray<PowerPointCell *> *) cells;

@property double width;


@end

@interface PowerPointRow : PowerPointBaseObject

- (SBElementArray<PowerPointCell *> *) cells;

@property double height;


@end

@interface PowerPointTable : PowerPointBaseObject

- (SBElementArray<PowerPointColumn *> *) columns;
- (SBElementArray<PowerPointRow *> *) rows;

@property PowerPointPDrT tableDirection;

- (PowerPointCell *) getCellFromRow:(NSInteger)row column:(NSInteger)column;

@end

