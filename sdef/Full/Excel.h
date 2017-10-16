/*
 * Excel.h
 */

#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>


@class ExcelBaseObject, ExcelBaseApplication, ExcelBaseDocument, ExcelBasicWindow, ExcelPrintSettings, ExcelCommandBarControl, ExcelCommandBarButton, ExcelCommandBarCombobox, ExcelCommandBarPopup, ExcelCommandBar, ExcelDocumentProperty, ExcelCustomDocumentProperty, ExcelWebPageFont, ExcelExcelComment, ExcelODBCError, ExcelProtection, ExcelAboveAverageFormatCondition, ExcelAddIn, ExcelApplication, ExcelAutofilter, ExcelBorder, ExcelButton, ExcelCalculatedMember, ExcelCheckbox, ExcelColorScaleCriteria, ExcelColorScaleCriterion, ExcelColorScaleFormatCondition, ExcelColorstop, ExcelColorstops, ExcelConditionValue, ExcelCubeField, ExcelCustomView, ExcelDatabarBorder, ExcelDatabarFormatCondition, ExcelDefaultWebOptions, ExcelDialog, ExcelDisplayFormat, ExcelDropdown, ExcelFilter, ExcelFormatColor, ExcelFormatConditionIconObject, ExcelFormatConditionIconSet, ExcelFormatConditionIconSets, ExcelFormatCondition, ExcelGraphic, ExcelGroupbox, ExcelHorizontalPageBreak, ExcelHyperlink, ExcelIconCriteria, ExcelIconCriterion, ExcelIconSetFormatCondition, ExcelLabel, ExcelLinearGradient, ExcelListColumn, ExcelListObject, ExcelListRow, ExcelListbox, ExcelNamedItem, ExcelNegativeBarFormat, ExcelOptionButton, ExcelOutline, ExcelPageSetup, ExcelPane, ExcelPhonetic, ExcelPivotAxis, ExcelPivotCache, ExcelPivotCell, ExcelPivotField, ExcelCalculatedField, ExcelColumnField, ExcelDataField, ExcelHiddenField, ExcelPageField, ExcelPivotFilter, ExcelActiveFilter, ExcelPivotFormula, ExcelPivotItem, ExcelCalculatedItem, ExcelChildItem, ExcelColumnItem, ExcelHiddenItem, ExcelParentItem, ExcelPivotLine, ExcelPivotTable, ExcelQueryTable, ExcelRecentFile, ExcelRectangularGradient, ExcelRowField, ExcelRowItem, ExcelScenario, ExcelScrollbar, ExcelSheet, ExcelInternationalMacroSheet, ExcelMacroSheet, ExcelSlicer, ExcelSort, ExcelSortfield, ExcelSortfields, ExcelSpinner, ExcelTab, ExcelTableStyleElement, ExcelTableStyle, ExcelTextbox, ExcelTop10FormatCondition, ExcelTreeviewControl, ExcelUniqueValuesFormatCondition, ExcelValidation, ExcelValueChange, ExcelVerticalPageBreak, ExcelWebOptions, ExcelWindow, ExcelWorkbookConnection, ExcelWorkbook, ExcelDocument, ExcelWorksheet, ExcelXlspellingOptions, ExcelAdjustment, ExcelArc, ExcelBulletFormat, ExcelCalloutFormat, ExcelConnectorFormat, ExcelFillFormat, ExcelGlowFormat, ExcelGradientStop, ExcelLineFormat, ExcelLine, ExcelOfficeTheme, ExcelOval, ExcelParagraphFormat, ExcelPictureFormat, ExcelRectangle, ExcelReflectionFormat, ExcelRulerLevel, ExcelRuler, ExcelShadowFormat, ExcelShapeFont, ExcelShapeTextFrame, ExcelShape, ExcelCallout, ExcelPicture, ExcelShapeConnector, ExcelShapeLine, ExcelShapeTextbox, ExcelSoftEdgeFormat, ExcelTabStop, ExcelTextColumn, ExcelTextFrame, ExcelThemeColorScheme, ExcelThemeColor, ExcelThemeEffectScheme, ExcelThemeFontScheme, ExcelThemeFont, ExcelMajorThemeFont, ExcelMinorThemeFont, ExcelThreeDFormat, ExcelWordArtFormat, ExcelCharacter, ExcelFont, ExcelStyle, ExcelTextRange, ExcelParagraph, ExcelSentence, ExcelTextFlow, ExcelTextRangeCharacter, ExcelTextRangeLine, ExcelWord, ExcelRange, ExcelCell, ExcelColumn, ExcelRow, ExcelAutocorrect, ExcelAxisTitle, ExcelAxis, ExcelChartArea, ExcelChartFillFormat, ExcelChartFormat, ExcelChartGroup, ExcelAreaGroup, ExcelBarGroup, ExcelChartObject, ExcelChartTitle, ExcelChart, ExcelChartSheet, ExcelColumnGroup, ExcelCorners, ExcelDataLabel, ExcelDataTable, ExcelDisplayUnitLabel, ExcelDoughnutGroup, ExcelDownBars, ExcelDropLines, ExcelErrorBars, ExcelFloor, ExcelGridlines, ExcelHiloLines, ExcelInterior, ExcelLeaderLines, ExcelLegendEntry, ExcelLegendKey, ExcelLegend, ExcelLineGroup, ExcelPieGroup, ExcelPlotArea, ExcelRadarGroup, ExcelSeriesLines, ExcelSeriesPoint, ExcelSeries, ExcelTickLabels, ExcelTrendline, ExcelUpBars, ExcelWalls, ExcelXyGroup;

enum ExcelSavo {
	ExcelSavoYes = 'yes ' /* Save objects now */,
	ExcelSavoNo = 'no  ' /* Do not save objects */,
	ExcelSavoAsk = 'ask ' /* Ask the user whether to save */
};
typedef enum ExcelSavo ExcelSavo;

enum ExcelKfrm {
	ExcelKfrmIndex = 'indx' /* keyform designating indexed access */,
	ExcelKfrmNamed = 'name' /* keyform designating named access */,
	ExcelKfrmId = 'ID  ' /* keyform designating access by unique identifier */
};
typedef enum ExcelKfrm ExcelKfrm;

enum ExcelXLfd {
	ExcelXLfdMacintoshPath = 'mPth' /* Maintosh path i.e. Foo:bar.txt */,
	ExcelXLfdPosixPath = 'file' /* Posix path i.e. file:://foo/bar.txt */
};
typedef enum ExcelXLfd ExcelXLfd;

enum ExcelEnum {
	ExcelEnumStandard = 'lwst' /* Standard PostScript error handling   */,
	ExcelEnumDetailed = 'lwdt' /* print a detailed report of Postscript errors */
};
typedef enum ExcelEnum ExcelEnum;

enum ExcelMlDs {
	ExcelMlDsLineDashStyleUnset = '\000\222\377\376',
	ExcelMlDsLineDashStyleSolid = '\000\223\000\001',
	ExcelMlDsLineDashStyleSquareDot = '\000\223\000\002',
	ExcelMlDsLineDashStyleRoundDot = '\000\223\000\003',
	ExcelMlDsLineDashStyleDash = '\000\223\000\004',
	ExcelMlDsLineDashStyleDashDot = '\000\223\000\005',
	ExcelMlDsLineDashStyleDashDotDot = '\000\223\000\006',
	ExcelMlDsLineDashStyleLongDash = '\000\223\000\007',
	ExcelMlDsLineDashStyleLongDashDot = '\000\223\000\010',
	ExcelMlDsLineDashStyleLongDashDotDot = '\000\223\000\011',
	ExcelMlDsLineDashStyleSystemDash = '\000\223\000\012',
	ExcelMlDsLineDashStyleSystemDot = '\000\223\000\013',
	ExcelMlDsLineDashStyleSystemDashDot = '\000\223\000\014'
};
typedef enum ExcelMlDs ExcelMlDs;

enum ExcelMLnS {
	ExcelMLnSLineStyleUnset = '\000\224\377\376',
	ExcelMLnSSingleLine = '\000\225\000\001',
	ExcelMLnSThinThinLine = '\000\225\000\002',
	ExcelMLnSThinThickLine = '\000\225\000\003',
	ExcelMLnSThickThinLine = '\000\225\000\004',
	ExcelMLnSThickBetweenThinLine = '\000\225\000\005'
};
typedef enum ExcelMLnS ExcelMLnS;

enum ExcelMAhS {
	ExcelMAhSArrowheadStyleUnset = '\000\221\377\376',
	ExcelMAhSNoArrowhead = '\000\222\000\001',
	ExcelMAhSTriangleArrowhead = '\000\222\000\002',
	ExcelMAhSOpen_arrowhead = '\000\222\000\003',
	ExcelMAhSStealthArrowhead = '\000\222\000\004',
	ExcelMAhSDiamondArrowhead = '\000\222\000\005',
	ExcelMAhSOvalArrowhead = '\000\222\000\006'
};
typedef enum ExcelMAhS ExcelMAhS;

enum ExcelMAhW {
	ExcelMAhWArrowheadWidthUnset = '\000\220\377\376',
	ExcelMAhWNarrowWidthArrowhead = '\000\221\000\001',
	ExcelMAhWMediumWidthArrowhead = '\000\221\000\002',
	ExcelMAhWWideArrowhead = '\000\221\000\003'
};
typedef enum ExcelMAhW ExcelMAhW;

enum ExcelMAhL {
	ExcelMAhLArrowheadLengthUnset = '\000\223\377\376',
	ExcelMAhLShortArrowhead = '\000\224\000\001',
	ExcelMAhLMediumArrowhead = '\000\224\000\002',
	ExcelMAhLLongArrowhead = '\000\224\000\003'
};
typedef enum ExcelMAhL ExcelMAhL;

enum ExcelMFdT {
	ExcelMFdTFillUnset = '\000c\377\376',
	ExcelMFdTFillSolid = '\000d\000\001',
	ExcelMFdTFillPatterned = '\000d\000\002',
	ExcelMFdTFillGradient = '\000d\000\003',
	ExcelMFdTFillTextured = '\000d\000\004',
	ExcelMFdTFillBackground = '\000d\000\005',
	ExcelMFdTFillPicture = '\000d\000\006'
};
typedef enum ExcelMFdT ExcelMFdT;

enum ExcelMGdS {
	ExcelMGdSGradientUnset = '\000d\377\376',
	ExcelMGdSHorizontalGradient = '\000e\000\001',
	ExcelMGdSVerticalGradient = '\000e\000\002',
	ExcelMGdSDiagonalUpGradient = '\000e\000\003',
	ExcelMGdSDiagonalDownGradient = '\000e\000\004',
	ExcelMGdSFromCornerGradient = '\000e\000\005',
	ExcelMGdSFromTitleGradient = '\000e\000\006',
	ExcelMGdSFromCenterGradient = '\000e\000\007'
};
typedef enum ExcelMGdS ExcelMGdS;

enum ExcelMGCt {
	ExcelMGCtGradientTypeUnset = '\003\357\377\376',
	ExcelMGCtSingleShadeGradientType = '\003\360\000\001',
	ExcelMGCtTwoColorsGradientType = '\003\360\000\002',
	ExcelMGCtPresetColorsGradientType = '\003\360\000\003',
	ExcelMGCtMultiColorsGradientType = '\003\360\000\004'
};
typedef enum ExcelMGCt ExcelMGCt;

enum ExcelMxtT {
	ExcelMxtTTextureTypeTextureTypeUnset = '\003\360\377\376',
	ExcelMxtTTextureTypePresetTexture = '\003\361\000\001',
	ExcelMxtTTextureTypeUserDefinedTexture = '\003\361\000\002'
};
typedef enum ExcelMxtT ExcelMxtT;

enum ExcelMPzT {
	ExcelMPzTPresetTextureUnset = '\000e\377\376',
	ExcelMPzTTexturePapyrus = '\000f\000\001',
	ExcelMPzTTextureCanvas = '\000f\000\002',
	ExcelMPzTTextureDenim = '\000f\000\003',
	ExcelMPzTTextureWovenMat = '\000f\000\004',
	ExcelMPzTTextureWaterDroplets = '\000f\000\005',
	ExcelMPzTTexturePaperBag = '\000f\000\006',
	ExcelMPzTTextureFishFossil = '\000f\000\007',
	ExcelMPzTTextureSand = '\000f\000\010',
	ExcelMPzTTextureGreenMarble = '\000f\000\011',
	ExcelMPzTTextureWhiteMarble = '\000f\000\012',
	ExcelMPzTTextureBrownMarble = '\000f\000\013',
	ExcelMPzTTextureGranite = '\000f\000\014',
	ExcelMPzTTextureNewsprint = '\000f\000\015',
	ExcelMPzTTextureRecycledPaper = '\000f\000\016',
	ExcelMPzTTextureParchment = '\000f\000\017',
	ExcelMPzTTextureStationery = '\000f\000\020',
	ExcelMPzTTextureBlueTissuePaper = '\000f\000\021',
	ExcelMPzTTexturePinkTissuePaper = '\000f\000\022',
	ExcelMPzTTexturePurpleMesh = '\000f\000\023',
	ExcelMPzTTextureBouquet = '\000f\000\024',
	ExcelMPzTTextureCork = '\000f\000\025',
	ExcelMPzTTextureWalnut = '\000f\000\026',
	ExcelMPzTTextureOak = '\000f\000\027',
	ExcelMPzTTextureMediumWood = '\000f\000\030'
};
typedef enum ExcelMPzT ExcelMPzT;

enum ExcelPpTy {
	ExcelPpTyPatternUnset = '\000f\377\376',
	ExcelPpTyFivePercentPattern = '\000g\000\001',
	ExcelPpTyTenPercentPattern = '\000g\000\002',
	ExcelPpTyTwentyPercentPattern = '\000g\000\003',
	ExcelPpTyTwentyFivePercentPattern = '\000g\000\004',
	ExcelPpTyThirtyPercentPattern = '\000g\000\005',
	ExcelPpTyFortyPercentPattern = '\000g\000\006',
	ExcelPpTyFiftyPercentPattern = '\000g\000\007',
	ExcelPpTySixtyPercentPattern = '\000g\000\010',
	ExcelPpTySeventyPercentPattern = '\000g\000\011',
	ExcelPpTySeventyFivePercentPattern = '\000g\000\012',
	ExcelPpTyEightyPercentPattern = '\000g\000\013',
	ExcelPpTyNinetyPercentPattern = '\000g\000\014',
	ExcelPpTyDarkHorizontalPattern = '\000g\000\015',
	ExcelPpTyDarkVerticalPattern = '\000g\000\016',
	ExcelPpTyDarkDownwardDiagonalPattern = '\000g\000\017',
	ExcelPpTyDarkUpwardDiagonalPattern = '\000g\000\020',
	ExcelPpTySmallCheckerBoardPattern = '\000g\000\021',
	ExcelPpTyTrellisPattern = '\000g\000\022',
	ExcelPpTyLightHorizontalPattern = '\000g\000\023',
	ExcelPpTyLightVerticalPattern = '\000g\000\024',
	ExcelPpTyLightDownwardDiagonalPattern = '\000g\000\025',
	ExcelPpTyLightUpwardDiagonalPattern = '\000g\000\026',
	ExcelPpTySmallGridPattern = '\000g\000\027',
	ExcelPpTyDottedDiamondPattern = '\000g\000\030',
	ExcelPpTyWideDownwardDiagonal = '\000g\000\031',
	ExcelPpTyWideUpwardDiagonalPattern = '\000g\000\032',
	ExcelPpTyDashedUpwardDiagonalPattern = '\000g\000\033',
	ExcelPpTyDashedDownwardDiagonalPattern = '\000g\000\034',
	ExcelPpTyNarrowVerticalPattern = '\000g\000\035',
	ExcelPpTyNarrowHorizontalPattern = '\000g\000\036',
	ExcelPpTyDashedVerticalPattern = '\000g\000\037',
	ExcelPpTyDashedHorizontalPattern = '\000g\000 ',
	ExcelPpTyLargeConfettiPattern = '\000g\000!',
	ExcelPpTyLargeGridPattern = '\000g\000\"',
	ExcelPpTyHorizontalBrickPattern = '\000g\000#',
	ExcelPpTyLargeCheckerBoardPattern = '\000g\000$',
	ExcelPpTySmallConfettiPattern = '\000g\000%',
	ExcelPpTyZigZagPattern = '\000g\000&',
	ExcelPpTySolidDiamondPattern = '\000g\000\'',
	ExcelPpTyDiagonalBrickPattern = '\000g\000(',
	ExcelPpTyOutlinedDiamondPattern = '\000g\000)',
	ExcelPpTyPlaidPattern = '\000g\000*',
	ExcelPpTySpherePattern = '\000g\000+',
	ExcelPpTyWeavePattern = '\000g\000,',
	ExcelPpTyDottedGridPattern = '\000g\000-',
	ExcelPpTyDivotPattern = '\000g\000.',
	ExcelPpTyShinglePattern = '\000g\000/',
	ExcelPpTyWavePattern = '\000g\0000',
	ExcelPpTyHorizontalPattern = '\000g\0001',
	ExcelPpTyVerticalPattern = '\000g\0002',
	ExcelPpTyCrossPattern = '\000g\0003',
	ExcelPpTyDownwardDiagonalPattern = '\000g\0004',
	ExcelPpTyUpwardDiagonalPattern = '\000g\0005',
	ExcelPpTyDiagonalCrossPattern = '\000g\0005'
};
typedef enum ExcelPpTy ExcelPpTy;

enum ExcelMPGb {
	ExcelMPGbPresetGradientUnset = '\000g\377\376',
	ExcelMPGbGradientEarlySunset = '\000h\000\001',
	ExcelMPGbGradientLateSunset = '\000h\000\002',
	ExcelMPGbGradientNightfall = '\000h\000\003',
	ExcelMPGbGradientDaybreak = '\000h\000\004',
	ExcelMPGbGradientHorizon = '\000h\000\005',
	ExcelMPGbGradientDesert = '\000h\000\006',
	ExcelMPGbGradientOcean = '\000h\000\007',
	ExcelMPGbGradientCalmWater = '\000h\000\010',
	ExcelMPGbGradientFire = '\000h\000\011',
	ExcelMPGbGradientFog = '\000h\000\012',
	ExcelMPGbGradientMoss = '\000h\000\013',
	ExcelMPGbGradientPeacock = '\000h\000\014',
	ExcelMPGbGradientWheat = '\000h\000\015',
	ExcelMPGbGradientParchment = '\000h\000\016',
	ExcelMPGbGradientMahogany = '\000h\000\017',
	ExcelMPGbGradientRainbow = '\000h\000\020',
	ExcelMPGbGradientRainbow2 = '\000h\000\021',
	ExcelMPGbGradientGold = '\000h\000\022',
	ExcelMPGbGradientGold2 = '\000h\000\023',
	ExcelMPGbGradientBrass = '\000h\000\024',
	ExcelMPGbGradientChrome = '\000h\000\025',
	ExcelMPGbGradientChrome2 = '\000h\000\026',
	ExcelMPGbGradientSilver = '\000h\000\027',
	ExcelMPGbGradientSapphire = '\000h\000\030'
};
typedef enum ExcelMPGb ExcelMPGb;

enum ExcelMSdT {
	ExcelMSdTShadowUnset = '\003_\377\376',
	ExcelMSdTShadow1 = '\003`\000\001',
	ExcelMSdTShadow2 = '\003`\000\002',
	ExcelMSdTShadow3 = '\003`\000\003',
	ExcelMSdTShadow4 = '\003`\000\004',
	ExcelMSdTShadow5 = '\003`\000\005',
	ExcelMSdTShadow6 = '\003`\000\006',
	ExcelMSdTShadow7 = '\003`\000\007',
	ExcelMSdTShadow8 = '\003`\000\010',
	ExcelMSdTShadow9 = '\003`\000\011',
	ExcelMSdTShadow10 = '\003`\000\012',
	ExcelMSdTShadow11 = '\003`\000\013',
	ExcelMSdTShadow12 = '\003`\000\014',
	ExcelMSdTShadow13 = '\003`\000\015',
	ExcelMSdTShadow14 = '\003`\000\016',
	ExcelMSdTShadow15 = '\003`\000\017',
	ExcelMSdTShadow16 = '\003`\000\020',
	ExcelMSdTShadow17 = '\003`\000\021',
	ExcelMSdTShadow18 = '\003`\000\022',
	ExcelMSdTShadow19 = '\003`\000\023',
	ExcelMSdTShadow20 = '\003`\000\024',
	ExcelMSdTShadow21 = '\003`\000\025',
	ExcelMSdTShadow22 = '\003`\000\026',
	ExcelMSdTShadow23 = '\003`\000\027',
	ExcelMSdTShadow24 = '\003`\000\030',
	ExcelMSdTShadow25 = '\003`\000\031',
	ExcelMSdTShadow26 = '\003`\000\032',
	ExcelMSdTShadow27 = '\003`\000\033',
	ExcelMSdTShadow28 = '\003`\000\034',
	ExcelMSdTShadow29 = '\003`\000\035',
	ExcelMSdTShadow30 = '\003`\000\036',
	ExcelMSdTShadow31 = '\003`\000\037',
	ExcelMSdTShadow32 = '\003`\000 ',
	ExcelMSdTShadow33 = '\003`\000!',
	ExcelMSdTShadow34 = '\003`\000\"',
	ExcelMSdTShadow35 = '\003`\000#',
	ExcelMSdTShadow36 = '\003`\000$',
	ExcelMSdTShadow37 = '\003`\000%',
	ExcelMSdTShadow38 = '\003`\000&',
	ExcelMSdTShadow39 = '\003`\000\'',
	ExcelMSdTShadow40 = '\003`\000(',
	ExcelMSdTShadow41 = '\003`\000)',
	ExcelMSdTShadow42 = '\003`\000*',
	ExcelMSdTShadow43 = '\003`\000+'
};
typedef enum ExcelMSdT ExcelMSdT;

enum ExcelMPXF {
	ExcelMPXFWordartFormatUnset = '\003\361\377\376',
	ExcelMPXFWordartFormat1 = '\003\362\000\000',
	ExcelMPXFWordartFormat2 = '\003\362\000\001',
	ExcelMPXFWordartFormat3 = '\003\362\000\002',
	ExcelMPXFWordartFormat4 = '\003\362\000\003',
	ExcelMPXFWordartFormat5 = '\003\362\000\004',
	ExcelMPXFWordartFormat6 = '\003\362\000\005',
	ExcelMPXFWordartFormat7 = '\003\362\000\006',
	ExcelMPXFWordartFormat8 = '\003\362\000\007',
	ExcelMPXFWordartFormat9 = '\003\362\000\010',
	ExcelMPXFWordartFormat10 = '\003\362\000\011',
	ExcelMPXFWordartFormat11 = '\003\362\000\012',
	ExcelMPXFWordartFormat12 = '\003\362\000\013',
	ExcelMPXFWordartFormat13 = '\003\362\000\014',
	ExcelMPXFWordartFormat14 = '\003\362\000\015',
	ExcelMPXFWordartFormat15 = '\003\362\000\016',
	ExcelMPXFWordartFormat16 = '\003\362\000\017',
	ExcelMPXFWordartFormat17 = '\003\362\000\020',
	ExcelMPXFWordartFormat18 = '\003\362\000\021',
	ExcelMPXFWordartFormat19 = '\003\362\000\022',
	ExcelMPXFWordartFormat20 = '\003\362\000\023',
	ExcelMPXFWordartFormat21 = '\003\362\000\024',
	ExcelMPXFWordartFormat22 = '\003\362\000\025',
	ExcelMPXFWordartFormat23 = '\003\362\000\026',
	ExcelMPXFWordartFormat24 = '\003\362\000\027',
	ExcelMPXFWordartFormat25 = '\003\362\000\030',
	ExcelMPXFWordartFormat26 = '\003\362\000\031',
	ExcelMPXFWordartFormat27 = '\003\362\000\032',
	ExcelMPXFWordartFormat28 = '\003\362\000\033',
	ExcelMPXFWordartFormat29 = '\003\362\000\034',
	ExcelMPXFWordartFormat30 = '\003\362\000\035'
};
typedef enum ExcelMPXF ExcelMPXF;

enum ExcelMPTs {
	ExcelMPTsTextEffectShapeUnset = '\000\227\377\376',
	ExcelMPTsPlainText = '\000\230\000\001',
	ExcelMPTsStop = '\000\230\000\002',
	ExcelMPTsTriangleUp = '\000\230\000\003',
	ExcelMPTsTriangleDown = '\000\230\000\004',
	ExcelMPTsChevronUp = '\000\230\000\005',
	ExcelMPTsChevronDown = '\000\230\000\006',
	ExcelMPTsRingInside = '\000\230\000\007',
	ExcelMPTsRingOutside = '\000\230\000\010',
	ExcelMPTsArchUpCurve = '\000\230\000\011',
	ExcelMPTsArchDownCurve = '\000\230\000\012',
	ExcelMPTsCircleCurve = '\000\230\000\013',
	ExcelMPTsButtonCurve = '\000\230\000\014',
	ExcelMPTsArchUpPour = '\000\230\000\015',
	ExcelMPTsArchDownPour = '\000\230\000\016',
	ExcelMPTsCirclePour = '\000\230\000\017',
	ExcelMPTsButtonPour = '\000\230\000\020',
	ExcelMPTsCurveUp = '\000\230\000\021',
	ExcelMPTsCurveDown = '\000\230\000\022',
	ExcelMPTsCanUp = '\000\230\000\023',
	ExcelMPTsCanDown = '\000\230\000\024',
	ExcelMPTsWave1 = '\000\230\000\025',
	ExcelMPTsWave2 = '\000\230\000\026',
	ExcelMPTsDoubleWave1 = '\000\230\000\027',
	ExcelMPTsDoubleWave2 = '\000\230\000\030',
	ExcelMPTsInflate = '\000\230\000\031',
	ExcelMPTsDeflate = '\000\230\000\032',
	ExcelMPTsInflateBottom = '\000\230\000\033',
	ExcelMPTsDeflateBottom = '\000\230\000\034',
	ExcelMPTsInflateTop = '\000\230\000\035',
	ExcelMPTsDeflateTop = '\000\230\000\036',
	ExcelMPTsDeflateInflate = '\000\230\000\037',
	ExcelMPTsDeflateInflateDeflate = '\000\230\000 ',
	ExcelMPTsFadeRight = '\000\230\000!',
	ExcelMPTsFadeLeft = '\000\230\000\"',
	ExcelMPTsFadeUp = '\000\230\000#',
	ExcelMPTsFadeDown = '\000\230\000$',
	ExcelMPTsSlantUp = '\000\230\000%',
	ExcelMPTsSlantDown = '\000\230\000&',
	ExcelMPTsCascadeUp = '\000\230\000\'',
	ExcelMPTsCascadeDown = '\000\230\000('
};
typedef enum ExcelMPTs ExcelMPTs;

enum ExcelMTxA {
	ExcelMTxATextEffectAlignmentUnset = '\000\226\377\376',
	ExcelMTxALeftTextEffectAlignment = '\000\227\000\001',
	ExcelMTxACenteredTextEffectAlignment = '\000\227\000\002',
	ExcelMTxARightTextEffectAlignment = '\000\227\000\003',
	ExcelMTxAJustifyTextEffectAlignment = '\000\227\000\004',
	ExcelMTxAWordJustifyTextEffectAlignment = '\000\227\000\005',
	ExcelMTxAStretchJustifyTextEffectAlignment = '\000\227\000\006'
};
typedef enum ExcelMTxA ExcelMTxA;

enum ExcelMPLd {
	ExcelMPLdPresetLightingDirectionUnset = '\000\233\377\376',
	ExcelMPLdLightFromTopLeft = '\000\234\000\001',
	ExcelMPLdLightFromTop = '\000\234\000\002',
	ExcelMPLdLightFromTopRight = '\000\234\000\003',
	ExcelMPLdLightFromLeft = '\000\234\000\004',
	ExcelMPLdLightFromNone = '\000\234\000\005',
	ExcelMPLdLightFromRight = '\000\234\000\006',
	ExcelMPLdLightFromBottomLeft = '\000\234\000\007',
	ExcelMPLdLightFromBottom = '\000\234\000\010',
	ExcelMPLdLightFromBottomRight = '\000\234\000\011'
};
typedef enum ExcelMPLd ExcelMPLd;

enum ExcelMlSf {
	ExcelMlSfLightingSoftnessUnset = '\000\234\377\376',
	ExcelMlSfLightingDim = '\000\235\000\001',
	ExcelMlSfLightingNormal = '\000\235\000\002',
	ExcelMlSfLightingBright = '\000\235\000\003'
};
typedef enum ExcelMlSf ExcelMlSf;

enum ExcelMPMt {
	ExcelMPMtPresetMaterialUnset = '\000\235\377\376',
	ExcelMPMtMatte = '\000\236\000\001',
	ExcelMPMtPlastic = '\000\236\000\002',
	ExcelMPMtMetal = '\000\236\000\003',
	ExcelMPMtWireframe = '\000\236\000\004',
	ExcelMPMtMatte2 = '\000\236\000\005',
	ExcelMPMtPlastic2 = '\000\236\000\006',
	ExcelMPMtMetal2 = '\000\236\000\007',
	ExcelMPMtWarmMatte = '\000\236\000\010',
	ExcelMPMtTranslucentPowder = '\000\236\000\011',
	ExcelMPMtPowder = '\000\236\000\012',
	ExcelMPMtDarkEdge = '\000\236\000\013',
	ExcelMPMtSoftEdge = '\000\236\000\014',
	ExcelMPMtMaterialClear = '\000\236\000\015',
	ExcelMPMtFlat = '\000\236\000\016',
	ExcelMPMtSoftMetal = '\000\236\000\017'
};
typedef enum ExcelMPMt ExcelMPMt;

enum ExcelMExD {
	ExcelMExDPresetExtrusionDirectionUnset = '\000\231\377\376',
	ExcelMExDExtrudeBottomRight = '\000\232\000\001',
	ExcelMExDExtrudeBottom = '\000\232\000\002',
	ExcelMExDExtrudeBottomLeft = '\000\232\000\003',
	ExcelMExDExtrudeRight = '\000\232\000\004',
	ExcelMExDExtrudeNone = '\000\232\000\005',
	ExcelMExDExtrudeLeft = '\000\232\000\006',
	ExcelMExDExtrudeTopRight = '\000\232\000\007',
	ExcelMExDExtrudeTop = '\000\232\000\010',
	ExcelMExDExtrudeTopLeft = '\000\232\000\011'
};
typedef enum ExcelMExD ExcelMExD;

enum ExcelM3DF {
	ExcelM3DFPresetThreeDFormatUnset = '\000\230\377\376',
	ExcelM3DFFormat1 = '\000\231\000\001',
	ExcelM3DFFormat2 = '\000\231\000\002',
	ExcelM3DFFormat3 = '\000\231\000\003',
	ExcelM3DFFormat4 = '\000\231\000\004',
	ExcelM3DFFormat5 = '\000\231\000\005',
	ExcelM3DFFormat6 = '\000\231\000\006',
	ExcelM3DFFormat7 = '\000\231\000\007',
	ExcelM3DFFormat8 = '\000\231\000\010',
	ExcelM3DFFormat9 = '\000\231\000\011',
	ExcelM3DFFormat10 = '\000\231\000\012',
	ExcelM3DFFormat11 = '\000\231\000\013',
	ExcelM3DFFormat12 = '\000\231\000\014',
	ExcelM3DFFormat13 = '\000\231\000\015',
	ExcelM3DFFormat14 = '\000\231\000\016',
	ExcelM3DFFormat15 = '\000\231\000\017',
	ExcelM3DFFormat16 = '\000\231\000\020',
	ExcelM3DFFormat17 = '\000\231\000\021',
	ExcelM3DFFormat18 = '\000\231\000\022',
	ExcelM3DFFormat19 = '\000\231\000\023',
	ExcelM3DFFormat20 = '\000\231\000\024'
};
typedef enum ExcelM3DF ExcelM3DF;

enum ExcelMExC {
	ExcelMExCExtrusionColorUnset = '\000\232\377\376',
	ExcelMExCExtrusionColorAutomatic = '\000\233\000\001',
	ExcelMExCExtrusionColorCustom = '\000\233\000\002'
};
typedef enum ExcelMExC ExcelMExC;

enum ExcelMCtT {
	ExcelMCtTConnectorTypeUnset = '\000h\377\376',
	ExcelMCtTStraight = '\000i\000\001',
	ExcelMCtTElbow = '\000i\000\002',
	ExcelMCtTCurve = '\000i\000\003'
};
typedef enum ExcelMCtT ExcelMCtT;

enum ExcelMHzA {
	ExcelMHzAHorizontalAnchorUnset = '\000\236\377\376',
	ExcelMHzAHorizontalAnchorNone = '\000\237\000\001',
	ExcelMHzAHorizontalAnchorCenter = '\000\237\000\002'
};
typedef enum ExcelMHzA ExcelMHzA;

enum ExcelMVtA {
	ExcelMVtAVerticalAnchorUnset = '\000\237\377\376',
	ExcelMVtAAnchorTop = '\000\240\000\001',
	ExcelMVtAAnchorTopBaseline = '\000\240\000\002',
	ExcelMVtAAnchorMiddle = '\000\240\000\003',
	ExcelMVtAAnchorBottom = '\000\240\000\004',
	ExcelMVtAAnchorBottomBaseline = '\000\240\000\005'
};
typedef enum ExcelMVtA ExcelMVtA;

enum ExcelMAsT {
	ExcelMAsTAutoshapeShapeTypeUnset = '\000i\377\376',
	ExcelMAsTAutoshapeRectangle = '\000j\000\001',
	ExcelMAsTAutoshapeParallelogram = '\000j\000\002',
	ExcelMAsTAutoshapeTrapezoid = '\000j\000\003',
	ExcelMAsTAutoshapeDiamond = '\000j\000\004',
	ExcelMAsTAutoshapeRoundedRectangle = '\000j\000\005',
	ExcelMAsTAutoshapeOctagon = '\000j\000\006',
	ExcelMAsTAutoshapeIsoscelesTriangle = '\000j\000\007',
	ExcelMAsTAutoshapeRightTriangle = '\000j\000\010',
	ExcelMAsTAutoshapeOval = '\000j\000\011',
	ExcelMAsTAutoshapeHexagon = '\000j\000\012',
	ExcelMAsTAutoshapeCross = '\000j\000\013',
	ExcelMAsTAutoshapeRegularPentagon = '\000j\000\014',
	ExcelMAsTAutoshapeCan = '\000j\000\015',
	ExcelMAsTAutoshapeCube = '\000j\000\016',
	ExcelMAsTAutoshapeBevel = '\000j\000\017',
	ExcelMAsTAutoshapeFoldedCorner = '\000j\000\020',
	ExcelMAsTAutoshapeSmileyFace = '\000j\000\021',
	ExcelMAsTAutoshapeDonut = '\000j\000\022',
	ExcelMAsTAutoshapeNoSymbol = '\000j\000\023',
	ExcelMAsTAutoshapeBlockArc = '\000j\000\024',
	ExcelMAsTAutoshapeHeart = '\000j\000\025',
	ExcelMAsTAutoshapeLightningBolt = '\000j\000\026',
	ExcelMAsTAutoshapeSun = '\000j\000\027',
	ExcelMAsTAutoshapeMoon = '\000j\000\030',
	ExcelMAsTAutoshapeArc = '\000j\000\031',
	ExcelMAsTAutoshapeDoubleBracket = '\000j\000\032',
	ExcelMAsTAutoshapeDoubleBrace = '\000j\000\033',
	ExcelMAsTAutoshapePlaque = '\000j\000\034',
	ExcelMAsTAutoshapeLeftBracket = '\000j\000\035',
	ExcelMAsTAutoshapeRightBracket = '\000j\000\036',
	ExcelMAsTAutoshapeLeftBrace = '\000j\000\037',
	ExcelMAsTAutoshapeRightBrace = '\000j\000 ',
	ExcelMAsTAutoshapeRightArrow = '\000j\000!',
	ExcelMAsTAutoshapeLeftArrow = '\000j\000\"',
	ExcelMAsTAutoshapeUpArrow = '\000j\000#',
	ExcelMAsTAutoshapeDownArrow = '\000j\000$',
	ExcelMAsTAutoshapeLeftRightArrow = '\000j\000%',
	ExcelMAsTAutoshapeUpDownArrow = '\000j\000&',
	ExcelMAsTAutoshapeQuadArrow = '\000j\000\'',
	ExcelMAsTAutoshapeLeftRightUpArrow = '\000j\000(',
	ExcelMAsTAutoshapeBentArrow = '\000j\000)',
	ExcelMAsTAutoshapeUTurnArrow = '\000j\000*',
	ExcelMAsTAutoshapeLeftUpArrow = '\000j\000+',
	ExcelMAsTAutoshapeBentUpArrow = '\000j\000,',
	ExcelMAsTAutoshapeCurvedRightArrow = '\000j\000-',
	ExcelMAsTAutoshapeCurvedLeftArrow = '\000j\000.',
	ExcelMAsTAutoshapeCurvedUpArrow = '\000j\000/',
	ExcelMAsTAutoshapeCurvedDownArrow = '\000j\0000',
	ExcelMAsTAutoshapeStripedRightArrow = '\000j\0001',
	ExcelMAsTAutoshapeNotchedRightArrow = '\000j\0002',
	ExcelMAsTAutoshapePentagon = '\000j\0003',
	ExcelMAsTAutoshapeChevron = '\000j\0004',
	ExcelMAsTAutoshapeRightArrowCallout = '\000j\0005',
	ExcelMAsTAutoshapeLeftArrowCallout = '\000j\0006',
	ExcelMAsTAutoshapeUpArrowCallout = '\000j\0007',
	ExcelMAsTAutoshapeDownArrowCallout = '\000j\0008',
	ExcelMAsTAutoshapeLeftRightArrowCallout = '\000j\0009',
	ExcelMAsTAutoshapeUpDownArrowCallout = '\000j\000:',
	ExcelMAsTAutoshapeQuadArrowCallout = '\000j\000;',
	ExcelMAsTAutoshapeCircularArrow = '\000j\000<',
	ExcelMAsTAutoshapeFlowchartProcess = '\000j\000=',
	ExcelMAsTAutoshapeFlowchartAlternateProcess = '\000j\000>',
	ExcelMAsTAutoshapeFlowchartDecision = '\000j\000\?',
	ExcelMAsTAutoshapeFlowchartData = '\000j\000@',
	ExcelMAsTAutoshapeFlowchartPredefinedProcess = '\000j\000A',
	ExcelMAsTAutoshapeFlowchartInternalStorage = '\000j\000B',
	ExcelMAsTAutoshapeFlowchartDocument = '\000j\000C',
	ExcelMAsTAutoshapeFlowchartMultiDocument = '\000j\000D',
	ExcelMAsTAutoshapeFlowchartTerminator = '\000j\000E',
	ExcelMAsTAutoshapeFlowchartPreparation = '\000j\000F',
	ExcelMAsTAutoshapeFlowchartManualInput = '\000j\000G',
	ExcelMAsTAutoshapeFlowchartManualOperation = '\000j\000H',
	ExcelMAsTAutoshapeFlowchartConnector = '\000j\000I',
	ExcelMAsTAutoshapeFlowchartOffpageConnector = '\000j\000J',
	ExcelMAsTAutoshapeFlowchartCard = '\000j\000K',
	ExcelMAsTAutoshapeFlowchartPunchedTape = '\000j\000L',
	ExcelMAsTAutoshapeFlowchartSummingJunction = '\000j\000M',
	ExcelMAsTAutoshapeFlowchartOr = '\000j\000N',
	ExcelMAsTAutoshapeFlowchartCollate = '\000j\000O',
	ExcelMAsTAutoshapeFlowchartSort = '\000j\000P',
	ExcelMAsTAutoshapeFlowchartExtract = '\000j\000Q',
	ExcelMAsTAutoshapeFlowchartMerge = '\000j\000R',
	ExcelMAsTAutoshapeFlowchartStoredData = '\000j\000S',
	ExcelMAsTAutoshapeFlowchartDelay = '\000j\000T',
	ExcelMAsTAutoshapeFlowchartSequentialAccessStorage = '\000j\000U',
	ExcelMAsTAutoshapeFlowchartMagneticDisk = '\000j\000V',
	ExcelMAsTAutoshapeFlowchartDirectAccessStorage = '\000j\000W',
	ExcelMAsTAutoshapeFlowchartDisplay = '\000j\000X',
	ExcelMAsTAutoshapeExplosionOne = '\000j\000Y',
	ExcelMAsTAutoshapeExplosionTwo = '\000j\000Z',
	ExcelMAsTAutoshapeFourPointStar = '\000j\000[',
	ExcelMAsTAutoshapeFivePointStar = '\000j\000\\',
	ExcelMAsTAutoshapeEightPointStar = '\000j\000]',
	ExcelMAsTAutoshapeSixteenPointStar = '\000j\000^',
	ExcelMAsTAutoshapeTwentyFourPointStar = '\000j\000_',
	ExcelMAsTAutoshapeThirtyTwoPointStar = '\000j\000`',
	ExcelMAsTAutoshapeUpRibbon = '\000j\000a',
	ExcelMAsTAutoshapeDownRibbon = '\000j\000b',
	ExcelMAsTAutoshapeCurvedUpRibbon = '\000j\000c',
	ExcelMAsTAutoshapeCurvedDownRibbon = '\000j\000d',
	ExcelMAsTAutoshapeVerticalScroll = '\000j\000e',
	ExcelMAsTAutoshapeHorizontalScroll = '\000j\000f',
	ExcelMAsTAutoshapeWave = '\000j\000g',
	ExcelMAsTAutoshapeDoubleWave = '\000j\000h',
	ExcelMAsTAutoshapeRectangularCallout = '\000j\000i',
	ExcelMAsTAutoshapeRoundedRectangularCallout = '\000j\000j',
	ExcelMAsTAutoshapeOvalCallout = '\000j\000k',
	ExcelMAsTAutoshapeCloudCallout = '\000j\000l',
	ExcelMAsTAutoshapeLineCalloutOne = '\000j\000m',
	ExcelMAsTAutoshapeLineCalloutTwo = '\000j\000n',
	ExcelMAsTAutoshapeLineCalloutThree = '\000j\000o',
	ExcelMAsTAutoshapeLineCalloutFour = '\000j\000p',
	ExcelMAsTAutoshapeLineCalloutOneAccentBar = '\000j\000q',
	ExcelMAsTAutoshapeLineCalloutTwoAccentBar = '\000j\000r',
	ExcelMAsTAutoshapeLineCalloutThreeAccentBar = '\000j\000s',
	ExcelMAsTAutoshapeLineCalloutFourAccentBar = '\000j\000t',
	ExcelMAsTAutoshapeLineCalloutOneNoBorder = '\000j\000u',
	ExcelMAsTAutoshapeLineCalloutTwoNoBorder = '\000j\000v',
	ExcelMAsTAutoshapeLineCalloutThreeNoBorder = '\000j\000w',
	ExcelMAsTAutoshapeLineCalloutFourNoBorder = '\000j\000x',
	ExcelMAsTAutoshapeCalloutOneBorderAndAccentBar = '\000j\000y',
	ExcelMAsTAutoshapeCalloutTwoBorderAndAccentBar = '\000j\000z',
	ExcelMAsTAutoshapeCalloutThreeBorderAndAccentBar = '\000j\000{',
	ExcelMAsTAutoshapeCalloutFourBorderAndAccentBar = '\000j\000|',
	ExcelMAsTAutoshapeActionButtonCustom = '\000j\000}',
	ExcelMAsTAutoshapeActionButtonHome = '\000j\000~',
	ExcelMAsTAutoshapeActionButtonHelp = '\000j\000\177',
	ExcelMAsTAutoshapeActionButtonInformation = '\000j\000\200',
	ExcelMAsTAutoshapeActionButtonBackOrPrevious = '\000j\000\201',
	ExcelMAsTAutoshapeActionButtonForwardOrNext = '\000j\000\202',
	ExcelMAsTAutoshapeActionButtonBeginning = '\000j\000\203',
	ExcelMAsTAutoshapeActionButtonEnd = '\000j\000\204',
	ExcelMAsTAutoshapeActionButtonReturn = '\000j\000\205',
	ExcelMAsTAutoshapeActionButtonDocument = '\000j\000\206',
	ExcelMAsTAutoshapeActionButtonSound = '\000j\000\207',
	ExcelMAsTAutoshapeActionButtonMovie = '\000j\000\210',
	ExcelMAsTAutoshapeBalloon = '\000j\000\211',
	ExcelMAsTAutoshapeNotPrimitive = '\000j\000\212',
	ExcelMAsTAutoshapeFlowchartOfflineStorage = '\000j\000\213',
	ExcelMAsTAutoshapeLeftRightRibbon = '\000j\000\214',
	ExcelMAsTAutoshapeDiagonalStripe = '\000j\000\215',
	ExcelMAsTAutoshapePie = '\000j\000\216',
	ExcelMAsTAutoshapeNonIsoscelesTrapezoid = '\000j\000\217',
	ExcelMAsTAutoshapeDecagon = '\000j\000\220',
	ExcelMAsTAutoshapeHeptagon = '\000j\000\221',
	ExcelMAsTAutoshapeDodecagon = '\000j\000\222',
	ExcelMAsTAutoshapeSixPointsStar = '\000j\000\223',
	ExcelMAsTAutoshapeSevenPointsStar = '\000j\000\224',
	ExcelMAsTAutoshapeTenPointsStar = '\000j\000\225',
	ExcelMAsTAutoshapeTwelvePointsStar = '\000j\000\226',
	ExcelMAsTAutoshapeRoundOneRectangle = '\000j\000\227',
	ExcelMAsTAutoshapeRoundTwoSameRectangle = '\000j\000\230',
	ExcelMAsTAutoshapeRoundTwoDiagonalRectangle = '\000j\000\231',
	ExcelMAsTAutoshapeSnipRoundRectangle = '\000j\000\232',
	ExcelMAsTAutoshapeSnipOneRectangle = '\000j\000\233',
	ExcelMAsTAutoshapeSnipTwoSameRectangle = '\000j\000\234',
	ExcelMAsTAutoshapeSnipTwoDiagonalRectangle = '\000j\000\235',
	ExcelMAsTAutoshapeFrame = '\000j\000\236',
	ExcelMAsTAutoshapeHalfFrame = '\000j\000\237',
	ExcelMAsTAutoshapeTear = '\000j\000\240',
	ExcelMAsTAutoshapeChord = '\000j\000\241',
	ExcelMAsTAutoshapeCorner = '\000j\000\242',
	ExcelMAsTAutoshapeMathPlus = '\000j\000\243',
	ExcelMAsTAutoshapeMathMinus = '\000j\000\244',
	ExcelMAsTAutoshapeMathMultiply = '\000j\000\245',
	ExcelMAsTAutoshapeMathDivide = '\000j\000\246',
	ExcelMAsTAutoshapeMathEqual = '\000j\000\247',
	ExcelMAsTAutoshapeMathNotEqual = '\000j\000\250',
	ExcelMAsTAutoshapeCornerTabs = '\000j\000\251',
	ExcelMAsTAutoshapeSquareTabs = '\000j\000\252',
	ExcelMAsTAutoshapePlaqueTabs = '\000j\000\253',
	ExcelMAsTAutoshapeGearSix = '\000j\000\254',
	ExcelMAsTAutoshapeGearNine = '\000j\000\255',
	ExcelMAsTAutoshapeFunnel = '\000j\000\256',
	ExcelMAsTAutoshapePieWedge = '\000j\000\257',
	ExcelMAsTAutoshapeLeftCircularArrow = '\000j\000\260',
	ExcelMAsTAutoshapeLeftRightCircularArrow = '\000j\000\261',
	ExcelMAsTAutoshapeSwooshArrow = '\000j\000\262',
	ExcelMAsTAutoshapeCloud = '\000j\000\263',
	ExcelMAsTAutoshapeChartX = '\000j\000\264',
	ExcelMAsTAutoshapeChartStar = '\000j\000\265',
	ExcelMAsTAutoshapeChartPlus = '\000j\000\266',
	ExcelMAsTAutoshapeLineInverse = '\000j\000\267'
};
typedef enum ExcelMAsT ExcelMAsT;

enum ExcelMShp {
	ExcelMShpShapeTypeUnset = '\000\213\377\376',
	ExcelMShpShapeTypeAuto = '\000\214\000\001',
	ExcelMShpShapeTypeCallout = '\000\214\000\002',
	ExcelMShpShapeTypeChart = '\000\214\000\003',
	ExcelMShpShapeTypeComment = '\000\214\000\004',
	ExcelMShpShapeTypeFreeForm = '\000\214\000\005',
	ExcelMShpShapeTypeGroup = '\000\214\000\006',
	ExcelMShpShapeTypeEmbeddedOLEControl = '\000\214\000\007',
	ExcelMShpShapeTypeFormControl = '\000\214\000\010',
	ExcelMShpShapeTypeLine = '\000\214\000\011',
	ExcelMShpShapeTypeLinkedOLEObject = '\000\214\000\012',
	ExcelMShpShapeTypeLinkedPicture = '\000\214\000\013',
	ExcelMShpShapeTypeOLEControl = '\000\214\000\014',
	ExcelMShpShapeTypePicture = '\000\214\000\015',
	ExcelMShpShapeTypePlaceHolder = '\000\214\000\016',
	ExcelMShpShapeTypeWordArt = '\000\214\000\017',
	ExcelMShpShapeTypeMedia = '\000\214\000\020',
	ExcelMShpShapeTypeTextBox = '\000\214\000\021',
	ExcelMShpShapeTypeTable = '\000\214\000\022',
	ExcelMShpShapeTypeCanvas = '\000\214\000\023',
	ExcelMShpShapeTypeDiagram = '\000\214\000\024',
	ExcelMShpShapeTypeInk = '\000\214\000\025',
	ExcelMShpShapeTypeInkComment = '\000\214\000\026',
	ExcelMShpShapeTypeSmartartGraphic = '\000\214\000\027',
	ExcelMShpShapeTypeSlicer = '\000\214\000\030'
};
typedef enum ExcelMShp ExcelMShp;

enum ExcelMCrT {
	ExcelMCrTColorTypeUnset = '\000j\377\376',
	ExcelMCrTRGB = '\000k\000\001',
	ExcelMCrTScheme = '\000k\000\002'
};
typedef enum ExcelMCrT ExcelMCrT;

enum ExcelMPc {
	ExcelMPcPictureColorTypeUnset = '\000\265\377\376',
	ExcelMPcPictureColorAutomatic = '\000\266\000\001',
	ExcelMPcPictureColorGrayScale = '\000\266\000\002',
	ExcelMPcPictureColorBlackAndWhite = '\000\266\000\003',
	ExcelMPcPictureColorWatermark = '\000\266\000\004'
};
typedef enum ExcelMPc ExcelMPc;

enum ExcelMCAt {
	ExcelMCAtAngleUnset = '\000k\377\376',
	ExcelMCAtAngleAutomatic = '\000l\000\001',
	ExcelMCAtAngle30 = '\000l\000\002',
	ExcelMCAtAngle45 = '\000l\000\003',
	ExcelMCAtAngle60 = '\000l\000\004',
	ExcelMCAtAngle90 = '\000l\000\005'
};
typedef enum ExcelMCAt ExcelMCAt;

enum ExcelMCDt {
	ExcelMCDtDropUnset = '\000l\377\376',
	ExcelMCDtDropCustom = '\000m\000\001',
	ExcelMCDtDropTop = '\000m\000\002',
	ExcelMCDtDropCenter = '\000m\000\003',
	ExcelMCDtDropBottom = '\000m\000\004'
};
typedef enum ExcelMCDt ExcelMCDt;

enum ExcelMCot {
	ExcelMCotCalloutUnset = '\000m\377\376',
	ExcelMCotCalloutOne = '\000n\000\001',
	ExcelMCotCalloutTwo = '\000n\000\002',
	ExcelMCotCalloutThree = '\000n\000\003',
	ExcelMCotCalloutFour = '\000n\000\004'
};
typedef enum ExcelMCot ExcelMCot;

enum ExcelTxOr {
	ExcelTxOrTextOrientationUnset = '\000\215\377\376',
	ExcelTxOrHorizontal = '\000\216\000\001',
	ExcelTxOrUpward = '\000\216\000\002',
	ExcelTxOrDownward = '\000\216\000\003',
	ExcelTxOrVerticalEastAsian = '\000\216\000\004',
	ExcelTxOrVertical = '\000\216\000\005',
	ExcelTxOrHorizontalRotatedEastAsian = '\000\216\000\006'
};
typedef enum ExcelTxOr ExcelTxOr;

enum ExcelMSFr {
	ExcelMSFrScaleFromTopLeft = '\000o\000\000',
	ExcelMSFrScaleFromMiddle = '\000o\000\001',
	ExcelMSFrScaleFromBottomRight = '\000o\000\002'
};
typedef enum ExcelMSFr ExcelMSFr;

enum ExcelMPzC {
	ExcelMPzCPresetCameraUnset = '\000\256\377\376',
	ExcelMPzCCameraLegacyObliqueFromTopLeft = '\000\257\000\001',
	ExcelMPzCCameraLegacyObliqueFromTop = '\000\257\000\002',
	ExcelMPzCCameraLegacyObliqueFromTopright = '\000\257\000\003',
	ExcelMPzCCameraLegacyObliqueFromLeft = '\000\257\000\004',
	ExcelMPzCCameraLegacyObliqueFromFront = '\000\257\000\005',
	ExcelMPzCCameraLegacyObliqueFromRight = '\000\257\000\006',
	ExcelMPzCCameraLegacyObliqueFromBottomLeft = '\000\257\000\007',
	ExcelMPzCCameraLegacyObliqueFromBottom = '\000\257\000\010',
	ExcelMPzCCameraLegacyObliqueFromBottomRight = '\000\257\000\011',
	ExcelMPzCCameraLegacyPerspectiveFromTopLeft = '\000\257\000\012',
	ExcelMPzCCameraLegacyPerspectiveFromTop = '\000\257\000\013',
	ExcelMPzCCameraLegacyPerspectiveFromTopRight = '\000\257\000\014',
	ExcelMPzCCameraLegacyPerspectiveFromLeft = '\000\257\000\015',
	ExcelMPzCCameraLegacyPerspectiveFromFront = '\000\257\000\016',
	ExcelMPzCCameraLegacyPerspectiveFromRight = '\000\257\000\017',
	ExcelMPzCCameraLegacyPerspectiveFromBottomLeft = '\000\257\000\020',
	ExcelMPzCCameraLegacyPerspectiveFromBottom = '\000\257\000\021',
	ExcelMPzCCameraLegacyPerspectiveFromBottomRight = '\000\257\000\022',
	ExcelMPzCCameraOrthographic = '\000\257\000\023',
	ExcelMPzCCameraIsometricFromTopUp = '\000\257\000\024',
	ExcelMPzCCameraIsometricFromTopDown = '\000\257\000\025',
	ExcelMPzCCameraIsometricFromBottomUp = '\000\257\000\026',
	ExcelMPzCCameraIsometricFromBottomDown = '\000\257\000\027',
	ExcelMPzCCameraIsometricFromLeftUp = '\000\257\000\030',
	ExcelMPzCCameraIsometricFromLeftDown = '\000\257\000\031',
	ExcelMPzCCameraIsometricFromRightUp = '\000\257\000\032',
	ExcelMPzCCameraIsometricFromRightDown = '\000\257\000\033',
	ExcelMPzCCameraIsometricOffAxis1FromLeft = '\000\257\000\034',
	ExcelMPzCCameraIsometricOffAxis1FromRight = '\000\257\000\035',
	ExcelMPzCCameraIsometricOffAxis1FromTop = '\000\257\000\036',
	ExcelMPzCCameraIsometricOffAxis2FromLeft = '\000\257\000\037',
	ExcelMPzCCameraIsometricOffAxis2FromRight = '\000\257\000 ',
	ExcelMPzCCameraIsometricOffAxis2FromTop = '\000\257\000!',
	ExcelMPzCCameraIsometricOffAxis3FromLeft = '\000\257\000\"',
	ExcelMPzCCameraIsometricOffAxis3FromRight = '\000\257\000#',
	ExcelMPzCCameraIsometricOffAxis3FromBottom = '\000\257\000$',
	ExcelMPzCCameraIsometricOffAxis4FromLeft = '\000\257\000%',
	ExcelMPzCCameraIsometricOffAxis4FromRight = '\000\257\000&',
	ExcelMPzCCameraIsometricOffAxis4FromBottom = '\000\257\000\'',
	ExcelMPzCCameraObliqueFromTopLeft = '\000\257\000(',
	ExcelMPzCCameraObliqueFromTop = '\000\257\000)',
	ExcelMPzCCameraObliqueFromTopRight = '\000\257\000*',
	ExcelMPzCCameraObliqueFromLeft = '\000\257\000+',
	ExcelMPzCCameraObliqueFromRight = '\000\257\000,',
	ExcelMPzCCameraObliqueFromBottomLeft = '\000\257\000-',
	ExcelMPzCCameraObliqueFromBottom = '\000\257\000.',
	ExcelMPzCCameraObliqueFromBottomRight = '\000\257\000/',
	ExcelMPzCCameraPerspectiveFromFront = '\000\257\0000',
	ExcelMPzCCameraPerspectiveFromLeft = '\000\257\0001',
	ExcelMPzCCameraPerspectiveFromRight = '\000\257\0002',
	ExcelMPzCCameraPerspectiveFromAbove = '\000\257\0003',
	ExcelMPzCCameraPerspectiveFromBelow = '\000\257\0004',
	ExcelMPzCCameraPerspectiveFromAboveFacingLeft = '\000\257\0005',
	ExcelMPzCCameraPerspectiveFromAboveFacingRight = '\000\257\0006',
	ExcelMPzCCameraPerspectiveContrastingFacingLeft = '\000\257\0007',
	ExcelMPzCCameraPerspectiveContrastingFacingRight = '\000\257\0008',
	ExcelMPzCCameraPerspectiveHeroicFacingLeft = '\000\257\0009',
	ExcelMPzCCameraPerspectiveHeroicFacingRight = '\000\257\000:',
	ExcelMPzCCameraPerspectiveHeroicExtremeFacingLeft = '\000\257\000;',
	ExcelMPzCCameraPerspectiveHeroicExtremeFacingRight = '\000\257\000<',
	ExcelMPzCCameraPerspectiveRelaxed = '\000\257\000=',
	ExcelMPzCCameraPerspectiveRelaxedModerately = '\000\257\000>'
};
typedef enum ExcelMPzC ExcelMPzC;

enum ExcelMLtT {
	ExcelMLtTLightRigUnset = '\000\257\377\376',
	ExcelMLtTLightRigFlat1 = '\000\260\000\001',
	ExcelMLtTLightRigFlat2 = '\000\260\000\002',
	ExcelMLtTLightRigFlat3 = '\000\260\000\003',
	ExcelMLtTLightRigFlat4 = '\000\260\000\004',
	ExcelMLtTLightRigNormal1 = '\000\260\000\005',
	ExcelMLtTLightRigNormal2 = '\000\260\000\006',
	ExcelMLtTLightRigNormal3 = '\000\260\000\007',
	ExcelMLtTLightRigNormal4 = '\000\260\000\010',
	ExcelMLtTLightRigHarsh1 = '\000\260\000\011',
	ExcelMLtTLightRigHarsh2 = '\000\260\000\012',
	ExcelMLtTLightRigHarsh3 = '\000\260\000\013',
	ExcelMLtTLightRigHarsh4 = '\000\260\000\014',
	ExcelMLtTLightRigThreePoint = '\000\260\000\015',
	ExcelMLtTLightRigBalanced = '\000\260\000\016',
	ExcelMLtTLightRigSoft = '\000\260\000\017',
	ExcelMLtTLightRigHarsh = '\000\260\000\020',
	ExcelMLtTLightRigFlood = '\000\260\000\021',
	ExcelMLtTLightRigContrasting = '\000\260\000\022',
	ExcelMLtTLightRigMorning = '\000\260\000\023',
	ExcelMLtTLightRigSunrise = '\000\260\000\024',
	ExcelMLtTLightRigSunset = '\000\260\000\025',
	ExcelMLtTLightRigChilly = '\000\260\000\026',
	ExcelMLtTLightRigFreezing = '\000\260\000\027',
	ExcelMLtTLightRigFlat = '\000\260\000\030',
	ExcelMLtTLightRigTwoPoint = '\000\260\000\031',
	ExcelMLtTLightRigGlow = '\000\260\000\032',
	ExcelMLtTLightRigBrightRoom = '\000\260\000\033'
};
typedef enum ExcelMLtT ExcelMLtT;

enum ExcelMBlT {
	ExcelMBlTBevelTypeUnset = '\000\260\377\376',
	ExcelMBlTBevelNone = '\000\261\000\001',
	ExcelMBlTBevelRelaxedInset = '\000\261\000\002',
	ExcelMBlTBevelCircle = '\000\261\000\003',
	ExcelMBlTBevelSlope = '\000\261\000\004',
	ExcelMBlTBevelCross = '\000\261\000\005',
	ExcelMBlTBevelAngle = '\000\261\000\006',
	ExcelMBlTBevelSoftRound = '\000\261\000\007',
	ExcelMBlTBevelConvex = '\000\261\000\010',
	ExcelMBlTBevelCoolSlant = '\000\261\000\011',
	ExcelMBlTBevelDivot = '\000\261\000\012',
	ExcelMBlTBevelRiblet = '\000\261\000\013',
	ExcelMBlTBevelHardEdge = '\000\261\000\014',
	ExcelMBlTBevelArtDeco = '\000\261\000\015'
};
typedef enum ExcelMBlT ExcelMBlT;

enum ExcelMSSt {
	ExcelMSStShadowStyleUnset = '\000\261\377\376',
	ExcelMSStShadowStyleInner = '\000\262\000\001',
	ExcelMSStShadowStyleOuter = '\000\262\000\002'
};
typedef enum ExcelMSSt ExcelMSSt;

enum ExcelPpgA {
	ExcelPpgAParagraphAlignmentUnset = '\000\346\377\376',
	ExcelPpgAParagraphAlignLeft = '\000\347\000\000',
	ExcelPpgAParagraphAlignCenter = '\000\347\000\001',
	ExcelPpgAParagraphAlignRight = '\000\347\000\002',
	ExcelPpgAParagraphAlignJustify = '\000\347\000\003',
	ExcelPpgAParagraphAlignDistribute = '\000\347\000\004',
	ExcelPpgAParagraphAlignThai = '\000\347\000\005',
	ExcelPpgAParagraphAlignJustifyLow = '\000\347\000\006'
};
typedef enum ExcelPpgA ExcelPpgA;

enum ExcelMTSt {
	ExcelMTStStrikeUnset = '\000\263\377\376',
	ExcelMTStNoStrike = '\000\264\000\000',
	ExcelMTStSingleStrike = '\000\264\000\001',
	ExcelMTStDoubleStrike = '\000\264\000\002'
};
typedef enum ExcelMTSt ExcelMTSt;

enum ExcelMTCa {
	ExcelMTCaCapsUnset = '\000\264\377\376',
	ExcelMTCaNoCaps = '\000\265\000\000',
	ExcelMTCaSmallCaps = '\000\265\000\001',
	ExcelMTCaAllCaps = '\000\265\000\002'
};
typedef enum ExcelMTCa ExcelMTCa;

enum ExcelMTUl {
	ExcelMTUlUnderlineUnset = '\003\356\377\376',
	ExcelMTUlNoUnderline = '\003\357\000\000',
	ExcelMTUlUnderlineWordsOnly = '\003\357\000\001',
	ExcelMTUlUnderlineSingleLine = '\003\357\000\002',
	ExcelMTUlUnderlineDoubleLine = '\003\357\000\003',
	ExcelMTUlUnderlineHeavyLine = '\003\357\000\004',
	ExcelMTUlUnderlineDottedLine = '\003\357\000\005',
	ExcelMTUlUnderlineHeavyDottedLine = '\003\357\000\006',
	ExcelMTUlUnderlineDashLine = '\003\357\000\007',
	ExcelMTUlUnderlineHeavyDashLine = '\003\357\000\010',
	ExcelMTUlUnderlineLongDashLine = '\003\357\000\011',
	ExcelMTUlUnderlineHeavyLongDashLine = '\003\357\000\012',
	ExcelMTUlUnderlineDotDashLine = '\003\357\000\013',
	ExcelMTUlUnderlineHeavyDotDashLine = '\003\357\000\014',
	ExcelMTUlUnderlineDotDotDashLine = '\003\357\000\015',
	ExcelMTUlUnderlineHeavyDotDotDashLine = '\003\357\000\016',
	ExcelMTUlUnderlineWavyLine = '\003\357\000\017',
	ExcelMTUlUnderlineHeavyWavyLine = '\003\357\000\020',
	ExcelMTUlUnderlineWavyDoubleLine = '\003\357\000\021'
};
typedef enum ExcelMTUl ExcelMTUl;

enum ExcelMTTA {
	ExcelMTTATabUnset = '\000\266\377\376',
	ExcelMTTALeftTab = '\000\267\000\000',
	ExcelMTTACenterTab = '\000\267\000\001',
	ExcelMTTARightTab = '\000\267\000\002',
	ExcelMTTADecimalTab = '\000\267\000\003'
};
typedef enum ExcelMTTA ExcelMTTA;

enum ExcelMTCW {
	ExcelMTCWCharacterWrapUnset = '\000\267\377\376',
	ExcelMTCWNoCharacterWrap = '\000\270\000\000',
	ExcelMTCWStandardCharacterWrap = '\000\270\000\001',
	ExcelMTCWStrictCharacterWrap = '\000\270\000\002',
	ExcelMTCWCustomCharacterWrap = '\000\270\000\003'
};
typedef enum ExcelMTCW ExcelMTCW;

enum ExcelMTFA {
	ExcelMTFAFontAlignmentUnset = '\000\270\377\376',
	ExcelMTFAAutomaticAlignment = '\000\271\000\000',
	ExcelMTFATopAlignment = '\000\271\000\001',
	ExcelMTFACenterAlignment = '\000\271\000\002',
	ExcelMTFABaselineAlignment = '\000\271\000\003',
	ExcelMTFABottomAlignment = '\000\271\000\004'
};
typedef enum ExcelMTFA ExcelMTFA;

enum ExcelPAtS {
	ExcelPAtSAutoSizeUnset = '\000\344\377\376',
	ExcelPAtSAutoSizeNone = '\000\345\000\000',
	ExcelPAtSShapeToFitText = '\000\345\000\001',
	ExcelPAtSTextToFitShape = '\000\345\000\002'
};
typedef enum ExcelPAtS ExcelPAtS;

enum ExcelMPFo {
	ExcelMPFoPathTypeUnset = '\000\272\377\376',
	ExcelMPFoNoPathType = '\000\273\000\000',
	ExcelMPFoPathType1 = '\000\273\000\001',
	ExcelMPFoPathType2 = '\000\273\000\002',
	ExcelMPFoPathType3 = '\000\273\000\003',
	ExcelMPFoPathType4 = '\000\273\000\004'
};
typedef enum ExcelMPFo ExcelMPFo;

enum ExcelMWFo {
	ExcelMWFoWarpFormatUnset = '\000\273\377\376',
	ExcelMWFoWarpFormat1 = '\000\274\000\000',
	ExcelMWFoWarpFormat2 = '\000\274\000\001',
	ExcelMWFoWarpFormat3 = '\000\274\000\002',
	ExcelMWFoWarpFormat4 = '\000\274\000\003',
	ExcelMWFoWarpFormat5 = '\000\274\000\004',
	ExcelMWFoWarpFormat6 = '\000\274\000\005',
	ExcelMWFoWarpFormat7 = '\000\274\000\006',
	ExcelMWFoWarpFormat8 = '\000\274\000\007',
	ExcelMWFoWarpFormat9 = '\000\274\000\010',
	ExcelMWFoWarpFormat10 = '\000\274\000\011',
	ExcelMWFoWarpFormat11 = '\000\274\000\012',
	ExcelMWFoWarpFormat12 = '\000\274\000\013',
	ExcelMWFoWarpFormat13 = '\000\274\000\014',
	ExcelMWFoWarpFormat14 = '\000\274\000\015',
	ExcelMWFoWarpFormat15 = '\000\274\000\016',
	ExcelMWFoWarpFormat16 = '\000\274\000\017',
	ExcelMWFoWarpFormat17 = '\000\274\000\020',
	ExcelMWFoWarpFormat18 = '\000\274\000\021',
	ExcelMWFoWarpFormat19 = '\000\274\000\022',
	ExcelMWFoWarpFormat20 = '\000\274\000\023',
	ExcelMWFoWarpFormat21 = '\000\274\000\024',
	ExcelMWFoWarpFormat22 = '\000\274\000\025',
	ExcelMWFoWarpFormat23 = '\000\274\000\026',
	ExcelMWFoWarpFormat24 = '\000\274\000\027',
	ExcelMWFoWarpFormat25 = '\000\274\000\030',
	ExcelMWFoWarpFormat26 = '\000\274\000\031',
	ExcelMWFoWarpFormat27 = '\000\274\000\032',
	ExcelMWFoWarpFormat28 = '\000\274\000\033',
	ExcelMWFoWarpFormat29 = '\000\274\000\034',
	ExcelMWFoWarpFormat30 = '\000\274\000\035',
	ExcelMWFoWarpFormat31 = '\000\274\000\036',
	ExcelMWFoWarpFormat32 = '\000\274\000\037',
	ExcelMWFoWarpFormat33 = '\000\274\000 ',
	ExcelMWFoWarpFormat34 = '\000\274\000!',
	ExcelMWFoWarpFormat35 = '\000\274\000\"',
	ExcelMWFoWarpFormat36 = '\000\274\000#'
};
typedef enum ExcelMWFo ExcelMWFo;

enum ExcelPCgC {
	ExcelPCgCCaseSentence = '\000\344\000\001',
	ExcelPCgCCaseLower = '\000\344\000\002',
	ExcelPCgCCaseUpper = '\000\344\000\003',
	ExcelPCgCCaseTitle = '\000\344\000\004',
	ExcelPCgCCaseToggle = '\000\344\000\005'
};
typedef enum ExcelPCgC ExcelPCgC;

enum ExcelMDTF {
	ExcelMDTFDateTimeFormatUnset = '\000\275\377\376',
	ExcelMDTFDateTimeFormatMdyy = '\000\276\000\001',
	ExcelMDTFDateTimeFormatDdddMMMMddyyyy = '\000\276\000\002',
	ExcelMDTFDateTimeFormatDMMMMyyyy = '\000\276\000\003',
	ExcelMDTFDateTimeFormatMMMMdyyyy = '\000\276\000\004',
	ExcelMDTFDateTimeFormatDMMMyy = '\000\276\000\005',
	ExcelMDTFDateTimeFormatMMMMyy = '\000\276\000\006',
	ExcelMDTFDateTimeFormatMMyy = '\000\276\000\007',
	ExcelMDTFDateTimeFormatMMddyyHmm = '\000\276\000\010',
	ExcelMDTFDateTimeFormatMMddyyhmmAMPM = '\000\276\000\011',
	ExcelMDTFDateTimeFormatHmm = '\000\276\000\012',
	ExcelMDTFDateTimeFormatHmmss = '\000\276\000\013',
	ExcelMDTFDateTimeFormatHmmAMPM = '\000\276\000\014',
	ExcelMDTFDateTimeFormatHmmssAMPM = '\000\276\000\015',
	ExcelMDTFDateTimeFormatFigureOut = '\000\276\000\016'
};
typedef enum ExcelMDTF ExcelMDTF;

enum ExcelMSET {
	ExcelMSETSoftEdgeUnset = '\000\277\377\376',
	ExcelMSETNoSoftEdge = '\000\300\000\000',
	ExcelMSETSoftEdgeType1 = '\000\300\000\001',
	ExcelMSETSoftEdgeType2 = '\000\300\000\002',
	ExcelMSETSoftEdgeType3 = '\000\300\000\003',
	ExcelMSETSoftEdgeType4 = '\000\300\000\004',
	ExcelMSETSoftEdgeType5 = '\000\300\000\005',
	ExcelMSETSoftEdgeType6 = '\000\300\000\006'
};
typedef enum ExcelMSET ExcelMSET;

enum ExcelMCSI {
	ExcelMCSIFirstDarkSchemeColor = '\000\301\000\001',
	ExcelMCSIFirstLightSchemeColor = '\000\301\000\002',
	ExcelMCSISecondDarkSchemeColor = '\000\301\000\003',
	ExcelMCSISecondLightSchemeColor = '\000\301\000\004',
	ExcelMCSIFirstAccentSchemeColor = '\000\301\000\005',
	ExcelMCSISecondAccentSchemeColor = '\000\301\000\006',
	ExcelMCSIThirdAccentSchemeColor = '\000\301\000\007',
	ExcelMCSIFourthAccentSchemeColor = '\000\301\000\010',
	ExcelMCSIFifthAccentSchemeColor = '\000\301\000\011',
	ExcelMCSISixthAccentSchemeColor = '\000\301\000\012',
	ExcelMCSIHyperlinkSchemeColor = '\000\301\000\013',
	ExcelMCSIFollowedHyperlinkSchemeColor = '\000\301\000\014'
};
typedef enum ExcelMCSI ExcelMCSI;

enum ExcelMCoI {
	ExcelMCoIThemeColorUnset = '\000\301\377\376',
	ExcelMCoINoThemeColor = '\000\302\000\000',
	ExcelMCoIFirstDarkThemeColor = '\000\302\000\001',
	ExcelMCoIFirstLightThemeColor = '\000\302\000\002',
	ExcelMCoISecondDarkThemeColor = '\000\302\000\003',
	ExcelMCoISecondLightThemeColor = '\000\302\000\004',
	ExcelMCoIFirstAccentThemeColor = '\000\302\000\005',
	ExcelMCoISecondAccentThemeColor = '\000\302\000\006',
	ExcelMCoIThirdAccentThemeColor = '\000\302\000\007',
	ExcelMCoIFourthAccentThemeColor = '\000\302\000\010',
	ExcelMCoIFifthAccentThemeColor = '\000\302\000\011',
	ExcelMCoISixthAccentThemeColor = '\000\302\000\012',
	ExcelMCoIHyperlinkThemeColor = '\000\302\000\013',
	ExcelMCoIFollowedHyperlinkThemeColor = '\000\302\000\014',
	ExcelMCoIFirstTextThemeColor = '\000\302\000\015',
	ExcelMCoIFirstBackgroundThemeColor = '\000\302\000\016',
	ExcelMCoISecondTextThemeColor = '\000\302\000\017',
	ExcelMCoISecondBackgroundThemeColor = '\000\302\000\020'
};
typedef enum ExcelMCoI ExcelMCoI;

enum ExcelMFLI {
	ExcelMFLIThemeFontLatin = '\000\303\000\001',
	ExcelMFLIThemeFontComplexScript = '\000\303\000\002',
	ExcelMFLIThemeFontHighAnsi = '\000\303\000\003',
	ExcelMFLIThemeFontEastAsian = '\000\303\000\004'
};
typedef enum ExcelMFLI ExcelMFLI;

enum ExcelMSSI {
	ExcelMSSIShapeStyleUnset = '\000\303\377\376',
	ExcelMSSIShapeNotAPreset = '\000\304\000\000',
	ExcelMSSIShapePreset1 = '\000\304\000\001',
	ExcelMSSIShapePreset2 = '\000\304\000\002',
	ExcelMSSIShapePreset3 = '\000\304\000\003',
	ExcelMSSIShapePreset4 = '\000\304\000\004',
	ExcelMSSIShapePreset5 = '\000\304\000\005',
	ExcelMSSIShapePreset6 = '\000\304\000\006',
	ExcelMSSIShapePreset7 = '\000\304\000\007',
	ExcelMSSIShapePreset8 = '\000\304\000\010',
	ExcelMSSIShapePreset9 = '\000\304\000\011',
	ExcelMSSIShapePreset10 = '\000\304\000\012',
	ExcelMSSIShapePreset11 = '\000\304\000\013',
	ExcelMSSIShapePreset12 = '\000\304\000\014',
	ExcelMSSIShapePreset13 = '\000\304\000\015',
	ExcelMSSIShapePreset14 = '\000\304\000\016',
	ExcelMSSIShapePreset15 = '\000\304\000\017',
	ExcelMSSIShapePreset16 = '\000\304\000\020',
	ExcelMSSIShapePreset17 = '\000\304\000\021',
	ExcelMSSIShapePreset18 = '\000\304\000\022',
	ExcelMSSIShapePreset19 = '\000\304\000\023',
	ExcelMSSIShapePreset20 = '\000\304\000\024',
	ExcelMSSIShapePreset21 = '\000\304\000\025',
	ExcelMSSIShapePreset22 = '\000\304\000\026',
	ExcelMSSIShapePreset23 = '\000\304\000\027',
	ExcelMSSIShapePreset24 = '\000\304\000\030',
	ExcelMSSIShapePreset25 = '\000\304\000\031',
	ExcelMSSIShapePreset26 = '\000\304\000\032',
	ExcelMSSIShapePreset27 = '\000\304\000\033',
	ExcelMSSIShapePreset28 = '\000\304\000\034',
	ExcelMSSIShapePreset29 = '\000\304\000\035',
	ExcelMSSIShapePreset30 = '\000\304\000\036',
	ExcelMSSIShapePreset31 = '\000\304\000\037',
	ExcelMSSIShapePreset32 = '\000\304\000 ',
	ExcelMSSIShapePreset33 = '\000\304\000!',
	ExcelMSSIShapePreset34 = '\000\304\000\"',
	ExcelMSSIShapePreset35 = '\000\304\000#',
	ExcelMSSIShapePreset36 = '\000\304\000$',
	ExcelMSSIShapePreset37 = '\000\304\000%',
	ExcelMSSIShapePreset38 = '\000\304\000&',
	ExcelMSSIShapePreset39 = '\000\304\000\'',
	ExcelMSSIShapePreset40 = '\000\304\000(',
	ExcelMSSIShapePreset41 = '\000\304\000)',
	ExcelMSSIShapePreset42 = '\000\304\000*',
	ExcelMSSILinePreset1 = '\000\304\'\021',
	ExcelMSSILinePreset2 = '\000\304\'\022',
	ExcelMSSILinePreset3 = '\000\304\'\023',
	ExcelMSSILinePreset4 = '\000\304\'\024',
	ExcelMSSILinePreset5 = '\000\304\'\025',
	ExcelMSSILinePreset6 = '\000\304\'\026',
	ExcelMSSILinePreset7 = '\000\304\'\027',
	ExcelMSSILinePreset8 = '\000\304\'\030',
	ExcelMSSILinePreset9 = '\000\304\'\031',
	ExcelMSSILinePreset10 = '\000\304\'\032',
	ExcelMSSILinePreset11 = '\000\304\'\033',
	ExcelMSSILinePreset12 = '\000\304\'\034',
	ExcelMSSILinePreset13 = '\000\304\'\035',
	ExcelMSSILinePreset14 = '\000\304\'\036',
	ExcelMSSILinePreset15 = '\000\304\'\037',
	ExcelMSSILinePreset16 = '\000\304\' ',
	ExcelMSSILinePreset17 = '\000\304\'!',
	ExcelMSSILinePreset18 = '\000\304\'\"',
	ExcelMSSILinePreset19 = '\000\304\'#',
	ExcelMSSILinePreset20 = '\000\304\'$',
	ExcelMSSILinePreset21 = '\000\304\'%'
};
typedef enum ExcelMSSI ExcelMSSI;

enum ExcelMBSI {
	ExcelMBSIBackgroundUnset = '\000\304\377\376',
	ExcelMBSIBackgroundNotAPreset = '\000\305\000\000',
	ExcelMBSIBackgroundPreset1 = '\000\305\000\001',
	ExcelMBSIBackgroundPreset2 = '\000\305\000\002',
	ExcelMBSIBackgroundPreset3 = '\000\305\000\003',
	ExcelMBSIBackgroundPreset4 = '\000\305\000\004',
	ExcelMBSIBackgroundPreset5 = '\000\305\000\005',
	ExcelMBSIBackgroundPreset6 = '\000\305\000\006',
	ExcelMBSIBackgroundPreset7 = '\000\305\000\007',
	ExcelMBSIBackgroundPreset8 = '\000\305\000\010',
	ExcelMBSIBackgroundPreset9 = '\000\305\000\011',
	ExcelMBSIBackgroundPreset10 = '\000\305\000\012',
	ExcelMBSIBackgroundPreset11 = '\000\305\000\013',
	ExcelMBSIBackgroundPreset12 = '\000\305\000\014'
};
typedef enum ExcelMBSI ExcelMBSI;

enum ExcelPDrT {
	ExcelPDrTTextDirectionUnset = '\000\352\377\376',
	ExcelPDrTLeftToRight = '\000\353\000\001',
	ExcelPDrTRightToLeft = '\000\353\000\002'
};
typedef enum ExcelPDrT ExcelPDrT;

enum ExcelPBtT {
	ExcelPBtTBulletTypeUnset = '\000\347\377\376',
	ExcelPBtTBulletTypeNone = '\000\350\000\000',
	ExcelPBtTBulletTypeUnnumbered = '\000\350\000\001',
	ExcelPBtTBulletTypeNumbered = '\000\350\000\002',
	ExcelPBtTPictureBulletType = '\000\350\000\003'
};
typedef enum ExcelPBtT ExcelPBtT;

enum ExcelPBtS {
	ExcelPBtSNumberedBulletStyleUnset = '\000\350\377\376',
	ExcelPBtSNumberedBulletStyleAlphaLowercasePeriod = '\000\351\000\000',
	ExcelPBtSNumberedBulletStyleAlphaUppercasePeriod = '\000\351\000\001',
	ExcelPBtSNumberedBulletStyleArabicRightParen = '\000\351\000\002',
	ExcelPBtSNumberedBulletStyleArabicPeriod = '\000\351\000\003',
	ExcelPBtSNumberedBulletStyleRomanLowercaseParenBoth = '\000\351\000\004',
	ExcelPBtSNumberedBulletStyleRomanLowercaseParenRight = '\000\351\000\005',
	ExcelPBtSNumberedBulletStyleRomanLowercasePeriod = '\000\351\000\006',
	ExcelPBtSNumberedBulletStyleRomanUppercasePeriod = '\000\351\000\007',
	ExcelPBtSNumberedBulletStyleAlphaLowercaseParenBoth = '\000\351\000\010',
	ExcelPBtSNumberedBulletStyleAlphaLowercaseParenRight = '\000\351\000\011',
	ExcelPBtSNumberedBulletStyleAlphaUppercaseParenBoth = '\000\351\000\012',
	ExcelPBtSNumberedBulletStyleAlphaUppercaseParenRight = '\000\351\000\013',
	ExcelPBtSNumberedBulletStyleArabicParenBoth = '\000\351\000\014',
	ExcelPBtSNumberedBulletStyleArabicPlain = '\000\351\000\015',
	ExcelPBtSNumberedBulletStyleRomanUppercaseParenBoth = '\000\351\000\016',
	ExcelPBtSNumberedBulletStyleRomanUppercaseParenRight = '\000\351\000\017',
	ExcelPBtSNumberedBulletStyleSimplifiedChinesePlain = '\000\351\000\020',
	ExcelPBtSNumberedBulletStyleSimplifiedChinesePeriod = '\000\351\000\021',
	ExcelPBtSNumberedBulletStyleCircleNumberPlain = '\000\351\000\022',
	ExcelPBtSNumberedBulletStyleCircleNumberWhitePlain = '\000\351\000\023',
	ExcelPBtSNumberedBulletStyleCircleNumberBlackPlain = '\000\351\000\024',
	ExcelPBtSNumberedBulletStyleTraditionalChinesePlain = '\000\351\000\025',
	ExcelPBtSNumberedBulletStyleTraditionalChinesePeriod = '\000\351\000\026',
	ExcelPBtSNumberedBulletStyleArabicAlphaDash = '\000\351\000\027',
	ExcelPBtSNumberedBulletStyleArabicAbjadDash = '\000\351\000\030',
	ExcelPBtSNumberedBulletStyleHebrewAlphaDash = '\000\351\000\031',
	ExcelPBtSNumberedBulletStyleKanjiKoreanPlain = '\000\351\000\032',
	ExcelPBtSNumberedBulletStyleKanjiKoreanPeriod = '\000\351\000\033',
	ExcelPBtSNumberedBulletStyleArabicDBPlain = '\000\351\000\034',
	ExcelPBtSNumberedBulletStyleArabicDBPeriod = '\000\351\000\035',
	ExcelPBtSNumberedBulletStyleThaiAlphaPeriod = '\000\351\000\036',
	ExcelPBtSNumberedBulletStyleThaiAlphaParenRight = '\000\351\000\037',
	ExcelPBtSNumberedBulletStyleThaiAlphaParenBoth = '\000\351\000 ',
	ExcelPBtSNumberedBulletStyleThaiNumberPeriod = '\000\351\000!',
	ExcelPBtSNumberedBulletStyleThaiNumberParenRight = '\000\351\000\"',
	ExcelPBtSNumberedBulletStyleThaiParenBoth = '\000\351\000#',
	ExcelPBtSNumberedBulletStyleHindiAlphaPeriod = '\000\351\000$',
	ExcelPBtSNumberedBulletStyleHindiNumberPeriod = '\000\351\000%',
	ExcelPBtSNumberedBulletStyleKanjiSimpifiedChineseDBPeriod = '\000\351\000&',
	ExcelPBtSNumberedBulletStyleHindiNumberParenRight = '\000\351\000\'',
	ExcelPBtSNumberedBulletStyleHindiAlpha1Period = '\000\351\000('
};
typedef enum ExcelPBtS ExcelPBtS;

enum ExcelPTSp {
	ExcelPTSpTabstopUnset = '\000\364\377\376',
	ExcelPTSpTabstopLeft = '\000\365\000\001',
	ExcelPTSpTabstopCenter = '\000\365\000\002',
	ExcelPTSpTabstopRight = '\000\365\000\003',
	ExcelPTSpTabstopDecimal = '\000\365\000\004'
};
typedef enum ExcelPTSp ExcelPTSp;

enum ExcelMRfT {
	ExcelMRfTReflectionUnset = '\003\351\377\376',
	ExcelMRfTReflectionTypeNone = '\003\352\000\000',
	ExcelMRfTReflectionType1 = '\003\352\000\001',
	ExcelMRfTReflectionType2 = '\003\352\000\002',
	ExcelMRfTReflectionType3 = '\003\352\000\003',
	ExcelMRfTReflectionType4 = '\003\352\000\004',
	ExcelMRfTReflectionType5 = '\003\352\000\005',
	ExcelMRfTReflectionType6 = '\003\352\000\006',
	ExcelMRfTReflectionType7 = '\003\352\000\007',
	ExcelMRfTReflectionType8 = '\003\352\000\010',
	ExcelMRfTReflectionType9 = '\003\352\000\011'
};
typedef enum ExcelMRfT ExcelMRfT;

enum ExcelMTtA {
	ExcelMTtATextureUnset = '\003\352\377\376',
	ExcelMTtATextureTopLeft = '\003\353\000\000',
	ExcelMTtATextureTop = '\003\353\000\001',
	ExcelMTtATextureTopRight = '\003\353\000\002',
	ExcelMTtATextureLeft = '\003\353\000\003',
	ExcelMTtATextureCenter = '\003\353\000\004',
	ExcelMTtATextureRight = '\003\353\000\005',
	ExcelMTtATextureBottomLeft = '\003\353\000\006',
	ExcelMTtATextureBotton = '\003\353\000\007',
	ExcelMTtATextureBottomRight = '\003\353\000\010'
};
typedef enum ExcelMTtA ExcelMTtA;

enum ExcelPBlA {
	ExcelPBlATextBaselineAlignmentUnset = '\003\353\377\376',
	ExcelPBlATextBaselineAlignBaseline = '\003\354\000\001',
	ExcelPBlATextBaselineAlignTop = '\003\354\000\002',
	ExcelPBlATextBaselineAlignCenter = '\003\354\000\003',
	ExcelPBlATextBaselineAlignEastAsian50 = '\003\354\000\004',
	ExcelPBlATextBaselineAlignAutomatic = '\003\354\000\005'
};
typedef enum ExcelPBlA ExcelPBlA;

enum ExcelMCbF {
	ExcelMCbFClipboardFormatUnset = '\003\354\377\376',
	ExcelMCbFNativeClipboardFormat = '\003\355\000\001',
	ExcelMCbFHTMlClipboardFormat = '\003\355\000\002',
	ExcelMCbFRTFClipboardFormat = '\003\355\000\003',
	ExcelMCbFPlainTextClipboardFormat = '\003\355\000\004'
};
typedef enum ExcelMCbF ExcelMCbF;

enum ExcelMTiP {
	ExcelMTiPInsertBefore = '\003\356\000\000',
	ExcelMTiPInsertAfter = '\003\356\000\001'
};
typedef enum ExcelMTiP ExcelMTiP;

enum ExcelMPiT {
	ExcelMPiTSaveAsDefault = '\003\362\377\376',
	ExcelMPiTSaveAsPNGFile = '\003\363\000\000',
	ExcelMPiTSaveAsBMPFile = '\003\363\000\001',
	ExcelMPiTSaveAsGIFFile = '\003\363\000\002',
	ExcelMPiTSaveAsJPGFile = '\003\363\000\003',
	ExcelMPiTSaveAsPDFFile = '\003\363\000\004'
};
typedef enum ExcelMPiT ExcelMPiT;

enum ExcelMPeT {
	ExcelMPeTNoEffect = '\003\364\000\000',
	ExcelMPeTEffectBackgroundRemoval = '\003\364\000\001',
	ExcelMPeTEffectBlur = '\003\364\000\002',
	ExcelMPeTEffectBrightnessContrast = '\003\364\000\003',
	ExcelMPeTEffectCement = '\003\364\000\004',
	ExcelMPeTEffectCrisscrossEtching = '\003\364\000\005',
	ExcelMPeTEffectChalkSketch = '\003\364\000\006',
	ExcelMPeTEffectColorTemperature = '\003\364\000\007',
	ExcelMPeTEffectCutout = '\003\364\000\010',
	ExcelMPeTEffectFilmGrain = '\003\364\000\011',
	ExcelMPeTEffectGlass = '\003\364\000\012',
	ExcelMPeTEffectGlowDiffused = '\003\364\000\013',
	ExcelMPeTEffectGlowEdges = '\003\364\000\014',
	ExcelMPeTEffectLightScreen = '\003\364\000\015',
	ExcelMPeTEffectLineDrawing = '\003\364\000\016',
	ExcelMPeTEffectMarker = '\003\364\000\017',
	ExcelMPeTEffectMosiaicBubbles = '\003\364\000\020',
	ExcelMPeTEffectPaintBrush = '\003\364\000\021',
	ExcelMPeTEffectPaintStrokes = '\003\364\000\022',
	ExcelMPeTEffectPastelsSmooth = '\003\364\000\023',
	ExcelMPeTEffectPencilGrayscale = '\003\364\000\024',
	ExcelMPeTEffectPencilSketch = '\003\364\000\025',
	ExcelMPeTEffectPhotocopy = '\003\364\000\026',
	ExcelMPeTEffectPlasticWrap = '\003\364\000\027',
	ExcelMPeTEffectSaturation = '\003\364\000\030',
	ExcelMPeTEffectSharpenSoften = '\003\364\000\031',
	ExcelMPeTEffectTexturizer = '\003\364\000\032',
	ExcelMPeTEffectWatercolorSponge = '\003\364\000\033'
};
typedef enum ExcelMPeT ExcelMPeT;

enum ExcelMSgT {
	ExcelMSgTLine = '\000\217\000\000',
	ExcelMSgTCurve = '\000\217\000\001'
};
typedef enum ExcelMSgT ExcelMSgT;

enum ExcelMEdT {
	ExcelMEdTAuto = '\000\220\000\000',
	ExcelMEdTCorner = '\000\220\000\001',
	ExcelMEdTSmooth = '\000\220\000\002',
	ExcelMEdTSymmetric = '\000\220\000\003'
};
typedef enum ExcelMEdT ExcelMEdT;

enum ExcelSANP {
	ExcelSANPDefaultNodePosition = '\003\365\000\001',
	ExcelSANPAfterNode = '\003\365\000\002',
	ExcelSANPBeforeNode = '\003\365\000\003',
	ExcelSANPAboveNode = '\003\365\000\004',
	ExcelSANPBelowNode = '\003\365\000\005'
};
typedef enum ExcelSANP ExcelSANP;

enum ExcelSANT {
	ExcelSANTDefaultNode = '\003\366\000\001',
	ExcelSANTAssistantNode = '\003\366\000\002'
};
typedef enum ExcelSANT ExcelSANT;

enum ExcelMOCT {
	ExcelMOCTOrgChartLayoutUnset = '\003\366\377\376',
	ExcelMOCTOrgChartLayoutStandard = '\003\367\000\001',
	ExcelMOCTOrgChartLayoutBothHanging = '\003\367\000\002',
	ExcelMOCTOrgChartLayoutLeftHanging = '\003\367\000\003',
	ExcelMOCTOrgChartLayoutRightHanging = '\003\367\000\004',
	ExcelMOCTOrgChartLayoutDefault = '\003\367\000\005'
};
typedef enum ExcelMOCT ExcelMOCT;

enum ExcelMAlC {
	ExcelMAlCAlignLefts = '\000\000\000\000',
	ExcelMAlCAlignCenters = '\000\000\000\001',
	ExcelMAlCAlignRights = '\000\000\000\002',
	ExcelMAlCAlignTops = '\000\000\000\003',
	ExcelMAlCAlignMiddles = '\000\000\000\004',
	ExcelMAlCAlignBottoms = '\000\000\000\005'
};
typedef enum ExcelMAlC ExcelMAlC;

enum ExcelMDsC {
	ExcelMDsCDistributeHorizontally = '\000\000\000\000',
	ExcelMDsCDistributeVertically = '\000\000\000\001'
};
typedef enum ExcelMDsC ExcelMDsC;

enum ExcelMOrT {
	ExcelMOrTOrientationUnset = '\000\214\377\376',
	ExcelMOrTHorizontalOrientation = '\000\215\000\001',
	ExcelMOrTVerticalOrientation = '\000\215\000\002'
};
typedef enum ExcelMOrT ExcelMOrT;

enum ExcelMZoC {
	ExcelMZoCBringShapeToFront = '\000p\000\000',
	ExcelMZoCSendShapeToBack = '\000p\000\001',
	ExcelMZoCBringShapeForward = '\000p\000\002',
	ExcelMZoCSendShapeBackward = '\000p\000\003',
	ExcelMZoCBringShapeInFrontOfText = '\000p\000\004',
	ExcelMZoCSendShapeBehindText = '\000p\000\005'
};
typedef enum ExcelMZoC ExcelMZoC;

enum ExcelMFlC {
	ExcelMFlCFlipHorizontal = '\000q\000\000',
	ExcelMFlCFlipVertical = '\000q\000\001'
};
typedef enum ExcelMFlC ExcelMFlC;

enum ExcelMTri {
	ExcelMTriTrue = '\000\240\377\377',
	ExcelMTriFalse = '\000\241\000\000',
	ExcelMTriCTrue = '\000\241\000\001',
	ExcelMTriToggle = '\000\240\377\375',
	ExcelMTriTriStateUnset = '\000\240\377\376'
};
typedef enum ExcelMTri ExcelMTri;

enum ExcelMBW {
	ExcelMBWBlackAndWhiteUnset = '\000\254\377\376',
	ExcelMBWBlackAndWhiteModeAutomatic = '\000\255\000\001',
	ExcelMBWBlackAndWhiteModeGrayScale = '\000\255\000\002',
	ExcelMBWBlackAndWhiteModeLightGrayScale = '\000\255\000\003',
	ExcelMBWBlackAndWhiteModeInverseGrayScale = '\000\255\000\004',
	ExcelMBWBlackAndWhiteModeGrayOutline = '\000\255\000\005',
	ExcelMBWBlackAndWhiteModeBlackTextAndLine = '\000\255\000\006',
	ExcelMBWBlackAndWhiteModeHighContrast = '\000\255\000\007',
	ExcelMBWBlackAndWhiteModeBlack = '\000\255\000\010',
	ExcelMBWBlackAndWhiteModeWhite = '\000\255\000\011',
	ExcelMBWBlackAndWhiteModeDontShow = '\000\255\000\012'
};
typedef enum ExcelMBW ExcelMBW;

enum ExcelMBPS {
	ExcelMBPSBarLeft = '\000r\000\000',
	ExcelMBPSBarTop = '\000r\000\001',
	ExcelMBPSBarRight = '\000r\000\002',
	ExcelMBPSBarBottom = '\000r\000\003',
	ExcelMBPSBarFloating = '\000r\000\004',
	ExcelMBPSBarPopUp = '\000r\000\005',
	ExcelMBPSBarMenu = '\000r\000\006'
};
typedef enum ExcelMBPS ExcelMBPS;

enum ExcelMBPt {
	ExcelMBPtNoProtection = '\000s\000\000',
	ExcelMBPtNoCustomize = '\000s\000\001',
	ExcelMBPtNoResize = '\000s\000\002',
	ExcelMBPtNoMove = '\000s\000\004',
	ExcelMBPtNoChangeVisible = '\000s\000\010',
	ExcelMBPtNoChangeDock = '\000s\000\020',
	ExcelMBPtNoVerticalDock = '\000s\000 ',
	ExcelMBPtNoHorizontalDock = '\000s\000@'
};
typedef enum ExcelMBPt ExcelMBPt;

enum ExcelMBTY {
	ExcelMBTYNormalCommandBar = '\000t\000\000',
	ExcelMBTYMenubarCommandBar = '\000t\000\001',
	ExcelMBTYPopupCommandBar = '\000t\000\002'
};
typedef enum ExcelMBTY ExcelMBTY;

enum ExcelMCLT {
	ExcelMCLTControlCustom = '\000u\000\000',
	ExcelMCLTControlButton = '\000u\000\001',
	ExcelMCLTControlEdit = '\000u\000\002',
	ExcelMCLTControlDropDown = '\000u\000\003',
	ExcelMCLTControlCombobox = '\000u\000\004',
	ExcelMCLTButtonDropDown = '\000u\000\005',
	ExcelMCLTSplitDropDown = '\000u\000\006',
	ExcelMCLTOCXDropDown = '\000u\000\007',
	ExcelMCLTGenericDropDown = '\000u\000\010',
	ExcelMCLTGraphicDropDown = '\000u\000\011',
	ExcelMCLTControlPopup = '\000u\000\012',
	ExcelMCLTGraphicPopup = '\000u\000\013',
	ExcelMCLTButtonPopup = '\000u\000\014',
	ExcelMCLTSplitButtonPopup = '\000u\000\015',
	ExcelMCLTSplitButtonMRUPopup = '\000u\000\016',
	ExcelMCLTControlLabel = '\000u\000\017',
	ExcelMCLTExpandingGrid = '\000u\000\020',
	ExcelMCLTSplitExpandingGrid = '\000u\000\021',
	ExcelMCLTControlGrid = '\000u\000\022',
	ExcelMCLTControlGauge = '\000u\000\023',
	ExcelMCLTGraphicCombobox = '\000u\000\024',
	ExcelMCLTControlPane = '\000u\000\025',
	ExcelMCLTActiveX = '\000u\000\026',
	ExcelMCLTControlGroup = '\000u\000\027',
	ExcelMCLTControlTab = '\000u\000\030',
	ExcelMCLTControlSpinner = '\000u\000\031'
};
typedef enum ExcelMCLT ExcelMCLT;

enum ExcelMBns {
	ExcelMBnsButtonStateUp = '\000v\000\000',
	ExcelMBnsButtonStateDown = '\000u\377\377',
	ExcelMBnsButtonStateUnset = '\000v\000\002'
};
typedef enum ExcelMBns ExcelMBns;

enum ExcelMcOu {
	ExcelMcOuNeither = '\000w\000\000',
	ExcelMcOuServer = '\000w\000\001',
	ExcelMcOuClient = '\000w\000\002',
	ExcelMcOuBoth = '\000w\000\003'
};
typedef enum ExcelMcOu ExcelMcOu;

enum ExcelMBTs {
	ExcelMBTsButtonAutomatic = '\000x\000\000',
	ExcelMBTsButtonIcon = '\000x\000\001',
	ExcelMBTsButtonCaption = '\000x\000\002',
	ExcelMBTsButtonIconAndCaption = '\000x\000\003'
};
typedef enum ExcelMBTs ExcelMBTs;

enum ExcelMXcb {
	ExcelMXcbComboboxStyleNormal = '\000y\000\000',
	ExcelMXcbComboboxStyleLabel = '\000y\000\001'
};
typedef enum ExcelMXcb ExcelMXcb;

enum ExcelMMNA {
	ExcelMMNANone = '\000{\000\000',
	ExcelMMNARandom = '\000{\000\001',
	ExcelMMNAUnfold = '\000{\000\002',
	ExcelMMNASlide = '\000{\000\003'
};
typedef enum ExcelMMNA ExcelMMNA;

enum ExcelMHlT {
	ExcelMHlTHyperlinkTypeTextRange = '\000\226\000\000',
	ExcelMHlTHyperlinkTypeShape = '\000\226\000\001',
	ExcelMHlTHyperlinkTypeInlineShape = '\000\226\000\002'
};
typedef enum ExcelMHlT ExcelMHlT;

enum ExcelMXiM {
	ExcelMXiMAppendString = '\000\256\000\000',
	ExcelMXiMPostString = '\000\256\000\001'
};
typedef enum ExcelMXiM ExcelMXiM;

enum ExcelMANT {
	ExcelMANTIdle = '\000|\000\001',
	ExcelMANTGreeting = '\000|\000\002',
	ExcelMANTGoodbye = '\000|\000\003',
	ExcelMANTBeginSpeaking = '\000|\000\004',
	ExcelMANTCharacterSuccessMajor = '\000|\000\006',
	ExcelMANTGetAttentionMajor = '\000|\000\013',
	ExcelMANTGetAttentionMinor = '\000|\000\014',
	ExcelMANTSearching = '\000|\000\015',
	ExcelMANTPrinting = '\000|\000\022',
	ExcelMANTGestureRight = '\000|\000\023',
	ExcelMANTWritingNotingSomething = '\000|\000\026',
	ExcelMANTWorkingAtSomething = '\000|\000\027',
	ExcelMANTThinking = '\000|\000\030',
	ExcelMANTSendingMail = '\000|\000\031',
	ExcelMANTListensToComputer = '\000|\000\032',
	ExcelMANTDisappear = '\000|\000\037',
	ExcelMANTAppear = '\000|\000 ',
	ExcelMANTGetArtsy = '\000|\000d',
	ExcelMANTGetTechy = '\000|\000e',
	ExcelMANTGetWizardy = '\000|\000f',
	ExcelMANTCheckingSomething = '\000|\000g',
	ExcelMANTLookDown = '\000|\000h',
	ExcelMANTLookDownLeft = '\000|\000i',
	ExcelMANTLookDownRight = '\000|\000j',
	ExcelMANTLookLeft = '\000|\000k',
	ExcelMANTLookRight = '\000|\000l',
	ExcelMANTLookUp = '\000|\000m',
	ExcelMANTLookUpLeft = '\000|\000n',
	ExcelMANTLookUpRight = '\000|\000o',
	ExcelMANTSaving = '\000|\000p',
	ExcelMANTGestureDown = '\000|\000q',
	ExcelMANTGestureLeft = '\000|\000r',
	ExcelMANTGestureUp = '\000|\000s',
	ExcelMANTEmptyTrash = '\000|\000t'
};
typedef enum ExcelMANT ExcelMANT;

enum ExcelMBSt {
	ExcelMBStButtonNone = '\000}\000\000',
	ExcelMBStButtonOk = '\000}\000\001',
	ExcelMBStButtonCancel = '\000}\000\002',
	ExcelMBStButtonsOkCancel = '\000}\000\003',
	ExcelMBStButtonsYesNo = '\000}\000\004',
	ExcelMBStButtonsYesNoCancel = '\000}\000\005',
	ExcelMBStButtonsBackClose = '\000}\000\006',
	ExcelMBStButtonsNextClose = '\000}\000\007',
	ExcelMBStButtonsBackNextClose = '\000}\000\010',
	ExcelMBStButtonsRetryCancel = '\000}\000\011',
	ExcelMBStButtonsAbortRetryIgnore = '\000}\000\012',
	ExcelMBStButtonsSearchClose = '\000}\000\013',
	ExcelMBStButtonsBackNextSnooze = '\000}\000\014',
	ExcelMBStButtonsTipsOptionsClose = '\000}\000\015',
	ExcelMBStButtonsYesAllNoCancel = '\000}\000\016'
};
typedef enum ExcelMBSt ExcelMBSt;

enum ExcelMIct {
	ExcelMIctIconNone = '\000~\000\000',
	ExcelMIctIconApplication = '\000~\000\001',
	ExcelMIctIconAlert = '\000~\000\002',
	ExcelMIctIconTip = '\000~\000\003',
	ExcelMIctIconAlertCritical = '\000~\000e',
	ExcelMIctIconAlertWarning = '\000~\000g',
	ExcelMIctIconAlertInfo = '\000~\000h'
};
typedef enum ExcelMIct ExcelMIct;

enum ExcelMWAt {
	ExcelMWAtInactive = '\000\202\000\000',
	ExcelMWAtActive = '\000\202\000\001',
	ExcelMWAtSuspend = '\000\202\000\002',
	ExcelMWAtResume = '\000\202\000\003'
};
typedef enum ExcelMWAt ExcelMWAt;

enum ExcelMeDP {
	ExcelMeDPPropertyTypeNumber = '\000\242\000\001',
	ExcelMeDPPropertyTypeBoolean = '\000\242\000\002',
	ExcelMeDPPropertyTypeDate = '\000\242\000\003',
	ExcelMeDPPropertyTypeString = '\000\242\000\004',
	ExcelMeDPPropertyTypeFloat = '\000\242\000\005'
};
typedef enum ExcelMeDP ExcelMeDP;

enum ExcelMASc {
	ExcelMAScMsoAutomationSecurityLow = '\000\243\000\001',
	ExcelMAScMsoAutomationSecurityByUI = '\000\243\000\002',
	ExcelMAScMsoAutomationSecurityForceDisable = '\000\243\000\003'
};
typedef enum ExcelMASc ExcelMASc;

enum ExcelMSsz {
	ExcelMSszResolution544x376 = '\000\204\000\000',
	ExcelMSszResolution640x480 = '\000\204\000\001',
	ExcelMSszResolution720x512 = '\000\204\000\002',
	ExcelMSszResolution800x600 = '\000\204\000\003',
	ExcelMSszResolution1024x768 = '\000\204\000\004',
	ExcelMSszResolution1152x882 = '\000\204\000\005',
	ExcelMSszResolution1152x900 = '\000\204\000\006',
	ExcelMSszResolution1280x1024 = '\000\204\000\007',
	ExcelMSszResolution1600x1200 = '\000\204\000\010',
	ExcelMSszResolution1800x1440 = '\000\204\000\011',
	ExcelMSszResolution1920x1200 = '\000\204\000\012'
};
typedef enum ExcelMSsz ExcelMSsz;

enum ExcelMChS {
	ExcelMChSArabicCharacterSet = '\000\205\000\001',
	ExcelMChSCyrillicCharacterSet = '\000\205\000\002',
	ExcelMChSEnglishCharacterSet = '\000\205\000\003',
	ExcelMChSGreekCharacterSet = '\000\205\000\004',
	ExcelMChSHebrewCharacterSet = '\000\205\000\005',
	ExcelMChSJapaneseCharacterSet = '\000\205\000\006',
	ExcelMChSKoreanCharacterSet = '\000\205\000\007',
	ExcelMChSMultilingualUnicodeCharacterSet = '\000\205\000\010',
	ExcelMChSSimplifiedChineseCharacterSet = '\000\205\000\011',
	ExcelMChSThaiCharacterSet = '\000\205\000\012',
	ExcelMChSTraditionalChineseCharacterSet = '\000\205\000\013',
	ExcelMChSVietnameseCharacterSet = '\000\205\000\014'
};
typedef enum ExcelMChS ExcelMChS;

enum ExcelMtEn {
	ExcelMtEnEncodingThai = '\000\213\003j',
	ExcelMtEnEncodingJapaneseShiftJIS = '\000\213\003\244',
	ExcelMtEnEncodingSimplifiedChinese = '\000\213\003\250',
	ExcelMtEnEncodingKorean = '\000\213\003\265',
	ExcelMtEnEncodingBig5TraditionalChinese = '\000\213\003\266',
	ExcelMtEnEncodingLittleEndian = '\000\213\004\260',
	ExcelMtEnEncodingBigEndian = '\000\213\004\261',
	ExcelMtEnEncodingCentralEuropean = '\000\213\004\342',
	ExcelMtEnEncodingCyrillic = '\000\213\004\343',
	ExcelMtEnEncodingWestern = '\000\213\004\344',
	ExcelMtEnEncodingGreek = '\000\213\004\345',
	ExcelMtEnEncodingTurkish = '\000\213\004\346',
	ExcelMtEnEncodingHebrew = '\000\213\004\347',
	ExcelMtEnEncodingArabic = '\000\213\004\350',
	ExcelMtEnEncodingBaltic = '\000\213\004\351',
	ExcelMtEnEncodingVietnamese = '\000\213\004\352',
	ExcelMtEnEncodingISO88591Latin1 = '\000\213o\257',
	ExcelMtEnEncodingISO88592CentralEurope = '\000\213o\260',
	ExcelMtEnEncodingISO88593Latin3 = '\000\213o\261',
	ExcelMtEnEncodingISO88594Baltic = '\000\213o\262',
	ExcelMtEnEncodingISO88595Cyrillic = '\000\213o\263',
	ExcelMtEnEncodingISO88596Arabic = '\000\213o\264',
	ExcelMtEnEncodingISO88597Greek = '\000\213o\265',
	ExcelMtEnEncodingISO88598Hebrew = '\000\213o\266',
	ExcelMtEnEncodingISO88599Turkish = '\000\213o\267',
	ExcelMtEnEncodingISO885915Latin9 = '\000\213o\275',
	ExcelMtEnEncodingISO2022JapaneseNoHalfWidthKatakana = '\000\213\304,',
	ExcelMtEnEncodingISO2022JapaneseJISX02021984 = '\000\213\304-',
	ExcelMtEnEncodingISO2022JapaneseJISX02011989 = '\000\213\304.',
	ExcelMtEnEncodingISO2022KR = '\000\213\3041',
	ExcelMtEnEncodingISO2022CNTraditionalChinese = '\000\213\3043',
	ExcelMtEnEncodingISO2022CNSimplifiedChinese = '\000\213\3045',
	ExcelMtEnEncodingMacRoman = '\000\213\'\020',
	ExcelMtEnEncodingMacJapanese = '\000\213\'\021',
	ExcelMtEnEncodingMacTraditionalChinese = '\000\213\'\022',
	ExcelMtEnEncodingMacKorean = '\000\213\'\023',
	ExcelMtEnEncodingMacArabic = '\000\213\'\024',
	ExcelMtEnEncodingMacHebrew = '\000\213\'\025',
	ExcelMtEnEncodingMacGreek1 = '\000\213\'\026',
	ExcelMtEnEncodingMacCyrillic = '\000\213\'\027',
	ExcelMtEnEncodingMacSimplifiedChineseGB2312 = '\000\213\'\030',
	ExcelMtEnEncodingMacRomania = '\000\213\'\032',
	ExcelMtEnEncodingMacUkraine = '\000\213\'!',
	ExcelMtEnEncodingMacLatin2 = '\000\213\'-',
	ExcelMtEnEncodingMacIcelandic = '\000\213\'_',
	ExcelMtEnEncodingMacTurkish = '\000\213\'a',
	ExcelMtEnEncodingMacCroatia = '\000\213\'b',
	ExcelMtEnEncodingEBCDICUSCanada = '\000\213\000%',
	ExcelMtEnEncodingEBCDICInternational = '\000\213\001\364',
	ExcelMtEnEncodingEBCDICMultilingualROECELatin2 = '\000\213\003f',
	ExcelMtEnEncodingEBCDICGreekModern = '\000\213\003k',
	ExcelMtEnEncodingEBCDICTurkishLatin5 = '\000\213\004\002',
	ExcelMtEnEncodingEBCDICGermany = '\000\213O1',
	ExcelMtEnEncodingEBCDICDenmarkNorway = '\000\213O5',
	ExcelMtEnEncodingEBCDICFinlandSweden = '\000\213O6',
	ExcelMtEnEncodingEBCDICItaly = '\000\213O8',
	ExcelMtEnEncodingEBCDICLatinAmericaSpain = '\000\213O<',
	ExcelMtEnEncodingEBCDICUnitedKingdom = '\000\213O=',
	ExcelMtEnEncodingEBCDICJapaneseKatakanaExtended = '\000\213OB',
	ExcelMtEnEncodingEBCDICFrance = '\000\213OI',
	ExcelMtEnEncodingEBCDICArabic = '\000\213O\304',
	ExcelMtEnEncodingEBCDICGreek = '\000\213O\307',
	ExcelMtEnEncodingEBCDICHebrew = '\000\213O\310',
	ExcelMtEnEncodingEBCDICKoreanExtended = '\000\213Qa',
	ExcelMtEnEncodingEBCDICThai = '\000\213Qf',
	ExcelMtEnEncodingEBCDICIcelandic = '\000\213Q\207',
	ExcelMtEnEncodingEBCDICTurkish = '\000\213Q\251',
	ExcelMtEnEncodingEBCDICRussian = '\000\213Q\220',
	ExcelMtEnEncodingEBCDICSerbianBulgarian = '\000\213R!',
	ExcelMtEnEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = '\000\213\306\362',
	ExcelMtEnEncodingEBCDICUSCanadaAndJapanese = '\000\213\306\363',
	ExcelMtEnEncodingEBCDICExtendedAndKorean = '\000\213\306\365',
	ExcelMtEnEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = '\000\213\306\367',
	ExcelMtEnEncodingEBCDICUSCanadaAndTraditionalChinese = '\000\213\306\371',
	ExcelMtEnEncodingEBCDICJapaneseLatinExtendedAndJapanese = '\000\213\306\373',
	ExcelMtEnEncodingOEMUnitedStates = '\000\213\001\265',
	ExcelMtEnEncodingOEMGreek = '\000\213\002\341',
	ExcelMtEnEncodingOEMBaltic = '\000\213\003\007',
	ExcelMtEnEncodingOEMMultilingualLatinI = '\000\213\003R',
	ExcelMtEnEncodingOEMMultilingualLatinII = '\000\213\003T',
	ExcelMtEnEncodingOEMCyrillic = '\000\213\003W',
	ExcelMtEnEncodingOEMTurkish = '\000\213\003Y',
	ExcelMtEnEncodingOEMPortuguese = '\000\213\003\\',
	ExcelMtEnEncodingOEMIcelandic = '\000\213\003]',
	ExcelMtEnEncodingOEMHebrew = '\000\213\003^',
	ExcelMtEnEncodingOEMCanadianFrench = '\000\213\003_',
	ExcelMtEnEncodingOEMArabic = '\000\213\003`',
	ExcelMtEnEncodingOEMNordic = '\000\213\003a',
	ExcelMtEnEncodingOEMCyrillicII = '\000\213\003b',
	ExcelMtEnEncodingOEMModernGreek = '\000\213\003e',
	ExcelMtEnEncodingEUCJapanese = '\000\213\312\334',
	ExcelMtEnEncodingEUCChineseSimplifiedChinese = '\000\213\312\340',
	ExcelMtEnEncodingEUCKorean = '\000\213\312\355',
	ExcelMtEnEncodingEUCTaiwaneseTraditionalChinese = '\000\213\312\356',
	ExcelMtEnEncodingDevanagari = '\000\213\336\252',
	ExcelMtEnEncodingBengali = '\000\213\336\253',
	ExcelMtEnEncodingTamil = '\000\213\336\254',
	ExcelMtEnEncodingTelugu = '\000\213\336\255',
	ExcelMtEnEncodingAssamese = '\000\213\336\256',
	ExcelMtEnEncodingOriya = '\000\213\336\257',
	ExcelMtEnEncodingKannada = '\000\213\336\260',
	ExcelMtEnEncodingMalayalam = '\000\213\336\261',
	ExcelMtEnEncodingGujarati = '\000\213\336\262',
	ExcelMtEnEncodingPunjabi = '\000\213\336\263',
	ExcelMtEnEncodingArabicASMO = '\000\213\002\304',
	ExcelMtEnEncodingArabicTransparentASMO = '\000\213\002\320',
	ExcelMtEnEncodingKoreanJohab = '\000\213\005Q',
	ExcelMtEnEncodingTaiwanCNS = '\000\213N ',
	ExcelMtEnEncodingTaiwanTCA = '\000\213N!',
	ExcelMtEnEncodingTaiwanEten = '\000\213N\"',
	ExcelMtEnEncodingTaiwanIBM5550 = '\000\213N#',
	ExcelMtEnEncodingTaiwanTeletext = '\000\213N$',
	ExcelMtEnEncodingTaiwanWang = '\000\213N%',
	ExcelMtEnEncodingIA5IRV = '\000\213N\211',
	ExcelMtEnEncodingIA5German = '\000\213N\212',
	ExcelMtEnEncodingIA5Swedish = '\000\213N\213',
	ExcelMtEnEncodingIA5Norwegian = '\000\213N\214',
	ExcelMtEnEncodingUSASCII = '\000\213N\237',
	ExcelMtEnEncodingT61 = '\000\213O%',
	ExcelMtEnEncodingISO6937NonspacingAccent = '\000\213O-',
	ExcelMtEnEncodingKOI8R = '\000\213Q\202',
	ExcelMtEnEncodingExtAlphaLowercase = '\000\213R#',
	ExcelMtEnEncodingKOI8U = '\000\213Uj',
	ExcelMtEnEncodingEuropa3 = '\000\213qI',
	ExcelMtEnEncodingHZGBSimplifiedChinese = '\000\213\316\310',
	ExcelMtEnEncodingUTF7 = '\000\213\375\350',
	ExcelMtEnEncodingUTF8 = '\000\213\375\351'
};
typedef enum ExcelMtEn ExcelMtEn;

enum Excel4000 {
	Excel4000CommandBar = 'msCB',
	Excel4000CommandBarControl = 'mCBC'
};
typedef enum Excel4000 Excel4000;

enum ExcelE102 {
	ExcelE102BuiltInChart = '\001\366\000\025',
	ExcelE102UserDefined = '\001\366\000\026',
	ExcelE102AnyGallery = '\001\366\000\027'
};
typedef enum ExcelE102 ExcelE102;

enum ExcelE103 {
	ExcelE103ColorIndexAutomatic = '\001\366\357\367',
	ExcelE103ColorIndexNone = '\001\366\357\322',
	ExcelE103AColorIndexInteger = '\001\367\000\000'
};
typedef enum ExcelE103 ExcelE103;

enum ExcelE104 {
	ExcelE104Cap = '\001\370\000\001',
	ExcelE104NoCap = '\001\370\000\002'
};
typedef enum ExcelE104 ExcelE104;

enum ExcelE105 {
	ExcelE105ByColumns = '\001\371\000\002',
	ExcelE105ByRows = '\001\371\000\001'
};
typedef enum ExcelE105 ExcelE105;

enum ExcelE106 {
	ExcelE106ScaleLinear = '\001\371\357\334',
	ExcelE106ScaleLogarithmic = '\001\371\357\333'
};
typedef enum ExcelE106 ExcelE106;

enum ExcelE107 {
	ExcelE107AutofillSeries = '\001\373\000\004',
	ExcelE107ChronologicalSeries = '\001\373\000\003',
	ExcelE107GrowthSeries = '\001\373\000\002',
	ExcelE107DataSeriesLinear = '\001\372\357\334'
};
typedef enum ExcelE107 ExcelE107;

enum ExcelE108 {
	ExcelE108AxisCrossesAutomatic = '\001\373\357\367',
	ExcelE108AxisCrossesCustom = '\001\373\357\356',
	ExcelE108AxisCrossesMaximum = '\001\374\000\002',
	ExcelE108AxisCrossesMinimum = '\001\374\000\004'
};
typedef enum ExcelE108 ExcelE108;

enum ExcelE109 {
	ExcelE109PrimaryAxis = '\001\375\000\001',
	ExcelE109SecondaryAxis = '\001\375\000\002'
};
typedef enum ExcelE109 ExcelE109;

enum ExcelE110 {
	ExcelE110BackgroundAutomatic = '\001\375\357\367',
	ExcelE110BackgroundOpaque = '\001\376\000\003',
	ExcelE110BackgroundTransparent = '\001\376\000\002'
};
typedef enum ExcelE110 ExcelE110;

enum ExcelE111 {
	ExcelE111WindowStateMaximized = '\001\376\357\327',
	ExcelE111WindowStateMinimized = '\001\376\357\324',
	ExcelE111WindowStateNormal = '\001\376\357\321'
};
typedef enum ExcelE111 ExcelE111;

enum ExcelE112 {
	ExcelE112CategoryAxis = '\002\000\000\001',
	ExcelE112SeriesAxis = '\002\000\000\003',
	ExcelE112ValueAxis = '\002\000\000\002'
};
typedef enum ExcelE112 ExcelE112;

enum ExcelE113 {
	ExcelE113ArrowheadLengthLong = '\002\001\000\003',
	ExcelE113ArrowheadLengthMedium = '\002\000\357\326',
	ExcelE113ArrowheadLengthShort = '\002\001\000\001'
};
typedef enum ExcelE113 ExcelE113;

enum ExcelE114 {
	ExcelE114ValignBottom = '\002\001\357\365',
	ExcelE114ValignCenter = '\002\001\357\364',
	ExcelE114ValignDistributed = '\002\001\357\353',
	ExcelE114ValignJustify = '\002\001\357\336',
	ExcelE114ValignTop = '\002\001\357\300'
};
typedef enum ExcelE114 ExcelE114;

enum ExcelE115 {
	ExcelE115TickMarkCross = '\002\003\000\004',
	ExcelE115TickMarkInside = '\002\003\000\002',
	ExcelE115TickMarkNone = '\002\002\357\322',
	ExcelE115TickMarkOutside = '\002\003\000\003'
};
typedef enum ExcelE115 ExcelE115;

enum ExcelE116 {
	ExcelE116ErrorBarDirectionX = '\002\003\357\270',
	ExcelE116ErrorBarDirectionY = '\002\004\000\001'
};
typedef enum ExcelE116 ExcelE116;

enum ExcelE117 {
	ExcelE117ErrorBarIncludeBoth = '\002\005\000\001',
	ExcelE117ErrorBarIncludeMinusValues = '\002\005\000\003',
	ExcelE117ErrorBarIncludeNone = '\002\004\357\322',
	ExcelE117ErrorBarIncludePlusValues = '\002\005\000\002'
};
typedef enum ExcelE117 ExcelE117;

enum ExcelE118 {
	ExcelE118Interpolated = '\002\006\000\003',
	ExcelE118NotPlotted = '\002\006\000\001',
	ExcelE118Zero = '\002\006\000\002'
};
typedef enum ExcelE118 ExcelE118;

enum ExcelE119 {
	ExcelE119ArrowheadStyleClosed = '\002\007\000\003',
	ExcelE119ArrowheadStyleDoubleClosed = '\002\007\000\005',
	ExcelE119ArrowheadStyleDoubleOpen = '\002\007\000\004',
	ExcelE119ArrowheadStyleNone = '\002\006\357\322',
	ExcelE119ArrowheadStyleOpen = '\002\007\000\002'
};
typedef enum ExcelE119 ExcelE119;

enum ExcelE120 {
	ExcelE120ArrowheadWidthMedium = '\002\007\357\326',
	ExcelE120ArrowheadWidthNarrow = '\002\010\000\001',
	ExcelE120ArrowheadWidthWide = '\002\010\000\003'
};
typedef enum ExcelE120 ExcelE120;

enum ExcelE121 {
	ExcelE121HorizontalAlignCenter = '\002\010\357\364',
	ExcelE121HorizontalAlignCenterAcrossSelection = '\002\011\000\007',
	ExcelE121HorizontalAlignDistributed = '\002\010\357\353',
	ExcelE121HorizontalAlignFill = '\002\011\000\005',
	ExcelE121HorizontalAlignGeneral = '\002\011\000\001',
	ExcelE121HorizontalAlignJustify = '\002\010\357\336',
	ExcelE121HorizontalAlignLeft = '\002\010\357\335',
	ExcelE121HorizontalAlignRight = '\002\010\357\310'
};
typedef enum ExcelE121 ExcelE121;

enum ExcelE122 {
	ExcelE122TickLabelPositionHigh = '\002\011\357\341',
	ExcelE122TickLabelPositionLow = '\002\011\357\332',
	ExcelE122TickLabelPositionNextToAxis = '\002\012\000\004',
	ExcelE122TickLabelPositionNone = '\002\011\357\322'
};
typedef enum ExcelE122 ExcelE122;

enum ExcelE123 {
	ExcelE123LegendPositionBottom = '\002\012\357\365',
	ExcelE123LegendPositionCorner = '\002\013\000\002',
	ExcelE123LegendPositionLeft = '\002\012\357\335',
	ExcelE123LegendPositionRight = '\002\012\357\310',
	ExcelE123LegendPositionTop = '\002\012\357\300'
};
typedef enum ExcelE123 ExcelE123;

enum ExcelE124 {
	ExcelE124ChartPictureTypeStackScale = '\002\014\000\003',
	ExcelE124ChartPictureTypeStack = '\002\014\000\002',
	ExcelE124ChartPictureTypeStretch = '\002\014\000\001'
};
typedef enum ExcelE124 ExcelE124;

enum ExcelE125 {
	ExcelE125Sides = '\002\015\000\001',
	ExcelE125End = '\002\015\000\002',
	ExcelE125EndSides = '\002\015\000\003',
	ExcelE125Front = '\002\015\000\004',
	ExcelE125FrontSides = '\002\015\000\005',
	ExcelE125FrontEnd = '\002\015\000\006',
	ExcelE125AllFaces = '\002\015\000\007'
};
typedef enum ExcelE125 ExcelE125;

enum ExcelE126 {
	ExcelE126OrientationDownward = '\002\015\357\266',
	ExcelE126OrientationHorizontal = '\002\015\357\340',
	ExcelE126OrientationUpward = '\002\015\357\265',
	ExcelE126OrientationVertical = '\002\015\357\272'
};
typedef enum ExcelE126 ExcelE126;

enum ExcelE127 {
	ExcelE127TickLabelOrientationAutomatic = '\002\016\357\367',
	ExcelE127TickLabelOrientationDownward = '\002\016\357\266',
	ExcelE127TickLabelOrientationHorizontal = '\002\016\357\340',
	ExcelE127TickLabelOrientationUpward = '\002\016\357\265',
	ExcelE127TickLabelOrientationVertical = '\002\016\357\272'
};
typedef enum ExcelE127 ExcelE127;

enum ExcelE128 {
	ExcelE128BorderWeightHairline = '\002\020\000\001',
	ExcelE128BorderWeightMedium = '\002\017\357\326',
	ExcelE128BorderWeightThick = '\002\020\000\004',
	ExcelE128BorderWeightThin = '\002\020\000\002'
};
typedef enum ExcelE128 ExcelE128;

enum ExcelE129 {
	ExcelE129SeriesDateDay = '\002\021\000\001',
	ExcelE129SeriesDateMonth = '\002\021\000\003',
	ExcelE129SeriesDateWeekday = '\002\021\000\002',
	ExcelE129SeriesDateYear = '\002\021\000\004'
};
typedef enum ExcelE129 ExcelE129;

enum ExcelE130 {
	ExcelE130UnderlineStyleDouble = '\002\021\357\351',
	ExcelE130UnderlineStyleDoubleAccounting = '\002\022\000\005',
	ExcelE130UnderlineStyleNone = '\002\021\357\322',
	ExcelE130UnderlineStyleSingle = '\002\022\000\002',
	ExcelE130UnderlineStyleSingleAccounting = '\002\022\000\004'
};
typedef enum ExcelE130 ExcelE130;

enum ExcelE131 {
	ExcelE131ErrorBarTypeCustom = '\002\022\357\356',
	ExcelE131ErrorBarTypeFixedValue = '\002\023\000\001',
	ExcelE131ErrorBarTypePercent = '\002\023\000\002',
	ExcelE131ErrorBarTypeStandardDeviation = '\002\022\357\305',
	ExcelE131ErrorBarTypeStandardError = '\002\023\000\004'
};
typedef enum ExcelE131 ExcelE131;

enum ExcelE132 {
	ExcelE132Exponential = '\002\024\000\005',
	ExcelE132Linear = '\002\023\357\334',
	ExcelE132Logarithmic = '\002\023\357\333',
	ExcelE132MovingAverage = '\002\024\000\006',
	ExcelE132Polynomial = '\002\024\000\003',
	ExcelE132Power = '\002\024\000\004'
};
typedef enum ExcelE132 ExcelE132;

enum ExcelE133 {
	ExcelE133Continuous = '\002\025\000\001',
	ExcelE133Dash = '\002\024\357\355',
	ExcelE133DashDot = '\002\025\000\004',
	ExcelE133DashDotDot = '\002\025\000\005',
	ExcelE133Dot = '\002\024\357\352',
	ExcelE133Double = '\002\024\357\351',
	ExcelE133SlantDashDot = '\002\025\000\015',
	ExcelE133LineStyleNone = '\002\024\357\322'
};
typedef enum ExcelE133 ExcelE133;

enum ExcelE134 {
	ExcelE134DataLabelsShowNone = '\002\025\357\322',
	ExcelE134DataLabelsShowValue = '\002\026\000\002',
	ExcelE134DataLabelsShowPercent = '\002\026\000\003',
	ExcelE134DataLabelsShowLabel = '\002\026\000\004',
	ExcelE134DataLabelsShowLabelAndPercent = '\002\026\000\005',
	ExcelE134DataLabelsShowBubbleSizes = '\002\026\000\006'
};
typedef enum ExcelE134 ExcelE134;

enum ExcelE135 {
	ExcelE135MarkerStyleAutomatic = '\002\026\357\367',
	ExcelE135MarkerStyleCircle = '\002\027\000\010',
	ExcelE135MarkerStyleDash = '\002\026\357\355',
	ExcelE135MarkerStyleDiamond = '\002\027\000\002',
	ExcelE135MarkerStyleDot = '\002\026\357\352',
	ExcelE135MarkerStyleNone = '\002\026\357\322',
	ExcelE135MarkerStylePicture = '\002\026\357\315',
	ExcelE135MarkerStylePlus = '\002\027\000\011',
	ExcelE135MarkerStyleSquare = '\002\027\000\001',
	ExcelE135MarkerStyleStar = '\002\027\000\005',
	ExcelE135MarkerStyleTriangle = '\002\027\000\003',
	ExcelE135MarkerStyleX = '\002\026\357\270'
};
typedef enum ExcelE135 ExcelE135;

enum ExcelE137 {
	ExcelE137PatternAutomatic = '\002\030\357\367',
	ExcelE137PatternChecker = '\002\031\000\011',
	ExcelE137PatternCrissCross = '\002\031\000\020',
	ExcelE137PatternDown = '\002\030\357\347',
	ExcelE137PatternGray16 = '\002\031\000\021',
	ExcelE137PatternGray25 = '\002\030\357\344',
	ExcelE137PatternGray50 = '\002\030\357\343',
	ExcelE137PatternGray75 = '\002\030\357\342',
	ExcelE137PatternGray8 = '\002\031\000\022',
	ExcelE137PatternGrid = '\002\031\000\017',
	ExcelE137PatternHorizontal = '\002\030\357\340',
	ExcelE137PatternLightDown = '\002\031\000\015',
	ExcelE137PatternLightHorizontal = '\002\031\000\013',
	ExcelE137PatternLightUp = '\002\031\000\016',
	ExcelE137PatternLightVertical = '\002\031\000\014',
	ExcelE137PatternNone = '\002\030\357\322',
	ExcelE137PatternSemiGray75 = '\002\031\000\012',
	ExcelE137PatternSolid = '\002\031\000\001',
	ExcelE137PatternUp = '\002\030\357\276',
	ExcelE137PatternVertical = '\002\030\357\272',
	ExcelE137PatternLinearGradient = '\002\031\017\240',
	ExcelE137PatternRectangularGradient = '\002\031\017\241'
};
typedef enum ExcelE137 ExcelE137;

enum ExcelE138 {
	ExcelE138SplitByPosition = '\002\032\000\001',
	ExcelE138SplitByPercentValue = '\002\032\000\003',
	ExcelE138SplitByCustomSplit = '\002\032\000\004',
	ExcelE138SplitByValue = '\002\032\000\002'
};
typedef enum ExcelE138 ExcelE138;

enum ExcelE139 {
	ExcelE139Hundreds = '\002\032\377\376',
	ExcelE139Thousands = '\002\032\377\375',
	ExcelE139TenThousands = '\002\032\377\374',
	ExcelE139HundredThousands = '\002\032\377\373',
	ExcelE139Millions = '\002\032\377\372',
	ExcelE139TenMillions = '\002\032\377\371',
	ExcelE139HundredMillions = '\002\032\377\370',
	ExcelE139ThousandMillions = '\002\032\377\367',
	ExcelE139MillionMillions = '\002\032\377\366',
	ExcelE139CustomDisplayUnit = '\002\032\357\356'
};
typedef enum ExcelE139 ExcelE139;

enum ExcelE140 {
	ExcelE140LabelPositionCenter = '\002\033\357\364',
	ExcelE140LabelPositionAbove = '\002\034\000\000',
	ExcelE140LabelPositionBelow = '\002\034\000\001',
	ExcelE140LabelPositionLeft = '\002\033\357\335',
	ExcelE140LabelPositionRight = '\002\033\357\310',
	ExcelE140LabelPositionOutsideEnd = '\002\034\000\002',
	ExcelE140LabelPositionInsideEnd = '\002\034\000\003',
	ExcelE140LabelPositionInsideBase = '\002\034\000\004',
	ExcelE140LabelPositionBestFit = '\002\034\000\005',
	ExcelE140LabelPositionMixed = '\002\034\000\006',
	ExcelE140LabelPositionCustom = '\002\034\000\007'
};
typedef enum ExcelE140 ExcelE140;

enum ExcelE141 {
	ExcelE141Days = '\002\035\000\000',
	ExcelE141Months = '\002\035\000\001',
	ExcelE141Years = '\002\035\000\002'
};
typedef enum ExcelE141 ExcelE141;

enum ExcelE142 {
	ExcelE142CategoryScale = '\002\036\000\002',
	ExcelE142TimeScale = '\002\036\000\003',
	ExcelE142AutomaticScale = '\002\035\357\367'
};
typedef enum ExcelE142 ExcelE142;

enum ExcelE143 {
	ExcelE143Box = '\002\037\000\000',
	ExcelE143PyramidToPoint = '\002\037\000\001',
	ExcelE143PyramidToMax = '\002\037\000\002',
	ExcelE143Cylinder = '\002\037\000\003',
	ExcelE143ConeToPoint = '\002\037\000\004',
	ExcelE143ConeToMax = '\002\037\000\005'
};
typedef enum ExcelE143 ExcelE143;

enum ExcelE144 {
	ExcelE144ColumnClustered = '\002 \0003',
	ExcelE144ColumnStacked = '\002 \0004',
	ExcelE144ColumnStacked100 = '\002 \0005',
	ExcelE144ThreeDColumnClustered = '\002 \0006',
	ExcelE144ThreeDColumnStacked = '\002 \0007',
	ExcelE144ThreeDColumnStacked100 = '\002 \0008',
	ExcelE144BarClustered = '\002 \0009',
	ExcelE144BarStacked = '\002 \000:',
	ExcelE144BarStacked100 = '\002 \000;',
	ExcelE144ThreeDBarClustered = '\002 \000<',
	ExcelE144ThreeDBarStacked = '\002 \000=',
	ExcelE144ThreeDBarStacked100 = '\002 \000>',
	ExcelE144LineStacked = '\002 \000\?',
	ExcelE144LineStacked100 = '\002 \000@',
	ExcelE144LineMarkers = '\002 \000A',
	ExcelE144LineMarkersStacked = '\002 \000B',
	ExcelE144LineMarkersStacked100 = '\002 \000C',
	ExcelE144PieOfPie = '\002 \000D',
	ExcelE144PieExploded = '\002 \000E',
	ExcelE144ThreeDPieExploded = '\002 \000F',
	ExcelE144BarOfPie = '\002 \000G',
	ExcelE144XyScatterSmooth = '\002 \000H',
	ExcelE144XyScatterSmoothNoMarkers = '\002 \000I',
	ExcelE144XyScatterLines = '\002 \000J',
	ExcelE144XyScatterLinesNoMarkers = '\002 \000K',
	ExcelE144AreaStacked = '\002 \000L',
	ExcelE144AreaStacked100 = '\002 \000M',
	ExcelE144ThreeDAreaStacked = '\002 \000N',
	ExcelE144ThreeDAreaStacked100 = '\002 \000O',
	ExcelE144DoughnutExploded = '\002 \000P',
	ExcelE144RadarMarkers = '\002 \000Q',
	ExcelE144RadarFilled = '\002 \000R',
	ExcelE144Surface = '\002 \000S',
	ExcelE144SurfaceWireframe = '\002 \000T',
	ExcelE144SurfaceTopView = '\002 \000U',
	ExcelE144SurfaceTopViewWireframe = '\002 \000V',
	ExcelE144Bubble = '\002 \000\017',
	ExcelE144BubbleThreeDEffect = '\002 \000W',
	ExcelE144StockHLC = '\002 \000X',
	ExcelE144StockOHLC = '\002 \000Y',
	ExcelE144StockVHLC = '\002 \000Z',
	ExcelE144StockVOHLC = '\002 \000[',
	ExcelE144CylinderColumnClustered = '\002 \000\\',
	ExcelE144CylinderColumnStacked = '\002 \000]',
	ExcelE144CylinderColumnStacked100 = '\002 \000^',
	ExcelE144CylinderBarClustered = '\002 \000_',
	ExcelE144CylinderBarStacked = '\002 \000`',
	ExcelE144CylinderBarStacked100 = '\002 \000a',
	ExcelE144CylinderColumn = '\002 \000b',
	ExcelE144ConeColumnClustered = '\002 \000c',
	ExcelE144ConeColumnStacked = '\002 \000d',
	ExcelE144ConeColumnStacked100 = '\002 \000e',
	ExcelE144ConeBarClustered = '\002 \000f',
	ExcelE144ConeBarStacked = '\002 \000g',
	ExcelE144ConeBarStacked100 = '\002 \000h',
	ExcelE144ConeCol = '\002 \000i',
	ExcelE144PyramidColumnClustered = '\002 \000j',
	ExcelE144PyramidColumnStacked = '\002 \000k',
	ExcelE144PyramidColumnStacked100 = '\002 \000l',
	ExcelE144PyramidBarClustered = '\002 \000m',
	ExcelE144PyramidBarStacked = '\002 \000n',
	ExcelE144PyramidBarStacked100 = '\002 \000o',
	ExcelE144PyramidColumn = '\002 \000p',
	ExcelE144ThreeDColumn = '\002\037\357\374',
	ExcelE144LineChart = '\002 \000\004',
	ExcelE144ThreeDLine = '\002\037\357\373',
	ExcelE144ThreeDPie = '\002\037\357\372',
	ExcelE144PieChart = '\002 \000\005',
	ExcelE144Xyscatter = '\002\037\357\267',
	ExcelE144ThreeDArea = '\002\037\357\376',
	ExcelE144AreaChart = '\002 \000\001',
	ExcelE144Doughnut = '\002\037\357\350',
	ExcelE144Radar = '\002\037\357\311',
	ExcelE144CombinationChart = '\002\037\357\361'
};
typedef enum ExcelE144 ExcelE144;

enum ExcelE145 {
	ExcelE145DataLabel = '\002!\000\000',
	ExcelE145AChartArea = '\002!\000\002',
	ExcelE145ASeries = '\002!\000\003',
	ExcelE145AChartTitle = '\002!\000\004',
	ExcelE145Walls = '\002!\000\005',
	ExcelE145ACornersObject = '\002!\000\006',
	ExcelE145DataTable = '\002!\000\007',
	ExcelE145Trendline = '\002!\000\010',
	ExcelE145ErrorBarsObject = '\002!\000\011',
	ExcelE145XerrorBars = '\002!\000\012',
	ExcelE145YerrorBars = '\002!\000\013',
	ExcelE145LegendEntry = '\002!\000\014',
	ExcelE145LegendKey = '\002!\000\015',
	ExcelE145Shape = '\002!\000\016',
	ExcelE145MajorGridlines = '\002!\000\017',
	ExcelE145MinorGridlines = '\002!\000\020',
	ExcelE145AxisTitle = '\002!\000\021',
	ExcelE145UpBars = '\002!\000\022',
	ExcelE145PlotArea = '\002!\000\023',
	ExcelE145DownBars = '\002!\000\024',
	ExcelE145Axis = '\002!\000\025',
	ExcelE145SeriesLines = '\002!\000\026',
	ExcelE145Floor = '\002!\000\027',
	ExcelE145Legend = '\002!\000\030',
	ExcelE145HiLoLines = '\002!\000\031',
	ExcelE145DropLines = '\002!\000\032',
	ExcelE145RadarAxisLabels = '\002!\000\033',
	ExcelE145Nothing = '\002!\000\034',
	ExcelE145LeaderLines = '\002!\000\035',
	ExcelE145DisplayUnitLabel = '\002!\000\036'
};
typedef enum ExcelE145 ExcelE145;

enum ExcelE146 {
	ExcelE146SizeIsWidth = '\002\"\000\002',
	ExcelE146SizeIsArea = '\002\"\000\001'
};
typedef enum ExcelE146 ExcelE146;

enum ExcelE147 {
	ExcelE147ShiftDown = '\002\"\357\347',
	ExcelE147ShiftToRight = '\002\"\357\277'
};
typedef enum ExcelE147 ExcelE147;

enum ExcelE148 {
	ExcelE148ShiftToLeft = '\002#\357\301',
	ExcelE148ShiftUp = '\002#\357\276'
};
typedef enum ExcelE148 ExcelE148;

enum ExcelE149 {
	ExcelE149TowardTheBottom = '\002$\357\347',
	ExcelE149TowardTheLeft = '\002$\357\301',
	ExcelE149TowardTheRight = '\002$\357\277',
	ExcelE149TowardTheTop = '\002$\357\276'
};
typedef enum ExcelE149 ExcelE149;

enum ExcelE150 {
	ExcelE150DoAverage = '\002%\357\366',
	ExcelE150DoCount = '\002%\357\360',
	ExcelE150DoCountNumbers = '\002%\357\357',
	ExcelE150DoMaximum = '\002%\357\330',
	ExcelE150DoMinimum = '\002%\357\325',
	ExcelE150DoProduct = '\002%\357\313',
	ExcelE150DoStandardDeviation = '\002%\357\305',
	ExcelE150DoStandardDeviationP = '\002%\357\304',
	ExcelE150DoSum = '\002%\357\303',
	ExcelE150DoVar = '\002%\357\274',
	ExcelE150DoVarP = '\002%\357\273'
};
typedef enum ExcelE150 ExcelE150;

enum ExcelE151 {
	ExcelE151SheetTypeChart = '\002&\357\363',
	ExcelE151SheetTypeDialogSheet = '\002&\357\354',
	ExcelE151SheetTypeExcel4IntlMacroSheet = '\002\'\000\004',
	ExcelE151SheetTypeExcel4MacroSheet = '\002\'\000\003',
	ExcelE151SheetTypeWorksheet = '\002&\357\271'
};
typedef enum ExcelE151 ExcelE151;

enum ExcelE152 {
	ExcelE152ColumnHeader = '\002\'\357\362',
	ExcelE152ColumnItem = '\002(\000\005',
	ExcelE152DataHeader = '\002(\000\003',
	ExcelE152DataItem = '\002(\000\007',
	ExcelE152PageHeader = '\002(\000\002',
	ExcelE152PageItem = '\002(\000\006',
	ExcelE152RowHeader = '\002\'\357\307',
	ExcelE152RowItem = '\002(\000\004',
	ExcelE152TableBody = '\002(\000\010'
};
typedef enum ExcelE152 ExcelE152;

enum ExcelE153 {
	ExcelE153Formulas = '\002(\357\345',
	ExcelE153Comments = '\002(\357\320',
	ExcelE153Values = '\002(\357\275'
};
typedef enum ExcelE153 ExcelE153;

enum ExcelE154 {
	ExcelE154WindowTypeChartAsWindow = '\002*\000\005',
	ExcelE154WindowTypeChartInPlace = '\002*\000\004',
	ExcelE154WindowTypeClipboard = '\002*\000\003',
	ExcelE154WindowTypeInfo = '\002)\357\337',
	ExcelE154WindowTypeWorkbook = '\002*\000\001'
};
typedef enum ExcelE154 ExcelE154;

enum ExcelE155 {
	ExcelE155PivotFieldTypeDate = '\002+\000\002',
	ExcelE155PivotFieldTypeNumber = '\002*\357\317',
	ExcelE155PivotFieldTypeText = '\002*\357\302'
};
typedef enum ExcelE155 ExcelE155;

enum ExcelE156 {
	ExcelE156Bitmap = '\002,\000\002',
	ExcelE156Picture = '\002+\357\315'
};
typedef enum ExcelE156 ExcelE156;

enum ExcelE157 {
	ExcelE157Consolidation = '\002-\000\003',
	ExcelE157Database = '\002-\000\001',
	ExcelE157External = '\002-\000\002',
	ExcelE157PivotTable = '\002,\357\314'
};
typedef enum ExcelE157 ExcelE157;

enum ExcelE158 {
	ExcelE158A1 = '\002.\000\001',
	ExcelE158R1C1 = '\002-\357\312'
};
typedef enum ExcelE158 ExcelE158;

enum ExcelE159 {
	ExcelE159MicrosoftAccess = '\002/\000\004',
	ExcelE159MicrosoftFoxPro = '\002/\000\005',
	ExcelE159MicrosoftMail = '\002/\000\003',
	ExcelE159MicrosoftPowerPoint = '\002/\000\002',
	ExcelE159MicrosoftProject = '\002/\000\006',
	ExcelE159MicrosoftSchedulePlus = '\002/\000\007',
	ExcelE159MicrosoftWord = '\002/\000\001'
};
typedef enum ExcelE159 ExcelE159;

enum ExcelE160 {
	ExcelE160NoButton = '\0020\000\000',
	ExcelE160PrimaryButton = '\0020\000\001',
	ExcelE160SecondaryButton = '\0020\000\002'
};
typedef enum ExcelE160 ExcelE160;

enum ExcelE161 {
	ExcelE161CopyMode = '\0021\000\001',
	ExcelE161CutMode = '\0021\000\002'
};
typedef enum ExcelE161 ExcelE161;

enum ExcelE163 {
	ExcelE163FilterCopy = '\0023\000\002',
	ExcelE163FilterInPlace = '\0023\000\001'
};
typedef enum ExcelE163 ExcelE163;

enum ExcelE164 {
	ExcelE164DownThenOver = '\0024\000\001',
	ExcelE164OverThenDown = '\0024\000\002'
};
typedef enum ExcelE164 ExcelE164;

enum ExcelE165 {
	ExcelE165LinkTypeExcelLinks = '\0025\000\001',
	ExcelE165LinkTypeOLELinks = '\0025\000\002'
};
typedef enum ExcelE165 ExcelE165;

enum ExcelE166 {
	ExcelE166ColumnThenRow = '\0026\000\002',
	ExcelE166RowThenColumn = '\0026\000\001'
};
typedef enum ExcelE166 ExcelE166;

enum ExcelE167 {
	ExcelE167CancelKeyDisabled = '\0027\000\000',
	ExcelE167ErrorHandler = '\0027\000\002',
	ExcelE167Interrupt = '\0027\000\001'
};
typedef enum ExcelE167 ExcelE167;

enum ExcelE168 {
	ExcelE168PageBreakAutomatic = '\0027\357\367',
	ExcelE168PageBreakManual = '\0027\357\331',
	ExcelE168PageBreakNone = '\0028\000\000'
};
typedef enum ExcelE168 ExcelE168;

enum ExcelE170 {
	ExcelE170Landscape = '\002:\000\002',
	ExcelE170Portrait = '\002:\000\001'
};
typedef enum ExcelE170 ExcelE170;

enum ExcelE171 {
	ExcelE171EditionDate = '\002;\000\002',
	ExcelE171UpdateState = '\002;\000\001'
};
typedef enum ExcelE171 ExcelE171;

enum ExcelE172 {
	ExcelE172CommandUnderlinesAutomatic = '\002;\357\367',
	ExcelE172CommandUnderlinesOff = '\002;\357\316',
	ExcelE172CommandUnderlinesOn = '\002<\000\001'
};
typedef enum ExcelE172 ExcelE172;

enum ExcelE173 {
	ExcelE173VerbOpen = '\002=\000\002',
	ExcelE173VerbPrimary = '\002=\000\001'
};
typedef enum ExcelE173 ExcelE173;

enum ExcelE174 {
	ExcelE174CalculationAutomatic = '\002=\357\367',
	ExcelE174CalculationManual = '\002=\357\331',
	ExcelE174CalculationSemiautomatic = '\002>\000\002'
};
typedef enum ExcelE174 ExcelE174;

enum ExcelE175 {
	ExcelE175WorkbookReadOnly = '\002\?\000\003',
	ExcelE175WorkbookReadWrite = '\002\?\000\002'
};
typedef enum ExcelE175 ExcelE175;

enum ExcelE176 {
	ExcelE176FitToPage = '\002@\000\002',
	ExcelE176FullPage = '\002@\000\003',
	ExcelE176FullScreen = '\002@\000\001'
};
typedef enum ExcelE176 ExcelE176;

enum ExcelE177 {
	ExcelE177Part = '\002A\000\002',
	ExcelE177Whole = '\002A\000\001'
};
typedef enum ExcelE177 ExcelE177;

enum ExcelE178 {
	ExcelE178MAPI = '\002B\000\001',
	ExcelE178NoMailSystem = '\002B\000\000',
	ExcelE178PowerTalk = '\002B\000\002'
};
typedef enum ExcelE178 ExcelE178;

enum ExcelE179 {
	ExcelE179LinkInfoOlelinks = '\002C\000\002',
	ExcelE179LinkInfoPublishers = '\002C\000\005',
	ExcelE179LinkInfoSubscribers = '\002C\000\006'
};
typedef enum ExcelE179 ExcelE179;

enum ExcelE182 {
	ExcelE182CellTypeBlanks = '\002F\000\004',
	ExcelE182CellTypeConstants = '\002F\000\002',
	ExcelE182CellTypeFormulas = '\002E\357\345',
	ExcelE182CellTypeLastCell = '\002F\000\013',
	ExcelE182CellTypeComments = '\002E\357\320',
	ExcelE182CellTypeVisible = '\002F\000\014',
	ExcelE182CellTypeAllFormatConditions = '\002E\357\264',
	ExcelE182CellTypeSameFormatConditions = '\002E\357\263',
	ExcelE182CellTypeAllValidation = '\002E\357\262',
	ExcelE182CellTypeSameValidation = '\002E\357\261'
};
typedef enum ExcelE182 ExcelE182;

enum ExcelE183 {
	ExcelE183ArrangeStyleCascade = '\002G\000\007',
	ExcelE183ArrangeStyleHorizontal = '\002F\357\340',
	ExcelE183ArrangeStyleTiled = '\002G\000\001',
	ExcelE183ArrangeStyleVertical = '\002F\357\272'
};
typedef enum ExcelE183 ExcelE183;

enum ExcelE184 {
	ExcelE184IBeamCursor = '\002H\000\003',
	ExcelE184DefaultCursor = '\002G\357\321',
	ExcelE184NorthwestArrowCursor = '\002H\000\001',
	ExcelE184WaitCursor = '\002H\000\002'
};
typedef enum ExcelE184 ExcelE184;

enum ExcelE185 {
	ExcelE185FillCopy = '\002I\000\001',
	ExcelE185FillDays = '\002I\000\005',
	ExcelE185FillDefault = '\002I\000\000',
	ExcelE185FillFormats = '\002I\000\003',
	ExcelE185FillMonths = '\002I\000\007',
	ExcelE185FillSeries = '\002I\000\002',
	ExcelE185FillValues = '\002I\000\004',
	ExcelE185FillWeekdays = '\002I\000\006',
	ExcelE185FillYears = '\002I\000\010',
	ExcelE185GrowthTrend = '\002I\000\012',
	ExcelE185LinearTrend = '\002I\000\011'
};
typedef enum ExcelE185 ExcelE185;

enum ExcelE186 {
	ExcelE186AutofilterAnd = '\002J\000\001',
	ExcelE186Bottom10Items = '\002J\000\004',
	ExcelE186Bottom10Percent = '\002J\000\006',
	ExcelE186AutofilterOr = '\002J\000\002',
	ExcelE186Top10Items = '\002J\000\003',
	ExcelE186Top10Percent = '\002J\000\005',
	ExcelE186FilterByValue = '\002J\000\007',
	ExcelE186FilterByCellColor = '\002J\000\010',
	ExcelE186FilterByFontColor = '\002J\000\011',
	ExcelE186FilterByIcon = '\002J\000\012',
	ExcelE186FilterDynamic = '\002J\000\013',
	ExcelE186FilterNoFill = '\002J\000\014',
	ExcelE186FilterByAutomaticFontColor = '\002J\000\015',
	ExcelE186FilterByNoIcon = '\002J\000\016'
};
typedef enum ExcelE186 ExcelE186;

enum ExcelE187 {
	ExcelE187ClipboardFormatBiff = '\002K\000\010',
	ExcelE187ClipboardFormatBiff2 = '\002K\000\022',
	ExcelE187ClipboardFormatBiff3 = '\002K\000\024',
	ExcelE187ClipboardFormatBiff4 = '\002K\000\036',
	ExcelE187ClipboardFormatBinary = '\002K\000\017',
	ExcelE187ClipboardFormatBitmap = '\002K\000\011',
	ExcelE187ClipboardFormatCgm = '\002K\000\015',
	ExcelE187ClipboardFormatCsv = '\002K\000\005',
	ExcelE187ClipboardFormatDif = '\002K\000\004',
	ExcelE187ClipboardFormatDspText = '\002K\000\014',
	ExcelE187ClipboardFormatEmbeddedObject = '\002K\000\025',
	ExcelE187ClipboardFormatEmbedSource = '\002K\000\026',
	ExcelE187ClipboardFormatLink = '\002K\000\013',
	ExcelE187ClipboardFormatLinkSource = '\002K\000\027',
	ExcelE187ClipboardFormatLinkSourceDesc = '\002K\000 ',
	ExcelE187ClipboardFormatMovie = '\002K\000\030',
	ExcelE187ClipboardFormatNative = '\002K\000\016',
	ExcelE187ClipboardFormatObjectDesc = '\002K\000\037',
	ExcelE187ClipboardFormatObjectLink = '\002K\000\023',
	ExcelE187ClipboardFormatOwnerLink = '\002K\000\021',
	ExcelE187ClipboardFormatPict = '\002K\000\002',
	ExcelE187ClipboardFormatPrintPict = '\002K\000\003',
	ExcelE187ClipboardFormatRtf = '\002K\000\007',
	ExcelE187ClipboardFormatScreenPict = '\002K\000\035',
	ExcelE187ClipboardFormatStandardFont = '\002K\000\034',
	ExcelE187ClipboardFormatStandardScale = '\002K\000\033',
	ExcelE187ClipboardFormatSylk = '\002K\000\006',
	ExcelE187ClipboardFormatTable = '\002K\000\020',
	ExcelE187ClipboardFormatText = '\002K\000\000',
	ExcelE187ClipboardFormatToolFace = '\002K\000\031',
	ExcelE187ClipboardFormatToolFacePict = '\002K\000\032',
	ExcelE187ClipboardFormatValu = '\002K\000\001',
	ExcelE187ClipboardFormatWk1 = '\002K\000\012',
	ExcelE187ClipboardFormatUnicodeText = '\002K\000.',
	ExcelE187ClipboardFormatStyleText = '\002K\0005',
	ExcelE187ClipboardFormatUnicodeStyleText = '\002K\0007',
	ExcelE187ClipboardFormatBiff5 = '\002K\000!',
	ExcelE187ClipboardFormatPictureBuild = '\002K\000\"',
	ExcelE187ClipboardFormatOdbcConn = '\002K\000#',
	ExcelE187ClipboardFormatOdbcSql = '\002K\000$',
	ExcelE187ClipboardFormat3dPicture = '\002K\000%',
	ExcelE187ClipboardFormatUnexpected38 = '\002K\000&',
	ExcelE187ClipboardFormatDrawingDragDrop = '\002K\000\'',
	ExcelE187ClipboardFormatDrawing = '\002K\000(',
	ExcelE187ClipboardFormatUnexpected41 = '\002K\000)',
	ExcelE187ClipboardFormatUnexpected42 = '\002K\000*',
	ExcelE187ClipboardFormatUnexpected43 = '\002K\000+',
	ExcelE187ClipboardFormatHyperlink = '\002K\000,',
	ExcelE187ClipboardFormatUnexpected45 = '\002K\000-',
	ExcelE187ClipboardFormatWindowsBitmap = '\002K\000/',
	ExcelE187ClipboardFormatUniformResourceLocator = '\002K\0000',
	ExcelE187ClipboardFormatFileName = '\002K\0001',
	ExcelE187ClipboardFormatUnexpected50 = '\002K\0002',
	ExcelE187ClipboardFormatUnexpected51 = '\002K\0003',
	ExcelE187ClipboardFormatHypertextMarkupLanguage = '\002K\0004',
	ExcelE187ClipboardFormatOfficeScrapbookInfo = '\002K\0006',
	ExcelE187ClipboardFormatPortableDocumentFormat = '\002K\0008',
	ExcelE187ClipboardFormatExcelInternalShape = '\002K\0009',
	ExcelE187ClipboardFormatOfficeArtShape = '\002K\000:'
};
typedef enum ExcelE187 ExcelE187;

enum ExcelE188 {
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
};
typedef enum ExcelE188 ExcelE188;

enum ExcelE189 {
	ExcelE189Twenty_four_hourClock = '\002M\000!',
	ExcelE189FourDigitYears = '\002M\000+',
	ExcelE189AlternateArraySeparator = '\002M\000\020',
	ExcelE189ColumnSeparator = '\002M\000\016',
	ExcelE189Country_code = '\002M\000\001',
	ExcelE189Country_setting = '\002M\000\002',
	ExcelE189Currency_before = '\002M\000%',
	ExcelE189Currency_code = '\002M\000\031',
	ExcelE189Currency_digits = '\002M\000\033',
	ExcelE189Currency_leading_zeros = '\002M\000(',
	ExcelE189Currency_minus_sign = '\002M\000&',
	ExcelE189Currency_negative = '\002M\000\034',
	ExcelE189Currency_space_before = '\002M\000$',
	ExcelE189Currency_trailing_zeros = '\002M\000\'',
	ExcelE189Date_order = '\002M\000 ',
	ExcelE189Date_separator = '\002M\000\021',
	ExcelE189DayCode = '\002M\000\025',
	ExcelE189DayLeadingZero = '\002M\000*',
	ExcelE189DecimalSeparator = '\002M\000\003',
	ExcelE189GeneralFormatName = '\002M\000\032',
	ExcelE189HourCode = '\002M\000\026',
	ExcelE189LeftBrace = '\002M\000\014',
	ExcelE189LeftBracket = '\002M\000\012',
	ExcelE189ListSeparator = '\002M\000\005',
	ExcelE189LowerCaseColumnLetter = '\002M\000\011',
	ExcelE189LowerCaseRowLetter = '\002M\000\010',
	ExcelE189Mdy = '\002M\000,',
	ExcelE189Metric = '\002M\000#',
	ExcelE189Minute_code = '\002M\000\027',
	ExcelE189Month_code = '\002M\000\024',
	ExcelE189Month_leading_zero = '\002M\000)',
	ExcelE189Month_name_chars = '\002M\000\036',
	ExcelE189Noncurrency_digits = '\002M\000\035',
	ExcelE189NonEnglishFunctions = '\002M\000\"',
	ExcelE189RightBrace = '\002M\000\015',
	ExcelE189RightBracket = '\002M\000\013',
	ExcelE189RowSeparator = '\002M\000\017',
	ExcelE189SecondCode = '\002M\000\030',
	ExcelE189ThousandsSeparator = '\002M\000\004',
	ExcelE189TimeLeadingZero = '\002M\000-',
	ExcelE189TimeSeparator = '\002M\000\022',
	ExcelE189UpperCaseColumnLetter = '\002M\000\007',
	ExcelE189UpperCaseRowLetter = '\002M\000\006',
	ExcelE189Weekday_name_chars = '\002M\000\037',
	ExcelE189YearCode = '\002M\000\023'
};
typedef enum ExcelE189 ExcelE189;

enum ExcelE190 {
	ExcelE190PageBreakFull = '\002N\000\001',
	ExcelE190PageBreakPartial = '\002N\000\002'
};
typedef enum ExcelE190 ExcelE190;

enum ExcelE191 {
	ExcelE191OverwriteCells = '\002O\000\000',
	ExcelE191InsertDeleteCells = '\002O\000\001',
	ExcelE191InsertEntireRows = '\002O\000\002'
};
typedef enum ExcelE191 ExcelE191;

enum ExcelE192 {
	ExcelE192NoLabels = '\002O\357\322',
	ExcelE192RowLabels = '\002P\000\001',
	ExcelE192ColumnLabels = '\002P\000\002',
	ExcelE192MixedLabels = '\002P\000\003'
};
typedef enum ExcelE192 ExcelE192;

enum ExcelE193 {
	ExcelE193SinceMyLastSave = '\002Q\000\001',
	ExcelE193AllChanges = '\002Q\000\002',
	ExcelE193NotYetReviewed = '\002Q\000\003'
};
typedef enum ExcelE193 ExcelE193;

enum ExcelE194 {
	ExcelE194NoIndicator = '\002R\000\000',
	ExcelE194CommentIndicatorOnly = '\002Q\377\377',
	ExcelE194CommentAndIndicator = '\002R\000\001'
};
typedef enum ExcelE194 ExcelE194;

enum ExcelE195 {
	ExcelE195CellValue = '\002S\000\001',
	ExcelE195Expression = '\002S\000\002',
	ExcelE195ColorScale = '\002S\000\003',
	ExcelE195Databar = '\002S\000\004',
	ExcelE195Top10 = '\002S\000\005',
	ExcelE195IconSets = '\002S\000\006',
	ExcelE195UniqueValues = '\002S\000\007',
	ExcelE195TextString = '\002S\000\011',
	ExcelE195BlanksCondition = '\002S\000\012',
	ExcelE195TimePeriod = '\002S\000\013',
	ExcelE195AboveAverageCondition = '\002S\000\014',
	ExcelE195NoBlanksCondition = '\002S\000\015',
	ExcelE195ErrorsCondition = '\002S\000\020',
	ExcelE195NoErrorsCondition = '\002S\000\021'
};
typedef enum ExcelE195 ExcelE195;

enum ExcelE196 {
	ExcelE196OperatorBetween = '\002T\000\001',
	ExcelE196OperatorNotBetween = '\002T\000\002',
	ExcelE196OperatorEqual = '\002T\000\003',
	ExcelE196OperatorNotEqual = '\002T\000\004',
	ExcelE196OperatorGreater = '\002T\000\005',
	ExcelE196OperatorLess = '\002T\000\006',
	ExcelE196OperatorGreaterEqual = '\002T\000\007',
	ExcelE196OperatorLessEqual = '\002T\000\010'
};
typedef enum ExcelE196 ExcelE196;

enum ExcelE197 {
	ExcelE197NoRestrictions = '\002U\000\000',
	ExcelE197UnlockedCells = '\002U\000\001',
	ExcelE197NoSelection = '\002T\357\322'
};
typedef enum ExcelE197 ExcelE197;

enum ExcelE198 {
	ExcelE198ValidateInputOnly = '\002V\000\000',
	ExcelE198ValidateWholeNumber = '\002V\000\001',
	ExcelE198ValidateDecimal = '\002V\000\002',
	ExcelE198ValidateList = '\002V\000\003',
	ExcelE198ValidatedDate = '\002V\000\004',
	ExcelE198ValidateTime = '\002V\000\005',
	ExcelE198ValidateTextLength = '\002V\000\006',
	ExcelE198ValidateCustom = '\002V\000\007'
};
typedef enum ExcelE198 ExcelE198;

enum ExcelE199 {
	ExcelE199IMEModeNoControl = '\002W\000\000',
	ExcelE199IMEModeOn = '\002W\000\001',
	ExcelE199IMEModeOff = '\002W\000\002',
	ExcelE199IMEModeDisable = '\002W\000\003',
	ExcelE199IMEModeHiragana = '\002W\000\004',
	ExcelE199IMEModeKatakana = '\002W\000\005',
	ExcelE199IMEModeKatakanaHalf = '\002W\000\006',
	ExcelE199IMEModeAlphaFull = '\002W\000\007',
	ExcelE199IMEModeAlpha = '\002W\000\010',
	ExcelE199IMEModeHangulFull = '\002W\000\011',
	ExcelE199IMEModeHangul = '\002W\000\012'
};
typedef enum ExcelE199 ExcelE199;

enum ExcelE200 {
	ExcelE200ValidAlertNone = '\002W\377\377',
	ExcelE200ValidAlertStop = '\002X\000\001',
	ExcelE200ValidAlertWarning = '\002X\000\002',
	ExcelE200ValidAlertInformation = '\002X\000\003'
};
typedef enum ExcelE200 ExcelE200;

enum ExcelE201 {
	ExcelE201LocationAsNewSheet = '\002Y\000\001',
	ExcelE201LocationAsObject = '\002Y\000\002',
	ExcelE201LocationAutomatic = '\002Y\000\003'
};
typedef enum ExcelE201 ExcelE201;

enum ExcelE216 {
	ExcelE216Automatic = '\002Y\357\367',
	ExcelE216Custom = '\002Y\357\356'
};
typedef enum ExcelE216 ExcelE216;

enum ExcelE900 {
	ExcelE900PivotTableVersion2000 = '\003\204\000\000',
	ExcelE900PivotTableVersion10 = '\003\204\000\001',
	ExcelE900PivotTableVersion11 = '\003\204\000\002',
	ExcelE900PivotTableVersion12 = '\003\204\000\003',
	ExcelE900PivotTableVersion14 = '\003\204\000\004',
	ExcelE900PivotTableVersionCurrent = '\003\203\377\377'
};
typedef enum ExcelE900 ExcelE900;

enum ExcelE901 {
	ExcelE901CompactRow = '\003\205\000\000',
	ExcelE901TabularRow = '\003\205\000\001',
	ExcelE901OutlineRow = '\003\205\000\002'
};
typedef enum ExcelE901 ExcelE901;

enum ExcelE902 {
	ExcelE902AtTop = '\003\206\000\001',
	ExcelE902AtBottom = '\003\206\000\002'
};
typedef enum ExcelE902 ExcelE902;

enum ExcelE903 {
	ExcelE903ManualAllocation = '\003\207\000\001',
	ExcelE903AutomaticAllocation = '\003\207\000\002'
};
typedef enum ExcelE903 ExcelE903;

enum ExcelE904 {
	ExcelE904AllocateValue = '\003\210\000\001',
	ExcelE904AllocateIncrement = '\003\210\000\002'
};
typedef enum ExcelE904 ExcelE904;

enum ExcelE905 {
	ExcelE905EqualAllocation = '\003\211\000\001',
	ExcelE905WeightAllocation = '\003\211\000\002'
};
typedef enum ExcelE905 ExcelE905;

enum ExcelE906 {
	ExcelE906DoNotRepeatLabels = '\003\212\000\001',
	ExcelE906RepeatLabels = '\003\212\000\002'
};
typedef enum ExcelE906 ExcelE906;

enum ExcelE907 {
	ExcelE907MissingItemsDefault = '\003\212\377\377',
	ExcelE907MissingItemsNone = '\003\213\000\000',
	ExcelE907MissingItemsMax = '\003\213~\364',
	ExcelE907MissingItemsMax2 = '\003\233\000\000'
};
typedef enum ExcelE907 ExcelE907;

enum ExcelE908 {
	ExcelE908PivotCellValue = '\003\214\000\000',
	ExcelE908PivotCellPivotItem = '\003\214\000\001',
	ExcelE908PivotCellSubtotal = '\003\214\000\002',
	ExcelE908PivotCellGrandTotal = '\003\214\000\003',
	ExcelE908PivotCellDataField = '\003\214\000\004',
	ExcelE908PivotCellPivotField = '\003\214\000\005',
	ExcelE908PivotCellPageFieldItem = '\003\214\000\006',
	ExcelE908PivotCellCustomSubtotal = '\003\214\000\007',
	ExcelE908PivotCellDataPivotField = '\003\214\000\010',
	ExcelE908PivotCellBlankCell = '\003\214\000\011'
};
typedef enum ExcelE908 ExcelE908;

enum ExcelE909 {
	ExcelE909CellNotChanged = '\003\215\000\001',
	ExcelE909CellChanged = '\003\215\000\002',
	ExcelE909CellChangeApplied = '\003\215\000\003'
};
typedef enum ExcelE909 ExcelE909;

enum ExcelE910 {
	ExcelE910Tabular = '\003\216\000\000',
	ExcelE910Outline = '\003\216\000\001'
};
typedef enum ExcelE910 ExcelE910;

enum ExcelE911 {
	ExcelE911PivotTopCount = '\003\217\000\001',
	ExcelE911PivotBottomCount = '\003\217\000\002',
	ExcelE911PivotTopPercent = '\003\217\000\003',
	ExcelE911PivotBottomPercent = '\003\217\000\004',
	ExcelE911PivotTopSum = '\003\217\000\005',
	ExcelE911PivotBottomSum = '\003\217\000\006',
	ExcelE911PivotValueEquals = '\003\217\000\007',
	ExcelE911PivotValueIsNotEqual = '\003\217\000\010',
	ExcelE911PivotValueIsGreaterThan = '\003\217\000\011',
	ExcelE911PivotValueIsGreaterThanOrEqualTo = '\003\217\000\012',
	ExcelE911PivotValueIsLessThan = '\003\217\000\013',
	ExcelE911PivotValueIsLessThanOrEqualTo = '\003\217\000\014',
	ExcelE911PivotValueIsBetween = '\003\217\000\015',
	ExcelE911PivotValueIsNotBetween = '\003\217\000\016',
	ExcelE911PivotCaptionEquals = '\003\217\000\017',
	ExcelE911PivotCaptionDoesNotEqual = '\003\217\000\020',
	ExcelE911PivotCaptionBeginsWith = '\003\217\000\021',
	ExcelE911PivotCaptionDoesNotBeginWith = '\003\217\000\022',
	ExcelE911PivotCaptionEndsWith = '\003\217\000\023',
	ExcelE911PivotCaptionDoesNotEndWith = '\003\217\000\024',
	ExcelE911PivotCaptionContains = '\003\217\000\025',
	ExcelE911PivotCaptionDoesNotContain = '\003\217\000\026',
	ExcelE911PivotCaptionIsGreaterThan = '\003\217\000\027',
	ExcelE911PivotCaptionIsGreaterThanOrEqualTo = '\003\217\000\030',
	ExcelE911PivotCaptionIsLessThan = '\003\217\000\031',
	ExcelE911PivotCaptionIsLessThanOrEqualTo = '\003\217\000\032',
	ExcelE911PivotCaptionIsBetween = '\003\217\000\033',
	ExcelE911PivotCaptionIsNowBetween = '\003\217\000\034',
	ExcelE911PivotSpecificDate = '\003\217\000\035',
	ExcelE911PivotNotSpecificDate = '\003\217\000\036',
	ExcelE911PivotBefore = '\003\217\000\037',
	ExcelE911PivotBeforeOrEqualTo = '\003\217\000 ',
	ExcelE911PivotAfter = '\003\217\000!',
	ExcelE911PivotAfterOrEqualTo = '\003\217\000\"',
	ExcelE911PivotBetween = '\003\217\000#',
	ExcelE911PivotNotBetween = '\003\217\000$',
	ExcelE911PivotTomorrow = '\003\217\000%',
	ExcelE911PivotToday = '\003\217\000&',
	ExcelE911PivotYesterday = '\003\217\000\'',
	ExcelE911PivotNextWeek = '\003\217\000(',
	ExcelE911PivotThisWeek = '\003\217\000)',
	ExcelE911PivotLastWeek = '\003\217\000*',
	ExcelE911PivotNextMonth = '\003\217\000+',
	ExcelE911PivotThisMonth = '\003\217\000,',
	ExcelE911PivotLastMonth = '\003\217\000-',
	ExcelE911PivotNextQuarter = '\003\217\000.',
	ExcelE911PivotThisQuarter = '\003\217\000/',
	ExcelE911PivotLastQuarter = '\003\217\0000',
	ExcelE911PivotNextYear = '\003\217\0001',
	ExcelE911PivotThisYear = '\003\217\0002',
	ExcelE911PivotLastYear = '\003\217\0003',
	ExcelE911PivotYearToDate = '\003\217\0004',
	ExcelE911PivotAllDatesInPeriodQuarter1 = '\003\217\0005',
	ExcelE911PivotAllDatesInPeriodQuarter2 = '\003\217\0006',
	ExcelE911PivotAllDatesInPeriodQuarter3 = '\003\217\0007',
	ExcelE911PivotAllDatesInPeriodQuarter4 = '\003\217\0008',
	ExcelE911PivotAllDatesInPeriodJanuary = '\003\217\0009',
	ExcelE911PivotAllDatesInPeriodFeberary = '\003\217\000:',
	ExcelE911PivotAllDatesInPeriodMarch = '\003\217\000;',
	ExcelE911PivotAllDatesInPeriodApril = '\003\217\000<',
	ExcelE911PivotAllDatesInPeriodMay = '\003\217\000=',
	ExcelE911PivotAllDatesInPeriodJune = '\003\217\000>',
	ExcelE911PivotAllDatesInPeriodJuly = '\003\217\000\?',
	ExcelE911PivotAllDatesInPeriodAugust = '\003\217\000@',
	ExcelE911PivotAllDatesInPeriodSeptember = '\003\217\000A',
	ExcelE911PivotAllDatesInPeriodOctober = '\003\217\000B',
	ExcelE911PivotAllDatesInPeriodNovember = '\003\217\000C',
	ExcelE911PivotAllDatesInPeriodDecember = '\003\217\000D'
};
typedef enum ExcelE911 ExcelE911;

enum ExcelE912 {
	ExcelE912PivotLineRegular = '\003\220\000\000',
	ExcelE912PivotLineSubtotal = '\003\220\000\001',
	ExcelE912PivotLineGrandtotal = '\003\220\000\002',
	ExcelE912PivotLineBlank = '\003\220\000\003'
};
typedef enum ExcelE912 ExcelE912;

enum ExcelE913 {
	ExcelE913Hierarchy = '\003\221\000\001',
	ExcelE913Measure = '\003\221\000\002',
	ExcelE913Set = '\003\221\000\003'
};
typedef enum ExcelE913 ExcelE913;

enum ExcelE914 {
	ExcelE914CubeHierarchy = '\003\222\000\001',
	ExcelE914CubeMeasure = '\003\222\000\002',
	ExcelE914CubeSet = '\003\222\000\003',
	ExcelE914CubeAttribute = '\003\222\000\004',
	ExcelE914CubeCalculatedMeasure = '\003\222\000\005',
	ExcelE914CubeKPIValue = '\003\222\000\006',
	ExcelE914CubeKPIGoal = '\003\222\000\007',
	ExcelE914CubeKPIStatus = '\003\222\000\010',
	ExcelE914CubeKPITrend = '\003\222\000\011',
	ExcelE914CubeKPIWeight = '\003\222\000\012'
};
typedef enum ExcelE914 ExcelE914;

enum ExcelE915 {
	ExcelE915DisplayPropertyInPivotTable = '\003\223\000\001',
	ExcelE915DisplayPropertyInTooltip = '\003\223\000\002',
	ExcelE915DisplayPropertyInPivotTableAndTooltip = '\003\223\000\003'
};
typedef enum ExcelE915 ExcelE915;

enum ExcelE916 {
	ExcelE916CalculatedMember = '\003\224\000\000',
	ExcelE916CalculatedSet = '\003\224\000\001'
};
typedef enum ExcelE916 ExcelE916;

enum ExcelE917 {
	ExcelE917ConnectionTypeOLEDB = '\003\225\000\001',
	ExcelE917ConnectionTypeODBC = '\003\225\000\002',
	ExcelE917ConnectionTypeXMLMAP = '\003\225\000\003',
	ExcelE917ConnectionTypeTEXT = '\003\225\000\004',
	ExcelE917ConnectionTypeWEB = '\003\225\000\005'
};
typedef enum ExcelE917 ExcelE917;

enum ExcelE203 {
	ExcelE203PasteSpecialOperationAdd = '\002[\000\002',
	ExcelE203PasteSpecialOperationDivide = '\002[\000\005',
	ExcelE203PasteSpecialOperationMultiply = '\002[\000\004',
	ExcelE203PasteSpecialOperationNone = '\002Z\357\322',
	ExcelE203PasteSpecialOperationSubtract = '\002[\000\003'
};
typedef enum ExcelE203 ExcelE203;

enum ExcelE204 {
	ExcelE204PasteAll = '\002[\357\370',
	ExcelE204PasteAllUsingSourceTheme = '\002\\\000\015',
	ExcelE204PasteAllExceptBorders = '\002\\\000\007',
	ExcelE204PasteFormats = '\002[\357\346',
	ExcelE204PasteFormulas = '\002[\357\345',
	ExcelE204PasteComments = '\002[\357\320',
	ExcelE204PasteValues = '\002[\357\275',
	ExcelE204PasteColumnWidths = '\002\\\000\010',
	ExcelE204PasteValidation = '\002\\\000\006',
	ExcelE204PasteFormulasAndNumberFormats = '\002\\\000\013',
	ExcelE204PasteValuesAndNumberFormats = '\002\\\000\014'
};
typedef enum ExcelE204 ExcelE204;

enum ExcelE205 {
	ExcelE205PhoneticCharacterHalfWidthKatakana = '\002]\000\000',
	ExcelE205PhoneticCharacterFullWidthKatakana = '\002]\000\001',
	ExcelE205PhoneticCharacterHiragana = '\002]\000\002',
	ExcelE205NoPhoneticCharacterConversion = '\002]\000\003'
};
typedef enum ExcelE205 ExcelE205;

enum ExcelE206 {
	ExcelE206PhoneticAlignNoControl = '\002^\000\000',
	ExcelE206PhoneticAlignLeft = '\002^\000\001',
	ExcelE206PhoneticAlignCenter = '\002^\000\002',
	ExcelE206PhoneticAlignDistributed = '\002^\000\003'
};
typedef enum ExcelE206 ExcelE206;

enum ExcelE207 {
	ExcelE207Printer = '\002_\000\002',
	ExcelE207Screen = '\002_\000\001'
};
typedef enum ExcelE207 ExcelE207;

enum ExcelE208 {
	ExcelE208OrientAsColumnField = '\002`\000\002',
	ExcelE208OrientAsDataField = '\002`\000\004',
	ExcelE208OrientAsHidden = '\002`\000\000',
	ExcelE208OrientAsPageField = '\002`\000\003',
	ExcelE208OrientAsRowField = '\002`\000\001'
};
typedef enum ExcelE208 ExcelE208;

enum ExcelE209 {
	ExcelE209PivotFieldCalculationDifferenceFrom = '\002a\000\002',
	ExcelE209PivotFieldCalculationIndex = '\002a\000\011',
	ExcelE209PivotFieldCalculationNoAdditionalCalculation = '\002`\357\321',
	ExcelE209PivotFieldCalculationPercentDifferenceFrom = '\002a\000\004',
	ExcelE209PivotFieldCalculationPercentOf = '\002a\000\003',
	ExcelE209PivotFieldCalculationPercentOfColumn = '\002a\000\007',
	ExcelE209PivotFieldCalculationPercentOfRow = '\002a\000\006',
	ExcelE209PivotFieldCalculationPercentOfTotal = '\002a\000\010',
	ExcelE209PivotFieldCalculationRunningTotal = '\002a\000\005'
};
typedef enum ExcelE209 ExcelE209;

enum ExcelE210 {
	ExcelE210PlacementFreeFloating = '\002b\000\003',
	ExcelE210PlacementMove = '\002b\000\002',
	ExcelE210PlacementMoveAndSize = '\002b\000\001'
};
typedef enum ExcelE210 ExcelE210;

enum ExcelE211 {
	ExcelE211Macintosh = '\002c\000\001',
	ExcelE211MSDos = '\002c\000\003',
	ExcelE211MSWindows = '\002c\000\002'
};
typedef enum ExcelE211 ExcelE211;

enum ExcelE212 {
	ExcelE212PrintSheetEnd = '\002d\000\001',
	ExcelE212PrintInPlace = '\002d\000\020',
	ExcelE212PrintNoComments = '\002c\357\322'
};
typedef enum ExcelE212 ExcelE212;

enum ExcelE213 {
	ExcelE213PriorityHigh = '\002d\357\341',
	ExcelE213PriorityLow = '\002d\357\332',
	ExcelE213PriorityNormal = '\002d\357\321'
};
typedef enum ExcelE213 ExcelE213;

enum ExcelE214 {
	ExcelE214SelectionModeLabelOnly = '\002f\000\001',
	ExcelE214SelectionModeDataAndLabel = '\002f\000\000',
	ExcelE214SelectionModeDataOnly = '\002f\000\002',
	ExcelE214SelectionModeOrigin = '\002f\000\003',
	ExcelE214SelectionModeButton = '\002f\000\017',
	ExcelE214SelectionModeBlanks = '\002f\000\004'
};
typedef enum ExcelE214 ExcelE214;

enum ExcelE215 {
	ExcelE215RangeAutoformatThreeDEffects1 = '\002g\000\015',
	ExcelE215RangeAutoformatThreeDEffects2 = '\002g\000\016',
	ExcelE215RangeAutoformatAccounting1 = '\002g\000\004',
	ExcelE215RangeAutoformatAccounting2 = '\002g\000\005',
	ExcelE215RangeAutoformatAccounting3 = '\002g\000\006',
	ExcelE215RangeAutoformatAccounting4 = '\002g\000\021',
	ExcelE215RangeAutoformatClassic1 = '\002g\000\001',
	ExcelE215RangeAutoformatClassic2 = '\002g\000\002',
	ExcelE215RangeAutoformatClassic3 = '\002g\000\003',
	ExcelE215RangeAutoformatColor1 = '\002g\000\007',
	ExcelE215RangeAutoformatColor2 = '\002g\000\010',
	ExcelE215RangeAutoformatColor3 = '\002g\000\011',
	ExcelE215RangeAutoformatList1 = '\002g\000\012',
	ExcelE215RangeAutoformatList2 = '\002g\000\013',
	ExcelE215RangeAutoformatList3 = '\002g\000\014',
	ExcelE215RangeAutoformatLocalFormat1 = '\002g\000\017',
	ExcelE215RangeAutoformatLocalFormat2 = '\002g\000\020',
	ExcelE215RangeAutoformatLocalFormat3 = '\002g\000\023',
	ExcelE215RangeAutoformatLocalFormat4 = '\002g\000\024',
	ExcelE215RangeAutoformatNone = '\002f\357\322',
	ExcelE215RangeAutoformatSimple = '\002f\357\306'
};
typedef enum ExcelE215 ExcelE215;

enum ExcelE217 {
	ExcelE217AllAtOnce = '\002i\000\002',
	ExcelE217OneAfterAnother = '\002i\000\001'
};
typedef enum ExcelE217 ExcelE217;

enum ExcelE218 {
	ExcelE218NotYetRouted = '\002j\000\000',
	ExcelE218RoutingComplete = '\002j\000\002',
	ExcelE218RoutingInProgress = '\002j\000\001'
};
typedef enum ExcelE218 ExcelE218;

enum ExcelE219 {
	ExcelE219AutoActivate = '\002k\000\003',
	ExcelE219AutoClose = '\002k\000\002',
	ExcelE219AutoDeactivate = '\002k\000\004',
	ExcelE219AutoOpen = '\002k\000\001'
};
typedef enum ExcelE219 ExcelE219;

enum ExcelE221 {
	ExcelE221Exclusive = '\002m\000\003',
	ExcelE221NoChange = '\002m\000\001',
	ExcelE221Shared = '\002m\000\002'
};
typedef enum ExcelE221 ExcelE221;

enum ExcelE222 {
	ExcelE222LocalSessionChanges = '\002n\000\002',
	ExcelE222OtherSessionChanges = '\002n\000\003',
	ExcelE222UserResolution = '\002n\000\001'
};
typedef enum ExcelE222 ExcelE222;

enum ExcelE223 {
	ExcelE223SearchNext = '\002o\000\001',
	ExcelE223SearchPrevious = '\002o\000\002'
};
typedef enum ExcelE223 ExcelE223;

enum ExcelE224 {
	ExcelE224ByColumns = '\002p\000\002',
	ExcelE224ByRows = '\002p\000\001'
};
typedef enum ExcelE224 ExcelE224;

enum ExcelE225 {
	ExcelE225SheetVisible = '\002p\377\377',
	ExcelE225SheetHidden = '\002q\000\000',
	ExcelE225SheetVeryHidden = '\002q\000\002'
};
typedef enum ExcelE225 ExcelE225;

enum ExcelE226 {
	ExcelE226PinYin = '\002r\000\001' /* Phonetic Chinese/Japanese sort order for characters. This is the default value. */,
	ExcelE226Stroke = '\002r\000\002' /* Sort by the quantity of strokes in each character. */
};
typedef enum ExcelE226 ExcelE226;

enum ExcelE228 {
	ExcelE228SortAscending = '\002t\000\001' /* Sorts the specified field in ascending order. This is the default value. */,
	ExcelE228SortDescending = '\002t\000\002' /* Sorts the specified field in descending order. */,
	ExcelE228SortManual = '\002s\357\331' /* It is not supported. */
};
typedef enum ExcelE228 ExcelE228;

enum ExcelE229 {
	ExcelE229SortRows = '\002u\000\002' /* Sorts by row. this is the default value. */,
	ExcelE229SortColumns = '\002u\000\001' /* Sorts by column. */
};
typedef enum ExcelE229 ExcelE229;

enum ExcelE230 {
	ExcelE230SortLabels = '\002v\000\002' /* Sorts the PivotTable report by labels. */,
	ExcelE230SortValues = '\002v\000\001' /* Sorts the PivotTable report by values. */
};
typedef enum ExcelE230 ExcelE230;

enum ExcelE231 {
	ExcelE231Errors = '\002w\000\020',
	ExcelE231Logical = '\002w\000\004',
	ExcelE231Numbers = '\002w\000\001',
	ExcelE231TextValues = '\002w\000\002'
};
typedef enum ExcelE231 ExcelE231;

enum ExcelE232 {
	ExcelE232SummaryAbove = '\002x\000\000',
	ExcelE232SummaryBelow = '\002x\000\001'
};
typedef enum ExcelE232 ExcelE232;

enum ExcelE233 {
	ExcelE233SummaryOnLeft = '\002x\357\335',
	ExcelE233SummaryOnRight = '\002x\357\310'
};
typedef enum ExcelE233 ExcelE233;

enum ExcelE234 {
	ExcelE234SummaryPivotTable = '\002y\357\314',
	ExcelE234StandardSummary = '\002z\000\001'
};
typedef enum ExcelE234 ExcelE234;

enum ExcelE236 {
	ExcelE236Delimited = '\002|\000\001',
	ExcelE236FixedWidth = '\002|\000\002'
};
typedef enum ExcelE236 ExcelE236;

enum ExcelE237 {
	ExcelE237TextQualifierDoubleQuote = '\002}\000\001',
	ExcelE237TextQualifierNone = '\002|\357\322',
	ExcelE237TextQualifierSingleQuote = '\002}\000\002'
};
typedef enum ExcelE237 ExcelE237;

enum ExcelE238 {
	ExcelE238Chart = '\002}\357\363',
	ExcelE238Excel4IntlMacroSheet = '\002~\000\004',
	ExcelE238Excel4MacroSheet = '\002~\000\003',
	ExcelE238Worksheet = '\002}\357\271'
};
typedef enum ExcelE238 ExcelE238;

enum ExcelE239 {
	ExcelE239NormalView = '\002\177\000\001',
	ExcelE239PageLayoutView = '\002\177\000\003'
};
typedef enum ExcelE239 ExcelE239;

enum ExcelE240 {
	ExcelE240MacroTypeCommand = '\002\200\000\002',
	ExcelE240MacroTypeFunction = '\002\200\000\001',
	ExcelE240MacroTypeNotXLM = '\002\200\000\003'
};
typedef enum ExcelE240 ExcelE240;

enum ExcelE241 {
	ExcelE241HeaderGuess = '\002\201\000\000' /* Default value. Excel determines whether there is a header, and where it is, if there is one. */,
	ExcelE241HeaderNo = '\002\201\000\002' /* The entire range should be sorted. */,
	ExcelE241HeaderYes = '\002\201\000\001' /* The entire range should not be sorted. */
};
typedef enum ExcelE241 ExcelE241;

enum ExcelE242 {
	ExcelE242DisplayShapes = '\002\201\357\370',
	ExcelE242Hide = '\002\202\000\003',
	ExcelE242Placeholders = '\002\202\000\002'
};
typedef enum ExcelE242 ExcelE242;

enum ExcelE243 {
	ExcelE243InsideHorizontal = '\002\203\000\014',
	ExcelE243InsideVertical = '\002\203\000\013',
	ExcelE243DiagonalDown = '\002\203\000\005',
	ExcelE243DiagonalUp = '\002\203\000\006',
	ExcelE243EdgeBottom = '\002\203\000\011',
	ExcelE243EdgeLeft = '\002\203\000\007',
	ExcelE243EdgeRight = '\002\203\000\012',
	ExcelE243EdgeTop = '\002\203\000\010',
	ExcelE243BorderBottom = '\002\202\357\365',
	ExcelE243BorderLeft = '\002\202\357\335',
	ExcelE243BorderRight = '\002\202\357\310',
	ExcelE243BorderTop = '\002\202\357\300'
};
typedef enum ExcelE243 ExcelE243;

enum ExcelE244 {
	ExcelE244NoButtonChanges = '\002\204\000\001',
	ExcelE244NoChanges = '\002\204\000\004',
	ExcelE244NoDockingChanges = '\002\204\000\003',
	ExcelE244ToolbarProtectionNone = '\002\203\357\321',
	ExcelE244NoShapeChanges = '\002\204\000\002'
};
typedef enum ExcelE244 ExcelE244;

enum ExcelE245 {
	ExcelE245DialogOpen = '\002\205\000\001',
	ExcelE245DialogOpenLinks = '\002\205\000\002',
	ExcelE245DialogSaveAs = '\002\205\000\005',
	ExcelE245DialogFileDelete = '\002\205\000\006',
	ExcelE245DialogPageSetup = '\002\205\000\007',
	ExcelE245DialogPrint = '\002\205\000\010',
	ExcelE245DialogPrinterSetup = '\002\205\000\011',
	ExcelE245DialogArrangeAll = '\002\205\000\014',
	ExcelE245DialogWindowSize = '\002\205\000\015',
	ExcelE245DialogWindowMove = '\002\205\000\016',
	ExcelE245DialogRun = '\002\205\000\021',
	ExcelE245DialogSetPrintTitles = '\002\205\000\027',
	ExcelE245DialogFont = '\002\205\000\032',
	ExcelE245DialogDisplay = '\002\205\000\033',
	ExcelE245DialogProtectDocument = '\002\205\000\034',
	ExcelE245DialogCalculation = '\002\205\000 ',
	ExcelE245DialogExtract = '\002\205\000#',
	ExcelE245DialogDataDelete = '\002\205\000$',
	ExcelE245DialogSort = '\002\205\000\'',
	ExcelE245DialogDataSeries = '\002\205\000(',
	ExcelE245DialogTable = '\002\205\000)',
	ExcelE245DialogFormatNumber = '\002\205\000*',
	ExcelE245DialogAlignment = '\002\205\000+',
	ExcelE245DialogStyle = '\002\205\000,',
	ExcelE245DialogBorder = '\002\205\000-',
	ExcelE245DialogCellProtection = '\002\205\000.',
	ExcelE245DialogColumnWidth = '\002\205\000/',
	ExcelE245DialogClear = '\002\205\0004',
	ExcelE245DialogPasteSpecial = '\002\205\0005',
	ExcelE245DialogEditDelete = '\002\205\0006',
	ExcelE245DialogInsert = '\002\205\0007',
	ExcelE245DialogPasteNames = '\002\205\000:',
	ExcelE245DialogDefineName = '\002\205\000=',
	ExcelE245DialogCreateNames = '\002\205\000>',
	ExcelE245DialogFormulaGoto = '\002\205\000\?',
	ExcelE245DialogFormulaFind = '\002\205\000@',
	ExcelE245DialogGalleryArea = '\002\205\000C',
	ExcelE245DialogGalleryBar = '\002\205\000D',
	ExcelE245DialogGalleryColumn = '\002\205\000E',
	ExcelE245DialogGalleryLine = '\002\205\000F',
	ExcelE245DialogGalleryPie = '\002\205\000G',
	ExcelE245DialogGalleryScatter = '\002\205\000H',
	ExcelE245DialogCombination = '\002\205\000I',
	ExcelE245DialogGridlines = '\002\205\000L',
	ExcelE245DialogAxes = '\002\205\000N',
	ExcelE245DialogAttachText = '\002\205\000P',
	ExcelE245DialogPatterns = '\002\205\000T',
	ExcelE245DialogMainChart = '\002\205\000U',
	ExcelE245DialogOverlay = '\002\205\000V',
	ExcelE245DialogScale = '\002\205\000W',
	ExcelE245DialogFormatLegend = '\002\205\000X',
	ExcelE245DialogFormatText = '\002\205\000Y',
	ExcelE245DialogParse = '\002\205\000[',
	ExcelE245DialogUnhide = '\002\205\000^',
	ExcelE245DialogWorkspace = '\002\205\000_',
	ExcelE245DialogActivate = '\002\205\000g',
	ExcelE245DialogCopyPicture = '\002\205\000l',
	ExcelE245DialogDeleteName = '\002\205\000n',
	ExcelE245DialogDeleteFormat = '\002\205\000o',
	ExcelE245DialogNew = '\002\205\000w',
	ExcelE245DialogRowHeight = '\002\205\000\177',
	ExcelE245DialogFormatMove = '\002\205\000\200',
	ExcelE245DialogFormatSize = '\002\205\000\201',
	ExcelE245DialogFormulaReplace = '\002\205\000\202',
	ExcelE245DialogSelectSpecial = '\002\205\000\204',
	ExcelE245DialogApplyNames = '\002\205\000\205',
	ExcelE245DialogReplaceFont = '\002\205\000\206',
	ExcelE245DialogSplit = '\002\205\000\211',
	ExcelE245DialogOutline = '\002\205\000\216',
	ExcelE245DialogSaveWorkbook = '\002\205\000\221',
	ExcelE245DialogCopyChart = '\002\205\000\223',
	ExcelE245DialogFormatFont = '\002\205\000\226',
	ExcelE245DialogNote = '\002\205\000\232',
	ExcelE245DialogSetUpdateStatus = '\002\205\000\237',
	ExcelE245DialogColorPalette = '\002\205\000\241',
	ExcelE245DialogChangeLink = '\002\205\000\246',
	ExcelE245DialogAppMove = '\002\205\000\252',
	ExcelE245DialogAppSize = '\002\205\000\253',
	ExcelE245DialogMainChartType = '\002\205\000\271',
	ExcelE245DialogOverlayChartType = '\002\205\000\272',
	ExcelE245DialogOpenMail = '\002\205\000\274',
	ExcelE245DialogSendMail = '\002\205\000\275',
	ExcelE245DialogStandardFont = '\002\205\000\276',
	ExcelE245DialogConsolidate = '\002\205\000\277',
	ExcelE245DialogSortSpecial = '\002\205\000\300',
	ExcelE245DialogGalleryThreeDArea = '\002\205\000\301',
	ExcelE245DialogGalleryThreeDColumn = '\002\205\000\302',
	ExcelE245DialogGalleryThreeDLine = '\002\205\000\303',
	ExcelE245DialogGalleryThreeDPie = '\002\205\000\304',
	ExcelE245DialogViewThreeD = '\002\205\000\305',
	ExcelE245DialogGoalSeek = '\002\205\000\306',
	ExcelE245DialogWorkgroup = '\002\205\000\307',
	ExcelE245DialogFillGroup = '\002\205\000\310',
	ExcelE245DialogUpdateLink = '\002\205\000\311',
	ExcelE245DialogPromote = '\002\205\000\312',
	ExcelE245DialogDemote = '\002\205\000\313',
	ExcelE245DialogShowDetail = '\002\205\000\314',
	ExcelE245DialogObjectProperties = '\002\205\000\317',
	ExcelE245DialogSaveNewObject = '\002\205\000\320',
	ExcelE245DialogApplyStyle = '\002\205\000\324',
	ExcelE245DialogAssignToObject = '\002\205\000\325',
	ExcelE245DialogObjectProtection = '\002\205\000\326',
	ExcelE245DialogShowToolbar = '\002\205\000\334',
	ExcelE245DialogPrintPreview = '\002\205\000\336',
	ExcelE245DialogEditColor = '\002\205\000\337',
	ExcelE245DialogFormatMain = '\002\205\000\341',
	ExcelE245DialogFormatOverlay = '\002\205\000\342',
	ExcelE245DialogEditSeries = '\002\205\000\344',
	ExcelE245DialogDefineStyle = '\002\205\000\345',
	ExcelE245DialogGalleryRadar = '\002\205\000\371',
	ExcelE245DialogZoom = '\002\205\001\000',
	ExcelE245DialogInsertObject = '\002\205\001\003',
	ExcelE245DialogSize = '\002\205\001\005',
	ExcelE245DialogMove = '\002\205\001\006',
	ExcelE245DialogFormatAuto = '\002\205\001\015',
	ExcelE245DialogGalleryThreeDBar = '\002\205\001\020',
	ExcelE245DialogGalleryThreeDSurface = '\002\205\001\021',
	ExcelE245DialogCustomizeToolbar = '\002\205\001\024',
	ExcelE245DialogWorkbookAdd = '\002\205\001\031',
	ExcelE245DialogWorkbookMove = '\002\205\001\032',
	ExcelE245DialogWorkbookCopy = '\002\205\001\033',
	ExcelE245DialogWorkbookOptions = '\002\205\001\034',
	ExcelE245DialogSaveWorkspace = '\002\205\001\035',
	ExcelE245DialogChartWizard = '\002\205\001 ',
	ExcelE245DialogAssignToTool = '\002\205\001%',
	ExcelE245DialogPlacement = '\002\205\001,',
	ExcelE245DialogFillWorkgroup = '\002\205\001-',
	ExcelE245DialogWorkbookNew = '\002\205\001.',
	ExcelE245DialogScenarioCells = '\002\205\0011',
	ExcelE245DialogScenarioAdd = '\002\205\0013',
	ExcelE245DialogScenarioEdit = '\002\205\0014',
	ExcelE245DialogScenarioSummary = '\002\205\0017',
	ExcelE245DialogPivotTableWizard = '\002\205\0018',
	ExcelE245DialogPivotFieldProperties = '\002\205\0019',
	ExcelE245DialogOptionsCalculation = '\002\205\001>',
	ExcelE245DialogOptionsEdit = '\002\205\001\?',
	ExcelE245DialogOptionsView = '\002\205\001@',
	ExcelE245DialogAddInManager = '\002\205\001A',
	ExcelE245DialogMenuEditor = '\002\205\001B',
	ExcelE245DialogAttachToolbars = '\002\205\001C',
	ExcelE245DialogOptionsChart = '\002\205\001E',
	ExcelE245DialogVbaInsertFile = '\002\205\001H',
	ExcelE245DialogVbaProcedureDefinition = '\002\205\001J',
	ExcelE245DialogRoutingSlip = '\002\205\001P',
	ExcelE245DialogMailLogon = '\002\205\001S',
	ExcelE245DialogInsertPicture = '\002\205\001V',
	ExcelE245DialogGalleryDoughnut = '\002\205\001X',
	ExcelE245DialogChartTrend = '\002\205\001^',
	ExcelE245DialogWorkbookInsert = '\002\205\001b',
	ExcelE245DialogOptionsTransition = '\002\205\001c',
	ExcelE245DialogOptionsGeneral = '\002\205\001d',
	ExcelE245DialogFilterAdvanced = '\002\205\001r',
	ExcelE245DialogMailNextLetter = '\002\205\001z',
	ExcelE245DialogDataLabel = '\002\205\001{',
	ExcelE245DialogInsertTitle = '\002\205\001|',
	ExcelE245DialogFontProperties = '\002\205\001}',
	ExcelE245DialogMacroOptions = '\002\205\001~',
	ExcelE245DialogWorkbookUnhide = '\002\205\001\200',
	ExcelE245DialogWorkbookName = '\002\205\001\202',
	ExcelE245DialogGalleryCustom = '\002\205\001\204',
	ExcelE245DialogAddChartAutoformat = '\002\205\001\206',
	ExcelE245DialogChartAddData = '\002\205\001\210',
	ExcelE245DialogTabOrder = '\002\205\001\212',
	ExcelE245DialogSubtotalCreate = '\002\205\001\216',
	ExcelE245DialogWorkbookTabSplit = '\002\205\001\237',
	ExcelE245DialogWorkbookProtect = '\002\205\001\241',
	ExcelE245DialogScrollbarProperties = '\002\205\001\244',
	ExcelE245DialogPivotShowPages = '\002\205\001\245',
	ExcelE245DialogTextToColumns = '\002\205\001\246',
	ExcelE245DialogFormatCharttype = '\002\205\001\247',
	ExcelE245DialogPivotFieldGroup = '\002\205\001\261',
	ExcelE245DialogPivotFieldUngroup = '\002\205\001\262',
	ExcelE245DialogCheckboxProperties = '\002\205\001\263',
	ExcelE245DialogLabelProperties = '\002\205\001\264',
	ExcelE245DialogListboxProperties = '\002\205\001\265',
	ExcelE245DialogEditboxProperties = '\002\205\001\266',
	ExcelE245DialogOpenText = '\002\205\001\271',
	ExcelE245DialogPushbuttonProperties = '\002\205\001\275',
	ExcelE245DialogFilter = '\002\205\001\277',
	ExcelE245DialogFunctionWizard = '\002\205\001\302',
	ExcelE245DialogSaveCopyAs = '\002\205\001\310',
	ExcelE245DialogOptionsListsAdd = '\002\205\001\312',
	ExcelE245DialogSeriesAxes = '\002\205\001\314',
	ExcelE245DialogSeriesX = '\002\205\001\315',
	ExcelE245DialogSeriesY = '\002\205\001\316',
	ExcelE245DialogErrorbarX = '\002\205\001\317',
	ExcelE245DialogErrorbarY = '\002\205\001\320',
	ExcelE245DialogFormatChart = '\002\205\001\321',
	ExcelE245DialogSeriesOrder = '\002\205\001\322',
	ExcelE245DialogMailEditMailer = '\002\205\001\326',
	ExcelE245DialogStandardWidth = '\002\205\001\330',
	ExcelE245DialogScenarioMerge = '\002\205\001\331',
	ExcelE245DialogProperties = '\002\205\001\332',
	ExcelE245DialogSummaryInfo = '\002\205\001\332',
	ExcelE245DialogFindFile = '\002\205\001\333',
	ExcelE245DialogActiveCellFont = '\002\205\001\334',
	ExcelE245DialogVbaMakeAddIn = '\002\205\001\336',
	ExcelE245DialogFileSharing = '\002\205\001\341',
	ExcelE245DialogAutocorrect = '\002\205\001\345',
	ExcelE245DialogCustomViews = '\002\205\001\355',
	ExcelE245DialogInsertNameLabel = '\002\205\001\360',
	ExcelE245DialogSeriesShape = '\002\205\001\370',
	ExcelE245DialogChartOptionsDataLabels = '\002\205\001\371',
	ExcelE245DialogChartOptionsDataTable = '\002\205\001\372',
	ExcelE245DialogSetBackgroundPicture = '\002\205\001\375',
	ExcelE245DialogDataValidation = '\002\205\002\015',
	ExcelE245DialogChartType = '\002\205\002\016',
	ExcelE245DialogChartLocation = '\002\205\002\017',
	ExcelE245DialogChartSourceData = '\002\205\002\035',
	ExcelE245DialogSeriesOptions = '\002\205\002-',
	ExcelE245DialogPivotTableOptions = '\002\205\0027',
	ExcelE245DialogPivotSolveOrder = '\002\205\0028',
	ExcelE245DialogPivotCalculatedField = '\002\205\002:',
	ExcelE245DialogPivotCalculatedItem = '\002\205\002<',
	ExcelE245DialogConditionalFormatting = '\002\205\002G',
	ExcelE245DialogInsertHyperlink = '\002\205\002T',
	ExcelE245DialogProtectSharing = '\002\205\002l',
	ExcelE245DialogPhonetic = '\002\205\002\213',
	ExcelE245DialogImportTextFile = '\002\205\002\232',
	ExcelE245DialogWebOptionsGeneral = '\002\205\002\264',
	ExcelE245DialogWebOptionsPictures = '\002\205\002\266',
	ExcelE245DialogWebOptionsFiles = '\002\205\002\265',
	ExcelE245DialogWebOptionsFonts = '\002\205\002\270',
	ExcelE245DialogWebOptionsEncoding = '\002\205\002\267'
};
typedef enum ExcelE245 ExcelE245;

enum ExcelE246 {
	ExcelE246Prompt = '\002\206\000\000',
	ExcelE246Constant = '\002\206\000\001',
	ExcelE246Range = '\002\206\000\002'
};
typedef enum ExcelE246 ExcelE246;

enum ExcelE247 {
	ExcelE247ParamTypeUnknown = '\002\207\000\000',
	ExcelE247ParamTypeChar = '\002\207\000\001',
	ExcelE247ParamTypeNumeric = '\002\207\000\002',
	ExcelE247ParamTypeDecimal = '\002\207\000\003',
	ExcelE247ParamTypeNumber = '\002\207\000\004',
	ExcelE247ParamTypeSmallInt = '\002\207\000\005',
	ExcelE247ParamTypeFloat = '\002\207\000\006',
	ExcelE247ParamTypeReal = '\002\207\000\007',
	ExcelE247ParamTypeDouble = '\002\207\000\010',
	ExcelE247ParamTypeVarChar = '\002\207\000\014',
	ExcelE247ParamTypeDate = '\002\207\000\011',
	ExcelE247ParamTypeTime = '\002\207\000\012',
	ExcelE247ParamTypeTimestamp = '\002\207\000\013',
	ExcelE247ParamTypeLongVarChar = '\002\206\377\377',
	ExcelE247ParamTypeBinary = '\002\206\377\376',
	ExcelE247ParamTypeVarBinary = '\002\206\377\375',
	ExcelE247ParamTypeLongVarBinary = '\002\206\377\374',
	ExcelE247ParamTypeBigInt = '\002\206\377\373',
	ExcelE247ParamTypeTinyInt = '\002\206\377\372',
	ExcelE247ParamTypeBit = '\002\206\377\371'
};
typedef enum ExcelE247 ExcelE247;

enum ExcelE248 {
	ExcelE248ButtonControl = '\002\210\000\000',
	ExcelE248CheckBox = '\002\210\000\001',
	ExcelE248DropDown = '\002\210\000\002',
	ExcelE248EditBox = '\002\210\000\003',
	ExcelE248GroupBox = '\002\210\000\004',
	ExcelE248Label = '\002\210\000\005',
	ExcelE248ListBox = '\002\210\000\006',
	ExcelE248OptionButton = '\002\210\000\007',
	ExcelE248ScrollBar = '\002\210\000\010',
	ExcelE248Spinner = '\002\210\000\011'
};
typedef enum ExcelE248 ExcelE248;

enum ExcelE249 {
	ExcelE249GeneralFormat = '\002\211\000\001',
	ExcelE249TextFormat = '\002\211\000\002',
	ExcelE249MDYFormat = '\002\211\000\003',
	ExcelE249DMYFormat = '\002\211\000\004',
	ExcelE249YMDFormat = '\002\211\000\005',
	ExcelE249MYDFormat = '\002\211\000\006',
	ExcelE249DYMFormat = '\002\211\000\007',
	ExcelE249YDMFormat = '\002\211\000\010',
	ExcelE249SkipColumn = '\002\211\000\011'
};
typedef enum ExcelE249 ExcelE249;

enum ExcelE250 {
	ExcelE250ODBCQuery = '\002\212\000\001',
	ExcelE250DAORecordSet = '\002\212\000\002',
	ExcelE250WebQuery = '\002\212\000\004',
	ExcelE250OLEDBQuery = '\002\212\000\005',
	ExcelE250TextImport = '\002\212\000\006',
	ExcelE250ADORecordset = '\002\212\000\007',
	ExcelE250FileMakerQuery = '\002\212\000\010'
};
typedef enum ExcelE250 ExcelE250;

enum ExcelE251 {
	ExcelE251CmdCube = '\002\213\000\001',
	ExcelE251CmdSql = '\002\213\000\002',
	ExcelE251CmdTable = '\002\213\000\003',
	ExcelE251CmdDefault = '\002\213\000\004',
	ExcelE251CmdList = '\002\213\000\005'
};
typedef enum ExcelE251 ExcelE251;

enum ExcelE253 {
	ExcelE253SrcNone = '\002\215\000\001',
	ExcelE253SrcRange = '\002\215\000\002',
	ExcelE253SrcExternal = '\002\215\000\003'
};
typedef enum ExcelE253 ExcelE253;

enum ExcelE257 {
	ExcelE257CriteriaEquals = '\002\221\000\000',
	ExcelE257CriteriaLessThanOrEqualTo = '\002\221\000\001',
	ExcelE257CriteriaGreaterThanOrEqualTo = '\002\221\000\002',
	ExcelE257CriteriaLessThan = '\002\221\000\003',
	ExcelE257CriteriaGreaterThan = '\002\221\000\004',
	ExcelE257CriteriaBeginsWith = '\002\221\000\005',
	ExcelE257CriteriaEndsWith = '\002\221\000\006',
	ExcelE257CriteriaContains = '\002\221\000\007'
};
typedef enum ExcelE257 ExcelE257;

enum ExcelE258 {
	ExcelE258NoCondition = '\002\222\000\000',
	ExcelE258AndCondition = '\002\222\000\001',
	ExcelE258OrCondition = '\002\222\000\002'
};
typedef enum ExcelE258 ExcelE258;

enum ExcelE259 {
	ExcelE259RangeValueDefault = '\002\223\000\012',
	ExcelE259RangeValueXMLSpreadsheet = '\002\223\000\013',
	ExcelE259RangeValueMSPersistXML = '\002\223\000\014'
};
typedef enum ExcelE259 ExcelE259;

enum ExcelE260 {
	ExcelE260Inches = '\002\224\000\001',
	ExcelE260Centimeters = '\002\224\000\002',
	ExcelE260Millimeters = '\002\224\000\003'
};
typedef enum ExcelE260 ExcelE260;

enum ExcelE261 {
	ExcelE261SubtotalAutomatic = '\002\225\000\001',
	ExcelE261SubtotalSum = '\002\225\000\002',
	ExcelE261SubtotalCount = '\002\225\000\003',
	ExcelE261SubtotalAverage = '\002\225\000\004',
	ExcelE261SubtotalMax = '\002\225\000\005',
	ExcelE261SubtotalMin = '\002\225\000\006',
	ExcelE261SubtotalProduct = '\002\225\000\007',
	ExcelE261SubtotalCountNumbers = '\002\225\000\010',
	ExcelE261SubtotalStandardDeviation = '\002\225\000\011',
	ExcelE261SubtotalStandardDeviationP = '\002\225\000\012',
	ExcelE261SubtotalVariable = '\002\225\000\013',
	ExcelE261SubtotalVariableP = '\002\225\000\014'
};
typedef enum ExcelE261 ExcelE261;

enum ExcelE262 {
	ExcelE262DataEntryOn = '\002\226\000\001',
	ExcelE262DataEntryStrict = '\002\226\000\002',
	ExcelE262DataEntryOff = '\002\225\357\316'
};
typedef enum ExcelE262 ExcelE262;

enum ExcelE263 {
	ExcelE263StatusText = '\002\226\377\377',
	ExcelE263ABoolean = '\002\227\000\000'
};
typedef enum ExcelE263 ExcelE263;

enum ExcelE264 {
	ExcelE264ExcelMenus = '\002\230\000\001'
};
typedef enum ExcelE264 ExcelE264;

enum ExcelE265 {
	ExcelE265LeftToRight = '\002\230\354u',
	ExcelE265RightToLeft = '\002\230\354t',
	ExcelE265Context = '\002\230\354v'
};
typedef enum ExcelE265 ExcelE265;

enum ExcelE266 {
	ExcelE266NormalCursor = '\002\232\000\000',
	ExcelE266LogicalCursor = '\002\232\000\001',
	ExcelE266VisualCursor = '\002\232\000\002'
};
typedef enum ExcelE266 ExcelE266;

enum ExcelE267 {
	ExcelE267RangeObject = '\002\233\000\001' /* range object */,
	ExcelE267A1StyleRangeReference = '\002\233\000\002' /* range R1C1 reference */,
	ExcelE267NamedRange = '\002\233\000\003' /* range R1C1 reference */
};
typedef enum ExcelE267 ExcelE267;

enum ExcelE268 {
	ExcelE268AutomaticSubtotal = '\002\234\000\001',
	ExcelE268SumSubtotal = '\002\234\000\002',
	ExcelE268CountSubtotal = '\002\234\000\003',
	ExcelE268AverageSubtotal = '\002\234\000\004',
	ExcelE268MaximumValue = '\002\234\000\005',
	ExcelE268MinimumValue = '\002\234\000\006',
	ExcelE268ProductSubtotal = '\002\234\000\007',
	ExcelE268CountNumbersSubtotal = '\002\234\000\010',
	ExcelE268StandardDeviation = '\002\234\000\011',
	ExcelE268StandardDeviationP = '\002\234\000\012',
	ExcelE268VarianceSubtotal = '\002\234\000\013',
	ExcelE268VariancePSubtotal = '\002\234\000\014'
};
typedef enum ExcelE268 ExcelE268;

enum ExcelE269 {
	ExcelE269Type_automatic = '\002\234\357\367',
	ExcelE269Type_manual = '\002\234\357\331'
};
typedef enum ExcelE269 ExcelE269;

enum ExcelE270 {
	ExcelE270PositionTop = '\002\235\357\300',
	ExcelE270PositionBottom = '\002\235\357\365'
};
typedef enum ExcelE270 ExcelE270;

enum ExcelE271 {
	ExcelE271ScrollTabPositionFirst = '\002\237\000\000',
	ExcelE271ScrollTabPositionLast = '\002\237\000\001'
};
typedef enum ExcelE271 ExcelE271;

enum ExcelE272 {
	ExcelE272Range = '\002\240\000\000',
	ExcelE272AListOfRanges = '\002\240\000\001',
	ExcelE272ReportName = '\002\240\000\002',
	ExcelE272AListOfStringThatIsASQLQuery = '\002\240\000\003'
};
typedef enum ExcelE272 ExcelE272;

enum ExcelE273 {
	ExcelE273BuiltInChartTemplate = '\002\241\000\001',
	ExcelE273FormatName = '\002\241\000\002'
};
typedef enum ExcelE273 ExcelE273;

enum ExcelE274 {
	ExcelE274BuiltInChartType = '\002\242\000\025',
	ExcelE274CustomChart = '\002\241\357\356'
};
typedef enum ExcelE274 ExcelE274;

enum ExcelE275 {
	ExcelE275RangeObject = '\002\243\000\001' /* range object */,
	ExcelE275A1StyleRangeReference = '\002\243\000\002' /* range R1C1 reference */,
	ExcelE275NamedRange = '\002\243\000\003' /* range R1C1 reference */,
	ExcelE275ListOfStrings = '\002\243\000\004'
};
typedef enum ExcelE275 ExcelE275;

enum ExcelE276 {
	ExcelE276RangeObject = '\002\244\000\001' /* range object */,
	ExcelE276A1StyleRangeReference = '\002\244\000\002' /* range R1C1 reference */,
	ExcelE276NamedRange = '\002\244\000\003' /* range R1C1 reference */,
	ExcelE276InputDefaultAsString = '\002\244\000\004'
};
typedef enum ExcelE276 ExcelE276;

enum ExcelE277 {
	ExcelE277ANumber = '\002\245\000\001' /* range object */,
	ExcelE277InputTypeAsString = '\002\245\000\002' /* range R1C1 reference */,
	ExcelE277ANumberOrAString = '\002\245\000\003' /* range R1C1 reference */,
	ExcelE277ABool = '\002\245\000\004',
	ExcelE277RangeObject = '\002\245\000\010',
	ExcelE277ListOfNumbers = '\002\245\000A',
	ExcelE277ListOfStrings = '\002\245\000B',
	ExcelE277ListOfNumberOrString = '\002\245\000C',
	ExcelE277ListOfBools = '\002\245\000D',
	ExcelE277ListOfRangeObjects = '\002\245\000H'
};
typedef enum ExcelE277 ExcelE277;

enum ExcelE278 {
	ExcelE278ANumber = '\002\246\000\001',
	ExcelE278ABool = '\002\246\000\004'
};
typedef enum ExcelE278 ExcelE278;

enum ExcelE279 {
	ExcelE279RangeObject = '\002\247\000\001' /* range object */,
	ExcelE279A1StyleRangeReference = '\002\247\000\002' /* range R1C1 reference */,
	ExcelE279NamedRange = '\002\247\000\003' /* range R1C1 reference */,
	ExcelE279ListOfStrings = '\002\247\000\004' /* A list of SQL query strings */
};
typedef enum ExcelE279 ExcelE279;

enum ExcelE280 {
	ExcelE280Percentable = '\002\250\000\001' /* A percentage between 10 and 400 */,
	ExcelE280ABool = '\002\250\000\004'
};
typedef enum ExcelE280 ExcelE280;

enum ExcelE281 {
	ExcelE281Script = '\002\251\000\001' /* A script object */,
	ExcelE281ScriptText = '\002\251\000\002'
};
typedef enum ExcelE281 ExcelE281;

enum ExcelE282 {
	ExcelE282Application = '\002\252\000\001',
	ExcelE282Worksheet = '\002\252\000\002',
	ExcelE282A1StyleRangeReference = '\002\252\000\003'
};
typedef enum ExcelE282 ExcelE282;

enum ExcelE283 {
	ExcelE283HorizontalAligmentBottom = '\002\252\357\365',
	ExcelE283HorizontalAligmentLeft = '\002\252\357\335',
	ExcelE283HorizontalAligmentRight = '\002\252\357\310',
	ExcelE283HorizontalAligmentTop = '\002\252\357\300'
};
typedef enum ExcelE283 ExcelE283;

enum ExcelE284 {
	ExcelE284VerticalAlignmentTop = '\002\253\357\300',
	ExcelE284VerticalAlignmentCenter = '\002\253\357\364',
	ExcelE284VerticalAlignmentBottom = '\002\253\357\365',
	ExcelE284VerticalAlignmentJustify = '\002\253\357\336',
	ExcelE284VerticalAlignmentDistributed = '\002\253\357\353'
};
typedef enum ExcelE284 ExcelE284;

enum ExcelE285 {
	ExcelE285CheckboxOff = '\002\254\357\316',
	ExcelE285CheckboxOn = '\002\255\000\001',
	ExcelE285CheckboxMixed = '\002\255\000\002'
};
typedef enum ExcelE285 ExcelE285;

enum ExcelE286 {
	ExcelE286Text = '\002\255\357\302',
	ExcelE286ANumber = '\002\256\000\002',
	ExcelE286XlNumber = '\002\255\357\317',
	ExcelE286Reference = '\002\256\000\004',
	ExcelE286Formula = '\002\256\000\005'
};
typedef enum ExcelE286 ExcelE286;

enum ExcelE290 {
	ExcelE290SelectNone = '\002\261\357\322',
	ExcelE290SelectSimple = '\002\261\357\306',
	ExcelE290SelectExtended = '\002\262\000\003'
};
typedef enum ExcelE290 ExcelE290;

enum ExcelE291 {
	ExcelE291TextToReplace = '\002\263\000\001',
	ExcelE291ReplacementText = '\002\263\000\002'
};
typedef enum ExcelE291 ExcelE291;

enum ExcelE292 {
	ExcelE292RangeObject = '\002\264\000\001' /* range object */,
	ExcelE292A1StyleRangeReference = '\002\264\000\002' /* range R1C1 reference */,
	ExcelE292NamedRange = '\002\264\000\003' /* range R1C1 reference */,
	ExcelE292ListOfCategoryNames = '\002\264\000\004' /* A list category names */
};
typedef enum ExcelE292 ExcelE292;

enum ExcelE294 {
	ExcelE294DoNotUpdateLinks = '\002\266\000\000',
	ExcelE294UpdateExternalLinksOnly = '\002\266\000\001',
	ExcelE294UpdateRemoteLinksOnly = '\002\266\000\002',
	ExcelE294UpdateRemoteAndExternalLinks = '\002\266\000\003'
};
typedef enum ExcelE294 ExcelE294;

enum ExcelE295 {
	ExcelE295TabDelimiter = '\002\267\000\001',
	ExcelE295CommasDelimiter = '\002\267\000\002',
	ExcelE295SpacesDelimiter = '\002\267\000\003',
	ExcelE295SemicolonDelimiter = '\002\267\000\004',
	ExcelE295NoDelimiter = '\002\267\000\005',
	ExcelE295CustomCharacterDelimiter = '\002\267\000\006'
};
typedef enum ExcelE295 ExcelE295;

enum ExcelE296 {
	ExcelE296VaryByColor = '\002\270\000\001',
	ExcelE296VaryByShade = '\002\270\000\002',
	ExcelE296VaryByGrayscale = '\002\270\000\003',
	ExcelE296VaryBySameColor = '\002\270\000\004'
};
typedef enum ExcelE296 ExcelE296;

enum ExcelE297 {
	ExcelE297RangeObject = '\002\271\000\001' /* range object */,
	ExcelE297A1StyleRangeReference = '\002\271\000\002' /* range R1C1 reference */,
	ExcelE297NamedRange = '\002\271\000\003' /* range R1C1 reference */
};
typedef enum ExcelE297 ExcelE297;

enum ExcelE298 {
	ExcelE298WorksheetObject = '\002\272\000\001' /* worksheet object */,
	ExcelE298WorksheetName = '\002\272\000\002' /* worksheet name */
};
typedef enum ExcelE298 ExcelE298;

enum ExcelE299 {
	ExcelE299AlignTickLabelCenter = '\002\272\357\364',
	ExcelE299AlignTickLabelLeft = '\002\272\357\335',
	ExcelE299AlignTickLabelRight = '\002\272\357\310'
};
typedef enum ExcelE299 ExcelE299;

enum ExcelE300 {
	ExcelE300Basque = '\002\274\004-',
	ExcelE300Catalan = '\002\274\004\003',
	ExcelE300Chinese = '\002\274\010\004',
	ExcelE300ChineseTaiwan = '\002\274\004\004',
	ExcelE300Czech = '\002\274\004\005',
	ExcelE300Danish = '\002\274\004\006',
	ExcelE300Dutch = '\002\274\004\023',
	ExcelE300EnglishUS = '\002\274\004\011',
	ExcelE300EnglishAUS = '\002\274\014\011',
	ExcelE300EnglishBritish = '\002\274\010\011',
	ExcelE300EnglishCAN = '\002\274\020\011',
	ExcelE300Finnish = '\002\274\004\013',
	ExcelE300French = '\002\274\004\014',
	ExcelE300FrenchCanadian = '\002\274\014\014',
	ExcelE300German = '\002\274\004\007',
	ExcelE300GermanAustria = '\002\274\014\007',
	ExcelE300SwissGerman = '\002\274\010\007',
	ExcelE300Greek = '\002\274\004\010',
	ExcelE300Hungarian = '\002\274\004\016',
	ExcelE300Italian = '\002\274\004\020',
	ExcelE300Japanese = '\002\274\004\021',
	ExcelE300Korean = '\002\274\004\022',
	ExcelE300Malaysian = '\002\274\004>',
	ExcelE300NorwegianBokmal = '\002\274\004\024',
	ExcelE300Norwegian = '\002\274\004,',
	ExcelE300Polish = '\002\274\004\025',
	ExcelE300PortugueseBrazil = '\002\274\004\026',
	ExcelE300PortugueseIberian = '\002\274\010\026',
	ExcelE300Russian = '\002\274\004\031',
	ExcelE300Slovak = '\002\274\004\033',
	ExcelE300Slovenian = '\002\274\004$',
	ExcelE300Spanish = '\002\274\004\012',
	ExcelE300Swedish = '\002\274\004\035',
	ExcelE300Turkish = '\002\274\004\037'
};
typedef enum ExcelE300 ExcelE300;

enum ExcelE301 {
	ExcelE301SortOnCellValue = '\002\275\000\000' /* Values. */,
	ExcelE301SortOnCellColor = '\002\275\000\001' /* Cell color. */,
	ExcelE301SortOnFontColor = '\002\275\000\002' /* Font color. */,
	ExcelE301SortOnIcon = '\002\275\000\003' /* Icon. */
};
typedef enum ExcelE301 ExcelE301;

enum ExcelE302 {
	ExcelE302SortNormal = '\002\276\000\000' /* Default. Sorts numeric and text data separately. */,
	ExcelE302SortTextAsNumbers = '\002\276\000\001' /* Treat text as numeric data for the sort. */
};
typedef enum ExcelE302 ExcelE302;

enum ExcelE303 {
	ExcelE303NoneTotalsCalc = '\002\277\000\000',
	ExcelE303AverageTotalsCalc = '\002\277\000\001',
	ExcelE303CountTotalsCalc = '\002\277\000\002',
	ExcelE303CountNumberTotalsCalc = '\002\277\000\003',
	ExcelE303MaxTotalsCalc = '\002\277\000\004',
	ExcelE303MinTotalsCalc = '\002\277\000\005',
	ExcelE303SumTotalsCalc = '\002\277\000\006',
	ExcelE303DeviationTotalsCalc = '\002\277\000\007',
	ExcelE303VarTotalsCalc = '\002\277\000\010',
	ExcelE303CustomTotalsCalc = '\002\277\000\011'
};
typedef enum ExcelE303 ExcelE303;

enum ExcelCCET {
	ExcelCCETNoChartTitle = '\003\206\000\000',
	ExcelCCETChartTitleCenteredOverlay = '\003\206\000\001',
	ExcelCCETChartTitleAboveChart = '\003\206\000\002',
	ExcelCCETNoLegend = '\003\206\000d',
	ExcelCCETLegendRight = '\003\206\000e',
	ExcelCCETLegendTop = '\003\206\000f',
	ExcelCCETLegendLeft = '\003\206\000g',
	ExcelCCETLegendBottom = '\003\206\000h',
	ExcelCCETLegendRightOverlay = '\003\206\000i',
	ExcelCCETLegendLeftOverlay = '\003\206\000j',
	ExcelCCETNoDataLabel = '\003\206\000\310',
	ExcelCCETShowDataLabel = '\003\206\000\311',
	ExcelCCETDataLabelCenter = '\003\206\000\312',
	ExcelCCETDataLabelInsideEnd = '\003\206\000\313',
	ExcelCCETDataLabelInsideBase = '\003\206\000\314',
	ExcelCCETDataLabelOutsideEnd = '\003\206\000\315',
	ExcelCCETDataLabelLeft = '\003\206\000\316',
	ExcelCCETDataLabelRight = '\003\206\000\317',
	ExcelCCETDataLabelTop = '\003\206\000\320',
	ExcelCCETDataLabelBottom = '\003\206\000\321',
	ExcelCCETDataLabelBestFit = '\003\206\000\322',
	ExcelCCETNoPrimaryCategoryAxisTitle = '\003\206\001,',
	ExcelCCETPrimaryCategoryAxisTitleAdjacentToAxis = '\003\206\001-',
	ExcelCCETPrimaryCategoryAxisTitleBelowAxis = '\003\206\001.',
	ExcelCCETPrimaryCategoryAxisTitleRotated = '\003\206\001/',
	ExcelCCETPrimaryCategoryAxisTitleVertical = '\003\206\0010',
	ExcelCCETPrimaryCategoryAxisTitleHorizontal = '\003\206\0011',
	ExcelCCETNoPrimaryValueAxisTitle = '\003\206\0012',
	ExcelCCETPrimaryValueAxisTitleAdjacentToAxis = '\003\206\0013',
	ExcelCCETPrimaryValueAxisTitleBelowAxis = '\003\206\0014',
	ExcelCCETPrimaryValueAxisTitleRotated = '\003\206\0015',
	ExcelCCETPrimaryValueAxisTitleVertical = '\003\206\0016',
	ExcelCCETPrimaryValueAxisTitleHorizontal = '\003\206\0017',
	ExcelCCETNoSecondaryCategoryAxisTitle = '\003\206\0018',
	ExcelCCETSecondaryCategoryAxisTitleAdjacentToAxis = '\003\206\0019',
	ExcelCCETSecondaryCategoryAxisTitleBelowAxis = '\003\206\001:',
	ExcelCCETSecondaryCategoryAxisTitleRotated = '\003\206\001;',
	ExcelCCETSecondaryCategoryAxisTitleVertical = '\003\206\001<',
	ExcelCCETSecondaryCategoryAxisTitleHorizontal = '\003\206\001=',
	ExcelCCETNoSecondaryValueAxisTitle = '\003\206\001>',
	ExcelCCETSecondaryValueAxisTitleAdjacentToAxis = '\003\206\001\?',
	ExcelCCETSecondaryValueAxisTitleBelowAxis = '\003\206\001@',
	ExcelCCETSecondaryValueAxisTitleRotated = '\003\206\001A',
	ExcelCCETSecondaryValueAxisTitleVertical = '\003\206\001B',
	ExcelCCETSecondaryValueAxisTitleHorizontal = '\003\206\001C',
	ExcelCCETNoSeriesAxisTitle = '\003\206\001D',
	ExcelCCETSeriesAxisTitleRotated = '\003\206\001E',
	ExcelCCETSeriesAxisTitleVertical = '\003\206\001F',
	ExcelCCETSeriesAxisTitleHorizontal = '\003\206\001G',
	ExcelCCETNoPrimaryValueGridLines = '\003\206\001H',
	ExcelCCETPrimaryValueGridLinesMinor = '\003\206\001I',
	ExcelCCETPrimaryValueGridLinesMajor = '\003\206\001J',
	ExcelCCETPrimaryValueGridLinesMinorMajor = '\003\206\001K',
	ExcelCCETNoPrimaryCategoryGridLines = '\003\206\001L',
	ExcelCCETPrimaryCategoryGridLinesMinor = '\003\206\001M',
	ExcelCCETPrimaryCategoryGridLinesMajor = '\003\206\001N',
	ExcelCCETPrimaryCategoryGridLinesMinorMajor = '\003\206\001O',
	ExcelCCETNoSecondaryValueGridLines = '\003\206\001P',
	ExcelCCETSecondaryValueGridLinesMinor = '\003\206\001Q',
	ExcelCCETSecondaryValueGridLinesMajor = '\003\206\001R',
	ExcelCCETSecondaryValueGridLinesMinorMajor = '\003\206\001S',
	ExcelCCETNoSecondaryCategoryGridLines = '\003\206\001T',
	ExcelCCETSecondaryCategoryGridLinesMinor = '\003\206\001U',
	ExcelCCETSecondaryCategoryGridLinesMajor = '\003\206\001V',
	ExcelCCETSecondaryCategoryGridLinesMinorMajor = '\003\206\001W',
	ExcelCCETNoSeriesAxisGridLines = '\003\206\001X',
	ExcelCCETSeriesAxisGridLinesMinor = '\003\206\001Y',
	ExcelCCETSeriesAxisGridLinesMajor = '\003\206\001Z',
	ExcelCCETSeriesAxisGridLinesMinorMajor = '\003\206\001[',
	ExcelCCETNoPrimaryCategoryAxis = '\003\206\001\\',
	ExcelCCETPrimaryCategoryAxisShow = '\003\206\001]',
	ExcelCCETPrimaryCategoryAxisWithoutLabels = '\003\206\001^',
	ExcelCCETPrimaryCategoryAxisReverse = '\003\206\001_',
	ExcelCCETNoPrimaryValueAxis = '\003\206\001`',
	ExcelCCETShowPrimaryValueAxis = '\003\206\001a',
	ExcelCCETPrimaryValueAxisThousands = '\003\206\001b',
	ExcelCCETPrimaryValueAxisMillions = '\003\206\001c',
	ExcelCCETPrimaryValueAxisBillions = '\003\206\001d',
	ExcelCCETPrimaryValueAxisLogScale = '\003\206\001e',
	ExcelCCETNoSecondaryCategoryAxis = '\003\206\001f',
	ExcelCCETShowSecondaryCategoryAxis = '\003\206\001g',
	ExcelCCETSecondaryCategoryAxisWithoutLabels = '\003\206\001h',
	ExcelCCETSecondaryCategoryAxisReverse = '\003\206\001i',
	ExcelCCETNoSecondaryValueAxis = '\003\206\001j',
	ExcelCCETShowSecondaryValueAxis = '\003\206\001k',
	ExcelCCETSecondaryValueAxisThousands = '\003\206\001l',
	ExcelCCETSecondaryValueAxisMillions = '\003\206\001m',
	ExcelCCETSecondaryValueAxisBillions = '\003\206\001n',
	ExcelCCETSecondaryValueAxisLogScale = '\003\206\001o',
	ExcelCCETNoSeriesAxis = '\003\206\001p',
	ExcelCCETShowSeriesAxis = '\003\206\001q',
	ExcelCCETSeriesAxisWithoutLabeling = '\003\206\001r',
	ExcelCCETSeriesAxisReverse = '\003\206\001s',
	ExcelCCETPrimaryCategoryAxisThousands = '\003\206\001t',
	ExcelCCETPrimaryCategoryAxisMillions = '\003\206\001u',
	ExcelCCETPrimaryCategoryAxisBillions = '\003\206\001v',
	ExcelCCETPrimaryCategoryAxisLogScale = '\003\206\001w',
	ExcelCCETSecondaryCategoryAxisThousands = '\003\206\001x',
	ExcelCCETSecondaryCategoryAxisMillions = '\003\206\001y',
	ExcelCCETSecondaryCategoryAxisBillions = '\003\206\001z',
	ExcelCCETSecondaryCategoryAxisLogScale = '\003\206\001{',
	ExcelCCETNoDataTable = '\003\206\001\364',
	ExcelCCETShowDataTable = '\003\206\001\365',
	ExcelCCETDataTableWithLegendKeys = '\003\206\001\366',
	ExcelCCETNoTrendline = '\003\206\002X',
	ExcelCCETTrendlineAddLinear = '\003\206\002Y',
	ExcelCCETTrendlineAddExponential = '\003\206\002Z',
	ExcelCCETTrendlineAddLinearForecast = '\003\206\002[',
	ExcelCCETTrendlineAddTwoPeriodMovingAverage = '\003\206\002\\',
	ExcelCCETNoErrorBar = '\003\206\002\274',
	ExcelCCETErrorBarStandardError = '\003\206\002\275',
	ExcelCCETErrorBarPercentage = '\003\206\002\276',
	ExcelCCETErrorBarStandardDeviation = '\003\206\002\277',
	ExcelCCETNoLine = '\003\206\003 ',
	ExcelCCETLineDropLine = '\003\206\003!',
	ExcelCCETLineHiLoLine = '\003\206\003\"',
	ExcelCCETLineSeriesLine = '\003\206\003#',
	ExcelCCETLineDropHiloLine = '\003\206\003$',
	ExcelCCETNoUpDownBars = '\003\206\003\204',
	ExcelCCETShowUpDownBars = '\003\206\003\205',
	ExcelCCETNoPlotArea = '\003\206\003\350',
	ExcelCCETShowPlotArea = '\003\206\003\351',
	ExcelCCETNoChartWall = '\003\206\004L',
	ExcelCCETShowChartWall = '\003\206\004M',
	ExcelCCETNoChartFloor = '\003\206\004\260',
	ExcelCCETShowChartFloor = '\003\206\004\261'
};
typedef enum ExcelCCET ExcelCCET;

enum ExcelXDFC {
	ExcelXDFCFilterAboveAverage = '\003\207\000!' /* Filter all above-average values. */,
	ExcelXDFCFilterAllDatesInApril = '\003\207\000\030' /* Filter all dates in April. */,
	ExcelXDFCFilterAllDatesInAugust = '\003\207\000\034' /* Filter all dates in August. */,
	ExcelXDFCFilterAllDatesInDecember = '\003\207\000 ' /* Filter all dates in December. */,
	ExcelXDFCFilterAllDatesInFebruary = '\003\207\000\026' /* Filter all dates in February */,
	ExcelXDFCFilterAllDatesInJanuary = '\003\207\000\025' /* Filter all dates in January. */,
	ExcelXDFCFilterAllDatesInJuly = '\003\207\000\033' /* Filter all dates in July. */,
	ExcelXDFCFilterAllDatesInJune = '\003\207\000\032' /* Filter all dates in June. */,
	ExcelXDFCFilterAllDatesInMarch = '\003\207\000\027' /* Filter all dates in March. */,
	ExcelXDFCFilterAllDatesInMay = '\003\207\000\031' /* Filter all dates in May. */,
	ExcelXDFCFilterAllDatesInNovember = '\003\207\000\037' /* Filter all dates in November. */,
	ExcelXDFCFilterAllDatesInOctober = '\003\207\000\036' /* Filter all dates in October. */,
	ExcelXDFCFilterAllDatesInQuarter1 = '\003\207\000\021' /* Filter all dates in Quarter1. */,
	ExcelXDFCFilterAllDatesInQuarter2 = '\003\207\000\022' /* Filter all dates in Quarter2. */,
	ExcelXDFCFilterAllDatesInQuarter3 = '\003\207\000\023' /* Filter all dates in Quarter3. */,
	ExcelXDFCFilterAllDatesInQuarter4 = '\003\207\000\024' /* Filter all dates in Quarter4. */,
	ExcelXDFCFilterAllDatesInSeptember = '\003\207\000\035' /* Filter all dates in September. */,
	ExcelXDFCFilterBelowAverage = '\003\207\000\"' /* Filter all below-average values. */,
	ExcelXDFCFilterLastMonth = '\003\207\000\010' /* Filter all values related to last month. */,
	ExcelXDFCFilterLastQuarter = '\003\207\000\013' /* Filter all values related to last quarter. */,
	ExcelXDFCFilterLastWeek = '\003\207\000\005' /* Filter all values related to last week. */,
	ExcelXDFCFilterLastYear = '\003\207\000\016' /* Filter all values related to last year. */,
	ExcelXDFCFilterNextMonth = '\003\207\000\011' /* Filter all values related to next month. */,
	ExcelXDFCFilterNextQuarter = '\003\207\000\014' /* Filter all values related to next quarter. */,
	ExcelXDFCFilterNextWeek = '\003\207\000\006' /* Filter all values related to next week. */,
	ExcelXDFCFilterNextYear = '\003\207\000\017' /* Filter all values related to next year. */,
	ExcelXDFCFilterThisMonth = '\003\207\000\007' /* Filter all values related to the current month. */,
	ExcelXDFCFilterThisQuarter = '\003\207\000\012' /* Filter all values related to the current quarter. */,
	ExcelXDFCFilterThisWeek = '\003\207\000\004' /* Filter all values related to the current week. */,
	ExcelXDFCFilterThisYear = '\003\207\000\015' /* Filter all values related to the current year. */,
	ExcelXDFCFilterToday = '\003\207\000\001' /* Filter all values related to the current date. */,
	ExcelXDFCFilterTomorrow = '\003\207\000\003' /* Filter all values related to tomorrow. */,
	ExcelXDFCFilterYearToDate = '\003\207\000\020' /* Filter all values from today until a year ago. */,
	ExcelXDFCFilterYesterday = '\003\207\000\002' /* Filter all values related to yesterday. */
};
typedef enum ExcelXDFC ExcelXDFC;

enum ExcelE304 {
	ExcelE304ThemeFontIndexNone = '\002\300\000\000',
	ExcelE304ThemeFontIndexMajor = '\002\300\000\001',
	ExcelE304ThemeFontIndexMinor = '\002\300\000\002'
};
typedef enum ExcelE304 ExcelE304;

enum ExcelE305 {
	ExcelE305ColorIndexNone = '\002\300\357\322',
	ExcelE305FirstDarkThemeColor = '\002\301\000\001',
	ExcelE305FirstLightThemeColor = '\002\301\000\002',
	ExcelE305SecondDarkThemeColor = '\002\301\000\003',
	ExcelE305SecondLightThemeColor = '\002\301\000\004',
	ExcelE305FirstAccentThemeColor = '\002\301\000\005',
	ExcelE305SecondAccentThemeColor = '\002\301\000\006',
	ExcelE305ThirdAccentThemeColor = '\002\301\000\007',
	ExcelE305FourthAccentThemeColor = '\002\301\000\010',
	ExcelE305FifthAccentThemeColor = '\002\301\000\011',
	ExcelE305SixthAccentThemeColor = '\002\301\000\012',
	ExcelE305HyperlinkThemeColor = '\002\301\000\013',
	ExcelE305FollowedHyperlinkThemeColor = '\002\301\000\014'
};
typedef enum ExcelE305 ExcelE305;

enum ExcelCivt {
	ExcelCivtMinorVersion = '\002\304\000\000',
	ExcelCivtMajorVersion = '\002\304\000\001',
	ExcelCivtOverwriteCurrentVersion = '\002\304\000\002'
};
typedef enum ExcelCivt ExcelCivt;

enum ExcelExWt {
	ExcelExWtEntirePage = '\002\305\000\000',
	ExcelExWtAllTables = '\002\305\000\001',
	ExcelExWtSpecifiedTables = '\002\305\000\002'
};
typedef enum ExcelExWt ExcelExWt;

enum ExcelEbfT {
	ExcelEbfTWebFormattingAll = '\002\306\000\000',
	ExcelEbfTWebFormattingRtf = '\002\306\000\001',
	ExcelEbfTWebFormattingNone = '\002\306\000\002'
};
typedef enum ExcelEbfT ExcelEbfT;

enum ExcelXRbC {
	ExcelXRbCAsRequired = '\002\307\000\000',
	ExcelXRbCAlways = '\002\307\000\001',
	ExcelXRbCNever = '\002\307\000\002'
};
typedef enum ExcelXRbC ExcelXRbC;

enum ExcelE306 {
	ExcelE306ConditionValueNone = '\002\307\377\377',
	ExcelE306ConditionValueNumber = '\002\310\000\000',
	ExcelE306ConditionValueLowestValue = '\002\310\000\001',
	ExcelE306ConditionValueHighestValue = '\002\310\000\002',
	ExcelE306ConditionValuePercent = '\002\310\000\003',
	ExcelE306ConditionValueFormula = '\002\310\000\004',
	ExcelE306ConditionValuePercentile = '\002\310\000\005',
	ExcelE306ConditionValueAutomaticMinimum = '\002\310\000\006',
	ExcelE306ConditionValueAutomaticMaximum = '\002\310\000\007'
};
typedef enum ExcelE306 ExcelE306;

enum ExcelE307 {
	ExcelE307PivotConditionSelectionScope = '\002\311\000\000',
	ExcelE307PivotConditionFieldsScope = '\002\311\000\001',
	ExcelE307PivotConditionDataFieldScope = '\002\311\000\002'
};
typedef enum ExcelE307 ExcelE307;

enum ExcelE308 {
	ExcelE308DatabarFillSolid = '\002\312\000\000',
	ExcelE308DatabarFillGradient = '\002\312\000\001'
};
typedef enum ExcelE308 ExcelE308;

enum ExcelE309 {
	ExcelE309DatabarBorderNone = '\002\313\000\000',
	ExcelE309DatabarBorderSolid = '\002\313\000\001'
};
typedef enum ExcelE309 ExcelE309;

enum ExcelE310 {
	ExcelE310DatabarAxisAutomatic = '\002\314\000\000',
	ExcelE310DatabarAxisMidpoint = '\002\314\000\001',
	ExcelE310DatabarAxisNone = '\002\314\000\002'
};
typedef enum ExcelE310 ExcelE310;

enum ExcelE311 {
	ExcelE311DatabarAutomatic = '\002\315\000\000',
	ExcelE311DatabarPositiveFormat = '\002\315\000\001',
	ExcelE311DatabarCustomFormat = '\002\315\000\002'
};
typedef enum ExcelE311 ExcelE311;

enum ExcelE312 {
	ExcelE312FormatConditionIconNoCellIcon = '\002\315\377\377',
	ExcelE312FormatConditionIconGreenUpArrow = '\002\316\000\001',
	ExcelE312FormatConditionIconYellowSideArrow = '\002\316\000\002',
	ExcelE312FormatConditionIconRedDownArrow = '\002\316\000\003',
	ExcelE312FormatConditionIconGrayUpArrow = '\002\316\000\004',
	ExcelE312FormatConditionIconGraySideArrow = '\002\316\000\005',
	ExcelE312FormatConditionIconGrayDownArrow = '\002\316\000\006',
	ExcelE312FormatConditionIconGreenFlag = '\002\316\000\007',
	ExcelE312FormatConditionIconYellowFlag = '\002\316\000\010',
	ExcelE312FormatConditionIconRedFlag = '\002\316\000\011',
	ExcelE312FormatConditionIconGreenCircle = '\002\316\000\012',
	ExcelE312FormatConditionIconYellowCircle = '\002\316\000\013',
	ExcelE312FormatConditionIconRedCircleWithBorder = '\002\316\000\014',
	ExcelE312FormatConditionIconBlackCircleWithBorder = '\002\316\000\015',
	ExcelE312FormatConditionIconGreenTrafficLight = '\002\316\000\016',
	ExcelE312FormatConditionIconYellowTrafficLight = '\002\316\000\017',
	ExcelE312FormatConditionIconRedTrafficLight = '\002\316\000\020',
	ExcelE312FormatConditionIconYellowTriangle = '\002\316\000\021',
	ExcelE312FormatConditionIconRedDiamond = '\002\316\000\022',
	ExcelE312FormatConditionIconGreenCheckSymbol = '\002\316\000\023',
	ExcelE312FormatConditionIconYellowExclamationSymbol = '\002\316\000\024',
	ExcelE312FormatConditionIconRedCrossSymbol = '\002\316\000\025',
	ExcelE312FormatConditionIconGreenCheck = '\002\316\000\026',
	ExcelE312FormatConditionIconYellowExclamation = '\002\316\000\027',
	ExcelE312FormatConditionIconRedCross = '\002\316\000\030',
	ExcelE312FormatConditionIconYellowUpInclineArrow = '\002\316\000\031',
	ExcelE312FormatConditionIconYellowDownInclineArrow = '\002\316\000\032',
	ExcelE312FormatConditionIconGrayUpInclineArrow = '\002\316\000\033',
	ExcelE312FormatConditionIconGrayDownInclineArrow = '\002\316\000\034',
	ExcelE312FormatConditionIconRedCircle = '\002\316\000\035',
	ExcelE312FormatConditionIconPinkCircle = '\002\316\000\036',
	ExcelE312FormatConditionIconGrayCircle = '\002\316\000\037',
	ExcelE312FormatConditionIconBlackCircle = '\002\316\000 ',
	ExcelE312FormatConditionIconCircleWithOneWhiteQuarter = '\002\316\000!',
	ExcelE312FormatConditionIconCircleWithTwoWhiteQuarters = '\002\316\000\"',
	ExcelE312FormatConditionIconCircleWithThreeWhiteQuarters = '\002\316\000#',
	ExcelE312FormatConditionIconWhiteCircleAllWhiteQuarters = '\002\316\000$',
	ExcelE312FormatConditionIcon0Bars = '\002\316\000%',
	ExcelE312FormatConditionIcon1Bar = '\002\316\000&',
	ExcelE312FormatConditionIcon2Bars = '\002\316\000\'',
	ExcelE312FormatConditionIcon3Bars = '\002\316\000(',
	ExcelE312FormatConditionIcon4Bars = '\002\316\000)',
	ExcelE312FormatConditionIconGoldStar = '\002\316\000*',
	ExcelE312FormatConditionIconHalfGoldStar = '\002\316\000+',
	ExcelE312FormatConditionIconSilverStar = '\002\316\000,',
	ExcelE312FormatConditionIconGreenUpTriangle = '\002\316\000-',
	ExcelE312FormatConditionIconYellowDash = '\002\316\000.',
	ExcelE312FormatConditionIconRedDownTriangle = '\002\316\000/',
	ExcelE312FormatConditionIcon4FilledBoxes = '\002\316\0000',
	ExcelE312FormatConditionIcon3FilledBoxes = '\002\316\0001',
	ExcelE312FormatConditionIcon2FilledBoxes = '\002\316\0002',
	ExcelE312FormatConditionIcon1FilledBox = '\002\316\0003',
	ExcelE312FormatConditionIcon0FilledBoxes = '\002\316\0004'
};
typedef enum ExcelE312 ExcelE312;

enum ExcelE313 {
	ExcelE313IconSetCustom = '\002\316\377\377',
	ExcelE313IconSet3Arrows = '\002\317\000\001',
	ExcelE313IconSet3ArrowsGray = '\002\317\000\002',
	ExcelE313IconSet3Flags = '\002\317\000\003',
	ExcelE313IconSet3TrafficLights1 = '\002\317\000\004',
	ExcelE313IconSet3TrafficLights2 = '\002\317\000\005',
	ExcelE313IconSet3Signs = '\002\317\000\006',
	ExcelE313IconSet3Symbols = '\002\317\000\007',
	ExcelE313IconSet3Symbols2 = '\002\317\000\010',
	ExcelE313IconSet4Arrows = '\002\317\000\011',
	ExcelE313IconSet4ArrowsGray = '\002\317\000\012',
	ExcelE313IconSet4RedToBlack = '\002\317\000\013',
	ExcelE313IconSet4CRV = '\002\317\000\014',
	ExcelE313IconSet4TrafficLights = '\002\317\000\015',
	ExcelE313IconSet5Arrows = '\002\317\000\016',
	ExcelE313IconSet5ArrowsGray = '\002\317\000\017',
	ExcelE313IconSet5CRV = '\002\317\000\020',
	ExcelE313IconSet5Quarters = '\002\317\000\021',
	ExcelE313IconSet3Stars = '\002\317\000\022',
	ExcelE313IconSet3Triangles = '\002\317\000\023',
	ExcelE313IconSet5Boxes = '\002\317\000\024'
};
typedef enum ExcelE313 ExcelE313;

enum ExcelE314 {
	ExcelE314Top10Top = '\002\320\000\001',
	ExcelE314Top10Bottom = '\002\320\000\000'
};
typedef enum ExcelE314 ExcelE314;

enum ExcelE315 {
	ExcelE315CalcForAllValues = '\002\321\000\000',
	ExcelE315CalcForRowGroups = '\002\321\000\001',
	ExcelE315CalcForColGroups = '\002\321\000\002'
};
typedef enum ExcelE315 ExcelE315;

enum ExcelE316 {
	ExcelE316FormatAboveAverage = '\002\322\000\000',
	ExcelE316FormatBelowAverage = '\002\322\000\001',
	ExcelE316FormatEqualAboveAverage = '\002\322\000\002',
	ExcelE316FormatEqualBelowAverage = '\002\322\000\003',
	ExcelE316FormatAboveStandardDeviation = '\002\322\000\004',
	ExcelE316FormatBelowStandardDeviation = '\002\322\000\005'
};
typedef enum ExcelE316 ExcelE316;

enum ExcelE317 {
	ExcelE317FormatUniqueValues = '\002\323\000\000',
	ExcelE317FormatDuplicateValues = '\002\323\000\001'
};
typedef enum ExcelE317 ExcelE317;

enum ExcelE318 {
	ExcelE318TextContains = '\002\324\000\000',
	ExcelE318TextDoesNotContain = '\002\324\000\001',
	ExcelE318TextBeginsWith = '\002\324\000\002',
	ExcelE318TextEndsWith = '\002\324\000\003'
};
typedef enum ExcelE318 ExcelE318;

enum ExcelE319 {
	ExcelE319DateIsToday = '\002\325\000\000',
	ExcelE319DateIsYesterday = '\002\325\000\001',
	ExcelE319DateIsWithinTheLastSevenDays = '\002\325\000\002',
	ExcelE319DateIsThisWeek = '\002\325\000\003',
	ExcelE319DateIsLastWeek = '\002\325\000\004',
	ExcelE319DateIsLastMonth = '\002\325\000\005',
	ExcelE319DateIsTomorrow = '\002\325\000\006',
	ExcelE319DateIsNextWeek = '\002\325\000\007',
	ExcelE319DateIsNextMonth = '\002\325\000\010',
	ExcelE319DateIsThisMonth = '\002\325\000\011'
};
typedef enum ExcelE319 ExcelE319;

enum ExcelE320 {
	ExcelE320DatabarColorTypeColor = '\002\326\000\000',
	ExcelE320DatabarColorTypeSameAsPositive = '\002\326\000\001'
};
typedef enum ExcelE320 ExcelE320;

enum Excel4022 {
	Excel4022Window = 'cwin',
	Excel4022Pane = 'X189'
};
typedef enum Excel4022 Excel4022;

enum Excel4002 {
	Excel4002Window = 'cwin',
	Excel4002Sheet = 'X128',
	Excel4002Workbook = 'X141'
};
typedef enum Excel4002 Excel4002;

enum Excel4003 {
	Excel4003Window = 'cwin',
	Excel4003Sheet = 'X128',
	Excel4003Workbook = 'X141'
};
typedef enum Excel4003 Excel4003;

enum Excel4023 {
	Excel4023Window = 'cwin',
	Excel4023Pane = 'X189'
};
typedef enum Excel4023 Excel4023;

enum Excel4004 {
	Excel4004Application = 'capp',
	Excel4004Sheet = 'X128'
};
typedef enum Excel4004 Excel4004;

enum Excel4010 {
	Excel4010Sheet = 'X128',
	Excel4010Button = 'Xbtn',
	Excel4010Checkbox = 'Xckb',
	Excel4010OptionButton = 'XObn',
	Excel4010Groupbox = 'XGBc',
	Excel4010Label = 'Xlbl',
	Excel4010Textbox = 'XTbx'
};
typedef enum Excel4010 Excel4010;

enum Excel4013 {
	Excel4013Button = 'Xbtn',
	Excel4013Checkbox = 'Xckb',
	Excel4013OptionButton = 'XObn',
	Excel4013Scrollbar = 'XSrl',
	Excel4013Listbox = 'XLbx',
	Excel4013Groupbox = 'XGBc',
	Excel4013Dropdown = 'XdpD',
	Excel4013Spinner = 'XSpn',
	Excel4013Label = 'Xlbl',
	Excel4013Textbox = 'XTbx'
};
typedef enum Excel4013 Excel4013;

enum Excel4014 {
	Excel4014Button = 'Xbtn',
	Excel4014Checkbox = 'Xckb',
	Excel4014OptionButton = 'XObn',
	Excel4014Scrollbar = 'XSrl',
	Excel4014Listbox = 'XLbx',
	Excel4014Groupbox = 'XGBc',
	Excel4014Dropdown = 'XdpD',
	Excel4014Spinner = 'XSpn',
	Excel4014Label = 'Xlbl',
	Excel4014Textbox = 'XTbx'
};
typedef enum Excel4014 Excel4014;

enum Excel4024 {
	Excel4024Dialog = 'X165',
	Excel4024Scenario = 'X191'
};
typedef enum Excel4024 Excel4024;

enum Excel4005 {
	Excel4005Sheet = 'X128',
	Excel4005Workbook = 'X141'
};
typedef enum Excel4005 Excel4005;

enum Excel4011 {
	Excel4011Button = 'Xbtn',
	Excel4011Checkbox = 'Xckb',
	Excel4011OptionButton = 'XObn',
	Excel4011Scrollbar = 'XSrl',
	Excel4011Listbox = 'XLbx',
	Excel4011Groupbox = 'XGBc',
	Excel4011Dropdown = 'XdpD',
	Excel4011Spinner = 'XSpn',
	Excel4011Label = 'Xlbl',
	Excel4011Textbox = 'XTbx'
};
typedef enum Excel4011 Excel4011;

enum Excel4015 {
	Excel4015Button = 'Xbtn',
	Excel4015Checkbox = 'Xckb',
	Excel4015OptionButton = 'XObn',
	Excel4015Scrollbar = 'XSrl',
	Excel4015Listbox = 'XLbx',
	Excel4015Groupbox = 'XGBc',
	Excel4015Dropdown = 'XdpD',
	Excel4015Spinner = 'XSpn',
	Excel4015Label = 'Xlbl',
	Excel4015Textbox = 'XTbx'
};
typedef enum Excel4015 Excel4015;

enum Excel4027 {
	Excel4027FormatCondition = 'X227',
	Excel4027ColorScaleFormatCondition = 'X325',
	Excel4027DatabarFormatCondition = 'X312',
	Excel4027IconSetFormatCondition = 'X315',
	Excel4027Top10FormatCondition = 'X321',
	Excel4027AboveAverageFormatCondition = 'X322',
	Excel4027UniqueValuesFormatCondition = 'X323'
};
typedef enum Excel4027 Excel4027;

enum Excel4028 {
	Excel4028FormatCondition = 'X227',
	Excel4028ColorScaleFormatCondition = 'X325',
	Excel4028DatabarFormatCondition = 'X312',
	Excel4028IconSetFormatCondition = 'X315',
	Excel4028Top10FormatCondition = 'X321',
	Excel4028AboveAverageFormatCondition = 'X322',
	Excel4028UniqueValuesFormatCondition = 'X323'
};
typedef enum Excel4028 Excel4028;

enum Excel4029 {
	Excel4029FormatCondition = 'X227',
	Excel4029ColorScaleFormatCondition = 'X325',
	Excel4029DatabarFormatCondition = 'X312',
	Excel4029IconSetFormatCondition = 'X315',
	Excel4029Top10FormatCondition = 'X321',
	Excel4029AboveAverageFormatCondition = 'X322',
	Excel4029UniqueValuesFormatCondition = 'X323'
};
typedef enum Excel4029 Excel4029;

enum Excel4008 {
	Excel4008PivotTable = 'X155',
	Excel4008PivotField = 'X157'
};
typedef enum Excel4008 Excel4008;

enum Excel4006 {
	Excel4006CubeField = 'X900',
	Excel4006CalculatedMember = 'X901',
	Excel4006PivotFilter = 'X903',
	Excel4006ValueChange = 'X905'
};
typedef enum Excel4006 Excel4006;

enum Excel4009 {
	Excel4009PivotField = 'X157',
	Excel4009PivotItem = 'X160'
};
typedef enum Excel4009 Excel4009;

enum Excel4007 {
	Excel4007CubeField = 'X900',
	Excel4007PivotField = 'X157'
};
typedef enum Excel4007 Excel4007;

enum Excel4025 {
	Excel4025PivotCache = 'X151',
	Excel4025QueryTable = 'X231'
};
typedef enum Excel4025 Excel4025;

enum Excel4016 {
	Excel4016Listbox = 'XLbx',
	Excel4016Dropdown = 'XdpD'
};
typedef enum Excel4016 Excel4016;

enum Excel4026 {
	Excel4026FormatCondition = 'X227',
	Excel4026DisplayFormat = 'X306',
	Excel4026Top10FormatCondition = 'X321',
	Excel4026AboveAverageFormatCondition = 'X322',
	Excel4026UniqueValuesFormatCondition = 'X323'
};
typedef enum Excel4026 Excel4026;

enum Excel4021 {
	Excel4021Listbox = 'XLbx',
	Excel4021Dropdown = 'XdpD'
};
typedef enum Excel4021 Excel4021;

enum Excel4018 {
	Excel4018Listbox = 'XLbx',
	Excel4018Dropdown = 'XdpD'
};
typedef enum Excel4018 Excel4018;

enum Excel4012 {
	Excel4012Button = 'Xbtn',
	Excel4012Checkbox = 'Xckb',
	Excel4012OptionButton = 'XObn',
	Excel4012Scrollbar = 'XSrl',
	Excel4012Listbox = 'XLbx',
	Excel4012Groupbox = 'XGBc',
	Excel4012Dropdown = 'XdpD',
	Excel4012Spinner = 'XSpn',
	Excel4012Label = 'Xlbl',
	Excel4012Textbox = 'XTbx'
};
typedef enum Excel4012 Excel4012;

enum Excel4017 {
	Excel4017Listbox = 'XLbx',
	Excel4017Dropdown = 'XdpD'
};
typedef enum Excel4017 Excel4017;

enum Excel4019 {
	Excel4019Listbox = 'XLbx',
	Excel4019Dropdown = 'XdpD'
};
typedef enum Excel4019 Excel4019;

enum Excel4020 {
	Excel4020Listbox = 'XLbx',
	Excel4020Dropdown = 'XdpD'
};
typedef enum Excel4020 Excel4020;

enum Excel4001 {
	Excel4001Window = 'cwin',
	Excel4001Sheet = 'X128',
	Excel4001Workbook = 'X141',
	Excel4001Pane = 'X189'
};
typedef enum Excel4001 Excel4001;

enum ExcelVrOf {
	ExcelVrOfOverflow = '\002\302\000\000',
	ExcelVrOfClip = '\002\302\000\001',
	ExcelVrOfEllipsis = '\002\302\000\002'
};
typedef enum ExcelVrOf ExcelVrOf;

enum ExcelHzOf {
	ExcelHzOfOverflow = '\002\303\000\000',
	ExcelHzOfClip = '\002\303\000\001'
};
typedef enum ExcelHzOf ExcelHzOf;

enum Excel4036 {
	Excel4036CalloutFormat = 'X101',
	Excel4036Callout = 'cD00'
};
typedef enum Excel4036 Excel4036;

enum Excel4037 {
	Excel4037CalloutFormat = 'X101',
	Excel4037Callout = 'cD00'
};
typedef enum Excel4037 Excel4037;

enum Excel4038 {
	Excel4038CalloutFormat = 'X101',
	Excel4038Callout = 'cD00'
};
typedef enum Excel4038 Excel4038;

enum Excel4035 {
	Excel4035Rectangle = 'XRct',
	Excel4035Oval = 'XOvl',
	Excel4035Arc = 'Xarc'
};
typedef enum Excel4035 Excel4035;

enum Excel4032 {
	Excel4032Line = 'Xlne',
	Excel4032Rectangle = 'XRct',
	Excel4032Oval = 'XOvl',
	Excel4032Arc = 'Xarc',
	Excel4032Shape = 'pShp'
};
typedef enum Excel4032 Excel4032;

enum Excel4033 {
	Excel4033Line = 'Xlne',
	Excel4033Rectangle = 'XRct',
	Excel4033Oval = 'XOvl',
	Excel4033Arc = 'Xarc',
	Excel4033Shape = 'pShp'
};
typedef enum Excel4033 Excel4033;

enum Excel4030 {
	Excel4030Line = 'Xlne',
	Excel4030Rectangle = 'XRct',
	Excel4030Oval = 'XOvl',
	Excel4030Arc = 'Xarc'
};
typedef enum Excel4030 Excel4030;

enum Excel4034 {
	Excel4034Line = 'Xlne',
	Excel4034Rectangle = 'XRct',
	Excel4034Oval = 'XOvl',
	Excel4034Arc = 'Xarc'
};
typedef enum Excel4034 Excel4034;

enum Excel4031 {
	Excel4031Line = 'Xlne',
	Excel4031Rectangle = 'XRct',
	Excel4031Oval = 'XOvl',
	Excel4031Arc = 'Xarc',
	Excel4031Shape = 'pShp'
};
typedef enum Excel4031 Excel4031;

enum Excel4041 {
	Excel4041ChartFillFormat = 'X253',
	Excel4041ChartTitle = 'X256',
	Excel4041AxisTitle = 'X257',
	Excel4041SeriesPoint = 'X262',
	Excel4041Series = 'X263',
	Excel4041DataLabel = 'X265',
	Excel4041LegendKey = 'X269',
	Excel4041DownBars = 'X279',
	Excel4041Floor = 'X280',
	Excel4041Walls = 'X281',
	Excel4041PlotArea = 'X283',
	Excel4041ChartArea = 'X284',
	Excel4041Legend = 'X285',
	Excel4041DisplayUnitLabel = 'X299'
};
typedef enum Excel4041 Excel4041;

enum Excel4053 {
	Excel4053ChartArea = 'X284',
	Excel4053Legend = 'X285'
};
typedef enum Excel4053 Excel4053;

enum Excel4051 {
	Excel4051SeriesPoint = 'X262',
	Excel4051Series = 'X263',
	Excel4051LegendKey = 'X269',
	Excel4051Trendline = 'X271',
	Excel4051Floor = 'X280',
	Excel4051Walls = 'X281',
	Excel4051PlotArea = 'X283',
	Excel4051ChartArea = 'X284',
	Excel4051ErrorBars = 'X286'
};
typedef enum Excel4051 Excel4051;

enum Excel4049 {
	Excel4049Chart = 'X119',
	Excel4049SeriesPoint = 'X262',
	Excel4049Series = 'X263'
};
typedef enum Excel4049 Excel4049;

enum Excel4052 {
	Excel4052Floor = 'X280',
	Excel4052Walls = 'X281'
};
typedef enum Excel4052 Excel4052;

enum Excel4042 {
	Excel4042ChartFillFormat = 'X253',
	Excel4042ChartTitle = 'X256',
	Excel4042AxisTitle = 'X257',
	Excel4042SeriesPoint = 'X262',
	Excel4042Series = 'X263',
	Excel4042DataLabel = 'X265',
	Excel4042LegendKey = 'X269',
	Excel4042DownBars = 'X279',
	Excel4042Floor = 'X280',
	Excel4042Walls = 'X281',
	Excel4042PlotArea = 'X283',
	Excel4042ChartArea = 'X284',
	Excel4042Legend = 'X285',
	Excel4042DisplayUnitLabel = 'X299'
};
typedef enum Excel4042 Excel4042;

enum Excel4045 {
	Excel4045ChartFillFormat = 'X253',
	Excel4045ChartTitle = 'X256',
	Excel4045AxisTitle = 'X257',
	Excel4045SeriesPoint = 'X262',
	Excel4045Series = 'X263',
	Excel4045DataLabel = 'X265',
	Excel4045LegendKey = 'X269',
	Excel4045DownBars = 'X279',
	Excel4045Floor = 'X280',
	Excel4045Walls = 'X281',
	Excel4045PlotArea = 'X283',
	Excel4045ChartArea = 'X284',
	Excel4045Legend = 'X285',
	Excel4045DisplayUnitLabel = 'X299'
};
typedef enum Excel4045 Excel4045;

enum Excel4046 {
	Excel4046ChartFillFormat = 'X253',
	Excel4046ChartTitle = 'X256',
	Excel4046AxisTitle = 'X257',
	Excel4046SeriesPoint = 'X262',
	Excel4046Series = 'X263',
	Excel4046DataLabel = 'X265',
	Excel4046LegendKey = 'X269',
	Excel4046DownBars = 'X279',
	Excel4046Floor = 'X280',
	Excel4046Walls = 'X281',
	Excel4046PlotArea = 'X283',
	Excel4046ChartArea = 'X284',
	Excel4046Legend = 'X285',
	Excel4046DisplayUnitLabel = 'X299'
};
typedef enum Excel4046 Excel4046;

enum Excel4050 {
	Excel4050ChartObject = 'X221',
	Excel4050SeriesPoint = 'X262',
	Excel4050Series = 'X263',
	Excel4050ChartArea = 'X284'
};
typedef enum Excel4050 Excel4050;

enum Excel4044 {
	Excel4044ChartFillFormat = 'X253',
	Excel4044ChartTitle = 'X256',
	Excel4044AxisTitle = 'X257',
	Excel4044SeriesPoint = 'X262',
	Excel4044Series = 'X263',
	Excel4044DataLabel = 'X265',
	Excel4044LegendKey = 'X269',
	Excel4044DownBars = 'X279',
	Excel4044Floor = 'X280',
	Excel4044Walls = 'X281',
	Excel4044PlotArea = 'X283',
	Excel4044ChartArea = 'X284',
	Excel4044Legend = 'X285',
	Excel4044DisplayUnitLabel = 'X299'
};
typedef enum Excel4044 Excel4044;

enum Excel4040 {
	Excel4040ChartObject = 'X221',
	Excel4040Axis = 'X255'
};
typedef enum Excel4040 Excel4040;

enum Excel4048 {
	Excel4048ChartFillFormat = 'X253',
	Excel4048ChartTitle = 'X256',
	Excel4048AxisTitle = 'X257',
	Excel4048SeriesPoint = 'X262',
	Excel4048Series = 'X263',
	Excel4048DataLabel = 'X265',
	Excel4048LegendKey = 'X269',
	Excel4048DownBars = 'X279',
	Excel4048Floor = 'X280',
	Excel4048Walls = 'X281',
	Excel4048PlotArea = 'X283',
	Excel4048ChartArea = 'X284',
	Excel4048Legend = 'X285',
	Excel4048DisplayUnitLabel = 'X299'
};
typedef enum Excel4048 Excel4048;

enum Excel4047 {
	Excel4047ChartFillFormat = 'X253',
	Excel4047ChartTitle = 'X256',
	Excel4047AxisTitle = 'X257',
	Excel4047SeriesPoint = 'X262',
	Excel4047Series = 'X263',
	Excel4047DataLabel = 'X265',
	Excel4047LegendKey = 'X269',
	Excel4047DownBars = 'X279',
	Excel4047Floor = 'X280',
	Excel4047Walls = 'X281',
	Excel4047PlotArea = 'X283',
	Excel4047ChartArea = 'X284',
	Excel4047Legend = 'X285',
	Excel4047DisplayUnitLabel = 'X299'
};
typedef enum Excel4047 Excel4047;

enum Excel4043 {
	Excel4043ChartFillFormat = 'X253',
	Excel4043ChartTitle = 'X256',
	Excel4043AxisTitle = 'X257',
	Excel4043SeriesPoint = 'X262',
	Excel4043Series = 'X263',
	Excel4043DataLabel = 'X265',
	Excel4043LegendKey = 'X269',
	Excel4043DownBars = 'X279',
	Excel4043Floor = 'X280',
	Excel4043Walls = 'X281',
	Excel4043PlotArea = 'X283',
	Excel4043ChartArea = 'X284',
	Excel4043Legend = 'X285',
	Excel4043DisplayUnitLabel = 'X299'
};
typedef enum Excel4043 Excel4043;

enum Excel4039 {
	Excel4039Chart = 'X119',
	Excel4039ChartObject = 'X221'
};
typedef enum Excel4039 Excel4039;

@protocol ExcelGenericMethods

- (void) closeSaving:(ExcelSavo)saving savingIn:(ExcelXLfd)savingIn;  // Close an object
- (NSInteger) dataSizeAs:(NSNumber *)as;  // Return the size in bytes of an object
- (void) delete;  // Delete an element from an object
- (SBObject *) duplicateTo:(SBObject *)to;  // Duplicate object(s)
- (BOOL) exists;  // Verify if an object exists
- (SBObject *) moveTo:(SBObject *)to;  // Move object(s) to a new location
- (void) open;  // Open the specified object(s)
- (void) printWithProperties:(ExcelPrintSettings *)withProperties printDialog:(BOOL)printDialog;  // Print the specified object(s)
- (void) saveIn:(ExcelXLfd)in_ as:(ExcelE188)as;  // Save an object
- (void) select;  // Make a selection
- (void) applyFilter;  // Apply the defined filter
- (void) applySort;  // Sorts the range based on the currently applied sort states.
- (BOOL) canCheckOutFileName:(NSString *)fileName;  // Returns True if Excel can check out a specified workbook from a server.
- (void) checkOutFileName:(NSString *)fileName;  // Copies a specified workbook from a server to a local computer for editing. Returns a String that represents the local path and file name of the workbook checked out.
- (void) clearColorstops;  // Clears the represented object.
- (void) clearSortfields;  // Clears all the sortfields object.
- (void) deleteSortfield;  // Removes the specified sortfield object from the SortFields collection.
- (void) openDataBaseFilename:(NSString *)filename commandText:(NSString *)commandText rcommandType:(NSInteger)rcommandType backGroundQuery:(id)backGroundQuery importDataAs:(NSInteger)importDataAs;  // Open a data base
- (void) openXmlFilename:(NSString *)filename styleSheets:(NSInteger)styleSheets loadOption:(NSInteger)loadOption;  // Open an XML file
- (void) showAll;  // Show all the hidden data, but do not exist the filter mode
- (void) removeDuplicates;  // Removes duplicate values from a range of values.

@end



/*
 * Standard Suite
 */

// A scriptable object
@interface ExcelBaseObject : SBObject <ExcelGenericMethods>

@property (copy) NSDictionary *properties;  // All of the object's properties


@end

// A basic application program
@interface ExcelBaseApplication : ExcelBaseObject

@property (readonly) BOOL frontmost;  // Is this the frontmost application?
@property (copy, readonly) NSString *name;  // the name
@property (copy, readonly) NSString *version;  // the version of the application


@end

// A basic document
@interface ExcelBaseDocument : ExcelBaseObject

@property (readonly) BOOL modified;  // Has the document been modified since the last save?
@property (copy, readonly) NSString *name;  // the name


@end

// A basic window
@interface ExcelBasicWindow : ExcelBaseObject

@property (copy) ExcelRectangle *bounds;  // the boundary rectangle for the window
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

@interface ExcelPrintSettings : SBObject <ExcelGenericMethods>

@property NSInteger copies;  // the number of copies of a document to be printed 
@property BOOL collating;  // Should printed copies be collated?
@property NSInteger startingPage;  // the first page of the document to be printed
@property NSInteger endingPage;  // the last page of the document to be printed
@property NSInteger pagesAcross;  // number of logical pages laid across a physical page
@property NSInteger pagesDown;  // number of logical pages laid out down a physical page
@property (copy) NSDate *requestedPrintTime;  // the time at which the desktop printer should print the document...
@property ExcelEnum errorHandling;  // how errors are handled
@property (copy) NSString *faxNumber;  // for fax number
@property (copy) NSString *targetPrinter;  // the queue name of the target printer


@end



/*
 * Microsoft Office Suite
 */

// A control within a command bar.
@interface ExcelCommandBarControl : ExcelBaseObject

@property BOOL beginGroup;  // Returns or sets if the command bar control appears at the beginning of a group of controls on the command bar.
@property (readonly) BOOL builtIn;  // Returns true if the command bar control is a built-in command bar control.
@property (readonly) ExcelMCLT controlType;  // Returns the type of the command bar control.
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
@interface ExcelCommandBarButton : ExcelCommandBarControl

@property (readonly) BOOL buttonFaceIsDefault;  // Returns if the face of a command bar button control is the original built-in face.
@property ExcelMBns buttonState;  // Returns or set the appearance of a command bar button control.  The property is read-only for built-in command bar buttons.
@property ExcelMBTs buttonStyle;  // Returns or sets the way a command button control is displayed.
@property NSInteger faceId;  // Returns or sets the Id number for the face of the command bar button control.


@end

// A combobox menu control within a command bar.
@interface ExcelCommandBarCombobox : ExcelCommandBarControl

@property ExcelMXcb comboboxStyle;  // Returns or sets the way a command bar combobox control is displayed.
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
@interface ExcelCommandBarPopup : ExcelCommandBarControl

- (SBElementArray<ExcelCommandBarControl *> *) commandBarControls;


@end

// Toolbars used in all of the Office applications.
@interface ExcelCommandBar : ExcelBaseObject

- (SBElementArray<ExcelCommandBarControl *> *) commandBarControls;

@property ExcelMBPS barPosition;  // Returns or sets the position of the command bar.
@property (readonly) ExcelMBTY barType;  // Returns the type of this command bar.
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

@interface ExcelDocumentProperty : ExcelBaseObject

@property (copy) NSNumber *documentPropertyType;  // Returns or sets the document property type.
@property (copy) NSString *linkSource;  // Returns or sets the source of a lined custom document property.
@property BOOL linkToContent;  // True if the value of the document property is lined to the content of the container document.  False if the value is static.  This only applies to custom document properties.  For built-in properties this is always false.
@property (copy) NSString *name;  // Returns or sets the name of the document property.
@property (copy) NSString *value;  // Returns or sets the value of the document property.


@end

@interface ExcelCustomDocumentProperty : ExcelDocumentProperty


@end

@interface ExcelWebPageFont : ExcelBaseObject

@property (copy) NSString *fixedWidthFont;  // Returns or sets the fixed-width font setting.
@property double fixedWidthFontSize;  // Returns or sets the fixed-width font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.
@property (copy) NSString *proportionalFont;  // Returns or sets the proportional font setting.
@property double proportionalFontSize;  // Returns or sets the proportional font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.


@end



/*
 * Microsoft Excel Suite
 */

// Represents a cell comment.
@interface ExcelExcelComment : ExcelBaseObject

@property (copy, readonly) NSString *author;  // Returns the author of the comment.
@property (copy, readonly) ExcelShape *shapeObject;  // Returns a shape object that represents the shape attached to the specified comment.
@property BOOL visible;  // Returns or sets if the specified object is visible.

- (NSString *) ExcelCommentTextText:(NSString *)text start:(NSInteger)start overWrite:(BOOL)overWrite;  // Returns or sets the text of the comment
- (ExcelExcelComment *) nextExcelComment;  // Returns the next Excel comment object
- (ExcelExcelComment *) previousExcelComment;  // Returns the previous Excel comment object

@end

@interface ExcelODBCError : ExcelBaseObject

@property (copy, readonly) NSString *errorString;  // Returns the ODBC error string.
@property (copy, readonly) NSString *sqlState;  // Returns the SQL state error.


@end

// Represents the various types of protection options available for a worksheet
@interface ExcelProtection : ExcelBaseObject

@property (readonly) BOOL allowDeletingColumns;  // Returns True if the deleting of columns is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowDeletingRows;  // Returns True if the deleting of rows is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowFiltering;  // Returns True if the filtering is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowFormattingCells;  // Returns True if the formatting of cells is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowFormattingColumns;  // Returns True if the formatting of columns is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowFormattingRows;  // Returns True if the formatting of rows is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowInsertingColumns;  // Returns True if the inserting of columns is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowInsertingHyperlinks;  // Returns True if the inserting of hyperlinks is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowInsertingRows;  // Returns True if the inserting of rows is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowSorting;  // Returns True if the sorting is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowUsingPivotTable;  // Returns True if the using of pivot table is allowed on a protected worksheet. Read-only Boolean.


@end

@interface ExcelAboveAverageFormatCondition : ExcelBaseObject

@property ExcelE316 aboveOrBelow;
@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property ExcelE315 calcFor;
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) ExcelE195 formatConditionType;  // Return the conditional format type.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property NSInteger numberOfStandardDeviations;
@property ExcelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

// Represents a single add-in, either installed or not installed.
@interface ExcelAddIn : ExcelBaseObject

@property (copy, readonly) NSString *fullName;  // Returns the name of the add in, including its path on disk, as a string.
@property BOOL installed;  // Returns or sets if the add in is installed.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy, readonly) NSString *path;  // Returns the complete path of the object, excluding the final separator and name of the add in.


@end

// A representation of the Microsoft Excel application.
@interface ExcelApplication : SBApplication

- (SBElementArray<ExcelAddIn *> *) addIns;
- (SBElementArray<ExcelChartSheet *> *) chartSheets;
- (SBElementArray<ExcelCommandBar *> *) commandBars;
- (SBElementArray<ExcelNamedItem *> *) namedItems;
- (SBElementArray<ExcelRange *> *) ranges;
- (SBElementArray<ExcelCell *> *) cells;
- (SBElementArray<ExcelRow *> *) rows;
- (SBElementArray<ExcelColumn *> *) columns;
- (SBElementArray<ExcelWindow *> *) windows;
- (SBElementArray<ExcelWorkbook *> *) workbooks;
- (SBElementArray<ExcelSheet *> *) sheets;
- (SBElementArray<ExcelWorksheet *> *) worksheets;
- (SBElementArray<ExcelInternationalMacroSheet *> *) internationalMacroSheets;
- (SBElementArray<ExcelMacroSheet *> *) macroSheets;
- (SBElementArray<ExcelRecentFile *> *) recentFiles;
- (SBElementArray<ExcelODBCError *> *) ODBCErrors;

@property BOOL AutoFormatAsYouTypeReplaceHyperlinks;  // True if Microsoft Excel automatically formats hyperlinks as you type. False if Excel does not automatically format hyperlinks as you type.
@property ExcelE184 ExcelCursor;  // Returns or sets the appearance of the mouse pointer in Microsoft Excel.
@property NSInteger ODBCTimeout;  // Returns or sets the ODBC query time limit, in seconds. The default value is 45 seconds. 
@property (copy, readonly) ExcelCell *activeCell;  // Returns a range object that represents the active cell in the active window, the window on top, or in the specified window. If the window isn't displaying a worksheet, this property fails.
@property (copy, readonly) ExcelChart *activeChart;  // Returns a chart object that represents the active chart, either an embedded chart or a chart sheet. An embedded chart is considered active when it's either selected or activated.
@property (copy) NSString *activePrinter;  // Returns or sets the name of the active printer.
@property (copy, readonly) ExcelSheet *activeSheet;  // Returns an object that represents the active sheet, the sheet on top, in the active workbook or in the specified window or workbook.
@property (copy, readonly) ExcelWindow *activeWindow;  // Returns a window object that represents the active window, the window on top.
@property (copy, readonly) ExcelWorkbook *activeWorkbook;  // Returns a workbook object that represents the workbook in the active window, the window on top.
@property BOOL alertBeforeOverwriting;  // Returns or sets if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation.
@property (copy) NSString *altStartupPath;  // Returns or sets the name of the alternate startup folder.
@property (readonly) BOOL arbitraryXMLSupportAvailable;  // Returns a Boolean value that indicates whether the XML features in Microsoft Excel are available
@property BOOL askToUpdateLinks;  // Returns or sets if Microsoft Excel asks the user to update links when opening files with links.
@property (copy, readonly) ExcelAutocorrect *autocorrectObject;  // Returns an autocorrect object that represents the Microsoft Excel AutoCorrect attributes.
@property ExcelMASc automationSecurity;
@property (readonly) NSInteger build;  // Returns the Microsoft Excel build number.
@property BOOL calculateBeforeSave;  // Returns or sets if workbooks are calculated before they're saved to disk.
@property ExcelE174 calculation;  // Returns or sets the calculation mode.
@property (readonly) NSInteger calculationVersion;  // Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits, on the left, are the major version of Microsoft Excel.
@property (copy, readonly) NSString *caption;  // Returns the name of the application.
@property BOOL cellDragAndDrop;  // Returns or sets if dragging and dropping cells is enabled.
@property ExcelE172 commandUnderlines;  // Returns or sets the state of the command underlines in Microsoft Excel.
@property BOOL copyObjectsWithCells;  // Returns or sets if objects are cut, copied, extracted, and sorted with cells.
@property (readonly) NSInteger customListCount;  // Returns the number of defined custom lists, including built-in lists.
@property ExcelE161 cutCopyMode;  // Returns or sets the status of cut or copy mode.
@property ExcelE262 dataEntryMode;  // Returns or sets data entry mode, as shown in the following table. When in data entry mode, you can enter data only in the unlocked cells in the currently selected range. 
@property (copy) NSString *defaultFilePath;  // Returns or sets the default path that Microsoft Excel uses when it opens files. 
@property ExcelE188 defaultSaveFormat;  // Returns or sets the default format for saving files.
@property (copy, readonly) ExcelDefaultWebOptions *defaultWebOptionsObject;  // Returns the default web object.
@property BOOL displayAlerts;  // Returns or sets if Microsoft Excel displays certain alerts and messages while handling events from AppleScript.
@property ExcelE194 displayCommentIndicator;  // Returns or sets the way cells display comments and indicators.
@property BOOL displayExcel4Menus;  // Returns or sets if Microsoft Excel displays version 4.0 menu bars.
@property BOOL displayFormulaBar;  // Returns or sets  if the formula bar is displayed.
@property BOOL displayFullScreen;  // Returns or sets if Microsoft Excel is in full-screen mode.
@property BOOL displayFunctionTooltips;  // Returns or sets if function tool tips can be displayed.
@property BOOL displayInsertOptions;  // Returns or sets if the insert options button should be displayed. 
@property BOOL displayNoteIndicator;  // Returns or sets if cells containing notes display cell tips and contain note indicators, small dots in their upper-right corners.
@property BOOL displayRecentFiles;  // Returns or sets if the list of recently used files is displayed on the file menu.
@property BOOL displayScrollBars;  // Returns or sets if scroll bars are visible for all workbooks.
@property BOOL displayStatusBar;  // Returns or sets if the status bar is displayed.
@property BOOL editDirectlyInCell;  // Returns or sets if Microsoft Excel allows editing in cells.
@property BOOL enableAnimations;  // Returns or sets if animated insertion and deletion is enabled.
@property BOOL enableAutocomplete;  // Returns or sets if the autocomplete feature is enabled.
@property ExcelE167 enableCancelKey;  // Controls how Microsoft Excel handles CTRL-BREAK, ESC, or cmd-period user interruptions to the running procedure.
@property BOOL enableEvents;  // Returns or sets if events are enabled for the application object.
@property BOOL enableFormulaAutocomplete;  // Returns or sets if the formula autocomplete feature is enabled.
@property BOOL enableFormulaTypeAhead;  // Returns or sets if the formula autocomplete type ahead is enabled.
@property BOOL enableSound;  // Returns or sets if sound is enabled for Microsoft Office.
@property BOOL extendList;  // Returns or sets if Microsoft Excel automatically extends formatting and formulas to new data that is added to a list.
@property BOOL fixedDecimal;  // All data entered after this property is set to true will be formatted with the number of fixed decimal places set by the fixed decimal places property.
@property NSInteger fixedDecimalPlaces;  // Returns or sets the number of fixed decimal places used when the fixed decimal property is set to true.
@property NSInteger formulaAutocompleteWait;  // Returns or sets number of characters to wait before formula type ahead.
@property (copy, readonly) id frontmost;  // Returns if the application is the frontmost window.
@property BOOL includeEmptyCellsInLists;  // Returns or sets if empty cells are included in range lists.
@property BOOL iteration;  // Returns or sets  if Microsoft Excel will use iteration to resolve circular references.
@property BOOL keepFourDigitYears;  // Returns or sets if years values should be kept as four digits instead of two.
@property NSInteger keyboardScript;
@property (copy, readonly) NSString *libraryPath;  // Returns the path to the Library folder.
@property (readonly) NSInteger localizedLanguage;  // Returns the Windows language ID for the locale that Microsoft Excel is using.
@property (readonly) BOOL mathCoprocessorAvailable;  // Returns true if a math coprocessor is available.
@property double maxChange;  // Returns or sets the maximum amount of change between each iteration as Microsoft Excel resolves circular references. 
@property NSInteger maxIterations;  // Returns or sets the maximum number of iterations that Microsoft Excel can use to resolve a circular reference.
@property ExcelE260 measurementUnit;  // Returns or set the current unit of measure.
@property (readonly) NSInteger memoryFree;  // Returns the amount of memory that's still available for Microsoft Excel to use, in bytes.
@property (readonly) NSInteger memoryTotal;  // Returns the total amount of memory, in bytes, that's available to Microsoft Excel, including memory already in use.
@property (readonly) NSInteger memoryUsed;  // Returns the amount of memory that Microsoft Excel is currently using, in bytes.
@property BOOL moveAfterReturn;  // Returns or sets if the active cell will be moved as soon as the ENTER/RETURN key is pressed.
@property ExcelE149 moveAfterReturnDirection;  // Returns or sets the direction in which the active cell is moved when the user presses ENTER.
@property (copy, readonly) NSString *name;  // Returns the name of the application.
@property (copy, readonly) NSString *networkTemplatesPath;  // Returns the network path where templates are stored. If the network path doesn't exist, this property returns an empty string. 
@property (copy, readonly) NSString *operatingSystem;  // Returns the name and version number of the current operating system.
@property (copy, readonly) NSString *organizationName;  // Returns the registered organization name.
@property (copy, readonly) NSString *path;  // Returns the complete path of the application, excluding the final separator and name of the application.
@property (copy, readonly) NSString *pathSeparator;  // Returns the path separator character.
@property BOOL pivotTableSelection;  // Returns or sets if pivot tables use structured selection.
@property BOOL promptForSummaryInfo;  // Returns or sets if Microsoft Excel asks for summary information when files are first saved.
@property ExcelE158 referenceStyle;  // Returns or sets how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style.
@property BOOL ribbonExpanded;  // Returns or sets a Boolean value that indicates whether the Ribbon in Microsoft Excel is expanded
@property BOOL rollZoom;  // Returns or sets if the IntelliMouse zooms instead of scrolling.
@property NSInteger saveInterval;
@property BOOL screenUpdating;  // Returns or sets if screen updating is turned on. Turn screen updating off to speed up your AppleScript code. You won't be able to see what the code is doing, but it will run faster.
@property (copy, readonly) ExcelRange *selection;  // Returns the selected object in the active window.
@property NSInteger sheetsInNewWorkbook;  // Returns or sets the number of sheets that Microsoft Excel automatically inserts into new workbooks.
@property BOOL showChartTipNames;  // Returns or sets if charts show chart tip names.
@property BOOL showChartTipValues;  // Returns or sets if charts show chart tip values.
@property BOOL showRibbon;  // Returns or sets a Boolean value that indicates whether the Ribbon in Microsoft Excel is shown
@property BOOL showToolTips;  // Returns or sets if tool tips are turned on.
@property (copy, readonly) ExcelXlspellingOptions *spellingOptions;  // Returns the default spelling options object.
@property (copy) NSString *standardFont;  // Returns or sets the name of the standard font.
@property double standardFontSize;  // Returns or sets the standard font size, in points.
@property BOOL startupDialog;  // Returns or sets if the startup dialog should be shown when starting up the application.
@property (copy, readonly) NSString *startupPath;  // Returns the complete path of the startup folder, excluding the final separator.
@property (copy) NSString *statusBar;  // Returns or sets the text in the status bar.
@property (copy, readonly) NSString *templatesPath;  // Returns the local path where templates are stored.
@property (copy, readonly) id thisCell;  // Returns the cell in which the user-defined function is being called from as a Range object.
@property (copy) NSString *transitionMenuKey;  // Returns or sets the alternate menu or help key.
@property ExcelE264 transitionMenuKeyAction;  // Returns or sets the action taken when the alternate menu key is pressed.
@property NSInteger twoDigitCutoffYear;  // Returns or sets the two digit value  after which year values are shown as four digits.
@property (readonly) double usableHeight;  // Returns the maximum height of the space that a window can occupy in points.
@property (readonly) double usableWidth;  // Returns the maximum width of the space that a window can occupy in points.
@property (copy) NSString *userName;  // Returns or sets the name of the current user.

- (void) quitSaving:(ExcelSavo)saving;  // Quit an application program
- (void) select;  // Make a selection
- (void) reset:(Excel4000)x;  // Resets a built-in command bar or command bar control to its default configuration.
- (void) ExcelRepeat;  // Repeats the last user-interface action.
- (void) activateObject:(Excel4001)x;  // Make the object active
- (void) addChartAutoformatChart:(ExcelChart *)chart name:(NSString *)name objectDescription:(NSString *)objectDescription;  // Adds a custom chart autoformat to the list of available chart autoformats.
- (void) addCustomListListArray:(ExcelE275)listArray byRow:(BOOL)byRow;  // Adds a custom list for custom autofill and/or custom sort.
- (void) addItemToList:(Excel4016)x itemText:(NSString *)itemText entry_index:(NSInteger)entry_index;  // Adds a string to the list control
- (void) arrange_windowsArrangeStyle:(ExcelE183)arrangeStyle activeWorkbook:(BOOL)activeWorkbook syncHorizontal:(BOOL)syncHorizontal syncVertical:(BOOL)syncVertical;  // Arranges the windows on the screen
- (void) bringToFront:(Excel4011)x;  // Bring the object to the front of the z-order of objects.
- (void) calculate:(Excel4004)x;  // Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet..
- (void) calculateFull;  // Forces a full calculation of the data in all open workbooks.
- (void) calculateFullRebuild;  // For all open workbooks, forces a full calculation of the data and rebuilds the dependencies
- (double) centimetersToPointsCentimeters:(double)centimeters;  // Converts a measurement from centimeters to points, one point equals 0.035 centimeters.
- (void) checkSpelling:(Excel4010)x customDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase alwaysSuggest:(BOOL)alwaysSuggest;  // Checks the spelling of an object.
- (BOOL) checkSpellingForTextToCheck:(NSString *)textToCheck customDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase;  // Checks the spelling of a single word. Returns True if the word is found in one of the dictionaries.
- (void) clearAllFilters:(Excel4008)x;  // The ClearAllFilters method deletes all filters currently applied to the PivotTable.
- (void) clearManualFilter:(Excel4007)x;  // The ClearManualFilter method provides an easy way to set the Visible property to True for all items of a PivotField in PivotTables, and to empty the HiddenItemsList or VisibleItemsList collections in OLAP PivotTables.
- (NSString *) convertFormulaFormulaToConvert:(NSString *)formulaToConvert fromReferenceStyle:(ExcelE158)fromReferenceStyle toReferenceStyle:(ExcelE158)toReferenceStyle toAbsolute:(ExcelE158)toAbsolute relativeTo:(ExcelE267)relativeTo;  // Converts cell references in a formula between the A1 and R1C1 reference styles, between relative and absolute references, or both
- (void) copyObject:(Excel4012)x NS_RETURNS_NOT_RETAINED;  // Copies the object to the clipboard.
- (void) copyPicture:(Excel4013)x appearance:(ExcelE207)appearance format:(ExcelE156)format NS_RETURNS_NOT_RETAINED;  // Copies the selected object to the clipboard as a picture.
- (void) cut:(Excel4014)x;  // Cuts the object to the clipboard.
- (void) delete:(Excel4006)x;  // Deletes the object.
- (void) deleteChartAutoformatName:(NSString *)name;  // Removes a custom chart autoformat from the list of available chart autoformats.
- (void) deleteCustomListListNum:(NSInteger)listNum;  // Deletes a custom list.
- (void) doubleClick;  // Equivalent to double-clicking the active cell.
- (void) drillTo:(Excel4009)x field:(NSString *)field;  // The DrillTo method supports drilling to a specified PivotField from another PivotField.
- (id) evaluateName:(id)name;  // Converts a Microsoft Excel name to an object or value.
- (ExcelBorder *) getBorder:(Excel4026)x whichBorder:(ExcelE243)whichBorder;  // Returns the specified border object.
- (NSArray *) getClipboardFormats;  // Returns a list of the  formats that are currently on the clipboard.
- (NSArray *) getCustomListContentsListNum:(NSInteger)listNum;  // Returns a custom list of strings.
- (NSInteger) getCustomListNumListArray:(NSArray *)listArray;  // Returns the custom list number for an array of strings. You can use this method to match both built-in lists and custom-defined lists.
- (ExcelDialog *) getDialog:(ExcelE245)x;  // Returns a the specified dialog object.
- (NSArray *) getFileConverters;  // Returns a list of all of the file converter objects.
- (id) getInternationalDataType:(ExcelE189)dataType;  // Returns information about a specific international setting.
- (NSString *) getListItem:(Excel4017)x entry_index:(NSInteger)entry_index;  // Returns a string from the list
- (NSString *) getOpenFilenameFileFilter:(NSString *)fileFilter buttonText:(NSString *)buttonText multiSelect:(BOOL)multiSelect;  // Displays the standard Open dialog box and gets a file name from the user without actually opening any files.
- (NSArray *) getPreviousSelections;  // Returns a list of the last four ranges or names selected. Each element in the list is a range object.
- (NSArray *) getRegisteredFunctions;  // Returns information about functions in code resources that were registered with the REGISTER or REGISTER.ID macro functions.
- (NSString *) getSaveAsFilenameInitialFilename:(NSString *)initialFilename fileFilter:(NSString *)fileFilter filterIndex:(NSInteger)filterIndex buttonText:(NSString *)buttonText;  // Displays the standard save as dialog box and gets a file name from the user without actually saving any files.
- (ExcelWebPageFont *) getWebpageFont:(ExcelMChS)x;  // Returns the specified web page font object.
- (void) gotoReference:(ExcelE267)reference scroll:(BOOL)scroll;  // Selects any range in any workbook, and activates that workbook if it's not already active.
- (void) helpHelpFile:(NSString *)helpFile helpContextId:(NSInteger)helpContextId;  // Displays the Help window.
- (double) inchesToPointsInches:(double)inches;  // Converts a measurement from inches to points.
- (SBObject *) inputBoxPrompt:(NSString *)prompt title:(NSString *)title default:(ExcelE276)default_ leftPosition:(NSInteger)leftPosition top:(NSInteger)top type:(ExcelE277)type;  // Displays a dialog box for user input. Returns the information entered in the dialog box.
- (ExcelRange *) intersectRange1:(ExcelRange *)range1 range2:(ExcelRange *)range2 range3:(ExcelRange *)range3 range4:(ExcelRange *)range4 range5:(ExcelRange *)range5 range6:(ExcelRange *)range6 range7:(ExcelRange *)range7 range8:(ExcelRange *)range8 range9:(ExcelRange *)range9 range10:(ExcelRange *)range10 range11:(ExcelRange *)range11 range12:(ExcelRange *)range12 range13:(ExcelRange *)range13 range14:(ExcelRange *)range14 range15:(ExcelRange *)range15 range16:(ExcelRange *)range16 range17:(ExcelRange *)range17 range18:(ExcelRange *)range18 range19:(ExcelRange *)range19 range20:(ExcelRange *)range20 range21:(ExcelRange *)range21 range22:(ExcelRange *)range22 range23:(ExcelRange *)range23 range24:(ExcelRange *)range24 range25:(ExcelRange *)range25 range26:(ExcelRange *)range26 range27:(ExcelRange *)range27 range28:(ExcelRange *)range28 range29:(ExcelRange *)range29 range30:(ExcelRange *)range30;  // Returns a range object that represents the rectangular intersection of two or more ranges.
- (BOOL) itemSelected:(Excel4021)x entry_index:(NSInteger)entry_index;  // Returns if a particular item in the list is selected
- (void) largeScroll:(Excel4022)x down:(NSInteger)down up:(NSInteger)up toRight:(NSInteger)toRight toLeft:(NSInteger)toLeft;  // Scrolls the contents of the window by pages.
- (void) modifyAppliesToRange:(Excel4029)x range:(ExcelRange *)range;  // Changes the range that this format condition applies to.
- (void) onKeyKey:(NSString *)key commandKeyPressed:(BOOL)commandKeyPressed shiftKeyPressed:(BOOL)shiftKeyPressed optionKeyPressed:(BOOL)optionKeyPressed controlKeyPressed:(BOOL)controlKeyPressed procedure:(ExcelE281)procedure;  // Runs a specified procedure when a particular key or key combination is pressed.
- (void) openFileMakerFileFilename:(NSString *)filename;  // Open a FileMaker file as an Excel worksheet.
- (void) openTextFileFilename:(NSString *)filename origin:(ExcelE211)origin startRow:(NSInteger)startRow dataType:(ExcelE236)dataType textQualifier:(ExcelE237)textQualifier consecutiveDelimiter:(BOOL)consecutiveDelimiter tab:(BOOL)tab semicolon:(BOOL)semicolon comma:(BOOL)comma space:(BOOL)space useOther:(BOOL)useOther otherChar:(NSString *)otherChar fieldInfo:(NSArray *)fieldInfo decimalSeparator:(NSString *)decimalSeparator thousandsSeparator:(NSString *)thousandsSeparator;  // Loads and parses a text file as a new workbook with a single sheet that contains the parsed text-file data.
- (ExcelWorkbook *) openWorkbookWorkbookFileName:(NSString *)workbookFileName updateLinks:(ExcelE294)updateLinks readOnly:(BOOL)readOnly format:(ExcelE295)format password:(NSString *)password writeReservedPassword:(NSString *)writeReservedPassword ignoreReadOnlyRecommended:(BOOL)ignoreReadOnlyRecommended origin:(ExcelE211)origin delimiter:(NSString *)delimiter editable:(BOOL)editable notify:(BOOL)notify converter:(NSInteger)converter addToMru:(BOOL)addToMru;  // Opens a workbook.
- (void) printOut:(Excel4002)x from:(NSInteger)from to:(NSInteger)to copies:(NSInteger)copies preview:(BOOL)preview activePrinter:(NSString *)activePrinter printToFile:(BOOL)printToFile collate:(BOOL)collate;  // Prints the object
- (void) printPreview:(Excel4003)x enableChanges:(BOOL)enableChanges;  // Shows a preview of the object as it would look when printed. This function has been deprecated.
- (BOOL) registerXllFilename:(NSString *)filename;  // Loads an XLL code resource and automatically registers the functions and commands contained in the resource.
- (void) removeAllItems:(Excel4019)x;  // Removes all of the strings from the list
- (void) removeItem:(Excel4020)x entry_index:(NSInteger)entry_index count:(NSInteger)count;  // Removes a specified string from the list
- (void) resetTimer:(Excel4025)x;  // Resets the refresh timer for the specified PivotTable report to the last interval you set using the RefreshPeriod property.
- (NSString *) runVBMacro:(ExcelE297)x arg1:(id)arg1 arg2:(id)arg2 arg3:(id)arg3 arg4:(id)arg4 arg5:(id)arg5 arg6:(id)arg6 arg7:(id)arg7 arg8:(id)arg8 arg9:(id)arg9 arg10:(id)arg10 arg11:(id)arg11 arg12:(id)arg12 arg13:(id)arg13 arg14:(id)arg14 arg15:(id)arg15 arg16:(id)arg16 arg17:(id)arg17 arg18:(id)arg18 arg19:(id)arg19 arg20:(id)arg20 arg21:(id)arg21 arg22:(id)arg22 arg23:(id)arg23 arg24:(id)arg24 arg25:(id)arg25 arg26:(id)arg26 arg27:(id)arg27 arg28:(id)arg28 arg29:(id)arg29 arg30:(id)arg30;  // Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.
- (NSString *) runXLMMacro:(ExcelE297)x arg1:(id)arg1 arg2:(id)arg2 arg3:(id)arg3 arg4:(id)arg4 arg5:(id)arg5 arg6:(id)arg6 arg7:(id)arg7 arg8:(id)arg8 arg9:(id)arg9 arg10:(id)arg10 arg11:(id)arg11 arg12:(id)arg12 arg13:(id)arg13 arg14:(id)arg14 arg15:(id)arg15 arg16:(id)arg16 arg17:(id)arg17 arg18:(id)arg18 arg19:(id)arg19 arg20:(id)arg20 arg21:(id)arg21 arg22:(id)arg22 arg23:(id)arg23 arg24:(id)arg24 arg25:(id)arg25 arg26:(id)arg26 arg27:(id)arg27 arg28:(id)arg28 arg29:(id)arg29 arg30:(id)arg30;  // Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.
- (void) saveWorkspaceWorkspaceFileName:(NSString *)workspaceFileName;  // Saves the current workspace.
- (void) sendToBack:(Excel4015)x;  // Sends the object to the back of the z-order.
- (void) setFirstPriority:(Excel4027)x;  // Sets this condition to the highest priority.
- (void) setLastPriority:(Excel4028)x;  // Sets this condition to the lowest priority.
- (void) setListItem:(Excel4018)x entry_index:(NSInteger)entry_index itemText:(NSString *)itemText;  // Set a string in the list
- (void) show:(Excel4024)x;  // Displays the built-in dialog box and waits for the user to input data.
- (void) smallScroll:(Excel4023)x down:(NSInteger)down up:(NSInteger)up toRight:(NSInteger)toRight toLeft:(NSInteger)toLeft;  // Scrolls the contents of the window by rows or columns.
- (void) undo;  // Cancels the last user-interface action.
- (ExcelRange *) unionRange1:(ExcelRange *)range1 range2:(ExcelRange *)range2 range3:(ExcelRange *)range3 range4:(ExcelRange *)range4 range5:(ExcelRange *)range5 range6:(ExcelRange *)range6 range7:(ExcelRange *)range7 range8:(ExcelRange *)range8 range9:(ExcelRange *)range9 range10:(ExcelRange *)range10 range11:(ExcelRange *)range11 range12:(ExcelRange *)range12 range13:(ExcelRange *)range13 range14:(ExcelRange *)range14 range15:(ExcelRange *)range15 range16:(ExcelRange *)range16 range17:(ExcelRange *)range17 range18:(ExcelRange *)range18 range19:(ExcelRange *)range19 range20:(ExcelRange *)range20 range21:(ExcelRange *)range21 range22:(ExcelRange *)range22 range23:(ExcelRange *)range23 range24:(ExcelRange *)range24 range25:(ExcelRange *)range25 range26:(ExcelRange *)range26 range27:(ExcelRange *)range27 range28:(ExcelRange *)range28 range29:(ExcelRange *)range29 range30:(ExcelRange *)range30;  // Returns the union of two or more ranges.
- (void) unprotect:(Excel4005)x password:(NSString *)password;  // Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.
- (void) bringToFront:(Excel4030)x;  // Bring the object to the front of the z-order of objects.
- (void) checkSpelling:(Excel4035)x customDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase alwaysSuggest:(BOOL)alwaysSuggest;  // Checks the spelling of an object.
- (void) copyObject:(Excel4031)x NS_RETURNS_NOT_RETAINED;  // Copies the object to the clipboard.
- (void) copyPicture:(Excel4032)x appearance:(ExcelE207)appearance format:(ExcelE156)format NS_RETURNS_NOT_RETAINED;  // Copies the selected object to the clipboard as a picture.
- (void) customDrop:(Excel4036)x drop:(double)drop;  // Sets the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.
- (void) customLength:(Excel4037)x length:(double)length;  // Specifies that the first segment of the callout line, the segment attached to the text callout box, retain a fixed length whenever the callout is moved.
- (void) cut:(Excel4033)x;  // Cuts the object to the clipboard.
- (void) presetDrop:(Excel4038)x dropType:(ExcelMCDt)dropType;  // Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.
- (void) sendToBack:(Excel4034)x;  // Sends the object to the back of the z-order.
- (void) clearHyperlinks;  // clear all the hyperlinks in a range
- (void) dirty;  // Designates a range to be recalculated when the next recalculation occurs.
- (void) activateObject:(Excel4039)x;  // Activates the object.
- (void) applyDataLabels:(Excel4049)x type:(ExcelE134)type legendKey:(BOOL)legendKey autoText:(BOOL)autoText hasLeaderLines:(BOOL)hasLeaderLines showSeriesName:(BOOL)showSeriesName showCategoryName:(BOOL)showCategoryName showValue:(BOOL)showValue showPercentage:(BOOL)showPercentage showBubbleSize:(BOOL)showBubbleSize separator:(NSString *)separator;  // Applies data labels to all the series in a chart, a single series or a series point.
- (void) chartOneColorGradient:(Excel4042)x gradientStyle:(ExcelMGdS)gradientStyle variant:(NSInteger)variant degree:(double)degree;  // Sets the specified fill to a one-color gradient.
- (void) chartPatterned:(Excel4041)x pattern:(ExcelPpTy)pattern;  // Sets the specified fill to a pattern.
- (void) chartSolid:(Excel4046)x;  // Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.
- (void) chartTwoColorGradient:(Excel4045)x gradientStyle:(ExcelMGdS)gradientStyle variant:(NSInteger)variant;  // Sets the specified fill to a two-color gradient.
- (void) chartUserPicture:(Excel4044)x pictureFile:(NSString *)pictureFile pictureFormat:(ExcelE124)pictureFormat pictureStackUnit:(NSInteger)pictureStackUnit picturePlacement:(ExcelE125)picturePlacement;  // Fills the specified shape with an image.
- (void) chartUserTextured:(Excel4043)x textureFile:(NSString *)textureFile;  // Fills the specified shape with small tiles of an image.
- (void) clear:(Excel4053)x;  // Clear the object.
- (void) clearFormats:(Excel4051)x;  // Clears the formatting of the object.
- (void) copyObject:(Excel4050)x NS_RETURNS_NOT_RETAINED;  // Copies the object to the clipboard.
- (void) deleteObject:(Excel4040)x;  // Deletes the object.
- (void) paste:(Excel4052)x;
- (void) presetChartGradient:(Excel4048)x gradientStyle:(ExcelMGdS)gradientStyle variant:(NSInteger)variant presetGradientType:(ExcelMPGb)presetGradientType;  // Sets the specified fill to a preset gradient.
- (void) presetChartTextured:(Excel4047)x presetTextureForChart:(ExcelMPzT)presetTextureForChart;  // Sets the specified fill format to a preset texture.

@end

// Represents autofiltering for the specified worksheet.
@interface ExcelAutofilter : ExcelBaseObject

- (SBElementArray<ExcelFilter *> *) filters;

@property (readonly) BOOL autofiltermode;  // Returns True if filters have been defined else False.
@property (copy, readonly) ExcelRange *rangeObject;  // Returns a range object that represents the range to which the specified autofilter applies.
@property (copy, readonly) ExcelSort *sortObject;  // Returns the sort object in the auto filter object.


@end

// Represents the border of an object.
@interface ExcelBorder : ExcelBaseObject

@property (copy) NSColor *color;  // Returns or sets the primary color of the object. 
@property ExcelE103 colorIndex;  // Returns or sets the color of the border. The color is specified as an index value into the current color palette.
@property ExcelE133 lineStyle;  // Returns or sets the line style for the border.
@property NSInteger lineWeight;  // Returns or sets the line weight of the border.
@property ExcelE305 themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.
@property ExcelE128 weight;  // Returns or sets the weight of the border.


@end

// Represents a button control.
@interface ExcelButton : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property BOOL autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property BOOL cancelButton;  // Returns or sets if this button is the cancel button.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL defaultButton;  // Returns or sets if this button is the default button.
@property BOOL dismissButton;  // Returns or sets if this button is the dismiss button.
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property double height;  // Returns or set the height of the object.
@property BOOL helpButton;  // Returns or sets if this button is the help button.
@property ExcelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelE126 orientation;  // May also be a number value from -90 to 90 degrees.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property ExcelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property ExcelE284 verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property BOOL wrapAutoText;  // Returns or sets if the auto text is wrapped.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents the calculated fields and calculated items for PivotTables with Online Analytical Processing data sources.
@interface ExcelCalculatedMember : ExcelBaseObject

@property (copy, readonly) NSString *_default;
@property (copy, readonly) NSString *displayFolder;  // An ST_Xstring attribute that specifies the display folder of this PivotTable named set.
@property (readonly) BOOL dynamic;
@property (readonly) BOOL flattenHierarchies;
@property (copy, readonly) NSString *formula;  // Returns the member's formula in multidimensional expressions syntax.
@property (readonly) BOOL hierarchizeDistinct;
@property (readonly) BOOL isValid;  // Returns a Boolean that indicates whether the specified calculated member has been successfully instantiated with the OLAP provider during the current session.
@property (copy, readonly) NSString *name;  // Calculated Member Name.
@property (readonly) NSInteger solveOrder;  // Calculated Members Solve Order.
@property (copy, readonly) NSString *sourceName;  // Returns the specified object's name as it appears in the original source data for the specified PivotTable report.
@property (readonly) ExcelE916 type;


@end

@interface ExcelCheckbox : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL displayThreeDShading;
@property BOOL enabled;  // Returns or sets if the control is enabled
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or sets the height of the control.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the control
@property (copy) NSString *linkedCell;
@property BOOL locked;  // Returns or sets whether the control is locked for editing.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of this control
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property ExcelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property ExcelE285 value;
@property BOOL visible;  // Returns or sets the current value of the control
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

@interface ExcelColorScaleCriteria : ExcelBaseObject


@end

@interface ExcelColorScaleCriterion : ExcelBaseObject

@property (readonly) NSInteger colorScaleCriterionIndex;  // The index of the color scale criterion. Read only.
@property ExcelE306 colorScaleCriterionType;  // Returns or sets the condition value type of the color scale criterion. Read/Write.
@property (copy) id colorScaleCriterionValue;
@property (copy, readonly) ExcelFormatColor *formatColor;  // Returns the FormatColor for the ColorScaleCriterion object.


@end

@interface ExcelColorScaleFormatCondition : ExcelBaseObject

@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property (copy, readonly) ExcelColorScaleCriteria *colorScaleCriteria;  // Returns the ColorScaleCriteria for the ColorScale object.
@property NSInteger colorScaleType;
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) ExcelE195 formatConditionType;  // Return the conditional format type.
@property (copy) NSString *formula;  // Returns or sets the value or expression associated with the conditional format or data validation.
@property ExcelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

// Represents the color stop point for a gradient fill in an range or selection.
@interface ExcelColorstop : ExcelBaseObject

@property (copy) NSColor *color;  // Returns or sets the primary color of the object.
@property double colorstopPosition;  // Returns or sets the position of the ColorStop.
@property ExcelE305 themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.

- (void) deleteColorstop;  // Clears the represented object.

@end

// A collection of all the ColorStop objects for the specified series.
@interface ExcelColorstops : ExcelBaseObject

- (ExcelColorstop *) addColorstopPosition:(double)position;  // Adds ColorStop object to the ColorStops object.

@end

@interface ExcelConditionValue : ExcelBaseObject

@property (readonly) ExcelE306 conditionValueType;
@property (copy, readonly) id conditionValueValue;

- (void) modifyConditionValueType:(ExcelE306)type conditionValue:(id)conditionValue;  // Modifies an existing condition value.

@end

// Represents a hierarchy or measure field from an OLAP cube
@interface ExcelCubeField : ExcelBaseObject

- (SBElementArray<ExcelPivotField *> *) pivotFields;

@property (readonly) BOOL allItemsVisible;  // The AllItemsVisible property checks whether manual filtering is applied to a PivotField or CubeField.
@property (copy) NSString *caption;  // The label text for the cube field.
@property (readonly) ExcelE914 cubeFieldSubType;  // Specifies the type of a CubeField.
@property (readonly) ExcelE913 cubeFieldType;  // Indicates whether the OLAP cube field is a hierarchy field or a measure field.
@property (copy) NSString *currentPageName;  // Returns or sets the page name for a CubeField.
@property BOOL dragToColumn;  // True if the specified field can be dragged to the column position.
@property BOOL dragToData;  // True if the specified field can be dragged to the data position.
@property BOOL dragToHide;  // True if the specified field can be dragged to the column position.
@property BOOL dragToPage;  // True if the field can be dragged to the page position.
@property BOOL dragToRow;  // True if the field can be dragged to the row position.
@property BOOL enableMultiplePageItems;  // True to allow multiple items in the page field area for OLAP PivotTables to be selected.
@property BOOL flattenHierarchies;
@property (readonly) BOOL hasMemberProperties;  // Returns True when there are member properties specified to be displayed for the cube field.
@property BOOL hierarchizeDistinct;
@property BOOL includeNewItemsInFilter;  // The IncludeNewItemsInFilter property is used to track included or excluded items in OLAP PivotTables.
@property (readonly) BOOL isDate;  // Returns True if the CubeField is a date.
@property (readonly) ExcelE910 layoutForm;  // Returns or sets the way the specified PivotTable items appear -- in table format or in outline format.
@property ExcelE902 layoutSubtotalLocation;  // Returns or sets the position of the PivotTable field subtotals in relation to, either above or below, the specified field.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property ExcelE208 orientation;  // The location of the field in the specified PivotTable report.
@property NSInteger position;  // Position of the item in its field if the item is currently showing.
@property BOOL showInFieldList;  // When set to True, a CubeField object will be shown in the field list.
@property (copy, readonly) ExcelTreeviewControl *treeviewControl;
@property (copy, readonly) NSString *value;  // Returns the name of the specified field.

- (void) addMemberPropertyFieldProperty:(NSString *)property propertyOrder:(NSString *)propertyOrder propertyDisplayedIn:(ExcelE915)propertyDisplayedIn;  // Adds a member property field to the display for the cube field.
- (void) createPivotFields;  // The CreatePivotFields method enables users to create all PivotFields of a CubeField.

@end

// Represents a custom workbook view.
@interface ExcelCustomView : ExcelBaseObject

@property (readonly) BOOL customViewPrintSettings;  // True if print settings are included in the custom view.
@property (copy, readonly) NSString *name;  // Returns the name of this object.
@property (readonly) BOOL rowColSettings;  // Returns true if the custom view includes settings for hidden rows and columns, including filter information.

- (void) showCustomView;  // Displays the custom view

@end

@interface ExcelDatabarBorder : ExcelBaseObject

@property (copy, readonly) ExcelFormatColor *databarBorderColor;
@property ExcelE309 databarBorderType;  // Returns or sets the type of border of the databar


@end

@interface ExcelDatabarFormatCondition : ExcelBaseObject

@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property (copy, readonly) ExcelFormatColor *databarAxisColor;
@property ExcelE310 databarAxisPosition;  // Returns or sets the position of the databar axis
@property (copy, readonly) ExcelFormatColor *databarBarColor;
@property (copy, readonly) ExcelDatabarBorder *databarBorder;  // Returns the DataBarBorder for the Databar object.
@property NSInteger databarDirection;  // Specifies the direction of the databar. Read/Write.
@property ExcelE308 databarFillType;  // Returns or sets the type of fill of the databar
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property BOOL formatConditionShowValue;  // Specifies whether to show the cell value along with the databar. Read/Write.
@property (readonly) ExcelE195 formatConditionType;  // Return the conditional format type.
@property (copy) NSString *formula;  // Returns or sets the value or expression associated with the conditional format or data validation.
@property (copy, readonly) ExcelConditionValue *maxPointConditionValue;  // Returns the ConditionValue for the maximum point of the DataBar object.
@property NSInteger maxPointPercent;  // Specifies the percentage of the data bar to draw at the maximum point. Read/Write.
@property (copy, readonly) ExcelConditionValue *minPointConditionValue;  // Returns the ConditionValue for the minimum point of the DataBar object.
@property NSInteger minPointPercent;  // Specifies the percentage of the data bar to draw at the minimum point. Read/Write.
@property (copy, readonly) ExcelNegativeBarFormat *negativeBarFormat;  // Returns the NegativeBarFormat for the Databar object.
@property ExcelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

// Contains global application-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.
@interface ExcelDefaultWebOptions : ExcelBaseObject

@property BOOL allowPng;  // Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save documents as a Web page.
@property BOOL alwaysSaveInDefaultEncoding;  // Returns or sets if the default encoding is used when you save a Web page or plain text document, independent of the file's original encoding when opened.
@property ExcelMtEn encoding;  // Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document.
@property (copy) NSString *locationOfComponents;  // Returns or sets the central URL, on the intranet or Web, or path, local or network, to the location from which authorized users can download Microsoft Office Web components when viewing your saved document.
@property NSInteger pixelsPerInch;  // Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120.
@property ExcelMSsz screenSize;  // Returns or sets the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser.
@property BOOL updateLinksOnSave;  // Returns or sets if hyperlinks and paths to all supporting files are automatically updated before you save the document as a Web page, ensuring that the links are up-to-date at the time the document is saved. If false, the links are not updated.


@end

// Represents a built-in Microsoft Excel dialog box.
@interface ExcelDialog : ExcelBaseObject


@end

// Represents the formatting shown to the user.
@interface ExcelDisplayFormat : ExcelBaseObject

@property (readonly) BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (readonly) BOOL formulaHidden;  // Returns or sets if the formula will be hidden when the workbook or worksheet is protected.
@property (readonly) ExcelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the range.
@property (readonly) NSInteger indentLevel;  // Returns or sets the indent level for the range or style. Can be an integer from 0 to 15.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (readonly) BOOL locked;  // Returns or sets  if the range is locked.
@property (readonly) BOOL mergeCells;  // Returns or sets if the range contains merged cells. 
@property (copy, readonly) NSString *numberFormat;  // Returns or sets the format code for the object.
@property (readonly) ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property (readonly) BOOL shrinkToFit;  // Returns or sets if text automatically shrinks to fit in the available column width.
@property (copy, readonly) ExcelStyle *styleObject;  // Returns or sets a style object that represents the style of the specified range.
@property (readonly) ExcelE126 textOrientation;  // The text orientation. Can be a number value from -90 to 90 degrees.
@property (readonly) ExcelE284 verticalAlignment;  // Returns or sets the vertical alignment of the range.
@property (readonly) BOOL wrapText;  // Returns or sets if Microsoft Excel wraps the text in the object.


@end

@interface ExcelDropdown : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying. 
@property (copy) NSString *caption;  // Returns or sets the caption for the control.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control.
@property BOOL displayThreeDShading;  // Returns or sets whether a 3-d effect will be employed when displaying the control
@property NSInteger dropDownLines;  // Returns or sets the number of dropdown items.
@property BOOL enabled;  // Returns or sets if the control is enabled
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or sets the height of the control.
@property double leftPosition;  // Returns or sets the left position of the control
@property (copy) NSString *linkedCell;  // Returns or sets reference to a call which contains the value of the control.
@property (copy) NSString *listFillRange;  // Returns or sets the items which are contained in the drop down.
@property NSInteger listIndex;  // Returns or sets the selected item.
@property BOOL locked;  // Returns or sets whether the control is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of this control.
@property (readonly) NSInteger numberOfItemsInList;  // Returns the number of total items in the list.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property NSInteger value;  // Returns or sets the current value of the control
@property BOOL visible;  // Returns or sets if the control is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// Represents a filter for a single column. 
@interface ExcelFilter : ExcelBaseObject

@property (copy, readonly) id criteria1;  // Returns the first filtered value for the specified column in a filtered range.
@property (copy, readonly) id criteria2;  // Returns the second filtered value for the specified column in a filtered range.
@property (readonly) BOOL filterOn;  // Returns true if the specified filter is on.
- (ExcelE186) operator;  // set or return the operator that associates the two criteria applied by the specified filter.
- (void) setOperator: (ExcelE186) operator;


@end

// Represents a color that a format condition can be colored with.
@interface ExcelFormatColor : ExcelBaseObject

@property (copy) NSColor *color;  // Returns or sets the primary color of the object.
@property ExcelE103 colorIndex;
@property ExcelE305 themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.


@end

@interface ExcelFormatConditionIconObject : ExcelBaseObject

@property (readonly) NSInteger formatConditionIconIndex;  // The index of the icon. Read only.


@end

@interface ExcelFormatConditionIconSet : ExcelBaseObject

@property (readonly) ExcelE313 iconSetId;


@end

@interface ExcelFormatConditionIconSets : ExcelBaseObject


@end

// Represents a conditional format.
@interface ExcelFormatCondition : ExcelBaseObject

@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property (readonly) ExcelE196 conditionOperator;  // Returns the operator for the conditional format.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property ExcelE319 formatConditionDateOperator;
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (copy) NSString *formatConditionText;  // Returns or sets the text for the specified object.
@property ExcelE318 formatConditionTextOperator;
@property (readonly) ExcelE195 formatConditionType;  // Return the conditional format type.
@property (copy, readonly) NSString *formula1;  // Returns the value or expression associated with the conditional format or data validation.
@property (copy, readonly) NSString *formula2;  // Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is operator between or operator not between.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property ExcelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read/Write.

- (void) modifyConditionType:(ExcelE195)type operator:(ExcelE196)operator_ formula1:(NSString *)formula1 formula2:(NSString *)formula2 string:(NSString *)string operator2:(ExcelE196)operator2;  // Modifies an existing conditional format.

@end

// Contains properties that apply to header and footer picture objects.
@interface ExcelGraphic : ExcelBaseObject

@property double brightness;  // Returns or sets the brightness of the specified picture. The value for this property must be a number from 0.0, dimmest, to 1.0, brightest.
@property ExcelMPc colorType;  // Returns or sets the type of color transformation applied to the specified picture.
@property double contrast;  // Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast, to 1.0, the greatest contrast.
@property double cropBottom;  // Returns or sets the number of points that are cropped off the bottom of the specified picture.
@property double cropLeft;  // Returns or sets the number of points that are cropped off the left side of the specified picture.
@property double cropRight;  // Returns or sets the number of points that are cropped off the right side of the specified picture.
@property double cropTop;  // Returns or sets the number of points that are cropped off the top of the specified picture.
@property (copy) NSString *fileName;  // Returns or sets the URL, on the intranet or the Web, or path, local or network, to the location where the specified source object was saved.
@property double height;  // Returns or sets the height of the object.
@property BOOL lockAspectRatio;  // Returns or sets if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it.
@property double width;  // Returns or sets the width of the object.


@end

@interface ExcelGroupbox : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL displayThreeDShading;
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or set the height of the object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property ExcelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property BOOL visible;  // Returns or sets if the control is visible
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// Represents a horizontal page break.
@interface ExcelHorizontalPageBreak : ExcelBaseObject

@property (readonly) ExcelE190 extent;  // Returns the type of the specified page break: full-screen or only within a print area.
@property ExcelE168 horizontalPageBreakType;  // Returns or sets the type of horizontal page break.
@property (copy) ExcelRange *location;  // Returns or sets where this horizontal page break is located.


@end

// Represents a hyperlink.
@interface ExcelHyperlink : ExcelBaseObject

@property (copy) NSString *address;  // Returns or sets the address of the target document.
@property (copy) NSString *emailSubject;  // Returns or sets the text string of the specified hyperlink's e-mail subject line. The subject line is appended to the hyperlink's address.
@property (readonly) ExcelMHlT hyperlinkType;  // Returns the hyperlink type, what the hyperlink is associated with.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy, readonly) ExcelRange *rangeObject;  // Returns a range object that represents the range the specified hyperlink is attached to.
@property (copy) NSString *screenTip;  // Returns or sets the screen tip text for the specified hyperlink.
@property (copy, readonly) ExcelShape *shapeObject;  // Returns a shape object that represents the shape attached to the specified hyperlink.
@property (copy) NSString *subAddress;  // Returns or sets the location within the document associated with the hyperlink.
@property (copy) NSString *textToDisplay;  // Returns or sets the text to be displayed for the specified hyperlink. The default value is the address of the hyperlink.

- (void) createNewDocumentFileName:(NSString *)fileName editNow:(BOOL)editNow overwrite:(BOOL)overwrite;  // Creates a new document linked to the specified hyperlink.
- (void) followNewWindow:(BOOL)newWindow extraInfo:(NSString *)extraInfo method:(ExcelMXiM)method headerInfo:(NSString *)headerInfo;  // Displays a cached document, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.

@end

@interface ExcelIconCriteria : ExcelBaseObject


@end

@interface ExcelIconCriterion : ExcelBaseObject

@property ExcelE196 conditionOperator;  // Returns the operator for the conditional format.
@property ExcelE312 iconCriterionIcon;
@property (readonly) NSInteger iconCriterionIndex;  // The index of the icon criterion. Read only.
@property ExcelE306 iconCriterionType;  // Returns or sets the condition value type of the icon criterion. Read/Write.
@property (copy) id iconCriterionValue;


@end

@interface ExcelIconSetFormatCondition : ExcelBaseObject

@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property (copy) ExcelFormatConditionIconSet *formatConditionIconSet;
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) ExcelE195 formatConditionType;  // Return the conditional format type.
@property (copy) NSString *formula;  // Returns or sets the value or expression associated with the conditional format or data validation.
@property (copy, readonly) ExcelIconCriteria *iconCriteria;  // Returns the IconCriteria for the IconSetCondition object.
@property BOOL percentileValues;
@property ExcelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property BOOL reverseIconSetOrder;
@property BOOL showIconOnly;
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

@interface ExcelLabel : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for the control.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL enabled;  // Returns or sets if the control is enabled
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or sets the height of the control.
@property double leftPosition;  // Returns or sets the left position of the control
@property BOOL locked;  // Returns or sets whether the control is locked for editing.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of this control
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property ExcelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property BOOL visible;  // Returns or sets if the control is visible
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// The LinearGradient object transitions through a series of colors in a linear manner along a specific angle.
@interface ExcelLinearGradient : ExcelBaseObject

@property (copy, readonly) ExcelColorstops *colorstops;  // Returns the ColorStops for the LinearGradient object.
@property double linearGradientDegree;  // The angle of the linear gradient fill within a selection.


@end

// Represents a column in a list object.
@interface ExcelListColumn : ExcelBaseObject

@property (copy, readonly) ExcelCell *cellTable;  // Returns the cell table from a specified list column. 
@property (readonly) NSInteger index;
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) ExcelRange *rangeObject;  // Returns a range object that represents the range to which the specified list column.
@property (copy, readonly) ExcelRange *totalRow;  // Returns the totals row, if any, from a specified list column.
@property ExcelE303 totalsCalculation;


@end

// Represents a list object on a worksheet.
@interface ExcelListObject : ExcelBaseObject

- (SBElementArray<ExcelListColumn *> *) listColumns;
- (SBElementArray<ExcelListRow *> *) listRows;

@property (readonly) BOOL active;
@property (copy, readonly) ExcelAutofilter *autofilterObject;  // Returns the autofilter object associated with this list object.
@property (copy, readonly) ExcelCell *cellTable;  // Returns the cell table from a specified list object. 
@property (copy) NSString *comment;  // Returns or sets the comment of the object.
- (NSString *) default;
@property (copy) NSString *displayName;  // Returns or sets the display name of the object.
@property (readonly) BOOL displayRightToLeft;
@property (copy, readonly) ExcelRange *headerRow;  // Returns a range object that represents the used range on the specified worksheet. 
@property (copy, readonly) ExcelRange *insertRow;
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) ExcelQueryTable *queryTable;  // Returns a query table object that represents the query table that intersects the specified range.
@property (copy) ExcelRange *rangeObject;  // Returns or sets a range object that represents the range to which the specified list column, object, or row applies.
@property BOOL showAutofilter;  // Returns or sets if the autofilter is implemented in a list object.
@property BOOL showHeaders;  // Returns or sets if headers is implemented in a list object.
@property BOOL showTableStyleColumnStripes;  // The ShowTableStyleColumnStripes property displays banded columns in which even columns are formatted differently from odd columns.
@property BOOL showTableStyleFirstColumn;  // Returns or sets if the first column is displayed for the specified ListObject object.
@property BOOL showTableStyleLastColumn;  // Returns or sets if the last column is displayed for the specified ListObject object.
@property BOOL showTableStyleRowStripes;  // The ShowTableStyleRowStripes property displays banded rows in which even rows are formatted differently from odd rows.
@property (copy, readonly) ExcelSort *sortObject;  // Returns the sort object associated with this list object.
@property (readonly) ExcelE253 sourceType;
@property (copy) id tableStyle;  // Returns or sets the style used in the table body.
@property BOOL total;  // Returns or sets if a totals row be implemented in a list object.
@property (copy, readonly) ExcelRange *totalRow;  // Returns the totals row, if any, from a specified list object.

- (void) clearContents;  // Clears the formulas from the range. Clears the data from a chart but leaves the formatting. Clears all the data, formatting, and formulas from a list object.
- (ExcelRange *) convertToRange;  // Converts a list object to a normal Excel range.
- (void) resizeRange:(ExcelRange *)range;

@end

// Represents a row in a list object.
@interface ExcelListRow : ExcelBaseObject

@property (readonly) NSInteger index;
@property (copy, readonly) ExcelRange *rangeObject;  // Returns a range object that represents the range to which the specified list row applies.


@end

@interface ExcelListbox : ExcelBaseObject

@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property BOOL displayThreeDShading;  // Returns or sets whether a 3-d effect will be employed when displaying the control
@property BOOL enabled;  // Returns or sets if the control is enabled
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or sets the height of the control.
@property double leftPosition;  // Returns or sets the left position of the control
@property (copy) NSString *linkedCell;  // Returns or sets reference to a call which contains the value of the control.
@property (copy) NSString *listFillRange;  // Returns or sets the items which are contained in the control.
@property NSInteger listIndex;  // Returns or sets the selected item.
@property BOOL locked;  // Returns or sets whether the control is locked for editing.
@property ExcelE290 multiSelect;  // Returns or sets the multiple selection setting.
@property (copy) NSString *name;  // Returns or sets the name of this control
@property (readonly) NSInteger numberOfItemsInList;
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned. 
@property NSInteger value;  // Returns or sets the current value of the control.
@property BOOL visible;  // Returns or sets if the control is visible
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// Represents a defined name for a range of cells. Named items can be either built-in names, such as Database, Print_Area, and Auto_Open, or custom names.
@interface ExcelNamedItem : ExcelBaseObject

@property (copy) NSString *category;  // Returns or sets the category for the specified name in the language of the macro.
@property (copy) NSString *categoryLocal;  // Returns or sets the category for the specified name, in the language of the user, if the name refers to a custom function or command.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property ExcelE240 macroType;  // Returns or sets what the name refers to.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *nameLocal;  // Returns or sets the name of the object, in the language of the user.
@property (copy) NSString *referenceLocal;  // Returns or sets the formula that the name refers to. The formula is in the language of the user, and it's in A1-style notation, beginning with an equal sign.
@property (copy) NSString *referenceLocalR1c1;  // Returns or sets the formula that the name refers to. This formula is in the language of the user, and it's in R1C1-style notation, beginning with an equal sign.
@property (copy) NSString *referenceR1c1;  // Returns or sets the formula that the name refers to. The formula is in the language of the macro, and it's in R1C1-style notation, beginning with an equal sign.
@property (copy, readonly) ExcelRange *referenceRange;  // Returns the range object referred to by this object.
@property (copy) NSString *references;  // Returns or sets the formula that the name is defined to refer to, in the language of the macro and in A1-style notation, beginning with an equal sign.
@property (copy) NSString *shortcutKey;  // Returns or sets the shortcut key for a name defined as a custom Microsoft Excel 4.0 macro command.
@property (copy) NSString *value;  // Returns or sets a string containing the formula that the name is defined to refer to. The string is in A1-style notation in the language of the macro, and it begins with an equal sign.
@property BOOL visible;  // Returns or sets if the object is visible.


@end

@interface ExcelNegativeBarFormat : ExcelBaseObject

@property (copy, readonly) ExcelFormatColor *databarBarColor;
@property (copy, readonly) ExcelFormatColor *databarBorderColor;  // Returns the DataBarBorder for the Databar object.
@property ExcelE320 negativeBarBorderColorType;  // Returns or sets the type of border of the databar
@property ExcelE320 negativeBarColorType;  // Returns or sets the type of fill of the databar


@end

@interface ExcelOptionButton : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL displayThreeDShading;
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) ExcelGroupbox *groupBox;
@property double height;  // Returns or set the height of the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property (copy) NSString *linkedCell;
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property ExcelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property ExcelE285 value;
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents an outline on a worksheet.
@interface ExcelOutline : ExcelBaseObject

@property BOOL automaticStyles;  // Returns or sets if the outline uses automatic styles.
@property ExcelE233 summaryColumn;  // Returns or sets the location of the summary columns in the outline, as shown in the following table.
@property ExcelE232 summaryRow;  // Returns or sets the location of the summary rows in the outline, as shown in the following table.

- (void) showLevelsRowLevels:(NSInteger)rowLevels columnLevels:(NSInteger)columnLevels;  // Displays the specified number of row and/or column levels of an outline.

@end

// Represents the page setup description. The page setup object contains all page setup attributes, left margin, bottom margin, paper size, and so on, as properties.
@interface ExcelPageSetup : ExcelBaseObject

@property BOOL blackAndWhite;  // Returns or sets if elements of the document will be printed in black and white.
@property double bottomMargin;  // Returns or sets the size of the bottom margin, in points.
@property (copy) NSString *centerFooter;  // Returns or sets the center part of the footer.
@property (copy, readonly) ExcelGraphic *centerFooterPicture;  // Returns or sets a graphic object that represents the picture for the center section of the footer. Used to set attributes about the picture.
@property (copy) NSString *centerHeader;  // Returns or sets the center part of the header.
@property (copy, readonly) ExcelGraphic *centerHeaderPicture;  // Returns or sets a graphic object that represents the picture for the center section of the footer. Used to set attributes about the picture.
@property BOOL centerHorizontally;  // Returns or sets  if the sheet is centered horizontally on the page when it's printed.
@property BOOL centerVertically;  // Returns or sets if the sheet is centered vertically on the page when it's printed.
@property ExcelE176 chartSize;  // Returns or sets the way a chart is scaled to fit on a page.
@property BOOL draft;  // Returns or sets if the sheet will be printed without graphics.
@property NSInteger firstPageNumber;  // Returns or sets the first page number that will be used when this sheet is printed.
@property NSInteger fitToPagesTall;  // Returns or sets the number of pages tall the worksheet will be scaled to when it's printed. Applies only to worksheets.
@property NSInteger fitToPagesWide;  // Returns or sets the number of pages wide the worksheet will be scaled to when it's printed. Applies only to worksheets.
@property double footerMargin;  // Returns or sets the distance from the bottom of the page to the footer, in points.
@property double headerMargin;  // Returns or sets the distance from the top of the page to the header, in points.
@property (copy) NSString *leftFooter;  // Returns or sets the left part of the footer.
@property (copy, readonly) ExcelGraphic *leftFooterPicture;  // Returns or sets a graphic object that represents the picture for the left section of the footer. Used to set attributes about the picture.
@property (copy) NSString *leftHeader;  // Returns or sets the left part of the header.
@property (copy, readonly) ExcelGraphic *leftHeaderPicture;  // Returns or sets a graphic object that represents the picture for the left section of the header. Used to set attributes about the picture.
@property double leftMargin;  // Returns or sets the size of the left margin, in points.
@property ExcelE164 order;  // Returns or sets the order that Microsoft Excel uses to number pages when printing a large worksheet.
@property ExcelE170 pageOrientation;  // Returns or set if the printing mode will be portrait or landscape.
@property ExcelE212 printExcelComments;  // Returns or sets the way comments are printed with the sheet.
@property (copy) NSString *printArea;  // Returns or sets the range to be printed, as a string using A1-style references in the language of the macro.
@property BOOL printGridlines;  // Returns or sets if cell gridlines are printed on the page. Applies only to worksheets.
@property BOOL printHeadings;  // Returns or sets if row and column headings are printed with this page. Applies only to worksheets.
@property BOOL printNotes;  // Returns or sets if cell notes are printed as end notes with the sheet. Applies only to worksheets.
@property (copy) NSArray *printQuality;  // Set/Gets a two element list where 1 is the horizontal print quality and 2 is the vertical print quality
@property (copy) NSString *printTitleColumns;  // Returns or sets the columns that contain the cells to be repeated on the left side of each page, as a string in A1-style notation in the language of the macro.
@property (copy) NSString *printTitleRows;  // Returns or sets the rows that contain the cells to be repeated at the top of each page, as a string in A1-style notation in the language of the macro.
@property (copy) NSString *rightFooter;  // Returns or sets the right part of the footer.
@property (copy, readonly) ExcelGraphic *rightFooterPicture;  // Returns or sets a graphic object that represents the picture for the right section of the footer. Used to set attributes about the picture.
@property (copy) NSString *rightHeader;  // Returns or sets the right part of the header.
@property (copy, readonly) ExcelGraphic *rightHeaderPicture;  // Returns or sets a graphic object that represents the picture for the right section of the header. Used to set attributes about the picture.
@property double rightMargin;  // Returns or sets the size of the right margin, in points.
@property double topMargin;  // Returns or sets the size of the top margin, in points.
@property ExcelE280 zoom;  // Returns or sets a percentage, between 10 and 400 percent, by which Microsoft Excel will scale the worksheet for printing. Applies only to worksheets.


@end

// Represents a pane of a window. Pane objects exist only for worksheets and Microsoft Excel 4.0 macro sheets.
@interface ExcelPane : ExcelBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property NSInteger scrollColumn;  // Returns or sets the number of the leftmost column in the pane
@property NSInteger scrollRow;  // Returns or sets the number of the row that appears at the top of the pane.
@property (copy, readonly) ExcelRange *visibleRange;  // Returns a Range object that represents the range of cells that are visible in the pane. If a column or row is partially visible, it's included in the range.


@end

// Contains information about a specific phonetic text string in a cell.
@interface ExcelPhonetic : ExcelBaseObject

@property ExcelE205 characterType;  // Returns or sets the type of phonetic text in the specified cell.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property ExcelE206 phoneticAlignment;  // Returns or sets the alignment for the specified phonetic text.
@property (copy) NSString *phoneticText;  // Returns or sets the text for the specified object.
@property BOOL visible;  // Returns or sets if the object is visible


@end

// Used as the base class for the PivotDataAxis, PivotFilterAxis, and PivotGroupAxis objects. There are no properties or methods which with you can use the PivotAxis object directly.
@interface ExcelPivotAxis : ExcelBaseObject

- (SBElementArray<ExcelPivotLine *> *) pivotLines;


@end

// Represents the memory cache for a PivotTable report.
@interface ExcelPivotCache : ExcelBaseObject

@property (copy, readonly) id ADOConnection;  // Returns an ADO connection object if the PivotTable cache is connected to an OLE DB data source.
@property (readonly) BOOL OLAP;  // Returns True if the PivotTable cache is connected to an Online Analytical Processing server.
@property (copy) NSString *SQLQuery;  // Returns or sets the SQL query string used with the specified ODBC data source.
@property BOOL backgroundQuery;  // Returns or sets if queries for the pivot table report or query table are performed asynchronously, in the background.
@property (copy) NSString *commandText;  // Returns or sets the command string for the specified data source.
@property ExcelE251 commandType;  // Returns or sets a constant describing the value type of the CommandText property.
@property (copy) NSString *connection;  // Returns or sets a string that contains one of the following. ODBC settings that enable Microsoft Excel to connect to an ODBC data source, a URL that enables Microsoft Excel to connect to a Web data source, or a file that specifies a database or Web query.
@property BOOL enableRefresh;  // Returns or sets if the pivot table cache can be refreshed by the user.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (readonly) BOOL isConnected;  // Returns True if the MaintainConnection property is True and the PivotTable cache is currently connected to its source.
@property (copy) NSString *localConnection;  // Returns or sets the connection string to an offline cube file.
@property BOOL maintainConnection;  // True if the connection to the specified data source is maintained after the refresh and until the workbook is closed.
@property (readonly) NSInteger memoryUsed;  // Returns the amount of memory currently being used by the object, in bytes.
@property ExcelE907 missingItemsLimit;  // Returns or sets the maximum quantity of unique items per PivotTable field that are retained even when they have no supporting data in the cache records.
@property BOOL optimizeCache;  // Returns or set if the pivot table cache is optimized when it's constructed.
@property (readonly) ExcelE250 queryType;  // Indicates the type of query used by Microsoft Excel to populate the PivotTable cache.
@property (readonly) NSInteger recordCount;  // Returns the number of records in the pivot table cache or the number of cache records that contain the specified item.
@property (copy, readonly) NSDate *refreshDate;  // Returns the date on which the pivot cache was last refreshed.
@property (copy, readonly) NSString *refreshName;  // Returns the name of the person who last refreshed the pivot cache. 
@property BOOL refreshOnFileOpen;  // Returns or sets if the pivot table cache or query table is automatically updated each time the workbook is opened.
@property NSInteger refreshPeriod;  // Returns or sets the number of minutes between refreshes.
@property ExcelXRbC robustConnect;  // Returns or sets how the PivotTable cache connects to its data source.
@property BOOL savePassword;  // Returns or set if password information in an ODBC connection string is saved with the specified query. if false, the password is removed.
@property (copy) NSString *sourceConnectionFile;  // Returns or sets a String indicating the Microsoft Office Data Connection file or similar file that was used to create the PivotTable.
@property ExcelE279 sourceData;  // Returns or sets the data source for the pivot table.
@property (readonly) ExcelE157 sourceType;  // Returns a value that identifies the type of item being published.
@property BOOL upgradeOnRefresh;  // Contains information on whether to upgrade the PivotCache and all connected PivotTables on the next refresh.
@property BOOL useLocalConnection;  // True if the LocalConnection property is used to specify the string that enables Microsoft Excel to connect to a data source.
@property (readonly) ExcelE900 version;  // Returns the version of Microsoft Excel in which the PivotCache was created.
@property (copy, readonly) id workbookConnection;  // Establishes a connection between the current workbook and the PivotCache object.

- (void) createPivotTableTableDestination:(NSString *)tableDestination tableName:(NSString *)tableName readData:(NSString *)readData defaultVersion:(NSString *)defaultVersion;  // Creates a PivotTable report based on a PivotCache object.
- (void) refresh;  // Updates the pivot table cache or query table.

@end

// Represents a cell in a PivotTable report.
@interface ExcelPivotCell : ExcelBaseObject

@property (copy, readonly) NSString *MDX;
@property (readonly) ExcelE909 cellChanged;
@property (readonly) ExcelE150 customSubtotalFunction;  // Returns the custom subtotal function field setting of a PivotCell object.
@property (copy, readonly) ExcelPivotField *dataField;  // Returns a PivotField object that corresponds to the selected data field.
@property (copy, readonly) id dataSourceValue;
@property (readonly) ExcelE908 pivotCellType;  // Returns a constant that identifies the PivotTable entity the cell corresponds to.
@property (copy, readonly) ExcelPivotLine *pivotColumnLine;  // Returns the PivotLine on a column for a specific PivotCell object.
@property (copy, readonly) ExcelPivotField *pivotField;  // Returns a PivotField object that represents the PivotTable field containing the upper-left corner of the specified range.
@property (copy, readonly) ExcelPivotItem *pivotItem;  // Returns a PivotItem object that represents the PivotTable item containing the upper-left corner of the specified range.
@property (copy, readonly) ExcelPivotLine *pivotRowLine;  // Returns the PivotLine on a row for a specific PivotCell object.
@property (copy, readonly) ExcelPivotTable *pivotTable;  // Returns a PivotTable object that represents the PivotTable report associated with the PivotCell.
@property (copy, readonly) ExcelRange *range;  // Returns a Range object that represents the range the specified PivotCell applies to.
@property (copy) id rowItems;  // Returns a PivotItemList collection that corresponds to the items on the category axis that represent the selected cell.

- (void) allocateChange;
- (void) discardChange;

@end

// Represents a field in a pivot table.
@interface ExcelPivotField : ExcelBaseObject

- (SBElementArray<ExcelChildItem *> *) childItems;
- (SBElementArray<ExcelHiddenItem *> *) hiddenItems;
- (SBElementArray<ExcelParentItem *> *) parentItems;
- (SBElementArray<ExcelPivotItem *> *) pivotItems;
- (SBElementArray<ExcelCalculatedItem *> *) calculatedItems;
- (SBElementArray<ExcelPivotFilter *> *) pivotFilters;

@property (readonly) BOOL allItemsVisible;  // Used to retrieve a Boolean value that indicates whether or not any manual filtering is applied to the PivotField.
@property (readonly) NSInteger autoShowCount;  // Returns the number of top or bottom items that are automatically shown in the pivot field.
@property (copy, readonly) NSString *autoShowField;  // Returns the name of the data field used to determine the top or bottom items that are automatically shown in the pivot field.
@property (readonly) ExcelE270 autoShowRange;  // Returns position top if the top items are shown automatically in the pivot field, returns position bottom if the bottom items are shown.
@property (readonly) ExcelE269 autoShowType;  // Returns type_automatic if auto show is enabled for the pivot field, or  returns type_manual if auto show is disabled.
@property (readonly) NSInteger autoSortCustomSubtotal;  // Returns the name of the custom subtotal used to sort the specified PivotTable field automatically.
@property (copy, readonly) NSString *autoSortField;  // Returns the name of the data field used to sort the pivot field automatically.
@property (readonly) ExcelE228 autoSortOrder;  // Returns the order used to sort the pivot field automatically. 
@property (copy, readonly) ExcelPivotLine *autoSortPivotLine;  // Returns the name of the PivotLine used to sort the specified PivotTable field automatically.
@property (copy) NSString *baseField;  // Returns or sets the base field for a custom calculation.
@property (copy) NSString *baseItem;  // Returns or sets the item in the base field for a custom calculation.
@property ExcelE209 calculation;  // Returns or sets the type of calculation done by the specified pivot field.
@property (copy) NSString *caption;  // The label text for the pivot field.
@property (copy, readonly) ExcelPivotField *childField;  // Returns a pivot field object that represents the child pivot field for the specified field, if the field is grouped and has a child field.
@property (copy, readonly) ExcelCubeField *cubeField;  // Returns the CubeField object from which the specified PivotTable field is descended.
@property (copy) NSString *currentPage;  // Returns or sets the current page showing for the page field, only valid for page fields.
@property (copy) id currentPageList;  // Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report.
@property (copy) NSString *currentPageName;  // Returns or sets the currently displayed page of the specified PivotTable report.
@property (copy, readonly) ExcelRange *dataRange;  // Returns a range object.  For a data field the range is the data contained in the field, for a row, column or page field it is the items in the field.
@property BOOL databaseSort;  // When set to True, manual repositioning of items in a PivotTable field is allowed.
@property (readonly) BOOL displayAsCaption;  // This property is used to display member properties of PivotFields as captions.
@property BOOL displayAsTooltip;  // This property is used to specify whether or not a specific member property PivotField is displayed in tooltips.
@property BOOL displayInReport;  // This property is used to specify whether the specified member property PivotField is displayed in the PivotTable or not.
@property BOOL dragToColumn;  // Returns or sets if the pivot field can be dragged to the column position.
@property BOOL dragToData;  // Returns or sets if the field can be dragged to the data position.
@property BOOL dragToHide;  // Returns or sets if the field can be hidden by being dragged off the pivot table.
@property BOOL dragToPage;  // Returns or sets if the field can be dragged to the page position.
@property BOOL dragToRow;  // Returns or sets if the field can be dragged to the row position.
@property BOOL drilledDown;  // True if the flag for the specified PivotTable field or PivotTable item is set to drilled.
@property BOOL enableItemSelection;  // When set to False, disables the ability to use the field dropdown in the user interface.
@property BOOL enableMultiplePageItems;  // Used for specifying whether or not check boxes are present in the filter drop-down list for fields in the page area.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property ExcelE150 function;  // Returns or sets the function used to summarize the pivot field.
@property (readonly) NSInteger groupLevel;  // Returns the placement of the specified field within a group of fields, if the field is a member of a grouped set of fields.
@property BOOL hidden;  // This property is used to hide the individual levels of an OLAP hierarchy.
@property (copy) id hiddenItemsList;  // Returns or sets an Object specifying an array of strings that are hidden items for a PivotTable field.
@property BOOL includeNewItemsInFilter;  // Allows developers to specify whether excluded or included items should be tracked when manual filtering is applied to the PivotField.
@property (readonly) BOOL isCalculated;  // Returns true if the pivot field or item is a calculated field or item.
@property (readonly) BOOL isMemberProperty;  // Returns True when the PivotField contains member properties.
@property (copy, readonly) ExcelRange *labelRange;  // Returns a range object that represents the cell, or cells, that contain the field label.
@property BOOL layoutBlankLine;  // True if a blank row is inserted after the specified row field in a PivotTable report.
@property BOOL layoutCompactRow;  // Specifies whether or not a PivotField is compacted , items of multiple PivotFields are displayed in a single column, when rows are selected.
@property ExcelE910 layoutForm;  // Returns or sets the way the specified PivotTable items appear.
@property BOOL layoutPagebreak;  // True if a page break is inserted after each field.
@property ExcelE902 layoutSubtotalLocation;  // Returns or sets the position of the PivotTable field subtotals in relation to either above or below the specified field.
@property (copy) NSString *memberPropertyCaption;  // Setting the MemberPropertyCaption property controls which member property is used as caption for a given level.
@property (readonly) NSInteger memoryUsed;  // Returns the amount of memory currently being used by the object, in bytes.
@property (copy) NSString *name;  // Returns or sets the name of the pivot field.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the pivot field. 
@property (copy, readonly) ExcelPivotField *parentField;  // Returns a pivot field object that represents the pivot field that's the group parent of the object. The field must be grouped and have a parent field.
@property (readonly) ExcelE155 pivotFieldDataType;  // Returns an enumeration describing the type of data in the pivot field.
@property ExcelE208 pivotFieldOrientation;  // The location of the field in the pivot table.
@property NSInteger position;  // Position of the field, first, second, third, and so on, among all the fields in its orientation, rows, columns, pages, data.
@property NSInteger propertyOrder;  // Valid only for PivotTable fields that are member property fields.
@property (copy, readonly) ExcelPivotField *propertyParentField;  // Returns a PivotField object representing the field to which the properties in this field pertain.
@property BOOL repeatLabels;
@property BOOL serverBased;  // Returns or sets if the pivot table's data source is external and only the items matching the page field selection are retrieved.
@property BOOL showAllItems;  // Returns or sets if all items in the pivot table are displayed, even if they don't contain summary data.
@property BOOL showDetail;  // Gets or sets whether the specified PivotField is showing detail.
@property (readonly) BOOL showingInAxis;  // Indicates if the PivotField is currently visible in the PivotTable or not.
@property (copy, readonly) NSString *sourceCaption;  // The SourceCaption property is applicable only for OLAP PivotTables, and returns the original caption from the OLAP server for a PivotField.
@property (copy, readonly) NSString *sourceName;  // Return the specified object's name, as it appears in the original source data for the pivot table. This might be different from the current item name if the user renamed the item after creating the pivot table.
@property (copy) NSString *standardFormula;  // Returns or sets a String specifying formulas with standard English formatting.
@property (copy) NSString *subtotalName;  // Returns or sets the text string label displayed in the subtotal column or row heading in the specified PivotTable report.
@property (readonly) NSInteger totalLevels;  // Returns the total number of fields in the current field group. If the field isn't grouped, or if the data source is OLAP-based, TotalLevels returns the value 1.
@property BOOL useMemberPropertyAsCaption;  // This property is used to control whether member property captions are used for PivotItem captions of the PivotField.
@property (copy) NSString *value;  // Returns or sets the name of the specified field in the pivot table.
@property (copy, readonly) NSArray *visibleItems;  // Returns a list of all the visible pivot items in the specified field.
@property (copy) id visibleItemsList;  // Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report.

- (void) addPageItemItem:(NSString *)item clearList:(NSString *)clearList;  // Adds an additional item to a multiple item page field.
- (void) autoShowType:(ExcelE269)type range:(ExcelE270)range count:(NSInteger)count field:(NSString *)field;  // Displays the number of top or bottom items for a row, page, or column field in the specified PivotTable report
- (void) autoSortSortOrder:(ExcelE228)sortOrder sortField:(NSString *)sortField;  // Establishes automatic pivot table field-sorting rules.
- (void) clearLabelFilters;  // This method deletes all label filters or all date filters in the PivotFilters collection of the PivotField.
- (void) clearValueFilters;  // Calling this method deletes all value filters in the PivotFilters collection of the PivotField.
- (BOOL) getSubtotalsSubtotalIndex:(ExcelE261)subtotalIndex;  // Returns subtotals displayed with the specified field. Valid only for nondata fields.
- (void) setSubtotalsSubtotalIndex:(ExcelE268)subtotalIndex value:(BOOL)value;  // Sets subtotals displayed with the specified field. Valid only for nondata fields.

@end

@interface ExcelCalculatedField : ExcelPivotField


@end

@interface ExcelColumnField : ExcelPivotField


@end

@interface ExcelDataField : ExcelPivotField


@end

@interface ExcelHiddenField : ExcelPivotField


@end

@interface ExcelPageField : ExcelPivotField


@end

// PivotTable Advanced Filter.
@interface ExcelPivotFilter : ExcelBaseObject

@property (readonly) BOOL active;  // Returns whether the specified PivotFilter is active.
@property (copy, readonly) ExcelCubeField *dataCubeField;  // This property is applicable only to non-OLAP PivotTables and provides the Value field ,PivotField in the Values area, being filtered by for a value filter.
@property (copy, readonly) ExcelPivotField *dataField;  // This property is applicable only to non-OLAP PivotTables and provides the Value field ,PivotField in the Values area, being filtered by for a value filter.
@property (copy, readonly) NSString *objectDescription;  // Provides an optional description for the PivotFilter object.
@property (readonly) ExcelE911 filterType;  // Specifies the type of filter to be applied.
@property (readonly) BOOL isMemberPropertyFilter;  // Specifies whether the label filter is based on the PivotItem captions of a member property of the field or on the PivotItem captions of the PivotField itself.
@property (copy, readonly) ExcelPivotField *memberPropertyField;  // This property specifies the member property PivotField on which the label filter is based.
@property (copy, readonly) NSString *name;  // This property provides the option of naming filters for reference.
@property NSInteger order;  // Specifies the evaluation order of the filter among all Value filters applied to the entire PivotTable.
@property (copy, readonly) ExcelPivotField *pivotField;  // Specifies the PivotField to which the filter is applied.
@property (copy, readonly) id value1;  // This property is a user-supplied parameter to define a filter for a PivotField.
@property (copy, readonly) id value2;  // This property is a user-supplied parameter to define a filter for a PivotField.


@end

@interface ExcelActiveFilter : ExcelPivotFilter


@end

// Represents a formula used to calculate results in a pivot table report.
@interface ExcelPivotFormula : ExcelBaseObject

@property NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property (copy) NSString *standardFormula;  // Returns or sets a String specifying formulas with standard English formatting.
@property (copy) NSString *value;  // Returns or sets the name of the specified formula in the pivot table formula.


@end

// Represents an item in a pivot field. The items are the individual data entries in a field category.
@interface ExcelPivotItem : ExcelBaseObject

- (SBElementArray<ExcelChildItem *> *) childItems;

@property (copy) NSString *caption;  // The label text for the pivot item.
@property (copy, readonly) ExcelRange *dataRange;  // Returns a range object specifying the data qualified by the item.
@property BOOL drilledDown;  // True if the flag for the specified PivotTable field or PivotTable item is set to 
@property (copy) NSString *formula;  // Returns or sets the pivot item's formula, in A1-style notation and in the language of the macro.
@property (readonly) BOOL isCalculated;  // Returns true if the pivot  item is a calculated item.
@property (copy, readonly) ExcelRange *labelRange;  // Returns a range object that represents all the pivot table cells that contain the item.
@property (copy) NSString *name;  // Returns or sets the name of the pivot item.
@property (copy, readonly) ExcelParentItem *parentItem;  // Returns a pivot item object that represents the parent pivot item in the parent pivot field object, the field must be grouped so that it has a parent.
@property (readonly) BOOL parentShowDetail;  // Return true if the specified item is showing because one of its parents is showing detail. False if the specified item isn't showing because one of its parents is hiding detail. This property is available only if the item is grouped.
@property NSInteger position;  // Returns or sets the position of the item in its field, if the item is currently showing.
@property (readonly) NSInteger recordCount;  // Returns the number of records in the pivot table cache or the number of cache records that contain the specified item.
@property BOOL showDetail;  // Returns  or sets if the pivot item is showing detail.
@property (copy, readonly) NSString *sourceName;  // Return the specified object's name, as it appears in the original source data for the pivot table. This might be different from the current item name if the user renamed the item after creating the pivot table.
@property (copy, readonly) NSString *sourceNameStandard;  // Returns a String that represents the PivotTable item's source name in standard English format settings.
@property (copy) NSString *standardFormula;  // Returns or sets a String specifying formulas with standard English formatting.
@property (copy) NSString *value;  // Returns or sets set the name of the specified item in the pivot table field.
@property BOOL visible;  // Returns or sets if the pivot item is visible.


@end

@interface ExcelCalculatedItem : ExcelPivotItem


@end

@interface ExcelChildItem : ExcelPivotItem


@end

@interface ExcelColumnItem : ExcelPivotItem


@end

@interface ExcelHiddenItem : ExcelPivotItem


@end

@interface ExcelParentItem : ExcelPivotItem


@end

// The PivotLines object is a collection of lines in a PivotTable, containing all lines on rows or columns of the pivot.
@interface ExcelPivotLine : ExcelBaseObject

@property (readonly) ExcelE912 lineType;  // Returns an XlPivotLineType constant that indicates the type of PivotLine.
@property (copy, readonly) id pivotLineCells;  // Returns a collection of PivotCell objects in a PivotLine.
@property (readonly) NSInteger position;  // Returns or sets the position of the PivotLine object.


@end

// Represents a pivot table on a worksheet.
@interface ExcelPivotTable : ExcelBaseObject

- (SBElementArray<ExcelColumnField *> *) columnFields;
- (SBElementArray<ExcelDataField *> *) dataFields;
- (SBElementArray<ExcelHiddenField *> *) hiddenFields;
- (SBElementArray<ExcelPageField *> *) pageFields;
- (SBElementArray<ExcelPivotField *> *) pivotFields;
- (SBElementArray<ExcelRowField *> *) rowFields;
- (SBElementArray<ExcelCalculatedField *> *) calculatedFields;
- (SBElementArray<ExcelPivotFormula *> *) pivotFormulas;
- (SBElementArray<ExcelCalculatedMember *> *) calculatedMembers;
- (SBElementArray<ExcelActiveFilter *> *) activeFilters;
- (SBElementArray<ExcelCubeField *> *) cubeFields;
- (SBElementArray<ExcelSlicer *> *) slicers;

@property NSInteger CompactRowIndent;  // Returns or sets the indent increment for PivotItems when compact row layout form is turned on.
@property ExcelE903 allocation;
@property ExcelE905 allocationMethod;
@property ExcelE904 allocationValue;
@property (copy) NSString *allocationWeightExpression;
@property BOOL allowMultipleFilters;  // Sets or retrieves a value that indicates whether a PivotField can have multiple filters applied to it at the same time.
@property (copy) NSString *alternativeText;  // Returns or sets the descriptive alternative text string for a ShapeRange object when the object is saved to a Web page.
@property NSInteger cacheIndex;  // Returns or sets the index number of the pivot table cache.
@property BOOL calculatedMembersInFilters;
@property (copy, readonly) id changeList;
@property BOOL columnGrand;  // Returns or sets if the pivot table shows grand totals for columns.
@property (copy, readonly) ExcelRange *columnRange;  // Returns a range object that represents the range that contains the pivot table column area.
@property (copy) NSString *compactLayoutColumnHeader;  // Specifies the caption that is displayed in the column header of a PivotTable when in compact row layout form.
@property (copy) NSString *compactLayoutRowHeader;  // Specifies the caption that is displayed in the row header of a PivotTable when in compact row layout form.
@property (copy, readonly) ExcelRange *dataBodyRange;  // Returns a range object that represents the range that contains the pivot table data area.
@property (copy, readonly) ExcelRange *dataLabelRange;  // Returns a range object that represents the range that contains the labels for the pivot table data fields. 
@property (copy, readonly) ExcelPivotField *dataPivotField;  // Returns a PivotField object that represents all the data fields in a PivotTable.
@property BOOL displayContextTooltips;  // Controls whether or not tooltips are displayed for PivotTable cells.
@property BOOL displayEmptyColumn;  // Returns True when the nonempty MDX keyword is included in the query to the OLAP provider for the value axis.
@property BOOL displayEmptyRow;  // Returns True when the nonempty MDX keyword is included in the query to the OLAP provider for the category axis.
@property BOOL displayErrorString;  // Returns or sets if the pivot table displays a custom error string in cells that contain errors.
@property BOOL displayFieldCaptions;  // Controls whether or not filter buttons and PivotField captions for rows and columns are displayed in the grid.
@property BOOL displayImmediateItems;  // Returns or sets a Boolean that indicates whether items in the row and column areas are visible when the data area of the PivotTable is empty.
@property BOOL displayMemberPropertyTooltips;  // Controls whether or not to display member properties in tooltips.
@property BOOL displayNullString;  // Returns or sets if the pivot table displays a custom string in cells that contain null values.
@property BOOL enableDataValueEditing;  // True to disable the alert for when the user overwrites values in the data area of the PivotTable.
@property BOOL enableDrilldown;  // Returns or sets if drilldown is enabled.
@property BOOL enableFieldDialog;  // Returns or sets if the pivot table field dialog box is available when the user double-clicks the pivot table field.
@property BOOL enableFieldList;  // False to disable the ability to display the field well for the PivotTable.
@property BOOL enableWizard;  // Returns or sets if the pivot table wizard is available.
@property BOOL enableWriteback;
@property (copy) NSString *errorString;  //  Returns or sets the string displayed in cells that contain errors when the display error string property is true. The default value is an empty string.
@property BOOL fileListSortAscending;  // Controls the sort order of fields in the PivotTable Field List.
@property (copy) NSString *grandTotalName;  // Returns or sets the text string label that is displayed in the grand total column or row heading in the specified PivotTable report
@property BOOL hasAutoformat;  // Returns or sets if the pivot table is automatically formatted when it's refreshed or when fields are moved.
@property BOOL inGridDropZones;  // This property is used to toggle in-grid drop zones for a PivotTable object.
@property (copy) NSString *innerDetail;  // Returns or sets the name of the field that will be shown as detail when the show detail property is true for the innermost row or column field.
@property ExcelE901 layoutRowDefault;  // This property specifies the layout settings for PivotFields when they are added to the PivotTable for the first time.
@property (copy) NSString *location;  // Gets or sets a String that represents the top-left cell in the body of the specified PivotTable.
@property BOOL manualUpdate;  // Returns or sets if the pivot table is recalculated only at the user's request.
@property (copy, readonly) NSString *mdx;  // Returns a String indicating the MDX, Multidimensional-Expression, that would be sent to the provider to populate the current PivotTable view.
@property BOOL mergeLabels;  // Returns or sets if pivot table outer-row item, column item, subtotal, and grand total labels use merged cells.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *nullString;  // Returns or sets the string displayed in cells that contain null values when the display null string property is true. The default value is an empty string.
@property ExcelE164 pageFieldOrder;  // Returns or sets the order in which page fields are added to the pivot table layout.
@property (copy) NSString *pageFieldStyle;  // Returns or sets the style used in the bound page field area.
@property NSInteger pageFieldWrapCount;  // Returns or sets the number of pivot table page fields in each column or row.
@property (copy, readonly) ExcelRange *pageRange;  // Returns a range object that represents the range that contains the pivot table page area.
@property (copy, readonly) ExcelRange *pageRangeCells;  // Returns a range object that represents the cells in the pivot table containing only the page fields and item drop-down lists.
@property (copy, readonly) ExcelPivotCache *pivotCache;  // Returns a pivot cache object that represents the cache for the specified pivot table
@property (copy, readonly) ExcelPivotAxis *pivotColumnAxis;  // Returns a PivotAxis object representing the entire column axis 
@property (copy, readonly) ExcelPivotAxis *pivotRowAxis;  // Returns a PivotAxis object representing the entire row axis 
@property (copy) NSString *pivotSelection;  // Returns or sets the pivot table selection, in standard pivot table selection format.
@property (copy) NSString *pivotSelectionStandard;  // Returns or sets a String indicating the PivotTable selection in standard PivotTable report format using English settings.
@property BOOL preserveFormatting;  // Returns or sets if pivot table formatting is preserved when the pivot table is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.
@property BOOL printDrillIndicators;  // Specifies whether or not drill indicators are printed with the PivotTable.
@property BOOL printTitles;  // set/get whether the print titles for the worksheet are set based on the PivotTable report.
@property (copy, readonly) NSDate *refreshDate;  // Returns the date on which the pivot table was last refreshed.
@property (copy, readonly) NSString *refreshName;  // Returns the name of the person who last refreshed the pivot table data.
@property BOOL repeatItemsOnEachPrintedPage;  // True if row, column, and item labels appear on the first row of each page when the specified PivotTable report is printed. False if labels are printed only on the first page.
@property BOOL rowGrand;  // Returns or sets if the pivot table shows grand totals for rows.
@property (copy, readonly) ExcelRange *rowRange;  // Returns a range object that represents the range including the pivot table row area.
@property BOOL saveData;  // Returns or sets if data for the pivot table is saved with the workbook. If false only the pivot table definition is saved.
@property ExcelE214 selectionMode;  // Returns or sets the pivot table structured selection mode.
@property BOOL showDrillIndicators;  // When the ShowDrillIndicators property is set to False, the PrintDrillIndicators property has no effect.
@property BOOL showPageMultipleLabel;  // When set to True, Multiple Items will appear in the PivotTable cell on the worksheet whenever items are hidden and an aggregate of nonhidden items is shown in the PivotTable view.
@property BOOL showTableStyleColumnHeaders;  // The ShowTableStyleColumnHeaders property is set to True if the coulmn headers should be displayed in the PivotTable.
@property BOOL showTableStyleColumnStripes;  // The ShowTableStyleColumnStripes property displays banded columns in which even columns are formatted differently from odd columns.
@property BOOL showTableStyleLastColumn;  // Returns or sets if the last column is displayed for the specified ListObject object.
@property BOOL showTableStyleRowHeaders;  // The ShowTableStyleRowHeaders property is set to True if the row headers should be displayed in the PivotTable.
@property BOOL showTableStyleRowStripes;  // The ShowTableStyleRowStripes property displays banded rows in which even rows are formatted differently from odd rows.
@property BOOL showValuesRow;
@property BOOL smallGrid;  // Returns or sets if Microsoft Excel uses a grid that's two cells wide and two cells deep for a newly created pivot table report. False if Excel uses a blank stencil outline.
@property BOOL sortUsingCustomLists;  // The SortUsingCustomLists property controls whether custom lists are used for sorting items of fields.
@property ExcelE279 sourceData;  // Either a range reference as a string or a list of SQL commands
@property BOOL subtotalHiddenPageItems;  // Returns or sets if hidden page field items in the pivot table are included in row and column subtotals, block totals, and grand totals.
@property (copy) NSString *summary;
@property (copy, readonly) ExcelRange *tableRange1;  // Returns a range object that represents the range containing the entire pivot table, but doesn't include page fields.
@property (copy, readonly) ExcelRange *tableRange2;  // Returns a range object that represents the range containing the entire pivot table, including page fields.
@property (copy) NSString *tableStyle;  // Returns or sets the style used in the pivot table body.
@property (copy) id tableStyle2;  // The TableStyle2 property specifies the PivotTable style currently applied to the PivotTable.
@property (copy) NSString *tag;  // Returns or sets a string saved with the pivot table.
@property BOOL totalsAnnotation;  // True if an asterisk is displayed next to each subtotal and grand total value in the specified PivotTable report if the report is based on an OLAP data source.
@property (copy) NSString *vacatedStyle;  // Returns or sets the style applied to cells vacated when the pivot table is refreshed.
@property (copy) NSString *value;  // Returns or set the name of the pivot table.
@property (readonly) ExcelE900 version;
@property BOOL viewCalculatedMembers;  // When set to True, calculated members for Online Analytical Processing, OLAP, PivotTables can be viewed.
@property BOOL visualTotals;  // True to enable Online Analytical Processing, OLAP, PivotTables to retotal after an item has been hidden from view.
@property BOOL visualTotalsForSets;

- (ExcelPivotField *) addDataFieldField:(id)field caption:(NSString *)caption function:(NSString *)function;  // Adds a data field to a PivotTable report.
- (void) addFieldsToPivotTableRowFields:(NSArray *)rowFields columnFields:(NSArray *)columnFields pageFields:(NSArray *)pageFields addToTable:(BOOL)addToTable;  // Adds row, column, and page fields to a pivot table.
- (void) allocateChanges;
- (void) changeConnectionConnection:(ExcelWorkbookConnection *)connection;  // Changes the connection of the specified PivotTable.
- (void) changePivotCachePivotCache:(NSString *)pivotCache;  // Changes the PivotCache of the specified PivotTable.
- (void) clearTable;  // The ClearTable method is used for clearing a PivotTable.
- (void) commitChanges;
- (void) convertToFormulasConvertFilters:(BOOL)convertFilters;  // The ConvertToFormulas method is new in 1st_Excel12 and is used for converting a PivotTable to cube formulas.
- (NSString *) createCubeFileFile:(NSString *)file measures:(NSString *)measures levels:(NSString *)levels members:(NSString *)members properties:(NSString *)properties;  // Creates a cube file from a PivotTable report connected to an Online Analytical Processing data source.
- (void) discardChanges;
- (ExcelRange *) getPivotDataDataField:(NSString *)dataField field1:(NSString *)field1 item1:(NSString *)item1 field2:(NSString *)field2 item2:(NSString *)item2 field3:(NSString *)field3 item3:(NSString *)item3 field4:(NSString *)field4 item4:(NSString *)item4 field5:(NSString *)field5 item5:(NSString *)item5 field6:(NSString *)field6 item6:(NSString *)item6 field7:(NSString *)field7 item7:(NSString *)item7 field8:(NSString *)field8 item8:(NSString *)item8 field9:(NSString *)field9 item9:(NSString *)item9 field10:(NSString *)field10 item10:(NSString *)item10 field11:(NSString *)field11 item11:(NSString *)item11 field12:(NSString *)field12 item12:(NSString *)item12 field13:(NSString *)field13 item13:(NSString *)item13 field14:(NSString *)field14 item14:(NSString *)item14;  // Returns a Range object with information about a data item in a PivotTable report.
- (double) getPivotTableDataName:(NSString *)name;  // Retrieve data from a pivot table
- (NSArray *) getVisibleFields;  // Returns a list of all the visible pivot fields. Visible pivot fields are shown as row, column, page or data fields.
- (void) listFormulas;  // Creates a list of calculated pivot table items and fields on a separate worksheet.
- (void) makeConnection;  // Establishes a connection for the specified PivotTable cache.
- (void) pivotSelectName:(NSString *)name mode:(ExcelE214)mode;  // Selects part of a pivot table.
- (void) refreshDataSourceValues;
- (BOOL) refreshTable;  // Refreshes the pivot table from the source data. Returns true if it's successful.
- (void) repeatAllLabelsRepeat:(ExcelE906)repeat;
- (void) rowAxisLayoutLayout:(ExcelE901)layout;  // This method is used for simultaneously setting layout options for all existing PivotFields.
- (void) saveAsODCODCFileName:(NSString *)ODCFileName objectDescription:(NSString *)objectDescription keywords:(NSString *)keywords;  // Saves the PivotTable cache source as a Microsoft Office Data Connection file.
- (void) showPagesPageField:(NSString *)pageField;  // Creates a new pivot table for each item in the page field. Each new pivot table is created on a new worksheet.
- (void) subtotalLocationLocation:(ExcelE902)location;  // This method changes the subtotal location for all existing PivotFields.
- (void) update;  // Updates the link or pivot table.

@end

// Represents a worksheet table built from data returned from an external data source, such as an SQL server.
@interface ExcelQueryTable : ExcelBaseObject

@property (copy) NSArray *FileMakerFields;
@property NSInteger FileMakerNumCriteria;
@property BOOL FileMakerUseTable;  // Returns or sets if the query uses a table rather than a layout. Only applicable for FileMaker 7 or above
@property BOOL adjustColumnWidth;  // Returns or sets if the column widths are automatically adjusted for the best fit each time you refresh the specified query table.
@property BOOL backgroundQuery;  // Returns or sets if queries for the query table are performed asynchronously.
@property ExcelE251 commandType;  // Returns or sets one of the XlCmdType constants listed in the following table in the remarks section.
@property (copy) NSString *connection;  // Returns or sets a string that contains one of the following: ODBC settings that enable Microsoft Excel to connect to an ODBC data source; a URL that enables Microsoft Excel to connect to a Web data source; or a file that specifies a database or Web query.
@property (copy, readonly) ExcelRange *destination;  // Returns the cell in the upper-left corner of the query table destination range, the range where the resulting query table will be placed. The destination range must be on the worksheet that contains the query table object.
@property BOOL enableEditing;  // Returns or sets if the user can edit the specified query table. False if the user can only refresh the query table.
@property BOOL enableRefresh;  // Returns or sets if the query table can be refreshed by the user.
@property (readonly) BOOL fetchedRowOverflow;  // Returns true if the number of rows returned by the last use of the refresh method is greater than the number of rows available on the worksheet.
@property BOOL fieldNames;  // Returns or sets if field names from the data source appear as column headings for the returned data.
@property BOOL fillAdjacentFormulas;  // Returns or sets if formulas to the right of the specified query table are automatically updated whenever the query table is refreshed.
@property BOOL hasAutoformat;  // Returns or sets if the query table is automatically formatted when it's refreshed or when fields are moved.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *postText;  // Returns or sets the string used with the post method of inputting data into a Web server to return data from a Web query.
@property (readonly) ExcelE250 queryType;  // Returns the type of query used by Microsoft Excel to populate the query table.
@property BOOL refreshOnFileOpen;  // Returns or sets if the query table is automatically updated each time the workbook is opened.
@property ExcelE191 refreshStyle;  // Returns or sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a recordset returned by a query.
@property (readonly) BOOL refreshing;  // Returns true if there's a background query in progress for the specified query table.
@property (copy, readonly) ExcelRange *resultRange;  // Returns a range object that represents the area of the worksheet occupied by the specified query table.
@property BOOL rowNumbers;  // Returns or sets if row numbers are added as the first column of the specified query table.
@property BOOL saveData;  // Returns or sets if data for the query table is saved with the workbook.
@property BOOL savePassword;  // Returns or sets if password information in an ODBC connection string is saved with the specified query.
@property (copy) NSString *sql;  // Returns or sets the SQL query string used with the specified ODBC data source.
@property BOOL tablesOnlyFromHtml;  // Returns or sets if only the HTML tables in the document are read when a query table is refreshed. False if the entire HTML document is read when a query table is refreshed. This property has an effect only when the query table is using a URL connection.
@property (copy) NSArray *textFileColumnDataTypes;  // A list of enums: general format, text format, MDY format, DMY format, YMD format, MYD format, DYM format, YDM format, skip column
@property BOOL textFileCommaDelimiter;  // Returns or sets if the comma character is the delimiter when you import a text file into a query table.
@property BOOL textFileConsecutiveDelimiter;  // Returns or sets if consecutive delimiters are treated as a single delimiter when you import a text file into a query table.
@property (copy) NSString *textFileDecimalSeparator;  // Returns or sets the decimal separator character that Microsoft Excel uses when you import a text file into a query table. The default is the system decimal separator character.
@property (copy) NSArray *textFileFixedColumnWidths;  // Returns or sets a list of numbers that correspond to the widths of the columns in the text file that you are importing into a query table. Valid widths are from 1 through 32767 characters.
@property (copy) NSString *textFileOtherDelimiter;  // Returns or sets the character used as the delimiter when you import a text file into a query table.
@property ExcelE236 textFileParseType;  // Returns or sets the column format for the data in the text file that you're importing into a query table.
@property ExcelE211 textFilePlatform;  // Returns or sets the origin of the text file you're importing into the query table. This property determines which code page is used during the data import.
@property BOOL textFilePromptOnRefresh;  // Returns or sets if you want to specify the name of the imported text file each time the query table is refreshed.
@property BOOL textFileSemicolonDelimiter;  // Returns or sets if the semicolon character is the delimiter when you import a text file into a query table.
@property BOOL textFileSpaceDelimiter;  // Returns or sets if the space character is the delimiter when you import a text file into a query table.
@property NSInteger textFileStartRow;  // Returns or sets the row number at which text parsing will begin when you import a text file into a query table. Valid values are integers from 1 through 32767. The default value is 1.
@property BOOL textFileTabDelimiter;  // Returns or sets if the tab character is the delimiter when you import a text file into a query table.
@property ExcelE237 textFileTextQualifier;  // Returns or sets the text qualifier when you import a text file into a query table. The text qualifier specifies that the enclosed data is in text format.
@property (copy) NSString *textFileThousandsSeparator;  // Returns or sets the thousands separator character thatMicrosoft Excel uses when you import a text file into a query table. The default is the system thousands separator character.
@property BOOL useListObject;

- (void) cancelRefresh;  // Cancels all background queries for the specified query table. Use the refreshing property to determine whether a background query is currently in progress.
- (NSDictionary *) getFileMakerCriteriaCriteriaIndex:(NSInteger)criteriaIndex;  // Get the criteria from an existing Excel query that was created against a FileMaker database
- (BOOL) refreshQueryTableBackgroundQuery:(BOOL)backgroundQuery;  // Updates the PivotTable cache or query table.
- (void) setFileMakerCriteriaCriteriaIndex:(NSInteger)criteriaIndex fieldName:(NSString *)fieldName operator:(ExcelE257)operator_ clauseText:(NSString *)clauseText condition:(ExcelE258)condition;

@end

// Represents a file in the list of recently used files.
@interface ExcelRecentFile : ExcelBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy, readonly) NSString *path;  // Returns the complete path to the file, excluding the final separator and name of the file.


@end

// The RectangularGradient object transitions through a series of colors from the center to one or more sides.
@interface ExcelRectangularGradient : ExcelBaseObject

@property (copy, readonly) ExcelColorstops *colorstops;  // Returns the ColorStops for the LinearGradient object.
@property double rectangularGradientBottom;  // Represents the point or vector that the gradient fill converges to.
@property double rectangularGradientLeft;  // Represents the point or vector that the gradient fill converges to.
@property double rectangularGradientRight;  // Represents the point or vector that the gradient fill converges to.
@property double rectangularGradientTop;  // Represents the point or vector that the gradient fill converges to.


@end

@interface ExcelRowField : ExcelPivotField


@end

@interface ExcelRowItem : ExcelPivotItem


@end

// Represents a scenario on a worksheet. Represents a scenario on a worksheet. A scenario is a group of input values, called changing cells, that's named and saved.
@interface ExcelScenario : ExcelBaseObject

@property (copy) NSString *ExcelComment;  // Returns or sets the comment associated with the scenario. The comment text cannot exceed 255 characters.
@property (copy, readonly) ExcelRange *changingCells;  // Returns a range object that represents the changing cells for a scenario.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property BOOL hidden;  //  Returns or sets if the scenario is hidden.
@property BOOL locked;  // Returns or sets if the object is locked.
@property (copy) NSString *name;  // Returns or sets the name of the object.

- (void) changeScenarioChangingCells:(ExcelE267)changingCells values:(NSArray *)values;  // Changes the scenario to have a new set of changing cells and, optionally, scenario values.
- (NSArray *) getValues;  // Returns or sets a list that contains the current values of the changing cells for the scenario.

@end

@interface ExcelScrollbar : ExcelBaseObject

@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property BOOL displayThreeDShading;
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or set the height of the object.
@property NSInteger largeChange;
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property (copy) NSString *linkedCell;
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property NSInteger maximumValue;
@property NSInteger minimumValue;
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property NSInteger smallChange;
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property NSInteger value;
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents a worksheet.
@interface ExcelSheet : ExcelBaseObject

- (SBElementArray<ExcelShape *> *) shapes;
- (SBElementArray<ExcelArc *> *) arcs;
- (SBElementArray<ExcelButton *> *) buttons;
- (SBElementArray<ExcelChartObject *> *) chartObjects;
- (SBElementArray<ExcelCheckbox *> *) checkboxes;
- (SBElementArray<ExcelDropdown *> *) dropdowns;
- (SBElementArray<ExcelGroupbox *> *) groupboxes;
- (SBElementArray<ExcelLabel *> *) labels;
- (SBElementArray<ExcelLine *> *) lines;
- (SBElementArray<ExcelListbox *> *) listboxes;
- (SBElementArray<ExcelNamedItem *> *) namedItems;
- (SBElementArray<ExcelOptionButton *> *) optionButtons;
- (SBElementArray<ExcelOval *> *) ovals;
- (SBElementArray<ExcelPivotTable *> *) pivotTables;
- (SBElementArray<ExcelRange *> *) ranges;
- (SBElementArray<ExcelCell *> *) cells;
- (SBElementArray<ExcelRow *> *) rows;
- (SBElementArray<ExcelColumn *> *) columns;
- (SBElementArray<ExcelRectangle *> *) rectangles;
- (SBElementArray<ExcelScenario *> *) scenarios;
- (SBElementArray<ExcelScrollbar *> *) scrollbars;
- (SBElementArray<ExcelSpinner *> *) spinners;
- (SBElementArray<ExcelTextbox *> *) textboxes;
- (SBElementArray<ExcelHorizontalPageBreak *> *) horizontalPageBreaks;
- (SBElementArray<ExcelVerticalPageBreak *> *) verticalPageBreaks;
- (SBElementArray<ExcelQueryTable *> *) queryTables;
- (SBElementArray<ExcelExcelComment *> *) ExcelComments;
- (SBElementArray<ExcelHyperlink *> *) hyperlinks;
- (SBElementArray<ExcelListObject *> *) listObjects;

@property BOOL autofilterMode;  // Returns or sets if the autofilter drop-down arrows are currently displayed on the sheet. This property is independent of the filter mode property. 
@property (copy, readonly) ExcelAutofilter *autofilterObject;  // Returns the autofilter object associated with this sheet.
@property (copy, readonly) ExcelRange *circularReference;  // Returns a range object that represents the range containing the first circular reference on the sheet, or returns missing value if there's no circular reference on the sheet.
@property (readonly) ExcelE150 consolidationFunction;  // Returns the function code used for the current consolidation.
@property (copy, readonly) NSArray *consolidationOptions;  // Returns a three-element list of boolean values. If the element is true, that option is set. Element 1 is use labels in top row, element 2 is use labels in left column, and element 3 is create links to source data.
@property (copy, readonly) NSArray *consolidationSources;  // Returns an list of string values that name the source sheets for the worksheet's current consolidation.
@property BOOL displayPageBreaks;  // Returns or sets if page breaks, both automatic and manual, on the specified worksheet are displayed.
@property BOOL enableAutofilter;  // Returns or sets if autofilter arrows are enabled when user-interface-only protection is turned on.
@property BOOL enableCalculation;  // Returns or sets if Microsoft Excel automatically recalculates the worksheet when necessary. If false, the user cannot request a recalculation, Microsoft Excel never recalculates the sheet automatically.
@property BOOL enableOutlining;  // Returns or sets if outlining symbols are enabled when user-interface-only protection is turned on.
@property BOOL enablePivotTable;  // Returns or sets if pivot table controls and actions are enabled when user-interface-only protection is turned on.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (readonly) BOOL filterMode;  // Returns true if the worksheet is in filter mode.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) ExcelSheet *next;  // Returns a worksheet object that represents the next sheet.
@property (copy, readonly) ExcelOutline *outlineObject;  // Returns an outline object that represents the outline for the specified worksheet.
@property (copy, readonly) ExcelPageSetup *pageSetupObject;  // Returns the page setup object associated with this chart.
@property (copy, readonly) ExcelSheet *previous;  // Returns a worksheet object that represents the previous sheet.
@property (readonly) BOOL protectContents;  // Returns true if the contents of the sheet are protected.
@property (readonly) BOOL protectDrawingObjects;  // Returns true if shapes are protected.
@property (readonly) BOOL protectionMode;  // Returns true if user-interface-only protection is turned on. To turn on user interface protection, use the protect method with the user interface only argument set to true.
@property (copy, readonly) ExcelProtection *protectionObject;  // Returns a Protection object that represents the protection options of the worksheet.
@property (copy) NSString *scrollArea;  // Returns or sets the range where scrolling is allowed, as an A1-style range reference. Cells outside the scroll area cannot be selected. 
@property (copy, readonly) ExcelTab *sheetTab;  // Returns the sheet tab of the work sheet
@property (copy, readonly) ExcelSort *sortObject;  // Returns the sort object in the sheet.
@property (readonly) double standardHeight;  // Returns the standard default height of all the rows in the worksheet, in points. 
@property double standardWidth;  // Returns or sets the standard default width of all the columns in the worksheet.
@property BOOL transitionExpressionEvaluation;  // Returns or sets if Microsoft Excel uses Lotus 1-2-3 expression evaluation rules for the worksheet.
@property (copy, readonly) ExcelRange *usedRange;  // Returns a range object that represents the used range on the specified worksheet.
@property ExcelE225 visible;  // Returns or sets if the worksheet is visible.
@property (readonly) ExcelE151 worksheetType;  // Returns the type of this worksheet.

- (void) circleInvalid;  // Circles invalid entries on the worksheet.
- (void) clearArrows;  // Clears the tracer arrows from the worksheet. Tracer arrows are added by using the auditing feature.
- (void) clearCircles;  // Clears circles from invalid entries on the worksheet.
- (void) copyWorksheetBefore:(ExcelSheet *)before after:(ExcelSheet *)after NS_RETURNS_NOT_RETAINED;  // Copies the sheet to another location in the workbook.
- (void) pasteSpecialOnWorksheetFormat:(NSString *)format link:(BOOL)link displayAsIcon:(BOOL)displayAsIcon iconFileName:(NSString *)iconFileName iconIndex:(NSInteger)iconIndex iconLabel:(NSString *)iconLabel noHTMLFormatting:(BOOL)noHTMLFormatting;  // Pastes the contents of the clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.
- (void) pasteWorksheetDestination:(ExcelE267)destination link:(BOOL)link;  // Pastes the contents of the Clipboard onto the sheet.
- (void) protectWorksheetPassword:(NSString *)password drawingObjects:(BOOL)drawingObjects worksheetContents:(BOOL)worksheetContents scenarios:(BOOL)scenarios userInterfaceOnly:(BOOL)userInterfaceOnly allowFormattingCells:(BOOL)allowFormattingCells allowFormattingColumns:(BOOL)allowFormattingColumns allowFormattingRows:(BOOL)allowFormattingRows allowInsertingColumns:(BOOL)allowInsertingColumns allowInsertingRows:(BOOL)allowInsertingRows allowInsertingHyperlinks:(BOOL)allowInsertingHyperlinks allowDeletingColumns:(BOOL)allowDeletingColumns allowDeletingRows:(BOOL)allowDeletingRows allowSorting:(BOOL)allowSorting allowFiltering:(BOOL)allowFiltering allowUsingPivotTable:(BOOL)allowUsingPivotTable;  // Protects a worksheet so that it cannot be modified.
- (void) resetAllPageBreaks;  // Resets all page breaks on the specified worksheet.
- (void) saveAsFilename:(NSString *)filename fileFormat:(ExcelE188)fileFormat password:(NSString *)password writeReservationPassword:(NSString *)writeReservationPassword readOnlyRecommended:(BOOL)readOnlyRecommended createBackup:(BOOL)createBackup addToMostRecentlyUsedList:(BOOL)addToMostRecentlyUsedList overwrite:(BOOL)overwrite saveAsLocalLanguage:(BOOL)saveAsLocalLanguage;  // Saves changes into a different file.
- (void) setBackgroundPicturePictureFileName:(NSString *)pictureFileName;  // Sets the background graphic for a worksheet.
- (void) showAllData;  // Makes all rows of the currently filtered list visible. If autofilter is in use, this method changes the arrows to All.
- (void) showDataForm;  // Displays the data form associated with the worksheet.

@end

@interface ExcelInternationalMacroSheet : ExcelSheet

@property ExcelE197 enableSelection;  // Returns or sets what can be selected on the sheet.


@end

@interface ExcelMacroSheet : ExcelSheet

@property ExcelE197 enableSelection;  // Returns or sets what can be selected on the sheet.


@end

// A slicer is a mechanism for filtering data in PivotTable views and cube functions.
@interface ExcelSlicer : ExcelBaseObject


@end

// Represents a sort of a range of data.
@interface ExcelSort : ExcelBaseObject

@property BOOL matchCase;  // Set to true to perform a case-sensitive sort or set to false to perform non-case sensitive sort. Read/Write.
@property ExcelE241 sortHeader;  // Specifies whether the first row contains header information. Read/Write.
@property ExcelE226 sortMethod;  // Specifies the sort method for Chinese/Japanese languages. Read/Write.
@property ExcelE229 sortOrientation;  // Specifies the orientation for the sort. Read/Write.
@property (copy, readonly) id sortfields;  // Stores the sort state for workbooks, lists, and autofilters. Read-Only.
@property (copy, readonly) ExcelRange *sortrange;  // Returns a range object that represents the range to which the specified sort applies. Read Only.

- (void) setSortRangeRng:(ExcelRange *)rng;  // Sets the starting and ending character positions for Sort object.

@end

// The sortfield object contains all the sort information for the worksheet, lists, and autofilter object.
@interface ExcelSortfield : ExcelBaseObject

@property (copy) id customOrder;  // Specifies a custom order to sort the fields. Read/Write.
@property ExcelE302 sortDataOption;  // Specifies how to sort text in the range specified in sortfield object. Read/Write.
@property (copy, readonly) ExcelRange *sortKey;  // Specifies the range that is currently being sorted on. Read only.
@property ExcelE301 sortOn;  // Returns or sets what attribute of the cell to sort on. Read/Write.
@property (copy, readonly) id sortOnValues;  // Return the value on which the sort is performed for the specified sortfield object.Read Only.
@property ExcelE228 sortOrder;  // Determines the sort order for the values specified in the key. Read/Write.
@property NSInteger sortPriority;  // Specifies the priority for the sortfield. Read/Write.

- (void) modifySortKeyRng:(ExcelRange *)rng;  // Modify the key value by which values are sorted in the field.

@end

// The sortfields collection is a collection of sortfield objects. It allows developers to store a sort state on worksheets, lists, and autofilters.
@interface ExcelSortfields : ExcelBaseObject

- (void) addSortfieldKey:(ExcelRange *)key sorton:(ExcelE301)sorton order:(ExcelE228)order customorder:(id)customorder dataoption:(ExcelE302)dataoption;  // Creates a new sort field and returns a sortfields object.

@end

@interface ExcelSpinner : ExcelBaseObject

@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property BOOL displayThreeDShading;  // Returns or sets whether a 3-d effect will be employed when displaying the control
@property BOOL enabled;  // Returns or sets if the control is enabled
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or sets the height of the control.
@property double leftPosition;  // Returns or sets the left position of the control
@property (copy) NSString *linkedCell;  // Returns or sets reference to a call which contains the value of the control.
@property BOOL locked;  // Returns or sets whether the control is locked for editing.
@property NSInteger maximumValue;  // Returns or sets the maximum value allowed
@property NSInteger minimumValue;  // Returns or sets the minimum value allowed
@property (copy) NSString *name;  // Returns or sets the name of this control
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property NSInteger smallChange;  // Returns or sets the change in value of one click on the spinner control.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property NSInteger value;  // Returns or sets the current value of the control
@property BOOL visible;  // Returns or sets if the control is visible
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// Represents the sheet tab of a work sheet or chart sheet.
@interface ExcelTab : ExcelBaseObject

@property (copy) NSColor *color;
@property ExcelE103 colorIndex;
@property ExcelE305 themeColor;
@property double tintAndShade;


@end

// Represents a table style element
@interface ExcelTableStyleElement : ExcelBaseObject

@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (readonly) BOOL hasFormat;  // Returns or sets if table style element has format.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns a interior object that represents the interior for the specified table style element. 


@end

// Represents a table style
@interface ExcelTableStyle : ExcelBaseObject

- (SBElementArray<ExcelTableStyleElement *> *) tableStyleElements;

@property (readonly) BOOL builtIn;  // if this is a built in table style or not
- (NSString *) default;
@property (copy, readonly) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) NSString *namelocal;  // Returns or sets the local name of the object.
@property BOOL showAsAvailablePivotTableStyle;  // whether to show as avaliable pivot table style
@property BOOL showAsAvailableTableStyle;  // whether to show as avaliable table style


@end

@interface ExcelTextbox : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property BOOL autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property double height;  // Returns or set the height of the object.
@property ExcelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelE126 orientation;  // May also be a number value from -90 to 90 degrees.
@property ExcelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL roundedCorners;
@property BOOL shadow;
@property (copy) NSString *stringValue;  // Returns or sets the text of the specified object.
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property ExcelE284 verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property BOOL wrapAutoText;  // Returns or sets if the auto text is wrapped.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

@interface ExcelTop10FormatCondition : ExcelBaseObject

@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property ExcelE315 calcFor;
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) ExcelE195 formatConditionType;  // Return the conditional format type.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property ExcelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.
@property BOOL top10Percentage;
@property NSInteger top10Rank;
@property ExcelE314 topOrBottom;


@end

// Represents the hierarchical member-selection control of a cube field.
@interface ExcelTreeviewControl : ExcelBaseObject

@property (copy) id drilled;
@property (copy) id hidden;


@end

@interface ExcelUniqueValuesFormatCondition : ExcelBaseObject

@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property ExcelE317 duplicateOrUnique;
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) ExcelE195 formatConditionType;  // Return the conditional format type.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property ExcelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

// Represents data validation for a worksheet range.
@interface ExcelValidation : ExcelBaseObject

@property ExcelE199 IMEMode;  // Returns or sets the description of the Japanese input rules.
@property (readonly) ExcelE200 alertStyle;  // Returns the validation alert style.
@property (copy) NSString *errorMessage;  // Returns or sets the data validation error message.
@property (copy) NSString *errorTitle;  // Returns or sets the title of the data-validation error dialog box.
@property (copy, readonly) NSString *formula1;  // Returns the value or expression associated with the conditional format or data validation.
@property (copy, readonly) NSString *formula2;  // Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is operator between or operator not between.
@property BOOL ignoreBlank;  // Returns or sets if blank values are permitted by the range data validation.
@property BOOL inCellDropdown;  // Returns or sets if data validation displays a drop-down list that contains acceptable values.
@property (copy) NSString *inputMessage;  // Returns or sets the data validation input message.
@property (copy) NSString *inputTitle;  // Returns or sets the title of the data-validation input dialog box.
@property BOOL showError;  // Returns or sets if the data validation error message will be displayed whenever the user enters invalid data.
@property BOOL showInput;  // Returns or sets if the data validation input message will be displayed whenever the user selects a cell in the data validation range.
@property (readonly) ExcelE196 validationOperator;  // Returns the operator for the conditional format or data validation.
@property (readonly) ExcelE198 validationType;  // Returns the data validation type.
@property (readonly) BOOL value;  // Returns true if all the validation criteria are met, that is, if the range contains valid data.

- (void) addDataValidationType:(ExcelE198)type alertStyle:(ExcelE200)alertStyle operator:(ExcelE196)operator_ formula1:(NSString *)formula1 formula2:(NSString *)formula2;  // Adds data validation to the specified range.
- (void) modifyType:(ExcelE198)type alertStyle:(ExcelE200)alertStyle operator:(ExcelE196)operator_ formula1:(NSString *)formula1 formula2:(NSString *)formula2;  // Modifies data validation for a range.

@end

@interface ExcelValueChange : ExcelBaseObject

@property (readonly) ExcelE905 allocationMethod;
@property (readonly) ExcelE904 allocationValue;
@property (copy, readonly) NSString *allocationWeightExpression;
@property (readonly) NSInteger order;
@property (copy, readonly) ExcelPivotCell *pivotCell;
@property (copy, readonly) NSString *tuple;
@property (readonly) double value;
@property (readonly) BOOL visibleInPivotTable;


@end

// Represents a vertical page break. 
@interface ExcelVerticalPageBreak : ExcelBaseObject

@property (readonly) ExcelE190 extent;  // Returns the type of the specified page break: full-screen or only within a print area.
@property (copy) ExcelRange *location;  // Returns or sets where this vertical page break is located.
@property ExcelE168 verticalPageBreakType;  // Returns or sets the type of vertical page break.


@end

// Contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.
@interface ExcelWebOptions : ExcelBaseObject

@property BOOL allowPng;  // Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save documents as a Web page.
@property ExcelMtEn encoding;  // Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document.
@property (copy) NSString *locationOfComponents;  // Returns or sets the central URL, on the intranet or Web, or path, local or network, to the location from which authorized users can download Microsoft Office Web components when viewing your saved document.
@property NSInteger pixelsPerInch;  // Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120.
@property (readonly) ExcelMSsz screenSize;  // Returns the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser.
@property (copy) NSString *webPageKeywords;  // Returns or sets the keywords for the Web page.
@property (copy) NSString *webPageTitle;  // Returns or sets the title for the web page.

- (void) useDefaultFolderSuffix;  // Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.

@end

// Represents a window. Many worksheet characteristics, such as scroll bars and gridlines, are actually properties of the window.
@interface ExcelWindow : ExcelBasicWindow

- (SBElementArray<ExcelPane *> *) panes;

@property (copy, readonly) ExcelCell *activeCell;  // Returns a range object that represents the active cell in the active window, the window on top,  or in the specified window. If the window isn't displaying a worksheet, this property fails.
@property (copy, readonly) ExcelChart *activeChart;  // Returns a chart object that represents the active chart, either an embedded chart or a chart sheet. An embedded chart is considered active when it's either selected or activated.
@property (copy, readonly) ExcelPane *activePane;  // Returns a pane object that represents the active pane in the active window, the window on top,  or in the specified window. If the window isn't displaying a worksheet, this property fails.
@property (copy, readonly) ExcelSheet *activeSheet;  // Returns an object that represents the active sheet, the sheet on top, in the active workbook or in the specified window or workbook.
@property (copy) NSString *caption;  // Returns or sets the name that appears in the title bar of the document window. 
@property BOOL dateGrouping;  // Returns or sets the date grouping flag in the window.
@property BOOL displayFormulas;  // Returns or set if the window is displaying formulas.  If false, the window is displaying values.
@property BOOL displayGridlines;  // Returns or sets if gridlines are displayed.
@property BOOL displayHeadings;  // Returns or sets if both row and column headings are displayed.
@property BOOL displayHorizontalScrollBar;  // Returns or sets if the horizontal scroll bar is displayed.
@property BOOL displayOutline;  // Returns or sets if outline symbols are displayed.
@property BOOL displayVerticalScrollBar;  // Returns or sets if the vertical scroll bar is displayed.
@property BOOL displayWorkbookTabs;  // Returns or sets if the workbook tabs are displayed.
@property BOOL displayZeros;  // Returns or sets if zero values are displayed.
@property BOOL enableResize;  // Returns or sets if the window can be resized. 
@property (readonly) NSInteger entry_index;  // Returns the index of this item in the element list of windows.
@property BOOL freezePanes;  // Returns or sets if split panes are frozen. It's possible for freeze panes to be true and split to be false, or vice versa. This property applies only to worksheets and macro sheets.
@property (copy) NSColor *gridlineColor;  // Returns or sets the gridline color as an RGB value. 
@property ExcelE103 gridlineColorIndex;  // Returns or sets the gridline color as an index into the current color palette.
@property double height;  // Returns or sets the height of the window. Use the usable height property to determine the maximum size for the window. You cannot set this property if the window is maximized or minimized. Use the window state property to determine the window state.
@property double leftPosition;  // Returns or sets the distance from the left edge of the client area to the left edge of the window.
@property (copy, readonly) ExcelRange *rangeSelection;  // Returns a range object that represents the selected cells on the worksheet in the specified window even if a graphic object is active or selected on the worksheet.
@property NSInteger scrollColumn;  // Returns or sets the number of the leftmost column in the window. 
@property NSInteger scrollRow;  // Returns or sets the number of the row that appears at the top of the window.
@property (copy, readonly) NSArray *selectedSheets;  // Returns a list of sheet objects that represents all the selected sheets in the specified window.
@property (copy, readonly) ExcelRange *selection;  // Returns the selected range object in the specified window.
@property BOOL split;  // Returns or sets if the window is split. 
@property NSInteger splitColumn;  // Returns or sets the column number where the window is split into panes, the number of columns to the left of the split line.
@property double splitHorizontal;  // Returns or sets the location of the horizontal window split, in points.
@property NSInteger splitRow;  // Returns or sets the row number where the window is split into panes, the number of rows above the split.
@property double splitVertical;  // Returns or sets the location of the vertical window split, in points.
@property double tabRatio;  // Returns or sets the ratio of the width of the workbook's tab area to the width of the window's horizontal scroll bar, as a number between 0 and 1, the default value is 0.75. 
@property double top;  //  The distance from the top edge of the window to the top edge of the usable area, below the menus, any toolbars docked at the top, and the formula bar. You cannot set this property for a maximized window.
@property (readonly) double usableHeight;  // Returns the maximum height of the space that a window can occupy in points.
@property (readonly) double usableWidth;  // Returns the maximum width of the space that a window can occupy in the application window area, in points.
@property ExcelE239 view;  // Returns or sets the view showing in the window.
@property BOOL visible;  // Returns or sets if the window is visible.
@property (copy, readonly) ExcelRange *visibleRange;  // Returns a range object that represents the range of cells that are visible in the window or pane. If a column or row is partially visible, it's included in the range. 
@property double width;  // Returns or sets the width of the window.
@property (readonly) NSInteger windowNumber;  // Returns the window number. For example, a window named Book1.xls:2 has 2 as its window number. Most windows have the window number 1.
@property ExcelE111 windowState;  // Returns or sets the state of the window.
@property (readonly) ExcelE154 windowType;  // Returns the window type.
@property ExcelE278 zoom;  // Returns or sets the display size of the window, as a percentage,100 equals normal size, 200 equals double size, and so on. 

- (void) activateNext;  // Activates the current window, sends it to the back of the window z-order, then activates the next window according to the z-order.
- (void) activatePrevious;  // Activates the specified window and then activates the window at the back of the window z-order.
- (void) scrollWorkbookTabsSheets:(NSInteger)sheets position:(ExcelE271)position;  // Scrolls through the workbook tabs at the bottom of the window. Doesn't affect the active sheet in the workbook.

@end

// A connection is a set of information needed to obtain data from an external data source other than an 1st_Excel12 workbook.
@interface ExcelWorkbookConnection : ExcelBaseObject

@property (copy, readonly) NSString *_default;
@property (copy, readonly) NSString *objectDescription;  // Returns or sets a brief description for a WorkbookConnection object.
@property (copy, readonly) NSString *name;  // Returns or sets the name of the workbook connection object.
@property (copy, readonly) id ranges;  // Returns the range of object for the specified WorkbookConnection object.
@property (readonly) ExcelE917 type;


@end

// Represents a Microsoft Excel workbook.
@interface ExcelWorkbook : ExcelBaseObject

- (SBElementArray<ExcelDocumentProperty *> *) documentProperties;
- (SBElementArray<ExcelChartSheet *> *) chartSheets;
- (SBElementArray<ExcelCommandBar *> *) commandBars;
- (SBElementArray<ExcelCustomDocumentProperty *> *) customDocumentProperties;
- (SBElementArray<ExcelNamedItem *> *) namedItems;
- (SBElementArray<ExcelPivotCache *> *) pivotCaches;
- (SBElementArray<ExcelSheet *> *) sheets;
- (SBElementArray<ExcelStyle *> *) styles;
- (SBElementArray<ExcelCustomView *> *) customViews;
- (SBElementArray<ExcelWindow *> *) windows;
- (SBElementArray<ExcelWorksheet *> *) worksheets;
- (SBElementArray<ExcelInternationalMacroSheet *> *) internationalMacroSheets;
- (SBElementArray<ExcelMacroSheet *> *) macroSheets;
- (SBElementArray<ExcelTableStyle *> *) tableStyles;

@property BOOL acceptLabelsInFormulas;  // Returns or sets if labels can be used in worksheet formulas.
@property NSInteger accuracyVersion;  // Returns or sets the accuracy version for this workbook.
@property (copy, readonly) ExcelChart *activeChart;  // Returns a chart object that represents the active chart, either an embedded chart or a chart sheet. An embedded chart is considered active when it's either selected or activated.
@property (copy, readonly) ExcelSheet *activeSheet;  // Returns an object that represents the active sheet, the sheet on top, in the specified window.
@property NSInteger autoUpdateFrequency;  // Returns or sets the number of minutes between automatic updates to the shared workbook. If this property is set to zero  updates occur only when the workbook is saved.
@property BOOL autoUpdateSaveChanges;  // Returns or sets if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. if false changes aren't posted, this workbook is still synchronized with changes made by other users.
@property (readonly) NSInteger calculationVersion;  // Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits, on the left, are the major version of Microsoft Excel.
@property NSInteger changeHistoryDuration;  // Returns or sets the number of days shown in the shared workbook's change history.
@property ExcelE222 conflictResolution;  // Returns or sets the way conflicts are to be resolved whenever a shared workbook is updated.
@property (readonly) BOOL createBackup;  // Returns true if a backup file is created when this file is saved.
@property BOOL date1904;  // Returns or sets if the workbook uses the 1904 date system.
@property (copy) id defaultPivottableStyle;  // Set or get the default pivot table style for the current workbook
@property (copy) id defaultTableStyle;  // Set or get the default table style for the current workbook
@property ExcelE242 displayDrawingObjects;  // Returns or sets how shapes are displayed.
@property BOOL enableAutoRecover;  // Gets or sets a value that indicates whether Microsoft Office Excel saves changed files, of all formats, on a timed interval.
@property (readonly) BOOL excel8CompatibilityMode;  // Gets a value that indicates whether the workbook is in compatibility mode.
@property (readonly) ExcelE188 fileFormat;  // Returns the file format of the workbook.
@property (copy, readonly) NSString *fullName;  // Returns the name of the workbook, including its path on disk, as a string.
@property (copy, readonly) NSString *fullNameUrlencoded;  // Returns a String indicating the name of the object, including its path on disk, as a string.
@property (readonly) BOOL hasPassword;  // Returns true if the workbook has a protection password.
@property (readonly) BOOL hasVbProject;  // Gets a value that indicates whether a workbook has an attached Microsoft Visual Basic for Applications <VBA> project.
@property BOOL highlightChangesOnScreen;  // Returns or sets if changes to the shared workbook are highlighted on-screen.
@property BOOL inactiveListBorderVisible;  // Gets or sets a value that indicates whether list borders are visible when a list is not active.
@property BOOL isAddIn;  // Returns or sets if the workbook is running as an add in.
@property BOOL keepChangeHistory;  // Returns or sets if change tracking is enabled for the shared workbook.
@property BOOL listChangesOnNewSheet;  // Returns or sets if changes to the shared workbook are shown on a separate worksheet.
@property (readonly) BOOL multiUserEditing;  // Returns true if the workbook is open as a shared list.
@property (copy, readonly) NSString *name;  // Returns or sets the name of the workbook.
@property (copy) NSString *password;  // Returns or sets the password that must be supplied to open the specified workbook.
@property (copy, readonly) NSString *path;  // Returns the complete path of the object, excluding the final separator and name of the object.
@property BOOL personalViewListSettings;  // Returns or sets if filter and sort settings for lists are included in the user's personal view of the shared workbook.
@property BOOL personalViewPrintSettings;  // Returns or sets if print settings are included in the user's personal view of the shared workbook.
@property BOOL precisionAsDisplayed;  // Returns or sets if calculations in this workbook will be done using only the precision of the numbers as they're displayed.
@property (readonly) BOOL protectStructure;  // Returns true if the order of the sheets in the workbook is protected.
@property (readonly) BOOL protectWindows;  // Returns true if the windows of the workbook are protected.
@property (readonly) BOOL readOnly;  // Returns true if the workbook has been opened as read-only.
@property BOOL readOnlyRecommended;  // Returns or sets if the workbook was saved as read-only recommended.
@property BOOL removePersonalInformation;  // Returns or sets if personal information can be removed from the specified workbook.
@property (readonly) NSInteger revisionNumber;  // Returns the number of times the workbook has been saved while open as a shared list. If the workbook is open in exclusive mode, this property returns zero.
@property BOOL saveLinkValues;  // Returns or sets if Microsoft Excel saves external link values with the workbook.
@property BOOL saved;  // Returns or sets if no changes have been made to the specified workbook since it was last saved.
@property BOOL showConflictHistory;  // Returns or sets if the conflict history worksheet is visible in the workbook that's open as a shared list.
@property BOOL templateRemoveExternalData;  // Returns or sets if external data references are removed when the workbook is saved as a template.
@property (copy, readonly) ExcelOfficeTheme *theme;
@property BOOL updateRemoteReferences;  // Returns or sets if Microsoft Excel updates remote references in for the workbook.
@property (copy, readonly) NSArray *userStatus;  // Returns a list of lists that provides information about each user who has the workbook open as a shared list. The list is name of user, date and time user opened the workbook, and a number 1 for exclusive, 2 for shared access.
@property (copy, readonly) ExcelWebOptions *webOptions;  // Returns the web options object, which contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.
@property (copy) NSString *workbookComments;  // Returns or sets the comment string for this workbook.
@property (copy) NSString *writePassword;  // Returns or sets a string for the write password of a workbook.
@property (readonly) BOOL writeReserved;  // Return true if the workbook is write-reserved
@property (copy, readonly) NSString *writeReservedBy;  // Returns the name of the user who currently has write permission for the workbook.

- (void) acceptAllChangesWhen:(NSString *)when who:(NSString *)who where:(NSString *)where;  // Accepts all changes in the specified shared workbook.
- (void) applyThemeFileName:(NSString *)fileName;  // Applies a theme or design template to the specified workbook
- (void) breakLinkName:(NSString *)name type:(ExcelE165)type;  // Converts formulas linked to other Microsoft Excel sources to values.
- (BOOL) canCheckIn;  // Returns True if Excel can check in a specified workbook to a server.
- (void) changeFileAccessMode:(ExcelE175)mode writePassword:(NSString *)writePassword notify:(BOOL)notify;  // Changes the access permissions for the workbook. This may require an updated version to be loaded from the disk.
- (void) changeLinkName:(NSString *)name newName:(NSString *)newName type:(ExcelE165)type;  // Changes a link from one document to another.
- (void) checkInSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic;  // Returns a workbook from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) checkInWithVersionSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic versionType:(ExcelCivt)versionType;  // Returns a workbook from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) deleteNumberFormatNumberFormat:(NSString *)numberFormat;  // Deletes a custom number format from the workbook.
- (BOOL) exclusiveAccess;  // Assigns the current user exclusive access to the workbook that's open as a shared list.
- (void) followHyperlinkAddress:(NSString *)address subAddress:(NSString *)subAddress newWindow:(BOOL)newWindow extraInfo:(NSString *)extraInfo method:(ExcelMANT)method headerInfo:(NSString *)headerInfo;  // Displays a cached document, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.
- (void) highlightChangesOptionsWhen:(ExcelE193)when who:(NSString *)who where:(ExcelE267)where;  // Controls how changes are shown in a shared workbook.
- (SBObject *) linkInfoName:(NSString *)name linkInfo:(ExcelE171)linkInfo type:(ExcelE179)type;  // Returns the link date and update status.
- (NSArray *) linkSourcesType:(ExcelE165)type;  // Returns a list of links in the workbook. The names in the array are the names of the linked documents. Returns empty if there are no links.
- (void) mergeWorkbookFileName:(NSString *)fileName;  // Merges changes from one workbook into an open workbook.
- (ExcelWindow *) newWindowOnWorkbook NS_RETURNS_NOT_RETAINED;  // Creates a new window or a copy of the specified workbook window.
- (void) openLinksName:(NSString *)name readOnly:(BOOL)readOnly type:(ExcelE165)type;  // Opens the supporting documents for a link or links.
- (void) protectSharingFileName:(NSString *)fileName password:(NSString *)password writeReservationPassword:(NSString *)writeReservationPassword readOnlyRecommended:(BOOL)readOnlyRecommended createBackup:(BOOL)createBackup sharingPassword:(NSString *)sharingPassword fileFormat:(ExcelE188)fileFormat;  // Saves the workbook and protects it for sharing.
- (void) protectWorkbookPassword:(NSString *)password structure:(BOOL)structure windows:(BOOL)windows;  // Protect workbook structure from changes.
- (void) purgeChangeHistoryNowDays:(NSInteger)days sharingPassword:(NSString *)sharingPassword;  // Removes entries from the change log for the specified workbook.
- (void) refreshAll;  // Refreshes all external data ranges and PivotTables in the workbook.
- (void) rejectAllChangesWhen:(NSString *)when who:(NSString *)who where:(NSString *)where;  // Rejects all changes in the specified shared workbook.
- (void) removeUserEntry_index:(NSInteger)entry_index;  // Disconnects the specified user from the shared workbook.
- (void) resetColors;  // Resets the color palette to the default colors.
- (void) saveWorkbookAsFilename:(NSString *)filename fileFormat:(ExcelE188)fileFormat password:(NSString *)password writeReservationPassword:(NSString *)writeReservationPassword readOnlyRecommended:(BOOL)readOnlyRecommended createBackup:(BOOL)createBackup accessMode:(ExcelE221)accessMode conflictResolution:(ExcelE222)conflictResolution addToMostRecentlyUsedList:(BOOL)addToMostRecentlyUsedList overwrite:(BOOL)overwrite;  // Saves changes into a different file.
- (void) sendHtmlMail;  // Opens a message window in Microsoft Outlook for sending the specified document, formatted as html.
- (void) sendMail;  // Opens a message window in your registered mail program for sending the specified document as an attachment.
- (void) toggleFormsDesign;  // Toggles Microsoft Office Excel into and out of design mode.
- (void) unprotectSharingSharingPassword:(NSString *)sharingPassword;  // Turns off protection for sharing and saves the workbook.
- (void) updateFromFile;  // Updates a read-only workbook from the saved disk version of the workbook if the disk version is more recent than the copy of the workbook that is loaded in memory.
- (void) updateLinkName:(NSString *)name type:(ExcelE165)type;  // Updates a Microsoft Excel  link.
- (void) webPagePreview;  // Displays a preview of the specified workbook as it would look if saved as a Web page.

@end

@interface ExcelDocument : ExcelWorkbook


@end

@interface ExcelWorksheet : ExcelSheet

@property ExcelE197 enableSelection;  // Returns or sets what can be selected on the sheet.
@property (readonly) BOOL protectScenarios;  // Returns true if the worksheet scenarios are protected.

- (void) createSummaryForScenariosReportType:(ExcelE234)reportType resultCells:(ExcelE267)resultCells;  // Creates a new worksheet that contains a summary report for the scenarios on the specified worksheet.
- (void) mergeScenariosMergeSource:(ExcelE298)mergeSource;  // Merges the scenarios from the merge source worksheet into this worksheet

@end

// Contains global application-level spelling options
@interface ExcelXlspellingOptions : ExcelBaseObject

@property ExcelE300 dictionaryLang;  // Returns or sets the LCID used by the proofing tools


@end



/*
 * Drawing Suite
 */

@interface ExcelAdjustment : ExcelBaseObject

@property double adjustment_value;


@end

// Represents an arc graphic.
@interface ExcelArc : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property NSInteger addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property NSInteger autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (readonly) NSInteger bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property NSInteger caption;  // Returns or sets the caption for this object.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (readonly) NSInteger fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property NSInteger height;  // Returns or set the height of the object.
@property NSInteger horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (readonly) NSInteger interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property NSInteger leftPosition;  // Returns or sets the position of the specified object, in points.
@property NSInteger locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property NSInteger lockedText;  // Returns or sets whether the control's text is locked for editing.
@property NSInteger name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property NSInteger orientation;  // May also be a number value from -90 to 90 degrees.
@property NSInteger placement;  // Returns or sets how the object is placed on the worksheet.
@property NSInteger printObject;  // Returns or sets if this object is printed.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property NSInteger stringValue;  // Returns or sets the text of the specified object.
@property NSInteger top;  // Returns the top position of the specified object, in points.
@property (readonly) NSInteger topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property NSInteger verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property NSInteger visible;  // Returns or sets if the object is visible.
@property NSInteger width;  // Returns or sets  the width of the object.
@property BOOL wrapAutoText;  // Returns or sets if the auto text is wrapped.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

@interface ExcelBulletFormat : ExcelBaseObject

@property (copy) NSString *bulletCharacter;  // Returns or sets the Unicode character value that is used for bullets in the specified text.
@property (copy, readonly) ExcelShapeFont *bulletFont;  // Returns a font object that represents character formatting for a bullet format object.
@property (readonly) NSInteger bulletNumber;  // Returns the bullet number of a paragraph.
@property NSInteger bulletStartValue;  // Gets or sets the beginning value of a bulleted list.
@property ExcelPBtS bulletStyle;  // Returns or sets a constant that represents the style of a bullet.
@property ExcelPBtT bulletType;  // Returns or sets a constant that represents the type of bullet.
@property double relativeSize;  // Returns or sets the bullet size relative to the size of the first text character in the paragraph.
@property BOOL useTextColor;  // Determines whether the specified bullets are set to the color of the first text character in the paragraph.
@property BOOL useTextFont;  // Determines whether the specified bullets are set to the font of the first text character in the paragraph.
@property BOOL visible;  // Returns or sets a value that specifies whether the bullet is visible.

- (void) setBulletPictureFileName:(NSString *)FileName;  // Sets the graphics file to be used for bullets in a bulleted list.

@end

// Contains properties and methods that apply to line callouts.
@interface ExcelCalloutFormat : ExcelBaseObject

@property BOOL accent;  // Returns or sets if a vertical accent bar separates the callout text from the callout line.
@property ExcelMCAt angle;  // Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box.
@property BOOL autoAttach;  // Returns or sets if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line, where the callout points to, is to the left or right of the callout text box.
@property (readonly) BOOL autoLength;  // Returns if the length of the callout line is automatically set. Use the automatic length method to set this property to true, and use the custom length method to set this property to false.
@property BOOL border;  // Returns or sets whether the text in the specified callout is surrounded by a border.
@property (readonly) double calloutFormatLength;  // When the auto length property of the specified callout is set to false, the length property returns the length in points of the first segment of the callout line, the segment attached to the text callout box.
@property ExcelMCot calloutFormatType;  // Returns or sets the callout type.
@property (readonly) double drop;  // For callouts with an explicitly set drop value, this property returns the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.
@property (readonly) ExcelMCDt dropType;  // Returns a value that indicates where the callout line attaches to the callout text box.
@property double gap;  // Returns or sets the horizontal distance in points between the end of the callout line and the text bounding box.

- (void) automaticLength;  // Specifies that the first segment of the callout line, the segment attached to the text callout box, be scaled automatically when the callout is moved.

@end

// Contains properties and methods that apply to connectors. A connector is a line that attaches two other shapes at points called connection sites.
@interface ExcelConnectorFormat : ExcelBaseObject

@property (readonly) BOOL beginConnected;  // Returns true if the beginning of the specified connector is connected to a shape.
@property (copy, readonly) ExcelShape *beginConnectedShape;  // Returns a shape object that represents the shape that the beginning of the specified connector is attached to.
@property (readonly) NSInteger beginConnectionSite;  // Returns an integer that specifies the connection site that the beginning of a connector is connected to.
@property ExcelMCtT connectorFormatType;  // Returns or sets the connector type.
@property (readonly) BOOL endConnected;  // Returns true if the end of the specified connector is connected to a shape.
@property (copy, readonly) ExcelShape *endConnectedShape;  // Returns a shape object that represents the shape that the end of the specified connector is attached to.
@property (readonly) NSInteger endConnectionSite;  // Returns an integer that specifies the connection site that the end of a connector is connected to.

- (void) beginConnectConnectedShape:(ExcelShape *)connectedShape connectionSite:(NSInteger)connectionSite;  // Attaches the beginning of the specified connector to a specified shape. If there's already a connection between the beginning of the connector and another shape, that connection is broken.
- (void) beginDisconnect;  // Detaches the beginning of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the beginning of the connector remains positioned at a connection site but is no longer connected.
- (void) endConnectConnectedShape:(ExcelShape *)connectedShape connectionSite:(NSInteger)connectionSite;  // Attaches the end of the specified connector to a specified shape. If there's already a connection between the end of the connector and another shape, that connection is broken.
- (void) endDisconnect;  // Detaches the end of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the end of the connector remains positioned at a connection site but is no longer connected.

@end

// Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.
@interface ExcelFillFormat : ExcelBaseObject

- (SBElementArray<ExcelGradientStop *> *) gradientStops;

@property (copy) NSColor *backColor;  // Returns or sets a RGB color that represents the background color for the specified fill or patterned line.
@property ExcelMCoI backColorThemeIndex;  // Returns or sets the specified fill background color.
@property (readonly) ExcelMFdT fillFormatType;  // Returns the shape fill format type.
@property (copy) NSColor *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property ExcelMCoI foreColorThemeIndex;  // Returns or sets the specified foreground fill or solid color.
@property (readonly) ExcelMGCt gradientColorType;  // Returns the gradient color type for the specified fill.
@property (readonly) double gradientDegree;  // Returns a value that indicates how dark or light a one-color gradient fill is. A value of zero means that black is mixed in with the shape's foreground color to form the gradient; a value of 1 means that white is mixed in. Values between 1 and zero blend.
@property (readonly) ExcelMGdS gradientStyle;  // Returns the gradient style for the specified fill.
@property (readonly) NSInteger gradientVariant;  // Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. If the gradient style is from center gradient, this property returns either 1 or 2.
@property (readonly) ExcelPpTy pattern;  // Returns a value that represents the pattern applied to the specified fill or line.
@property (readonly) ExcelMPGb presetGradientType;  // Returns the preset gradient type for the specified fill.
@property (readonly) ExcelMPzT presetTexture;  // Returns the preset texture for the specified fill.
@property BOOL rotateWithObject;  // Returns or sets whether the fill rotates with the specified shape.
@property ExcelMTtA textureAlignment;  // Returns or sets the texture alignment for the specified object.
@property double textureHorizontalScale;  // Returns or sets the texture alignment for the specified object.
@property (copy, readonly) NSString *textureName;  // Returns the name of the custom texture file for the specified fill.
@property double textureOffsetX;  // Returns or sets the texture alignment for the specified object.
@property double textureOffsetY;  // Returns or sets the texture alignment for the specified object.
@property BOOL textureTile;  // Returns the texture tile style for the specified fill.
@property (readonly) ExcelMxtT textureType;  // Returns the texture type for the specified fill.
@property double textureVerticalScale;  // Returns or sets the texture alignment for the specified object.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque, and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.

- (void) deleteGradientStopStopIndex:(NSInteger)stopIndex;  // Removes a gradient stop.
- (void) insertGradientStopCustomColor:(NSColor *)customColor position:(double)position transparency:(double)transparency stopIndex:(NSInteger)stopIndex;  // Adds a stop to a gradient.
- (void) oneColorGradientGradientStyle:(ExcelMGdS)gradientStyle variant:(NSInteger)variant degree:(double)degree;  // Sets the specified fill to a one-color gradient.
- (void) patternedPattern:(ExcelPpTy)pattern;  // Sets the specified fill to a pattern.
- (void) presetGradientGradientStyle:(ExcelMGdS)gradientStyle variant:(NSInteger)variant presetGradientType:(ExcelMPGb)presetGradientType;  // Sets the specified fill to a preset gradient.
- (void) presetTexturedPresetTexture:(ExcelMPzT)presetTexture;  // Sets the specified fill to a preset texture.
- (void) solid;  // Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.
- (void) twoColorGradientGradientStyle:(ExcelMGdS)gradientStyle variant:(NSInteger)variant;  // Sets the specified fill to a two-color gradient.
- (void) userPicturePictureFile:(NSString *)pictureFile;  // Fills the specified shape with one large image.
- (void) userTexturedTextureFile:(NSString *)textureFile;  // Fills the specified shape with small tiles of an image.

@end

// Represents the glow formatting for a shape or range of shapes
@interface ExcelGlowFormat : ExcelBaseObject

@property (copy) NSColor *color;  // Returns or sets the color for the specified glow format.
@property ExcelMCoI colorThemeIndex;  // Returns or sets the color for the specified glow format.
@property double radius;  // Returns or sets the length of the radius for the specified glow format.


@end

// Represents one gradient stop.
@interface ExcelGradientStop : ExcelBaseObject

@property (copy) NSColor *color;  // Returns or sets the color for the specified the gradient stop.
@property ExcelMCoI colorThemeIndex;  // Returns or sets the color for the specified gradient stop.
@property double position;  // Returns or sets the position for the specified gradient stop expressed as a percent.
@property double transparency;  // Returns or sets a value representing the transparency of the gradient fill expressed as a percent.


@end

// Represents line and arrowhead formatting. For a line, the line format object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.
@interface ExcelLineFormat : ExcelBaseObject

@property (copy) NSColor *backColor;  // Returns or sets a RGB color that represents the background color for the specified fill or patterned line.
@property ExcelMCoI backColorThemeIndex;  // Returns or sets the background color for a patterned line.
@property ExcelMAhL beginArrowheadLength;  // Returns or sets the length of the arrowhead at the beginning of the specified line.
@property ExcelMAhS beginArrowheadStyle;  // Returns or sets the style of the arrowhead at the beginning of the specified line.
@property ExcelMAhW beginArrowheadWidth;  // Returns or sets the width of the arrowhead at the beginning of the specified line.
@property ExcelMlDs dashStyle;  // Returns or sets the dash style for the specified line.
@property ExcelMAhL endArrowheadLength;  // Returns or sets the length of the arrowhead at the end of the specified line.
@property ExcelMAhS endArrowheadStyle;  // Returns or sets the style of the arrowhead at the end of the specified line.
@property ExcelMAhW endArrowheadWidth;  // Returns or sets the width of the arrowhead at the end of the specified line.
@property (copy) NSColor *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property ExcelMCoI foreColorThemeIndex;  // Returns or sets the foreground color for the line.
@property ExcelMLnS lineStyle;  // Returns or sets the line format style.
@property ExcelPpTy pattern;  // Returns or sets a value that represents the pattern applied to the specified fill or line.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.
@property double weight;  // Returns or sets the thickness of the specified line in points.


@end

// Represents a line graphic object.
@interface ExcelLine : ExcelBaseObject

@property ExcelE113 arrowheadLength;  // Returns or sets the length of an arrowhead
@property ExcelE119 arrowheadStyle;  // Returns or sets the style of an arrowhead.
@property ExcelE120 arrowheadWidth;  // Returns or sets the width of an arrowhead.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or set the height of the object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents a Microsoft Office theme.
@interface ExcelOfficeTheme : ExcelBaseObject

@property (copy, readonly) ExcelThemeColorScheme *themeColorScheme;  // Returns the color scheme of a Microsoft Office theme.
@property (copy, readonly) ExcelThemeEffectScheme *themeEffectScheme;  // Returns the effects scheme of a Microsoft Office theme.
@property (copy, readonly) ExcelThemeFontScheme *themeFontScheme;  // Returns the font scheme of a Microsoft Office theme.


@end

// Represents an oval graphic.
@interface ExcelOval : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property NSInteger addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property NSInteger autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (readonly) NSInteger bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property NSInteger caption;  // Returns or sets the caption for this object.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (readonly) NSInteger fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property NSInteger height;  // Returns or set the height of the object.
@property NSInteger horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (readonly) NSInteger interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property NSInteger leftPosition;  // Returns or sets the position of the specified object, in points.
@property NSInteger locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property NSInteger lockedText;  // Returns or sets whether the control's text is locked for editing.
@property NSInteger name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property NSInteger orientation;  // May also be a number value from -90 to 90 degrees.
@property NSInteger placement;  // Returns or sets how the object is placed on the worksheet.
@property NSInteger printObject;  // Returns or sets if this object is printed.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property NSInteger shadow;  // Returns or sets if the object has a shadow.
@property NSInteger stringValue;  // Returns or sets the text of the specified object.
@property NSInteger top;  // Returns the top position of the specified object, in points.
@property (readonly) NSInteger topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property NSInteger verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property NSInteger visible;  // Returns or sets if the object is visible.
@property NSInteger width;  // Returns or sets  the width of the object.
@property BOOL wrapAutoText;  // Returns or sets if the auto text is wrapped.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

@interface ExcelParagraphFormat : ExcelBaseObject

- (SBElementArray<ExcelTabStop *> *) tabStops;

@property ExcelPpgA alignment;  // Returns or sets a value specifying the alignment of the paragraph.
@property ExcelPBlA baselineAlignment;  // Returns or sets a constant that represents the vertical position of fonts in a paragraph.
@property (copy, readonly) ExcelBulletFormat *bullet;  // Returns a bullet format object for the paragraph.
@property BOOL eastAsianLineBreakLevel;  // Returns or sets the East Asian line break control level for the specified paragraph.
@property double firstLineIndent;  // Returns or sets the value, in points, for a first line or hanging indent.
@property BOOL hangingPunctuation;  // Determines whether hanging punctuation is enabled for the specified paragraphs.
@property NSInteger indentLevel;  // Returns or sets a value representing the indent level assigned to text in the selected paragraph.
@property double leftIndent;  // Returns or sets a value that represents the left indent value, in points, for the specified paragraphs.
@property BOOL lineRuleAfter;  // Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines.
@property BOOL lineRuleBefore;  // Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines.
@property BOOL lineRuleWithin;  // Determines whether line spacing between base lines is set to a specific number of points or lines.
@property double rightIndent;  // Returns or sets the right indent, in points, for the specified paragraphs.
@property double spaceAfter;  // Returns or sets the amount of spacing, in points, after the specified paragraph.
@property double spaceBefore;  // Returns or sets the spacing, in points, before the specified paragraphs.
@property double spaceWithin;  // Returns or sets the amount of space between base lines in the specified paragraph, in points or lines.
@property (copy) id textDirection;  // Returns or sets the text direction for the specified paragraph.
@property BOOL wordWrap;  // Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs.


@end

// Contains properties and methods that apply to pictures.
@interface ExcelPictureFormat : ExcelBaseObject

@property double brightness;  // Returns or sets the brightness of the specified picture . The value for this property must be a number from 0.0, dimmest to 1.0, brightest.
@property ExcelMPc colorType;  // Returns or sets the type of color transformation applied to the specified picture.
@property double contrast;  // Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast to 1.0, the greatest contrast.
@property double cropBottom;  // Returns or sets the number of points that are cropped off the bottom of the specified picture.
@property double cropLeft;  // Returns or sets the number of points that are cropped off the left side of the specified picture.
@property double cropRight;  // Returns or sets the number of points that are cropped off the right side of the specified picture.
@property double cropTop;  // Returns or sets the number of points that are cropped off the top of the specified picture.
@property (copy) NSColor *transparencyColor;  // Returns or sets the transparent color for the specified picture as aRGB color. For this property to take effect, the transparent background property must be set to true.
@property BOOL transparentBackground;  // Returns or sets if the parts of the picture that are defined with a transparent color actually appear transparent.


@end

// Represents a rectangle graphic object.
@interface ExcelRectangle : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property NSInteger addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property NSInteger autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (readonly) NSInteger bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property NSInteger caption;  // Returns or sets the caption for this object.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (readonly) NSInteger fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property NSInteger height;  // Returns or set the height of the object.
@property NSInteger horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (readonly) NSInteger interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property NSInteger leftPosition;  // Returns or sets the position of the specified object, in points.
@property NSInteger locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property NSInteger lockedText;  // Returns or sets whether the control's text is locked for editing.
@property NSInteger name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property NSInteger orientation;  // May also be a number value from -90 to 90 degrees.
@property NSInteger placement;  // Returns or sets how the object is placed on the worksheet.
@property NSInteger printObject;  // Returns or sets if this object is printed.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property NSInteger roundedCorners;  // Returns or sets if the rectangle has rounded corners
@property NSInteger shadow;  // Returns or sets if the object has a shadow.
@property NSInteger stringValue;  // Returns or sets the text of the specified object.
@property NSInteger top;  // Returns the top position of the specified object, in points.
@property (readonly) NSInteger topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property NSInteger verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property NSInteger visible;  // Returns or sets if the object is visible.
@property NSInteger width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents the reflection effect in Office graphics.
@interface ExcelReflectionFormat : ExcelBaseObject

@property ExcelMRfT reflectionType;  // Returns or sets the type of the reflection format object.


@end

@interface ExcelRulerLevel : ExcelBaseObject

@property double firstMargin;  // Returns or sets the first-line indent for the specified outline level, in points.
@property double leftMargin;  // Returns or sets the left indent for the specified outline level, in points.


@end

// Represents the ruler for the text in the specified shape or for all text in the specified text style. Contains tab stops and the indentation settings for text outline levels.
@interface ExcelRuler : ExcelBaseObject

- (SBElementArray<ExcelTabStop *> *) tabStops;
- (SBElementArray<ExcelRulerLevel *> *) rulerLevels;


@end

// Represents shadow formatting for a shape.
@interface ExcelShadowFormat : ExcelBaseObject

@property double blur;  // Returns or sets the blur, in points, of the specified shadow.
@property (copy) NSColor *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property ExcelMCoI foreColorThemeIndex;  // Returns or sets the foreground color for the shadow format.
@property BOOL obscured;  // Returns or sets if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. If false the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.
@property double offsetX;  // Returns or sets the horizontal offset in points of the shadow from the specified shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.
@property double offsetY;  // Returns or sets the vertical offset in points of the shadow from the specified shape. A positive value offsets the shadow below the shape; a negative value offsets it above the shape.
@property BOOL rotateWithShape;  // Returns or sets whether to rotate the shadow when rotating the shape.
@property ExcelMSSt shadowStyle;  // Returns or sets the style of shadow formatting to apply to a shape.
@property ExcelMSdT shadowType;  // Returns or sets the shape shadow type.
@property double size;  // Returns or sets the width of the shadow.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

// Contains font attributes such as font name, size, and color, for an object.
@interface ExcelShapeFont : ExcelBaseObject

@property (copy) NSString *ASCIIName;  // Returns or sets the font used for Latin text; characters with character codes from 0 through 127.
@property BOOL autoRotateNumbers;  // Returns or sets a value that specifies whether the numbers in a numbered list should be rotated when the text is rotated.
@property double baseLineOffset;  // Returns or sets a value specifying the horizontaol offset of the selected font.
@property BOOL bold;  // Returns or sets a value specifying whether the font should be bold.
@property ExcelMTCa capsType;  // Returns or sets a value specifying how the text should be capitalized.
@property (copy) NSString *eastAsianName;  // Returns or sets the font name used for Asian text.
@property (readonly) BOOL embedable;  // Returns a value indicating whether the font can be embedded in a page.
@property (readonly) BOOL embedded;  // Returns a value specifying whether the font is embedded in a page.
@property BOOL equalizeCharacterHeight;  // Returns or sets a value specifying whether the text should have the same horizontal height.
@property (copy, readonly) ExcelFillFormat *fillFormat;  // Returns a fill format object that contains fill formatting properties for the specified font.
@property (copy) NSColor *fontColor;
@property ExcelMCoI fontColorThemeIndex;  // Returns or sets the color for the specified font.
@property (copy) NSString *fontName;  // Returns or sets a value specifying the font to use for a selection.
@property (copy) NSString *fontNameOther;  // Returns or sets the font used for characters whose character set numbers are greater than 127.
@property double fontSize;  // Returns or sets the font size.
@property (copy, readonly) ExcelGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property (copy) NSColor *highlightColor;  // Returns or sets the text highlight color for object.
@property ExcelMCoI highlightColorThemeIndex;  // Returns or sets the highlight color for the specified text.
@property BOOL italic;  // Returns or sets a value specifying whether the text for a selection is italic.
@property double kerning;  // Returns or sets a value specifying the amount of spacing between text characters.
@property (copy, readonly) ExcelLineFormat *lineFormat;  // Returns a value specifiying the format of a line.
@property (copy, readonly) ExcelReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property (copy, readonly) ExcelShadowFormat *shadowFormat;  // Returns the value specifying the type of shadow effect for the selection of text.
@property ExcelMSET softEdgeType;  // Returns or sets the type soft edge format object.
@property double spacing;  // Returns or sets a value specifying the spacing between characters in a selection of text.
@property ExcelMTSt strikeType;  // Returns or sets a value specifying the strike format used for a selection of text.
@property BOOL subscript;  // Returns or sets a value specifying that the selected text should be displayed a subscript.
@property BOOL superscript;  // Returns or sets a value specifying that the selected text should be displayed a superscript.
@property (copy) NSColor *underlineColor;  // Returns a value specifying the color of the underline for the selected text.
@property ExcelMCoI underlineColorThemeIndex;  // Returns a value specifying the color of the underline for the selected text.
@property ExcelMTUl underlineStyle;  // Returns or sets a value specifying the text effect for the selected text.
@property ExcelMPXF wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.


@end

// Represents the shape text frame in a shape object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.
@interface ExcelShapeTextFrame : ExcelBaseObject

@property (readonly) BOOL hasText;  // Returns whether the specified text frame has text.
@property ExcelMHzA horizontalAnchor;
@property double marginBottom;  // Returns or sets the distance, in points, between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text.
@property double marginLeft;  // Returns or sets the distance, in points, between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text.
@property double marginRight;  // Returns or sets the distance, in points, between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text.
@property double marginTop;  // Returns or sets the distance, in points, between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text.
@property ExcelTxOr orientation;  // Returns or sets the text orientation.
@property ExcelMPFo pathFormat;  // Returns or sets the path type for the specified text frame.
@property (copy, readonly) ExcelRuler *ruler;
@property (copy, readonly) ExcelTextColumn *textColumn;  // Returns the text column object that represents the columns within the text frame.
@property (copy, readonly) ExcelTextRange *textRange;
@property (copy, readonly) ExcelThreeDFormat *threeDFormat;  // Returns the 3-D-effect formatting properties for the specified text.
@property ExcelMVtA verticalAnchor;
@property ExcelMWFo warpFormat;  // Returns or sets the warp type for the specified text frame.
@property (readonly) ExcelMPXF wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.
@property BOOL wordWrap;  // Returns or sets text break lines within or past the boundaries of the shape.
@property ExcelPAtS wordartAutoSize;  // The size of the specified object that changes automatically to fit text within its boundaries.
@property ExcelMPXF wordartFormat;  // Returns or sets the WordArt type for the specified text frame.

- (void) ungroup;

@end

// Represents an object in the drawing layer.
@interface ExcelShape : ExcelBaseObject

- (SBElementArray<ExcelAdjustment *> *) adjustments;

@property (copy) NSString *alternativeText;  // Returns or sets the descriptive alternative text string for a Shape object when the object is saved to a Web page.
@property ExcelMAsT autoShapeType;  // Returns or sets the shape type for the specified shape object, which must represent an auto-shape.
@property ExcelMBSI backgroundStyle;  // Returns or sets the background style.
@property ExcelMBW blackWhiteMode;  // Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode.
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns a range object that represents the cell that lies under the lower-right corner of the object.
@property (copy, readonly) ExcelChart *chart;  // Returns a chart object that represents the chart contained in the shape.
@property (readonly) BOOL child;  // True if the shape is a child shape.
@property (readonly) NSInteger connectionSiteCount;  // Returns the number of connection sites on the specified shape.
@property (readonly) BOOL connector;  // Returns true if the specified shape is a connector.
@property (copy, readonly) ExcelConnectorFormat *connectorFormat;  // Returns a connector format object that contains connector formatting properties if this shape is a connector.
@property (readonly) ExcelMCtT connectorType;  // Returns the connector type if this shape is a connector.
@property (copy, readonly) ExcelFillFormat *fillFormat;  // Returns a fill format object that contains fill formatting properties for the specified shape.
@property (copy, readonly) ExcelGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property (readonly) BOOL hasChart;  // True if the specified shape has a chart.
@property double height;  // Returns or sets the height of the object.
@property (readonly) BOOL horizontalFlip;  // Returns true if the specified shape is flipped around the horizontal axis.
@property (copy, readonly) ExcelHyperlink *hyperlink;  // Returns a hyperlink object that represents the hyperlink for the shape.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) ExcelLineFormat *lineFormat;  // Returns a line format object for this shape.
@property BOOL lockAspectRatio;  // Returns or sets if the specified shape retains its original proportions when you resize it. If false, you can change the height and width of the shape independently of one another when you resize it.
@property BOOL locked;  // Returns or sets if the object is locked. If false, the object can be modified when the sheet is protected.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) ExcelShape *parentgroup;  // Returns a Shape object that represents the common parent shape of a child shape or a range of child shapes.
@property ExcelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property (copy, readonly) ExcelReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property double rotation;  // Returns or sets the rotation of the shape, in degrees.
@property (copy, readonly) ExcelShadowFormat *shadowFormat;  // Returns a shadow format object for this shape.
@property (copy) NSString *shapeOnAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelMSSI shapeStyle;  // Returns or sets the shape style corresponding to the Quick Styles.
@property (copy, readonly) ExcelShapeTextFrame *shapeTextFrame;  // Returns a shape text frame object that contains the alignment and anchoring properties for the specified shape.
@property (readonly) ExcelMShp shapeType;  // Returns the shape type.
@property (copy, readonly) ExcelSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) ExcelTextFrame *textFrame;  // Returns a text frame object that contains the alignment and anchoring properties for the specified shape.
@property (copy, readonly) ExcelThreeDFormat *threeDFormat;  // Returns a threeD format object that contains 3-D-effect formatting properties for the specified shape.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object.
@property (readonly) BOOL verticalFlip;  // Returns true if the specified shape is flipped around the vertical axis.
@property BOOL visible;  // Returns or sets if the object is visible
@property double width;  // Returns or sets  the width of the object.
@property (copy, readonly) ExcelWordArtFormat *wordArtFormat;  // Returns the formatting properties for a word art effect.
@property (readonly) NSInteger zOrderPosition;  // Returns the position of the specified shape in the z-order. To set the shape's position in the z-order, use the z order method.

- (void) apply;  // Applies to the specified shape formatting that's been copied by using the pick up method.
- (void) flipFlipCmd:(ExcelMFlC)flipCmd;  // Flips the specified shape around its horizontal or vertical axis.
- (void) pickUp;  // Copies the formatting of the specified shape. Use the apply method to apply the copied formatting to another shape.
- (void) rerouteConnections;  // Reroutes connectors so that they take the shortest possible path between the shapes they connect. To do this, the reroute connections method may detach the ends of a connector and reattach them to different connecting sites on the connected shapes.
- (void) saveAsPicturePictureType:(ExcelMPiT)pictureType fileName:(NSString *)fileName;  // Saves the shape in the requested file using the stated graphic format
- (void) setShapesDefaultProperties;  // Applies the formatting for the specified shape to the default shape. Shapes created after this method has been used will have this formatting applied by default.
- (void) zOrderZOrderCommand:(ExcelMZoC)zOrderCommand;  // Moves the specified shape in front of or behind other shapes in the collection, that is, changes the shape's position in the z-order.

@end

@interface ExcelCallout : ExcelShape

@property (copy, readonly) ExcelCalloutFormat *calloutFormat;  // Returns a connector format object that contains connector formatting properties.
@property (readonly) ExcelMCot calloutType;  // Returns the type of callout.


@end

@interface ExcelPicture : ExcelShape

@property (copy, readonly) NSString *fileName;  // Returns he name of the file that has the picture.
@property (readonly) BOOL linkToFile;  // Returns if the picture is lined to the file.
@property (copy, readonly) ExcelPictureFormat *pictureFormat;  // Returns a picture format object for this picture.
@property (readonly) BOOL saveWithDocument;  // Returns if the picture should be saved with the document.

- (void) scaleHeightFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(ExcelMSFr)scale;  // Scales the height of the shape by a specified factor.
- (void) scaleWidthFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(ExcelMSFr)scale;  // Scales the width of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original size or relative to the current size.

@end

@interface ExcelShapeConnector : ExcelShape

@property (copy, readonly) ExcelConnectorFormat *connectorFormat;  // Returns a connector format object that contains connector formatting properties.
@property (readonly) ExcelMCtT connectorType;  // Returns the connector type.


@end

// The line shape uses begin line X, begin line Y, end line X, and end line Y when created
@interface ExcelShapeLine : ExcelShape

@property double beginLineX;  // Returns or sets the beginning X position of the line.
@property double beginLineY;  // Returns or sets the beginning Y position of the line.
@property double endLineX;  // Returns or sets the ending X position of the line.
@property double endLineY;  // Returns or sets the ending Y position of the line.


@end

@interface ExcelShapeTextbox : ExcelShape

@property (readonly) ExcelTxOr textOrientation;  // Returns the text orientation of the object.


@end

// Represents the soft edge formatting for a shape or range of shapes
@interface ExcelSoftEdgeFormat : ExcelBaseObject

@property ExcelMSET softEdgeType;  // Returns or sets the type soft edge format object.


@end

// Represents a single tab stop.
@interface ExcelTabStop : ExcelBaseObject

@property double tabPosition;  // Returns or sets the position of a tab stop relative to the left margin.
@property ExcelPTSp tabStopType;  // Returns or sets the type of the tab stop object.


@end

// Represents a single text column.
@interface ExcelTextColumn : ExcelBaseObject

@property NSInteger columnNumber;  // Returns or sets the index of the text column object.
@property double spacing;  // Returns or sets the spacing between text columns in a text column object.
@property (copy) id textDirection;  // Returns or sets the direction of text in the text column object.


@end

// Represents the text frame in a shape object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.
@interface ExcelTextFrame : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property BOOL autoMargins;  // Returns or sets if Microsoft Excel automatically calculates text frame margins.
@property BOOL autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property ExcelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property ExcelHzOf horizontalOverflow;  // Returns or sets the horizontal overflow.
@property double marginBottom;  // Returns or sets the distance, in points, between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text.
@property double marginLeft;  // Returns or sets the distance, in points, between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text.
@property double marginRight;  // Returns or sets the distance, in points, between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text.
@property double marginTop;  // Returns or sets the distance, in points, between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text.
@property ExcelTxOr orientation;  // Returns or sets the text orientation.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property ExcelE114 verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property ExcelVrOf verticalOverflow;  // Returns or sets the vertical overflow.
@property BOOL wrapAutoText;  // Returns or sets if the auto text is wrapped.


@end

// Represents the color scheme of an Office Theme
@interface ExcelThemeColorScheme : ExcelBaseObject

- (SBElementArray<ExcelThemeColor *> *) themeColors;

- (NSColor *) getCustomColorName:(NSString *)name;  // Returns the custom color for the specified Microsoft Office theme.
- (void) loadThemeColorSchemeFileName:(NSString *)fileName;  // Loads the color scheme of a Microsoft Office theme from a file
- (void) saveThemeColorSchemeFileName:(NSString *)fileName;  // Saves the color scheme of a Microsoft Office theme to a file.

@end

// Represents a color in the color scheme of a Microsoft Office 2007 theme.
@interface ExcelThemeColor : ExcelBaseObject

@property (copy) NSColor *RGB;  // Returns or sets a value of a color in the color scheme of a Microsoft Office theme.
@property (readonly) ExcelMCSI themeColorSchemeIndex;  // Returns the index value a color scheme of a Microsoft Office theme.


@end

// Represents the effect scheme of a Microsoft Office theme.
@interface ExcelThemeEffectScheme : ExcelBaseObject

- (void) loadThemeEffectSchemeFileName:(NSString *)fileName;  // Loads the effects scheme of a Microsoft Office theme from a file

@end

// Represents the font scheme of a Microsoft Office theme.
@interface ExcelThemeFontScheme : ExcelBaseObject

- (SBElementArray<ExcelMinorThemeFont *> *) minorThemeFonts;
- (SBElementArray<ExcelMajorThemeFont *> *) majorThemeFonts;

- (void) loadThemeFontSchemeFileName:(NSString *)fileName;  // Loads the font scheme of a Microsoft Office theme from a file.
- (void) saveThemeFontSchemeFileName:(NSString *)fileName;  // Saves the font scheme of a Microsoft Office theme to a file.

@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface ExcelThemeFont : ExcelBaseObject

@property (copy) NSString *name;  // Returns or sets a value specifying the font to use for a selection.


@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface ExcelMajorThemeFont : ExcelThemeFont


@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface ExcelMinorThemeFont : ExcelThemeFont


@end

// Represents a shape's three-dimensional formatting.
@interface ExcelThreeDFormat : ExcelBaseObject

@property double ZDistance;  // Returns or sets a Single that represents the distance from the center of an object or text.
@property double bevelBottomDepth;  // Returns or sets the depth/height of the bottom bevel.
@property double bevelBottomInset;  // Returns or sets the inset size/width for the bottom bevel.
@property ExcelMBlT bevelBottomType;  // Returns or sets the bevel type for the bottom bevel.
@property double bevelTopDepth;  // Returns or sets the depth/height of the top bevel.
@property double bevelTopInset;  // Returns or sets the inset size/width for the top bevel.
@property ExcelMBlT bevelTopType;  // Returns or sets the bevel type for the top bevel.
@property (copy, readonly) NSColor *contourColor;  // Returns or sets the color of the contour of an object or text.
@property ExcelMCoI contourColorThemeIndex;  // Returns or sets the color for the specified contour.
@property double contourWidth;  // Returns or sets the width of the contour of an object or text.
@property double depth;  // Returns or sets the depth of the shape's extrusion.
@property (copy, readonly) NSColor *extrusionColor;  // Returns or sets a RGB color that represents the color of the shape's extrusion.
@property ExcelMCoI extrusionColorThemeIndex;  // Returns or sets the color for the specified extrusion.
@property ExcelMExC extrusionColorType;  // Returns or sets a value that indicates what will determine the extrusion color.
@property double fieldOfView;  // Returns or sets the amount of perspective for an object or text.
@property double lightAngle;  // Returns or sets the angle of the lighting.
@property BOOL perspective;  // Returns or sets if the extrusion appears in perspective that is, if the walls of the extrusion narrow toward a vanishing point. If false, the extrusion is a parallel, or orthographic, projection that is, if the walls don't narrow toward a vanishing point.
@property (readonly) ExcelMPzC presetCamera;  // Returns a constant that represents the camera preset.
@property (readonly) ExcelMExD presetExtrusionDirection;  // Returns or sets the direction taken by the extrusion's sweep path leading away from the extruded shape, the front face of the extrusion.
@property ExcelMPLd presetLightingDirection;  // Returns or sets the position of the light source relative to the extrusion.
@property ExcelMLtT presetLightingRig;  // Returns a constant that represents the lighting preset.
@property ExcelMlSf presetLightingSoftness;  // Returns or sets the intensity of the extrusion lighting.
@property ExcelMPMt presetMaterial;  // Returns or sets the extrusion surface material.
@property (readonly) ExcelM3DF presetThreeDFormat;
@property BOOL projectText;  // Returns or sets whether text on a shape rotates with shape.
@property double rotationX;  // Returns or sets the rotation of the extruded shape around the x-axis in degrees. A positive value indicates upward rotation; a negative value indicates downward rotation.
@property double rotationY;  // Returns or sets the rotation of the extruded shape around the y-axis, in degrees. A positive value indicates rotation to the left; a negative value indicates rotation to the right.
@property double rotationZ;  // Returns or sets the rotation of the extruded shape around the z-axis, in degrees. A positive value indicates clockwise rotation; a negative value indicates anti-clockwise rotation.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.

- (void) resetRotation;  // Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.

@end

// Contains properties and methods that apply to WordArt objects.
@interface ExcelWordArtFormat : ExcelBaseObject

@property ExcelMTxA alignment;  // Returns or sets a constant that represents the alignment for the specified text effect.
@property BOOL bold;  // Returns if the text of the word art shape is formatted as bold.
@property (copy) NSString *fontName;  // Returns or sets the font name of the font used by this word art shape.
@property double fontSize;  // Returns or sets the font size of the font used by this word art shape.
@property BOOL italic;  // Returns if the text of the word art shape is formatted as italic.
@property BOOL kernedPairs;  // Returns or sets if character pairs in a WordArt object have been kerned. 
@property BOOL normalizedHeight;  // Returns or sets if all characters, both uppercase and lowercase, in the specified WordArt are the same height.
@property ExcelMPTs presetShape;  // Returns or sets the shape of the specified WordArt.
@property ExcelMPXF presetWordArtEffect;  // Returns or sets the style of the specified WordArt.
@property BOOL rotatedChars;  // Returns or sets if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. If false, characters in the specified WordArt retain their original orientation relative to the bounding shape.
@property double tracking;  // Returns or sets the ratio of the horizontal space allotted to each character in the specified WordArt in relation to the width of the character. Can be a value from zero through 5.
@property (copy) NSString *wordArtText;  // Returns or sets the text associated with this word art shape.

- (void) toggleVerticalText;

@end



/*
 * Text Suite
 */

// Represents characters in an object that contains text.
@interface ExcelCharacter : ExcelBaseObject

@property (copy) NSString *content;  // Returns or sets the text for the specified object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *phoneticCharacters;  // Returns or sets the phonetic text in the specified characters object.

- (void) insertIntoString:(NSString *)string;  // Inserts a string preceding the selected characters.

@end

// Contains font attributes, such as font name, size, and color, for an object.
@interface ExcelFont : ExcelBaseObject

@property BOOL bold;  // Returns or sets if the font is formatted as bold.
@property (copy) NSColor *color;  // Returns or sets the color for the font.
@property ExcelE110 fontBackground;  // Returns or sets the text background type.
@property ExcelE103 fontColorIndex;  // Returns or sets the color index of the font.
@property NSInteger fontSize;  // Returns or sets the font size.
@property (copy) NSString *fontStyle;  // Returns or sets the font style.
@property BOOL italic;  // Returns or sets if the font is formatted as italic.
@property (copy) NSString *name;  // Returns or sets the font name associated with this font object.
@property BOOL outlineFont;  // Returns or sets if the specified font is formatted as outline.
@property BOOL shadow;  // Returns or sets if the specified font is formatted as shadowed.
@property BOOL strikethrough;  // Returns or sets if the font is formatted as strikethrough text.
@property BOOL subscript;  // Returns or sets if the font is formatted as subscript.
@property BOOL superscript;  // Returns or sets if the font is formatted as superscript.
@property ExcelE305 themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property ExcelE304 themeFontIndex;  // Returns or sets the theme font in the applied font scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.
@property ExcelE130 underline;  // Returns or sets the type of underline applied to the font.


@end

// Represents a style description for a range.
@interface ExcelStyle : ExcelBaseObject

@property BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically. 
@property (readonly) BOOL builtIn;  // Returns true if the style is a built-in style.
@property (copy, readonly) ExcelFont *fontObject;  // Returns the font object associated with this style.
@property BOOL formulaHidden;  // Returns or sets if the formula will be hidden when the worksheet is protected.
@property ExcelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property BOOL includeAlignment;  // Returns or sets if the style includes the add indent, horizontal alignment, vertical alignment, wrap text, and orientation properties.
@property BOOL includeBorder;  // Returns or sets if the style includes the color, color index, line style, and weight border properties.
@property BOOL includeFont;  // Returns or sets if the style includes the background, bold, color, color index, font style, italic, name, outline font, shadow, size, strikethrough, subscript, superscript, and underline font properties.
@property BOOL includeNumber;  // Returns or sets if the style includes the number format property. 
@property BOOL includePatterns;  // Returns or sets if the style includes the color, color index, invert if negative, pattern, pattern color, and pattern color index interior properties.
@property BOOL includeProtection;  // Returns or sets if the style includes the formula hidden and locked protection properties.
@property NSInteger indentLevel;  // Returns or sets the indent level for the style. Can be an integer from 0 to 15.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified style.
@property BOOL locked;  // Returns or sets if the range using this style is locked.
@property BOOL mergedCells;  // Returns true if the style contains merged cells.
@property (copy, readonly) NSString *name;  // Returns the name of the style.
@property (copy, readonly) NSString *nameLocal;  // Returns the name of the style, in the language of the user.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property (copy) NSString *numberFormatLocal;  // Returns or sets the format code for the object as a string in the language of the user.
@property ExcelE126 orientation;  // May also be a number value from -90 to 90 degrees.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shrinkToFit;  // Returns or sets if text automatically shrinks to fit in the available column width.
@property (copy, readonly) NSString *value;  // Return the name of the specified style.
@property ExcelE114 verticalAlignment;  // Returns or sets the vertical alignment of the style.
@property BOOL wrapText;  // Returns or sets if Microsoft Excel wraps the text in the object.

- (ExcelBorder *) getBorderWhichBorder:(ExcelE243)whichBorder;  // Returns the specified border object.

@end

// Represents the text frame in a shape object.
@interface ExcelTextRange : ExcelBaseObject

- (SBElementArray<ExcelTextRangeCharacter *> *) textRangeCharacters;
- (SBElementArray<ExcelWord *> *) words;
- (SBElementArray<ExcelTextRangeLine *> *) textRangeLines;
- (SBElementArray<ExcelSentence *> *) sentences;
- (SBElementArray<ExcelParagraph *> *) paragraphs;
- (SBElementArray<ExcelTextFlow *> *) textFlows;

@property (readonly) double boundHeight;  // Returns the height, in points, of the text bounding box for the specified text.
@property (readonly) double boundLeft;  // Returns the left coordinate, in points, of the text bounding box for the specified text.
@property (readonly) double boundTop;  // Returns the top coordinate, in points, of the text bounding box for the specified text.
@property (readonly) double boundWidth;  // Returns the width, in points, of the text bounding box for the specified text.
@property (copy) NSString *content;  // Returns or sets the text in a text range.
@property (copy, readonly) ExcelShapeFont *font;  // Returns the character formatting for the object.
@property (copy, readonly) ExcelParagraphFormat *paragraphFormat;  // Returns the paragraph formatting for the specified text.
@property (readonly) NSInteger textLength;  // Returns the length of a text range.
@property (readonly) NSInteger textStart;  // Returns the starting point of the specified text range.

- (void) addPeriodsTo;  // Adds period punctuation to the end of the text contained in text range object.
- (void) changeCaseTo:(ExcelPCgC)to;  // Changes the case of a text range object.
- (void) copyTextRange NS_RETURNS_NOT_RETAINED;  // Copies a text range object.
- (void) cutTextRange;  // Removes a portion or all of the text from a range of text.
- (NSArray *) getRotatedTextBounds;  // Returns back a list containing the four point of the text bounds as follows  {x1, y1}, {x2, y2}, {x3, y3}, {x4, y4} }
- (void) insertTextTextRangeInsertWhere:(ExcelMTiP)insertWhere newText:(NSString *)newText;  // Adds new text to the text range object.
- (void) pasteTextRange;  // Pastes the contents of the Clipboard into the text range object.
- (void) removePeriodsFrom;  // Removes all period punctuation from the text in the text range object.

@end

@interface ExcelParagraph : ExcelTextRange


@end

@interface ExcelSentence : ExcelTextRange


@end

@interface ExcelTextFlow : ExcelTextRange


@end

@interface ExcelTextRangeCharacter : ExcelTextRange


@end

@interface ExcelTextRangeLine : ExcelTextRange


@end

@interface ExcelWord : ExcelTextRange


@end



/*
 * Table Suite
 */

// Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3-D range.
@interface ExcelRange : ExcelBaseObject

- (SBElementArray<ExcelRange *> *) ranges;
- (SBElementArray<ExcelCell *> *) cells;
- (SBElementArray<ExcelRow *> *) rows;
- (SBElementArray<ExcelColumn *> *) columns;
- (SBElementArray<ExcelCharacter *> *) characters;
- (SBElementArray<ExcelFormatCondition *> *) formatConditions;
- (SBElementArray<ExcelHyperlink *> *) hyperlinks;

@property (copy, readonly) ExcelExcelComment *ExcelComment;  //  Returns a comment object that represents the comment associated with the cell in the upper-left corner of the range.
@property BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property (copy, readonly) NSArray *areas;  // Returns a list of  range objects  that represents all the ranges in a multiple-area selection.
@property double columnWidth;  // Returns or sets the width of all columns in the specified range.
@property (copy, readonly) id countLarge;  // Counts the largest value in a given range of values. Read-only Variant.
@property (copy, readonly) ExcelRange *currentArray;  // If the specified cell is part of an array, returns a range object that represents the entire array.
@property (copy, readonly) ExcelRange *currentRegion;  // Returns a range object that represents the current region. The current region is a range bounded by any combination of blank rows and blank columns.
@property (copy, readonly) ExcelRange *dependents;  // Returns a range object that represents the range containing all the dependents of a cell. This can be a multiple selection, a union of Range objects, if there's more than one dependent.
@property (copy, readonly) ExcelRange *directDependents;  // Returns a range object that represents the range containing all the direct dependents of a cell. This can be a multiple selection, a union of Range objects, if there's more than one dependent.
@property (copy, readonly) ExcelRange *directPrecedents;  // Returns a range object that represents the range containing all the direct precedents of a cell. This can be a multiple selection, a union of Range objects,  if there's more than one precedent.
@property (copy, readonly) ExcelDisplayFormat *displayFormat;  // Returns the display format associated with this range.
@property (copy, readonly) ExcelRange *entireColumn;  // Returns a range object that represents the entire column, or columns,  that contains the specified range.
@property (copy, readonly) ExcelRange *entireRow;  // Returns a range object that represents the entire row, or rows,  that contains the specified range.
@property (readonly) NSInteger firstColumnIndex;  // Returns the number of the first column in the first area in the specified range.
@property (readonly) NSInteger firstRowIndex;  // Returns the number of the first row of the first area in the range.
@property (copy, readonly) ExcelFont *fontObject;  // Returns the font object associated with this range.
@property (copy) id formula;  // Returns or sets the object's formula, in A1-style notation.
@property (copy) NSString *formulaArray;  // Returns or sets the array formula of a range.
@property BOOL formulaHidden;  // Returns or sets if the formula will be hidden when the workbook or worksheet is protected.
@property ExcelE192 formulaLabel;  // Returns or sets the formula label type for the specified range. 
@property (copy) NSString *formulaLocal;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formulaR1c1;  // Returns or sets the formula for the object, using R1C1-style notation.
@property (copy) NSString *formulaR1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property (readonly) BOOL hasArray;  // Returns true if the specified cell is part of an array formula.
@property (readonly) BOOL hasFormula;  // Returns true if all cells in the range contain formulas.
@property (readonly) double height;  // Returns the height of the range.
@property BOOL hidden;  // Returns or sets if the rows or columns are hidden. The specified range must span an entire column or row.
@property ExcelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the range.
@property NSInteger indentLevel;  // Returns or sets the indent level for the range or style. Can be an integer from 0 to 15.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (readonly) double leftPosition;  // The distance from the left edge of column A to the left edge of the range. If the range is discontinuous, the first area is used. If the range is more than one column wide, the leftmost column in the range is used.
@property (readonly) NSInteger listHeaderRows;  // Returns the number of header rows for the specified range.
@property (copy, readonly) ExcelListObject *listObject;  // Returns the list object associated with this range.
@property (readonly) ExcelE152 locationInTable;  // Returns an enumeration that describes the part of the pivot table that contains the upper-left corner of the specified range.
@property BOOL locked;  // Returns or sets  if the range is locked.
@property (copy, readonly) ExcelRange *mergeArea;  // Returns a range object that represents the merged range containing the specified cell. If the specified cell isn't in a merged range, this property returns the specified cell.
@property BOOL mergeCells;  // Returns or sets if the range contains merged cells. 
@property (copy) NSString *name;  // Returns or sets the name of the range.
@property (copy, readonly) NSString *namedItem;  // Returns the named item associated with this range.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property (copy) NSString *numberFormatLocal;  // Returns or sets the format code for the object as a string in the language of the user.
@property NSInteger outlineLevel;  // Returns or sets the current outline level of the specified row or column
@property ExcelE168 pageBreak;  // Returns or sets the location of a page break.
@property (copy, readonly) ExcelPhonetic *phoneticObject;  // Returns the phonetic object, which contains information about a specific phonetic text string in a specified range.
@property (copy, readonly) ExcelPivotField *pivotField;  // Returns a pivot field object that represents the pivot field containing the upper-left corner of the specified range.
@property (copy, readonly) ExcelPivotItem *pivotItem;  // Returns a pivot item object that represents the pivot item containing the upper-left corner of the specified range.
@property (copy, readonly) ExcelPivotTable *pivotTable;  // Returns a pivot table object that represents the pivot table containing the upper-left corner of the specified range.
@property (copy, readonly) ExcelRange *precedents;  // Returns a range object that represents all the precedents of a cell. This can be a multiple selection, a union of Range objects, if there's more than one precedent.
@property (copy, readonly) NSString *prefixCharacter;  // Returns the prefix character for the cell.
@property (copy, readonly) ExcelQueryTable *queryTable;  // Returns a query table object that represents the query table that intersects the specified range.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property double rowHeight;  // Returns  or sets the height of all the rows in the range specified, measured in points.
@property BOOL showDetail;  // Returns or sets if the outline is expanded for the specified range, so that the detail of the column or row is visible. The specified range must be a single summary column or row in an outline.
@property BOOL shrinkToFit;  // Returns or sets if text automatically shrinks to fit in the available column width.
@property (copy, readonly) id stringValue;  // Returns the text for the range.
@property (copy) ExcelStyle *styleObject;  // Returns or sets a style object that represents the style of the specified range.
@property (readonly) BOOL summary;  // Returns true if the range is an outlining summary row or column. The range should be a row or a column.
@property ExcelE126 textOrientation;  // The text orientation. Can be a number value from -90 to 90 degrees.
@property (readonly) double top;  // The distance from the top edge of row 1 to the top edge of the range. If the range is discontinuous, the first area is used. If the range is more than one row high, the top, lowest numbered, row in the range is used.
@property BOOL useStandardHeight;  // Returns or sets if the row height of the range object equals the standard height of the sheet.
@property BOOL useStandardWidth;  // Returns or sets if the column width of the range object equals the standard width of the sheet.
@property (copy, readonly) ExcelValidation *validation;  // Returns the validation object that represents data validation for the specified range.
@property (copy) id value;  // Returns or sets the value of the range.
@property (copy) id value2;  // Returns or sets the value of the range. The difference between this property and Value is that Value2 uses the numerical representation of cells that are formatted as currency and date.
@property ExcelE284 verticalAlignment;  // Returns or sets the vertical alignment of the range.
@property (readonly) double width;  // Returns the width of the range.
@property (copy, readonly) ExcelSheet *worksheetObject;  // Returns a worksheet object that represents the worksheet containing the specified range.
@property BOOL wrapText;  // Returns or sets if Microsoft Excel wraps the text in the object.

- (void) activateObject;  // Activates the object.
- (ExcelExcelComment *) addCommentCommentText:(NSString *)commentText;  // Adds a comment to the range.
- (void) advancedFilterAction:(ExcelE163)action criteriaRange:(ExcelE292)criteriaRange copyToRange:(ExcelE292)copyToRange unique:(BOOL)unique;  // Filters or copies data from a list based on a criteria range. If the initial selection is a single cell, that cell's current region is used.
- (void) applyNamesNames:(NSArray *)names ignoreRelativeAbsolute:(BOOL)ignoreRelativeAbsolute useRowColumnNames:(BOOL)useRowColumnNames omitColumn:(BOOL)omitColumn omitRow:(BOOL)omitRow order:(ExcelE166)order appendLast:(BOOL)appendLast;  // Applies names to the cells in the specified range.
- (void) applyOutlineStyles;  // Applies outlining styles to the specified range.
- (void) autoOutline;  // Automatically creates an outline for the specified range. If the range is a single cell, Microsoft Excel creates an outline for the entire sheet. The new outline replaces any existing outline.
- (NSString *) autocompleteString:(NSString *)string;  // Returns an AutoComplete match from the list. If there's no AutoComplete match or if more than one entry in the list matches the string to complete, this method returns an empty string.
- (void) autofillDestination:(ExcelE292)destination type:(ExcelE185)type;  // Performs an autofill on the cells in the specified range.
- (void) autofilterRangeField:(NSInteger)field criteria1:(id)criteria1 operator:(ExcelE186)operator_ criteria2:(id)criteria2 visibleDropDown:(BOOL)visibleDropDown;  // Filters a list using the AutoFilter.
- (void) autofit;  // Changes the width of the specified column range or the height of the specified row range to achieve the best fit.
- (void) autoformatFormat:(ExcelE215)format includeNumber:(BOOL)includeNumber font:(BOOL)font alignment:(BOOL)alignment border:(BOOL)border pattern:(BOOL)pattern width:(BOOL)width;  // Automatically formats a range of cells, using a predefined format.
- (void) borderAroundLineStyle:(ExcelE133)lineStyle weight:(ExcelE128)weight colorIndex:(ExcelE103)colorIndex color:(NSColor *)color;  // Adds a border to a range and sets the color, line style, and weight properties for the new border.
- (void) calculate;  // Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet.
- (void) calculateRowMajorOrder;  // Calculates a specfied range of cells.
- (void) checkSpellingCustomDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase alwaysSuggest:(BOOL)alwaysSuggest;  // Checks the spelling of an object.
- (void) clearExcelComments;  // Clears all cell comments from the specified range.
- (void) clearContents;  // Clears all cell comments from the specified range.
- (void) clearOutline;  // Clears the outline for the specified range.
- (void) clearRange;  // Clears the entire object.
- (void) clearRangeFormats;  // Clears the formatting of the object.
- (ExcelRange *) columnDifferencesComparison:(ExcelE292)comparison;  // Returns a range object that represents all the cells whose contents are different from the comparison cell in each column.
- (void) consolidateSources:(NSArray *)sources consolidationFunction:(ExcelE150)consolidationFunction topRow:(BOOL)topRow leftColumn:(BOOL)leftColumn createLinks:(BOOL)createLinks;  // Consolidates data from multiple ranges on multiple worksheets into a single range on a single worksheet.
- (void) copyPictureAppearance:(ExcelE207)appearance format:(ExcelE156)format NS_RETURNS_NOT_RETAINED;  // Copies the selected object to the clipboard as a picture.
- (void) copyRangeDestination:(ExcelRange *)destination NS_RETURNS_NOT_RETAINED;  //  Copies the range to the specified range.
- (void) createNamesTop:(BOOL)top leftPosition:(BOOL)leftPosition bottom:(BOOL)bottom right:(BOOL)right;  // Creates names in the specified range, based on text labels in the sheet.
- (void) cutRangeDestinationOfCut:(ExcelE292)destinationOfCut;  // Cuts the range to the clipboard.
- (void) dataSeriesRowcol:(ExcelE105)rowcol dataSeriesType:(ExcelE107)dataSeriesType date:(ExcelE129)date increment:(NSInteger)increment stop:(NSInteger)stop trend:(BOOL)trend;  // Creates a data series in the specified range.
- (void) dataTableRowInput:(ExcelCell *)rowInput columnInput:(ExcelCell *)columnInput;  // Creates a data table based on input values and formulas that you define on a worksheet.
- (void) deleteRangeShift:(ExcelE148)shift;  // Deletes the range
- (void) fillDown;  // Fills down from the top cell or cells in the specified range to the bottom of the range. The contents and formatting of the cell or cells in the top row of a range are copied into the rest of the rows in the range.
- (void) fillLeft;  // Fills left from the rightmost cell or cells in the specified range. The contents and formatting of the cell or cells in the rightmost column of a range are copied into the rest of the columns in the range.
- (void) fillRight;  // Fills right from the leftmost cell or cells in the specified range. The contents and formatting of the cell or cells in the leftmost column of a range are copied into the rest of the columns in the range.
- (void) fillUp;  // Fills up from the bottom cell or cells in the specified range to the top of the range. The contents and formatting of the cell or cells in the bottom row of a range are copied into the rest of the rows in the range.
- (ExcelRange *) findWhat:(NSString *)what after:(ExcelE292)after lookIn:(ExcelE153)lookIn lookAt:(ExcelE177)lookAt searchOrder:(ExcelE224)searchOrder searchDirection:(ExcelE223)searchDirection matchCase:(BOOL)matchCase matchByte:(BOOL)matchByte;  // Finds specific information in a range, and returns a range object that represents the first cell where that information is found. Returns nothing if no match is found. Doesn't affect the selection or the active cell.
- (ExcelRange *) findNextAfter:(ExcelE292)after;  // Continues a search that was begun with the find method. Finds the next cell that matches those same conditions and returns a range object that represents that cell. Doesn't affect the selection or the active cell.
- (ExcelRange *) findPreviousAfter:(ExcelE292)after;  // Continues a search that was begun with the find method. Finds the previous cell that matches those same conditions and returns a range object that represents that cell. Doesn't affect the selection or the active cell.
- (void) functionWizard;  // Starts the Function Wizard for the upper-left cell of the range.
- (id) getXMLValue;  // Returns the value of the specified range as XML.
- (NSString *) getAddressRowAbsolute:(BOOL)rowAbsolute columnAbsolute:(BOOL)columnAbsolute referenceStyle:(ExcelE158)referenceStyle external:(BOOL)external relativeTo:(ExcelE292)relativeTo;  // Returns the range reference in the language of the macro.
- (NSString *) getAddressLocalRowAbsolute:(BOOL)rowAbsolute columnAbsolute:(BOOL)columnAbsolute referenceStyle:(ExcelE158)referenceStyle external:(BOOL)external relativeTo:(ExcelE292)relativeTo;  // Returns the range reference in the language of the user.
- (ExcelBorder *) getBorderWhichBorder:(ExcelE243)whichBorder;  // Returns the specified border object.
- (ExcelRange *) getEndDirection:(ExcelE149)direction;  // Returns a range object that represents the cell at the end of the region that contains the source range.
- (ExcelRange *) getOffsetRowOffset:(NSInteger)rowOffset columnOffset:(NSInteger)columnOffset;  // Returns a range object that represents a range that's offset from the specified range.
- (ExcelRange *) getResizeRowSize:(NSInteger)rowSize columnSize:(NSInteger)columnSize;  // Resizes the specified range. Returns a Range object that represents the resized range.
- (BOOL) goalSeekGoal:(double)goal changingCell:(ExcelRange *)changingCell;  // Calculates the values necessary to achieve a specific goal. If the goal is an amount returned by a formula, this calculates a value that, when supplied to your formula, causes the formula to return the number you want. Returns True if successful.
- (BOOL) groupStart:(NSInteger)start end:(NSInteger)end by:(NSInteger)by periods:(NSArray *)periods;  // When the Range object represents a single cell in a pivot table field's data range, the group method performs numeric or date-based grouping in that field.
- (void) insertIndentInsertAmount:(NSInteger)insertAmount;  // Adds an indent to the specified range.
- (void) insertIntoRangeShift:(ExcelE147)shift;  // Inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.
- (void) justify;  // Rearranges the text in a range so that it fills the range evenly.
- (void) listNames;  // Pastes a list of all non-hidden names onto the worksheet, beginning with the first cell in the range.
- (void) mergeAcross:(BOOL)across;  // Creates a merged cell from the specified Range object.
- (void) navigateArrowTowardPrecedent:(BOOL)towardPrecedent arrowNumber:(NSInteger)arrowNumber linkNumber:(NSInteger)linkNumber;  // Navigates a tracer arrow for the specified range to the precedent, dependent, or error-causing cell or cells. Selects the precedent, dependent, or error cells and returns a range object that represents the new selection.
- (void) parseParseLine:(NSString *)parseLine destination:(ExcelE292)destination;  // Parses a range of data and breaks it into multiple cells. Distributes the contents of the range to fill several adjacent columns; the range can be no more than one column wide.
- (void) pasteSpecialWhat:(ExcelE204)what operation:(ExcelE203)operation skipBlanks:(BOOL)skipBlanks transpose:(BOOL)transpose;  // Pastes the contents of the Clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.
- (void) printOutFrom:(NSInteger)from to:(NSInteger)to copies:(NSInteger)copies preview:(BOOL)preview activePrinter:(NSString *)activePrinter printToFile:(BOOL)printToFile collate:(BOOL)collate;  // Prints the object
- (void) printPreviewEnableChanges:(BOOL)enableChanges;  // Shows a preview of the object as it would look when printed. This function has been deprecated.
- (void) removeSubtotal;  // Removes subtotals from a list.
- (BOOL) replaceWhat:(NSString *)what replacement:(NSString *)replacement lookAt:(ExcelE177)lookAt searchOrder:(ExcelE224)searchOrder matchCase:(BOOL)matchCase matchByte:(BOOL)matchByte matchControlCharacters:(BOOL)matchControlCharacters matchDiacritics:(BOOL)matchDiacritics;  // Finds and replaces characters in cells within a range. Doesn't change the selection or the active cell.
- (ExcelRange *) rowDifferencesComparison:(ExcelCell *)comparison;  // Returns a range object that represents all the cells whose contents are different from those of the comparison cell in each row.
- (NSString *) runVBMacroArg1:(id)arg1 arg2:(id)arg2 arg3:(id)arg3 arg4:(id)arg4 arg5:(id)arg5 arg6:(id)arg6 arg7:(id)arg7 arg8:(id)arg8 arg9:(id)arg9 arg10:(id)arg10 arg11:(id)arg11 arg12:(id)arg12 arg13:(id)arg13 arg14:(id)arg14 arg15:(id)arg15 arg16:(id)arg16 arg17:(id)arg17 arg18:(id)arg18 arg19:(id)arg19 arg20:(id)arg20 arg21:(id)arg21 arg22:(id)arg22 arg23:(id)arg23 arg24:(id)arg24 arg25:(id)arg25 arg26:(id)arg26 arg27:(id)arg27 arg28:(id)arg28 arg29:(id)arg29 arg30:(id)arg30;  //  Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.
- (NSString *) runXLMMacroArg1:(id)arg1 arg2:(id)arg2 arg3:(id)arg3 arg4:(id)arg4 arg5:(id)arg5 arg6:(id)arg6 arg7:(id)arg7 arg8:(id)arg8 arg9:(id)arg9 arg10:(id)arg10 arg11:(id)arg11 arg12:(id)arg12 arg13:(id)arg13 arg14:(id)arg14 arg15:(id)arg15 arg16:(id)arg16 arg17:(id)arg17 arg18:(id)arg18 arg19:(id)arg19 arg20:(id)arg20 arg21:(id)arg21 arg22:(id)arg22 arg23:(id)arg23 arg24:(id)arg24 arg25:(id)arg25 arg26:(id)arg26 arg27:(id)arg27 arg28:(id)arg28 arg29:(id)arg29 arg30:(id)arg30;  //  Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.
- (void) setXMLValueRangeValue:(id)rangeValue;  // Set the value of the specified range using XML.
- (void) show;  // Scrolls through the contents of the active window to move the range into view. The range must consist of a single cell in the active document.
- (void) showDependentsRemove:(BOOL)remove;  // Draws tracer arrows to the direct dependents of the range.
- (void) showErrors;  // Draws tracer arrows through the precedents tree to the cell that's the source of the error, and returns the range that contains that cell.
- (void) showPrecedentsRemove:(BOOL)remove;  // Draws tracer arrows to the direct precedents of the range.
- (void) sortKey1:(ExcelRange *)key1 order1:(ExcelE228)order1 key2:(ExcelRange *)key2 sortType:(ExcelE230)sortType order2:(ExcelE228)order2 key3:(ExcelRange *)key3 order3:(ExcelE228)order3 header:(ExcelE241)header orderCustom:(NSInteger)orderCustom matchCase:(BOOL)matchCase orientation:(ExcelE229)orientation sortMethod:(ExcelE226)sortMethod ignoreControlCharacters:(BOOL)ignoreControlCharacters ignoreDiacritics:(BOOL)ignoreDiacritics dataoption1:(ExcelE302)dataoption1 dataoption2:(ExcelE302)dataoption2 dataoption3:(ExcelE302)dataoption3;  // Sorts a pivot table, a range, or the current region, if the specified range contains only one cell.
- (void) sortSpecialSortMethod:(ExcelE226)sortMethod key1:(ExcelRange *)key1 order1:(ExcelE228)order1 type:(id)type key2:(ExcelRange *)key2 order2:(ExcelE228)order2 key3:(ExcelRange *)key3 order3:(ExcelE228)order3 header:(ExcelE241)header orderCustom:(NSInteger)orderCustom matchCase:(BOOL)matchCase orientation:(ExcelE229)orientation dataoption1:(ExcelE302)dataoption1 dataoption2:(ExcelE302)dataoption2 dataoption3:(ExcelE302)dataoption3;  // Uses East Asian sorting methods to sort the range, or uses the current region if the range contains only one cell.
- (ExcelRange *) specialCellsType:(ExcelE182)type value:(ExcelE231)value;  // Returns a Range object that represents all the cells that match the specified type and value.
- (void) subtotalGroupBy:(NSInteger)groupBy function:(ExcelE150)function totalList:(NSArray *)totalList replace:(BOOL)replace pageBreaks:(BOOL)pageBreaks summaryBelowData:(ExcelE232)summaryBelowData;  // Creates subtotals for the range, or the current region, if the range is a single cell.
- (void) textToColumnsDestination:(ExcelE292)destination dataType:(ExcelE236)dataType textQualifier:(ExcelE237)textQualifier consecutiveDelimiter:(BOOL)consecutiveDelimiter tab:(BOOL)tab semicolon:(BOOL)semicolon comma:(BOOL)comma space:(BOOL)space useOther:(BOOL)useOther otherChar:(NSString *)otherChar fieldInfo:(NSArray *)fieldInfo decimalSeparator:(NSString *)decimalSeparator thousandsSeparator:(NSString *)thousandsSeparator;  // Parses a column of cells that contain text into several columns.
- (void) ungroup;  // Promotes a range in an outline, that is, decreases its outline level. The specified range must be a row or column, or a range of rows or columns. If the range is in a pivot table, this method ungroups the items contained in the range.
- (void) unmerge;  // Separates a merged area into individual cells.

@end

@interface ExcelCell : ExcelRange


@end

@interface ExcelColumn : ExcelRange


@end

@interface ExcelRow : ExcelRange


@end



/*
 * Proofing Suite
 */

// Contains Microsoft Excel autocorrect attributes, capitalization of names of days, correction of two initial capital letters, automatic correction list, and so on.
@interface ExcelAutocorrect : ExcelBaseObject

@property BOOL AutomaticallyExpandTablesAsIType;  // Returns or sets if resizes the table to include data entered into a neighboring column or row.
@property BOOL AutomaticallyFillFormulas;  // Returns or sets if we automatically copies the formula to all the cells in the column when a formula is entered into a cell.
@property BOOL correctCapsLock;  // Returns or sets if Microsoft Excel automatically corrects accidental use of the CAPS LOCK key.
@property BOOL correctDays;  // Returns or sets if the first letter of day names is capitalized automatically.
@property BOOL correctInitialCaps;  // Returns or sets if words that begin with two capital letters are corrected automatically.
@property BOOL correctSentenceCaps;  // Returns or sets if Microsoft Excel automatically corrects sentence, first word, capitalization.
@property BOOL replaceText;  // Returns or sets if text in the list of autocorrect replacements is replaced automatically.

- (void) addReplacementTextToReplace:(NSString *)textToReplace replacementText:(NSString *)replacementText;  // Adds an entry to the array of autocorrect replacements.
- (void) deleteReplacementWhat:(NSString *)what;  // Deletes an entry from the array of autocorrect replacements. 
- (NSArray *) getReplacementList;  // Returns a list of autocorrect replacements.

@end



/*
 * Chart Suite
 */

// Represents a chart axis title.
@interface ExcelAxisTitle : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy) NSString *axisTitleText;  // Returns or sets the text associated with this object.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property ExcelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property BOOL includeInLayout;  // Returns or sets if a axis title will occupy the chart layout space when a chart layout is being determined.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property NSInteger orientation;  // The text orientation.  Can be a number value from -90 to 90 degrees.
@property ExcelE216 position;  // Returns or sets the position of the axis title on the chart.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property ExcelE284 verticalAlignment;  // Returns or sets the vertical alignment of the object.


@end

// Represents a single axis in a chart.
@interface ExcelAxis : ExcelBaseObject

@property BOOL axisBetweenCategories;  // Returns or sets if the value axis crosses the category axis between categories.
@property (readonly) ExcelE109 axisGroup;  // Returns the group for the specified axis.
@property (copy, readonly) ExcelAxisTitle *axisTitle;  // Returns an axis title object that represents the title of the specified axis.
@property ExcelE112 axisType;  // Returns or sets the axis type
@property ExcelE141 baseUnit;  // Returns or sets the base unit for the specified category axis.
@property BOOL baseUnitIsAuto;  // Returns or sets if Microsoft Excel chooses appropriate base units for the specified category axis.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property ExcelE292 categoryNames;  // Returns or sets all the category names for the specified axis, as a text array. When you set this property, you can set it to either an array or a range object that contains the category names.
@property ExcelE142 categoryType;  // Returns or sets the category axis type.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the axis.
@property ExcelE108 crosses;  // Returns or sets the point on the specified axis where the other axis crosses.
@property double crossesAt;  // Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis.
@property ExcelE139 displayUnit;  // Returns or sets the unit label for the specified axis. 
@property double displayUnitCustom;  // If the value of the display unit property is custom display unit, the display unit custom property returns or sets the value of the displayed units.
@property (copy, readonly) ExcelDisplayUnitLabel *displayUnitLabel;  // Returns the display unit label object for the specified axis
@property BOOL hasDisplayUnitLabel;  // Returns or sets if the label specified by the display unit or display unit custom property is displayed on the specified axis. The default value is true.
@property BOOL hasMajorGridlines;  // Returns or sets if the axis has major gridlines. Only axes in the primary axis group can have gridlines.
@property BOOL hasMinorGridlines;  // Returns or sets if the axis has minor gridlines. Only axes in the primary axis group can have gridlines.
@property BOOL hasTitle;  // Returns or sets if the axis or chart has a visible title.
@property (readonly) double height;  // Returns the height of the object.
@property (readonly) double leftPosition;  // Returns the left position of the specified object, in points.
@property double logBase;  // Returns or sets the base of the logarithm when you are using log scales.
@property (copy, readonly) ExcelGridlines *majorGridlines;  // Returns a gridlines object that represents the major gridlines for the specified axis. Only axes in the primary axis group can have gridlines.
@property ExcelE115 majorTickMark;  // Returns or sets the type of major tick mark for the specified axis. 
@property double majorUnit;  // Returns or sets the major units for the axis.
@property BOOL majorUnitIsAuto;  // Returns or sets if Microsoft Excel calculates the major units for the axis.
@property ExcelE141 majorUnitScale;  // Returns or sets the major unit scale value for the category axis when the category type property is set to time scale.
@property double maximumScale;  // Returns or sets the maximum value on the axis.
@property BOOL maximumScaleIsAuto;  // Returns or sets if Microsoft Excel calculates the maximum value for the axis.
@property double minimumScale;  // Returns or sets the minimum value on the axis.
@property BOOL minimumScaleIsAuto;  // Returns or sets if Microsoft Excel calculates the minimum value for the axis.
@property (copy, readonly) ExcelGridlines *minorGridlines;  // Returns a gridlines object that represents the minor gridlines for the specified axis. Only axes in the primary axis group can have gridlines.
@property ExcelE115 minorTickMark;  // Returns or sets the type of minor tick mark for the specified axis.
@property double minorUnit;  // Returns or sets the minor units on the axis.
@property BOOL minorUnitIsAuto;  // Returns or sets if Microsoft Excel calculates minor units for the axis.
@property ExcelE141 minorUnitScale;  // Returns or sets the minor unit scale value for the category axis when the category type property is set to time scale. 
@property BOOL reversePlotOrder;  // Returns or sets if Microsoft Excel plots data points from last to first.
@property ExcelE106 scaleType;  // Returns or sets the value axis scale type.
@property ExcelE122 tickLabelPosition;  // Describes the position of tick-mark labels on the specified axis.
@property NSInteger tickLabelSpacing;  // Returns or sets the number of categories or series between tick-mark labels. Applies only to category and series axes.
@property BOOL tickLabelSpacingIsAuto;  // Returns or sets whether or not the tick label spacing is automatic.
@property (copy, readonly) ExcelTickLabels *tickLabels;  // Returns a tick labels object that represents the tick-mark labels for the specified axis.
@property NSInteger tickMarkSpacing;  // Returns or sets the number of categories or series between tick marks. Applies only to category and series axes.
@property (readonly) double top;  // Returns the top position of the specified object, in points.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents the chart area of a chart. The chart area on a 2-D chart contains the axes, the chart title, the axis titles, and the legend. The chart area on a 3-D chart contains the chart title and the legend; it doesn't include the plot area .
@interface ExcelChartArea : ExcelBaseObject

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the chart area.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property double height;  // Returns or sets the height of the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (readonly) BOOL roundedCorners;  // Returns or sets if the chart area has rounded corners.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property double width;  // Returns or sets  the width of the object.

- (void) clearContents;  // Clears the data from a chart but leaves the formatting. 

@end

// Used only with charts. Represents fill formatting for chart elements.
@interface ExcelChartFillFormat : ExcelBaseObject

@property (copy, readonly) NSColor *backColor;  // Returns the background color.
@property NSInteger backgroundSchemeColor;  // Returns or sets the background color as an index in the current color scheme.
@property (readonly) ExcelMFdT chartFillFormatType;  // The fill type. 
@property (copy, readonly) NSColor *foreColor;  // Returns the foreground color.
@property NSInteger foregroundSchemeColor;  // Returns or sets the foreground color as an index in the current color scheme.
@property (readonly) ExcelMGCt gradientColorType;  // Returns the gradient color type for the specified fill.
@property (readonly) double gradientDegree;  // Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 dark through 1.0 light.
@property (readonly) ExcelMGdS gradientStyle;  // Returns the gradient style for the specified fill.
@property (readonly) NSInteger gradientVariant;  // Returns the shade variant for the specified fill as an integer value from 1 through 4.
@property (readonly) ExcelPpTy pattern;  // Returns the fill pattern.
@property (readonly) ExcelMPGb presetGradientType;  // Returns the preset gradient type for the specified fill.
@property (readonly) ExcelMPzT presetTexture;  // Returns the preset texture for the specified fill.
@property (copy, readonly) NSString *textureName;  // Returns the name of the custom texture file for the specified fill.
@property (readonly) ExcelMxtT textureType;  // Returns the texture type for the specified fill.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque, and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object is visible.


@end

// Provides access to the Office Art formatting for chart elements.
@interface ExcelChartFormat : ExcelBaseObject

@property (copy, readonly) ExcelFillFormat *fillFormat;  // Returns a fill format object for the parent chart element that contains fill formatting properties for the chart element.
@property (copy, readonly) ExcelGlowFormat *glowFormat;  // Returns a glow format object for a specified chart that contains glow formatting properties for the chart.
@property (copy, readonly) ExcelLineFormat *lineFormat;  // Returns a line format object that contains line formatting properties for the specified chart element.
@property (copy, readonly) ExcelShadowFormat *shadowFormat;  // Returns a shadow format object that contains shadow formatting properties for the chart element.
@property (copy, readonly) ExcelShapeTextFrame *shapeTextFrame;  // Returns a shape text frame object that contains the alignment and anchoring properties for the specified chart.
@property (copy, readonly) ExcelSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) ExcelThreeDFormat *threeDFormat;  // Returns a threeD format object that contains 3-D-effect formatting properties for the specified chart.


@end

// Represents one or more series plotted in a chart with the same format. A chart contains one or more chart groups, each chart group contains one or more series, and each series contains one or more points.
@interface ExcelChartGroup : ExcelBaseObject

- (SBElementArray<ExcelSeries *> *) seriesCollection;

@property ExcelE109 axisGroup;
@property NSInteger bubbleScale;  // Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from zero to 300, corresponding to a percentage of the default size. Applies only to bubble charts.
@property NSInteger doughnutHoleSize;  // Returns or sets the size of the hole in a doughnut chart group. The hole size is expressed as a percentage of the chart size, between 10 and 90 percent.
@property (copy, readonly) ExcelDownBars *downBarsObject;  // Returns a down bars object that represents the down bars on a line chart. Applies only to line charts.
@property (copy, readonly) ExcelDropLines *dropLinesObject;  // Returns a drop lines object that represents the drop lines for a series on a line chart or area chart. Applies only to line charts or area charts.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property NSInteger firstSliceAngle;  // Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees, clockwise from vertical. Applies only to pie, 3-D pie, and doughnut charts.
@property NSInteger gapWidth;  // Returns or sets: For bar and column charts, the space between bar or column clusters, as a percentage of the bar or column width. For pie of pie and bar of pie charts, the space between the primary and secondary sections of the chart.
@property BOOL hasDropLines;  // Returns or sets if the line chart or area chart has drop lines. Applies only to line and area charts.
@property BOOL hasHiLoLines;  // Returns or sets if the line chart has high-low lines. Applies only to line charts.
@property BOOL hasRadarAxisLabels;  // Returns or sets if a radar chart has axis labels. Applies only to radar charts.
@property BOOL hasSeriesLines;  // Returns or sets if a stacked column chart or bar chart has series lines or if a pie of pie chart or bar of pie chart has connector lines between the two sections. Applies only to stacked column charts, bar charts, pie of pie charts, or bar of pie charts.
@property BOOL hasThreeDShading;  // Returns or sets if the chart group has three-dimensional shading.
@property BOOL hasUpDownBars;  // Returns or sets if a line chart has up and down bars. Applies only to line charts.
@property (copy, readonly) ExcelHiloLines *hiloLinesObject;  // Returns a hilo lines object that represents the high-low lines for a series on a line chart.
@property NSInteger overlap;  // Returns or sets how bars and columns are positioned. Can be a value between -100 and 100. Applies only to 2-D bar and 2-D column charts.
@property (copy, readonly) ExcelTickLabels *radarAxisLabels;  // Returns a tick labels object that represents the radar axis labels for the specified chart group.
@property NSInteger secondPlotSize;  // Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. 
@property (copy, readonly) ExcelSeriesLines *seriesLinesObject;  // Returns a series lines object that represents the series lines for a stacked bar chart or a stacked column chart. Applies only to stacked bar and stacked column charts.
@property BOOL showNegativeBubbles;  // Returns or sets if negative bubbles are shown for the chart group. Valid only for bubble charts.
@property ExcelE146 sizeRepresents;  // Returns or sets what the bubble size represents on a bubble chart.
@property ExcelE138 splitType;  // Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split.
@property NSInteger splitValue;  // Returns or sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart.
@property (copy, readonly) ExcelUpBars *upBarsObject;  // Returns an up bars object that represents the up bars on a line chart. Applies only to line charts.
@property BOOL varyByCategories;  // Returns or sets if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series.


@end

@interface ExcelAreaGroup : ExcelChartGroup


@end

@interface ExcelBarGroup : ExcelChartGroup


@end

// Represents an embedded chart on a worksheet. The chart object object acts as a container for a chart object. Properties and methods for the chart object object control the appearance and size of the embedded chart on the worksheet.
@interface ExcelChartObject : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns a range object that represents the cell that lies under the lower-right corner of the object.
@property (copy, readonly) ExcelChart *chart;  // Returns a chart object that represents the chart contained in the object.
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or set the height of the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelE210 placement;  // Returns or sets the way the object is attached to the cells below it.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property BOOL protectChartObject;  // Returns or sets if the embedded chart frame cannot be moved, resized, or deleted.
@property BOOL roundedCorners;  // Returns or sets if the embedded chart has rounded corners.
@property BOOL shadow;  // Returns or sets if the if the object has a shadow.
@property double top;  // Returns or sets the distance from the top edge of the object to the top of row 1, on a worksheet, or the top of the chart area, on a chart.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.

- (void) bringToFront;  // Brings the object to the front of the z-order.
- (void) copyPictureAppearance:(ExcelE207)appearance format:(ExcelE156)format NS_RETURNS_NOT_RETAINED;  // Copies the selected object to the clipboard as a picture.
- (void) cut;  // Cuts the object to the clipboard.
- (void) saveAsPicturePictureType:(ExcelMPiT)pictureType fileName:(NSString *)fileName;  // Saves the shape in the requested file using the stated graphic format
- (void) sendToBack;  // Sends the object to the back of the z-order.

@end

// Represents the chart title.
@interface ExcelChartTitle : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *caption;  // Returns or sets the chart title text.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the chart title.
@property (copy) NSString *chartTitleText;  // Returns or sets the text for the specified object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property (copy) NSString *formulaLocal;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formulaR1c1;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the macro.
@property (copy) NSString *formulaR1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property (readonly) double height;  // Returns the height of the object.
@property ExcelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property BOOL includeInLayout;  // Returns or sets if a chart title will occupy the chart layout space when a chart layout is being determined.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property ExcelE126 orientation;  // May also be a number value from -90 to 90 degrees.
@property ExcelE216 position;  // Returns or sets the position of the chart title on the chart.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property ExcelE284 verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents a chart in a workbook. The chart can be either an embedded chart, contained in a chart object, or a separate chart sheet.
@interface ExcelChart : ExcelBaseObject

- (SBElementArray<ExcelShape *> *) shapes;
- (SBElementArray<ExcelArc *> *) arcs;
- (SBElementArray<ExcelAreaGroup *> *) areaGroups;
- (SBElementArray<ExcelBarGroup *> *) barGroups;
- (SBElementArray<ExcelButton *> *) buttons;
- (SBElementArray<ExcelChartGroup *> *) chartGroups;
- (SBElementArray<ExcelChartObject *> *) chartObjects;
- (SBElementArray<ExcelCheckbox *> *) checkboxes;
- (SBElementArray<ExcelColumnGroup *> *) columnGroups;
- (SBElementArray<ExcelDoughnutGroup *> *) doughnutGroups;
- (SBElementArray<ExcelDropdown *> *) dropdowns;
- (SBElementArray<ExcelGroupbox *> *) groupboxes;
- (SBElementArray<ExcelHyperlink *> *) hyperlinks;
- (SBElementArray<ExcelLabel *> *) labels;
- (SBElementArray<ExcelLineGroup *> *) lineGroups;
- (SBElementArray<ExcelLine *> *) lines;
- (SBElementArray<ExcelListbox *> *) listboxes;
- (SBElementArray<ExcelOptionButton *> *) optionButtons;
- (SBElementArray<ExcelOval *> *) ovals;
- (SBElementArray<ExcelPieGroup *> *) pieGroups;
- (SBElementArray<ExcelRadarGroup *> *) radarGroups;
- (SBElementArray<ExcelRectangle *> *) rectangles;
- (SBElementArray<ExcelScrollbar *> *) scrollbars;
- (SBElementArray<ExcelSeries *> *) seriesCollection;
- (SBElementArray<ExcelSpinner *> *) spinners;
- (SBElementArray<ExcelTextbox *> *) textboxes;
- (SBElementArray<ExcelXyGroup *> *) xyGroups;

@property (copy, readonly) ExcelChartGroup *areaThreeDGroup;  // Returns a chart group object that represents the area chart group on a 3-D chart.
@property BOOL autoScaling;  // Returns or sets if Microsoft Excel scales a 3-D chart so that it's closer in size to the equivalent 2-D chart. The chart's right angle axes property must be true.
@property (copy, readonly) ExcelWalls *backWall;  // Returns a walls object that allows the user to individually format the back wall of a 3-D chart.
@property ExcelE143 barShape;  // Returns or sets the shape used with the 3-D bar or column chart.
@property (copy, readonly) ExcelChartGroup *barThreeDGroup;  // Returns a chart group object that represents the bar chart group on a 3-D chart.
@property (copy, readonly) ExcelChartArea *chartAreaObject;  // Returns a chart area object that represents the complete chart area for the chart.
@property NSInteger chartStyle;  // Returns or sets the chart type.
@property (copy, readonly) ExcelChartTitle *chartTitle;  // Returns a chart title object that represents the title of the specified chart.
@property ExcelE144 chartType;  // Returns or sets the chart type.
@property (copy, readonly) ExcelChartGroup *columnThreeDGroup;  // Returns a chart group object that represents the column chart group on a 3-D chart.
@property (copy, readonly) ExcelCorners *cornersObject;  // Returns a corners object that represents the corners of a 3-D chart.
@property (copy, readonly) ExcelDataTable *dataTableObject;  // Returns a data table object that represents the chart data table.
@property NSInteger depthPercent;  // Returns or sets the depth of a 3-D chart as a percentage of the chart width, between 20 and 2000 percent.
@property ExcelE118 displayBlanksAs;  // Returns or sets the way that blank cells are plotted on a chart.
@property NSInteger elevation;  // Returns or sets the elevation of the 3-D chart view, in degrees.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) ExcelFloor *floorObject;  // Returns a floor object that represents the floor of the 3-D chart.
@property NSInteger gapDepth;  // Returns or sets the distance between the data series in a 3-D chart, as a percentage of the marker width. The value of this property must be between 0 and 500.
@property BOOL hasDataTable;  // Returns or sets if the chart has a data table.
@property BOOL hasLegend;  // Returns or sets if the chart has a legend.
@property BOOL hasTitle;  // Returns or sets if the chart has a title.
@property NSInteger heightPercent;  // Returns or sets the height of a 3-D chart as a percentage of the chart width, between 5 and 500 percent.
@property (copy, readonly) ExcelLegend *legendObject;  // Returns a legend object that represents the legend for the chart.
@property (copy, readonly) ExcelChartGroup *lineThreeDGroup;  // Returns a chart group object that represents the line chart group on a 3-D chart.
@property (copy) NSString *name;  // Returns or sets the name of the chart.
@property (copy, readonly) ExcelChart *next;  // Returns a worksheet object that represents the next sheet.
@property (copy, readonly) ExcelPageSetup *pageSetupObject;  // Returns the page setup object associated with this chart.
@property NSInteger perspective;  // Returns or sets the perspective for the 3-D chart view. Must be between 0 and 100. This property is ignored if the right angle axes property is true.
@property (copy, readonly) ExcelChartGroup *pieThreeDGroup;  // Returns a chart group object that represents the pie chart group on a 3-D chart.
@property (copy, readonly) ExcelPlotArea *plotAreaObject;  // Returns a plot area object that represents the plot area of a chart.
@property ExcelE105 plotBy;  // Returns or sets the way columns or rows are used as data series on the chart.
@property BOOL plotVisibleOnly;  // Returns or sets if only visible cells are plotted. False if both visible and hidden cells are plotted.
@property (copy, readonly) id previous;  // Returns a worksheet object that represents the previous sheet.
@property (readonly) BOOL protectContents;  // Returns true if the contents of the sheet are protected.
@property BOOL protectData;  // Returns or sets if series formulas cannot be modified by the user.
@property (readonly) BOOL protectDrawingObjects;  // Returns true if shapes are protected.
@property BOOL protectFormatting;  // Returns or sets if chart formatting cannot be modified by the user.
@property BOOL protectGoalSeek;  // Returns or sets if the user cannot modify chart data points with mouse actions.
@property BOOL protectSelection;  // Returns or sets if chart elements cannot be selected.
@property (readonly) BOOL protectionMode;  // Returns true if user-interface-only protection is turned on. To turn on user interface protection, use the protect method with the user interface only argument set to true.
@property BOOL rightAngleAxes;  // Returns or sets if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3-D line, column, and bar charts.
@property NSInteger rotation;  // The rotation of the 3D chart view.  The value of must be from 0 to 360.
@property (copy, readonly) ExcelTab *sheetTab;  // Returns the sheet tab of the chart sheet
@property BOOL showDataLabelsOverMaximum;  // Returns or sets whether to show the data labels when the value is greater than the maximum value on the value axis.
@property BOOL showWindow;  // Returns or sets if the embedded chart is displayed in a separate window. The Chart object used with this property must refer to an embedded chart.
@property (copy, readonly) ExcelWalls *sideWall;  // Returns a walls object that allows the user to individually format the side wall of a 3-D chart.
@property BOOL sizeWithWindow;  // Returns or sets if Microsoft Excel resizes the chart to match the size of the chart sheet window. False if the chart size isn't attached to the window size. Applies only to chart sheets, doesn't apply to embedded charts.
@property (copy, readonly) ExcelChartGroup *surfaceGroup;  // Returns a chart group object that represents the surface chart group of a 3-D chart.
@property ExcelE225 visible;  // Returns or sets if the chart is visible.
@property BOOL wallsAndGridlinesTwoD;  // Returns or sets if gridlines are drawn two-dimensionally on a 3-D chart.
@property (copy, readonly) ExcelWalls *wallsObject;  // Returns a walls object that represents the walls of the 3-D chart.

- (void) applyCustomChartTypeChartType:(ExcelE102)chartType chartName:(NSString *)chartName;  // Applies a standard or custom chart type to a chart or series.
- (void) applyLayoutLayout:(NSInteger)layout chartType:(ExcelE144)chartType;  // Applies a chart layout.
- (ExcelChart *) chartLocationWhere:(ExcelE201)where name:(NSString *)name;  // Moves the chart to a new location.
- (void) chartWizardSource:(ExcelRange *)source gallery:(ExcelE144)gallery format:(NSInteger)format plotBy:(ExcelE105)plotBy categoryLabels:(NSInteger)categoryLabels seriesLabels:(NSInteger)seriesLabels hasLegend:(BOOL)hasLegend title:(NSString *)title categoryTitle:(NSString *)categoryTitle valueTitle:(NSString *)valueTitle extraTitle:(NSString *)extraTitle;  // Modifies the properties of the given chart. You can use this method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties.
- (void) checkSpellingCustomDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase alwaysSuggest:(BOOL)alwaysSuggest;  // Checks the spelling of an object.
- (void) clearToMatchStyle;  // Sets the chart formatting to the last predefined chart style applied.
- (void) copyChartAsPictureAppearance:(ExcelE207)appearance format:(ExcelE156)format outputSize:(ExcelE207)outputSize NS_RETURNS_NOT_RETAINED;  // Copies the selected object to the clipboard as a picture.
- (void) deselect;  // Cancels the selection for the specified chart.
- (ExcelAxis *) getAxisAxisType:(ExcelE112)axisType whichAxis:(ExcelE109)whichAxis;  // Returns an object that represents a specified axis object.
- (NSArray *) getChartElementX:(NSInteger)x y:(NSInteger)y;  // Returns information about the chart element at specified X and Y coordinates. This method returns a list  of three items.  Please refer to the documentation on the meaning of the returned values.
- (BOOL) getHasAxisAxisType:(ExcelE112)axisType axisGroup:(ExcelE109)axisGroup;  // Returns which axes exist on the chart.
- (void) pasteChartFormat:(ExcelE204)format;  // Pastes chart data from the clipboard into the specified chart.
- (void) printOutFrom:(NSInteger)from to:(NSInteger)to copies:(NSInteger)copies preview:(BOOL)preview activePrinter:(NSString *)activePrinter printToFile:(BOOL)printToFile collate:(BOOL)collate;  // Prints the object
- (void) printPreviewEnableChanges:(BOOL)enableChanges;  // Shows a preview of the object as it would look when printed. This function is deprecated.
- (void) protectChartPassword:(NSString *)password drawingObjects:(BOOL)drawingObjects chartContents:(BOOL)chartContents userInterfaceOnly:(BOOL)userInterfaceOnly;  // Protects a chart so that it cannot be modified.
- (void) refresh;  // Updates the cache of the chart object.
- (void) saveAsFilename:(NSString *)filename fileFormat:(ExcelE188)fileFormat password:(NSString *)password writeReservationPassword:(NSString *)writeReservationPassword readOnlyRecommended:(BOOL)readOnlyRecommended createBackup:(BOOL)createBackup addToMostRecentlyUsedList:(BOOL)addToMostRecentlyUsedList saveAsLocalLanguage:(BOOL)saveAsLocalLanguage;  // Saves changes into a different file.
- (void) setBackgroundPicturePictureFileName:(NSString *)pictureFileName;  // Sets the background graphic for a chart.
- (void) setChartElementChartElement:(id)chartElement;  // Sets chart elements on a chart.
- (void) setHasAxisAxisExists:(BOOL)axisExists axisType:(ExcelE112)axisType axisGroup:(ExcelE109)axisGroup;  // Sets which axes exist on the chart.
- (void) setSourceDataSource:(ExcelRange *)source plotBy:(ExcelE105)plotBy;  // Sets the source data range for the chart.
- (void) unprotectPassword:(NSString *)password;  // Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.

@end

// A chart sheet is a worksheet that contains only a chart.
@interface ExcelChartSheet : ExcelChart

@property (copy, readonly) ExcelChart *chart;  // Returns the chart for this chart sheet.
@property (readonly) ExcelE151 worksheetType;  // Returns the type of this worksheet.


@end

@interface ExcelColumnGroup : ExcelChartGroup


@end

// Represents the corners of a 3-D chart.
@interface ExcelCorners : ExcelBaseObject

@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the data label on a chart point or trendline.
@interface ExcelDataLabel : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property BOOL autoText;  // Returns or sets if the object automatically generates appropriate text based on context.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *caption;  // Returns or sets the caption of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the data label.
@property (copy) NSString *dataLabelText;  // Returns or sets the text for this object.
@property ExcelE134 dataLabelType;  // Returns or sets the type of the data label.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property (copy) NSString *formulaLocal;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formulaR1c1;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the macro.
@property (copy) NSString *formulaR1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property (readonly) double height;  // Returns the height of the object.
@property ExcelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property BOOL numberFormatLinked;  // Returns or sets if the number format is linked to the cells, so that the number format changes in the labels when it changes in the cells.
@property (copy) NSString *numberFormatLocal;  // Returns or sets the format code for the object as a string in the language of the user.
@property NSInteger orientation;  // The text orientation.  Can be a number value from -90 to 90 degrees.
@property ExcelE140 position;  // Returns or sets the position of the specified object.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property BOOL showLegendKey;  // Returns or sets if the data label legend key is visible.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property ExcelE284 verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents a chart data table.
@interface ExcelDataTable : ExcelBaseObject

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the data table.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property BOOL hasBorderHorizontal;  // Returns or sets if the chart data table has horizontal cell borders. 
@property BOOL hasBorderOutline;  // Returns or sets if the chart data table has outline borders.
@property BOOL hasBorderVertical;  // Returns or sets if the chart data table has vertical cell borders. 
@property BOOL showLegendKey;  // Returns or sets if the data label legend key is visible.


@end

// Represents a unit label on an axis in the specified chart. Unit labels are useful for charting large values-- for example, in the millions or billions.
@interface ExcelDisplayUnitLabel : ExcelBaseObject

- (SBElementArray<ExcelCharacter *> *) characters;

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *caption;  // Returns or sets the caption for the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the chart title.
@property (copy) NSString *displayLabelUnitText;  // Returns or sets the text for this object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property ExcelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property NSInteger orientation;  // The text orientation.  Can be a number value from -90 to 90 degrees.
@property ExcelE216 position;  // Returns or sets the position of the chart title on the chart.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property ExcelE284 verticalAlignment;  // Returns or sets the vertical alignment of the object.


@end

@interface ExcelDoughnutGroup : ExcelChartGroup


@end

// Represents the down bars in a chart group. Down bars connect points on the first series in the chart group with lower values on the last series, the lines go down from the first series.
@interface ExcelDownBars : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the down bars.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy, readonly) NSString *name;  // Returns the name of the object.


@end

// Represents the drop lines in a chart group. Drop lines connect the points in the chart with the x-axis. Only line and area chart groups can have drop lines.
@interface ExcelDropLines : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the drop lines.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the error bars on a chart series. Error bars indicate the degree of uncertainty for chart data.
@interface ExcelErrorBars : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the error bars.
@property ExcelE104 endStyle;  // Returns or sets the end style for the error bars.
@property (copy, readonly) NSString *name;  // Returns the name of the object.


@end

// Represents the floor of a 3-D chart.
@interface ExcelFloor : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the floor.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property ExcelE124 pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property NSInteger thickness;  // Returns or sets  the thickness of the floor.


@end

// Represents major or minor gridlines on a chart axis. Gridlines extend the tick marks on a chart axis to make it easier to see the values associated with the data markers.
@interface ExcelGridlines : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the gridlines.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the high-low lines in a chart group. High-low lines connect the highest point with the lowest point in every category in the chart group. Only 2-D line groups can have high-low lines.
@interface ExcelHiloLines : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the hilo lines.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the interior of an object.
@interface ExcelInterior : ExcelBaseObject

@property (copy) NSColor *color;  // Returns or sets the primary color of the object.
@property ExcelE103 colorIndex;  // Returns or sets the color of the interior. The color is specified as an index value into the current color palette.
@property BOOL invertIfNegative;  // Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.
@property (copy, readonly) ExcelLinearGradient *linearGradient;  // Returns or sets the Gradient property of an Interior object of a selection.
@property ExcelE137 pattern;  // Returns or sets the interior pattern.
@property (copy) NSColor *patternColor;  // Returns or sets the color of the interior pattern as an RGB value.
@property ExcelE103 patternColorIndex;  // Returns or sets the color of the interior pattern as an index into the current color palette.
@property ExcelE305 patternThemeColor;  // Returns or sets a theme color pattern for an Interior object.
@property double patternTintAndShade;  // Returns or sets a tint and shade pattern for an Interior object.
@property (copy, readonly) ExcelRectangularGradient *rectangularGradient;  // Returns or sets the Gradient property of an Interior object of a selection.
@property ExcelE305 themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.


@end

// Represents leader lines on a chart. Leader lines connect data labels to data points.
@interface ExcelLeaderLines : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the leader lines.


@end

// Represents a legend entry in a chart legend.
@interface ExcelLegendEntry : ExcelBaseObject

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the legend entry.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (readonly) double height;  // Returns the height of the object.
@property (readonly) double leftPosition;  // Returns the left position of the specified object, in points.
@property (copy, readonly) ExcelLegendKey *legendKey;  // Returns a legend key object that represents the legend key associated with the entry.
@property (readonly) double top;  // Returns the top position of the specified object, in points.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents a legend key in a chart legend. Each legend key is a graphic that visually links a legend entry with its associated series or trendline in the chart.
@interface ExcelLegendKey : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the legend key.
@property (readonly) double height;  // Returns the height of the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property BOOL invertIfNegative;  // Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.
@property (readonly) double leftPosition;  // Returns the left position of the specified object, in points.
@property (copy) NSColor *markerBackgroundColor;  // Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts.
@property ExcelE103 markerBackgroundColorIndex;  // Returns or sets the marker background color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property (copy) NSColor *markerForegroundColor;  // Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts.
@property ExcelE103 markerForegroundColorIndex;  // Returns or sets the marker foreground color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property NSInteger markerSize;  // Returns or sets the data-marker size, in points.
@property ExcelE135 markerStyle;  // Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart.
@property ExcelE124 pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property double pictureUnit;  // Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property BOOL smooth;  // Returns or sets if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts.
@property (readonly) double top;  // Returns the top position of the specified object, in points.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents the legend in a chart. Each chart can have only one legend.
@interface ExcelLegend : ExcelBaseObject

- (SBElementArray<ExcelLegendEntry *> *) legendEntries;

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the legend.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property double height;  // Returns or sets the height of the object.
@property BOOL includeInLayout;  // Returns or sets if a legend will occupy the chart layout space when a chart layout is being determined.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property ExcelE123 position;  // Returns or sets the position of the legend on the chart.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property double width;  // Returns or sets  the width of the object.


@end

@interface ExcelLineGroup : ExcelChartGroup


@end

@interface ExcelPieGroup : ExcelChartGroup


@end

// Represents the plot area of a chart. This is the area where your chart data is plotted.
@interface ExcelPlotArea : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the plot area.
@property double height;  // Returns or sets the height of the object.
@property (readonly) double insideHeight;  // Returns the inside height of the plot area, in points.
@property (readonly) double insideLeft;  // Returns the distance from the chart edge to the inside left edge of the plot area, in points.
@property (readonly) double insideTop;  // Returns the distance from the chart edge to the inside top edge of the plot area, in points.
@property (readonly) double insideWidth;  // Returns the inside width of the plot area, in points.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property ExcelE216 position;  // Returns or sets the position of the plot area on the chart.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property double width;  // Returns or sets  the width of the object.


@end

@interface ExcelRadarGroup : ExcelChartGroup


@end

// Represents series lines in a chart group. Series lines connect the data values from each series. Only 2-D stacked bar or column chart groups can have series lines.
@interface ExcelSeriesLines : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the series lines.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents a single point in a series in a chart.
@interface ExcelSeriesPoint : ExcelBaseObject

@property BOOL applyPictToEnd;  // Returns or sets if a picture is applied to the end of the point or all points in the series.
@property BOOL applyPictToFront;  // Returns or sets if a picture is applied to the front of the point or all points in the series.
@property BOOL applyPictToSides;  // Returns or sets if a picture is applied to the sides of the point or all points in the series.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the point.
@property (copy, readonly) ExcelDataLabel *dataLabelObject;  // Returns a data label object that represents the data label associated with the point or trendline.
@property NSInteger explosion;  // Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns zero if there's no explosion, the tip of the slice is in the center of the pie.
@property BOOL hasDataLabel;  // Returns or sets if the point has a data label.
@property BOOL hasThreeDEffect;  // Returns or sets if a point has a three-dimensional appearance. Applies only to bubble charts.
@property (readonly) double height;  // Returns the height of the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property BOOL invertIfNegative;  // Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.
@property (readonly) double leftPosition;  // Returns the left position of the specified object, in points.
@property (copy) NSColor *markerBackgroundColor;  // Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts.
@property ExcelE103 markerBackgroundColorIndex;  // Returns or sets the marker background color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property (copy) NSColor *markerForegroundColor;  // Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts.
@property ExcelE103 markerForegroundColorIndex;  // Returns or sets the marker foreground color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property NSInteger markerSize;  // Returns or sets the data-marker size, in points.
@property ExcelE135 markerStyle;  // Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property ExcelE124 pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property double pictureUnit;  // Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored.
@property BOOL secondaryPlot;  // Returns or sets if the point is in the secondary section of either a pie of pie chart or a bar of pie chart. Applies only to points on pie of pie charts or bar of pie charts.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property (readonly) double top;  // Returns the top position of the specified object, in points.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents a series in a chart.
@interface ExcelSeries : ExcelBaseObject

- (SBElementArray<ExcelDataLabel *> *) dataLabels;
- (SBElementArray<ExcelSeriesPoint *> *) seriesPoints;
- (SBElementArray<ExcelTrendline *> *) trendlines;

@property BOOL applyPictureToEnd;  // Returns or sets if a picture is applied to the end of the point or all points in the series.
@property BOOL applyPictureToFront;  // Returns or sets if a picture is applied to the front of the point or all points in the series.
@property BOOL applyPictureToSides;  // Returns or sets if a picture is applied to the sides of the point or all points in the series.
@property ExcelE109 axisGroup;  // Returns or sets the axis group for the specified series.
@property ExcelE143 barShape;  // Returns or sets the shape used with the 3-D bar or column chart.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *bubbleSizes;  // Returns or sets a string in A1-style notation that refers to the worksheet cells containing the size data for the bubble chart. Applies only to bubble charts. 
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the series.
@property ExcelE144 chartType;  // Returns or sets the chart type.
@property (copy, readonly) ExcelErrorBars *errorBars;  // Returns an error bars object that represents the error bars for the series.
@property NSInteger explosion;  // Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns zero if there's no explosion, the tip of the slice is in the center of the pie.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property (copy) NSString *formulaLocal;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formulaR1c1;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the macro.
@property (copy) NSString *formulaR1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property BOOL hasDataLabels;  // Returns or sets if the series has data labels.
@property BOOL hasErrorBars;  // Returns or set if the series has error bars. This property isn't available for 3-D charts.
@property BOOL hasLeaderLines;  // Returns or sets if the series has leader lines.
@property BOOL hasThreeDEffect;  // Returns or sets if the series has a three-dimensional appearance. Applies only to bubble charts.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property BOOL invertIfNegative;  // Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.
@property (copy, readonly) ExcelLeaderLines *leaderLines;  // Returns a leader lines object that represents the leader lines for the series.
@property (copy) NSColor *markerBackgroundColor;  // Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts.
@property ExcelE103 markerBackgroundColorIndex;  // Returns or sets the marker background color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property (copy) NSColor *markerForegroundColor;  // Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts.
@property ExcelE103 markerForegroundColorIndex;  // Returns or sets the marker foreground color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property NSInteger markerSize;  // Returns or sets the data-marker size, in points.
@property ExcelE135 markerStyle;  // Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property ExcelE124 pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property double pictureUnit;  // Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored.
@property (readonly) NSInteger plotColorIndex;  // Returns the plot color index of the series.
@property NSInteger plotOrder;  // Returns or sets the plot order for the selected series within the chart group.
@property ExcelE292 seriesValues;  // Returns or sets a list of all the values in the series. This can be a range on a worksheet or a list of constant values.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property BOOL smooth;  // Returns or sets if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts.
@property ExcelE292 xvalues;  // Returns or sets a list of x values for a chart series. The xvalues property can be set to a range on a worksheet or to a list of values.

- (void) errorBarDirection:(ExcelE116)direction include:(ExcelE117)include type:(ExcelE131)type amount:(double)amount minusValues:(double)minusValues;  // Applies error bars to the series.
- (void) pasteSeries;

@end

// Represents the tick-mark labels associated with tick marks on a chart axis.
@interface ExcelTickLabels : ExcelBaseObject

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the tick labels.
@property (readonly) NSInteger depth;  // Returns the number of levels of category tick labels.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property BOOL multiLevel;  // Returns or sets whether an axis is multilevel or not.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property BOOL numberFormatLinked;  // Returns or sets if the number format is linked to the cells, so that the number format changes in the labels when it changes in the cells.
@property (copy) NSString *numberFormatLocal;  // Returns or sets the format code for the object as a string in the language of the user.
@property NSInteger offset;  // Returns or sets the distance between the levels of labels, and the distance between the first level and the axis line. The value can be an integer percentage from 0 through 1000, relative to the axis label's font size. 
@property ExcelE127 orientation;  // The text orientation.  Can be a number value from -90 to 90 degrees.
@property ExcelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property ExcelE299 tickAlignment;  // Returns or sets the alignment for the specified tick label.


@end

// Represents a trendline in a chart. A trendline shows the trend, or direction, of data in a series.
@interface ExcelTrendline : ExcelBaseObject

@property double backward;  // Returns or sets the number of periods, or units on a scatter chart, that the trendline extends backward.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the trendline.
@property (copy, readonly) ExcelDataLabel *dataLabelObject;  // Returns a data label object that represents the data label associated with the point or trendline.
@property BOOL displayRSquared;  // Returns or sets if the R-squared value of the trendline is displayed on the chart, in the same data label as the equation. Setting this property to true automatically turns on data labels.
@property BOOL displayEquation;  // Returns or sets if the equation for the trendline is displayed on the chart, in the same data label as the R-squared value. Setting this property to true automatically turns on data labels.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double forward;  // Returns or sets the number of periods, or units on a scatter chart, that the trendline extends forward.
@property double intercept;  // Returns or sets the point where the trendline crosses the value axis.
@property BOOL interceptIsAuto;  // Returns or sets if the point where the trendline crosses the value axis is automatically determined by the regression.
@property (copy) NSString *name;  // Returns or sets the name of this object.
@property BOOL nameIsAuto;  // Returns or sets if Microsoft Excel automatically determines the name of the trendline.
@property NSInteger order;  // Returns or sets the trendline order, an integer greater than 1,  when the trendline type is polynomial.
@property NSInteger period;  // Returns or sets the period for the moving-average trendline.
@property ExcelE132 trendlineType;  // Returns or sets the type of this trend line


@end

// Represents the up bars in a chart group. Up bars connect points on series one with higher values on the last series in the chart group, the lines go up from series one. Only 2-D line groups that contain at least two series can have up bars.
@interface ExcelUpBars : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the up bars.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the walls of a 3-D chart.
@interface ExcelWalls : ExcelBaseObject

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the walls.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy, readonly) NSString *name;  // Returns or sets the name of the object.
@property ExcelE124 pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property NSInteger pictureUnit;  // Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored.
@property NSInteger thickness;  // Returns or sets  the thickness of the floor.


@end

@interface ExcelXyGroup : ExcelChartGroup


@end

