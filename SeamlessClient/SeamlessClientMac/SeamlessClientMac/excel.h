/*
 * excel.h
 */

#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>


@class excelBaseObject, excelBaseApplication, excelBaseDocument, excelBasicWindow, excelPrintSettings, excelCommandBarControl, excelCommandBarButton, excelCommandBarCombobox, excelCommandBarPopup, excelCommandBar, excelDocumentProperty, excelCustomDocumentProperty, excelWebPageFont, excelExcelComment, excelODBCError, excelProtection, excelAboveAverageFormatCondition, excelAddIn, excelApplication, excelAutofilter, excelBorder, excelButton, excelCalculatedMember, excelCheckbox, excelColorScaleCriteria, excelColorScaleCriterion, excelColorScaleFormatCondition, excelColorstop, excelColorstops, excelConditionValue, excelCubeField, excelCustomView, excelDatabarBorder, excelDatabarFormatCondition, excelDefaultWebOptions, excelDialog, excelDisplayFormat, excelDropdown, excelFilter, excelFormatColor, excelFormatConditionIconObject, excelFormatConditionIconSet, excelFormatConditionIconSets, excelFormatCondition, excelGraphic, excelGroupbox, excelHorizontalPageBreak, excelHyperlink, excelIconCriteria, excelIconCriterion, excelIconSetFormatCondition, excelLabel, excelLinearGradient, excelListColumn, excelListObject, excelListRow, excelListbox, excelNamedItem, excelNegativeBarFormat, excelOptionButton, excelOutline, excelPageSetup, excelPane, excelPhonetic, excelPivotAxis, excelPivotCache, excelPivotCell, excelPivotField, excelCalculatedField, excelColumnField, excelDataField, excelHiddenField, excelPageField, excelPivotFilter, excelActiveFilter, excelPivotFormula, excelPivotItem, excelCalculatedItem, excelChildItem, excelColumnItem, excelHiddenItem, excelParentItem, excelPivotLine, excelPivotTable, excelQueryTable, excelRecentFile, excelRectangularGradient, excelRowField, excelRowItem, excelScenario, excelScrollbar, excelSheet, excelInternationalMacroSheet, excelMacroSheet, excelSlicer, excelSort, excelSortfield, excelSortfields, excelSpinner, excelTab, excelTableStyleElement, excelTableStyle, excelTextbox, excelTop10FormatCondition, excelTreeviewControl, excelUniqueValuesFormatCondition, excelValidation, excelValueChange, excelVerticalPageBreak, excelWebOptions, excelWindow, excelWorkbookConnection, excelWorkbook, excelDocument, excelWorksheet, excelXlspellingOptions, excelAdjustment, excelArc, excelBulletFormat, excelCalloutFormat, excelConnectorFormat, excelFillFormat, excelGlowFormat, excelGradientStop, excelLineFormat, excelLine, excelOfficeTheme, excelOval, excelParagraphFormat, excelPictureFormat, excelRectangle, excelReflectionFormat, excelRulerLevel, excelRuler, excelShadowFormat, excelShapeFont, excelShapeTextFrame, excelShape, excelCallout, excelPicture, excelShapeConnector, excelShapeLine, excelShapeTextbox, excelSoftEdgeFormat, excelTabStop, excelTextColumn, excelTextFrame, excelThemeColorScheme, excelThemeColor, excelThemeEffectScheme, excelThemeFontScheme, excelThemeFont, excelMajorThemeFont, excelMinorThemeFont, excelThreeDFormat, excelWordArtFormat, excelCharacter, excelFont, excelStyle, excelTextRange, excelParagraph, excelSentence, excelTextFlow, excelTextRangeCharacter, excelTextRangeLine, excelWord, excelRange, excelCell, excelColumn, excelRow, excelAutocorrect, excelAxisTitle, excelAxis, excelChartArea, excelChartFillFormat, excelChartFormat, excelChartGroup, excelAreaGroup, excelBarGroup, excelChartObject, excelChartTitle, excelChart, excelChartSheet, excelColumnGroup, excelCorners, excelDataLabel, excelDataTable, excelDisplayUnitLabel, excelDoughnutGroup, excelDownBars, excelDropLines, excelErrorBars, excelFloor, excelGridlines, excelHiloLines, excelInterior, excelLeaderLines, excelLegendEntry, excelLegendKey, excelLegend, excelLineGroup, excelPieGroup, excelPlotArea, excelRadarGroup, excelSeriesLines, excelSeriesPoint, excelSeries, excelTickLabels, excelTrendline, excelUpBars, excelWalls, excelXyGroup;

enum excelSavo {
	excelSavoYes = 'yes ' /* Save objects now */,
	excelSavoNo = 'no  ' /* Do not save objects */,
	excelSavoAsk = 'ask ' /* Ask the user whether to save */
};
typedef enum excelSavo excelSavo;

enum excelKfrm {
	excelKfrmIndex = 'indx' /* keyform designating indexed access */,
	excelKfrmNamed = 'name' /* keyform designating named access */,
	excelKfrmId = 'ID  ' /* keyform designating access by unique identifier */
};
typedef enum excelKfrm excelKfrm;

enum excelXLfd {
	excelXLfdMacintoshPath = 'mPth' /* Maintosh path i.e. Foo:bar.txt */,
	excelXLfdPosixPath = 'file' /* Posix path i.e. file:://foo/bar.txt */
};
typedef enum excelXLfd excelXLfd;

enum excelEnum {
	excelEnumStandard = 'lwst' /* Standard PostScript error handling   */,
	excelEnumDetailed = 'lwdt' /* print a detailed report of Postscript errors */
};
typedef enum excelEnum excelEnum;

enum excelMlDs {
	excelMlDsLineDashStyleUnset = '\000\222\377\376',
	excelMlDsLineDashStyleSolid = '\000\223\000\001',
	excelMlDsLineDashStyleSquareDot = '\000\223\000\002',
	excelMlDsLineDashStyleRoundDot = '\000\223\000\003',
	excelMlDsLineDashStyleDash = '\000\223\000\004',
	excelMlDsLineDashStyleDashDot = '\000\223\000\005',
	excelMlDsLineDashStyleDashDotDot = '\000\223\000\006',
	excelMlDsLineDashStyleLongDash = '\000\223\000\007',
	excelMlDsLineDashStyleLongDashDot = '\000\223\000\010',
	excelMlDsLineDashStyleLongDashDotDot = '\000\223\000\011',
	excelMlDsLineDashStyleSystemDash = '\000\223\000\012',
	excelMlDsLineDashStyleSystemDot = '\000\223\000\013',
	excelMlDsLineDashStyleSystemDashDot = '\000\223\000\014'
};
typedef enum excelMlDs excelMlDs;

enum excelMLnS {
	excelMLnSLineStyleUnset = '\000\224\377\376',
	excelMLnSSingleLine = '\000\225\000\001',
	excelMLnSThinThinLine = '\000\225\000\002',
	excelMLnSThinThickLine = '\000\225\000\003',
	excelMLnSThickThinLine = '\000\225\000\004',
	excelMLnSThickBetweenThinLine = '\000\225\000\005'
};
typedef enum excelMLnS excelMLnS;

enum excelMAhS {
	excelMAhSArrowheadStyleUnset = '\000\221\377\376',
	excelMAhSNoArrowhead = '\000\222\000\001',
	excelMAhSTriangleArrowhead = '\000\222\000\002',
	excelMAhSOpen_arrowhead = '\000\222\000\003',
	excelMAhSStealthArrowhead = '\000\222\000\004',
	excelMAhSDiamondArrowhead = '\000\222\000\005',
	excelMAhSOvalArrowhead = '\000\222\000\006'
};
typedef enum excelMAhS excelMAhS;

enum excelMAhW {
	excelMAhWArrowheadWidthUnset = '\000\220\377\376',
	excelMAhWNarrowWidthArrowhead = '\000\221\000\001',
	excelMAhWMediumWidthArrowhead = '\000\221\000\002',
	excelMAhWWideArrowhead = '\000\221\000\003'
};
typedef enum excelMAhW excelMAhW;

enum excelMAhL {
	excelMAhLArrowheadLengthUnset = '\000\223\377\376',
	excelMAhLShortArrowhead = '\000\224\000\001',
	excelMAhLMediumArrowhead = '\000\224\000\002',
	excelMAhLLongArrowhead = '\000\224\000\003'
};
typedef enum excelMAhL excelMAhL;

enum excelMFdT {
	excelMFdTFillUnset = '\000c\377\376',
	excelMFdTFillSolid = '\000d\000\001',
	excelMFdTFillPatterned = '\000d\000\002',
	excelMFdTFillGradient = '\000d\000\003',
	excelMFdTFillTextured = '\000d\000\004',
	excelMFdTFillBackground = '\000d\000\005',
	excelMFdTFillPicture = '\000d\000\006'
};
typedef enum excelMFdT excelMFdT;

enum excelMGdS {
	excelMGdSGradientUnset = '\000d\377\376',
	excelMGdSHorizontalGradient = '\000e\000\001',
	excelMGdSVerticalGradient = '\000e\000\002',
	excelMGdSDiagonalUpGradient = '\000e\000\003',
	excelMGdSDiagonalDownGradient = '\000e\000\004',
	excelMGdSFromCornerGradient = '\000e\000\005',
	excelMGdSFromTitleGradient = '\000e\000\006',
	excelMGdSFromCenterGradient = '\000e\000\007'
};
typedef enum excelMGdS excelMGdS;

enum excelMGCt {
	excelMGCtGradientTypeUnset = '\003\357\377\376',
	excelMGCtSingleShadeGradientType = '\003\360\000\001',
	excelMGCtTwoColorsGradientType = '\003\360\000\002',
	excelMGCtPresetColorsGradientType = '\003\360\000\003',
	excelMGCtMultiColorsGradientType = '\003\360\000\004'
};
typedef enum excelMGCt excelMGCt;

enum excelMxtT {
	excelMxtTTextureTypeTextureTypeUnset = '\003\360\377\376',
	excelMxtTTextureTypePresetTexture = '\003\361\000\001',
	excelMxtTTextureTypeUserDefinedTexture = '\003\361\000\002'
};
typedef enum excelMxtT excelMxtT;

enum excelMPzT {
	excelMPzTPresetTextureUnset = '\000e\377\376',
	excelMPzTTexturePapyrus = '\000f\000\001',
	excelMPzTTextureCanvas = '\000f\000\002',
	excelMPzTTextureDenim = '\000f\000\003',
	excelMPzTTextureWovenMat = '\000f\000\004',
	excelMPzTTextureWaterDroplets = '\000f\000\005',
	excelMPzTTexturePaperBag = '\000f\000\006',
	excelMPzTTextureFishFossil = '\000f\000\007',
	excelMPzTTextureSand = '\000f\000\010',
	excelMPzTTextureGreenMarble = '\000f\000\011',
	excelMPzTTextureWhiteMarble = '\000f\000\012',
	excelMPzTTextureBrownMarble = '\000f\000\013',
	excelMPzTTextureGranite = '\000f\000\014',
	excelMPzTTextureNewsprint = '\000f\000\015',
	excelMPzTTextureRecycledPaper = '\000f\000\016',
	excelMPzTTextureParchment = '\000f\000\017',
	excelMPzTTextureStationery = '\000f\000\020',
	excelMPzTTextureBlueTissuePaper = '\000f\000\021',
	excelMPzTTexturePinkTissuePaper = '\000f\000\022',
	excelMPzTTexturePurpleMesh = '\000f\000\023',
	excelMPzTTextureBouquet = '\000f\000\024',
	excelMPzTTextureCork = '\000f\000\025',
	excelMPzTTextureWalnut = '\000f\000\026',
	excelMPzTTextureOak = '\000f\000\027',
	excelMPzTTextureMediumWood = '\000f\000\030'
};
typedef enum excelMPzT excelMPzT;

enum excelPpTy {
	excelPpTyPatternUnset = '\000f\377\376',
	excelPpTyFivePercentPattern = '\000g\000\001',
	excelPpTyTenPercentPattern = '\000g\000\002',
	excelPpTyTwentyPercentPattern = '\000g\000\003',
	excelPpTyTwentyFivePercentPattern = '\000g\000\004',
	excelPpTyThirtyPercentPattern = '\000g\000\005',
	excelPpTyFortyPercentPattern = '\000g\000\006',
	excelPpTyFiftyPercentPattern = '\000g\000\007',
	excelPpTySixtyPercentPattern = '\000g\000\010',
	excelPpTySeventyPercentPattern = '\000g\000\011',
	excelPpTySeventyFivePercentPattern = '\000g\000\012',
	excelPpTyEightyPercentPattern = '\000g\000\013',
	excelPpTyNinetyPercentPattern = '\000g\000\014',
	excelPpTyDarkHorizontalPattern = '\000g\000\015',
	excelPpTyDarkVerticalPattern = '\000g\000\016',
	excelPpTyDarkDownwardDiagonalPattern = '\000g\000\017',
	excelPpTyDarkUpwardDiagonalPattern = '\000g\000\020',
	excelPpTySmallCheckerBoardPattern = '\000g\000\021',
	excelPpTyTrellisPattern = '\000g\000\022',
	excelPpTyLightHorizontalPattern = '\000g\000\023',
	excelPpTyLightVerticalPattern = '\000g\000\024',
	excelPpTyLightDownwardDiagonalPattern = '\000g\000\025',
	excelPpTyLightUpwardDiagonalPattern = '\000g\000\026',
	excelPpTySmallGridPattern = '\000g\000\027',
	excelPpTyDottedDiamondPattern = '\000g\000\030',
	excelPpTyWideDownwardDiagonal = '\000g\000\031',
	excelPpTyWideUpwardDiagonalPattern = '\000g\000\032',
	excelPpTyDashedUpwardDiagonalPattern = '\000g\000\033',
	excelPpTyDashedDownwardDiagonalPattern = '\000g\000\034',
	excelPpTyNarrowVerticalPattern = '\000g\000\035',
	excelPpTyNarrowHorizontalPattern = '\000g\000\036',
	excelPpTyDashedVerticalPattern = '\000g\000\037',
	excelPpTyDashedHorizontalPattern = '\000g\000 ',
	excelPpTyLargeConfettiPattern = '\000g\000!',
	excelPpTyLargeGridPattern = '\000g\000\"',
	excelPpTyHorizontalBrickPattern = '\000g\000#',
	excelPpTyLargeCheckerBoardPattern = '\000g\000$',
	excelPpTySmallConfettiPattern = '\000g\000%',
	excelPpTyZigZagPattern = '\000g\000&',
	excelPpTySolidDiamondPattern = '\000g\000\'',
	excelPpTyDiagonalBrickPattern = '\000g\000(',
	excelPpTyOutlinedDiamondPattern = '\000g\000)',
	excelPpTyPlaidPattern = '\000g\000*',
	excelPpTySpherePattern = '\000g\000+',
	excelPpTyWeavePattern = '\000g\000,',
	excelPpTyDottedGridPattern = '\000g\000-',
	excelPpTyDivotPattern = '\000g\000.',
	excelPpTyShinglePattern = '\000g\000/',
	excelPpTyWavePattern = '\000g\0000',
	excelPpTyHorizontalPattern = '\000g\0001',
	excelPpTyVerticalPattern = '\000g\0002',
	excelPpTyCrossPattern = '\000g\0003',
	excelPpTyDownwardDiagonalPattern = '\000g\0004',
	excelPpTyUpwardDiagonalPattern = '\000g\0005',
	excelPpTyDiagonalCrossPattern = '\000g\0005'
};
typedef enum excelPpTy excelPpTy;

enum excelMPGb {
	excelMPGbPresetGradientUnset = '\000g\377\376',
	excelMPGbGradientEarlySunset = '\000h\000\001',
	excelMPGbGradientLateSunset = '\000h\000\002',
	excelMPGbGradientNightfall = '\000h\000\003',
	excelMPGbGradientDaybreak = '\000h\000\004',
	excelMPGbGradientHorizon = '\000h\000\005',
	excelMPGbGradientDesert = '\000h\000\006',
	excelMPGbGradientOcean = '\000h\000\007',
	excelMPGbGradientCalmWater = '\000h\000\010',
	excelMPGbGradientFire = '\000h\000\011',
	excelMPGbGradientFog = '\000h\000\012',
	excelMPGbGradientMoss = '\000h\000\013',
	excelMPGbGradientPeacock = '\000h\000\014',
	excelMPGbGradientWheat = '\000h\000\015',
	excelMPGbGradientParchment = '\000h\000\016',
	excelMPGbGradientMahogany = '\000h\000\017',
	excelMPGbGradientRainbow = '\000h\000\020',
	excelMPGbGradientRainbow2 = '\000h\000\021',
	excelMPGbGradientGold = '\000h\000\022',
	excelMPGbGradientGold2 = '\000h\000\023',
	excelMPGbGradientBrass = '\000h\000\024',
	excelMPGbGradientChrome = '\000h\000\025',
	excelMPGbGradientChrome2 = '\000h\000\026',
	excelMPGbGradientSilver = '\000h\000\027',
	excelMPGbGradientSapphire = '\000h\000\030'
};
typedef enum excelMPGb excelMPGb;

enum excelMSdT {
	excelMSdTShadowUnset = '\003_\377\376',
	excelMSdTShadow1 = '\003`\000\001',
	excelMSdTShadow2 = '\003`\000\002',
	excelMSdTShadow3 = '\003`\000\003',
	excelMSdTShadow4 = '\003`\000\004',
	excelMSdTShadow5 = '\003`\000\005',
	excelMSdTShadow6 = '\003`\000\006',
	excelMSdTShadow7 = '\003`\000\007',
	excelMSdTShadow8 = '\003`\000\010',
	excelMSdTShadow9 = '\003`\000\011',
	excelMSdTShadow10 = '\003`\000\012',
	excelMSdTShadow11 = '\003`\000\013',
	excelMSdTShadow12 = '\003`\000\014',
	excelMSdTShadow13 = '\003`\000\015',
	excelMSdTShadow14 = '\003`\000\016',
	excelMSdTShadow15 = '\003`\000\017',
	excelMSdTShadow16 = '\003`\000\020',
	excelMSdTShadow17 = '\003`\000\021',
	excelMSdTShadow18 = '\003`\000\022',
	excelMSdTShadow19 = '\003`\000\023',
	excelMSdTShadow20 = '\003`\000\024',
	excelMSdTShadow21 = '\003`\000\025',
	excelMSdTShadow22 = '\003`\000\026',
	excelMSdTShadow23 = '\003`\000\027',
	excelMSdTShadow24 = '\003`\000\030',
	excelMSdTShadow25 = '\003`\000\031',
	excelMSdTShadow26 = '\003`\000\032',
	excelMSdTShadow27 = '\003`\000\033',
	excelMSdTShadow28 = '\003`\000\034',
	excelMSdTShadow29 = '\003`\000\035',
	excelMSdTShadow30 = '\003`\000\036',
	excelMSdTShadow31 = '\003`\000\037',
	excelMSdTShadow32 = '\003`\000 ',
	excelMSdTShadow33 = '\003`\000!',
	excelMSdTShadow34 = '\003`\000\"',
	excelMSdTShadow35 = '\003`\000#',
	excelMSdTShadow36 = '\003`\000$',
	excelMSdTShadow37 = '\003`\000%',
	excelMSdTShadow38 = '\003`\000&',
	excelMSdTShadow39 = '\003`\000\'',
	excelMSdTShadow40 = '\003`\000(',
	excelMSdTShadow41 = '\003`\000)',
	excelMSdTShadow42 = '\003`\000*',
	excelMSdTShadow43 = '\003`\000+'
};
typedef enum excelMSdT excelMSdT;

enum excelMPXF {
	excelMPXFWordartFormatUnset = '\003\361\377\376',
	excelMPXFWordartFormat1 = '\003\362\000\000',
	excelMPXFWordartFormat2 = '\003\362\000\001',
	excelMPXFWordartFormat3 = '\003\362\000\002',
	excelMPXFWordartFormat4 = '\003\362\000\003',
	excelMPXFWordartFormat5 = '\003\362\000\004',
	excelMPXFWordartFormat6 = '\003\362\000\005',
	excelMPXFWordartFormat7 = '\003\362\000\006',
	excelMPXFWordartFormat8 = '\003\362\000\007',
	excelMPXFWordartFormat9 = '\003\362\000\010',
	excelMPXFWordartFormat10 = '\003\362\000\011',
	excelMPXFWordartFormat11 = '\003\362\000\012',
	excelMPXFWordartFormat12 = '\003\362\000\013',
	excelMPXFWordartFormat13 = '\003\362\000\014',
	excelMPXFWordartFormat14 = '\003\362\000\015',
	excelMPXFWordartFormat15 = '\003\362\000\016',
	excelMPXFWordartFormat16 = '\003\362\000\017',
	excelMPXFWordartFormat17 = '\003\362\000\020',
	excelMPXFWordartFormat18 = '\003\362\000\021',
	excelMPXFWordartFormat19 = '\003\362\000\022',
	excelMPXFWordartFormat20 = '\003\362\000\023',
	excelMPXFWordartFormat21 = '\003\362\000\024',
	excelMPXFWordartFormat22 = '\003\362\000\025',
	excelMPXFWordartFormat23 = '\003\362\000\026',
	excelMPXFWordartFormat24 = '\003\362\000\027',
	excelMPXFWordartFormat25 = '\003\362\000\030',
	excelMPXFWordartFormat26 = '\003\362\000\031',
	excelMPXFWordartFormat27 = '\003\362\000\032',
	excelMPXFWordartFormat28 = '\003\362\000\033',
	excelMPXFWordartFormat29 = '\003\362\000\034',
	excelMPXFWordartFormat30 = '\003\362\000\035'
};
typedef enum excelMPXF excelMPXF;

enum excelMPTs {
	excelMPTsTextEffectShapeUnset = '\000\227\377\376',
	excelMPTsPlainText = '\000\230\000\001',
	excelMPTsStop = '\000\230\000\002',
	excelMPTsTriangleUp = '\000\230\000\003',
	excelMPTsTriangleDown = '\000\230\000\004',
	excelMPTsChevronUp = '\000\230\000\005',
	excelMPTsChevronDown = '\000\230\000\006',
	excelMPTsRingInside = '\000\230\000\007',
	excelMPTsRingOutside = '\000\230\000\010',
	excelMPTsArchUpCurve = '\000\230\000\011',
	excelMPTsArchDownCurve = '\000\230\000\012',
	excelMPTsCircleCurve = '\000\230\000\013',
	excelMPTsButtonCurve = '\000\230\000\014',
	excelMPTsArchUpPour = '\000\230\000\015',
	excelMPTsArchDownPour = '\000\230\000\016',
	excelMPTsCirclePour = '\000\230\000\017',
	excelMPTsButtonPour = '\000\230\000\020',
	excelMPTsCurveUp = '\000\230\000\021',
	excelMPTsCurveDown = '\000\230\000\022',
	excelMPTsCanUp = '\000\230\000\023',
	excelMPTsCanDown = '\000\230\000\024',
	excelMPTsWave1 = '\000\230\000\025',
	excelMPTsWave2 = '\000\230\000\026',
	excelMPTsDoubleWave1 = '\000\230\000\027',
	excelMPTsDoubleWave2 = '\000\230\000\030',
	excelMPTsInflate = '\000\230\000\031',
	excelMPTsDeflate = '\000\230\000\032',
	excelMPTsInflateBottom = '\000\230\000\033',
	excelMPTsDeflateBottom = '\000\230\000\034',
	excelMPTsInflateTop = '\000\230\000\035',
	excelMPTsDeflateTop = '\000\230\000\036',
	excelMPTsDeflateInflate = '\000\230\000\037',
	excelMPTsDeflateInflateDeflate = '\000\230\000 ',
	excelMPTsFadeRight = '\000\230\000!',
	excelMPTsFadeLeft = '\000\230\000\"',
	excelMPTsFadeUp = '\000\230\000#',
	excelMPTsFadeDown = '\000\230\000$',
	excelMPTsSlantUp = '\000\230\000%',
	excelMPTsSlantDown = '\000\230\000&',
	excelMPTsCascadeUp = '\000\230\000\'',
	excelMPTsCascadeDown = '\000\230\000('
};
typedef enum excelMPTs excelMPTs;

enum excelMTxA {
	excelMTxATextEffectAlignmentUnset = '\000\226\377\376',
	excelMTxALeftTextEffectAlignment = '\000\227\000\001',
	excelMTxACenteredTextEffectAlignment = '\000\227\000\002',
	excelMTxARightTextEffectAlignment = '\000\227\000\003',
	excelMTxAJustifyTextEffectAlignment = '\000\227\000\004',
	excelMTxAWordJustifyTextEffectAlignment = '\000\227\000\005',
	excelMTxAStretchJustifyTextEffectAlignment = '\000\227\000\006'
};
typedef enum excelMTxA excelMTxA;

enum excelMPLd {
	excelMPLdPresetLightingDirectionUnset = '\000\233\377\376',
	excelMPLdLightFromTopLeft = '\000\234\000\001',
	excelMPLdLightFromTop = '\000\234\000\002',
	excelMPLdLightFromTopRight = '\000\234\000\003',
	excelMPLdLightFromLeft = '\000\234\000\004',
	excelMPLdLightFromNone = '\000\234\000\005',
	excelMPLdLightFromRight = '\000\234\000\006',
	excelMPLdLightFromBottomLeft = '\000\234\000\007',
	excelMPLdLightFromBottom = '\000\234\000\010',
	excelMPLdLightFromBottomRight = '\000\234\000\011'
};
typedef enum excelMPLd excelMPLd;

enum excelMlSf {
	excelMlSfLightingSoftnessUnset = '\000\234\377\376',
	excelMlSfLightingDim = '\000\235\000\001',
	excelMlSfLightingNormal = '\000\235\000\002',
	excelMlSfLightingBright = '\000\235\000\003'
};
typedef enum excelMlSf excelMlSf;

enum excelMPMt {
	excelMPMtPresetMaterialUnset = '\000\235\377\376',
	excelMPMtMatte = '\000\236\000\001',
	excelMPMtPlastic = '\000\236\000\002',
	excelMPMtMetal = '\000\236\000\003',
	excelMPMtWireframe = '\000\236\000\004',
	excelMPMtMatte2 = '\000\236\000\005',
	excelMPMtPlastic2 = '\000\236\000\006',
	excelMPMtMetal2 = '\000\236\000\007',
	excelMPMtWarmMatte = '\000\236\000\010',
	excelMPMtTranslucentPowder = '\000\236\000\011',
	excelMPMtPowder = '\000\236\000\012',
	excelMPMtDarkEdge = '\000\236\000\013',
	excelMPMtSoftEdge = '\000\236\000\014',
	excelMPMtMaterialClear = '\000\236\000\015',
	excelMPMtFlat = '\000\236\000\016',
	excelMPMtSoftMetal = '\000\236\000\017'
};
typedef enum excelMPMt excelMPMt;

enum excelMExD {
	excelMExDPresetExtrusionDirectionUnset = '\000\231\377\376',
	excelMExDExtrudeBottomRight = '\000\232\000\001',
	excelMExDExtrudeBottom = '\000\232\000\002',
	excelMExDExtrudeBottomLeft = '\000\232\000\003',
	excelMExDExtrudeRight = '\000\232\000\004',
	excelMExDExtrudeNone = '\000\232\000\005',
	excelMExDExtrudeLeft = '\000\232\000\006',
	excelMExDExtrudeTopRight = '\000\232\000\007',
	excelMExDExtrudeTop = '\000\232\000\010',
	excelMExDExtrudeTopLeft = '\000\232\000\011'
};
typedef enum excelMExD excelMExD;

enum excelM3DF {
	excelM3DFPresetThreeDFormatUnset = '\000\230\377\376',
	excelM3DFFormat1 = '\000\231\000\001',
	excelM3DFFormat2 = '\000\231\000\002',
	excelM3DFFormat3 = '\000\231\000\003',
	excelM3DFFormat4 = '\000\231\000\004',
	excelM3DFFormat5 = '\000\231\000\005',
	excelM3DFFormat6 = '\000\231\000\006',
	excelM3DFFormat7 = '\000\231\000\007',
	excelM3DFFormat8 = '\000\231\000\010',
	excelM3DFFormat9 = '\000\231\000\011',
	excelM3DFFormat10 = '\000\231\000\012',
	excelM3DFFormat11 = '\000\231\000\013',
	excelM3DFFormat12 = '\000\231\000\014',
	excelM3DFFormat13 = '\000\231\000\015',
	excelM3DFFormat14 = '\000\231\000\016',
	excelM3DFFormat15 = '\000\231\000\017',
	excelM3DFFormat16 = '\000\231\000\020',
	excelM3DFFormat17 = '\000\231\000\021',
	excelM3DFFormat18 = '\000\231\000\022',
	excelM3DFFormat19 = '\000\231\000\023',
	excelM3DFFormat20 = '\000\231\000\024'
};
typedef enum excelM3DF excelM3DF;

enum excelMExC {
	excelMExCExtrusionColorUnset = '\000\232\377\376',
	excelMExCExtrusionColorAutomatic = '\000\233\000\001',
	excelMExCExtrusionColorCustom = '\000\233\000\002'
};
typedef enum excelMExC excelMExC;

enum excelMCtT {
	excelMCtTConnectorTypeUnset = '\000h\377\376',
	excelMCtTStraight = '\000i\000\001',
	excelMCtTElbow = '\000i\000\002',
	excelMCtTCurve = '\000i\000\003'
};
typedef enum excelMCtT excelMCtT;

enum excelMHzA {
	excelMHzAHorizontalAnchorUnset = '\000\236\377\376',
	excelMHzAHorizontalAnchorNone = '\000\237\000\001',
	excelMHzAHorizontalAnchorCenter = '\000\237\000\002'
};
typedef enum excelMHzA excelMHzA;

enum excelMVtA {
	excelMVtAVerticalAnchorUnset = '\000\237\377\376',
	excelMVtAAnchorTop = '\000\240\000\001',
	excelMVtAAnchorTopBaseline = '\000\240\000\002',
	excelMVtAAnchorMiddle = '\000\240\000\003',
	excelMVtAAnchorBottom = '\000\240\000\004',
	excelMVtAAnchorBottomBaseline = '\000\240\000\005'
};
typedef enum excelMVtA excelMVtA;

enum excelMAsT {
	excelMAsTAutoshapeShapeTypeUnset = '\000i\377\376',
	excelMAsTAutoshapeRectangle = '\000j\000\001',
	excelMAsTAutoshapeParallelogram = '\000j\000\002',
	excelMAsTAutoshapeTrapezoid = '\000j\000\003',
	excelMAsTAutoshapeDiamond = '\000j\000\004',
	excelMAsTAutoshapeRoundedRectangle = '\000j\000\005',
	excelMAsTAutoshapeOctagon = '\000j\000\006',
	excelMAsTAutoshapeIsoscelesTriangle = '\000j\000\007',
	excelMAsTAutoshapeRightTriangle = '\000j\000\010',
	excelMAsTAutoshapeOval = '\000j\000\011',
	excelMAsTAutoshapeHexagon = '\000j\000\012',
	excelMAsTAutoshapeCross = '\000j\000\013',
	excelMAsTAutoshapeRegularPentagon = '\000j\000\014',
	excelMAsTAutoshapeCan = '\000j\000\015',
	excelMAsTAutoshapeCube = '\000j\000\016',
	excelMAsTAutoshapeBevel = '\000j\000\017',
	excelMAsTAutoshapeFoldedCorner = '\000j\000\020',
	excelMAsTAutoshapeSmileyFace = '\000j\000\021',
	excelMAsTAutoshapeDonut = '\000j\000\022',
	excelMAsTAutoshapeNoSymbol = '\000j\000\023',
	excelMAsTAutoshapeBlockArc = '\000j\000\024',
	excelMAsTAutoshapeHeart = '\000j\000\025',
	excelMAsTAutoshapeLightningBolt = '\000j\000\026',
	excelMAsTAutoshapeSun = '\000j\000\027',
	excelMAsTAutoshapeMoon = '\000j\000\030',
	excelMAsTAutoshapeArc = '\000j\000\031',
	excelMAsTAutoshapeDoubleBracket = '\000j\000\032',
	excelMAsTAutoshapeDoubleBrace = '\000j\000\033',
	excelMAsTAutoshapePlaque = '\000j\000\034',
	excelMAsTAutoshapeLeftBracket = '\000j\000\035',
	excelMAsTAutoshapeRightBracket = '\000j\000\036',
	excelMAsTAutoshapeLeftBrace = '\000j\000\037',
	excelMAsTAutoshapeRightBrace = '\000j\000 ',
	excelMAsTAutoshapeRightArrow = '\000j\000!',
	excelMAsTAutoshapeLeftArrow = '\000j\000\"',
	excelMAsTAutoshapeUpArrow = '\000j\000#',
	excelMAsTAutoshapeDownArrow = '\000j\000$',
	excelMAsTAutoshapeLeftRightArrow = '\000j\000%',
	excelMAsTAutoshapeUpDownArrow = '\000j\000&',
	excelMAsTAutoshapeQuadArrow = '\000j\000\'',
	excelMAsTAutoshapeLeftRightUpArrow = '\000j\000(',
	excelMAsTAutoshapeBentArrow = '\000j\000)',
	excelMAsTAutoshapeUTurnArrow = '\000j\000*',
	excelMAsTAutoshapeLeftUpArrow = '\000j\000+',
	excelMAsTAutoshapeBentUpArrow = '\000j\000,',
	excelMAsTAutoshapeCurvedRightArrow = '\000j\000-',
	excelMAsTAutoshapeCurvedLeftArrow = '\000j\000.',
	excelMAsTAutoshapeCurvedUpArrow = '\000j\000/',
	excelMAsTAutoshapeCurvedDownArrow = '\000j\0000',
	excelMAsTAutoshapeStripedRightArrow = '\000j\0001',
	excelMAsTAutoshapeNotchedRightArrow = '\000j\0002',
	excelMAsTAutoshapePentagon = '\000j\0003',
	excelMAsTAutoshapeChevron = '\000j\0004',
	excelMAsTAutoshapeRightArrowCallout = '\000j\0005',
	excelMAsTAutoshapeLeftArrowCallout = '\000j\0006',
	excelMAsTAutoshapeUpArrowCallout = '\000j\0007',
	excelMAsTAutoshapeDownArrowCallout = '\000j\0008',
	excelMAsTAutoshapeLeftRightArrowCallout = '\000j\0009',
	excelMAsTAutoshapeUpDownArrowCallout = '\000j\000:',
	excelMAsTAutoshapeQuadArrowCallout = '\000j\000;',
	excelMAsTAutoshapeCircularArrow = '\000j\000<',
	excelMAsTAutoshapeFlowchartProcess = '\000j\000=',
	excelMAsTAutoshapeFlowchartAlternateProcess = '\000j\000>',
	excelMAsTAutoshapeFlowchartDecision = '\000j\000\?',
	excelMAsTAutoshapeFlowchartData = '\000j\000@',
	excelMAsTAutoshapeFlowchartPredefinedProcess = '\000j\000A',
	excelMAsTAutoshapeFlowchartInternalStorage = '\000j\000B',
	excelMAsTAutoshapeFlowchartDocument = '\000j\000C',
	excelMAsTAutoshapeFlowchartMultiDocument = '\000j\000D',
	excelMAsTAutoshapeFlowchartTerminator = '\000j\000E',
	excelMAsTAutoshapeFlowchartPreparation = '\000j\000F',
	excelMAsTAutoshapeFlowchartManualInput = '\000j\000G',
	excelMAsTAutoshapeFlowchartManualOperation = '\000j\000H',
	excelMAsTAutoshapeFlowchartConnector = '\000j\000I',
	excelMAsTAutoshapeFlowchartOffpageConnector = '\000j\000J',
	excelMAsTAutoshapeFlowchartCard = '\000j\000K',
	excelMAsTAutoshapeFlowchartPunchedTape = '\000j\000L',
	excelMAsTAutoshapeFlowchartSummingJunction = '\000j\000M',
	excelMAsTAutoshapeFlowchartOr = '\000j\000N',
	excelMAsTAutoshapeFlowchartCollate = '\000j\000O',
	excelMAsTAutoshapeFlowchartSort = '\000j\000P',
	excelMAsTAutoshapeFlowchartExtract = '\000j\000Q',
	excelMAsTAutoshapeFlowchartMerge = '\000j\000R',
	excelMAsTAutoshapeFlowchartStoredData = '\000j\000S',
	excelMAsTAutoshapeFlowchartDelay = '\000j\000T',
	excelMAsTAutoshapeFlowchartSequentialAccessStorage = '\000j\000U',
	excelMAsTAutoshapeFlowchartMagneticDisk = '\000j\000V',
	excelMAsTAutoshapeFlowchartDirectAccessStorage = '\000j\000W',
	excelMAsTAutoshapeFlowchartDisplay = '\000j\000X',
	excelMAsTAutoshapeExplosionOne = '\000j\000Y',
	excelMAsTAutoshapeExplosionTwo = '\000j\000Z',
	excelMAsTAutoshapeFourPointStar = '\000j\000[',
	excelMAsTAutoshapeFivePointStar = '\000j\000\\',
	excelMAsTAutoshapeEightPointStar = '\000j\000]',
	excelMAsTAutoshapeSixteenPointStar = '\000j\000^',
	excelMAsTAutoshapeTwentyFourPointStar = '\000j\000_',
	excelMAsTAutoshapeThirtyTwoPointStar = '\000j\000`',
	excelMAsTAutoshapeUpRibbon = '\000j\000a',
	excelMAsTAutoshapeDownRibbon = '\000j\000b',
	excelMAsTAutoshapeCurvedUpRibbon = '\000j\000c',
	excelMAsTAutoshapeCurvedDownRibbon = '\000j\000d',
	excelMAsTAutoshapeVerticalScroll = '\000j\000e',
	excelMAsTAutoshapeHorizontalScroll = '\000j\000f',
	excelMAsTAutoshapeWave = '\000j\000g',
	excelMAsTAutoshapeDoubleWave = '\000j\000h',
	excelMAsTAutoshapeRectangularCallout = '\000j\000i',
	excelMAsTAutoshapeRoundedRectangularCallout = '\000j\000j',
	excelMAsTAutoshapeOvalCallout = '\000j\000k',
	excelMAsTAutoshapeCloudCallout = '\000j\000l',
	excelMAsTAutoshapeLineCalloutOne = '\000j\000m',
	excelMAsTAutoshapeLineCalloutTwo = '\000j\000n',
	excelMAsTAutoshapeLineCalloutThree = '\000j\000o',
	excelMAsTAutoshapeLineCalloutFour = '\000j\000p',
	excelMAsTAutoshapeLineCalloutOneAccentBar = '\000j\000q',
	excelMAsTAutoshapeLineCalloutTwoAccentBar = '\000j\000r',
	excelMAsTAutoshapeLineCalloutThreeAccentBar = '\000j\000s',
	excelMAsTAutoshapeLineCalloutFourAccentBar = '\000j\000t',
	excelMAsTAutoshapeLineCalloutOneNoBorder = '\000j\000u',
	excelMAsTAutoshapeLineCalloutTwoNoBorder = '\000j\000v',
	excelMAsTAutoshapeLineCalloutThreeNoBorder = '\000j\000w',
	excelMAsTAutoshapeLineCalloutFourNoBorder = '\000j\000x',
	excelMAsTAutoshapeCalloutOneBorderAndAccentBar = '\000j\000y',
	excelMAsTAutoshapeCalloutTwoBorderAndAccentBar = '\000j\000z',
	excelMAsTAutoshapeCalloutThreeBorderAndAccentBar = '\000j\000{',
	excelMAsTAutoshapeCalloutFourBorderAndAccentBar = '\000j\000|',
	excelMAsTAutoshapeActionButtonCustom = '\000j\000}',
	excelMAsTAutoshapeActionButtonHome = '\000j\000~',
	excelMAsTAutoshapeActionButtonHelp = '\000j\000\177',
	excelMAsTAutoshapeActionButtonInformation = '\000j\000\200',
	excelMAsTAutoshapeActionButtonBackOrPrevious = '\000j\000\201',
	excelMAsTAutoshapeActionButtonForwardOrNext = '\000j\000\202',
	excelMAsTAutoshapeActionButtonBeginning = '\000j\000\203',
	excelMAsTAutoshapeActionButtonEnd = '\000j\000\204',
	excelMAsTAutoshapeActionButtonReturn = '\000j\000\205',
	excelMAsTAutoshapeActionButtonDocument = '\000j\000\206',
	excelMAsTAutoshapeActionButtonSound = '\000j\000\207',
	excelMAsTAutoshapeActionButtonMovie = '\000j\000\210',
	excelMAsTAutoshapeBalloon = '\000j\000\211',
	excelMAsTAutoshapeNotPrimitive = '\000j\000\212',
	excelMAsTAutoshapeFlowchartOfflineStorage = '\000j\000\213',
	excelMAsTAutoshapeLeftRightRibbon = '\000j\000\214',
	excelMAsTAutoshapeDiagonalStripe = '\000j\000\215',
	excelMAsTAutoshapePie = '\000j\000\216',
	excelMAsTAutoshapeNonIsoscelesTrapezoid = '\000j\000\217',
	excelMAsTAutoshapeDecagon = '\000j\000\220',
	excelMAsTAutoshapeHeptagon = '\000j\000\221',
	excelMAsTAutoshapeDodecagon = '\000j\000\222',
	excelMAsTAutoshapeSixPointsStar = '\000j\000\223',
	excelMAsTAutoshapeSevenPointsStar = '\000j\000\224',
	excelMAsTAutoshapeTenPointsStar = '\000j\000\225',
	excelMAsTAutoshapeTwelvePointsStar = '\000j\000\226',
	excelMAsTAutoshapeRoundOneRectangle = '\000j\000\227',
	excelMAsTAutoshapeRoundTwoSameRectangle = '\000j\000\230',
	excelMAsTAutoshapeRoundTwoDiagonalRectangle = '\000j\000\231',
	excelMAsTAutoshapeSnipRoundRectangle = '\000j\000\232',
	excelMAsTAutoshapeSnipOneRectangle = '\000j\000\233',
	excelMAsTAutoshapeSnipTwoSameRectangle = '\000j\000\234',
	excelMAsTAutoshapeSnipTwoDiagonalRectangle = '\000j\000\235',
	excelMAsTAutoshapeFrame = '\000j\000\236',
	excelMAsTAutoshapeHalfFrame = '\000j\000\237',
	excelMAsTAutoshapeTear = '\000j\000\240',
	excelMAsTAutoshapeChord = '\000j\000\241',
	excelMAsTAutoshapeCorner = '\000j\000\242',
	excelMAsTAutoshapeMathPlus = '\000j\000\243',
	excelMAsTAutoshapeMathMinus = '\000j\000\244',
	excelMAsTAutoshapeMathMultiply = '\000j\000\245',
	excelMAsTAutoshapeMathDivide = '\000j\000\246',
	excelMAsTAutoshapeMathEqual = '\000j\000\247',
	excelMAsTAutoshapeMathNotEqual = '\000j\000\250',
	excelMAsTAutoshapeCornerTabs = '\000j\000\251',
	excelMAsTAutoshapeSquareTabs = '\000j\000\252',
	excelMAsTAutoshapePlaqueTabs = '\000j\000\253',
	excelMAsTAutoshapeGearSix = '\000j\000\254',
	excelMAsTAutoshapeGearNine = '\000j\000\255',
	excelMAsTAutoshapeFunnel = '\000j\000\256',
	excelMAsTAutoshapePieWedge = '\000j\000\257',
	excelMAsTAutoshapeLeftCircularArrow = '\000j\000\260',
	excelMAsTAutoshapeLeftRightCircularArrow = '\000j\000\261',
	excelMAsTAutoshapeSwooshArrow = '\000j\000\262',
	excelMAsTAutoshapeCloud = '\000j\000\263',
	excelMAsTAutoshapeChartX = '\000j\000\264',
	excelMAsTAutoshapeChartStar = '\000j\000\265',
	excelMAsTAutoshapeChartPlus = '\000j\000\266',
	excelMAsTAutoshapeLineInverse = '\000j\000\267'
};
typedef enum excelMAsT excelMAsT;

enum excelMShp {
	excelMShpShapeTypeUnset = '\000\213\377\376',
	excelMShpShapeTypeAuto = '\000\214\000\001',
	excelMShpShapeTypeCallout = '\000\214\000\002',
	excelMShpShapeTypeChart = '\000\214\000\003',
	excelMShpShapeTypeComment = '\000\214\000\004',
	excelMShpShapeTypeFreeForm = '\000\214\000\005',
	excelMShpShapeTypeGroup = '\000\214\000\006',
	excelMShpShapeTypeEmbeddedOLEControl = '\000\214\000\007',
	excelMShpShapeTypeFormControl = '\000\214\000\010',
	excelMShpShapeTypeLine = '\000\214\000\011',
	excelMShpShapeTypeLinkedOLEObject = '\000\214\000\012',
	excelMShpShapeTypeLinkedPicture = '\000\214\000\013',
	excelMShpShapeTypeOLEControl = '\000\214\000\014',
	excelMShpShapeTypePicture = '\000\214\000\015',
	excelMShpShapeTypePlaceHolder = '\000\214\000\016',
	excelMShpShapeTypeWordArt = '\000\214\000\017',
	excelMShpShapeTypeMedia = '\000\214\000\020',
	excelMShpShapeTypeTextBox = '\000\214\000\021',
	excelMShpShapeTypeTable = '\000\214\000\022',
	excelMShpShapeTypeCanvas = '\000\214\000\023',
	excelMShpShapeTypeDiagram = '\000\214\000\024',
	excelMShpShapeTypeInk = '\000\214\000\025',
	excelMShpShapeTypeInkComment = '\000\214\000\026',
	excelMShpShapeTypeSmartartGraphic = '\000\214\000\027',
	excelMShpShapeTypeSlicer = '\000\214\000\030'
};
typedef enum excelMShp excelMShp;

enum excelMCrT {
	excelMCrTColorTypeUnset = '\000j\377\376',
	excelMCrTRGB = '\000k\000\001',
	excelMCrTScheme = '\000k\000\002'
};
typedef enum excelMCrT excelMCrT;

enum excelMPc {
	excelMPcPictureColorTypeUnset = '\000\265\377\376',
	excelMPcPictureColorAutomatic = '\000\266\000\001',
	excelMPcPictureColorGrayScale = '\000\266\000\002',
	excelMPcPictureColorBlackAndWhite = '\000\266\000\003',
	excelMPcPictureColorWatermark = '\000\266\000\004'
};
typedef enum excelMPc excelMPc;

enum excelMCAt {
	excelMCAtAngleUnset = '\000k\377\376',
	excelMCAtAngleAutomatic = '\000l\000\001',
	excelMCAtAngle30 = '\000l\000\002',
	excelMCAtAngle45 = '\000l\000\003',
	excelMCAtAngle60 = '\000l\000\004',
	excelMCAtAngle90 = '\000l\000\005'
};
typedef enum excelMCAt excelMCAt;

enum excelMCDt {
	excelMCDtDropUnset = '\000l\377\376',
	excelMCDtDropCustom = '\000m\000\001',
	excelMCDtDropTop = '\000m\000\002',
	excelMCDtDropCenter = '\000m\000\003',
	excelMCDtDropBottom = '\000m\000\004'
};
typedef enum excelMCDt excelMCDt;

enum excelMCot {
	excelMCotCalloutUnset = '\000m\377\376',
	excelMCotCalloutOne = '\000n\000\001',
	excelMCotCalloutTwo = '\000n\000\002',
	excelMCotCalloutThree = '\000n\000\003',
	excelMCotCalloutFour = '\000n\000\004'
};
typedef enum excelMCot excelMCot;

enum excelTxOr {
	excelTxOrTextOrientationUnset = '\000\215\377\376',
	excelTxOrHorizontal = '\000\216\000\001',
	excelTxOrUpward = '\000\216\000\002',
	excelTxOrDownward = '\000\216\000\003',
	excelTxOrVerticalEastAsian = '\000\216\000\004',
	excelTxOrVertical = '\000\216\000\005',
	excelTxOrHorizontalRotatedEastAsian = '\000\216\000\006'
};
typedef enum excelTxOr excelTxOr;

enum excelMSFr {
	excelMSFrScaleFromTopLeft = '\000o\000\000',
	excelMSFrScaleFromMiddle = '\000o\000\001',
	excelMSFrScaleFromBottomRight = '\000o\000\002'
};
typedef enum excelMSFr excelMSFr;

enum excelMPzC {
	excelMPzCPresetCameraUnset = '\000\256\377\376',
	excelMPzCCameraLegacyObliqueFromTopLeft = '\000\257\000\001',
	excelMPzCCameraLegacyObliqueFromTop = '\000\257\000\002',
	excelMPzCCameraLegacyObliqueFromTopright = '\000\257\000\003',
	excelMPzCCameraLegacyObliqueFromLeft = '\000\257\000\004',
	excelMPzCCameraLegacyObliqueFromFront = '\000\257\000\005',
	excelMPzCCameraLegacyObliqueFromRight = '\000\257\000\006',
	excelMPzCCameraLegacyObliqueFromBottomLeft = '\000\257\000\007',
	excelMPzCCameraLegacyObliqueFromBottom = '\000\257\000\010',
	excelMPzCCameraLegacyObliqueFromBottomRight = '\000\257\000\011',
	excelMPzCCameraLegacyPerspectiveFromTopLeft = '\000\257\000\012',
	excelMPzCCameraLegacyPerspectiveFromTop = '\000\257\000\013',
	excelMPzCCameraLegacyPerspectiveFromTopRight = '\000\257\000\014',
	excelMPzCCameraLegacyPerspectiveFromLeft = '\000\257\000\015',
	excelMPzCCameraLegacyPerspectiveFromFront = '\000\257\000\016',
	excelMPzCCameraLegacyPerspectiveFromRight = '\000\257\000\017',
	excelMPzCCameraLegacyPerspectiveFromBottomLeft = '\000\257\000\020',
	excelMPzCCameraLegacyPerspectiveFromBottom = '\000\257\000\021',
	excelMPzCCameraLegacyPerspectiveFromBottomRight = '\000\257\000\022',
	excelMPzCCameraOrthographic = '\000\257\000\023',
	excelMPzCCameraIsometricFromTopUp = '\000\257\000\024',
	excelMPzCCameraIsometricFromTopDown = '\000\257\000\025',
	excelMPzCCameraIsometricFromBottomUp = '\000\257\000\026',
	excelMPzCCameraIsometricFromBottomDown = '\000\257\000\027',
	excelMPzCCameraIsometricFromLeftUp = '\000\257\000\030',
	excelMPzCCameraIsometricFromLeftDown = '\000\257\000\031',
	excelMPzCCameraIsometricFromRightUp = '\000\257\000\032',
	excelMPzCCameraIsometricFromRightDown = '\000\257\000\033',
	excelMPzCCameraIsometricOffAxis1FromLeft = '\000\257\000\034',
	excelMPzCCameraIsometricOffAxis1FromRight = '\000\257\000\035',
	excelMPzCCameraIsometricOffAxis1FromTop = '\000\257\000\036',
	excelMPzCCameraIsometricOffAxis2FromLeft = '\000\257\000\037',
	excelMPzCCameraIsometricOffAxis2FromRight = '\000\257\000 ',
	excelMPzCCameraIsometricOffAxis2FromTop = '\000\257\000!',
	excelMPzCCameraIsometricOffAxis3FromLeft = '\000\257\000\"',
	excelMPzCCameraIsometricOffAxis3FromRight = '\000\257\000#',
	excelMPzCCameraIsometricOffAxis3FromBottom = '\000\257\000$',
	excelMPzCCameraIsometricOffAxis4FromLeft = '\000\257\000%',
	excelMPzCCameraIsometricOffAxis4FromRight = '\000\257\000&',
	excelMPzCCameraIsometricOffAxis4FromBottom = '\000\257\000\'',
	excelMPzCCameraObliqueFromTopLeft = '\000\257\000(',
	excelMPzCCameraObliqueFromTop = '\000\257\000)',
	excelMPzCCameraObliqueFromTopRight = '\000\257\000*',
	excelMPzCCameraObliqueFromLeft = '\000\257\000+',
	excelMPzCCameraObliqueFromRight = '\000\257\000,',
	excelMPzCCameraObliqueFromBottomLeft = '\000\257\000-',
	excelMPzCCameraObliqueFromBottom = '\000\257\000.',
	excelMPzCCameraObliqueFromBottomRight = '\000\257\000/',
	excelMPzCCameraPerspectiveFromFront = '\000\257\0000',
	excelMPzCCameraPerspectiveFromLeft = '\000\257\0001',
	excelMPzCCameraPerspectiveFromRight = '\000\257\0002',
	excelMPzCCameraPerspectiveFromAbove = '\000\257\0003',
	excelMPzCCameraPerspectiveFromBelow = '\000\257\0004',
	excelMPzCCameraPerspectiveFromAboveFacingLeft = '\000\257\0005',
	excelMPzCCameraPerspectiveFromAboveFacingRight = '\000\257\0006',
	excelMPzCCameraPerspectiveContrastingFacingLeft = '\000\257\0007',
	excelMPzCCameraPerspectiveContrastingFacingRight = '\000\257\0008',
	excelMPzCCameraPerspectiveHeroicFacingLeft = '\000\257\0009',
	excelMPzCCameraPerspectiveHeroicFacingRight = '\000\257\000:',
	excelMPzCCameraPerspectiveHeroicExtremeFacingLeft = '\000\257\000;',
	excelMPzCCameraPerspectiveHeroicExtremeFacingRight = '\000\257\000<',
	excelMPzCCameraPerspectiveRelaxed = '\000\257\000=',
	excelMPzCCameraPerspectiveRelaxedModerately = '\000\257\000>'
};
typedef enum excelMPzC excelMPzC;

enum excelMLtT {
	excelMLtTLightRigUnset = '\000\257\377\376',
	excelMLtTLightRigFlat1 = '\000\260\000\001',
	excelMLtTLightRigFlat2 = '\000\260\000\002',
	excelMLtTLightRigFlat3 = '\000\260\000\003',
	excelMLtTLightRigFlat4 = '\000\260\000\004',
	excelMLtTLightRigNormal1 = '\000\260\000\005',
	excelMLtTLightRigNormal2 = '\000\260\000\006',
	excelMLtTLightRigNormal3 = '\000\260\000\007',
	excelMLtTLightRigNormal4 = '\000\260\000\010',
	excelMLtTLightRigHarsh1 = '\000\260\000\011',
	excelMLtTLightRigHarsh2 = '\000\260\000\012',
	excelMLtTLightRigHarsh3 = '\000\260\000\013',
	excelMLtTLightRigHarsh4 = '\000\260\000\014',
	excelMLtTLightRigThreePoint = '\000\260\000\015',
	excelMLtTLightRigBalanced = '\000\260\000\016',
	excelMLtTLightRigSoft = '\000\260\000\017',
	excelMLtTLightRigHarsh = '\000\260\000\020',
	excelMLtTLightRigFlood = '\000\260\000\021',
	excelMLtTLightRigContrasting = '\000\260\000\022',
	excelMLtTLightRigMorning = '\000\260\000\023',
	excelMLtTLightRigSunrise = '\000\260\000\024',
	excelMLtTLightRigSunset = '\000\260\000\025',
	excelMLtTLightRigChilly = '\000\260\000\026',
	excelMLtTLightRigFreezing = '\000\260\000\027',
	excelMLtTLightRigFlat = '\000\260\000\030',
	excelMLtTLightRigTwoPoint = '\000\260\000\031',
	excelMLtTLightRigGlow = '\000\260\000\032',
	excelMLtTLightRigBrightRoom = '\000\260\000\033'
};
typedef enum excelMLtT excelMLtT;

enum excelMBlT {
	excelMBlTBevelTypeUnset = '\000\260\377\376',
	excelMBlTBevelNone = '\000\261\000\001',
	excelMBlTBevelRelaxedInset = '\000\261\000\002',
	excelMBlTBevelCircle = '\000\261\000\003',
	excelMBlTBevelSlope = '\000\261\000\004',
	excelMBlTBevelCross = '\000\261\000\005',
	excelMBlTBevelAngle = '\000\261\000\006',
	excelMBlTBevelSoftRound = '\000\261\000\007',
	excelMBlTBevelConvex = '\000\261\000\010',
	excelMBlTBevelCoolSlant = '\000\261\000\011',
	excelMBlTBevelDivot = '\000\261\000\012',
	excelMBlTBevelRiblet = '\000\261\000\013',
	excelMBlTBevelHardEdge = '\000\261\000\014',
	excelMBlTBevelArtDeco = '\000\261\000\015'
};
typedef enum excelMBlT excelMBlT;

enum excelMSSt {
	excelMSStShadowStyleUnset = '\000\261\377\376',
	excelMSStShadowStyleInner = '\000\262\000\001',
	excelMSStShadowStyleOuter = '\000\262\000\002'
};
typedef enum excelMSSt excelMSSt;

enum excelPpgA {
	excelPpgAParagraphAlignmentUnset = '\000\346\377\376',
	excelPpgAParagraphAlignLeft = '\000\347\000\000',
	excelPpgAParagraphAlignCenter = '\000\347\000\001',
	excelPpgAParagraphAlignRight = '\000\347\000\002',
	excelPpgAParagraphAlignJustify = '\000\347\000\003',
	excelPpgAParagraphAlignDistribute = '\000\347\000\004',
	excelPpgAParagraphAlignThai = '\000\347\000\005',
	excelPpgAParagraphAlignJustifyLow = '\000\347\000\006'
};
typedef enum excelPpgA excelPpgA;

enum excelMTSt {
	excelMTStStrikeUnset = '\000\263\377\376',
	excelMTStNoStrike = '\000\264\000\000',
	excelMTStSingleStrike = '\000\264\000\001',
	excelMTStDoubleStrike = '\000\264\000\002'
};
typedef enum excelMTSt excelMTSt;

enum excelMTCa {
	excelMTCaCapsUnset = '\000\264\377\376',
	excelMTCaNoCaps = '\000\265\000\000',
	excelMTCaSmallCaps = '\000\265\000\001',
	excelMTCaAllCaps = '\000\265\000\002'
};
typedef enum excelMTCa excelMTCa;

enum excelMTUl {
	excelMTUlUnderlineUnset = '\003\356\377\376',
	excelMTUlNoUnderline = '\003\357\000\000',
	excelMTUlUnderlineWordsOnly = '\003\357\000\001',
	excelMTUlUnderlineSingleLine = '\003\357\000\002',
	excelMTUlUnderlineDoubleLine = '\003\357\000\003',
	excelMTUlUnderlineHeavyLine = '\003\357\000\004',
	excelMTUlUnderlineDottedLine = '\003\357\000\005',
	excelMTUlUnderlineHeavyDottedLine = '\003\357\000\006',
	excelMTUlUnderlineDashLine = '\003\357\000\007',
	excelMTUlUnderlineHeavyDashLine = '\003\357\000\010',
	excelMTUlUnderlineLongDashLine = '\003\357\000\011',
	excelMTUlUnderlineHeavyLongDashLine = '\003\357\000\012',
	excelMTUlUnderlineDotDashLine = '\003\357\000\013',
	excelMTUlUnderlineHeavyDotDashLine = '\003\357\000\014',
	excelMTUlUnderlineDotDotDashLine = '\003\357\000\015',
	excelMTUlUnderlineHeavyDotDotDashLine = '\003\357\000\016',
	excelMTUlUnderlineWavyLine = '\003\357\000\017',
	excelMTUlUnderlineHeavyWavyLine = '\003\357\000\020',
	excelMTUlUnderlineWavyDoubleLine = '\003\357\000\021'
};
typedef enum excelMTUl excelMTUl;

enum excelMTTA {
	excelMTTATabUnset = '\000\266\377\376',
	excelMTTALeftTab = '\000\267\000\000',
	excelMTTACenterTab = '\000\267\000\001',
	excelMTTARightTab = '\000\267\000\002',
	excelMTTADecimalTab = '\000\267\000\003'
};
typedef enum excelMTTA excelMTTA;

enum excelMTCW {
	excelMTCWCharacterWrapUnset = '\000\267\377\376',
	excelMTCWNoCharacterWrap = '\000\270\000\000',
	excelMTCWStandardCharacterWrap = '\000\270\000\001',
	excelMTCWStrictCharacterWrap = '\000\270\000\002',
	excelMTCWCustomCharacterWrap = '\000\270\000\003'
};
typedef enum excelMTCW excelMTCW;

enum excelMTFA {
	excelMTFAFontAlignmentUnset = '\000\270\377\376',
	excelMTFAAutomaticAlignment = '\000\271\000\000',
	excelMTFATopAlignment = '\000\271\000\001',
	excelMTFACenterAlignment = '\000\271\000\002',
	excelMTFABaselineAlignment = '\000\271\000\003',
	excelMTFABottomAlignment = '\000\271\000\004'
};
typedef enum excelMTFA excelMTFA;

enum excelPAtS {
	excelPAtSAutoSizeUnset = '\000\344\377\376',
	excelPAtSAutoSizeNone = '\000\345\000\000',
	excelPAtSShapeToFitText = '\000\345\000\001',
	excelPAtSTextToFitShape = '\000\345\000\002'
};
typedef enum excelPAtS excelPAtS;

enum excelMPFo {
	excelMPFoPathTypeUnset = '\000\272\377\376',
	excelMPFoNoPathType = '\000\273\000\000',
	excelMPFoPathType1 = '\000\273\000\001',
	excelMPFoPathType2 = '\000\273\000\002',
	excelMPFoPathType3 = '\000\273\000\003',
	excelMPFoPathType4 = '\000\273\000\004'
};
typedef enum excelMPFo excelMPFo;

enum excelMWFo {
	excelMWFoWarpFormatUnset = '\000\273\377\376',
	excelMWFoWarpFormat1 = '\000\274\000\000',
	excelMWFoWarpFormat2 = '\000\274\000\001',
	excelMWFoWarpFormat3 = '\000\274\000\002',
	excelMWFoWarpFormat4 = '\000\274\000\003',
	excelMWFoWarpFormat5 = '\000\274\000\004',
	excelMWFoWarpFormat6 = '\000\274\000\005',
	excelMWFoWarpFormat7 = '\000\274\000\006',
	excelMWFoWarpFormat8 = '\000\274\000\007',
	excelMWFoWarpFormat9 = '\000\274\000\010',
	excelMWFoWarpFormat10 = '\000\274\000\011',
	excelMWFoWarpFormat11 = '\000\274\000\012',
	excelMWFoWarpFormat12 = '\000\274\000\013',
	excelMWFoWarpFormat13 = '\000\274\000\014',
	excelMWFoWarpFormat14 = '\000\274\000\015',
	excelMWFoWarpFormat15 = '\000\274\000\016',
	excelMWFoWarpFormat16 = '\000\274\000\017',
	excelMWFoWarpFormat17 = '\000\274\000\020',
	excelMWFoWarpFormat18 = '\000\274\000\021',
	excelMWFoWarpFormat19 = '\000\274\000\022',
	excelMWFoWarpFormat20 = '\000\274\000\023',
	excelMWFoWarpFormat21 = '\000\274\000\024',
	excelMWFoWarpFormat22 = '\000\274\000\025',
	excelMWFoWarpFormat23 = '\000\274\000\026',
	excelMWFoWarpFormat24 = '\000\274\000\027',
	excelMWFoWarpFormat25 = '\000\274\000\030',
	excelMWFoWarpFormat26 = '\000\274\000\031',
	excelMWFoWarpFormat27 = '\000\274\000\032',
	excelMWFoWarpFormat28 = '\000\274\000\033',
	excelMWFoWarpFormat29 = '\000\274\000\034',
	excelMWFoWarpFormat30 = '\000\274\000\035',
	excelMWFoWarpFormat31 = '\000\274\000\036',
	excelMWFoWarpFormat32 = '\000\274\000\037',
	excelMWFoWarpFormat33 = '\000\274\000 ',
	excelMWFoWarpFormat34 = '\000\274\000!',
	excelMWFoWarpFormat35 = '\000\274\000\"',
	excelMWFoWarpFormat36 = '\000\274\000#'
};
typedef enum excelMWFo excelMWFo;

enum excelPCgC {
	excelPCgCCaseSentence = '\000\344\000\001',
	excelPCgCCaseLower = '\000\344\000\002',
	excelPCgCCaseUpper = '\000\344\000\003',
	excelPCgCCaseTitle = '\000\344\000\004',
	excelPCgCCaseToggle = '\000\344\000\005'
};
typedef enum excelPCgC excelPCgC;

enum excelMDTF {
	excelMDTFDateTimeFormatUnset = '\000\275\377\376',
	excelMDTFDateTimeFormatMdyy = '\000\276\000\001',
	excelMDTFDateTimeFormatDdddMMMMddyyyy = '\000\276\000\002',
	excelMDTFDateTimeFormatDMMMMyyyy = '\000\276\000\003',
	excelMDTFDateTimeFormatMMMMdyyyy = '\000\276\000\004',
	excelMDTFDateTimeFormatDMMMyy = '\000\276\000\005',
	excelMDTFDateTimeFormatMMMMyy = '\000\276\000\006',
	excelMDTFDateTimeFormatMMyy = '\000\276\000\007',
	excelMDTFDateTimeFormatMMddyyHmm = '\000\276\000\010',
	excelMDTFDateTimeFormatMMddyyhmmAMPM = '\000\276\000\011',
	excelMDTFDateTimeFormatHmm = '\000\276\000\012',
	excelMDTFDateTimeFormatHmmss = '\000\276\000\013',
	excelMDTFDateTimeFormatHmmAMPM = '\000\276\000\014',
	excelMDTFDateTimeFormatHmmssAMPM = '\000\276\000\015',
	excelMDTFDateTimeFormatFigureOut = '\000\276\000\016'
};
typedef enum excelMDTF excelMDTF;

enum excelMSET {
	excelMSETSoftEdgeUnset = '\000\277\377\376',
	excelMSETNoSoftEdge = '\000\300\000\000',
	excelMSETSoftEdgeType1 = '\000\300\000\001',
	excelMSETSoftEdgeType2 = '\000\300\000\002',
	excelMSETSoftEdgeType3 = '\000\300\000\003',
	excelMSETSoftEdgeType4 = '\000\300\000\004',
	excelMSETSoftEdgeType5 = '\000\300\000\005',
	excelMSETSoftEdgeType6 = '\000\300\000\006'
};
typedef enum excelMSET excelMSET;

enum excelMCSI {
	excelMCSIFirstDarkSchemeColor = '\000\301\000\001',
	excelMCSIFirstLightSchemeColor = '\000\301\000\002',
	excelMCSISecondDarkSchemeColor = '\000\301\000\003',
	excelMCSISecondLightSchemeColor = '\000\301\000\004',
	excelMCSIFirstAccentSchemeColor = '\000\301\000\005',
	excelMCSISecondAccentSchemeColor = '\000\301\000\006',
	excelMCSIThirdAccentSchemeColor = '\000\301\000\007',
	excelMCSIFourthAccentSchemeColor = '\000\301\000\010',
	excelMCSIFifthAccentSchemeColor = '\000\301\000\011',
	excelMCSISixthAccentSchemeColor = '\000\301\000\012',
	excelMCSIHyperlinkSchemeColor = '\000\301\000\013',
	excelMCSIFollowedHyperlinkSchemeColor = '\000\301\000\014'
};
typedef enum excelMCSI excelMCSI;

enum excelMCoI {
	excelMCoIThemeColorUnset = '\000\301\377\376',
	excelMCoINoThemeColor = '\000\302\000\000',
	excelMCoIFirstDarkThemeColor = '\000\302\000\001',
	excelMCoIFirstLightThemeColor = '\000\302\000\002',
	excelMCoISecondDarkThemeColor = '\000\302\000\003',
	excelMCoISecondLightThemeColor = '\000\302\000\004',
	excelMCoIFirstAccentThemeColor = '\000\302\000\005',
	excelMCoISecondAccentThemeColor = '\000\302\000\006',
	excelMCoIThirdAccentThemeColor = '\000\302\000\007',
	excelMCoIFourthAccentThemeColor = '\000\302\000\010',
	excelMCoIFifthAccentThemeColor = '\000\302\000\011',
	excelMCoISixthAccentThemeColor = '\000\302\000\012',
	excelMCoIHyperlinkThemeColor = '\000\302\000\013',
	excelMCoIFollowedHyperlinkThemeColor = '\000\302\000\014',
	excelMCoIFirstTextThemeColor = '\000\302\000\015',
	excelMCoIFirstBackgroundThemeColor = '\000\302\000\016',
	excelMCoISecondTextThemeColor = '\000\302\000\017',
	excelMCoISecondBackgroundThemeColor = '\000\302\000\020'
};
typedef enum excelMCoI excelMCoI;

enum excelMFLI {
	excelMFLIThemeFontLatin = '\000\303\000\001',
	excelMFLIThemeFontComplexScript = '\000\303\000\002',
	excelMFLIThemeFontHighAnsi = '\000\303\000\003',
	excelMFLIThemeFontEastAsian = '\000\303\000\004'
};
typedef enum excelMFLI excelMFLI;

enum excelMSSI {
	excelMSSIShapeStyleUnset = '\000\303\377\376',
	excelMSSIShapeNotAPreset = '\000\304\000\000',
	excelMSSIShapePreset1 = '\000\304\000\001',
	excelMSSIShapePreset2 = '\000\304\000\002',
	excelMSSIShapePreset3 = '\000\304\000\003',
	excelMSSIShapePreset4 = '\000\304\000\004',
	excelMSSIShapePreset5 = '\000\304\000\005',
	excelMSSIShapePreset6 = '\000\304\000\006',
	excelMSSIShapePreset7 = '\000\304\000\007',
	excelMSSIShapePreset8 = '\000\304\000\010',
	excelMSSIShapePreset9 = '\000\304\000\011',
	excelMSSIShapePreset10 = '\000\304\000\012',
	excelMSSIShapePreset11 = '\000\304\000\013',
	excelMSSIShapePreset12 = '\000\304\000\014',
	excelMSSIShapePreset13 = '\000\304\000\015',
	excelMSSIShapePreset14 = '\000\304\000\016',
	excelMSSIShapePreset15 = '\000\304\000\017',
	excelMSSIShapePreset16 = '\000\304\000\020',
	excelMSSIShapePreset17 = '\000\304\000\021',
	excelMSSIShapePreset18 = '\000\304\000\022',
	excelMSSIShapePreset19 = '\000\304\000\023',
	excelMSSIShapePreset20 = '\000\304\000\024',
	excelMSSIShapePreset21 = '\000\304\000\025',
	excelMSSIShapePreset22 = '\000\304\000\026',
	excelMSSIShapePreset23 = '\000\304\000\027',
	excelMSSIShapePreset24 = '\000\304\000\030',
	excelMSSIShapePreset25 = '\000\304\000\031',
	excelMSSIShapePreset26 = '\000\304\000\032',
	excelMSSIShapePreset27 = '\000\304\000\033',
	excelMSSIShapePreset28 = '\000\304\000\034',
	excelMSSIShapePreset29 = '\000\304\000\035',
	excelMSSIShapePreset30 = '\000\304\000\036',
	excelMSSIShapePreset31 = '\000\304\000\037',
	excelMSSIShapePreset32 = '\000\304\000 ',
	excelMSSIShapePreset33 = '\000\304\000!',
	excelMSSIShapePreset34 = '\000\304\000\"',
	excelMSSIShapePreset35 = '\000\304\000#',
	excelMSSIShapePreset36 = '\000\304\000$',
	excelMSSIShapePreset37 = '\000\304\000%',
	excelMSSIShapePreset38 = '\000\304\000&',
	excelMSSIShapePreset39 = '\000\304\000\'',
	excelMSSIShapePreset40 = '\000\304\000(',
	excelMSSIShapePreset41 = '\000\304\000)',
	excelMSSIShapePreset42 = '\000\304\000*',
	excelMSSILinePreset1 = '\000\304\'\021',
	excelMSSILinePreset2 = '\000\304\'\022',
	excelMSSILinePreset3 = '\000\304\'\023',
	excelMSSILinePreset4 = '\000\304\'\024',
	excelMSSILinePreset5 = '\000\304\'\025',
	excelMSSILinePreset6 = '\000\304\'\026',
	excelMSSILinePreset7 = '\000\304\'\027',
	excelMSSILinePreset8 = '\000\304\'\030',
	excelMSSILinePreset9 = '\000\304\'\031',
	excelMSSILinePreset10 = '\000\304\'\032',
	excelMSSILinePreset11 = '\000\304\'\033',
	excelMSSILinePreset12 = '\000\304\'\034',
	excelMSSILinePreset13 = '\000\304\'\035',
	excelMSSILinePreset14 = '\000\304\'\036',
	excelMSSILinePreset15 = '\000\304\'\037',
	excelMSSILinePreset16 = '\000\304\' ',
	excelMSSILinePreset17 = '\000\304\'!',
	excelMSSILinePreset18 = '\000\304\'\"',
	excelMSSILinePreset19 = '\000\304\'#',
	excelMSSILinePreset20 = '\000\304\'$',
	excelMSSILinePreset21 = '\000\304\'%'
};
typedef enum excelMSSI excelMSSI;

enum excelMBSI {
	excelMBSIBackgroundUnset = '\000\304\377\376',
	excelMBSIBackgroundNotAPreset = '\000\305\000\000',
	excelMBSIBackgroundPreset1 = '\000\305\000\001',
	excelMBSIBackgroundPreset2 = '\000\305\000\002',
	excelMBSIBackgroundPreset3 = '\000\305\000\003',
	excelMBSIBackgroundPreset4 = '\000\305\000\004',
	excelMBSIBackgroundPreset5 = '\000\305\000\005',
	excelMBSIBackgroundPreset6 = '\000\305\000\006',
	excelMBSIBackgroundPreset7 = '\000\305\000\007',
	excelMBSIBackgroundPreset8 = '\000\305\000\010',
	excelMBSIBackgroundPreset9 = '\000\305\000\011',
	excelMBSIBackgroundPreset10 = '\000\305\000\012',
	excelMBSIBackgroundPreset11 = '\000\305\000\013',
	excelMBSIBackgroundPreset12 = '\000\305\000\014'
};
typedef enum excelMBSI excelMBSI;

enum excelPDrT {
	excelPDrTTextDirectionUnset = '\000\352\377\376',
	excelPDrTLeftToRight = '\000\353\000\001',
	excelPDrTRightToLeft = '\000\353\000\002'
};
typedef enum excelPDrT excelPDrT;

enum excelPBtT {
	excelPBtTBulletTypeUnset = '\000\347\377\376',
	excelPBtTBulletTypeNone = '\000\350\000\000',
	excelPBtTBulletTypeUnnumbered = '\000\350\000\001',
	excelPBtTBulletTypeNumbered = '\000\350\000\002',
	excelPBtTPictureBulletType = '\000\350\000\003'
};
typedef enum excelPBtT excelPBtT;

enum excelPBtS {
	excelPBtSNumberedBulletStyleUnset = '\000\350\377\376',
	excelPBtSNumberedBulletStyleAlphaLowercasePeriod = '\000\351\000\000',
	excelPBtSNumberedBulletStyleAlphaUppercasePeriod = '\000\351\000\001',
	excelPBtSNumberedBulletStyleArabicRightParen = '\000\351\000\002',
	excelPBtSNumberedBulletStyleArabicPeriod = '\000\351\000\003',
	excelPBtSNumberedBulletStyleRomanLowercaseParenBoth = '\000\351\000\004',
	excelPBtSNumberedBulletStyleRomanLowercaseParenRight = '\000\351\000\005',
	excelPBtSNumberedBulletStyleRomanLowercasePeriod = '\000\351\000\006',
	excelPBtSNumberedBulletStyleRomanUppercasePeriod = '\000\351\000\007',
	excelPBtSNumberedBulletStyleAlphaLowercaseParenBoth = '\000\351\000\010',
	excelPBtSNumberedBulletStyleAlphaLowercaseParenRight = '\000\351\000\011',
	excelPBtSNumberedBulletStyleAlphaUppercaseParenBoth = '\000\351\000\012',
	excelPBtSNumberedBulletStyleAlphaUppercaseParenRight = '\000\351\000\013',
	excelPBtSNumberedBulletStyleArabicParenBoth = '\000\351\000\014',
	excelPBtSNumberedBulletStyleArabicPlain = '\000\351\000\015',
	excelPBtSNumberedBulletStyleRomanUppercaseParenBoth = '\000\351\000\016',
	excelPBtSNumberedBulletStyleRomanUppercaseParenRight = '\000\351\000\017',
	excelPBtSNumberedBulletStyleSimplifiedChinesePlain = '\000\351\000\020',
	excelPBtSNumberedBulletStyleSimplifiedChinesePeriod = '\000\351\000\021',
	excelPBtSNumberedBulletStyleCircleNumberPlain = '\000\351\000\022',
	excelPBtSNumberedBulletStyleCircleNumberWhitePlain = '\000\351\000\023',
	excelPBtSNumberedBulletStyleCircleNumberBlackPlain = '\000\351\000\024',
	excelPBtSNumberedBulletStyleTraditionalChinesePlain = '\000\351\000\025',
	excelPBtSNumberedBulletStyleTraditionalChinesePeriod = '\000\351\000\026',
	excelPBtSNumberedBulletStyleArabicAlphaDash = '\000\351\000\027',
	excelPBtSNumberedBulletStyleArabicAbjadDash = '\000\351\000\030',
	excelPBtSNumberedBulletStyleHebrewAlphaDash = '\000\351\000\031',
	excelPBtSNumberedBulletStyleKanjiKoreanPlain = '\000\351\000\032',
	excelPBtSNumberedBulletStyleKanjiKoreanPeriod = '\000\351\000\033',
	excelPBtSNumberedBulletStyleArabicDBPlain = '\000\351\000\034',
	excelPBtSNumberedBulletStyleArabicDBPeriod = '\000\351\000\035',
	excelPBtSNumberedBulletStyleThaiAlphaPeriod = '\000\351\000\036',
	excelPBtSNumberedBulletStyleThaiAlphaParenRight = '\000\351\000\037',
	excelPBtSNumberedBulletStyleThaiAlphaParenBoth = '\000\351\000 ',
	excelPBtSNumberedBulletStyleThaiNumberPeriod = '\000\351\000!',
	excelPBtSNumberedBulletStyleThaiNumberParenRight = '\000\351\000\"',
	excelPBtSNumberedBulletStyleThaiParenBoth = '\000\351\000#',
	excelPBtSNumberedBulletStyleHindiAlphaPeriod = '\000\351\000$',
	excelPBtSNumberedBulletStyleHindiNumberPeriod = '\000\351\000%',
	excelPBtSNumberedBulletStyleKanjiSimpifiedChineseDBPeriod = '\000\351\000&',
	excelPBtSNumberedBulletStyleHindiNumberParenRight = '\000\351\000\'',
	excelPBtSNumberedBulletStyleHindiAlpha1Period = '\000\351\000('
};
typedef enum excelPBtS excelPBtS;

enum excelPTSp {
	excelPTSpTabstopUnset = '\000\364\377\376',
	excelPTSpTabstopLeft = '\000\365\000\001',
	excelPTSpTabstopCenter = '\000\365\000\002',
	excelPTSpTabstopRight = '\000\365\000\003',
	excelPTSpTabstopDecimal = '\000\365\000\004'
};
typedef enum excelPTSp excelPTSp;

enum excelMRfT {
	excelMRfTReflectionUnset = '\003\351\377\376',
	excelMRfTReflectionTypeNone = '\003\352\000\000',
	excelMRfTReflectionType1 = '\003\352\000\001',
	excelMRfTReflectionType2 = '\003\352\000\002',
	excelMRfTReflectionType3 = '\003\352\000\003',
	excelMRfTReflectionType4 = '\003\352\000\004',
	excelMRfTReflectionType5 = '\003\352\000\005',
	excelMRfTReflectionType6 = '\003\352\000\006',
	excelMRfTReflectionType7 = '\003\352\000\007',
	excelMRfTReflectionType8 = '\003\352\000\010',
	excelMRfTReflectionType9 = '\003\352\000\011'
};
typedef enum excelMRfT excelMRfT;

enum excelMTtA {
	excelMTtATextureUnset = '\003\352\377\376',
	excelMTtATextureTopLeft = '\003\353\000\000',
	excelMTtATextureTop = '\003\353\000\001',
	excelMTtATextureTopRight = '\003\353\000\002',
	excelMTtATextureLeft = '\003\353\000\003',
	excelMTtATextureCenter = '\003\353\000\004',
	excelMTtATextureRight = '\003\353\000\005',
	excelMTtATextureBottomLeft = '\003\353\000\006',
	excelMTtATextureBotton = '\003\353\000\007',
	excelMTtATextureBottomRight = '\003\353\000\010'
};
typedef enum excelMTtA excelMTtA;

enum excelPBlA {
	excelPBlATextBaselineAlignmentUnset = '\003\353\377\376',
	excelPBlATextBaselineAlignBaseline = '\003\354\000\001',
	excelPBlATextBaselineAlignTop = '\003\354\000\002',
	excelPBlATextBaselineAlignCenter = '\003\354\000\003',
	excelPBlATextBaselineAlignEastAsian50 = '\003\354\000\004',
	excelPBlATextBaselineAlignAutomatic = '\003\354\000\005'
};
typedef enum excelPBlA excelPBlA;

enum excelMCbF {
	excelMCbFClipboardFormatUnset = '\003\354\377\376',
	excelMCbFNativeClipboardFormat = '\003\355\000\001',
	excelMCbFHTMlClipboardFormat = '\003\355\000\002',
	excelMCbFRTFClipboardFormat = '\003\355\000\003',
	excelMCbFPlainTextClipboardFormat = '\003\355\000\004'
};
typedef enum excelMCbF excelMCbF;

enum excelMTiP {
	excelMTiPInsertBefore = '\003\356\000\000',
	excelMTiPInsertAfter = '\003\356\000\001'
};
typedef enum excelMTiP excelMTiP;

enum excelMPiT {
	excelMPiTSaveAsDefault = '\003\362\377\376',
	excelMPiTSaveAsPNGFile = '\003\363\000\000',
	excelMPiTSaveAsBMPFile = '\003\363\000\001',
	excelMPiTSaveAsGIFFile = '\003\363\000\002',
	excelMPiTSaveAsJPGFile = '\003\363\000\003',
	excelMPiTSaveAsPDFFile = '\003\363\000\004'
};
typedef enum excelMPiT excelMPiT;

enum excelMPeT {
	excelMPeTNoEffect = '\003\364\000\000',
	excelMPeTEffectBackgroundRemoval = '\003\364\000\001',
	excelMPeTEffectBlur = '\003\364\000\002',
	excelMPeTEffectBrightnessContrast = '\003\364\000\003',
	excelMPeTEffectCement = '\003\364\000\004',
	excelMPeTEffectCrisscrossEtching = '\003\364\000\005',
	excelMPeTEffectChalkSketch = '\003\364\000\006',
	excelMPeTEffectColorTemperature = '\003\364\000\007',
	excelMPeTEffectCutout = '\003\364\000\010',
	excelMPeTEffectFilmGrain = '\003\364\000\011',
	excelMPeTEffectGlass = '\003\364\000\012',
	excelMPeTEffectGlowDiffused = '\003\364\000\013',
	excelMPeTEffectGlowEdges = '\003\364\000\014',
	excelMPeTEffectLightScreen = '\003\364\000\015',
	excelMPeTEffectLineDrawing = '\003\364\000\016',
	excelMPeTEffectMarker = '\003\364\000\017',
	excelMPeTEffectMosiaicBubbles = '\003\364\000\020',
	excelMPeTEffectPaintBrush = '\003\364\000\021',
	excelMPeTEffectPaintStrokes = '\003\364\000\022',
	excelMPeTEffectPastelsSmooth = '\003\364\000\023',
	excelMPeTEffectPencilGrayscale = '\003\364\000\024',
	excelMPeTEffectPencilSketch = '\003\364\000\025',
	excelMPeTEffectPhotocopy = '\003\364\000\026',
	excelMPeTEffectPlasticWrap = '\003\364\000\027',
	excelMPeTEffectSaturation = '\003\364\000\030',
	excelMPeTEffectSharpenSoften = '\003\364\000\031',
	excelMPeTEffectTexturizer = '\003\364\000\032',
	excelMPeTEffectWatercolorSponge = '\003\364\000\033'
};
typedef enum excelMPeT excelMPeT;

enum excelMSgT {
	excelMSgTLine = '\000\217\000\000',
	excelMSgTCurve = '\000\217\000\001'
};
typedef enum excelMSgT excelMSgT;

enum excelMEdT {
	excelMEdTAuto = '\000\220\000\000',
	excelMEdTCorner = '\000\220\000\001',
	excelMEdTSmooth = '\000\220\000\002',
	excelMEdTSymmetric = '\000\220\000\003'
};
typedef enum excelMEdT excelMEdT;

enum excelSANP {
	excelSANPDefaultNodePosition = '\003\365\000\001',
	excelSANPAfterNode = '\003\365\000\002',
	excelSANPBeforeNode = '\003\365\000\003',
	excelSANPAboveNode = '\003\365\000\004',
	excelSANPBelowNode = '\003\365\000\005'
};
typedef enum excelSANP excelSANP;

enum excelSANT {
	excelSANTDefaultNode = '\003\366\000\001',
	excelSANTAssistantNode = '\003\366\000\002'
};
typedef enum excelSANT excelSANT;

enum excelMOCT {
	excelMOCTOrgChartLayoutUnset = '\003\366\377\376',
	excelMOCTOrgChartLayoutStandard = '\003\367\000\001',
	excelMOCTOrgChartLayoutBothHanging = '\003\367\000\002',
	excelMOCTOrgChartLayoutLeftHanging = '\003\367\000\003',
	excelMOCTOrgChartLayoutRightHanging = '\003\367\000\004',
	excelMOCTOrgChartLayoutDefault = '\003\367\000\005'
};
typedef enum excelMOCT excelMOCT;

enum excelMAlC {
	excelMAlCAlignLefts = '\000\000\000\000',
	excelMAlCAlignCenters = '\000\000\000\001',
	excelMAlCAlignRights = '\000\000\000\002',
	excelMAlCAlignTops = '\000\000\000\003',
	excelMAlCAlignMiddles = '\000\000\000\004',
	excelMAlCAlignBottoms = '\000\000\000\005'
};
typedef enum excelMAlC excelMAlC;

enum excelMDsC {
	excelMDsCDistributeHorizontally = '\000\000\000\000',
	excelMDsCDistributeVertically = '\000\000\000\001'
};
typedef enum excelMDsC excelMDsC;

enum excelMOrT {
	excelMOrTOrientationUnset = '\000\214\377\376',
	excelMOrTHorizontalOrientation = '\000\215\000\001',
	excelMOrTVerticalOrientation = '\000\215\000\002'
};
typedef enum excelMOrT excelMOrT;

enum excelMZoC {
	excelMZoCBringShapeToFront = '\000p\000\000',
	excelMZoCSendShapeToBack = '\000p\000\001',
	excelMZoCBringShapeForward = '\000p\000\002',
	excelMZoCSendShapeBackward = '\000p\000\003',
	excelMZoCBringShapeInFrontOfText = '\000p\000\004',
	excelMZoCSendShapeBehindText = '\000p\000\005'
};
typedef enum excelMZoC excelMZoC;

enum excelMFlC {
	excelMFlCFlipHorizontal = '\000q\000\000',
	excelMFlCFlipVertical = '\000q\000\001'
};
typedef enum excelMFlC excelMFlC;

enum excelMTri {
	excelMTriTrue = '\000\240\377\377',
	excelMTriFalse = '\000\241\000\000',
	excelMTriCTrue = '\000\241\000\001',
	excelMTriToggle = '\000\240\377\375',
	excelMTriTriStateUnset = '\000\240\377\376'
};
typedef enum excelMTri excelMTri;

enum excelMBW {
	excelMBWBlackAndWhiteUnset = '\000\254\377\376',
	excelMBWBlackAndWhiteModeAutomatic = '\000\255\000\001',
	excelMBWBlackAndWhiteModeGrayScale = '\000\255\000\002',
	excelMBWBlackAndWhiteModeLightGrayScale = '\000\255\000\003',
	excelMBWBlackAndWhiteModeInverseGrayScale = '\000\255\000\004',
	excelMBWBlackAndWhiteModeGrayOutline = '\000\255\000\005',
	excelMBWBlackAndWhiteModeBlackTextAndLine = '\000\255\000\006',
	excelMBWBlackAndWhiteModeHighContrast = '\000\255\000\007',
	excelMBWBlackAndWhiteModeBlack = '\000\255\000\010',
	excelMBWBlackAndWhiteModeWhite = '\000\255\000\011',
	excelMBWBlackAndWhiteModeDontShow = '\000\255\000\012'
};
typedef enum excelMBW excelMBW;

enum excelMBPS {
	excelMBPSBarLeft = '\000r\000\000',
	excelMBPSBarTop = '\000r\000\001',
	excelMBPSBarRight = '\000r\000\002',
	excelMBPSBarBottom = '\000r\000\003',
	excelMBPSBarFloating = '\000r\000\004',
	excelMBPSBarPopUp = '\000r\000\005',
	excelMBPSBarMenu = '\000r\000\006'
};
typedef enum excelMBPS excelMBPS;

enum excelMBPt {
	excelMBPtNoProtection = '\000s\000\000',
	excelMBPtNoCustomize = '\000s\000\001',
	excelMBPtNoResize = '\000s\000\002',
	excelMBPtNoMove = '\000s\000\004',
	excelMBPtNoChangeVisible = '\000s\000\010',
	excelMBPtNoChangeDock = '\000s\000\020',
	excelMBPtNoVerticalDock = '\000s\000 ',
	excelMBPtNoHorizontalDock = '\000s\000@'
};
typedef enum excelMBPt excelMBPt;

enum excelMBTY {
	excelMBTYNormalCommandBar = '\000t\000\000',
	excelMBTYMenubarCommandBar = '\000t\000\001',
	excelMBTYPopupCommandBar = '\000t\000\002'
};
typedef enum excelMBTY excelMBTY;

enum excelMCLT {
	excelMCLTControlCustom = '\000u\000\000',
	excelMCLTControlButton = '\000u\000\001',
	excelMCLTControlEdit = '\000u\000\002',
	excelMCLTControlDropDown = '\000u\000\003',
	excelMCLTControlCombobox = '\000u\000\004',
	excelMCLTButtonDropDown = '\000u\000\005',
	excelMCLTSplitDropDown = '\000u\000\006',
	excelMCLTOCXDropDown = '\000u\000\007',
	excelMCLTGenericDropDown = '\000u\000\010',
	excelMCLTGraphicDropDown = '\000u\000\011',
	excelMCLTControlPopup = '\000u\000\012',
	excelMCLTGraphicPopup = '\000u\000\013',
	excelMCLTButtonPopup = '\000u\000\014',
	excelMCLTSplitButtonPopup = '\000u\000\015',
	excelMCLTSplitButtonMRUPopup = '\000u\000\016',
	excelMCLTControlLabel = '\000u\000\017',
	excelMCLTExpandingGrid = '\000u\000\020',
	excelMCLTSplitExpandingGrid = '\000u\000\021',
	excelMCLTControlGrid = '\000u\000\022',
	excelMCLTControlGauge = '\000u\000\023',
	excelMCLTGraphicCombobox = '\000u\000\024',
	excelMCLTControlPane = '\000u\000\025',
	excelMCLTActiveX = '\000u\000\026',
	excelMCLTControlGroup = '\000u\000\027',
	excelMCLTControlTab = '\000u\000\030',
	excelMCLTControlSpinner = '\000u\000\031'
};
typedef enum excelMCLT excelMCLT;

enum excelMBns {
	excelMBnsButtonStateUp = '\000v\000\000',
	excelMBnsButtonStateDown = '\000u\377\377',
	excelMBnsButtonStateUnset = '\000v\000\002'
};
typedef enum excelMBns excelMBns;

enum excelMcOu {
	excelMcOuNeither = '\000w\000\000',
	excelMcOuServer = '\000w\000\001',
	excelMcOuClient = '\000w\000\002',
	excelMcOuBoth = '\000w\000\003'
};
typedef enum excelMcOu excelMcOu;

enum excelMBTs {
	excelMBTsButtonAutomatic = '\000x\000\000',
	excelMBTsButtonIcon = '\000x\000\001',
	excelMBTsButtonCaption = '\000x\000\002',
	excelMBTsButtonIconAndCaption = '\000x\000\003'
};
typedef enum excelMBTs excelMBTs;

enum excelMXcb {
	excelMXcbComboboxStyleNormal = '\000y\000\000',
	excelMXcbComboboxStyleLabel = '\000y\000\001'
};
typedef enum excelMXcb excelMXcb;

enum excelMMNA {
	excelMMNANone = '\000{\000\000',
	excelMMNARandom = '\000{\000\001',
	excelMMNAUnfold = '\000{\000\002',
	excelMMNASlide = '\000{\000\003'
};
typedef enum excelMMNA excelMMNA;

enum excelMHlT {
	excelMHlTHyperlinkTypeTextRange = '\000\226\000\000',
	excelMHlTHyperlinkTypeShape = '\000\226\000\001',
	excelMHlTHyperlinkTypeInlineShape = '\000\226\000\002'
};
typedef enum excelMHlT excelMHlT;

enum excelMXiM {
	excelMXiMAppendString = '\000\256\000\000',
	excelMXiMPostString = '\000\256\000\001'
};
typedef enum excelMXiM excelMXiM;

enum excelMANT {
	excelMANTIdle = '\000|\000\001',
	excelMANTGreeting = '\000|\000\002',
	excelMANTGoodbye = '\000|\000\003',
	excelMANTBeginSpeaking = '\000|\000\004',
	excelMANTCharacterSuccessMajor = '\000|\000\006',
	excelMANTGetAttentionMajor = '\000|\000\013',
	excelMANTGetAttentionMinor = '\000|\000\014',
	excelMANTSearching = '\000|\000\015',
	excelMANTPrinting = '\000|\000\022',
	excelMANTGestureRight = '\000|\000\023',
	excelMANTWritingNotingSomething = '\000|\000\026',
	excelMANTWorkingAtSomething = '\000|\000\027',
	excelMANTThinking = '\000|\000\030',
	excelMANTSendingMail = '\000|\000\031',
	excelMANTListensToComputer = '\000|\000\032',
	excelMANTDisappear = '\000|\000\037',
	excelMANTAppear = '\000|\000 ',
	excelMANTGetArtsy = '\000|\000d',
	excelMANTGetTechy = '\000|\000e',
	excelMANTGetWizardy = '\000|\000f',
	excelMANTCheckingSomething = '\000|\000g',
	excelMANTLookDown = '\000|\000h',
	excelMANTLookDownLeft = '\000|\000i',
	excelMANTLookDownRight = '\000|\000j',
	excelMANTLookLeft = '\000|\000k',
	excelMANTLookRight = '\000|\000l',
	excelMANTLookUp = '\000|\000m',
	excelMANTLookUpLeft = '\000|\000n',
	excelMANTLookUpRight = '\000|\000o',
	excelMANTSaving = '\000|\000p',
	excelMANTGestureDown = '\000|\000q',
	excelMANTGestureLeft = '\000|\000r',
	excelMANTGestureUp = '\000|\000s',
	excelMANTEmptyTrash = '\000|\000t'
};
typedef enum excelMANT excelMANT;

enum excelMBSt {
	excelMBStButtonNone = '\000}\000\000',
	excelMBStButtonOk = '\000}\000\001',
	excelMBStButtonCancel = '\000}\000\002',
	excelMBStButtonsOkCancel = '\000}\000\003',
	excelMBStButtonsYesNo = '\000}\000\004',
	excelMBStButtonsYesNoCancel = '\000}\000\005',
	excelMBStButtonsBackClose = '\000}\000\006',
	excelMBStButtonsNextClose = '\000}\000\007',
	excelMBStButtonsBackNextClose = '\000}\000\010',
	excelMBStButtonsRetryCancel = '\000}\000\011',
	excelMBStButtonsAbortRetryIgnore = '\000}\000\012',
	excelMBStButtonsSearchClose = '\000}\000\013',
	excelMBStButtonsBackNextSnooze = '\000}\000\014',
	excelMBStButtonsTipsOptionsClose = '\000}\000\015',
	excelMBStButtonsYesAllNoCancel = '\000}\000\016'
};
typedef enum excelMBSt excelMBSt;

enum excelMIct {
	excelMIctIconNone = '\000~\000\000',
	excelMIctIconApplication = '\000~\000\001',
	excelMIctIconAlert = '\000~\000\002',
	excelMIctIconTip = '\000~\000\003',
	excelMIctIconAlertCritical = '\000~\000e',
	excelMIctIconAlertWarning = '\000~\000g',
	excelMIctIconAlertInfo = '\000~\000h'
};
typedef enum excelMIct excelMIct;

enum excelMWAt {
	excelMWAtInactive = '\000\202\000\000',
	excelMWAtActive = '\000\202\000\001',
	excelMWAtSuspend = '\000\202\000\002',
	excelMWAtResume = '\000\202\000\003'
};
typedef enum excelMWAt excelMWAt;

enum excelMeDP {
	excelMeDPPropertyTypeNumber = '\000\242\000\001',
	excelMeDPPropertyTypeBoolean = '\000\242\000\002',
	excelMeDPPropertyTypeDate = '\000\242\000\003',
	excelMeDPPropertyTypeString = '\000\242\000\004',
	excelMeDPPropertyTypeFloat = '\000\242\000\005'
};
typedef enum excelMeDP excelMeDP;

enum excelMASc {
	excelMAScMsoAutomationSecurityLow = '\000\243\000\001',
	excelMAScMsoAutomationSecurityByUI = '\000\243\000\002',
	excelMAScMsoAutomationSecurityForceDisable = '\000\243\000\003'
};
typedef enum excelMASc excelMASc;

enum excelMSsz {
	excelMSszResolution544x376 = '\000\204\000\000',
	excelMSszResolution640x480 = '\000\204\000\001',
	excelMSszResolution720x512 = '\000\204\000\002',
	excelMSszResolution800x600 = '\000\204\000\003',
	excelMSszResolution1024x768 = '\000\204\000\004',
	excelMSszResolution1152x882 = '\000\204\000\005',
	excelMSszResolution1152x900 = '\000\204\000\006',
	excelMSszResolution1280x1024 = '\000\204\000\007',
	excelMSszResolution1600x1200 = '\000\204\000\010',
	excelMSszResolution1800x1440 = '\000\204\000\011',
	excelMSszResolution1920x1200 = '\000\204\000\012'
};
typedef enum excelMSsz excelMSsz;

enum excelMChS {
	excelMChSArabicCharacterSet = '\000\205\000\001',
	excelMChSCyrillicCharacterSet = '\000\205\000\002',
	excelMChSEnglishCharacterSet = '\000\205\000\003',
	excelMChSGreekCharacterSet = '\000\205\000\004',
	excelMChSHebrewCharacterSet = '\000\205\000\005',
	excelMChSJapaneseCharacterSet = '\000\205\000\006',
	excelMChSKoreanCharacterSet = '\000\205\000\007',
	excelMChSMultilingualUnicodeCharacterSet = '\000\205\000\010',
	excelMChSSimplifiedChineseCharacterSet = '\000\205\000\011',
	excelMChSThaiCharacterSet = '\000\205\000\012',
	excelMChSTraditionalChineseCharacterSet = '\000\205\000\013',
	excelMChSVietnameseCharacterSet = '\000\205\000\014'
};
typedef enum excelMChS excelMChS;

enum excelMtEn {
	excelMtEnEncodingThai = '\000\213\003j',
	excelMtEnEncodingJapaneseShiftJIS = '\000\213\003\244',
	excelMtEnEncodingSimplifiedChinese = '\000\213\003\250',
	excelMtEnEncodingKorean = '\000\213\003\265',
	excelMtEnEncodingBig5TraditionalChinese = '\000\213\003\266',
	excelMtEnEncodingLittleEndian = '\000\213\004\260',
	excelMtEnEncodingBigEndian = '\000\213\004\261',
	excelMtEnEncodingCentralEuropean = '\000\213\004\342',
	excelMtEnEncodingCyrillic = '\000\213\004\343',
	excelMtEnEncodingWestern = '\000\213\004\344',
	excelMtEnEncodingGreek = '\000\213\004\345',
	excelMtEnEncodingTurkish = '\000\213\004\346',
	excelMtEnEncodingHebrew = '\000\213\004\347',
	excelMtEnEncodingArabic = '\000\213\004\350',
	excelMtEnEncodingBaltic = '\000\213\004\351',
	excelMtEnEncodingVietnamese = '\000\213\004\352',
	excelMtEnEncodingISO88591Latin1 = '\000\213o\257',
	excelMtEnEncodingISO88592CentralEurope = '\000\213o\260',
	excelMtEnEncodingISO88593Latin3 = '\000\213o\261',
	excelMtEnEncodingISO88594Baltic = '\000\213o\262',
	excelMtEnEncodingISO88595Cyrillic = '\000\213o\263',
	excelMtEnEncodingISO88596Arabic = '\000\213o\264',
	excelMtEnEncodingISO88597Greek = '\000\213o\265',
	excelMtEnEncodingISO88598Hebrew = '\000\213o\266',
	excelMtEnEncodingISO88599Turkish = '\000\213o\267',
	excelMtEnEncodingISO885915Latin9 = '\000\213o\275',
	excelMtEnEncodingISO2022JapaneseNoHalfWidthKatakana = '\000\213\304,',
	excelMtEnEncodingISO2022JapaneseJISX02021984 = '\000\213\304-',
	excelMtEnEncodingISO2022JapaneseJISX02011989 = '\000\213\304.',
	excelMtEnEncodingISO2022KR = '\000\213\3041',
	excelMtEnEncodingISO2022CNTraditionalChinese = '\000\213\3043',
	excelMtEnEncodingISO2022CNSimplifiedChinese = '\000\213\3045',
	excelMtEnEncodingMacRoman = '\000\213\'\020',
	excelMtEnEncodingMacJapanese = '\000\213\'\021',
	excelMtEnEncodingMacTraditionalChinese = '\000\213\'\022',
	excelMtEnEncodingMacKorean = '\000\213\'\023',
	excelMtEnEncodingMacArabic = '\000\213\'\024',
	excelMtEnEncodingMacHebrew = '\000\213\'\025',
	excelMtEnEncodingMacGreek1 = '\000\213\'\026',
	excelMtEnEncodingMacCyrillic = '\000\213\'\027',
	excelMtEnEncodingMacSimplifiedChineseGB2312 = '\000\213\'\030',
	excelMtEnEncodingMacRomania = '\000\213\'\032',
	excelMtEnEncodingMacUkraine = '\000\213\'!',
	excelMtEnEncodingMacLatin2 = '\000\213\'-',
	excelMtEnEncodingMacIcelandic = '\000\213\'_',
	excelMtEnEncodingMacTurkish = '\000\213\'a',
	excelMtEnEncodingMacCroatia = '\000\213\'b',
	excelMtEnEncodingEBCDICUSCanada = '\000\213\000%',
	excelMtEnEncodingEBCDICInternational = '\000\213\001\364',
	excelMtEnEncodingEBCDICMultilingualROECELatin2 = '\000\213\003f',
	excelMtEnEncodingEBCDICGreekModern = '\000\213\003k',
	excelMtEnEncodingEBCDICTurkishLatin5 = '\000\213\004\002',
	excelMtEnEncodingEBCDICGermany = '\000\213O1',
	excelMtEnEncodingEBCDICDenmarkNorway = '\000\213O5',
	excelMtEnEncodingEBCDICFinlandSweden = '\000\213O6',
	excelMtEnEncodingEBCDICItaly = '\000\213O8',
	excelMtEnEncodingEBCDICLatinAmericaSpain = '\000\213O<',
	excelMtEnEncodingEBCDICUnitedKingdom = '\000\213O=',
	excelMtEnEncodingEBCDICJapaneseKatakanaExtended = '\000\213OB',
	excelMtEnEncodingEBCDICFrance = '\000\213OI',
	excelMtEnEncodingEBCDICArabic = '\000\213O\304',
	excelMtEnEncodingEBCDICGreek = '\000\213O\307',
	excelMtEnEncodingEBCDICHebrew = '\000\213O\310',
	excelMtEnEncodingEBCDICKoreanExtended = '\000\213Qa',
	excelMtEnEncodingEBCDICThai = '\000\213Qf',
	excelMtEnEncodingEBCDICIcelandic = '\000\213Q\207',
	excelMtEnEncodingEBCDICTurkish = '\000\213Q\251',
	excelMtEnEncodingEBCDICRussian = '\000\213Q\220',
	excelMtEnEncodingEBCDICSerbianBulgarian = '\000\213R!',
	excelMtEnEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = '\000\213\306\362',
	excelMtEnEncodingEBCDICUSCanadaAndJapanese = '\000\213\306\363',
	excelMtEnEncodingEBCDICExtendedAndKorean = '\000\213\306\365',
	excelMtEnEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = '\000\213\306\367',
	excelMtEnEncodingEBCDICUSCanadaAndTraditionalChinese = '\000\213\306\371',
	excelMtEnEncodingEBCDICJapaneseLatinExtendedAndJapanese = '\000\213\306\373',
	excelMtEnEncodingOEMUnitedStates = '\000\213\001\265',
	excelMtEnEncodingOEMGreek = '\000\213\002\341',
	excelMtEnEncodingOEMBaltic = '\000\213\003\007',
	excelMtEnEncodingOEMMultilingualLatinI = '\000\213\003R',
	excelMtEnEncodingOEMMultilingualLatinII = '\000\213\003T',
	excelMtEnEncodingOEMCyrillic = '\000\213\003W',
	excelMtEnEncodingOEMTurkish = '\000\213\003Y',
	excelMtEnEncodingOEMPortuguese = '\000\213\003\\',
	excelMtEnEncodingOEMIcelandic = '\000\213\003]',
	excelMtEnEncodingOEMHebrew = '\000\213\003^',
	excelMtEnEncodingOEMCanadianFrench = '\000\213\003_',
	excelMtEnEncodingOEMArabic = '\000\213\003`',
	excelMtEnEncodingOEMNordic = '\000\213\003a',
	excelMtEnEncodingOEMCyrillicII = '\000\213\003b',
	excelMtEnEncodingOEMModernGreek = '\000\213\003e',
	excelMtEnEncodingEUCJapanese = '\000\213\312\334',
	excelMtEnEncodingEUCChineseSimplifiedChinese = '\000\213\312\340',
	excelMtEnEncodingEUCKorean = '\000\213\312\355',
	excelMtEnEncodingEUCTaiwaneseTraditionalChinese = '\000\213\312\356',
	excelMtEnEncodingDevanagari = '\000\213\336\252',
	excelMtEnEncodingBengali = '\000\213\336\253',
	excelMtEnEncodingTamil = '\000\213\336\254',
	excelMtEnEncodingTelugu = '\000\213\336\255',
	excelMtEnEncodingAssamese = '\000\213\336\256',
	excelMtEnEncodingOriya = '\000\213\336\257',
	excelMtEnEncodingKannada = '\000\213\336\260',
	excelMtEnEncodingMalayalam = '\000\213\336\261',
	excelMtEnEncodingGujarati = '\000\213\336\262',
	excelMtEnEncodingPunjabi = '\000\213\336\263',
	excelMtEnEncodingArabicASMO = '\000\213\002\304',
	excelMtEnEncodingArabicTransparentASMO = '\000\213\002\320',
	excelMtEnEncodingKoreanJohab = '\000\213\005Q',
	excelMtEnEncodingTaiwanCNS = '\000\213N ',
	excelMtEnEncodingTaiwanTCA = '\000\213N!',
	excelMtEnEncodingTaiwanEten = '\000\213N\"',
	excelMtEnEncodingTaiwanIBM5550 = '\000\213N#',
	excelMtEnEncodingTaiwanTeletext = '\000\213N$',
	excelMtEnEncodingTaiwanWang = '\000\213N%',
	excelMtEnEncodingIA5IRV = '\000\213N\211',
	excelMtEnEncodingIA5German = '\000\213N\212',
	excelMtEnEncodingIA5Swedish = '\000\213N\213',
	excelMtEnEncodingIA5Norwegian = '\000\213N\214',
	excelMtEnEncodingUSASCII = '\000\213N\237',
	excelMtEnEncodingT61 = '\000\213O%',
	excelMtEnEncodingISO6937NonspacingAccent = '\000\213O-',
	excelMtEnEncodingKOI8R = '\000\213Q\202',
	excelMtEnEncodingExtAlphaLowercase = '\000\213R#',
	excelMtEnEncodingKOI8U = '\000\213Uj',
	excelMtEnEncodingEuropa3 = '\000\213qI',
	excelMtEnEncodingHZGBSimplifiedChinese = '\000\213\316\310',
	excelMtEnEncodingUTF7 = '\000\213\375\350',
	excelMtEnEncodingUTF8 = '\000\213\375\351'
};
typedef enum excelMtEn excelMtEn;

enum excel4000 {
	excel4000CommandBar = 'msCB',
	excel4000CommandBarControl = 'mCBC'
};
typedef enum excel4000 excel4000;

enum excelE102 {
	excelE102BuiltInChart = '\001\366\000\025',
	excelE102UserDefined = '\001\366\000\026',
	excelE102AnyGallery = '\001\366\000\027'
};
typedef enum excelE102 excelE102;

enum excelE103 {
	excelE103ColorIndexAutomatic = '\001\366\357\367',
	excelE103ColorIndexNone = '\001\366\357\322',
	excelE103AColorIndexInteger = '\001\367\000\000'
};
typedef enum excelE103 excelE103;

enum excelE104 {
	excelE104Cap = '\001\370\000\001',
	excelE104NoCap = '\001\370\000\002'
};
typedef enum excelE104 excelE104;

enum excelE105 {
	excelE105ByColumns = '\001\371\000\002',
	excelE105ByRows = '\001\371\000\001'
};
typedef enum excelE105 excelE105;

enum excelE106 {
	excelE106ScaleLinear = '\001\371\357\334',
	excelE106ScaleLogarithmic = '\001\371\357\333'
};
typedef enum excelE106 excelE106;

enum excelE107 {
	excelE107AutofillSeries = '\001\373\000\004',
	excelE107ChronologicalSeries = '\001\373\000\003',
	excelE107GrowthSeries = '\001\373\000\002',
	excelE107DataSeriesLinear = '\001\372\357\334'
};
typedef enum excelE107 excelE107;

enum excelE108 {
	excelE108AxisCrossesAutomatic = '\001\373\357\367',
	excelE108AxisCrossesCustom = '\001\373\357\356',
	excelE108AxisCrossesMaximum = '\001\374\000\002',
	excelE108AxisCrossesMinimum = '\001\374\000\004'
};
typedef enum excelE108 excelE108;

enum excelE109 {
	excelE109PrimaryAxis = '\001\375\000\001',
	excelE109SecondaryAxis = '\001\375\000\002'
};
typedef enum excelE109 excelE109;

enum excelE110 {
	excelE110BackgroundAutomatic = '\001\375\357\367',
	excelE110BackgroundOpaque = '\001\376\000\003',
	excelE110BackgroundTransparent = '\001\376\000\002'
};
typedef enum excelE110 excelE110;

enum excelE111 {
	excelE111WindowStateMaximized = '\001\376\357\327',
	excelE111WindowStateMinimized = '\001\376\357\324',
	excelE111WindowStateNormal = '\001\376\357\321'
};
typedef enum excelE111 excelE111;

enum excelE112 {
	excelE112CategoryAxis = '\002\000\000\001',
	excelE112SeriesAxis = '\002\000\000\003',
	excelE112ValueAxis = '\002\000\000\002'
};
typedef enum excelE112 excelE112;

enum excelE113 {
	excelE113ArrowheadLengthLong = '\002\001\000\003',
	excelE113ArrowheadLengthMedium = '\002\000\357\326',
	excelE113ArrowheadLengthShort = '\002\001\000\001'
};
typedef enum excelE113 excelE113;

enum excelE114 {
	excelE114ValignBottom = '\002\001\357\365',
	excelE114ValignCenter = '\002\001\357\364',
	excelE114ValignDistributed = '\002\001\357\353',
	excelE114ValignJustify = '\002\001\357\336',
	excelE114ValignTop = '\002\001\357\300'
};
typedef enum excelE114 excelE114;

enum excelE115 {
	excelE115TickMarkCross = '\002\003\000\004',
	excelE115TickMarkInside = '\002\003\000\002',
	excelE115TickMarkNone = '\002\002\357\322',
	excelE115TickMarkOutside = '\002\003\000\003'
};
typedef enum excelE115 excelE115;

enum excelE116 {
	excelE116ErrorBarDirectionX = '\002\003\357\270',
	excelE116ErrorBarDirectionY = '\002\004\000\001'
};
typedef enum excelE116 excelE116;

enum excelE117 {
	excelE117ErrorBarIncludeBoth = '\002\005\000\001',
	excelE117ErrorBarIncludeMinusValues = '\002\005\000\003',
	excelE117ErrorBarIncludeNone = '\002\004\357\322',
	excelE117ErrorBarIncludePlusValues = '\002\005\000\002'
};
typedef enum excelE117 excelE117;

enum excelE118 {
	excelE118Interpolated = '\002\006\000\003',
	excelE118NotPlotted = '\002\006\000\001',
	excelE118Zero = '\002\006\000\002'
};
typedef enum excelE118 excelE118;

enum excelE119 {
	excelE119ArrowheadStyleClosed = '\002\007\000\003',
	excelE119ArrowheadStyleDoubleClosed = '\002\007\000\005',
	excelE119ArrowheadStyleDoubleOpen = '\002\007\000\004',
	excelE119ArrowheadStyleNone = '\002\006\357\322',
	excelE119ArrowheadStyleOpen = '\002\007\000\002'
};
typedef enum excelE119 excelE119;

enum excelE120 {
	excelE120ArrowheadWidthMedium = '\002\007\357\326',
	excelE120ArrowheadWidthNarrow = '\002\010\000\001',
	excelE120ArrowheadWidthWide = '\002\010\000\003'
};
typedef enum excelE120 excelE120;

enum excelE121 {
	excelE121HorizontalAlignCenter = '\002\010\357\364',
	excelE121HorizontalAlignCenterAcrossSelection = '\002\011\000\007',
	excelE121HorizontalAlignDistributed = '\002\010\357\353',
	excelE121HorizontalAlignFill = '\002\011\000\005',
	excelE121HorizontalAlignGeneral = '\002\011\000\001',
	excelE121HorizontalAlignJustify = '\002\010\357\336',
	excelE121HorizontalAlignLeft = '\002\010\357\335',
	excelE121HorizontalAlignRight = '\002\010\357\310'
};
typedef enum excelE121 excelE121;

enum excelE122 {
	excelE122TickLabelPositionHigh = '\002\011\357\341',
	excelE122TickLabelPositionLow = '\002\011\357\332',
	excelE122TickLabelPositionNextToAxis = '\002\012\000\004',
	excelE122TickLabelPositionNone = '\002\011\357\322'
};
typedef enum excelE122 excelE122;

enum excelE123 {
	excelE123LegendPositionBottom = '\002\012\357\365',
	excelE123LegendPositionCorner = '\002\013\000\002',
	excelE123LegendPositionLeft = '\002\012\357\335',
	excelE123LegendPositionRight = '\002\012\357\310',
	excelE123LegendPositionTop = '\002\012\357\300'
};
typedef enum excelE123 excelE123;

enum excelE124 {
	excelE124ChartPictureTypeStackScale = '\002\014\000\003',
	excelE124ChartPictureTypeStack = '\002\014\000\002',
	excelE124ChartPictureTypeStretch = '\002\014\000\001'
};
typedef enum excelE124 excelE124;

enum excelE125 {
	excelE125Sides = '\002\015\000\001',
	excelE125End = '\002\015\000\002',
	excelE125EndSides = '\002\015\000\003',
	excelE125Front = '\002\015\000\004',
	excelE125FrontSides = '\002\015\000\005',
	excelE125FrontEnd = '\002\015\000\006',
	excelE125AllFaces = '\002\015\000\007'
};
typedef enum excelE125 excelE125;

enum excelE126 {
	excelE126OrientationDownward = '\002\015\357\266',
	excelE126OrientationHorizontal = '\002\015\357\340',
	excelE126OrientationUpward = '\002\015\357\265',
	excelE126OrientationVertical = '\002\015\357\272'
};
typedef enum excelE126 excelE126;

enum excelE127 {
	excelE127TickLabelOrientationAutomatic = '\002\016\357\367',
	excelE127TickLabelOrientationDownward = '\002\016\357\266',
	excelE127TickLabelOrientationHorizontal = '\002\016\357\340',
	excelE127TickLabelOrientationUpward = '\002\016\357\265',
	excelE127TickLabelOrientationVertical = '\002\016\357\272'
};
typedef enum excelE127 excelE127;

enum excelE128 {
	excelE128BorderWeightHairline = '\002\020\000\001',
	excelE128BorderWeightMedium = '\002\017\357\326',
	excelE128BorderWeightThick = '\002\020\000\004',
	excelE128BorderWeightThin = '\002\020\000\002'
};
typedef enum excelE128 excelE128;

enum excelE129 {
	excelE129SeriesDateDay = '\002\021\000\001',
	excelE129SeriesDateMonth = '\002\021\000\003',
	excelE129SeriesDateWeekday = '\002\021\000\002',
	excelE129SeriesDateYear = '\002\021\000\004'
};
typedef enum excelE129 excelE129;

enum excelE130 {
	excelE130UnderlineStyleDouble = '\002\021\357\351',
	excelE130UnderlineStyleDoubleAccounting = '\002\022\000\005',
	excelE130UnderlineStyleNone = '\002\021\357\322',
	excelE130UnderlineStyleSingle = '\002\022\000\002',
	excelE130UnderlineStyleSingleAccounting = '\002\022\000\004'
};
typedef enum excelE130 excelE130;

enum excelE131 {
	excelE131ErrorBarTypeCustom = '\002\022\357\356',
	excelE131ErrorBarTypeFixedValue = '\002\023\000\001',
	excelE131ErrorBarTypePercent = '\002\023\000\002',
	excelE131ErrorBarTypeStandardDeviation = '\002\022\357\305',
	excelE131ErrorBarTypeStandardError = '\002\023\000\004'
};
typedef enum excelE131 excelE131;

enum excelE132 {
	excelE132Exponential = '\002\024\000\005',
	excelE132Linear = '\002\023\357\334',
	excelE132Logarithmic = '\002\023\357\333',
	excelE132MovingAverage = '\002\024\000\006',
	excelE132Polynomial = '\002\024\000\003',
	excelE132Power = '\002\024\000\004'
};
typedef enum excelE132 excelE132;

enum excelE133 {
	excelE133Continuous = '\002\025\000\001',
	excelE133Dash = '\002\024\357\355',
	excelE133DashDot = '\002\025\000\004',
	excelE133DashDotDot = '\002\025\000\005',
	excelE133Dot = '\002\024\357\352',
	excelE133Double = '\002\024\357\351',
	excelE133SlantDashDot = '\002\025\000\015',
	excelE133LineStyleNone = '\002\024\357\322'
};
typedef enum excelE133 excelE133;

enum excelE134 {
	excelE134DataLabelsShowNone = '\002\025\357\322',
	excelE134DataLabelsShowValue = '\002\026\000\002',
	excelE134DataLabelsShowPercent = '\002\026\000\003',
	excelE134DataLabelsShowLabel = '\002\026\000\004',
	excelE134DataLabelsShowLabelAndPercent = '\002\026\000\005',
	excelE134DataLabelsShowBubbleSizes = '\002\026\000\006'
};
typedef enum excelE134 excelE134;

enum excelE135 {
	excelE135MarkerStyleAutomatic = '\002\026\357\367',
	excelE135MarkerStyleCircle = '\002\027\000\010',
	excelE135MarkerStyleDash = '\002\026\357\355',
	excelE135MarkerStyleDiamond = '\002\027\000\002',
	excelE135MarkerStyleDot = '\002\026\357\352',
	excelE135MarkerStyleNone = '\002\026\357\322',
	excelE135MarkerStylePicture = '\002\026\357\315',
	excelE135MarkerStylePlus = '\002\027\000\011',
	excelE135MarkerStyleSquare = '\002\027\000\001',
	excelE135MarkerStyleStar = '\002\027\000\005',
	excelE135MarkerStyleTriangle = '\002\027\000\003',
	excelE135MarkerStyleX = '\002\026\357\270'
};
typedef enum excelE135 excelE135;

enum excelE137 {
	excelE137PatternAutomatic = '\002\030\357\367',
	excelE137PatternChecker = '\002\031\000\011',
	excelE137PatternCrissCross = '\002\031\000\020',
	excelE137PatternDown = '\002\030\357\347',
	excelE137PatternGray16 = '\002\031\000\021',
	excelE137PatternGray25 = '\002\030\357\344',
	excelE137PatternGray50 = '\002\030\357\343',
	excelE137PatternGray75 = '\002\030\357\342',
	excelE137PatternGray8 = '\002\031\000\022',
	excelE137PatternGrid = '\002\031\000\017',
	excelE137PatternHorizontal = '\002\030\357\340',
	excelE137PatternLightDown = '\002\031\000\015',
	excelE137PatternLightHorizontal = '\002\031\000\013',
	excelE137PatternLightUp = '\002\031\000\016',
	excelE137PatternLightVertical = '\002\031\000\014',
	excelE137PatternNone = '\002\030\357\322',
	excelE137PatternSemiGray75 = '\002\031\000\012',
	excelE137PatternSolid = '\002\031\000\001',
	excelE137PatternUp = '\002\030\357\276',
	excelE137PatternVertical = '\002\030\357\272',
	excelE137PatternLinearGradient = '\002\031\017\240',
	excelE137PatternRectangularGradient = '\002\031\017\241'
};
typedef enum excelE137 excelE137;

enum excelE138 {
	excelE138SplitByPosition = '\002\032\000\001',
	excelE138SplitByPercentValue = '\002\032\000\003',
	excelE138SplitByCustomSplit = '\002\032\000\004',
	excelE138SplitByValue = '\002\032\000\002'
};
typedef enum excelE138 excelE138;

enum excelE139 {
	excelE139Hundreds = '\002\032\377\376',
	excelE139Thousands = '\002\032\377\375',
	excelE139TenThousands = '\002\032\377\374',
	excelE139HundredThousands = '\002\032\377\373',
	excelE139Millions = '\002\032\377\372',
	excelE139TenMillions = '\002\032\377\371',
	excelE139HundredMillions = '\002\032\377\370',
	excelE139ThousandMillions = '\002\032\377\367',
	excelE139MillionMillions = '\002\032\377\366',
	excelE139CustomDisplayUnit = '\002\032\357\356'
};
typedef enum excelE139 excelE139;

enum excelE140 {
	excelE140LabelPositionCenter = '\002\033\357\364',
	excelE140LabelPositionAbove = '\002\034\000\000',
	excelE140LabelPositionBelow = '\002\034\000\001',
	excelE140LabelPositionLeft = '\002\033\357\335',
	excelE140LabelPositionRight = '\002\033\357\310',
	excelE140LabelPositionOutsideEnd = '\002\034\000\002',
	excelE140LabelPositionInsideEnd = '\002\034\000\003',
	excelE140LabelPositionInsideBase = '\002\034\000\004',
	excelE140LabelPositionBestFit = '\002\034\000\005',
	excelE140LabelPositionMixed = '\002\034\000\006',
	excelE140LabelPositionCustom = '\002\034\000\007'
};
typedef enum excelE140 excelE140;

enum excelE141 {
	excelE141Days = '\002\035\000\000',
	excelE141Months = '\002\035\000\001',
	excelE141Years = '\002\035\000\002'
};
typedef enum excelE141 excelE141;

enum excelE142 {
	excelE142CategoryScale = '\002\036\000\002',
	excelE142TimeScale = '\002\036\000\003',
	excelE142AutomaticScale = '\002\035\357\367'
};
typedef enum excelE142 excelE142;

enum excelE143 {
	excelE143Box = '\002\037\000\000',
	excelE143PyramidToPoint = '\002\037\000\001',
	excelE143PyramidToMax = '\002\037\000\002',
	excelE143Cylinder = '\002\037\000\003',
	excelE143ConeToPoint = '\002\037\000\004',
	excelE143ConeToMax = '\002\037\000\005'
};
typedef enum excelE143 excelE143;

enum excelE144 {
	excelE144ColumnClustered = '\002 \0003',
	excelE144ColumnStacked = '\002 \0004',
	excelE144ColumnStacked100 = '\002 \0005',
	excelE144ThreeDColumnClustered = '\002 \0006',
	excelE144ThreeDColumnStacked = '\002 \0007',
	excelE144ThreeDColumnStacked100 = '\002 \0008',
	excelE144BarClustered = '\002 \0009',
	excelE144BarStacked = '\002 \000:',
	excelE144BarStacked100 = '\002 \000;',
	excelE144ThreeDBarClustered = '\002 \000<',
	excelE144ThreeDBarStacked = '\002 \000=',
	excelE144ThreeDBarStacked100 = '\002 \000>',
	excelE144LineStacked = '\002 \000\?',
	excelE144LineStacked100 = '\002 \000@',
	excelE144LineMarkers = '\002 \000A',
	excelE144LineMarkersStacked = '\002 \000B',
	excelE144LineMarkersStacked100 = '\002 \000C',
	excelE144PieOfPie = '\002 \000D',
	excelE144PieExploded = '\002 \000E',
	excelE144ThreeDPieExploded = '\002 \000F',
	excelE144BarOfPie = '\002 \000G',
	excelE144XyScatterSmooth = '\002 \000H',
	excelE144XyScatterSmoothNoMarkers = '\002 \000I',
	excelE144XyScatterLines = '\002 \000J',
	excelE144XyScatterLinesNoMarkers = '\002 \000K',
	excelE144AreaStacked = '\002 \000L',
	excelE144AreaStacked100 = '\002 \000M',
	excelE144ThreeDAreaStacked = '\002 \000N',
	excelE144ThreeDAreaStacked100 = '\002 \000O',
	excelE144DoughnutExploded = '\002 \000P',
	excelE144RadarMarkers = '\002 \000Q',
	excelE144RadarFilled = '\002 \000R',
	excelE144Surface = '\002 \000S',
	excelE144SurfaceWireframe = '\002 \000T',
	excelE144SurfaceTopView = '\002 \000U',
	excelE144SurfaceTopViewWireframe = '\002 \000V',
	excelE144Bubble = '\002 \000\017',
	excelE144BubbleThreeDEffect = '\002 \000W',
	excelE144StockHLC = '\002 \000X',
	excelE144StockOHLC = '\002 \000Y',
	excelE144StockVHLC = '\002 \000Z',
	excelE144StockVOHLC = '\002 \000[',
	excelE144CylinderColumnClustered = '\002 \000\\',
	excelE144CylinderColumnStacked = '\002 \000]',
	excelE144CylinderColumnStacked100 = '\002 \000^',
	excelE144CylinderBarClustered = '\002 \000_',
	excelE144CylinderBarStacked = '\002 \000`',
	excelE144CylinderBarStacked100 = '\002 \000a',
	excelE144CylinderColumn = '\002 \000b',
	excelE144ConeColumnClustered = '\002 \000c',
	excelE144ConeColumnStacked = '\002 \000d',
	excelE144ConeColumnStacked100 = '\002 \000e',
	excelE144ConeBarClustered = '\002 \000f',
	excelE144ConeBarStacked = '\002 \000g',
	excelE144ConeBarStacked100 = '\002 \000h',
	excelE144ConeCol = '\002 \000i',
	excelE144PyramidColumnClustered = '\002 \000j',
	excelE144PyramidColumnStacked = '\002 \000k',
	excelE144PyramidColumnStacked100 = '\002 \000l',
	excelE144PyramidBarClustered = '\002 \000m',
	excelE144PyramidBarStacked = '\002 \000n',
	excelE144PyramidBarStacked100 = '\002 \000o',
	excelE144PyramidColumn = '\002 \000p',
	excelE144ThreeDColumn = '\002\037\357\374',
	excelE144LineChart = '\002 \000\004',
	excelE144ThreeDLine = '\002\037\357\373',
	excelE144ThreeDPie = '\002\037\357\372',
	excelE144PieChart = '\002 \000\005',
	excelE144Xyscatter = '\002\037\357\267',
	excelE144ThreeDArea = '\002\037\357\376',
	excelE144AreaChart = '\002 \000\001',
	excelE144Doughnut = '\002\037\357\350',
	excelE144Radar = '\002\037\357\311',
	excelE144CombinationChart = '\002\037\357\361'
};
typedef enum excelE144 excelE144;

enum excelE145 {
	excelE145DataLabel = '\002!\000\000',
	excelE145AChartArea = '\002!\000\002',
	excelE145ASeries = '\002!\000\003',
	excelE145AChartTitle = '\002!\000\004',
	excelE145Walls = '\002!\000\005',
	excelE145ACornersObject = '\002!\000\006',
	excelE145DataTable = '\002!\000\007',
	excelE145Trendline = '\002!\000\010',
	excelE145ErrorBarsObject = '\002!\000\011',
	excelE145XerrorBars = '\002!\000\012',
	excelE145YerrorBars = '\002!\000\013',
	excelE145LegendEntry = '\002!\000\014',
	excelE145LegendKey = '\002!\000\015',
	excelE145Shape = '\002!\000\016',
	excelE145MajorGridlines = '\002!\000\017',
	excelE145MinorGridlines = '\002!\000\020',
	excelE145AxisTitle = '\002!\000\021',
	excelE145UpBars = '\002!\000\022',
	excelE145PlotArea = '\002!\000\023',
	excelE145DownBars = '\002!\000\024',
	excelE145Axis = '\002!\000\025',
	excelE145SeriesLines = '\002!\000\026',
	excelE145Floor = '\002!\000\027',
	excelE145Legend = '\002!\000\030',
	excelE145HiLoLines = '\002!\000\031',
	excelE145DropLines = '\002!\000\032',
	excelE145RadarAxisLabels = '\002!\000\033',
	excelE145Nothing = '\002!\000\034',
	excelE145LeaderLines = '\002!\000\035',
	excelE145DisplayUnitLabel = '\002!\000\036'
};
typedef enum excelE145 excelE145;

enum excelE146 {
	excelE146SizeIsWidth = '\002\"\000\002',
	excelE146SizeIsArea = '\002\"\000\001'
};
typedef enum excelE146 excelE146;

enum excelE147 {
	excelE147ShiftDown = '\002\"\357\347',
	excelE147ShiftToRight = '\002\"\357\277'
};
typedef enum excelE147 excelE147;

enum excelE148 {
	excelE148ShiftToLeft = '\002#\357\301',
	excelE148ShiftUp = '\002#\357\276'
};
typedef enum excelE148 excelE148;

enum excelE149 {
	excelE149TowardTheBottom = '\002$\357\347',
	excelE149TowardTheLeft = '\002$\357\301',
	excelE149TowardTheRight = '\002$\357\277',
	excelE149TowardTheTop = '\002$\357\276'
};
typedef enum excelE149 excelE149;

enum excelE150 {
	excelE150DoAverage = '\002%\357\366',
	excelE150DoCount = '\002%\357\360',
	excelE150DoCountNumbers = '\002%\357\357',
	excelE150DoMaximum = '\002%\357\330',
	excelE150DoMinimum = '\002%\357\325',
	excelE150DoProduct = '\002%\357\313',
	excelE150DoStandardDeviation = '\002%\357\305',
	excelE150DoStandardDeviationP = '\002%\357\304',
	excelE150DoSum = '\002%\357\303',
	excelE150DoVar = '\002%\357\274',
	excelE150DoVarP = '\002%\357\273'
};
typedef enum excelE150 excelE150;

enum excelE151 {
	excelE151SheetTypeChart = '\002&\357\363',
	excelE151SheetTypeDialogSheet = '\002&\357\354',
	excelE151SheetTypeExcel4IntlMacroSheet = '\002\'\000\004',
	excelE151SheetTypeExcel4MacroSheet = '\002\'\000\003',
	excelE151SheetTypeWorksheet = '\002&\357\271'
};
typedef enum excelE151 excelE151;

enum excelE152 {
	excelE152ColumnHeader = '\002\'\357\362',
	excelE152ColumnItem = '\002(\000\005',
	excelE152DataHeader = '\002(\000\003',
	excelE152DataItem = '\002(\000\007',
	excelE152PageHeader = '\002(\000\002',
	excelE152PageItem = '\002(\000\006',
	excelE152RowHeader = '\002\'\357\307',
	excelE152RowItem = '\002(\000\004',
	excelE152TableBody = '\002(\000\010'
};
typedef enum excelE152 excelE152;

enum excelE153 {
	excelE153Formulas = '\002(\357\345',
	excelE153Comments = '\002(\357\320',
	excelE153Values = '\002(\357\275'
};
typedef enum excelE153 excelE153;

enum excelE154 {
	excelE154WindowTypeChartAsWindow = '\002*\000\005',
	excelE154WindowTypeChartInPlace = '\002*\000\004',
	excelE154WindowTypeClipboard = '\002*\000\003',
	excelE154WindowTypeInfo = '\002)\357\337',
	excelE154WindowTypeWorkbook = '\002*\000\001'
};
typedef enum excelE154 excelE154;

enum excelE155 {
	excelE155PivotFieldTypeDate = '\002+\000\002',
	excelE155PivotFieldTypeNumber = '\002*\357\317',
	excelE155PivotFieldTypeText = '\002*\357\302'
};
typedef enum excelE155 excelE155;

enum excelE156 {
	excelE156Bitmap = '\002,\000\002',
	excelE156Picture = '\002+\357\315'
};
typedef enum excelE156 excelE156;

enum excelE157 {
	excelE157Consolidation = '\002-\000\003',
	excelE157Database = '\002-\000\001',
	excelE157External = '\002-\000\002',
	excelE157PivotTable = '\002,\357\314'
};
typedef enum excelE157 excelE157;

enum excelE158 {
	excelE158A1 = '\002.\000\001',
	excelE158R1C1 = '\002-\357\312'
};
typedef enum excelE158 excelE158;

enum excelE159 {
	excelE159MicrosoftAccess = '\002/\000\004',
	excelE159MicrosoftFoxPro = '\002/\000\005',
	excelE159MicrosoftMail = '\002/\000\003',
	excelE159MicrosoftPowerPoint = '\002/\000\002',
	excelE159MicrosoftProject = '\002/\000\006',
	excelE159MicrosoftSchedulePlus = '\002/\000\007',
	excelE159MicrosoftWord = '\002/\000\001'
};
typedef enum excelE159 excelE159;

enum excelE160 {
	excelE160NoButton = '\0020\000\000',
	excelE160PrimaryButton = '\0020\000\001',
	excelE160SecondaryButton = '\0020\000\002'
};
typedef enum excelE160 excelE160;

enum excelE161 {
	excelE161CopyMode = '\0021\000\001',
	excelE161CutMode = '\0021\000\002'
};
typedef enum excelE161 excelE161;

enum excelE163 {
	excelE163FilterCopy = '\0023\000\002',
	excelE163FilterInPlace = '\0023\000\001'
};
typedef enum excelE163 excelE163;

enum excelE164 {
	excelE164DownThenOver = '\0024\000\001',
	excelE164OverThenDown = '\0024\000\002'
};
typedef enum excelE164 excelE164;

enum excelE165 {
	excelE165LinkTypeExcelLinks = '\0025\000\001',
	excelE165LinkTypeOLELinks = '\0025\000\002'
};
typedef enum excelE165 excelE165;

enum excelE166 {
	excelE166ColumnThenRow = '\0026\000\002',
	excelE166RowThenColumn = '\0026\000\001'
};
typedef enum excelE166 excelE166;

enum excelE167 {
	excelE167CancelKeyDisabled = '\0027\000\000',
	excelE167ErrorHandler = '\0027\000\002',
	excelE167Interrupt = '\0027\000\001'
};
typedef enum excelE167 excelE167;

enum excelE168 {
	excelE168PageBreakAutomatic = '\0027\357\367',
	excelE168PageBreakManual = '\0027\357\331',
	excelE168PageBreakNone = '\0028\000\000'
};
typedef enum excelE168 excelE168;

enum excelE170 {
	excelE170Landscape = '\002:\000\002',
	excelE170Portrait = '\002:\000\001'
};
typedef enum excelE170 excelE170;

enum excelE171 {
	excelE171EditionDate = '\002;\000\002',
	excelE171UpdateState = '\002;\000\001'
};
typedef enum excelE171 excelE171;

enum excelE172 {
	excelE172CommandUnderlinesAutomatic = '\002;\357\367',
	excelE172CommandUnderlinesOff = '\002;\357\316',
	excelE172CommandUnderlinesOn = '\002<\000\001'
};
typedef enum excelE172 excelE172;

enum excelE173 {
	excelE173VerbOpen = '\002=\000\002',
	excelE173VerbPrimary = '\002=\000\001'
};
typedef enum excelE173 excelE173;

enum excelE174 {
	excelE174CalculationAutomatic = '\002=\357\367',
	excelE174CalculationManual = '\002=\357\331',
	excelE174CalculationSemiautomatic = '\002>\000\002'
};
typedef enum excelE174 excelE174;

enum excelE175 {
	excelE175WorkbookReadOnly = '\002\?\000\003',
	excelE175WorkbookReadWrite = '\002\?\000\002'
};
typedef enum excelE175 excelE175;

enum excelE176 {
	excelE176FitToPage = '\002@\000\002',
	excelE176FullPage = '\002@\000\003',
	excelE176FullScreen = '\002@\000\001'
};
typedef enum excelE176 excelE176;

enum excelE177 {
	excelE177Part = '\002A\000\002',
	excelE177Whole = '\002A\000\001'
};
typedef enum excelE177 excelE177;

enum excelE178 {
	excelE178MAPI = '\002B\000\001',
	excelE178NoMailSystem = '\002B\000\000',
	excelE178PowerTalk = '\002B\000\002'
};
typedef enum excelE178 excelE178;

enum excelE179 {
	excelE179LinkInfoOlelinks = '\002C\000\002',
	excelE179LinkInfoPublishers = '\002C\000\005',
	excelE179LinkInfoSubscribers = '\002C\000\006'
};
typedef enum excelE179 excelE179;

enum excelE182 {
	excelE182CellTypeBlanks = '\002F\000\004',
	excelE182CellTypeConstants = '\002F\000\002',
	excelE182CellTypeFormulas = '\002E\357\345',
	excelE182CellTypeLastCell = '\002F\000\013',
	excelE182CellTypeComments = '\002E\357\320',
	excelE182CellTypeVisible = '\002F\000\014',
	excelE182CellTypeAllFormatConditions = '\002E\357\264',
	excelE182CellTypeSameFormatConditions = '\002E\357\263',
	excelE182CellTypeAllValidation = '\002E\357\262',
	excelE182CellTypeSameValidation = '\002E\357\261'
};
typedef enum excelE182 excelE182;

enum excelE183 {
	excelE183ArrangeStyleCascade = '\002G\000\007',
	excelE183ArrangeStyleHorizontal = '\002F\357\340',
	excelE183ArrangeStyleTiled = '\002G\000\001',
	excelE183ArrangeStyleVertical = '\002F\357\272'
};
typedef enum excelE183 excelE183;

enum excelE184 {
	excelE184IBeamCursor = '\002H\000\003',
	excelE184DefaultCursor = '\002G\357\321',
	excelE184NorthwestArrowCursor = '\002H\000\001',
	excelE184WaitCursor = '\002H\000\002'
};
typedef enum excelE184 excelE184;

enum excelE185 {
	excelE185FillCopy = '\002I\000\001',
	excelE185FillDays = '\002I\000\005',
	excelE185FillDefault = '\002I\000\000',
	excelE185FillFormats = '\002I\000\003',
	excelE185FillMonths = '\002I\000\007',
	excelE185FillSeries = '\002I\000\002',
	excelE185FillValues = '\002I\000\004',
	excelE185FillWeekdays = '\002I\000\006',
	excelE185FillYears = '\002I\000\010',
	excelE185GrowthTrend = '\002I\000\012',
	excelE185LinearTrend = '\002I\000\011'
};
typedef enum excelE185 excelE185;

enum excelE186 {
	excelE186AutofilterAnd = '\002J\000\001',
	excelE186Bottom10Items = '\002J\000\004',
	excelE186Bottom10Percent = '\002J\000\006',
	excelE186AutofilterOr = '\002J\000\002',
	excelE186Top10Items = '\002J\000\003',
	excelE186Top10Percent = '\002J\000\005',
	excelE186FilterByValue = '\002J\000\007',
	excelE186FilterByCellColor = '\002J\000\010',
	excelE186FilterByFontColor = '\002J\000\011',
	excelE186FilterByIcon = '\002J\000\012',
	excelE186FilterDynamic = '\002J\000\013',
	excelE186FilterNoFill = '\002J\000\014',
	excelE186FilterByAutomaticFontColor = '\002J\000\015',
	excelE186FilterByNoIcon = '\002J\000\016'
};
typedef enum excelE186 excelE186;

enum excelE187 {
	excelE187ClipboardFormatBiff = '\002K\000\010',
	excelE187ClipboardFormatBiff2 = '\002K\000\022',
	excelE187ClipboardFormatBiff3 = '\002K\000\024',
	excelE187ClipboardFormatBiff4 = '\002K\000\036',
	excelE187ClipboardFormatBinary = '\002K\000\017',
	excelE187ClipboardFormatBitmap = '\002K\000\011',
	excelE187ClipboardFormatCgm = '\002K\000\015',
	excelE187ClipboardFormatCsv = '\002K\000\005',
	excelE187ClipboardFormatDif = '\002K\000\004',
	excelE187ClipboardFormatDspText = '\002K\000\014',
	excelE187ClipboardFormatEmbeddedObject = '\002K\000\025',
	excelE187ClipboardFormatEmbedSource = '\002K\000\026',
	excelE187ClipboardFormatLink = '\002K\000\013',
	excelE187ClipboardFormatLinkSource = '\002K\000\027',
	excelE187ClipboardFormatLinkSourceDesc = '\002K\000 ',
	excelE187ClipboardFormatMovie = '\002K\000\030',
	excelE187ClipboardFormatNative = '\002K\000\016',
	excelE187ClipboardFormatObjectDesc = '\002K\000\037',
	excelE187ClipboardFormatObjectLink = '\002K\000\023',
	excelE187ClipboardFormatOwnerLink = '\002K\000\021',
	excelE187ClipboardFormatPict = '\002K\000\002',
	excelE187ClipboardFormatPrintPict = '\002K\000\003',
	excelE187ClipboardFormatRtf = '\002K\000\007',
	excelE187ClipboardFormatScreenPict = '\002K\000\035',
	excelE187ClipboardFormatStandardFont = '\002K\000\034',
	excelE187ClipboardFormatStandardScale = '\002K\000\033',
	excelE187ClipboardFormatSylk = '\002K\000\006',
	excelE187ClipboardFormatTable = '\002K\000\020',
	excelE187ClipboardFormatText = '\002K\000\000',
	excelE187ClipboardFormatToolFace = '\002K\000\031',
	excelE187ClipboardFormatToolFacePict = '\002K\000\032',
	excelE187ClipboardFormatValu = '\002K\000\001',
	excelE187ClipboardFormatWk1 = '\002K\000\012',
	excelE187ClipboardFormatUnicodeText = '\002K\000.',
	excelE187ClipboardFormatStyleText = '\002K\0005',
	excelE187ClipboardFormatUnicodeStyleText = '\002K\0007',
	excelE187ClipboardFormatBiff5 = '\002K\000!',
	excelE187ClipboardFormatPictureBuild = '\002K\000\"',
	excelE187ClipboardFormatOdbcConn = '\002K\000#',
	excelE187ClipboardFormatOdbcSql = '\002K\000$',
	excelE187ClipboardFormat3dPicture = '\002K\000%',
	excelE187ClipboardFormatUnexpected38 = '\002K\000&',
	excelE187ClipboardFormatDrawingDragDrop = '\002K\000\'',
	excelE187ClipboardFormatDrawing = '\002K\000(',
	excelE187ClipboardFormatUnexpected41 = '\002K\000)',
	excelE187ClipboardFormatUnexpected42 = '\002K\000*',
	excelE187ClipboardFormatUnexpected43 = '\002K\000+',
	excelE187ClipboardFormatHyperlink = '\002K\000,',
	excelE187ClipboardFormatUnexpected45 = '\002K\000-',
	excelE187ClipboardFormatWindowsBitmap = '\002K\000/',
	excelE187ClipboardFormatUniformResourceLocator = '\002K\0000',
	excelE187ClipboardFormatFileName = '\002K\0001',
	excelE187ClipboardFormatUnexpected50 = '\002K\0002',
	excelE187ClipboardFormatUnexpected51 = '\002K\0003',
	excelE187ClipboardFormatHypertextMarkupLanguage = '\002K\0004',
	excelE187ClipboardFormatOfficeScrapbookInfo = '\002K\0006',
	excelE187ClipboardFormatPortableDocumentFormat = '\002K\0008',
	excelE187ClipboardFormatExcelInternalShape = '\002K\0009',
	excelE187ClipboardFormatOfficeArtShape = '\002K\000:'
};
typedef enum excelE187 excelE187;

enum excelE188 {
	excelE188CSVFileFormat = '\002\274\000\006',
	excelE188CSVMacFileFormat = '\002\274\000\026',
	excelE188CSVMSDosFileFormat = '\002\274\000\030',
	excelE188CSVWindowsFileFormat = '\002\274\000\027',
	excelE188DBF3FileFormat = '\002\274\000\010',
	excelE188DBF4FileFormat = '\002\274\000\013',
	excelE188DIFFileFormat = '\002\274\000\011',
	excelE188Excel2FileFormat = '\002\274\000\020',
	excelE188Excel2EastAsianFileFormat = '\002\274\000\033',
	excelE188Excel3FileFormat = '\002\274\000\035',
	excelE188Excel4FileFormat = '\002\274\000!',
	excelE188Excel5FileFormat = '\002\274\000\'',
	excelE188Excel7FileFormat = '\002\274\000\'',
	excelE188Excel4WorkbookFileFormat = '\002\274\000#',
	excelE188InternationalAddInFileFormat = '\002\274\000\032',
	excelE188InternationalMacroFileFormat = '\002\274\000\031',
	excelE188WorkbookNormalFileFormat = '\002\273\357\321',
	excelE188SYLKFileFormat = '\002\274\000\002',
	excelE188CurrentPlatformTextFileFormat = '\002\273\357\302',
	excelE188TextMacFileFormat = '\002\274\000\023',
	excelE188TextMSDosFileFormat = '\002\274\000\025',
	excelE188TextPrinterFileFormat = '\002\274\000$',
	excelE188TextWindowsFileFormat = '\002\274\000\024',
	excelE188HTMLFileFormat = '\002\274\000,',
	excelE188XMLSpreadsheetFileFormat = '\002\274\000-',
	excelE188PDFFileFormat = '\002\274\000.',
	excelE188ExcelBinaryFileFormat = '\002\274\0003',
	excelE188ExcelXMLFileFormat = '\002\274\0004',
	excelE188MacroEnabledXMLFileFormat = '\002\274\0005',
	excelE188MacroEnabledTemplateFileFormat = '\002\274\0006',
	excelE188TemplateFileFormat = '\002\274\0007',
	excelE188AddInFileFormat = '\002\274\0008',
	excelE188Excel98to2004FileFormat = '\002\274\0009',
	excelE188Excel98to2004TemplateFileFormat = '\002\274\000\021',
	excelE188Excel98to2004AddInFileFormat = '\002\274\000\022'
};
typedef enum excelE188 excelE188;

enum excelE189 {
	excelE189Twenty_four_hourClock = '\002M\000!',
	excelE189FourDigitYears = '\002M\000+',
	excelE189AlternateArraySeparator = '\002M\000\020',
	excelE189ColumnSeparator = '\002M\000\016',
	excelE189Country_code = '\002M\000\001',
	excelE189Country_setting = '\002M\000\002',
	excelE189Currency_before = '\002M\000%',
	excelE189Currency_code = '\002M\000\031',
	excelE189Currency_digits = '\002M\000\033',
	excelE189Currency_leading_zeros = '\002M\000(',
	excelE189Currency_minus_sign = '\002M\000&',
	excelE189Currency_negative = '\002M\000\034',
	excelE189Currency_space_before = '\002M\000$',
	excelE189Currency_trailing_zeros = '\002M\000\'',
	excelE189Date_order = '\002M\000 ',
	excelE189Date_separator = '\002M\000\021',
	excelE189DayCode = '\002M\000\025',
	excelE189DayLeadingZero = '\002M\000*',
	excelE189DecimalSeparator = '\002M\000\003',
	excelE189GeneralFormatName = '\002M\000\032',
	excelE189HourCode = '\002M\000\026',
	excelE189LeftBrace = '\002M\000\014',
	excelE189LeftBracket = '\002M\000\012',
	excelE189ListSeparator = '\002M\000\005',
	excelE189LowerCaseColumnLetter = '\002M\000\011',
	excelE189LowerCaseRowLetter = '\002M\000\010',
	excelE189Mdy = '\002M\000,',
	excelE189Metric = '\002M\000#',
	excelE189Minute_code = '\002M\000\027',
	excelE189Month_code = '\002M\000\024',
	excelE189Month_leading_zero = '\002M\000)',
	excelE189Month_name_chars = '\002M\000\036',
	excelE189Noncurrency_digits = '\002M\000\035',
	excelE189NonEnglishFunctions = '\002M\000\"',
	excelE189RightBrace = '\002M\000\015',
	excelE189RightBracket = '\002M\000\013',
	excelE189RowSeparator = '\002M\000\017',
	excelE189SecondCode = '\002M\000\030',
	excelE189ThousandsSeparator = '\002M\000\004',
	excelE189TimeLeadingZero = '\002M\000-',
	excelE189TimeSeparator = '\002M\000\022',
	excelE189UpperCaseColumnLetter = '\002M\000\007',
	excelE189UpperCaseRowLetter = '\002M\000\006',
	excelE189Weekday_name_chars = '\002M\000\037',
	excelE189YearCode = '\002M\000\023'
};
typedef enum excelE189 excelE189;

enum excelE190 {
	excelE190PageBreakFull = '\002N\000\001',
	excelE190PageBreakPartial = '\002N\000\002'
};
typedef enum excelE190 excelE190;

enum excelE191 {
	excelE191OverwriteCells = '\002O\000\000',
	excelE191InsertDeleteCells = '\002O\000\001',
	excelE191InsertEntireRows = '\002O\000\002'
};
typedef enum excelE191 excelE191;

enum excelE192 {
	excelE192NoLabels = '\002O\357\322',
	excelE192RowLabels = '\002P\000\001',
	excelE192ColumnLabels = '\002P\000\002',
	excelE192MixedLabels = '\002P\000\003'
};
typedef enum excelE192 excelE192;

enum excelE193 {
	excelE193SinceMyLastSave = '\002Q\000\001',
	excelE193AllChanges = '\002Q\000\002',
	excelE193NotYetReviewed = '\002Q\000\003'
};
typedef enum excelE193 excelE193;

enum excelE194 {
	excelE194NoIndicator = '\002R\000\000',
	excelE194CommentIndicatorOnly = '\002Q\377\377',
	excelE194CommentAndIndicator = '\002R\000\001'
};
typedef enum excelE194 excelE194;

enum excelE195 {
	excelE195CellValue = '\002S\000\001',
	excelE195Expression = '\002S\000\002',
	excelE195ColorScale = '\002S\000\003',
	excelE195Databar = '\002S\000\004',
	excelE195Top10 = '\002S\000\005',
	excelE195IconSets = '\002S\000\006',
	excelE195UniqueValues = '\002S\000\007',
	excelE195TextString = '\002S\000\011',
	excelE195BlanksCondition = '\002S\000\012',
	excelE195TimePeriod = '\002S\000\013',
	excelE195AboveAverageCondition = '\002S\000\014',
	excelE195NoBlanksCondition = '\002S\000\015',
	excelE195ErrorsCondition = '\002S\000\020',
	excelE195NoErrorsCondition = '\002S\000\021'
};
typedef enum excelE195 excelE195;

enum excelE196 {
	excelE196OperatorBetween = '\002T\000\001',
	excelE196OperatorNotBetween = '\002T\000\002',
	excelE196OperatorEqual = '\002T\000\003',
	excelE196OperatorNotEqual = '\002T\000\004',
	excelE196OperatorGreater = '\002T\000\005',
	excelE196OperatorLess = '\002T\000\006',
	excelE196OperatorGreaterEqual = '\002T\000\007',
	excelE196OperatorLessEqual = '\002T\000\010'
};
typedef enum excelE196 excelE196;

enum excelE197 {
	excelE197NoRestrictions = '\002U\000\000',
	excelE197UnlockedCells = '\002U\000\001',
	excelE197NoSelection = '\002T\357\322'
};
typedef enum excelE197 excelE197;

enum excelE198 {
	excelE198ValidateInputOnly = '\002V\000\000',
	excelE198ValidateWholeNumber = '\002V\000\001',
	excelE198ValidateDecimal = '\002V\000\002',
	excelE198ValidateList = '\002V\000\003',
	excelE198ValidatedDate = '\002V\000\004',
	excelE198ValidateTime = '\002V\000\005',
	excelE198ValidateTextLength = '\002V\000\006',
	excelE198ValidateCustom = '\002V\000\007'
};
typedef enum excelE198 excelE198;

enum excelE199 {
	excelE199IMEModeNoControl = '\002W\000\000',
	excelE199IMEModeOn = '\002W\000\001',
	excelE199IMEModeOff = '\002W\000\002',
	excelE199IMEModeDisable = '\002W\000\003',
	excelE199IMEModeHiragana = '\002W\000\004',
	excelE199IMEModeKatakana = '\002W\000\005',
	excelE199IMEModeKatakanaHalf = '\002W\000\006',
	excelE199IMEModeAlphaFull = '\002W\000\007',
	excelE199IMEModeAlpha = '\002W\000\010',
	excelE199IMEModeHangulFull = '\002W\000\011',
	excelE199IMEModeHangul = '\002W\000\012'
};
typedef enum excelE199 excelE199;

enum excelE200 {
	excelE200ValidAlertNone = '\002W\377\377',
	excelE200ValidAlertStop = '\002X\000\001',
	excelE200ValidAlertWarning = '\002X\000\002',
	excelE200ValidAlertInformation = '\002X\000\003'
};
typedef enum excelE200 excelE200;

enum excelE201 {
	excelE201LocationAsNewSheet = '\002Y\000\001',
	excelE201LocationAsObject = '\002Y\000\002',
	excelE201LocationAutomatic = '\002Y\000\003'
};
typedef enum excelE201 excelE201;

enum excelE216 {
	excelE216Automatic = '\002Y\357\367',
	excelE216Custom = '\002Y\357\356'
};
typedef enum excelE216 excelE216;

enum excelE900 {
	excelE900PivotTableVersion2000 = '\003\204\000\000',
	excelE900PivotTableVersion10 = '\003\204\000\001',
	excelE900PivotTableVersion11 = '\003\204\000\002',
	excelE900PivotTableVersion12 = '\003\204\000\003',
	excelE900PivotTableVersion14 = '\003\204\000\004',
	excelE900PivotTableVersionCurrent = '\003\203\377\377'
};
typedef enum excelE900 excelE900;

enum excelE901 {
	excelE901CompactRow = '\003\205\000\000',
	excelE901TabularRow = '\003\205\000\001',
	excelE901OutlineRow = '\003\205\000\002'
};
typedef enum excelE901 excelE901;

enum excelE902 {
	excelE902AtTop = '\003\206\000\001',
	excelE902AtBottom = '\003\206\000\002'
};
typedef enum excelE902 excelE902;

enum excelE903 {
	excelE903ManualAllocation = '\003\207\000\001',
	excelE903AutomaticAllocation = '\003\207\000\002'
};
typedef enum excelE903 excelE903;

enum excelE904 {
	excelE904AllocateValue = '\003\210\000\001',
	excelE904AllocateIncrement = '\003\210\000\002'
};
typedef enum excelE904 excelE904;

enum excelE905 {
	excelE905EqualAllocation = '\003\211\000\001',
	excelE905WeightAllocation = '\003\211\000\002'
};
typedef enum excelE905 excelE905;

enum excelE906 {
	excelE906DoNotRepeatLabels = '\003\212\000\001',
	excelE906RepeatLabels = '\003\212\000\002'
};
typedef enum excelE906 excelE906;

enum excelE907 {
	excelE907MissingItemsDefault = '\003\212\377\377',
	excelE907MissingItemsNone = '\003\213\000\000',
	excelE907MissingItemsMax = '\003\213~\364',
	excelE907MissingItemsMax2 = '\003\233\000\000'
};
typedef enum excelE907 excelE907;

enum excelE908 {
	excelE908PivotCellValue = '\003\214\000\000',
	excelE908PivotCellPivotItem = '\003\214\000\001',
	excelE908PivotCellSubtotal = '\003\214\000\002',
	excelE908PivotCellGrandTotal = '\003\214\000\003',
	excelE908PivotCellDataField = '\003\214\000\004',
	excelE908PivotCellPivotField = '\003\214\000\005',
	excelE908PivotCellPageFieldItem = '\003\214\000\006',
	excelE908PivotCellCustomSubtotal = '\003\214\000\007',
	excelE908PivotCellDataPivotField = '\003\214\000\010',
	excelE908PivotCellBlankCell = '\003\214\000\011'
};
typedef enum excelE908 excelE908;

enum excelE909 {
	excelE909CellNotChanged = '\003\215\000\001',
	excelE909CellChanged = '\003\215\000\002',
	excelE909CellChangeApplied = '\003\215\000\003'
};
typedef enum excelE909 excelE909;

enum excelE910 {
	excelE910Tabular = '\003\216\000\000',
	excelE910Outline = '\003\216\000\001'
};
typedef enum excelE910 excelE910;

enum excelE911 {
	excelE911PivotTopCount = '\003\217\000\001',
	excelE911PivotBottomCount = '\003\217\000\002',
	excelE911PivotTopPercent = '\003\217\000\003',
	excelE911PivotBottomPercent = '\003\217\000\004',
	excelE911PivotTopSum = '\003\217\000\005',
	excelE911PivotBottomSum = '\003\217\000\006',
	excelE911PivotValueEquals = '\003\217\000\007',
	excelE911PivotValueIsNotEqual = '\003\217\000\010',
	excelE911PivotValueIsGreaterThan = '\003\217\000\011',
	excelE911PivotValueIsGreaterThanOrEqualTo = '\003\217\000\012',
	excelE911PivotValueIsLessThan = '\003\217\000\013',
	excelE911PivotValueIsLessThanOrEqualTo = '\003\217\000\014',
	excelE911PivotValueIsBetween = '\003\217\000\015',
	excelE911PivotValueIsNotBetween = '\003\217\000\016',
	excelE911PivotCaptionEquals = '\003\217\000\017',
	excelE911PivotCaptionDoesNotEqual = '\003\217\000\020',
	excelE911PivotCaptionBeginsWith = '\003\217\000\021',
	excelE911PivotCaptionDoesNotBeginWith = '\003\217\000\022',
	excelE911PivotCaptionEndsWith = '\003\217\000\023',
	excelE911PivotCaptionDoesNotEndWith = '\003\217\000\024',
	excelE911PivotCaptionContains = '\003\217\000\025',
	excelE911PivotCaptionDoesNotContain = '\003\217\000\026',
	excelE911PivotCaptionIsGreaterThan = '\003\217\000\027',
	excelE911PivotCaptionIsGreaterThanOrEqualTo = '\003\217\000\030',
	excelE911PivotCaptionIsLessThan = '\003\217\000\031',
	excelE911PivotCaptionIsLessThanOrEqualTo = '\003\217\000\032',
	excelE911PivotCaptionIsBetween = '\003\217\000\033',
	excelE911PivotCaptionIsNowBetween = '\003\217\000\034',
	excelE911PivotSpecificDate = '\003\217\000\035',
	excelE911PivotNotSpecificDate = '\003\217\000\036',
	excelE911PivotBefore = '\003\217\000\037',
	excelE911PivotBeforeOrEqualTo = '\003\217\000 ',
	excelE911PivotAfter = '\003\217\000!',
	excelE911PivotAfterOrEqualTo = '\003\217\000\"',
	excelE911PivotBetween = '\003\217\000#',
	excelE911PivotNotBetween = '\003\217\000$',
	excelE911PivotTomorrow = '\003\217\000%',
	excelE911PivotToday = '\003\217\000&',
	excelE911PivotYesterday = '\003\217\000\'',
	excelE911PivotNextWeek = '\003\217\000(',
	excelE911PivotThisWeek = '\003\217\000)',
	excelE911PivotLastWeek = '\003\217\000*',
	excelE911PivotNextMonth = '\003\217\000+',
	excelE911PivotThisMonth = '\003\217\000,',
	excelE911PivotLastMonth = '\003\217\000-',
	excelE911PivotNextQuarter = '\003\217\000.',
	excelE911PivotThisQuarter = '\003\217\000/',
	excelE911PivotLastQuarter = '\003\217\0000',
	excelE911PivotNextYear = '\003\217\0001',
	excelE911PivotThisYear = '\003\217\0002',
	excelE911PivotLastYear = '\003\217\0003',
	excelE911PivotYearToDate = '\003\217\0004',
	excelE911PivotAllDatesInPeriodQuarter1 = '\003\217\0005',
	excelE911PivotAllDatesInPeriodQuarter2 = '\003\217\0006',
	excelE911PivotAllDatesInPeriodQuarter3 = '\003\217\0007',
	excelE911PivotAllDatesInPeriodQuarter4 = '\003\217\0008',
	excelE911PivotAllDatesInPeriodJanuary = '\003\217\0009',
	excelE911PivotAllDatesInPeriodFeberary = '\003\217\000:',
	excelE911PivotAllDatesInPeriodMarch = '\003\217\000;',
	excelE911PivotAllDatesInPeriodApril = '\003\217\000<',
	excelE911PivotAllDatesInPeriodMay = '\003\217\000=',
	excelE911PivotAllDatesInPeriodJune = '\003\217\000>',
	excelE911PivotAllDatesInPeriodJuly = '\003\217\000\?',
	excelE911PivotAllDatesInPeriodAugust = '\003\217\000@',
	excelE911PivotAllDatesInPeriodSeptember = '\003\217\000A',
	excelE911PivotAllDatesInPeriodOctober = '\003\217\000B',
	excelE911PivotAllDatesInPeriodNovember = '\003\217\000C',
	excelE911PivotAllDatesInPeriodDecember = '\003\217\000D'
};
typedef enum excelE911 excelE911;

enum excelE912 {
	excelE912PivotLineRegular = '\003\220\000\000',
	excelE912PivotLineSubtotal = '\003\220\000\001',
	excelE912PivotLineGrandtotal = '\003\220\000\002',
	excelE912PivotLineBlank = '\003\220\000\003'
};
typedef enum excelE912 excelE912;

enum excelE913 {
	excelE913Hierarchy = '\003\221\000\001',
	excelE913Measure = '\003\221\000\002',
	excelE913Set = '\003\221\000\003'
};
typedef enum excelE913 excelE913;

enum excelE914 {
	excelE914CubeHierarchy = '\003\222\000\001',
	excelE914CubeMeasure = '\003\222\000\002',
	excelE914CubeSet = '\003\222\000\003',
	excelE914CubeAttribute = '\003\222\000\004',
	excelE914CubeCalculatedMeasure = '\003\222\000\005',
	excelE914CubeKPIValue = '\003\222\000\006',
	excelE914CubeKPIGoal = '\003\222\000\007',
	excelE914CubeKPIStatus = '\003\222\000\010',
	excelE914CubeKPITrend = '\003\222\000\011',
	excelE914CubeKPIWeight = '\003\222\000\012'
};
typedef enum excelE914 excelE914;

enum excelE915 {
	excelE915DisplayPropertyInPivotTable = '\003\223\000\001',
	excelE915DisplayPropertyInTooltip = '\003\223\000\002',
	excelE915DisplayPropertyInPivotTableAndTooltip = '\003\223\000\003'
};
typedef enum excelE915 excelE915;

enum excelE916 {
	excelE916CalculatedMember = '\003\224\000\000',
	excelE916CalculatedSet = '\003\224\000\001'
};
typedef enum excelE916 excelE916;

enum excelE917 {
	excelE917ConnectionTypeOLEDB = '\003\225\000\001',
	excelE917ConnectionTypeODBC = '\003\225\000\002',
	excelE917ConnectionTypeXMLMAP = '\003\225\000\003',
	excelE917ConnectionTypeTEXT = '\003\225\000\004',
	excelE917ConnectionTypeWEB = '\003\225\000\005'
};
typedef enum excelE917 excelE917;

enum excelE203 {
	excelE203PasteSpecialOperationAdd = '\002[\000\002',
	excelE203PasteSpecialOperationDivide = '\002[\000\005',
	excelE203PasteSpecialOperationMultiply = '\002[\000\004',
	excelE203PasteSpecialOperationNone = '\002Z\357\322',
	excelE203PasteSpecialOperationSubtract = '\002[\000\003'
};
typedef enum excelE203 excelE203;

enum excelE204 {
	excelE204PasteAll = '\002[\357\370',
	excelE204PasteAllUsingSourceTheme = '\002\\\000\015',
	excelE204PasteAllExceptBorders = '\002\\\000\007',
	excelE204PasteFormats = '\002[\357\346',
	excelE204PasteFormulas = '\002[\357\345',
	excelE204PasteComments = '\002[\357\320',
	excelE204PasteValues = '\002[\357\275',
	excelE204PasteColumnWidths = '\002\\\000\010',
	excelE204PasteValidation = '\002\\\000\006',
	excelE204PasteFormulasAndNumberFormats = '\002\\\000\013',
	excelE204PasteValuesAndNumberFormats = '\002\\\000\014'
};
typedef enum excelE204 excelE204;

enum excelE205 {
	excelE205PhoneticCharacterHalfWidthKatakana = '\002]\000\000',
	excelE205PhoneticCharacterFullWidthKatakana = '\002]\000\001',
	excelE205PhoneticCharacterHiragana = '\002]\000\002',
	excelE205NoPhoneticCharacterConversion = '\002]\000\003'
};
typedef enum excelE205 excelE205;

enum excelE206 {
	excelE206PhoneticAlignNoControl = '\002^\000\000',
	excelE206PhoneticAlignLeft = '\002^\000\001',
	excelE206PhoneticAlignCenter = '\002^\000\002',
	excelE206PhoneticAlignDistributed = '\002^\000\003'
};
typedef enum excelE206 excelE206;

enum excelE207 {
	excelE207Printer = '\002_\000\002',
	excelE207Screen = '\002_\000\001'
};
typedef enum excelE207 excelE207;

enum excelE208 {
	excelE208OrientAsColumnField = '\002`\000\002',
	excelE208OrientAsDataField = '\002`\000\004',
	excelE208OrientAsHidden = '\002`\000\000',
	excelE208OrientAsPageField = '\002`\000\003',
	excelE208OrientAsRowField = '\002`\000\001'
};
typedef enum excelE208 excelE208;

enum excelE209 {
	excelE209PivotFieldCalculationDifferenceFrom = '\002a\000\002',
	excelE209PivotFieldCalculationIndex = '\002a\000\011',
	excelE209PivotFieldCalculationNoAdditionalCalculation = '\002`\357\321',
	excelE209PivotFieldCalculationPercentDifferenceFrom = '\002a\000\004',
	excelE209PivotFieldCalculationPercentOf = '\002a\000\003',
	excelE209PivotFieldCalculationPercentOfColumn = '\002a\000\007',
	excelE209PivotFieldCalculationPercentOfRow = '\002a\000\006',
	excelE209PivotFieldCalculationPercentOfTotal = '\002a\000\010',
	excelE209PivotFieldCalculationRunningTotal = '\002a\000\005'
};
typedef enum excelE209 excelE209;

enum excelE210 {
	excelE210PlacementFreeFloating = '\002b\000\003',
	excelE210PlacementMove = '\002b\000\002',
	excelE210PlacementMoveAndSize = '\002b\000\001'
};
typedef enum excelE210 excelE210;

enum excelE211 {
	excelE211Macintosh = '\002c\000\001',
	excelE211MSDos = '\002c\000\003',
	excelE211MSWindows = '\002c\000\002'
};
typedef enum excelE211 excelE211;

enum excelE212 {
	excelE212PrintSheetEnd = '\002d\000\001',
	excelE212PrintInPlace = '\002d\000\020',
	excelE212PrintNoComments = '\002c\357\322'
};
typedef enum excelE212 excelE212;

enum excelE213 {
	excelE213PriorityHigh = '\002d\357\341',
	excelE213PriorityLow = '\002d\357\332',
	excelE213PriorityNormal = '\002d\357\321'
};
typedef enum excelE213 excelE213;

enum excelE214 {
	excelE214SelectionModeLabelOnly = '\002f\000\001',
	excelE214SelectionModeDataAndLabel = '\002f\000\000',
	excelE214SelectionModeDataOnly = '\002f\000\002',
	excelE214SelectionModeOrigin = '\002f\000\003',
	excelE214SelectionModeButton = '\002f\000\017',
	excelE214SelectionModeBlanks = '\002f\000\004'
};
typedef enum excelE214 excelE214;

enum excelE215 {
	excelE215RangeAutoformatThreeDEffects1 = '\002g\000\015',
	excelE215RangeAutoformatThreeDEffects2 = '\002g\000\016',
	excelE215RangeAutoformatAccounting1 = '\002g\000\004',
	excelE215RangeAutoformatAccounting2 = '\002g\000\005',
	excelE215RangeAutoformatAccounting3 = '\002g\000\006',
	excelE215RangeAutoformatAccounting4 = '\002g\000\021',
	excelE215RangeAutoformatClassic1 = '\002g\000\001',
	excelE215RangeAutoformatClassic2 = '\002g\000\002',
	excelE215RangeAutoformatClassic3 = '\002g\000\003',
	excelE215RangeAutoformatColor1 = '\002g\000\007',
	excelE215RangeAutoformatColor2 = '\002g\000\010',
	excelE215RangeAutoformatColor3 = '\002g\000\011',
	excelE215RangeAutoformatList1 = '\002g\000\012',
	excelE215RangeAutoformatList2 = '\002g\000\013',
	excelE215RangeAutoformatList3 = '\002g\000\014',
	excelE215RangeAutoformatLocalFormat1 = '\002g\000\017',
	excelE215RangeAutoformatLocalFormat2 = '\002g\000\020',
	excelE215RangeAutoformatLocalFormat3 = '\002g\000\023',
	excelE215RangeAutoformatLocalFormat4 = '\002g\000\024',
	excelE215RangeAutoformatNone = '\002f\357\322',
	excelE215RangeAutoformatSimple = '\002f\357\306'
};
typedef enum excelE215 excelE215;

enum excelE217 {
	excelE217AllAtOnce = '\002i\000\002',
	excelE217OneAfterAnother = '\002i\000\001'
};
typedef enum excelE217 excelE217;

enum excelE218 {
	excelE218NotYetRouted = '\002j\000\000',
	excelE218RoutingComplete = '\002j\000\002',
	excelE218RoutingInProgress = '\002j\000\001'
};
typedef enum excelE218 excelE218;

enum excelE219 {
	excelE219AutoActivate = '\002k\000\003',
	excelE219AutoClose = '\002k\000\002',
	excelE219AutoDeactivate = '\002k\000\004',
	excelE219AutoOpen = '\002k\000\001'
};
typedef enum excelE219 excelE219;

enum excelE221 {
	excelE221Exclusive = '\002m\000\003',
	excelE221NoChange = '\002m\000\001',
	excelE221Shared = '\002m\000\002'
};
typedef enum excelE221 excelE221;

enum excelE222 {
	excelE222LocalSessionChanges = '\002n\000\002',
	excelE222OtherSessionChanges = '\002n\000\003',
	excelE222UserResolution = '\002n\000\001'
};
typedef enum excelE222 excelE222;

enum excelE223 {
	excelE223SearchNext = '\002o\000\001',
	excelE223SearchPrevious = '\002o\000\002'
};
typedef enum excelE223 excelE223;

enum excelE224 {
	excelE224ByColumns = '\002p\000\002',
	excelE224ByRows = '\002p\000\001'
};
typedef enum excelE224 excelE224;

enum excelE225 {
	excelE225SheetVisible = '\002p\377\377',
	excelE225SheetHidden = '\002q\000\000',
	excelE225SheetVeryHidden = '\002q\000\002'
};
typedef enum excelE225 excelE225;

enum excelE226 {
	excelE226PinYin = '\002r\000\001' /* Phonetic Chinese/Japanese sort order for characters. This is the default value. */,
	excelE226Stroke = '\002r\000\002' /* Sort by the quantity of strokes in each character. */
};
typedef enum excelE226 excelE226;

enum excelE228 {
	excelE228SortAscending = '\002t\000\001' /* Sorts the specified field in ascending order. This is the default value. */,
	excelE228SortDescending = '\002t\000\002' /* Sorts the specified field in descending order. */,
	excelE228SortManual = '\002s\357\331' /* It is not supported. */
};
typedef enum excelE228 excelE228;

enum excelE229 {
	excelE229SortRows = '\002u\000\002' /* Sorts by row. this is the default value. */,
	excelE229SortColumns = '\002u\000\001' /* Sorts by column. */
};
typedef enum excelE229 excelE229;

enum excelE230 {
	excelE230SortLabels = '\002v\000\002' /* Sorts the PivotTable report by labels. */,
	excelE230SortValues = '\002v\000\001' /* Sorts the PivotTable report by values. */
};
typedef enum excelE230 excelE230;

enum excelE231 {
	excelE231Errors = '\002w\000\020',
	excelE231Logical = '\002w\000\004',
	excelE231Numbers = '\002w\000\001',
	excelE231TextValues = '\002w\000\002'
};
typedef enum excelE231 excelE231;

enum excelE232 {
	excelE232SummaryAbove = '\002x\000\000',
	excelE232SummaryBelow = '\002x\000\001'
};
typedef enum excelE232 excelE232;

enum excelE233 {
	excelE233SummaryOnLeft = '\002x\357\335',
	excelE233SummaryOnRight = '\002x\357\310'
};
typedef enum excelE233 excelE233;

enum excelE234 {
	excelE234SummaryPivotTable = '\002y\357\314',
	excelE234StandardSummary = '\002z\000\001'
};
typedef enum excelE234 excelE234;

enum excelE236 {
	excelE236Delimited = '\002|\000\001',
	excelE236FixedWidth = '\002|\000\002'
};
typedef enum excelE236 excelE236;

enum excelE237 {
	excelE237TextQualifierDoubleQuote = '\002}\000\001',
	excelE237TextQualifierNone = '\002|\357\322',
	excelE237TextQualifierSingleQuote = '\002}\000\002'
};
typedef enum excelE237 excelE237;

enum excelE238 {
	excelE238Chart = '\002}\357\363',
	excelE238Excel4IntlMacroSheet = '\002~\000\004',
	excelE238Excel4MacroSheet = '\002~\000\003',
	excelE238Worksheet = '\002}\357\271'
};
typedef enum excelE238 excelE238;

enum excelE239 {
	excelE239NormalView = '\002\177\000\001',
	excelE239PageLayoutView = '\002\177\000\003'
};
typedef enum excelE239 excelE239;

enum excelE240 {
	excelE240MacroTypeCommand = '\002\200\000\002',
	excelE240MacroTypeFunction = '\002\200\000\001',
	excelE240MacroTypeNotXLM = '\002\200\000\003'
};
typedef enum excelE240 excelE240;

enum excelE241 {
	excelE241HeaderGuess = '\002\201\000\000' /* Default value. Excel determines whether there is a header, and where it is, if there is one. */,
	excelE241HeaderNo = '\002\201\000\002' /* The entire range should be sorted. */,
	excelE241HeaderYes = '\002\201\000\001' /* The entire range should not be sorted. */
};
typedef enum excelE241 excelE241;

enum excelE242 {
	excelE242DisplayShapes = '\002\201\357\370',
	excelE242Hide = '\002\202\000\003',
	excelE242Placeholders = '\002\202\000\002'
};
typedef enum excelE242 excelE242;

enum excelE243 {
	excelE243InsideHorizontal = '\002\203\000\014',
	excelE243InsideVertical = '\002\203\000\013',
	excelE243DiagonalDown = '\002\203\000\005',
	excelE243DiagonalUp = '\002\203\000\006',
	excelE243EdgeBottom = '\002\203\000\011',
	excelE243EdgeLeft = '\002\203\000\007',
	excelE243EdgeRight = '\002\203\000\012',
	excelE243EdgeTop = '\002\203\000\010',
	excelE243BorderBottom = '\002\202\357\365',
	excelE243BorderLeft = '\002\202\357\335',
	excelE243BorderRight = '\002\202\357\310',
	excelE243BorderTop = '\002\202\357\300'
};
typedef enum excelE243 excelE243;

enum excelE244 {
	excelE244NoButtonChanges = '\002\204\000\001',
	excelE244NoChanges = '\002\204\000\004',
	excelE244NoDockingChanges = '\002\204\000\003',
	excelE244ToolbarProtectionNone = '\002\203\357\321',
	excelE244NoShapeChanges = '\002\204\000\002'
};
typedef enum excelE244 excelE244;

enum excelE245 {
	excelE245DialogOpen = '\002\205\000\001',
	excelE245DialogOpenLinks = '\002\205\000\002',
	excelE245DialogSaveAs = '\002\205\000\005',
	excelE245DialogFileDelete = '\002\205\000\006',
	excelE245DialogPageSetup = '\002\205\000\007',
	excelE245DialogPrint = '\002\205\000\010',
	excelE245DialogPrinterSetup = '\002\205\000\011',
	excelE245DialogArrangeAll = '\002\205\000\014',
	excelE245DialogWindowSize = '\002\205\000\015',
	excelE245DialogWindowMove = '\002\205\000\016',
	excelE245DialogRun = '\002\205\000\021',
	excelE245DialogSetPrintTitles = '\002\205\000\027',
	excelE245DialogFont = '\002\205\000\032',
	excelE245DialogDisplay = '\002\205\000\033',
	excelE245DialogProtectDocument = '\002\205\000\034',
	excelE245DialogCalculation = '\002\205\000 ',
	excelE245DialogExtract = '\002\205\000#',
	excelE245DialogDataDelete = '\002\205\000$',
	excelE245DialogSort = '\002\205\000\'',
	excelE245DialogDataSeries = '\002\205\000(',
	excelE245DialogTable = '\002\205\000)',
	excelE245DialogFormatNumber = '\002\205\000*',
	excelE245DialogAlignment = '\002\205\000+',
	excelE245DialogStyle = '\002\205\000,',
	excelE245DialogBorder = '\002\205\000-',
	excelE245DialogCellProtection = '\002\205\000.',
	excelE245DialogColumnWidth = '\002\205\000/',
	excelE245DialogClear = '\002\205\0004',
	excelE245DialogPasteSpecial = '\002\205\0005',
	excelE245DialogEditDelete = '\002\205\0006',
	excelE245DialogInsert = '\002\205\0007',
	excelE245DialogPasteNames = '\002\205\000:',
	excelE245DialogDefineName = '\002\205\000=',
	excelE245DialogCreateNames = '\002\205\000>',
	excelE245DialogFormulaGoto = '\002\205\000\?',
	excelE245DialogFormulaFind = '\002\205\000@',
	excelE245DialogGalleryArea = '\002\205\000C',
	excelE245DialogGalleryBar = '\002\205\000D',
	excelE245DialogGalleryColumn = '\002\205\000E',
	excelE245DialogGalleryLine = '\002\205\000F',
	excelE245DialogGalleryPie = '\002\205\000G',
	excelE245DialogGalleryScatter = '\002\205\000H',
	excelE245DialogCombination = '\002\205\000I',
	excelE245DialogGridlines = '\002\205\000L',
	excelE245DialogAxes = '\002\205\000N',
	excelE245DialogAttachText = '\002\205\000P',
	excelE245DialogPatterns = '\002\205\000T',
	excelE245DialogMainChart = '\002\205\000U',
	excelE245DialogOverlay = '\002\205\000V',
	excelE245DialogScale = '\002\205\000W',
	excelE245DialogFormatLegend = '\002\205\000X',
	excelE245DialogFormatText = '\002\205\000Y',
	excelE245DialogParse = '\002\205\000[',
	excelE245DialogUnhide = '\002\205\000^',
	excelE245DialogWorkspace = '\002\205\000_',
	excelE245DialogActivate = '\002\205\000g',
	excelE245DialogCopyPicture = '\002\205\000l',
	excelE245DialogDeleteName = '\002\205\000n',
	excelE245DialogDeleteFormat = '\002\205\000o',
	excelE245DialogNew = '\002\205\000w',
	excelE245DialogRowHeight = '\002\205\000\177',
	excelE245DialogFormatMove = '\002\205\000\200',
	excelE245DialogFormatSize = '\002\205\000\201',
	excelE245DialogFormulaReplace = '\002\205\000\202',
	excelE245DialogSelectSpecial = '\002\205\000\204',
	excelE245DialogApplyNames = '\002\205\000\205',
	excelE245DialogReplaceFont = '\002\205\000\206',
	excelE245DialogSplit = '\002\205\000\211',
	excelE245DialogOutline = '\002\205\000\216',
	excelE245DialogSaveWorkbook = '\002\205\000\221',
	excelE245DialogCopyChart = '\002\205\000\223',
	excelE245DialogFormatFont = '\002\205\000\226',
	excelE245DialogNote = '\002\205\000\232',
	excelE245DialogSetUpdateStatus = '\002\205\000\237',
	excelE245DialogColorPalette = '\002\205\000\241',
	excelE245DialogChangeLink = '\002\205\000\246',
	excelE245DialogAppMove = '\002\205\000\252',
	excelE245DialogAppSize = '\002\205\000\253',
	excelE245DialogMainChartType = '\002\205\000\271',
	excelE245DialogOverlayChartType = '\002\205\000\272',
	excelE245DialogOpenMail = '\002\205\000\274',
	excelE245DialogSendMail = '\002\205\000\275',
	excelE245DialogStandardFont = '\002\205\000\276',
	excelE245DialogConsolidate = '\002\205\000\277',
	excelE245DialogSortSpecial = '\002\205\000\300',
	excelE245DialogGalleryThreeDArea = '\002\205\000\301',
	excelE245DialogGalleryThreeDColumn = '\002\205\000\302',
	excelE245DialogGalleryThreeDLine = '\002\205\000\303',
	excelE245DialogGalleryThreeDPie = '\002\205\000\304',
	excelE245DialogViewThreeD = '\002\205\000\305',
	excelE245DialogGoalSeek = '\002\205\000\306',
	excelE245DialogWorkgroup = '\002\205\000\307',
	excelE245DialogFillGroup = '\002\205\000\310',
	excelE245DialogUpdateLink = '\002\205\000\311',
	excelE245DialogPromote = '\002\205\000\312',
	excelE245DialogDemote = '\002\205\000\313',
	excelE245DialogShowDetail = '\002\205\000\314',
	excelE245DialogObjectProperties = '\002\205\000\317',
	excelE245DialogSaveNewObject = '\002\205\000\320',
	excelE245DialogApplyStyle = '\002\205\000\324',
	excelE245DialogAssignToObject = '\002\205\000\325',
	excelE245DialogObjectProtection = '\002\205\000\326',
	excelE245DialogShowToolbar = '\002\205\000\334',
	excelE245DialogPrintPreview = '\002\205\000\336',
	excelE245DialogEditColor = '\002\205\000\337',
	excelE245DialogFormatMain = '\002\205\000\341',
	excelE245DialogFormatOverlay = '\002\205\000\342',
	excelE245DialogEditSeries = '\002\205\000\344',
	excelE245DialogDefineStyle = '\002\205\000\345',
	excelE245DialogGalleryRadar = '\002\205\000\371',
	excelE245DialogZoom = '\002\205\001\000',
	excelE245DialogInsertObject = '\002\205\001\003',
	excelE245DialogSize = '\002\205\001\005',
	excelE245DialogMove = '\002\205\001\006',
	excelE245DialogFormatAuto = '\002\205\001\015',
	excelE245DialogGalleryThreeDBar = '\002\205\001\020',
	excelE245DialogGalleryThreeDSurface = '\002\205\001\021',
	excelE245DialogCustomizeToolbar = '\002\205\001\024',
	excelE245DialogWorkbookAdd = '\002\205\001\031',
	excelE245DialogWorkbookMove = '\002\205\001\032',
	excelE245DialogWorkbookCopy = '\002\205\001\033',
	excelE245DialogWorkbookOptions = '\002\205\001\034',
	excelE245DialogSaveWorkspace = '\002\205\001\035',
	excelE245DialogChartWizard = '\002\205\001 ',
	excelE245DialogAssignToTool = '\002\205\001%',
	excelE245DialogPlacement = '\002\205\001,',
	excelE245DialogFillWorkgroup = '\002\205\001-',
	excelE245DialogWorkbookNew = '\002\205\001.',
	excelE245DialogScenarioCells = '\002\205\0011',
	excelE245DialogScenarioAdd = '\002\205\0013',
	excelE245DialogScenarioEdit = '\002\205\0014',
	excelE245DialogScenarioSummary = '\002\205\0017',
	excelE245DialogPivotTableWizard = '\002\205\0018',
	excelE245DialogPivotFieldProperties = '\002\205\0019',
	excelE245DialogOptionsCalculation = '\002\205\001>',
	excelE245DialogOptionsEdit = '\002\205\001\?',
	excelE245DialogOptionsView = '\002\205\001@',
	excelE245DialogAddInManager = '\002\205\001A',
	excelE245DialogMenuEditor = '\002\205\001B',
	excelE245DialogAttachToolbars = '\002\205\001C',
	excelE245DialogOptionsChart = '\002\205\001E',
	excelE245DialogVbaInsertFile = '\002\205\001H',
	excelE245DialogVbaProcedureDefinition = '\002\205\001J',
	excelE245DialogRoutingSlip = '\002\205\001P',
	excelE245DialogMailLogon = '\002\205\001S',
	excelE245DialogInsertPicture = '\002\205\001V',
	excelE245DialogGalleryDoughnut = '\002\205\001X',
	excelE245DialogChartTrend = '\002\205\001^',
	excelE245DialogWorkbookInsert = '\002\205\001b',
	excelE245DialogOptionsTransition = '\002\205\001c',
	excelE245DialogOptionsGeneral = '\002\205\001d',
	excelE245DialogFilterAdvanced = '\002\205\001r',
	excelE245DialogMailNextLetter = '\002\205\001z',
	excelE245DialogDataLabel = '\002\205\001{',
	excelE245DialogInsertTitle = '\002\205\001|',
	excelE245DialogFontProperties = '\002\205\001}',
	excelE245DialogMacroOptions = '\002\205\001~',
	excelE245DialogWorkbookUnhide = '\002\205\001\200',
	excelE245DialogWorkbookName = '\002\205\001\202',
	excelE245DialogGalleryCustom = '\002\205\001\204',
	excelE245DialogAddChartAutoformat = '\002\205\001\206',
	excelE245DialogChartAddData = '\002\205\001\210',
	excelE245DialogTabOrder = '\002\205\001\212',
	excelE245DialogSubtotalCreate = '\002\205\001\216',
	excelE245DialogWorkbookTabSplit = '\002\205\001\237',
	excelE245DialogWorkbookProtect = '\002\205\001\241',
	excelE245DialogScrollbarProperties = '\002\205\001\244',
	excelE245DialogPivotShowPages = '\002\205\001\245',
	excelE245DialogTextToColumns = '\002\205\001\246',
	excelE245DialogFormatCharttype = '\002\205\001\247',
	excelE245DialogPivotFieldGroup = '\002\205\001\261',
	excelE245DialogPivotFieldUngroup = '\002\205\001\262',
	excelE245DialogCheckboxProperties = '\002\205\001\263',
	excelE245DialogLabelProperties = '\002\205\001\264',
	excelE245DialogListboxProperties = '\002\205\001\265',
	excelE245DialogEditboxProperties = '\002\205\001\266',
	excelE245DialogOpenText = '\002\205\001\271',
	excelE245DialogPushbuttonProperties = '\002\205\001\275',
	excelE245DialogFilter = '\002\205\001\277',
	excelE245DialogFunctionWizard = '\002\205\001\302',
	excelE245DialogSaveCopyAs = '\002\205\001\310',
	excelE245DialogOptionsListsAdd = '\002\205\001\312',
	excelE245DialogSeriesAxes = '\002\205\001\314',
	excelE245DialogSeriesX = '\002\205\001\315',
	excelE245DialogSeriesY = '\002\205\001\316',
	excelE245DialogErrorbarX = '\002\205\001\317',
	excelE245DialogErrorbarY = '\002\205\001\320',
	excelE245DialogFormatChart = '\002\205\001\321',
	excelE245DialogSeriesOrder = '\002\205\001\322',
	excelE245DialogMailEditMailer = '\002\205\001\326',
	excelE245DialogStandardWidth = '\002\205\001\330',
	excelE245DialogScenarioMerge = '\002\205\001\331',
	excelE245DialogProperties = '\002\205\001\332',
	excelE245DialogSummaryInfo = '\002\205\001\332',
	excelE245DialogFindFile = '\002\205\001\333',
	excelE245DialogActiveCellFont = '\002\205\001\334',
	excelE245DialogVbaMakeAddIn = '\002\205\001\336',
	excelE245DialogFileSharing = '\002\205\001\341',
	excelE245DialogAutocorrect = '\002\205\001\345',
	excelE245DialogCustomViews = '\002\205\001\355',
	excelE245DialogInsertNameLabel = '\002\205\001\360',
	excelE245DialogSeriesShape = '\002\205\001\370',
	excelE245DialogChartOptionsDataLabels = '\002\205\001\371',
	excelE245DialogChartOptionsDataTable = '\002\205\001\372',
	excelE245DialogSetBackgroundPicture = '\002\205\001\375',
	excelE245DialogDataValidation = '\002\205\002\015',
	excelE245DialogChartType = '\002\205\002\016',
	excelE245DialogChartLocation = '\002\205\002\017',
	excelE245DialogChartSourceData = '\002\205\002\035',
	excelE245DialogSeriesOptions = '\002\205\002-',
	excelE245DialogPivotTableOptions = '\002\205\0027',
	excelE245DialogPivotSolveOrder = '\002\205\0028',
	excelE245DialogPivotCalculatedField = '\002\205\002:',
	excelE245DialogPivotCalculatedItem = '\002\205\002<',
	excelE245DialogConditionalFormatting = '\002\205\002G',
	excelE245DialogInsertHyperlink = '\002\205\002T',
	excelE245DialogProtectSharing = '\002\205\002l',
	excelE245DialogPhonetic = '\002\205\002\213',
	excelE245DialogImportTextFile = '\002\205\002\232',
	excelE245DialogWebOptionsGeneral = '\002\205\002\264',
	excelE245DialogWebOptionsPictures = '\002\205\002\266',
	excelE245DialogWebOptionsFiles = '\002\205\002\265',
	excelE245DialogWebOptionsFonts = '\002\205\002\270',
	excelE245DialogWebOptionsEncoding = '\002\205\002\267'
};
typedef enum excelE245 excelE245;

enum excelE246 {
	excelE246Prompt = '\002\206\000\000',
	excelE246Constant = '\002\206\000\001',
	excelE246Range = '\002\206\000\002'
};
typedef enum excelE246 excelE246;

enum excelE247 {
	excelE247ParamTypeUnknown = '\002\207\000\000',
	excelE247ParamTypeChar = '\002\207\000\001',
	excelE247ParamTypeNumeric = '\002\207\000\002',
	excelE247ParamTypeDecimal = '\002\207\000\003',
	excelE247ParamTypeNumber = '\002\207\000\004',
	excelE247ParamTypeSmallInt = '\002\207\000\005',
	excelE247ParamTypeFloat = '\002\207\000\006',
	excelE247ParamTypeReal = '\002\207\000\007',
	excelE247ParamTypeDouble = '\002\207\000\010',
	excelE247ParamTypeVarChar = '\002\207\000\014',
	excelE247ParamTypeDate = '\002\207\000\011',
	excelE247ParamTypeTime = '\002\207\000\012',
	excelE247ParamTypeTimestamp = '\002\207\000\013',
	excelE247ParamTypeLongVarChar = '\002\206\377\377',
	excelE247ParamTypeBinary = '\002\206\377\376',
	excelE247ParamTypeVarBinary = '\002\206\377\375',
	excelE247ParamTypeLongVarBinary = '\002\206\377\374',
	excelE247ParamTypeBigInt = '\002\206\377\373',
	excelE247ParamTypeTinyInt = '\002\206\377\372',
	excelE247ParamTypeBit = '\002\206\377\371'
};
typedef enum excelE247 excelE247;

enum excelE248 {
	excelE248ButtonControl = '\002\210\000\000',
	excelE248CheckBox = '\002\210\000\001',
	excelE248DropDown = '\002\210\000\002',
	excelE248EditBox = '\002\210\000\003',
	excelE248GroupBox = '\002\210\000\004',
	excelE248Label = '\002\210\000\005',
	excelE248ListBox = '\002\210\000\006',
	excelE248OptionButton = '\002\210\000\007',
	excelE248ScrollBar = '\002\210\000\010',
	excelE248Spinner = '\002\210\000\011'
};
typedef enum excelE248 excelE248;

enum excelE249 {
	excelE249GeneralFormat = '\002\211\000\001',
	excelE249TextFormat = '\002\211\000\002',
	excelE249MDYFormat = '\002\211\000\003',
	excelE249DMYFormat = '\002\211\000\004',
	excelE249YMDFormat = '\002\211\000\005',
	excelE249MYDFormat = '\002\211\000\006',
	excelE249DYMFormat = '\002\211\000\007',
	excelE249YDMFormat = '\002\211\000\010',
	excelE249SkipColumn = '\002\211\000\011'
};
typedef enum excelE249 excelE249;

enum excelE250 {
	excelE250ODBCQuery = '\002\212\000\001',
	excelE250DAORecordSet = '\002\212\000\002',
	excelE250WebQuery = '\002\212\000\004',
	excelE250OLEDBQuery = '\002\212\000\005',
	excelE250TextImport = '\002\212\000\006',
	excelE250ADORecordset = '\002\212\000\007',
	excelE250FileMakerQuery = '\002\212\000\010'
};
typedef enum excelE250 excelE250;

enum excelE251 {
	excelE251CmdCube = '\002\213\000\001',
	excelE251CmdSql = '\002\213\000\002',
	excelE251CmdTable = '\002\213\000\003',
	excelE251CmdDefault = '\002\213\000\004',
	excelE251CmdList = '\002\213\000\005'
};
typedef enum excelE251 excelE251;

enum excelE253 {
	excelE253SrcNone = '\002\215\000\001',
	excelE253SrcRange = '\002\215\000\002',
	excelE253SrcExternal = '\002\215\000\003'
};
typedef enum excelE253 excelE253;

enum excelE257 {
	excelE257CriteriaEquals = '\002\221\000\000',
	excelE257CriteriaLessThanOrEqualTo = '\002\221\000\001',
	excelE257CriteriaGreaterThanOrEqualTo = '\002\221\000\002',
	excelE257CriteriaLessThan = '\002\221\000\003',
	excelE257CriteriaGreaterThan = '\002\221\000\004',
	excelE257CriteriaBeginsWith = '\002\221\000\005',
	excelE257CriteriaEndsWith = '\002\221\000\006',
	excelE257CriteriaContains = '\002\221\000\007'
};
typedef enum excelE257 excelE257;

enum excelE258 {
	excelE258NoCondition = '\002\222\000\000',
	excelE258AndCondition = '\002\222\000\001',
	excelE258OrCondition = '\002\222\000\002'
};
typedef enum excelE258 excelE258;

enum excelE259 {
	excelE259RangeValueDefault = '\002\223\000\012',
	excelE259RangeValueXMLSpreadsheet = '\002\223\000\013',
	excelE259RangeValueMSPersistXML = '\002\223\000\014'
};
typedef enum excelE259 excelE259;

enum excelE260 {
	excelE260Inches = '\002\224\000\001',
	excelE260Centimeters = '\002\224\000\002',
	excelE260Millimeters = '\002\224\000\003'
};
typedef enum excelE260 excelE260;

enum excelE261 {
	excelE261SubtotalAutomatic = '\002\225\000\001',
	excelE261SubtotalSum = '\002\225\000\002',
	excelE261SubtotalCount = '\002\225\000\003',
	excelE261SubtotalAverage = '\002\225\000\004',
	excelE261SubtotalMax = '\002\225\000\005',
	excelE261SubtotalMin = '\002\225\000\006',
	excelE261SubtotalProduct = '\002\225\000\007',
	excelE261SubtotalCountNumbers = '\002\225\000\010',
	excelE261SubtotalStandardDeviation = '\002\225\000\011',
	excelE261SubtotalStandardDeviationP = '\002\225\000\012',
	excelE261SubtotalVariable = '\002\225\000\013',
	excelE261SubtotalVariableP = '\002\225\000\014'
};
typedef enum excelE261 excelE261;

enum excelE262 {
	excelE262DataEntryOn = '\002\226\000\001',
	excelE262DataEntryStrict = '\002\226\000\002',
	excelE262DataEntryOff = '\002\225\357\316'
};
typedef enum excelE262 excelE262;

enum excelE263 {
	excelE263StatusText = '\002\226\377\377',
	excelE263ABoolean = '\002\227\000\000'
};
typedef enum excelE263 excelE263;

enum excelE264 {
	excelE264ExcelMenus = '\002\230\000\001'
};
typedef enum excelE264 excelE264;

enum excelE265 {
	excelE265LeftToRight = '\002\230\354u',
	excelE265RightToLeft = '\002\230\354t',
	excelE265Context = '\002\230\354v'
};
typedef enum excelE265 excelE265;

enum excelE266 {
	excelE266NormalCursor = '\002\232\000\000',
	excelE266LogicalCursor = '\002\232\000\001',
	excelE266VisualCursor = '\002\232\000\002'
};
typedef enum excelE266 excelE266;

enum excelE267 {
	excelE267RangeObject = '\002\233\000\001' /* range object */,
	excelE267A1StyleRangeReference = '\002\233\000\002' /* range R1C1 reference */,
	excelE267NamedRange = '\002\233\000\003' /* range R1C1 reference */
};
typedef enum excelE267 excelE267;

enum excelE268 {
	excelE268AutomaticSubtotal = '\002\234\000\001',
	excelE268SumSubtotal = '\002\234\000\002',
	excelE268CountSubtotal = '\002\234\000\003',
	excelE268AverageSubtotal = '\002\234\000\004',
	excelE268MaximumValue = '\002\234\000\005',
	excelE268MinimumValue = '\002\234\000\006',
	excelE268ProductSubtotal = '\002\234\000\007',
	excelE268CountNumbersSubtotal = '\002\234\000\010',
	excelE268StandardDeviation = '\002\234\000\011',
	excelE268StandardDeviationP = '\002\234\000\012',
	excelE268VarianceSubtotal = '\002\234\000\013',
	excelE268VariancePSubtotal = '\002\234\000\014'
};
typedef enum excelE268 excelE268;

enum excelE269 {
	excelE269Type_automatic = '\002\234\357\367',
	excelE269Type_manual = '\002\234\357\331'
};
typedef enum excelE269 excelE269;

enum excelE270 {
	excelE270PositionTop = '\002\235\357\300',
	excelE270PositionBottom = '\002\235\357\365'
};
typedef enum excelE270 excelE270;

enum excelE271 {
	excelE271ScrollTabPositionFirst = '\002\237\000\000',
	excelE271ScrollTabPositionLast = '\002\237\000\001'
};
typedef enum excelE271 excelE271;

enum excelE272 {
	excelE272Range = '\002\240\000\000',
	excelE272AListOfRanges = '\002\240\000\001',
	excelE272ReportName = '\002\240\000\002',
	excelE272AListOfStringThatIsASQLQuery = '\002\240\000\003'
};
typedef enum excelE272 excelE272;

enum excelE273 {
	excelE273BuiltInChartTemplate = '\002\241\000\001',
	excelE273FormatName = '\002\241\000\002'
};
typedef enum excelE273 excelE273;

enum excelE274 {
	excelE274BuiltInChartType = '\002\242\000\025',
	excelE274CustomChart = '\002\241\357\356'
};
typedef enum excelE274 excelE274;

enum excelE275 {
	excelE275RangeObject = '\002\243\000\001' /* range object */,
	excelE275A1StyleRangeReference = '\002\243\000\002' /* range R1C1 reference */,
	excelE275NamedRange = '\002\243\000\003' /* range R1C1 reference */,
	excelE275ListOfStrings = '\002\243\000\004'
};
typedef enum excelE275 excelE275;

enum excelE276 {
	excelE276RangeObject = '\002\244\000\001' /* range object */,
	excelE276A1StyleRangeReference = '\002\244\000\002' /* range R1C1 reference */,
	excelE276NamedRange = '\002\244\000\003' /* range R1C1 reference */,
	excelE276InputDefaultAsString = '\002\244\000\004'
};
typedef enum excelE276 excelE276;

enum excelE277 {
	excelE277ANumber = '\002\245\000\001' /* range object */,
	excelE277InputTypeAsString = '\002\245\000\002' /* range R1C1 reference */,
	excelE277ANumberOrAString = '\002\245\000\003' /* range R1C1 reference */,
	excelE277ABool = '\002\245\000\004',
	excelE277RangeObject = '\002\245\000\010',
	excelE277ListOfNumbers = '\002\245\000A',
	excelE277ListOfStrings = '\002\245\000B',
	excelE277ListOfNumberOrString = '\002\245\000C',
	excelE277ListOfBools = '\002\245\000D',
	excelE277ListOfRangeObjects = '\002\245\000H'
};
typedef enum excelE277 excelE277;

enum excelE278 {
	excelE278ANumber = '\002\246\000\001',
	excelE278ABool = '\002\246\000\004'
};
typedef enum excelE278 excelE278;

enum excelE279 {
	excelE279RangeObject = '\002\247\000\001' /* range object */,
	excelE279A1StyleRangeReference = '\002\247\000\002' /* range R1C1 reference */,
	excelE279NamedRange = '\002\247\000\003' /* range R1C1 reference */,
	excelE279ListOfStrings = '\002\247\000\004' /* A list of SQL query strings */
};
typedef enum excelE279 excelE279;

enum excelE280 {
	excelE280Percentable = '\002\250\000\001' /* A percentage between 10 and 400 */,
	excelE280ABool = '\002\250\000\004'
};
typedef enum excelE280 excelE280;

enum excelE281 {
	excelE281Script = '\002\251\000\001' /* A script object */,
	excelE281ScriptText = '\002\251\000\002'
};
typedef enum excelE281 excelE281;

enum excelE282 {
	excelE282Application = '\002\252\000\001',
	excelE282Worksheet = '\002\252\000\002',
	excelE282A1StyleRangeReference = '\002\252\000\003'
};
typedef enum excelE282 excelE282;

enum excelE283 {
	excelE283HorizontalAligmentBottom = '\002\252\357\365',
	excelE283HorizontalAligmentLeft = '\002\252\357\335',
	excelE283HorizontalAligmentRight = '\002\252\357\310',
	excelE283HorizontalAligmentTop = '\002\252\357\300'
};
typedef enum excelE283 excelE283;

enum excelE284 {
	excelE284VerticalAlignmentTop = '\002\253\357\300',
	excelE284VerticalAlignmentCenter = '\002\253\357\364',
	excelE284VerticalAlignmentBottom = '\002\253\357\365',
	excelE284VerticalAlignmentJustify = '\002\253\357\336',
	excelE284VerticalAlignmentDistributed = '\002\253\357\353'
};
typedef enum excelE284 excelE284;

enum excelE285 {
	excelE285CheckboxOff = '\002\254\357\316',
	excelE285CheckboxOn = '\002\255\000\001',
	excelE285CheckboxMixed = '\002\255\000\002'
};
typedef enum excelE285 excelE285;

enum excelE286 {
	excelE286Text = '\002\255\357\302',
	excelE286ANumber = '\002\256\000\002',
	excelE286XlNumber = '\002\255\357\317',
	excelE286Reference = '\002\256\000\004',
	excelE286Formula = '\002\256\000\005'
};
typedef enum excelE286 excelE286;

enum excelE290 {
	excelE290SelectNone = '\002\261\357\322',
	excelE290SelectSimple = '\002\261\357\306',
	excelE290SelectExtended = '\002\262\000\003'
};
typedef enum excelE290 excelE290;

enum excelE291 {
	excelE291TextToReplace = '\002\263\000\001',
	excelE291ReplacementText = '\002\263\000\002'
};
typedef enum excelE291 excelE291;

enum excelE292 {
	excelE292RangeObject = '\002\264\000\001' /* range object */,
	excelE292A1StyleRangeReference = '\002\264\000\002' /* range R1C1 reference */,
	excelE292NamedRange = '\002\264\000\003' /* range R1C1 reference */,
	excelE292ListOfCategoryNames = '\002\264\000\004' /* A list category names */
};
typedef enum excelE292 excelE292;

enum excelE294 {
	excelE294DoNotUpdateLinks = '\002\266\000\000',
	excelE294UpdateExternalLinksOnly = '\002\266\000\001',
	excelE294UpdateRemoteLinksOnly = '\002\266\000\002',
	excelE294UpdateRemoteAndExternalLinks = '\002\266\000\003'
};
typedef enum excelE294 excelE294;

enum excelE295 {
	excelE295TabDelimiter = '\002\267\000\001',
	excelE295CommasDelimiter = '\002\267\000\002',
	excelE295SpacesDelimiter = '\002\267\000\003',
	excelE295SemicolonDelimiter = '\002\267\000\004',
	excelE295NoDelimiter = '\002\267\000\005',
	excelE295CustomCharacterDelimiter = '\002\267\000\006'
};
typedef enum excelE295 excelE295;

enum excelE296 {
	excelE296VaryByColor = '\002\270\000\001',
	excelE296VaryByShade = '\002\270\000\002',
	excelE296VaryByGrayscale = '\002\270\000\003',
	excelE296VaryBySameColor = '\002\270\000\004'
};
typedef enum excelE296 excelE296;

enum excelE297 {
	excelE297RangeObject = '\002\271\000\001' /* range object */,
	excelE297A1StyleRangeReference = '\002\271\000\002' /* range R1C1 reference */,
	excelE297NamedRange = '\002\271\000\003' /* range R1C1 reference */
};
typedef enum excelE297 excelE297;

enum excelE298 {
	excelE298WorksheetObject = '\002\272\000\001' /* worksheet object */,
	excelE298WorksheetName = '\002\272\000\002' /* worksheet name */
};
typedef enum excelE298 excelE298;

enum excelE299 {
	excelE299AlignTickLabelCenter = '\002\272\357\364',
	excelE299AlignTickLabelLeft = '\002\272\357\335',
	excelE299AlignTickLabelRight = '\002\272\357\310'
};
typedef enum excelE299 excelE299;

enum excelE300 {
	excelE300Basque = '\002\274\004-',
	excelE300Catalan = '\002\274\004\003',
	excelE300Chinese = '\002\274\010\004',
	excelE300ChineseTaiwan = '\002\274\004\004',
	excelE300Czech = '\002\274\004\005',
	excelE300Danish = '\002\274\004\006',
	excelE300Dutch = '\002\274\004\023',
	excelE300EnglishUS = '\002\274\004\011',
	excelE300EnglishAUS = '\002\274\014\011',
	excelE300EnglishBritish = '\002\274\010\011',
	excelE300EnglishCAN = '\002\274\020\011',
	excelE300Finnish = '\002\274\004\013',
	excelE300French = '\002\274\004\014',
	excelE300FrenchCanadian = '\002\274\014\014',
	excelE300German = '\002\274\004\007',
	excelE300GermanAustria = '\002\274\014\007',
	excelE300SwissGerman = '\002\274\010\007',
	excelE300Greek = '\002\274\004\010',
	excelE300Hungarian = '\002\274\004\016',
	excelE300Italian = '\002\274\004\020',
	excelE300Japanese = '\002\274\004\021',
	excelE300Korean = '\002\274\004\022',
	excelE300Malaysian = '\002\274\004>',
	excelE300NorwegianBokmal = '\002\274\004\024',
	excelE300Norwegian = '\002\274\004,',
	excelE300Polish = '\002\274\004\025',
	excelE300PortugueseBrazil = '\002\274\004\026',
	excelE300PortugueseIberian = '\002\274\010\026',
	excelE300Russian = '\002\274\004\031',
	excelE300Slovak = '\002\274\004\033',
	excelE300Slovenian = '\002\274\004$',
	excelE300Spanish = '\002\274\004\012',
	excelE300Swedish = '\002\274\004\035',
	excelE300Turkish = '\002\274\004\037'
};
typedef enum excelE300 excelE300;

enum excelE301 {
	excelE301SortOnCellValue = '\002\275\000\000' /* Values. */,
	excelE301SortOnCellColor = '\002\275\000\001' /* Cell color. */,
	excelE301SortOnFontColor = '\002\275\000\002' /* Font color. */,
	excelE301SortOnIcon = '\002\275\000\003' /* Icon. */
};
typedef enum excelE301 excelE301;

enum excelE302 {
	excelE302SortNormal = '\002\276\000\000' /* Default. Sorts numeric and text data separately. */,
	excelE302SortTextAsNumbers = '\002\276\000\001' /* Treat text as numeric data for the sort. */
};
typedef enum excelE302 excelE302;

enum excelE303 {
	excelE303NoneTotalsCalc = '\002\277\000\000',
	excelE303AverageTotalsCalc = '\002\277\000\001',
	excelE303CountTotalsCalc = '\002\277\000\002',
	excelE303CountNumberTotalsCalc = '\002\277\000\003',
	excelE303MaxTotalsCalc = '\002\277\000\004',
	excelE303MinTotalsCalc = '\002\277\000\005',
	excelE303SumTotalsCalc = '\002\277\000\006',
	excelE303DeviationTotalsCalc = '\002\277\000\007',
	excelE303VarTotalsCalc = '\002\277\000\010',
	excelE303CustomTotalsCalc = '\002\277\000\011'
};
typedef enum excelE303 excelE303;

enum excelCCET {
	excelCCETNoChartTitle = '\003\206\000\000',
	excelCCETChartTitleCenteredOverlay = '\003\206\000\001',
	excelCCETChartTitleAboveChart = '\003\206\000\002',
	excelCCETNoLegend = '\003\206\000d',
	excelCCETLegendRight = '\003\206\000e',
	excelCCETLegendTop = '\003\206\000f',
	excelCCETLegendLeft = '\003\206\000g',
	excelCCETLegendBottom = '\003\206\000h',
	excelCCETLegendRightOverlay = '\003\206\000i',
	excelCCETLegendLeftOverlay = '\003\206\000j',
	excelCCETNoDataLabel = '\003\206\000\310',
	excelCCETShowDataLabel = '\003\206\000\311',
	excelCCETDataLabelCenter = '\003\206\000\312',
	excelCCETDataLabelInsideEnd = '\003\206\000\313',
	excelCCETDataLabelInsideBase = '\003\206\000\314',
	excelCCETDataLabelOutsideEnd = '\003\206\000\315',
	excelCCETDataLabelLeft = '\003\206\000\316',
	excelCCETDataLabelRight = '\003\206\000\317',
	excelCCETDataLabelTop = '\003\206\000\320',
	excelCCETDataLabelBottom = '\003\206\000\321',
	excelCCETDataLabelBestFit = '\003\206\000\322',
	excelCCETNoPrimaryCategoryAxisTitle = '\003\206\001,',
	excelCCETPrimaryCategoryAxisTitleAdjacentToAxis = '\003\206\001-',
	excelCCETPrimaryCategoryAxisTitleBelowAxis = '\003\206\001.',
	excelCCETPrimaryCategoryAxisTitleRotated = '\003\206\001/',
	excelCCETPrimaryCategoryAxisTitleVertical = '\003\206\0010',
	excelCCETPrimaryCategoryAxisTitleHorizontal = '\003\206\0011',
	excelCCETNoPrimaryValueAxisTitle = '\003\206\0012',
	excelCCETPrimaryValueAxisTitleAdjacentToAxis = '\003\206\0013',
	excelCCETPrimaryValueAxisTitleBelowAxis = '\003\206\0014',
	excelCCETPrimaryValueAxisTitleRotated = '\003\206\0015',
	excelCCETPrimaryValueAxisTitleVertical = '\003\206\0016',
	excelCCETPrimaryValueAxisTitleHorizontal = '\003\206\0017',
	excelCCETNoSecondaryCategoryAxisTitle = '\003\206\0018',
	excelCCETSecondaryCategoryAxisTitleAdjacentToAxis = '\003\206\0019',
	excelCCETSecondaryCategoryAxisTitleBelowAxis = '\003\206\001:',
	excelCCETSecondaryCategoryAxisTitleRotated = '\003\206\001;',
	excelCCETSecondaryCategoryAxisTitleVertical = '\003\206\001<',
	excelCCETSecondaryCategoryAxisTitleHorizontal = '\003\206\001=',
	excelCCETNoSecondaryValueAxisTitle = '\003\206\001>',
	excelCCETSecondaryValueAxisTitleAdjacentToAxis = '\003\206\001\?',
	excelCCETSecondaryValueAxisTitleBelowAxis = '\003\206\001@',
	excelCCETSecondaryValueAxisTitleRotated = '\003\206\001A',
	excelCCETSecondaryValueAxisTitleVertical = '\003\206\001B',
	excelCCETSecondaryValueAxisTitleHorizontal = '\003\206\001C',
	excelCCETNoSeriesAxisTitle = '\003\206\001D',
	excelCCETSeriesAxisTitleRotated = '\003\206\001E',
	excelCCETSeriesAxisTitleVertical = '\003\206\001F',
	excelCCETSeriesAxisTitleHorizontal = '\003\206\001G',
	excelCCETNoPrimaryValueGridLines = '\003\206\001H',
	excelCCETPrimaryValueGridLinesMinor = '\003\206\001I',
	excelCCETPrimaryValueGridLinesMajor = '\003\206\001J',
	excelCCETPrimaryValueGridLinesMinorMajor = '\003\206\001K',
	excelCCETNoPrimaryCategoryGridLines = '\003\206\001L',
	excelCCETPrimaryCategoryGridLinesMinor = '\003\206\001M',
	excelCCETPrimaryCategoryGridLinesMajor = '\003\206\001N',
	excelCCETPrimaryCategoryGridLinesMinorMajor = '\003\206\001O',
	excelCCETNoSecondaryValueGridLines = '\003\206\001P',
	excelCCETSecondaryValueGridLinesMinor = '\003\206\001Q',
	excelCCETSecondaryValueGridLinesMajor = '\003\206\001R',
	excelCCETSecondaryValueGridLinesMinorMajor = '\003\206\001S',
	excelCCETNoSecondaryCategoryGridLines = '\003\206\001T',
	excelCCETSecondaryCategoryGridLinesMinor = '\003\206\001U',
	excelCCETSecondaryCategoryGridLinesMajor = '\003\206\001V',
	excelCCETSecondaryCategoryGridLinesMinorMajor = '\003\206\001W',
	excelCCETNoSeriesAxisGridLines = '\003\206\001X',
	excelCCETSeriesAxisGridLinesMinor = '\003\206\001Y',
	excelCCETSeriesAxisGridLinesMajor = '\003\206\001Z',
	excelCCETSeriesAxisGridLinesMinorMajor = '\003\206\001[',
	excelCCETNoPrimaryCategoryAxis = '\003\206\001\\',
	excelCCETPrimaryCategoryAxisShow = '\003\206\001]',
	excelCCETPrimaryCategoryAxisWithoutLabels = '\003\206\001^',
	excelCCETPrimaryCategoryAxisReverse = '\003\206\001_',
	excelCCETNoPrimaryValueAxis = '\003\206\001`',
	excelCCETShowPrimaryValueAxis = '\003\206\001a',
	excelCCETPrimaryValueAxisThousands = '\003\206\001b',
	excelCCETPrimaryValueAxisMillions = '\003\206\001c',
	excelCCETPrimaryValueAxisBillions = '\003\206\001d',
	excelCCETPrimaryValueAxisLogScale = '\003\206\001e',
	excelCCETNoSecondaryCategoryAxis = '\003\206\001f',
	excelCCETShowSecondaryCategoryAxis = '\003\206\001g',
	excelCCETSecondaryCategoryAxisWithoutLabels = '\003\206\001h',
	excelCCETSecondaryCategoryAxisReverse = '\003\206\001i',
	excelCCETNoSecondaryValueAxis = '\003\206\001j',
	excelCCETShowSecondaryValueAxis = '\003\206\001k',
	excelCCETSecondaryValueAxisThousands = '\003\206\001l',
	excelCCETSecondaryValueAxisMillions = '\003\206\001m',
	excelCCETSecondaryValueAxisBillions = '\003\206\001n',
	excelCCETSecondaryValueAxisLogScale = '\003\206\001o',
	excelCCETNoSeriesAxis = '\003\206\001p',
	excelCCETShowSeriesAxis = '\003\206\001q',
	excelCCETSeriesAxisWithoutLabeling = '\003\206\001r',
	excelCCETSeriesAxisReverse = '\003\206\001s',
	excelCCETPrimaryCategoryAxisThousands = '\003\206\001t',
	excelCCETPrimaryCategoryAxisMillions = '\003\206\001u',
	excelCCETPrimaryCategoryAxisBillions = '\003\206\001v',
	excelCCETPrimaryCategoryAxisLogScale = '\003\206\001w',
	excelCCETSecondaryCategoryAxisThousands = '\003\206\001x',
	excelCCETSecondaryCategoryAxisMillions = '\003\206\001y',
	excelCCETSecondaryCategoryAxisBillions = '\003\206\001z',
	excelCCETSecondaryCategoryAxisLogScale = '\003\206\001{',
	excelCCETNoDataTable = '\003\206\001\364',
	excelCCETShowDataTable = '\003\206\001\365',
	excelCCETDataTableWithLegendKeys = '\003\206\001\366',
	excelCCETNoTrendline = '\003\206\002X',
	excelCCETTrendlineAddLinear = '\003\206\002Y',
	excelCCETTrendlineAddExponential = '\003\206\002Z',
	excelCCETTrendlineAddLinearForecast = '\003\206\002[',
	excelCCETTrendlineAddTwoPeriodMovingAverage = '\003\206\002\\',
	excelCCETNoErrorBar = '\003\206\002\274',
	excelCCETErrorBarStandardError = '\003\206\002\275',
	excelCCETErrorBarPercentage = '\003\206\002\276',
	excelCCETErrorBarStandardDeviation = '\003\206\002\277',
	excelCCETNoLine = '\003\206\003 ',
	excelCCETLineDropLine = '\003\206\003!',
	excelCCETLineHiLoLine = '\003\206\003\"',
	excelCCETLineSeriesLine = '\003\206\003#',
	excelCCETLineDropHiloLine = '\003\206\003$',
	excelCCETNoUpDownBars = '\003\206\003\204',
	excelCCETShowUpDownBars = '\003\206\003\205',
	excelCCETNoPlotArea = '\003\206\003\350',
	excelCCETShowPlotArea = '\003\206\003\351',
	excelCCETNoChartWall = '\003\206\004L',
	excelCCETShowChartWall = '\003\206\004M',
	excelCCETNoChartFloor = '\003\206\004\260',
	excelCCETShowChartFloor = '\003\206\004\261'
};
typedef enum excelCCET excelCCET;

enum excelXDFC {
	excelXDFCFilterAboveAverage = '\003\207\000!' /* Filter all above-average values. */,
	excelXDFCFilterAllDatesInApril = '\003\207\000\030' /* Filter all dates in April. */,
	excelXDFCFilterAllDatesInAugust = '\003\207\000\034' /* Filter all dates in August. */,
	excelXDFCFilterAllDatesInDecember = '\003\207\000 ' /* Filter all dates in December. */,
	excelXDFCFilterAllDatesInFebruary = '\003\207\000\026' /* Filter all dates in February */,
	excelXDFCFilterAllDatesInJanuary = '\003\207\000\025' /* Filter all dates in January. */,
	excelXDFCFilterAllDatesInJuly = '\003\207\000\033' /* Filter all dates in July. */,
	excelXDFCFilterAllDatesInJune = '\003\207\000\032' /* Filter all dates in June. */,
	excelXDFCFilterAllDatesInMarch = '\003\207\000\027' /* Filter all dates in March. */,
	excelXDFCFilterAllDatesInMay = '\003\207\000\031' /* Filter all dates in May. */,
	excelXDFCFilterAllDatesInNovember = '\003\207\000\037' /* Filter all dates in November. */,
	excelXDFCFilterAllDatesInOctober = '\003\207\000\036' /* Filter all dates in October. */,
	excelXDFCFilterAllDatesInQuarter1 = '\003\207\000\021' /* Filter all dates in Quarter1. */,
	excelXDFCFilterAllDatesInQuarter2 = '\003\207\000\022' /* Filter all dates in Quarter2. */,
	excelXDFCFilterAllDatesInQuarter3 = '\003\207\000\023' /* Filter all dates in Quarter3. */,
	excelXDFCFilterAllDatesInQuarter4 = '\003\207\000\024' /* Filter all dates in Quarter4. */,
	excelXDFCFilterAllDatesInSeptember = '\003\207\000\035' /* Filter all dates in September. */,
	excelXDFCFilterBelowAverage = '\003\207\000\"' /* Filter all below-average values. */,
	excelXDFCFilterLastMonth = '\003\207\000\010' /* Filter all values related to last month. */,
	excelXDFCFilterLastQuarter = '\003\207\000\013' /* Filter all values related to last quarter. */,
	excelXDFCFilterLastWeek = '\003\207\000\005' /* Filter all values related to last week. */,
	excelXDFCFilterLastYear = '\003\207\000\016' /* Filter all values related to last year. */,
	excelXDFCFilterNextMonth = '\003\207\000\011' /* Filter all values related to next month. */,
	excelXDFCFilterNextQuarter = '\003\207\000\014' /* Filter all values related to next quarter. */,
	excelXDFCFilterNextWeek = '\003\207\000\006' /* Filter all values related to next week. */,
	excelXDFCFilterNextYear = '\003\207\000\017' /* Filter all values related to next year. */,
	excelXDFCFilterThisMonth = '\003\207\000\007' /* Filter all values related to the current month. */,
	excelXDFCFilterThisQuarter = '\003\207\000\012' /* Filter all values related to the current quarter. */,
	excelXDFCFilterThisWeek = '\003\207\000\004' /* Filter all values related to the current week. */,
	excelXDFCFilterThisYear = '\003\207\000\015' /* Filter all values related to the current year. */,
	excelXDFCFilterToday = '\003\207\000\001' /* Filter all values related to the current date. */,
	excelXDFCFilterTomorrow = '\003\207\000\003' /* Filter all values related to tomorrow. */,
	excelXDFCFilterYearToDate = '\003\207\000\020' /* Filter all values from today until a year ago. */,
	excelXDFCFilterYesterday = '\003\207\000\002' /* Filter all values related to yesterday. */
};
typedef enum excelXDFC excelXDFC;

enum excelE304 {
	excelE304ThemeFontIndexNone = '\002\300\000\000',
	excelE304ThemeFontIndexMajor = '\002\300\000\001',
	excelE304ThemeFontIndexMinor = '\002\300\000\002'
};
typedef enum excelE304 excelE304;

enum excelE305 {
	excelE305ColorIndexNone = '\002\300\357\322',
	excelE305FirstDarkThemeColor = '\002\301\000\001',
	excelE305FirstLightThemeColor = '\002\301\000\002',
	excelE305SecondDarkThemeColor = '\002\301\000\003',
	excelE305SecondLightThemeColor = '\002\301\000\004',
	excelE305FirstAccentThemeColor = '\002\301\000\005',
	excelE305SecondAccentThemeColor = '\002\301\000\006',
	excelE305ThirdAccentThemeColor = '\002\301\000\007',
	excelE305FourthAccentThemeColor = '\002\301\000\010',
	excelE305FifthAccentThemeColor = '\002\301\000\011',
	excelE305SixthAccentThemeColor = '\002\301\000\012',
	excelE305HyperlinkThemeColor = '\002\301\000\013',
	excelE305FollowedHyperlinkThemeColor = '\002\301\000\014'
};
typedef enum excelE305 excelE305;

enum excelCivt {
	excelCivtMinorVersion = '\002\304\000\000',
	excelCivtMajorVersion = '\002\304\000\001',
	excelCivtOverwriteCurrentVersion = '\002\304\000\002'
};
typedef enum excelCivt excelCivt;

enum excelExWt {
	excelExWtEntirePage = '\002\305\000\000',
	excelExWtAllTables = '\002\305\000\001',
	excelExWtSpecifiedTables = '\002\305\000\002'
};
typedef enum excelExWt excelExWt;

enum excelEbfT {
	excelEbfTWebFormattingAll = '\002\306\000\000',
	excelEbfTWebFormattingRtf = '\002\306\000\001',
	excelEbfTWebFormattingNone = '\002\306\000\002'
};
typedef enum excelEbfT excelEbfT;

enum excelXRbC {
	excelXRbCAsRequired = '\002\307\000\000',
	excelXRbCAlways = '\002\307\000\001',
	excelXRbCNever = '\002\307\000\002'
};
typedef enum excelXRbC excelXRbC;

enum excelE306 {
	excelE306ConditionValueNone = '\002\307\377\377',
	excelE306ConditionValueNumber = '\002\310\000\000',
	excelE306ConditionValueLowestValue = '\002\310\000\001',
	excelE306ConditionValueHighestValue = '\002\310\000\002',
	excelE306ConditionValuePercent = '\002\310\000\003',
	excelE306ConditionValueFormula = '\002\310\000\004',
	excelE306ConditionValuePercentile = '\002\310\000\005',
	excelE306ConditionValueAutomaticMinimum = '\002\310\000\006',
	excelE306ConditionValueAutomaticMaximum = '\002\310\000\007'
};
typedef enum excelE306 excelE306;

enum excelE307 {
	excelE307PivotConditionSelectionScope = '\002\311\000\000',
	excelE307PivotConditionFieldsScope = '\002\311\000\001',
	excelE307PivotConditionDataFieldScope = '\002\311\000\002'
};
typedef enum excelE307 excelE307;

enum excelE308 {
	excelE308DatabarFillSolid = '\002\312\000\000',
	excelE308DatabarFillGradient = '\002\312\000\001'
};
typedef enum excelE308 excelE308;

enum excelE309 {
	excelE309DatabarBorderNone = '\002\313\000\000',
	excelE309DatabarBorderSolid = '\002\313\000\001'
};
typedef enum excelE309 excelE309;

enum excelE310 {
	excelE310DatabarAxisAutomatic = '\002\314\000\000',
	excelE310DatabarAxisMidpoint = '\002\314\000\001',
	excelE310DatabarAxisNone = '\002\314\000\002'
};
typedef enum excelE310 excelE310;

enum excelE311 {
	excelE311DatabarAutomatic = '\002\315\000\000',
	excelE311DatabarPositiveFormat = '\002\315\000\001',
	excelE311DatabarCustomFormat = '\002\315\000\002'
};
typedef enum excelE311 excelE311;

enum excelE312 {
	excelE312FormatConditionIconNoCellIcon = '\002\315\377\377',
	excelE312FormatConditionIconGreenUpArrow = '\002\316\000\001',
	excelE312FormatConditionIconYellowSideArrow = '\002\316\000\002',
	excelE312FormatConditionIconRedDownArrow = '\002\316\000\003',
	excelE312FormatConditionIconGrayUpArrow = '\002\316\000\004',
	excelE312FormatConditionIconGraySideArrow = '\002\316\000\005',
	excelE312FormatConditionIconGrayDownArrow = '\002\316\000\006',
	excelE312FormatConditionIconGreenFlag = '\002\316\000\007',
	excelE312FormatConditionIconYellowFlag = '\002\316\000\010',
	excelE312FormatConditionIconRedFlag = '\002\316\000\011',
	excelE312FormatConditionIconGreenCircle = '\002\316\000\012',
	excelE312FormatConditionIconYellowCircle = '\002\316\000\013',
	excelE312FormatConditionIconRedCircleWithBorder = '\002\316\000\014',
	excelE312FormatConditionIconBlackCircleWithBorder = '\002\316\000\015',
	excelE312FormatConditionIconGreenTrafficLight = '\002\316\000\016',
	excelE312FormatConditionIconYellowTrafficLight = '\002\316\000\017',
	excelE312FormatConditionIconRedTrafficLight = '\002\316\000\020',
	excelE312FormatConditionIconYellowTriangle = '\002\316\000\021',
	excelE312FormatConditionIconRedDiamond = '\002\316\000\022',
	excelE312FormatConditionIconGreenCheckSymbol = '\002\316\000\023',
	excelE312FormatConditionIconYellowExclamationSymbol = '\002\316\000\024',
	excelE312FormatConditionIconRedCrossSymbol = '\002\316\000\025',
	excelE312FormatConditionIconGreenCheck = '\002\316\000\026',
	excelE312FormatConditionIconYellowExclamation = '\002\316\000\027',
	excelE312FormatConditionIconRedCross = '\002\316\000\030',
	excelE312FormatConditionIconYellowUpInclineArrow = '\002\316\000\031',
	excelE312FormatConditionIconYellowDownInclineArrow = '\002\316\000\032',
	excelE312FormatConditionIconGrayUpInclineArrow = '\002\316\000\033',
	excelE312FormatConditionIconGrayDownInclineArrow = '\002\316\000\034',
	excelE312FormatConditionIconRedCircle = '\002\316\000\035',
	excelE312FormatConditionIconPinkCircle = '\002\316\000\036',
	excelE312FormatConditionIconGrayCircle = '\002\316\000\037',
	excelE312FormatConditionIconBlackCircle = '\002\316\000 ',
	excelE312FormatConditionIconCircleWithOneWhiteQuarter = '\002\316\000!',
	excelE312FormatConditionIconCircleWithTwoWhiteQuarters = '\002\316\000\"',
	excelE312FormatConditionIconCircleWithThreeWhiteQuarters = '\002\316\000#',
	excelE312FormatConditionIconWhiteCircleAllWhiteQuarters = '\002\316\000$',
	excelE312FormatConditionIcon0Bars = '\002\316\000%',
	excelE312FormatConditionIcon1Bar = '\002\316\000&',
	excelE312FormatConditionIcon2Bars = '\002\316\000\'',
	excelE312FormatConditionIcon3Bars = '\002\316\000(',
	excelE312FormatConditionIcon4Bars = '\002\316\000)',
	excelE312FormatConditionIconGoldStar = '\002\316\000*',
	excelE312FormatConditionIconHalfGoldStar = '\002\316\000+',
	excelE312FormatConditionIconSilverStar = '\002\316\000,',
	excelE312FormatConditionIconGreenUpTriangle = '\002\316\000-',
	excelE312FormatConditionIconYellowDash = '\002\316\000.',
	excelE312FormatConditionIconRedDownTriangle = '\002\316\000/',
	excelE312FormatConditionIcon4FilledBoxes = '\002\316\0000',
	excelE312FormatConditionIcon3FilledBoxes = '\002\316\0001',
	excelE312FormatConditionIcon2FilledBoxes = '\002\316\0002',
	excelE312FormatConditionIcon1FilledBox = '\002\316\0003',
	excelE312FormatConditionIcon0FilledBoxes = '\002\316\0004'
};
typedef enum excelE312 excelE312;

enum excelE313 {
	excelE313IconSetCustom = '\002\316\377\377',
	excelE313IconSet3Arrows = '\002\317\000\001',
	excelE313IconSet3ArrowsGray = '\002\317\000\002',
	excelE313IconSet3Flags = '\002\317\000\003',
	excelE313IconSet3TrafficLights1 = '\002\317\000\004',
	excelE313IconSet3TrafficLights2 = '\002\317\000\005',
	excelE313IconSet3Signs = '\002\317\000\006',
	excelE313IconSet3Symbols = '\002\317\000\007',
	excelE313IconSet3Symbols2 = '\002\317\000\010',
	excelE313IconSet4Arrows = '\002\317\000\011',
	excelE313IconSet4ArrowsGray = '\002\317\000\012',
	excelE313IconSet4RedToBlack = '\002\317\000\013',
	excelE313IconSet4CRV = '\002\317\000\014',
	excelE313IconSet4TrafficLights = '\002\317\000\015',
	excelE313IconSet5Arrows = '\002\317\000\016',
	excelE313IconSet5ArrowsGray = '\002\317\000\017',
	excelE313IconSet5CRV = '\002\317\000\020',
	excelE313IconSet5Quarters = '\002\317\000\021',
	excelE313IconSet3Stars = '\002\317\000\022',
	excelE313IconSet3Triangles = '\002\317\000\023',
	excelE313IconSet5Boxes = '\002\317\000\024'
};
typedef enum excelE313 excelE313;

enum excelE314 {
	excelE314Top10Top = '\002\320\000\001',
	excelE314Top10Bottom = '\002\320\000\000'
};
typedef enum excelE314 excelE314;

enum excelE315 {
	excelE315CalcForAllValues = '\002\321\000\000',
	excelE315CalcForRowGroups = '\002\321\000\001',
	excelE315CalcForColGroups = '\002\321\000\002'
};
typedef enum excelE315 excelE315;

enum excelE316 {
	excelE316FormatAboveAverage = '\002\322\000\000',
	excelE316FormatBelowAverage = '\002\322\000\001',
	excelE316FormatEqualAboveAverage = '\002\322\000\002',
	excelE316FormatEqualBelowAverage = '\002\322\000\003',
	excelE316FormatAboveStandardDeviation = '\002\322\000\004',
	excelE316FormatBelowStandardDeviation = '\002\322\000\005'
};
typedef enum excelE316 excelE316;

enum excelE317 {
	excelE317FormatUniqueValues = '\002\323\000\000',
	excelE317FormatDuplicateValues = '\002\323\000\001'
};
typedef enum excelE317 excelE317;

enum excelE318 {
	excelE318TextContains = '\002\324\000\000',
	excelE318TextDoesNotContain = '\002\324\000\001',
	excelE318TextBeginsWith = '\002\324\000\002',
	excelE318TextEndsWith = '\002\324\000\003'
};
typedef enum excelE318 excelE318;

enum excelE319 {
	excelE319DateIsToday = '\002\325\000\000',
	excelE319DateIsYesterday = '\002\325\000\001',
	excelE319DateIsWithinTheLastSevenDays = '\002\325\000\002',
	excelE319DateIsThisWeek = '\002\325\000\003',
	excelE319DateIsLastWeek = '\002\325\000\004',
	excelE319DateIsLastMonth = '\002\325\000\005',
	excelE319DateIsTomorrow = '\002\325\000\006',
	excelE319DateIsNextWeek = '\002\325\000\007',
	excelE319DateIsNextMonth = '\002\325\000\010',
	excelE319DateIsThisMonth = '\002\325\000\011'
};
typedef enum excelE319 excelE319;

enum excelE320 {
	excelE320DatabarColorTypeColor = '\002\326\000\000',
	excelE320DatabarColorTypeSameAsPositive = '\002\326\000\001'
};
typedef enum excelE320 excelE320;

enum excel4022 {
	excel4022Window = 'cwin',
	excel4022Pane = 'X189'
};
typedef enum excel4022 excel4022;

enum excel4002 {
	excel4002Window = 'cwin',
	excel4002Sheet = 'X128',
	excel4002Workbook = 'X141'
};
typedef enum excel4002 excel4002;

enum excel4003 {
	excel4003Window = 'cwin',
	excel4003Sheet = 'X128',
	excel4003Workbook = 'X141'
};
typedef enum excel4003 excel4003;

enum excel4023 {
	excel4023Window = 'cwin',
	excel4023Pane = 'X189'
};
typedef enum excel4023 excel4023;

enum excel4004 {
	excel4004Application = 'capp',
	excel4004Sheet = 'X128'
};
typedef enum excel4004 excel4004;

enum excel4010 {
	excel4010Sheet = 'X128',
	excel4010Button = 'Xbtn',
	excel4010Checkbox = 'Xckb',
	excel4010OptionButton = 'XObn',
	excel4010Groupbox = 'XGBc',
	excel4010Label = 'Xlbl',
	excel4010Textbox = 'XTbx'
};
typedef enum excel4010 excel4010;

enum excel4013 {
	excel4013Button = 'Xbtn',
	excel4013Checkbox = 'Xckb',
	excel4013OptionButton = 'XObn',
	excel4013Scrollbar = 'XSrl',
	excel4013Listbox = 'XLbx',
	excel4013Groupbox = 'XGBc',
	excel4013Dropdown = 'XdpD',
	excel4013Spinner = 'XSpn',
	excel4013Label = 'Xlbl',
	excel4013Textbox = 'XTbx'
};
typedef enum excel4013 excel4013;

enum excel4014 {
	excel4014Button = 'Xbtn',
	excel4014Checkbox = 'Xckb',
	excel4014OptionButton = 'XObn',
	excel4014Scrollbar = 'XSrl',
	excel4014Listbox = 'XLbx',
	excel4014Groupbox = 'XGBc',
	excel4014Dropdown = 'XdpD',
	excel4014Spinner = 'XSpn',
	excel4014Label = 'Xlbl',
	excel4014Textbox = 'XTbx'
};
typedef enum excel4014 excel4014;

enum excel4024 {
	excel4024Dialog = 'X165',
	excel4024Scenario = 'X191'
};
typedef enum excel4024 excel4024;

enum excel4005 {
	excel4005Sheet = 'X128',
	excel4005Workbook = 'X141'
};
typedef enum excel4005 excel4005;

enum excel4011 {
	excel4011Button = 'Xbtn',
	excel4011Checkbox = 'Xckb',
	excel4011OptionButton = 'XObn',
	excel4011Scrollbar = 'XSrl',
	excel4011Listbox = 'XLbx',
	excel4011Groupbox = 'XGBc',
	excel4011Dropdown = 'XdpD',
	excel4011Spinner = 'XSpn',
	excel4011Label = 'Xlbl',
	excel4011Textbox = 'XTbx'
};
typedef enum excel4011 excel4011;

enum excel4015 {
	excel4015Button = 'Xbtn',
	excel4015Checkbox = 'Xckb',
	excel4015OptionButton = 'XObn',
	excel4015Scrollbar = 'XSrl',
	excel4015Listbox = 'XLbx',
	excel4015Groupbox = 'XGBc',
	excel4015Dropdown = 'XdpD',
	excel4015Spinner = 'XSpn',
	excel4015Label = 'Xlbl',
	excel4015Textbox = 'XTbx'
};
typedef enum excel4015 excel4015;

enum excel4027 {
	excel4027FormatCondition = 'X227',
	excel4027ColorScaleFormatCondition = 'X325',
	excel4027DatabarFormatCondition = 'X312',
	excel4027IconSetFormatCondition = 'X315',
	excel4027Top10FormatCondition = 'X321',
	excel4027AboveAverageFormatCondition = 'X322',
	excel4027UniqueValuesFormatCondition = 'X323'
};
typedef enum excel4027 excel4027;

enum excel4028 {
	excel4028FormatCondition = 'X227',
	excel4028ColorScaleFormatCondition = 'X325',
	excel4028DatabarFormatCondition = 'X312',
	excel4028IconSetFormatCondition = 'X315',
	excel4028Top10FormatCondition = 'X321',
	excel4028AboveAverageFormatCondition = 'X322',
	excel4028UniqueValuesFormatCondition = 'X323'
};
typedef enum excel4028 excel4028;

enum excel4029 {
	excel4029FormatCondition = 'X227',
	excel4029ColorScaleFormatCondition = 'X325',
	excel4029DatabarFormatCondition = 'X312',
	excel4029IconSetFormatCondition = 'X315',
	excel4029Top10FormatCondition = 'X321',
	excel4029AboveAverageFormatCondition = 'X322',
	excel4029UniqueValuesFormatCondition = 'X323'
};
typedef enum excel4029 excel4029;

enum excel4008 {
	excel4008PivotTable = 'X155',
	excel4008PivotField = 'X157'
};
typedef enum excel4008 excel4008;

enum excel4006 {
	excel4006CubeField = 'X900',
	excel4006CalculatedMember = 'X901',
	excel4006PivotFilter = 'X903',
	excel4006ValueChange = 'X905'
};
typedef enum excel4006 excel4006;

enum excel4009 {
	excel4009PivotField = 'X157',
	excel4009PivotItem = 'X160'
};
typedef enum excel4009 excel4009;

enum excel4007 {
	excel4007CubeField = 'X900',
	excel4007PivotField = 'X157'
};
typedef enum excel4007 excel4007;

enum excel4025 {
	excel4025PivotCache = 'X151',
	excel4025QueryTable = 'X231'
};
typedef enum excel4025 excel4025;

enum excel4016 {
	excel4016Listbox = 'XLbx',
	excel4016Dropdown = 'XdpD'
};
typedef enum excel4016 excel4016;

enum excel4026 {
	excel4026FormatCondition = 'X227',
	excel4026DisplayFormat = 'X306',
	excel4026Top10FormatCondition = 'X321',
	excel4026AboveAverageFormatCondition = 'X322',
	excel4026UniqueValuesFormatCondition = 'X323'
};
typedef enum excel4026 excel4026;

enum excel4021 {
	excel4021Listbox = 'XLbx',
	excel4021Dropdown = 'XdpD'
};
typedef enum excel4021 excel4021;

enum excel4018 {
	excel4018Listbox = 'XLbx',
	excel4018Dropdown = 'XdpD'
};
typedef enum excel4018 excel4018;

enum excel4012 {
	excel4012Button = 'Xbtn',
	excel4012Checkbox = 'Xckb',
	excel4012OptionButton = 'XObn',
	excel4012Scrollbar = 'XSrl',
	excel4012Listbox = 'XLbx',
	excel4012Groupbox = 'XGBc',
	excel4012Dropdown = 'XdpD',
	excel4012Spinner = 'XSpn',
	excel4012Label = 'Xlbl',
	excel4012Textbox = 'XTbx'
};
typedef enum excel4012 excel4012;

enum excel4017 {
	excel4017Listbox = 'XLbx',
	excel4017Dropdown = 'XdpD'
};
typedef enum excel4017 excel4017;

enum excel4019 {
	excel4019Listbox = 'XLbx',
	excel4019Dropdown = 'XdpD'
};
typedef enum excel4019 excel4019;

enum excel4020 {
	excel4020Listbox = 'XLbx',
	excel4020Dropdown = 'XdpD'
};
typedef enum excel4020 excel4020;

enum excel4001 {
	excel4001Window = 'cwin',
	excel4001Sheet = 'X128',
	excel4001Workbook = 'X141',
	excel4001Pane = 'X189'
};
typedef enum excel4001 excel4001;

enum excelVrOf {
	excelVrOfOverflow = '\002\302\000\000',
	excelVrOfClip = '\002\302\000\001',
	excelVrOfEllipsis = '\002\302\000\002'
};
typedef enum excelVrOf excelVrOf;

enum excelHzOf {
	excelHzOfOverflow = '\002\303\000\000',
	excelHzOfClip = '\002\303\000\001'
};
typedef enum excelHzOf excelHzOf;

enum excel4036 {
	excel4036CalloutFormat = 'X101',
	excel4036Callout = 'cD00'
};
typedef enum excel4036 excel4036;

enum excel4037 {
	excel4037CalloutFormat = 'X101',
	excel4037Callout = 'cD00'
};
typedef enum excel4037 excel4037;

enum excel4038 {
	excel4038CalloutFormat = 'X101',
	excel4038Callout = 'cD00'
};
typedef enum excel4038 excel4038;

enum excel4035 {
	excel4035Rectangle = 'XRct',
	excel4035Oval = 'XOvl',
	excel4035Arc = 'Xarc'
};
typedef enum excel4035 excel4035;

enum excel4032 {
	excel4032Line = 'Xlne',
	excel4032Rectangle = 'XRct',
	excel4032Oval = 'XOvl',
	excel4032Arc = 'Xarc',
	excel4032Shape = 'pShp'
};
typedef enum excel4032 excel4032;

enum excel4033 {
	excel4033Line = 'Xlne',
	excel4033Rectangle = 'XRct',
	excel4033Oval = 'XOvl',
	excel4033Arc = 'Xarc',
	excel4033Shape = 'pShp'
};
typedef enum excel4033 excel4033;

enum excel4030 {
	excel4030Line = 'Xlne',
	excel4030Rectangle = 'XRct',
	excel4030Oval = 'XOvl',
	excel4030Arc = 'Xarc'
};
typedef enum excel4030 excel4030;

enum excel4034 {
	excel4034Line = 'Xlne',
	excel4034Rectangle = 'XRct',
	excel4034Oval = 'XOvl',
	excel4034Arc = 'Xarc'
};
typedef enum excel4034 excel4034;

enum excel4031 {
	excel4031Line = 'Xlne',
	excel4031Rectangle = 'XRct',
	excel4031Oval = 'XOvl',
	excel4031Arc = 'Xarc',
	excel4031Shape = 'pShp'
};
typedef enum excel4031 excel4031;

enum excel4041 {
	excel4041ChartFillFormat = 'X253',
	excel4041ChartTitle = 'X256',
	excel4041AxisTitle = 'X257',
	excel4041SeriesPoint = 'X262',
	excel4041Series = 'X263',
	excel4041DataLabel = 'X265',
	excel4041LegendKey = 'X269',
	excel4041DownBars = 'X279',
	excel4041Floor = 'X280',
	excel4041Walls = 'X281',
	excel4041PlotArea = 'X283',
	excel4041ChartArea = 'X284',
	excel4041Legend = 'X285',
	excel4041DisplayUnitLabel = 'X299'
};
typedef enum excel4041 excel4041;

enum excel4053 {
	excel4053ChartArea = 'X284',
	excel4053Legend = 'X285'
};
typedef enum excel4053 excel4053;

enum excel4051 {
	excel4051SeriesPoint = 'X262',
	excel4051Series = 'X263',
	excel4051LegendKey = 'X269',
	excel4051Trendline = 'X271',
	excel4051Floor = 'X280',
	excel4051Walls = 'X281',
	excel4051PlotArea = 'X283',
	excel4051ChartArea = 'X284',
	excel4051ErrorBars = 'X286'
};
typedef enum excel4051 excel4051;

enum excel4049 {
	excel4049Chart = 'X119',
	excel4049SeriesPoint = 'X262',
	excel4049Series = 'X263'
};
typedef enum excel4049 excel4049;

enum excel4052 {
	excel4052Floor = 'X280',
	excel4052Walls = 'X281'
};
typedef enum excel4052 excel4052;

enum excel4042 {
	excel4042ChartFillFormat = 'X253',
	excel4042ChartTitle = 'X256',
	excel4042AxisTitle = 'X257',
	excel4042SeriesPoint = 'X262',
	excel4042Series = 'X263',
	excel4042DataLabel = 'X265',
	excel4042LegendKey = 'X269',
	excel4042DownBars = 'X279',
	excel4042Floor = 'X280',
	excel4042Walls = 'X281',
	excel4042PlotArea = 'X283',
	excel4042ChartArea = 'X284',
	excel4042Legend = 'X285',
	excel4042DisplayUnitLabel = 'X299'
};
typedef enum excel4042 excel4042;

enum excel4045 {
	excel4045ChartFillFormat = 'X253',
	excel4045ChartTitle = 'X256',
	excel4045AxisTitle = 'X257',
	excel4045SeriesPoint = 'X262',
	excel4045Series = 'X263',
	excel4045DataLabel = 'X265',
	excel4045LegendKey = 'X269',
	excel4045DownBars = 'X279',
	excel4045Floor = 'X280',
	excel4045Walls = 'X281',
	excel4045PlotArea = 'X283',
	excel4045ChartArea = 'X284',
	excel4045Legend = 'X285',
	excel4045DisplayUnitLabel = 'X299'
};
typedef enum excel4045 excel4045;

enum excel4046 {
	excel4046ChartFillFormat = 'X253',
	excel4046ChartTitle = 'X256',
	excel4046AxisTitle = 'X257',
	excel4046SeriesPoint = 'X262',
	excel4046Series = 'X263',
	excel4046DataLabel = 'X265',
	excel4046LegendKey = 'X269',
	excel4046DownBars = 'X279',
	excel4046Floor = 'X280',
	excel4046Walls = 'X281',
	excel4046PlotArea = 'X283',
	excel4046ChartArea = 'X284',
	excel4046Legend = 'X285',
	excel4046DisplayUnitLabel = 'X299'
};
typedef enum excel4046 excel4046;

enum excel4050 {
	excel4050ChartObject = 'X221',
	excel4050SeriesPoint = 'X262',
	excel4050Series = 'X263',
	excel4050ChartArea = 'X284'
};
typedef enum excel4050 excel4050;

enum excel4044 {
	excel4044ChartFillFormat = 'X253',
	excel4044ChartTitle = 'X256',
	excel4044AxisTitle = 'X257',
	excel4044SeriesPoint = 'X262',
	excel4044Series = 'X263',
	excel4044DataLabel = 'X265',
	excel4044LegendKey = 'X269',
	excel4044DownBars = 'X279',
	excel4044Floor = 'X280',
	excel4044Walls = 'X281',
	excel4044PlotArea = 'X283',
	excel4044ChartArea = 'X284',
	excel4044Legend = 'X285',
	excel4044DisplayUnitLabel = 'X299'
};
typedef enum excel4044 excel4044;

enum excel4040 {
	excel4040ChartObject = 'X221',
	excel4040Axis = 'X255'
};
typedef enum excel4040 excel4040;

enum excel4048 {
	excel4048ChartFillFormat = 'X253',
	excel4048ChartTitle = 'X256',
	excel4048AxisTitle = 'X257',
	excel4048SeriesPoint = 'X262',
	excel4048Series = 'X263',
	excel4048DataLabel = 'X265',
	excel4048LegendKey = 'X269',
	excel4048DownBars = 'X279',
	excel4048Floor = 'X280',
	excel4048Walls = 'X281',
	excel4048PlotArea = 'X283',
	excel4048ChartArea = 'X284',
	excel4048Legend = 'X285',
	excel4048DisplayUnitLabel = 'X299'
};
typedef enum excel4048 excel4048;

enum excel4047 {
	excel4047ChartFillFormat = 'X253',
	excel4047ChartTitle = 'X256',
	excel4047AxisTitle = 'X257',
	excel4047SeriesPoint = 'X262',
	excel4047Series = 'X263',
	excel4047DataLabel = 'X265',
	excel4047LegendKey = 'X269',
	excel4047DownBars = 'X279',
	excel4047Floor = 'X280',
	excel4047Walls = 'X281',
	excel4047PlotArea = 'X283',
	excel4047ChartArea = 'X284',
	excel4047Legend = 'X285',
	excel4047DisplayUnitLabel = 'X299'
};
typedef enum excel4047 excel4047;

enum excel4043 {
	excel4043ChartFillFormat = 'X253',
	excel4043ChartTitle = 'X256',
	excel4043AxisTitle = 'X257',
	excel4043SeriesPoint = 'X262',
	excel4043Series = 'X263',
	excel4043DataLabel = 'X265',
	excel4043LegendKey = 'X269',
	excel4043DownBars = 'X279',
	excel4043Floor = 'X280',
	excel4043Walls = 'X281',
	excel4043PlotArea = 'X283',
	excel4043ChartArea = 'X284',
	excel4043Legend = 'X285',
	excel4043DisplayUnitLabel = 'X299'
};
typedef enum excel4043 excel4043;

enum excel4039 {
	excel4039Chart = 'X119',
	excel4039ChartObject = 'X221'
};
typedef enum excel4039 excel4039;



/*
 * Standard Suite
 */

// A scriptable object
@interface excelBaseObject : SBObject

@property (copy) NSDictionary *properties;  // All of the object's properties

- (void) closeSaving:(excelSavo)saving savingIn:(excelXLfd)savingIn;  // Close an object
- (NSInteger) dataSizeAs:(NSNumber *)as;  // Return the size in bytes of an object
- (void) delete;  // Delete an element from an object
- (SBObject *) duplicateTo:(SBObject *)to;  // Duplicate object(s)
- (BOOL) exists;  // Verify if an object exists
- (SBObject *) moveTo:(SBObject *)to;  // Move object(s) to a new location
- (void) open;  // Open the specified object(s)
- (void) printWithProperties:(excelPrintSettings *)withProperties printDialog:(BOOL)printDialog;  // Print the specified object(s)
- (void) saveIn:(excelXLfd)in_ as:(excelE188)as;  // Save an object
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

// A basic application program
@interface excelBaseApplication : excelBaseObject

@property (readonly) BOOL frontmost;  // Is this the frontmost application?
@property (copy, readonly) NSString *name;  // the name
@property (copy, readonly) NSString *version;  // the version of the application


@end

// A basic document
@interface excelBaseDocument : excelBaseObject

@property (readonly) BOOL modified;  // Has the document been modified since the last save?
@property (copy, readonly) NSString *name;  // the name


@end

// A basic window
@interface excelBasicWindow : excelBaseObject

@property (copy) excelRectangle *bounds;  // the boundary rectangle for the window
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

@interface excelPrintSettings : SBObject

@property NSInteger copies;  // the number of copies of a document to be printed 
@property BOOL collating;  // Should printed copies be collated?
@property NSInteger startingPage;  // the first page of the document to be printed
@property NSInteger endingPage;  // the last page of the document to be printed
@property NSInteger pagesAcross;  // number of logical pages laid across a physical page
@property NSInteger pagesDown;  // number of logical pages laid out down a physical page
@property (copy) NSDate *requestedPrintTime;  // the time at which the desktop printer should print the document...
@property excelEnum errorHandling;  // how errors are handled
@property (copy) NSString *faxNumber;  // for fax number
@property (copy) NSString *targetPrinter;  // the queue name of the target printer

- (void) closeSaving:(excelSavo)saving savingIn:(excelXLfd)savingIn;  // Close an object
- (NSInteger) dataSizeAs:(NSNumber *)as;  // Return the size in bytes of an object
- (void) delete;  // Delete an element from an object
- (SBObject *) duplicateTo:(SBObject *)to;  // Duplicate object(s)
- (BOOL) exists;  // Verify if an object exists
- (SBObject *) moveTo:(SBObject *)to;  // Move object(s) to a new location
- (void) open;  // Open the specified object(s)
- (void) printWithProperties:(excelPrintSettings *)withProperties printDialog:(BOOL)printDialog;  // Print the specified object(s)
- (void) saveIn:(excelXLfd)in_ as:(excelE188)as;  // Save an object
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
 * Microsoft Office Suite
 */

// A control within a command bar.
@interface excelCommandBarControl : excelBaseObject

@property BOOL beginGroup;  // Returns or sets if the command bar control appears at the beginning of a group of controls on the command bar.
@property (readonly) BOOL builtIn;  // Returns true if the command bar control is a built-in command bar control.
@property (readonly) excelMCLT controlType;  // Returns the type of the command bar control.
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
@interface excelCommandBarButton : excelCommandBarControl

@property (readonly) BOOL buttonFaceIsDefault;  // Returns if the face of a command bar button control is the original built-in face.
@property excelMBns buttonState;  // Returns or set the appearance of a command bar button control.  The property is read-only for built-in command bar buttons.
@property excelMBTs buttonStyle;  // Returns or sets the way a command button control is displayed.
@property NSInteger faceId;  // Returns or sets the Id number for the face of the command bar button control.


@end

// A combobox menu control within a command bar.
@interface excelCommandBarCombobox : excelCommandBarControl

@property excelMXcb comboboxStyle;  // Returns or sets the way a command bar combobox control is displayed.
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
@interface excelCommandBarPopup : excelCommandBarControl

- (SBElementArray *) commandBarControls;


@end

// Toolbars used in all of the Office applications.
@interface excelCommandBar : excelBaseObject

- (SBElementArray *) commandBarControls;

@property excelMBPS barPosition;  // Returns or sets the position of the command bar.
@property (readonly) excelMBTY barType;  // Returns the type of this command bar.
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

@interface excelDocumentProperty : excelBaseObject

@property (copy) NSNumber *documentPropertyType;  // Returns or sets the document property type.
@property (copy) NSString *linkSource;  // Returns or sets the source of a lined custom document property.
@property BOOL linkToContent;  // True if the value of the document property is lined to the content of the container document.  False if the value is static.  This only applies to custom document properties.  For built-in properties this is always false.
@property (copy) NSString *name;  // Returns or sets the name of the document property.
@property (copy) NSString *value;  // Returns or sets the value of the document property.


@end

@interface excelCustomDocumentProperty : excelDocumentProperty


@end

@interface excelWebPageFont : excelBaseObject

@property (copy) NSString *fixedWidthFont;  // Returns or sets the fixed-width font setting.
@property double fixedWidthFontSize;  // Returns or sets the fixed-width font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.
@property (copy) NSString *proportionalFont;  // Returns or sets the proportional font setting.
@property double proportionalFontSize;  // Returns or sets the proportional font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.


@end



/*
 * Microsoft Excel Suite
 */

// Represents a cell comment.
@interface excelExcelComment : excelBaseObject

@property (copy, readonly) NSString *author;  // Returns the author of the comment.
@property (copy, readonly) excelShape *shapeObject;  // Returns a shape object that represents the shape attached to the specified comment.
@property BOOL visible;  // Returns or sets if the specified object is visible.

- (NSString *) ExcelCommentTextText:(NSString *)text start:(NSInteger)start overWrite:(BOOL)overWrite;  // Returns or sets the text of the comment
- (excelExcelComment *) nextExcelComment;  // Returns the next Excel comment object
- (excelExcelComment *) previousExcelComment;  // Returns the previous Excel comment object

@end

@interface excelODBCError : excelBaseObject

@property (copy, readonly) NSString *errorString;  // Returns the ODBC error string.
@property (copy, readonly) NSString *sqlState;  // Returns the SQL state error.


@end

// Represents the various types of protection options available for a worksheet
@interface excelProtection : excelBaseObject

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

@interface excelAboveAverageFormatCondition : excelBaseObject

@property excelE316 aboveOrBelow;
@property (copy, readonly) excelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property excelE315 calcFor;
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) excelE195 formatConditionType;  // Return the conditional format type.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property NSInteger numberOfStandardDeviations;
@property excelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

// Represents a single add-in, either installed or not installed.
@interface excelAddIn : excelBaseObject

@property (copy, readonly) NSString *fullName;  // Returns the name of the add in, including its path on disk, as a string.
@property BOOL installed;  // Returns or sets if the add in is installed.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy, readonly) NSString *path;  // Returns the complete path of the object, excluding the final separator and name of the add in.


@end

// A representation of the Microsoft Excel application.
@interface excelApplication : SBApplication

- (SBElementArray *) addIns;
- (SBElementArray *) chartSheets;
- (SBElementArray *) commandBars;
- (SBElementArray *) namedItems;
- (SBElementArray *) ranges;
- (SBElementArray *) cells;
- (SBElementArray *) rows;
- (SBElementArray *) columns;
- (SBElementArray *) windows;
- (SBElementArray *) workbooks;
- (SBElementArray *) sheets;
- (SBElementArray *) worksheets;
- (SBElementArray *) internationalMacroSheets;
- (SBElementArray *) macroSheets;
- (SBElementArray *) recentFiles;
- (SBElementArray *) ODBCErrors;

@property BOOL AutoFormatAsYouTypeReplaceHyperlinks;  // True if Microsoft Excel automatically formats hyperlinks as you type. False if Excel does not automatically format hyperlinks as you type.
@property excelE184 ExcelCursor;  // Returns or sets the appearance of the mouse pointer in Microsoft Excel.
@property NSInteger ODBCTimeout;  // Returns or sets the ODBC query time limit, in seconds. The default value is 45 seconds. 
@property (copy, readonly) excelCell *activeCell;  // Returns a range object that represents the active cell in the active window, the window on top, or in the specified window. If the window isn't displaying a worksheet, this property fails.
@property (copy, readonly) excelChart *activeChart;  // Returns a chart object that represents the active chart, either an embedded chart or a chart sheet. An embedded chart is considered active when it's either selected or activated.
@property (copy) NSString *activePrinter;  // Returns or sets the name of the active printer.
@property (copy, readonly) excelSheet *activeSheet;  // Returns an object that represents the active sheet, the sheet on top, in the active workbook or in the specified window or workbook.
@property (copy, readonly) excelWindow *activeWindow;  // Returns a window object that represents the active window, the window on top.
@property (copy, readonly) excelWorkbook *activeWorkbook;  // Returns a workbook object that represents the workbook in the active window, the window on top.
@property BOOL alertBeforeOverwriting;  // Returns or sets if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation.
@property (copy) NSString *altStartupPath;  // Returns or sets the name of the alternate startup folder.
@property (readonly) BOOL arbitraryXMLSupportAvailable;  // Returns a Boolean value that indicates whether the XML features in Microsoft Excel are available
@property BOOL askToUpdateLinks;  // Returns or sets if Microsoft Excel asks the user to update links when opening files with links.
@property (copy, readonly) excelAutocorrect *autocorrectObject;  // Returns an autocorrect object that represents the Microsoft Excel AutoCorrect attributes.
@property excelMASc automationSecurity;
@property (readonly) NSInteger build;  // Returns the Microsoft Excel build number.
@property BOOL calculateBeforeSave;  // Returns or sets if workbooks are calculated before they're saved to disk.
@property excelE174 calculation;  // Returns or sets the calculation mode.
@property (readonly) NSInteger calculationVersion;  // Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits, on the left, are the major version of Microsoft Excel.
@property (copy, readonly) NSString *caption;  // Returns the name of the application.
@property BOOL cellDragAndDrop;  // Returns or sets if dragging and dropping cells is enabled.
@property excelE172 commandUnderlines;  // Returns or sets the state of the command underlines in Microsoft Excel.
@property BOOL copyObjectsWithCells;  // Returns or sets if objects are cut, copied, extracted, and sorted with cells.
@property (readonly) NSInteger customListCount;  // Returns the number of defined custom lists, including built-in lists.
@property excelE161 cutCopyMode;  // Returns or sets the status of cut or copy mode.
@property excelE262 dataEntryMode;  // Returns or sets data entry mode, as shown in the following table. When in data entry mode, you can enter data only in the unlocked cells in the currently selected range. 
@property (copy) NSString *defaultFilePath;  // Returns or sets the default path that Microsoft Excel uses when it opens files. 
@property excelE188 defaultSaveFormat;  // Returns or sets the default format for saving files.
@property (copy, readonly) excelDefaultWebOptions *defaultWebOptionsObject;  // Returns the default web object.
@property BOOL displayAlerts;  // Returns or sets if Microsoft Excel displays certain alerts and messages while handling events from AppleScript.
@property excelE194 displayCommentIndicator;  // Returns or sets the way cells display comments and indicators.
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
@property excelE167 enableCancelKey;  // Controls how Microsoft Excel handles CTRL-BREAK, ESC, or cmd-period user interruptions to the running procedure.
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
@property excelE260 measurementUnit;  // Returns or set the current unit of measure.
@property (readonly) NSInteger memoryFree;  // Returns the amount of memory that's still available for Microsoft Excel to use, in bytes.
@property (readonly) NSInteger memoryTotal;  // Returns the total amount of memory, in bytes, that's available to Microsoft Excel, including memory already in use.
@property (readonly) NSInteger memoryUsed;  // Returns the amount of memory that Microsoft Excel is currently using, in bytes.
@property BOOL moveAfterReturn;  // Returns or sets if the active cell will be moved as soon as the ENTER/RETURN key is pressed.
@property excelE149 moveAfterReturnDirection;  // Returns or sets the direction in which the active cell is moved when the user presses ENTER.
@property (copy, readonly) NSString *name;  // Returns the name of the application.
@property (copy, readonly) NSString *networkTemplatesPath;  // Returns the network path where templates are stored. If the network path doesn't exist, this property returns an empty string. 
@property (copy, readonly) NSString *operatingSystem;  // Returns the name and version number of the current operating system.
@property (copy, readonly) NSString *organizationName;  // Returns the registered organization name.
@property (copy, readonly) NSString *path;  // Returns the complete path of the application, excluding the final separator and name of the application.
@property (copy, readonly) NSString *pathSeparator;  // Returns the path separator character.
@property BOOL pivotTableSelection;  // Returns or sets if pivot tables use structured selection.
@property BOOL promptForSummaryInfo;  // Returns or sets if Microsoft Excel asks for summary information when files are first saved.
@property excelE158 referenceStyle;  // Returns or sets how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style.
@property BOOL ribbonExpanded;  // Returns or sets a Boolean value that indicates whether the Ribbon in Microsoft Excel is expanded
@property BOOL rollZoom;  // Returns or sets if the IntelliMouse zooms instead of scrolling.
@property NSInteger saveInterval;
@property BOOL screenUpdating;  // Returns or sets if screen updating is turned on. Turn screen updating off to speed up your AppleScript code. You won't be able to see what the code is doing, but it will run faster.
@property (copy, readonly) excelRange *selection;  // Returns the selected object in the active window.
@property NSInteger sheetsInNewWorkbook;  // Returns or sets the number of sheets that Microsoft Excel automatically inserts into new workbooks.
@property BOOL showChartTipNames;  // Returns or sets if charts show chart tip names.
@property BOOL showChartTipValues;  // Returns or sets if charts show chart tip values.
@property BOOL showRibbon;  // Returns or sets a Boolean value that indicates whether the Ribbon in Microsoft Excel is shown
@property BOOL showToolTips;  // Returns or sets if tool tips are turned on.
@property (copy, readonly) excelXlspellingOptions *spellingOptions;  // Returns the default spelling options object.
@property (copy) NSString *standardFont;  // Returns or sets the name of the standard font.
@property double standardFontSize;  // Returns or sets the standard font size, in points.
@property BOOL startupDialog;  // Returns or sets if the startup dialog should be shown when starting up the application.
@property (copy, readonly) NSString *startupPath;  // Returns the complete path of the startup folder, excluding the final separator.
@property (copy) NSString *statusBar;  // Returns or sets the text in the status bar.
@property (copy, readonly) NSString *templatesPath;  // Returns the local path where templates are stored.
@property (copy, readonly) id thisCell;  // Returns the cell in which the user-defined function is being called from as a Range object.
@property (copy) NSString *transitionMenuKey;  // Returns or sets the alternate menu or help key.
@property excelE264 transitionMenuKeyAction;  // Returns or sets the action taken when the alternate menu key is pressed.
@property NSInteger twoDigitCutoffYear;  // Returns or sets the two digit value  after which year values are shown as four digits.
@property (readonly) double usableHeight;  // Returns the maximum height of the space that a window can occupy in points.
@property (readonly) double usableWidth;  // Returns the maximum width of the space that a window can occupy in points.
@property (copy) NSString *userName;  // Returns or sets the name of the current user.

- (void) quitSaving:(excelSavo)saving;  // Quit an application program
- (void) select;  // Make a selection
- (void) reset:(excel4000)x;  // Resets a built-in command bar or command bar control to its default configuration.
- (void) ExcelRepeat;  // Repeats the last user-interface action.
- (void) activateObject:(excel4001)x;  // Make the object active
- (void) addChartAutoformatChart:(excelChart *)chart name:(NSString *)name objectDescription:(NSString *)objectDescription;  // Adds a custom chart autoformat to the list of available chart autoformats.
- (void) addCustomListListArray:(excelE275)listArray byRow:(BOOL)byRow;  // Adds a custom list for custom autofill and/or custom sort.
- (void) addItemToList:(excel4016)x itemText:(NSString *)itemText entry_index:(NSInteger)entry_index;  // Adds a string to the list control
- (void) arrange_windowsArrangeStyle:(excelE183)arrangeStyle activeWorkbook:(BOOL)activeWorkbook syncHorizontal:(BOOL)syncHorizontal syncVertical:(BOOL)syncVertical;  // Arranges the windows on the screen
- (void) bringToFront:(excel4011)x;  // Bring the object to the front of the z-order of objects.
- (void) calculate:(excel4004)x;  // Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet..
- (void) calculateFull;  // Forces a full calculation of the data in all open workbooks.
- (void) calculateFullRebuild;  // For all open workbooks, forces a full calculation of the data and rebuilds the dependencies
- (double) centimetersToPointsCentimeters:(double)centimeters;  // Converts a measurement from centimeters to points, one point equals 0.035 centimeters.
- (void) checkSpelling:(excel4010)x customDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase alwaysSuggest:(BOOL)alwaysSuggest;  // Checks the spelling of an object.
- (BOOL) checkSpellingForTextToCheck:(NSString *)textToCheck customDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase;  // Checks the spelling of a single word. Returns True if the word is found in one of the dictionaries.
- (void) clearAllFilters:(excel4008)x;  // The ClearAllFilters method deletes all filters currently applied to the PivotTable.
- (void) clearManualFilter:(excel4007)x;  // The ClearManualFilter method provides an easy way to set the Visible property to True for all items of a PivotField in PivotTables, and to empty the HiddenItemsList or VisibleItemsList collections in OLAP PivotTables.
- (NSString *) convertFormulaFormulaToConvert:(NSString *)formulaToConvert fromReferenceStyle:(excelE158)fromReferenceStyle toReferenceStyle:(excelE158)toReferenceStyle toAbsolute:(excelE158)toAbsolute relativeTo:(excelE267)relativeTo;  // Converts cell references in a formula between the A1 and R1C1 reference styles, between relative and absolute references, or both
- (void) copyObject:(excel4012)x;  // Copies the object to the clipboard.
- (void) copyPicture:(excel4013)x appearance:(excelE207)appearance format:(excelE156)format;  // Copies the selected object to the clipboard as a picture.
- (void) cut:(excel4014)x;  // Cuts the object to the clipboard.
- (void) delete:(excel4006)x;  // Deletes the object.
- (void) deleteChartAutoformatName:(NSString *)name;  // Removes a custom chart autoformat from the list of available chart autoformats.
- (void) deleteCustomListListNum:(NSInteger)listNum;  // Deletes a custom list.
- (void) doubleClick;  // Equivalent to double-clicking the active cell.
- (void) drillTo:(excel4009)x field:(NSString *)field;  // The DrillTo method supports drilling to a specified PivotField from another PivotField.
- (id) evaluateName:(id)name;  // Converts a Microsoft Excel name to an object or value.
- (excelBorder *) getBorder:(excel4026)x whichBorder:(excelE243)whichBorder;  // Returns the specified border object.
- (NSArray *) getClipboardFormats;  // Returns a list of the  formats that are currently on the clipboard.
- (NSArray *) getCustomListContentsListNum:(NSInteger)listNum;  // Returns a custom list of strings.
- (NSInteger) getCustomListNumListArray:(NSArray *)listArray;  // Returns the custom list number for an array of strings. You can use this method to match both built-in lists and custom-defined lists.
- (excelDialog *) getDialog:(excelE245)x;  // Returns a the specified dialog object.
- (NSArray *) getFileConverters;  // Returns a list of all of the file converter objects.
- (id) getInternationalDataType:(excelE189)dataType;  // Returns information about a specific international setting.
- (NSString *) getListItem:(excel4017)x entry_index:(NSInteger)entry_index;  // Returns a string from the list
- (NSString *) getOpenFilenameFileFilter:(NSString *)fileFilter buttonText:(NSString *)buttonText multiSelect:(BOOL)multiSelect;  // Displays the standard Open dialog box and gets a file name from the user without actually opening any files.
- (NSArray *) getPreviousSelections;  // Returns a list of the last four ranges or names selected. Each element in the list is a range object.
- (NSArray *) getRegisteredFunctions;  // Returns information about functions in code resources that were registered with the REGISTER or REGISTER.ID macro functions.
- (NSString *) getSaveAsFilenameInitialFilename:(NSString *)initialFilename fileFilter:(NSString *)fileFilter filterIndex:(NSInteger)filterIndex buttonText:(NSString *)buttonText;  // Displays the standard save as dialog box and gets a file name from the user without actually saving any files.
- (excelWebPageFont *) getWebpageFont:(excelMChS)x;  // Returns the specified web page font object.
- (void) gotoReference:(excelE267)reference scroll:(BOOL)scroll;  // Selects any range in any workbook, and activates that workbook if it's not already active.
- (void) helpHelpFile:(NSString *)helpFile helpContextId:(NSInteger)helpContextId;  // Displays the Help window.
- (double) inchesToPointsInches:(double)inches;  // Converts a measurement from inches to points.
- (SBObject *) inputBoxPrompt:(NSString *)prompt title:(NSString *)title default:(excelE276)default_ leftPosition:(NSInteger)leftPosition top:(NSInteger)top type:(excelE277)type;  // Displays a dialog box for user input. Returns the information entered in the dialog box.
- (excelRange *) intersectRange1:(excelRange *)range1 range2:(excelRange *)range2 range3:(excelRange *)range3 range4:(excelRange *)range4 range5:(excelRange *)range5 range6:(excelRange *)range6 range7:(excelRange *)range7 range8:(excelRange *)range8 range9:(excelRange *)range9 range10:(excelRange *)range10 range11:(excelRange *)range11 range12:(excelRange *)range12 range13:(excelRange *)range13 range14:(excelRange *)range14 range15:(excelRange *)range15 range16:(excelRange *)range16 range17:(excelRange *)range17 range18:(excelRange *)range18 range19:(excelRange *)range19 range20:(excelRange *)range20 range21:(excelRange *)range21 range22:(excelRange *)range22 range23:(excelRange *)range23 range24:(excelRange *)range24 range25:(excelRange *)range25 range26:(excelRange *)range26 range27:(excelRange *)range27 range28:(excelRange *)range28 range29:(excelRange *)range29 range30:(excelRange *)range30;  // Returns a range object that represents the rectangular intersection of two or more ranges.
- (BOOL) itemSelected:(excel4021)x entry_index:(NSInteger)entry_index;  // Returns if a particular item in the list is selected
- (void) largeScroll:(excel4022)x down:(NSInteger)down up:(NSInteger)up toRight:(NSInteger)toRight toLeft:(NSInteger)toLeft;  // Scrolls the contents of the window by pages.
- (void) modifyAppliesToRange:(excel4029)x range:(excelRange *)range;  // Changes the range that this format condition applies to.
- (void) onKeyKey:(NSString *)key commandKeyPressed:(BOOL)commandKeyPressed shiftKeyPressed:(BOOL)shiftKeyPressed optionKeyPressed:(BOOL)optionKeyPressed controlKeyPressed:(BOOL)controlKeyPressed procedure:(excelE281)procedure;  // Runs a specified procedure when a particular key or key combination is pressed.
- (void) openFileMakerFileFilename:(NSString *)filename;  // Open a FileMaker file as an Excel worksheet.
- (void) openTextFileFilename:(NSString *)filename origin:(excelE211)origin startRow:(NSInteger)startRow dataType:(excelE236)dataType textQualifier:(excelE237)textQualifier consecutiveDelimiter:(BOOL)consecutiveDelimiter tab:(BOOL)tab semicolon:(BOOL)semicolon comma:(BOOL)comma space:(BOOL)space useOther:(BOOL)useOther otherChar:(NSString *)otherChar fieldInfo:(NSArray *)fieldInfo decimalSeparator:(NSString *)decimalSeparator thousandsSeparator:(NSString *)thousandsSeparator;  // Loads and parses a text file as a new workbook with a single sheet that contains the parsed text-file data.
- (excelWorkbook *) openWorkbookWorkbookFileName:(NSString *)workbookFileName updateLinks:(excelE294)updateLinks readOnly:(BOOL)readOnly format:(excelE295)format password:(NSString *)password writeReservedPassword:(NSString *)writeReservedPassword ignoreReadOnlyRecommended:(BOOL)ignoreReadOnlyRecommended origin:(excelE211)origin delimiter:(NSString *)delimiter editable:(BOOL)editable notify:(BOOL)notify converter:(NSInteger)converter addToMru:(BOOL)addToMru;  // Opens a workbook.
- (void) printOut:(excel4002)x from:(NSInteger)from to:(NSInteger)to copies:(NSInteger)copies preview:(BOOL)preview activePrinter:(NSString *)activePrinter printToFile:(BOOL)printToFile collate:(BOOL)collate;  // Prints the object
- (void) printPreview:(excel4003)x enableChanges:(BOOL)enableChanges;  // Shows a preview of the object as it would look when printed. This function has been deprecated.
- (BOOL) registerXllFilename:(NSString *)filename;  // Loads an XLL code resource and automatically registers the functions and commands contained in the resource.
- (void) removeAllItems:(excel4019)x;  // Removes all of the strings from the list
- (void) removeItem:(excel4020)x entry_index:(NSInteger)entry_index count:(NSInteger)count;  // Removes a specified string from the list
- (void) resetTimer:(excel4025)x;  // Resets the refresh timer for the specified PivotTable report to the last interval you set using the RefreshPeriod property.
- (NSString *) runVBMacro:(excelE297)x arg1:(id)arg1 arg2:(id)arg2 arg3:(id)arg3 arg4:(id)arg4 arg5:(id)arg5 arg6:(id)arg6 arg7:(id)arg7 arg8:(id)arg8 arg9:(id)arg9 arg10:(id)arg10 arg11:(id)arg11 arg12:(id)arg12 arg13:(id)arg13 arg14:(id)arg14 arg15:(id)arg15 arg16:(id)arg16 arg17:(id)arg17 arg18:(id)arg18 arg19:(id)arg19 arg20:(id)arg20 arg21:(id)arg21 arg22:(id)arg22 arg23:(id)arg23 arg24:(id)arg24 arg25:(id)arg25 arg26:(id)arg26 arg27:(id)arg27 arg28:(id)arg28 arg29:(id)arg29 arg30:(id)arg30;  // Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.
- (NSString *) runXLMMacro:(excelE297)x arg1:(id)arg1 arg2:(id)arg2 arg3:(id)arg3 arg4:(id)arg4 arg5:(id)arg5 arg6:(id)arg6 arg7:(id)arg7 arg8:(id)arg8 arg9:(id)arg9 arg10:(id)arg10 arg11:(id)arg11 arg12:(id)arg12 arg13:(id)arg13 arg14:(id)arg14 arg15:(id)arg15 arg16:(id)arg16 arg17:(id)arg17 arg18:(id)arg18 arg19:(id)arg19 arg20:(id)arg20 arg21:(id)arg21 arg22:(id)arg22 arg23:(id)arg23 arg24:(id)arg24 arg25:(id)arg25 arg26:(id)arg26 arg27:(id)arg27 arg28:(id)arg28 arg29:(id)arg29 arg30:(id)arg30;  // Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.
- (void) saveWorkspaceWorkspaceFileName:(NSString *)workspaceFileName;  // Saves the current workspace.
- (void) sendToBack:(excel4015)x;  // Sends the object to the back of the z-order.
- (void) setFirstPriority:(excel4027)x;  // Sets this condition to the highest priority.
- (void) setLastPriority:(excel4028)x;  // Sets this condition to the lowest priority.
- (void) setListItem:(excel4018)x entry_index:(NSInteger)entry_index itemText:(NSString *)itemText;  // Set a string in the list
- (void) show:(excel4024)x;  // Displays the built-in dialog box and waits for the user to input data.
- (void) smallScroll:(excel4023)x down:(NSInteger)down up:(NSInteger)up toRight:(NSInteger)toRight toLeft:(NSInteger)toLeft;  // Scrolls the contents of the window by rows or columns.
- (void) undo;  // Cancels the last user-interface action.
- (excelRange *) unionRange1:(excelRange *)range1 range2:(excelRange *)range2 range3:(excelRange *)range3 range4:(excelRange *)range4 range5:(excelRange *)range5 range6:(excelRange *)range6 range7:(excelRange *)range7 range8:(excelRange *)range8 range9:(excelRange *)range9 range10:(excelRange *)range10 range11:(excelRange *)range11 range12:(excelRange *)range12 range13:(excelRange *)range13 range14:(excelRange *)range14 range15:(excelRange *)range15 range16:(excelRange *)range16 range17:(excelRange *)range17 range18:(excelRange *)range18 range19:(excelRange *)range19 range20:(excelRange *)range20 range21:(excelRange *)range21 range22:(excelRange *)range22 range23:(excelRange *)range23 range24:(excelRange *)range24 range25:(excelRange *)range25 range26:(excelRange *)range26 range27:(excelRange *)range27 range28:(excelRange *)range28 range29:(excelRange *)range29 range30:(excelRange *)range30;  // Returns the union of two or more ranges.
- (void) unprotect:(excel4005)x password:(NSString *)password;  // Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.
- (void) bringToFront:(excel4030)x;  // Bring the object to the front of the z-order of objects.
- (void) checkSpelling:(excel4035)x customDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase alwaysSuggest:(BOOL)alwaysSuggest;  // Checks the spelling of an object.
- (void) copyObject:(excel4031)x;  // Copies the object to the clipboard.
- (void) copyPicture:(excel4032)x appearance:(excelE207)appearance format:(excelE156)format;  // Copies the selected object to the clipboard as a picture.
- (void) customDrop:(excel4036)x drop:(double)drop;  // Sets the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.
- (void) customLength:(excel4037)x length:(double)length;  // Specifies that the first segment of the callout line, the segment attached to the text callout box, retain a fixed length whenever the callout is moved.
- (void) cut:(excel4033)x;  // Cuts the object to the clipboard.
- (void) presetDrop:(excel4038)x dropType:(excelMCDt)dropType;  // Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.
- (void) sendToBack:(excel4034)x;  // Sends the object to the back of the z-order.
- (void) clearHyperlinks;  // clear all the hyperlinks in a range
- (void) dirty;  // Designates a range to be recalculated when the next recalculation occurs.
- (void) activateObject:(excel4039)x;  // Activates the object.
- (void) applyDataLabels:(excel4049)x type:(excelE134)type legendKey:(BOOL)legendKey autoText:(BOOL)autoText hasLeaderLines:(BOOL)hasLeaderLines showSeriesName:(BOOL)showSeriesName showCategoryName:(BOOL)showCategoryName showValue:(BOOL)showValue showPercentage:(BOOL)showPercentage showBubbleSize:(BOOL)showBubbleSize separator:(NSString *)separator;  // Applies data labels to all the series in a chart, a single series or a series point.
- (void) chartOneColorGradient:(excel4042)x gradientStyle:(excelMGdS)gradientStyle variant:(NSInteger)variant degree:(double)degree;  // Sets the specified fill to a one-color gradient.
- (void) chartPatterned:(excel4041)x pattern:(excelPpTy)pattern;  // Sets the specified fill to a pattern.
- (void) chartSolid:(excel4046)x;  // Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.
- (void) chartTwoColorGradient:(excel4045)x gradientStyle:(excelMGdS)gradientStyle variant:(NSInteger)variant;  // Sets the specified fill to a two-color gradient.
- (void) chartUserPicture:(excel4044)x pictureFile:(NSString *)pictureFile pictureFormat:(excelE124)pictureFormat pictureStackUnit:(NSInteger)pictureStackUnit picturePlacement:(excelE125)picturePlacement;  // Fills the specified shape with an image.
- (void) chartUserTextured:(excel4043)x textureFile:(NSString *)textureFile;  // Fills the specified shape with small tiles of an image.
- (void) clear:(excel4053)x;  // Clear the object.
- (void) clearFormats:(excel4051)x;  // Clears the formatting of the object.
- (void) copyObject:(excel4050)x;  // Copies the object to the clipboard.
- (void) deleteObject:(excel4040)x;  // Deletes the object.
- (void) paste:(excel4052)x;
- (void) presetChartGradient:(excel4048)x gradientStyle:(excelMGdS)gradientStyle variant:(NSInteger)variant presetGradientType:(excelMPGb)presetGradientType;  // Sets the specified fill to a preset gradient.
- (void) presetChartTextured:(excel4047)x presetTextureForChart:(excelMPzT)presetTextureForChart;  // Sets the specified fill format to a preset texture.

@end

// Represents autofiltering for the specified worksheet.
@interface excelAutofilter : excelBaseObject

- (SBElementArray *) filters;

@property (readonly) BOOL autofiltermode;  // Returns True if filters have been defined else False.
@property (copy, readonly) excelRange *rangeObject;  // Returns a range object that represents the range to which the specified autofilter applies.
@property (copy, readonly) excelSort *sortObject;  // Returns the sort object in the auto filter object.


@end

// Represents the border of an object.
@interface excelBorder : excelBaseObject

@property (copy) NSColor *color;  // Returns or sets the primary color of the object. 
@property excelE103 colorIndex;  // Returns or sets the color of the border. The color is specified as an index value into the current color palette.
@property excelE133 lineStyle;  // Returns or sets the line style for the border.
@property NSInteger lineWeight;  // Returns or sets the line weight of the border.
@property excelE305 themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.
@property excelE128 weight;  // Returns or sets the weight of the border.


@end

// Represents a button control.
@interface excelButton : excelBaseObject

- (SBElementArray *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property BOOL autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) excelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property BOOL cancelButton;  // Returns or sets if this button is the cancel button.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL defaultButton;  // Returns or sets if this button is the default button.
@property BOOL dismissButton;  // Returns or sets if this button is the dismiss button.
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property double height;  // Returns or set the height of the object.
@property BOOL helpButton;  // Returns or sets if this button is the help button.
@property excelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property excelE126 orientation;  // May also be a number value from -90 to 90 degrees.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property excelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) excelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property excelE284 verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property BOOL wrapAutoText;  // Returns or sets if the auto text is wrapped.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents the calculated fields and calculated items for PivotTables with Online Analytical Processing data sources.
@interface excelCalculatedMember : excelBaseObject

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
@property (readonly) excelE916 type;


@end

@interface excelCheckbox : excelBaseObject

- (SBElementArray *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object. 
@property (copy, readonly) excelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL displayThreeDShading;
@property BOOL enabled;  // Returns or sets if the control is enabled
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or sets the height of the control.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the control
@property (copy) NSString *linkedCell;
@property BOOL locked;  // Returns or sets whether the control is locked for editing.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of this control
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property excelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) excelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property excelE285 value;
@property BOOL visible;  // Returns or sets the current value of the control
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

@interface excelColorScaleCriteria : excelBaseObject


@end

@interface excelColorScaleCriterion : excelBaseObject

@property (readonly) NSInteger colorScaleCriterionIndex;  // The index of the color scale criterion. Read only.
@property excelE306 colorScaleCriterionType;  // Returns or sets the condition value type of the color scale criterion. Read/Write.
@property (copy) id colorScaleCriterionValue;
@property (copy, readonly) excelFormatColor *formatColor;  // Returns the FormatColor for the ColorScaleCriterion object.


@end

@interface excelColorScaleFormatCondition : excelBaseObject

@property (copy, readonly) excelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property (copy, readonly) excelColorScaleCriteria *colorScaleCriteria;  // Returns the ColorScaleCriteria for the ColorScale object.
@property NSInteger colorScaleType;
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) excelE195 formatConditionType;  // Return the conditional format type.
@property (copy) NSString *formula;  // Returns or sets the value or expression associated with the conditional format or data validation.
@property excelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

// Represents the color stop point for a gradient fill in an range or selection.
@interface excelColorstop : excelBaseObject

@property (copy) NSColor *color;  // Returns or sets the primary color of the object.
@property double colorstopPosition;  // Returns or sets the position of the ColorStop.
@property excelE305 themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.

- (void) deleteColorstop;  // Clears the represented object.

@end

// A collection of all the ColorStop objects for the specified series.
@interface excelColorstops : excelBaseObject

- (excelColorstop *) addColorstopPosition:(double)position;  // Adds ColorStop object to the ColorStops object.

@end

@interface excelConditionValue : excelBaseObject

@property (readonly) excelE306 conditionValueType;
@property (copy, readonly) id conditionValueValue;

- (void) modifyConditionValueType:(excelE306)type conditionValue:(id)conditionValue;  // Modifies an existing condition value.

@end

// Represents a hierarchy or measure field from an OLAP cube
@interface excelCubeField : excelBaseObject

- (SBElementArray *) pivotFields;

@property (readonly) BOOL allItemsVisible;  // The AllItemsVisible property checks whether manual filtering is applied to a PivotField or CubeField.
@property (copy) NSString *caption;  // The label text for the cube field.
@property (readonly) excelE914 cubeFieldSubType;  // Specifies the type of a CubeField.
@property (readonly) excelE913 cubeFieldType;  // Indicates whether the OLAP cube field is a hierarchy field or a measure field.
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
@property (readonly) excelE910 layoutForm;  // Returns or sets the way the specified PivotTable items appear -- in table format or in outline format.
@property excelE902 layoutSubtotalLocation;  // Returns or sets the position of the PivotTable field subtotals in relation to, either above or below, the specified field.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property excelE208 orientation;  // The location of the field in the specified PivotTable report.
@property NSInteger position;  // Position of the item in its field if the item is currently showing.
@property BOOL showInFieldList;  // When set to True, a CubeField object will be shown in the field list.
@property (copy, readonly) excelTreeviewControl *treeviewControl;
@property (copy, readonly) NSString *value;  // Returns the name of the specified field.

- (void) addMemberPropertyFieldProperty:(NSString *)property propertyOrder:(NSString *)propertyOrder propertyDisplayedIn:(excelE915)propertyDisplayedIn;  // Adds a member property field to the display for the cube field.
- (void) createPivotFields;  // The CreatePivotFields method enables users to create all PivotFields of a CubeField.

@end

// Represents a custom workbook view.
@interface excelCustomView : excelBaseObject

@property (readonly) BOOL customViewPrintSettings;  // True if print settings are included in the custom view.
@property (copy, readonly) NSString *name;  // Returns the name of this object.
@property (readonly) BOOL rowColSettings;  // Returns true if the custom view includes settings for hidden rows and columns, including filter information.

- (void) showCustomView;  // Displays the custom view

@end

@interface excelDatabarBorder : excelBaseObject

@property (copy, readonly) excelFormatColor *databarBorderColor;
@property excelE309 databarBorderType;  // Returns or sets the type of border of the databar


@end

@interface excelDatabarFormatCondition : excelBaseObject

@property (copy, readonly) excelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property (copy, readonly) excelFormatColor *databarAxisColor;
@property excelE310 databarAxisPosition;  // Returns or sets the position of the databar axis
@property (copy, readonly) excelFormatColor *databarBarColor;
@property (copy, readonly) excelDatabarBorder *databarBorder;  // Returns the DataBarBorder for the Databar object.
@property NSInteger databarDirection;  // Specifies the direction of the databar. Read/Write.
@property excelE308 databarFillType;  // Returns or sets the type of fill of the databar
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property BOOL formatConditionShowValue;  // Specifies whether to show the cell value along with the databar. Read/Write.
@property (readonly) excelE195 formatConditionType;  // Return the conditional format type.
@property (copy) NSString *formula;  // Returns or sets the value or expression associated with the conditional format or data validation.
@property (copy, readonly) excelConditionValue *maxPointConditionValue;  // Returns the ConditionValue for the maximum point of the DataBar object.
@property NSInteger maxPointPercent;  // Specifies the percentage of the data bar to draw at the maximum point. Read/Write.
@property (copy, readonly) excelConditionValue *minPointConditionValue;  // Returns the ConditionValue for the minimum point of the DataBar object.
@property NSInteger minPointPercent;  // Specifies the percentage of the data bar to draw at the minimum point. Read/Write.
@property (copy, readonly) excelNegativeBarFormat *negativeBarFormat;  // Returns the NegativeBarFormat for the Databar object.
@property excelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

// Contains global application-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.
@interface excelDefaultWebOptions : excelBaseObject

@property BOOL allowPng;  // Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save documents as a Web page.
@property BOOL alwaysSaveInDefaultEncoding;  // Returns or sets if the default encoding is used when you save a Web page or plain text document, independent of the file's original encoding when opened.
@property excelMtEn encoding;  // Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document.
@property (copy) NSString *locationOfComponents;  // Returns or sets the central URL, on the intranet or Web, or path, local or network, to the location from which authorized users can download Microsoft Office Web components when viewing your saved document.
@property NSInteger pixelsPerInch;  // Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120.
@property excelMSsz screenSize;  // Returns or sets the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser.
@property BOOL updateLinksOnSave;  // Returns or sets if hyperlinks and paths to all supporting files are automatically updated before you save the document as a Web page, ensuring that the links are up-to-date at the time the document is saved. If false, the links are not updated.


@end

// Represents a built-in Microsoft Excel dialog box.
@interface excelDialog : excelBaseObject


@end

// Represents the formatting shown to the user.
@interface excelDisplayFormat : excelBaseObject

@property (readonly) BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (readonly) BOOL formulaHidden;  // Returns or sets if the formula will be hidden when the workbook or worksheet is protected.
@property (readonly) excelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the range.
@property (readonly) NSInteger indentLevel;  // Returns or sets the indent level for the range or style. Can be an integer from 0 to 15.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (readonly) BOOL locked;  // Returns or sets  if the range is locked.
@property (readonly) BOOL mergeCells;  // Returns or sets if the range contains merged cells. 
@property (copy, readonly) NSString *numberFormat;  // Returns or sets the format code for the object.
@property (readonly) excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property (readonly) BOOL shrinkToFit;  // Returns or sets if text automatically shrinks to fit in the available column width.
@property (copy, readonly) excelStyle *styleObject;  // Returns or sets a style object that represents the style of the specified range.
@property (readonly) excelE126 textOrientation;  // The text orientation. Can be a number value from -90 to 90 degrees.
@property (readonly) excelE284 verticalAlignment;  // Returns or sets the vertical alignment of the range.
@property (readonly) BOOL wrapText;  // Returns or sets if Microsoft Excel wraps the text in the object.


@end

@interface excelDropdown : excelBaseObject

- (SBElementArray *) characters;

@property (copy, readonly) excelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying. 
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
@property excelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) excelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property NSInteger value;  // Returns or sets the current value of the control
@property BOOL visible;  // Returns or sets if the control is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// Represents a filter for a single column. 
@interface excelFilter : excelBaseObject

@property (copy, readonly) id criteria1;  // Returns the first filtered value for the specified column in a filtered range.
@property (copy, readonly) id criteria2;  // Returns the second filtered value for the specified column in a filtered range.
@property (readonly) BOOL filterOn;  // Returns true if the specified filter is on.
- (excelE186) operator;  // set or return the operator that associates the two criteria applied by the specified filter.
- (void) setOperator: (excelE186) operator;


@end

// Represents a color that a format condition can be colored with.
@interface excelFormatColor : excelBaseObject

@property (copy) NSColor *color;  // Returns or sets the primary color of the object.
@property excelE103 colorIndex;
@property excelE305 themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.


@end

@interface excelFormatConditionIconObject : excelBaseObject

@property (readonly) NSInteger formatConditionIconIndex;  // The index of the icon. Read only.


@end

@interface excelFormatConditionIconSet : excelBaseObject

@property (readonly) excelE313 iconSetId;


@end

@interface excelFormatConditionIconSets : excelBaseObject


@end

// Represents a conditional format.
@interface excelFormatCondition : excelBaseObject

@property (copy, readonly) excelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property (readonly) excelE196 conditionOperator;  // Returns the operator for the conditional format.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property excelE319 formatConditionDateOperator;
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (copy) NSString *formatConditionText;  // Returns or sets the text for the specified object.
@property excelE318 formatConditionTextOperator;
@property (readonly) excelE195 formatConditionType;  // Return the conditional format type.
@property (copy, readonly) NSString *formula1;  // Returns the value or expression associated with the conditional format or data validation.
@property (copy, readonly) NSString *formula2;  // Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is operator between or operator not between.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property excelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read/Write.

- (void) modifyConditionType:(excelE195)type operator:(excelE196)operator_ formula1:(NSString *)formula1 formula2:(NSString *)formula2 string:(NSString *)string operator2:(excelE196)operator2;  // Modifies an existing conditional format.

@end

// Contains properties that apply to header and footer picture objects.
@interface excelGraphic : excelBaseObject

@property double brightness;  // Returns or sets the brightness of the specified picture. The value for this property must be a number from 0.0, dimmest, to 1.0, brightest.
@property excelMPc colorType;  // Returns or sets the type of color transformation applied to the specified picture.
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

@interface excelGroupbox : excelBaseObject

- (SBElementArray *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property (copy, readonly) excelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
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
@property excelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) excelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property BOOL visible;  // Returns or sets if the control is visible
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// Represents a horizontal page break.
@interface excelHorizontalPageBreak : excelBaseObject

@property (readonly) excelE190 extent;  // Returns the type of the specified page break: full-screen or only within a print area.
@property excelE168 horizontalPageBreakType;  // Returns or sets the type of horizontal page break.
@property (copy) excelRange *location;  // Returns or sets where this horizontal page break is located.


@end

// Represents a hyperlink.
@interface excelHyperlink : excelBaseObject

@property (copy) NSString *address;  // Returns or sets the address of the target document.
@property (copy) NSString *emailSubject;  // Returns or sets the text string of the specified hyperlink's e-mail subject line. The subject line is appended to the hyperlink's address.
@property (readonly) excelMHlT hyperlinkType;  // Returns the hyperlink type, what the hyperlink is associated with.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy, readonly) excelRange *rangeObject;  // Returns a range object that represents the range the specified hyperlink is attached to.
@property (copy) NSString *screenTip;  // Returns or sets the screen tip text for the specified hyperlink.
@property (copy, readonly) excelShape *shapeObject;  // Returns a shape object that represents the shape attached to the specified hyperlink.
@property (copy) NSString *subAddress;  // Returns or sets the location within the document associated with the hyperlink.
@property (copy) NSString *textToDisplay;  // Returns or sets the text to be displayed for the specified hyperlink. The default value is the address of the hyperlink.

- (void) createNewDocumentFileName:(NSString *)fileName editNow:(BOOL)editNow overwrite:(BOOL)overwrite;  // Creates a new document linked to the specified hyperlink.
- (void) followNewWindow:(BOOL)newWindow extraInfo:(NSString *)extraInfo method:(excelMXiM)method headerInfo:(NSString *)headerInfo;  // Displays a cached document, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.

@end

@interface excelIconCriteria : excelBaseObject


@end

@interface excelIconCriterion : excelBaseObject

@property excelE196 conditionOperator;  // Returns the operator for the conditional format.
@property excelE312 iconCriterionIcon;
@property (readonly) NSInteger iconCriterionIndex;  // The index of the icon criterion. Read only.
@property excelE306 iconCriterionType;  // Returns or sets the condition value type of the icon criterion. Read/Write.
@property (copy) id iconCriterionValue;


@end

@interface excelIconSetFormatCondition : excelBaseObject

@property (copy, readonly) excelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property (copy) excelFormatConditionIconSet *formatConditionIconSet;
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) excelE195 formatConditionType;  // Return the conditional format type.
@property (copy) NSString *formula;  // Returns or sets the value or expression associated with the conditional format or data validation.
@property (copy, readonly) excelIconCriteria *iconCriteria;  // Returns the IconCriteria for the IconSetCondition object.
@property BOOL percentileValues;
@property excelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property BOOL reverseIconSetOrder;
@property BOOL showIconOnly;
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

@interface excelLabel : excelBaseObject

- (SBElementArray *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property (copy, readonly) excelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
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
@property excelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) excelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property BOOL visible;  // Returns or sets if the control is visible
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// The LinearGradient object transitions through a series of colors in a linear manner along a specific angle.
@interface excelLinearGradient : excelBaseObject

@property (copy, readonly) excelColorstops *colorstops;  // Returns the ColorStops for the LinearGradient object.
@property double linearGradientDegree;  // The angle of the linear gradient fill within a selection.


@end

// Represents a column in a list object.
@interface excelListColumn : excelBaseObject

@property (copy, readonly) excelCell *cellTable;  // Returns the cell table from a specified list column. 
@property (readonly) NSInteger index;
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) excelRange *rangeObject;  // Returns a range object that represents the range to which the specified list column.
@property (copy, readonly) excelRange *totalRow;  // Returns the totals row, if any, from a specified list column.
@property excelE303 totalsCalculation;


@end

// Represents a list object on a worksheet.
@interface excelListObject : excelBaseObject

- (SBElementArray *) listColumns;
- (SBElementArray *) listRows;

@property (readonly) BOOL active;
@property (copy, readonly) excelAutofilter *autofilterObject;  // Returns the autofilter object associated with this list object.
@property (copy, readonly) excelCell *cellTable;  // Returns the cell table from a specified list object. 
@property (copy) NSString *comment;  // Returns or sets the comment of the object.
- (NSString *) default;
@property (copy) NSString *displayName;  // Returns or sets the display name of the object.
@property (readonly) BOOL displayRightToLeft;
@property (copy, readonly) excelRange *headerRow;  // Returns a range object that represents the used range on the specified worksheet. 
@property (copy, readonly) excelRange *insertRow;
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) excelQueryTable *queryTable;  // Returns a query table object that represents the query table that intersects the specified range.
@property (copy) excelRange *rangeObject;  // Returns or sets a range object that represents the range to which the specified list column, object, or row applies.
@property BOOL showAutofilter;  // Returns or sets if the autofilter is implemented in a list object.
@property BOOL showHeaders;  // Returns or sets if headers is implemented in a list object.
@property BOOL showTableStyleColumnStripes;  // The ShowTableStyleColumnStripes property displays banded columns in which even columns are formatted differently from odd columns.
@property BOOL showTableStyleFirstColumn;  // Returns or sets if the first column is displayed for the specified ListObject object.
@property BOOL showTableStyleLastColumn;  // Returns or sets if the last column is displayed for the specified ListObject object.
@property BOOL showTableStyleRowStripes;  // The ShowTableStyleRowStripes property displays banded rows in which even rows are formatted differently from odd rows.
@property (copy, readonly) excelSort *sortObject;  // Returns the sort object associated with this list object.
@property (readonly) excelE253 sourceType;
@property (copy) id tableStyle;  // Returns or sets the style used in the table body.
@property BOOL total;  // Returns or sets if a totals row be implemented in a list object.
@property (copy, readonly) excelRange *totalRow;  // Returns the totals row, if any, from a specified list object.

- (void) clearContents;  // Clears the formulas from the range. Clears the data from a chart but leaves the formatting. Clears all the data, formatting, and formulas from a list object.
- (excelRange *) convertToRange;  // Converts a list object to a normal Excel range.
- (void) resizeRange:(excelRange *)range;

@end

// Represents a row in a list object.
@interface excelListRow : excelBaseObject

@property (readonly) NSInteger index;
@property (copy, readonly) excelRange *rangeObject;  // Returns a range object that represents the range to which the specified list row applies.


@end

@interface excelListbox : excelBaseObject

@property (copy, readonly) excelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property BOOL displayThreeDShading;  // Returns or sets whether a 3-d effect will be employed when displaying the control
@property BOOL enabled;  // Returns or sets if the control is enabled
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or sets the height of the control.
@property double leftPosition;  // Returns or sets the left position of the control
@property (copy) NSString *linkedCell;  // Returns or sets reference to a call which contains the value of the control.
@property (copy) NSString *listFillRange;  // Returns or sets the items which are contained in the control.
@property NSInteger listIndex;  // Returns or sets the selected item.
@property BOOL locked;  // Returns or sets whether the control is locked for editing.
@property excelE290 multiSelect;  // Returns or sets the multiple selection setting.
@property (copy) NSString *name;  // Returns or sets the name of this control
@property (readonly) NSInteger numberOfItemsInList;
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property excelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) excelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned. 
@property NSInteger value;  // Returns or sets the current value of the control.
@property BOOL visible;  // Returns or sets if the control is visible
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// Represents a defined name for a range of cells. Named items can be either built-in names, such as Database, Print_Area, and Auto_Open, or custom names.
@interface excelNamedItem : excelBaseObject

@property (copy) NSString *category;  // Returns or sets the category for the specified name in the language of the macro.
@property (copy) NSString *categoryLocal;  // Returns or sets the category for the specified name, in the language of the user, if the name refers to a custom function or command.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property excelE240 macroType;  // Returns or sets what the name refers to.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *nameLocal;  // Returns or sets the name of the object, in the language of the user.
@property (copy) NSString *referenceLocal;  // Returns or sets the formula that the name refers to. The formula is in the language of the user, and it's in A1-style notation, beginning with an equal sign.
@property (copy) NSString *referenceLocalR1c1;  // Returns or sets the formula that the name refers to. This formula is in the language of the user, and it's in R1C1-style notation, beginning with an equal sign.
@property (copy) NSString *referenceR1c1;  // Returns or sets the formula that the name refers to. The formula is in the language of the macro, and it's in R1C1-style notation, beginning with an equal sign.
@property (copy, readonly) excelRange *referenceRange;  // Returns the range object referred to by this object.
@property (copy) NSString *references;  // Returns or sets the formula that the name is defined to refer to, in the language of the macro and in A1-style notation, beginning with an equal sign.
@property (copy) NSString *shortcutKey;  // Returns or sets the shortcut key for a name defined as a custom Microsoft Excel 4.0 macro command.
@property (copy) NSString *value;  // Returns or sets a string containing the formula that the name is defined to refer to. The string is in A1-style notation in the language of the macro, and it begins with an equal sign.
@property BOOL visible;  // Returns or sets if the object is visible.


@end

@interface excelNegativeBarFormat : excelBaseObject

@property (copy, readonly) excelFormatColor *databarBarColor;
@property (copy, readonly) excelFormatColor *databarBorderColor;  // Returns the DataBarBorder for the Databar object.
@property excelE320 negativeBarBorderColorType;  // Returns or sets the type of border of the databar
@property excelE320 negativeBarColorType;  // Returns or sets the type of fill of the databar


@end

@interface excelOptionButton : excelBaseObject

- (SBElementArray *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object. 
@property (copy, readonly) excelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL displayThreeDShading;
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) excelGroupbox *groupBox;
@property double height;  // Returns or set the height of the object.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property (copy) NSString *linkedCell;
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property excelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) excelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property excelE285 value;
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents an outline on a worksheet.
@interface excelOutline : excelBaseObject

@property BOOL automaticStyles;  // Returns or sets if the outline uses automatic styles.
@property excelE233 summaryColumn;  // Returns or sets the location of the summary columns in the outline, as shown in the following table.
@property excelE232 summaryRow;  // Returns or sets the location of the summary rows in the outline, as shown in the following table.

- (void) showLevelsRowLevels:(NSInteger)rowLevels columnLevels:(NSInteger)columnLevels;  // Displays the specified number of row and/or column levels of an outline.

@end

// Represents the page setup description. The page setup object contains all page setup attributes, left margin, bottom margin, paper size, and so on, as properties.
@interface excelPageSetup : excelBaseObject

@property BOOL blackAndWhite;  // Returns or sets if elements of the document will be printed in black and white.
@property double bottomMargin;  // Returns or sets the size of the bottom margin, in points.
@property (copy) NSString *centerFooter;  // Returns or sets the center part of the footer.
@property (copy, readonly) excelGraphic *centerFooterPicture;  // Returns or sets a graphic object that represents the picture for the center section of the footer. Used to set attributes about the picture.
@property (copy) NSString *centerHeader;  // Returns or sets the center part of the header.
@property (copy, readonly) excelGraphic *centerHeaderPicture;  // Returns or sets a graphic object that represents the picture for the center section of the footer. Used to set attributes about the picture.
@property BOOL centerHorizontally;  // Returns or sets  if the sheet is centered horizontally on the page when it's printed.
@property BOOL centerVertically;  // Returns or sets if the sheet is centered vertically on the page when it's printed.
@property excelE176 chartSize;  // Returns or sets the way a chart is scaled to fit on a page.
@property BOOL draft;  // Returns or sets if the sheet will be printed without graphics.
@property NSInteger firstPageNumber;  // Returns or sets the first page number that will be used when this sheet is printed.
@property NSInteger fitToPagesTall;  // Returns or sets the number of pages tall the worksheet will be scaled to when it's printed. Applies only to worksheets.
@property NSInteger fitToPagesWide;  // Returns or sets the number of pages wide the worksheet will be scaled to when it's printed. Applies only to worksheets.
@property double footerMargin;  // Returns or sets the distance from the bottom of the page to the footer, in points.
@property double headerMargin;  // Returns or sets the distance from the top of the page to the header, in points.
@property (copy) NSString *leftFooter;  // Returns or sets the left part of the footer.
@property (copy, readonly) excelGraphic *leftFooterPicture;  // Returns or sets a graphic object that represents the picture for the left section of the footer. Used to set attributes about the picture.
@property (copy) NSString *leftHeader;  // Returns or sets the left part of the header.
@property (copy, readonly) excelGraphic *leftHeaderPicture;  // Returns or sets a graphic object that represents the picture for the left section of the header. Used to set attributes about the picture.
@property double leftMargin;  // Returns or sets the size of the left margin, in points.
@property excelE164 order;  // Returns or sets the order that Microsoft Excel uses to number pages when printing a large worksheet.
@property excelE170 pageOrientation;  // Returns or set if the printing mode will be portrait or landscape.
@property excelE212 printExcelComments;  // Returns or sets the way comments are printed with the sheet.
@property (copy) NSString *printArea;  // Returns or sets the range to be printed, as a string using A1-style references in the language of the macro.
@property BOOL printGridlines;  // Returns or sets if cell gridlines are printed on the page. Applies only to worksheets.
@property BOOL printHeadings;  // Returns or sets if row and column headings are printed with this page. Applies only to worksheets.
@property BOOL printNotes;  // Returns or sets if cell notes are printed as end notes with the sheet. Applies only to worksheets.
@property (copy) NSArray *printQuality;  // Set/Gets a two element list where 1 is the horizontal print quality and 2 is the vertical print quality
@property (copy) NSString *printTitleColumns;  // Returns or sets the columns that contain the cells to be repeated on the left side of each page, as a string in A1-style notation in the language of the macro.
@property (copy) NSString *printTitleRows;  // Returns or sets the rows that contain the cells to be repeated at the top of each page, as a string in A1-style notation in the language of the macro.
@property (copy) NSString *rightFooter;  // Returns or sets the right part of the footer.
@property (copy, readonly) excelGraphic *rightFooterPicture;  // Returns or sets a graphic object that represents the picture for the right section of the footer. Used to set attributes about the picture.
@property (copy) NSString *rightHeader;  // Returns or sets the right part of the header.
@property (copy, readonly) excelGraphic *rightHeaderPicture;  // Returns or sets a graphic object that represents the picture for the right section of the header. Used to set attributes about the picture.
@property double rightMargin;  // Returns or sets the size of the right margin, in points.
@property double topMargin;  // Returns or sets the size of the top margin, in points.
@property excelE280 zoom;  // Returns or sets a percentage, between 10 and 400 percent, by which Microsoft Excel will scale the worksheet for printing. Applies only to worksheets.


@end

// Represents a pane of a window. Pane objects exist only for worksheets and Microsoft Excel 4.0 macro sheets.
@interface excelPane : excelBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property NSInteger scrollColumn;  // Returns or sets the number of the leftmost column in the pane
@property NSInteger scrollRow;  // Returns or sets the number of the row that appears at the top of the pane.
@property (copy, readonly) excelRange *visibleRange;  // Returns a Range object that represents the range of cells that are visible in the pane. If a column or row is partially visible, it's included in the range.


@end

// Contains information about a specific phonetic text string in a cell.
@interface excelPhonetic : excelBaseObject

@property excelE205 characterType;  // Returns or sets the type of phonetic text in the specified cell.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property excelE206 phoneticAlignment;  // Returns or sets the alignment for the specified phonetic text.
@property (copy) NSString *phoneticText;  // Returns or sets the text for the specified object.
@property BOOL visible;  // Returns or sets if the object is visible


@end

// Used as the base class for the PivotDataAxis, PivotFilterAxis, and PivotGroupAxis objects. There are no properties or methods which with you can use the PivotAxis object directly.
@interface excelPivotAxis : excelBaseObject

- (SBElementArray *) pivotLines;


@end

// Represents the memory cache for a PivotTable report.
@interface excelPivotCache : excelBaseObject

@property (copy, readonly) id ADOConnection;  // Returns an ADO connection object if the PivotTable cache is connected to an OLE DB data source.
@property (readonly) BOOL OLAP;  // Returns True if the PivotTable cache is connected to an Online Analytical Processing server.
@property (copy) NSString *SQLQuery;  // Returns or sets the SQL query string used with the specified ODBC data source.
@property BOOL backgroundQuery;  // Returns or sets if queries for the pivot table report or query table are performed asynchronously, in the background.
@property (copy) NSString *commandText;  // Returns or sets the command string for the specified data source.
@property excelE251 commandType;  // Returns or sets a constant describing the value type of the CommandText property.
@property (copy) NSString *connection;  // Returns or sets a string that contains one of the following. ODBC settings that enable Microsoft Excel to connect to an ODBC data source, a URL that enables Microsoft Excel to connect to a Web data source, or a file that specifies a database or Web query.
@property BOOL enableRefresh;  // Returns or sets if the pivot table cache can be refreshed by the user.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (readonly) BOOL isConnected;  // Returns True if the MaintainConnection property is True and the PivotTable cache is currently connected to its source.
@property (copy) NSString *localConnection;  // Returns or sets the connection string to an offline cube file.
@property BOOL maintainConnection;  // True if the connection to the specified data source is maintained after the refresh and until the workbook is closed.
@property (readonly) NSInteger memoryUsed;  // Returns the amount of memory currently being used by the object, in bytes.
@property excelE907 missingItemsLimit;  // Returns or sets the maximum quantity of unique items per PivotTable field that are retained even when they have no supporting data in the cache records.
@property BOOL optimizeCache;  // Returns or set if the pivot table cache is optimized when it's constructed.
@property (readonly) excelE250 queryType;  // Indicates the type of query used by Microsoft Excel to populate the PivotTable cache.
@property (readonly) NSInteger recordCount;  // Returns the number of records in the pivot table cache or the number of cache records that contain the specified item.
@property (copy, readonly) NSDate *refreshDate;  // Returns the date on which the pivot cache was last refreshed.
@property (copy, readonly) NSString *refreshName;  // Returns the name of the person who last refreshed the pivot cache. 
@property BOOL refreshOnFileOpen;  // Returns or sets if the pivot table cache or query table is automatically updated each time the workbook is opened.
@property NSInteger refreshPeriod;  // Returns or sets the number of minutes between refreshes.
@property excelXRbC robustConnect;  // Returns or sets how the PivotTable cache connects to its data source.
@property BOOL savePassword;  // Returns or set if password information in an ODBC connection string is saved with the specified query. if false, the password is removed.
@property (copy) NSString *sourceConnectionFile;  // Returns or sets a String indicating the Microsoft Office Data Connection file or similar file that was used to create the PivotTable.
@property excelE279 sourceData;  // Returns or sets the data source for the pivot table.
@property (readonly) excelE157 sourceType;  // Returns a value that identifies the type of item being published.
@property BOOL upgradeOnRefresh;  // Contains information on whether to upgrade the PivotCache and all connected PivotTables on the next refresh.
@property BOOL useLocalConnection;  // True if the LocalConnection property is used to specify the string that enables Microsoft Excel to connect to a data source.
@property (readonly) excelE900 version;  // Returns the version of Microsoft Excel in which the PivotCache was created.
@property (copy, readonly) id workbookConnection;  // Establishes a connection between the current workbook and the PivotCache object.

- (void) createPivotTableTableDestination:(NSString *)tableDestination tableName:(NSString *)tableName readData:(NSString *)readData defaultVersion:(NSString *)defaultVersion;  // Creates a PivotTable report based on a PivotCache object.
- (void) refresh;  // Updates the pivot table cache or query table.

@end

// Represents a cell in a PivotTable report.
@interface excelPivotCell : excelBaseObject

@property (copy, readonly) NSString *MDX;
@property (readonly) excelE909 cellChanged;
@property (readonly) excelE150 customSubtotalFunction;  // Returns the custom subtotal function field setting of a PivotCell object.
@property (copy, readonly) excelPivotField *dataField;  // Returns a PivotField object that corresponds to the selected data field.
@property (copy, readonly) id dataSourceValue;
@property (readonly) excelE908 pivotCellType;  // Returns a constant that identifies the PivotTable entity the cell corresponds to.
@property (copy, readonly) excelPivotLine *pivotColumnLine;  // Returns the PivotLine on a column for a specific PivotCell object.
@property (copy, readonly) excelPivotField *pivotField;  // Returns a PivotField object that represents the PivotTable field containing the upper-left corner of the specified range.
@property (copy, readonly) excelPivotItem *pivotItem;  // Returns a PivotItem object that represents the PivotTable item containing the upper-left corner of the specified range.
@property (copy, readonly) excelPivotLine *pivotRowLine;  // Returns the PivotLine on a row for a specific PivotCell object.
@property (copy, readonly) excelPivotTable *pivotTable;  // Returns a PivotTable object that represents the PivotTable report associated with the PivotCell.
@property (copy, readonly) excelRange *range;  // Returns a Range object that represents the range the specified PivotCell applies to.
@property (copy) id rowItems;  // Returns a PivotItemList collection that corresponds to the items on the category axis that represent the selected cell.

- (void) allocateChange;
- (void) discardChange;

@end

// Represents a field in a pivot table.
@interface excelPivotField : excelBaseObject

- (SBElementArray *) childItems;
- (SBElementArray *) hiddenItems;
- (SBElementArray *) parentItems;
- (SBElementArray *) pivotItems;
- (SBElementArray *) calculatedItems;
- (SBElementArray *) pivotFilters;

@property (readonly) BOOL allItemsVisible;  // Used to retrieve a Boolean value that indicates whether or not any manual filtering is applied to the PivotField.
@property (readonly) NSInteger autoShowCount;  // Returns the number of top or bottom items that are automatically shown in the pivot field.
@property (copy, readonly) NSString *autoShowField;  // Returns the name of the data field used to determine the top or bottom items that are automatically shown in the pivot field.
@property (readonly) excelE270 autoShowRange;  // Returns position top if the top items are shown automatically in the pivot field, returns position bottom if the bottom items are shown.
@property (readonly) excelE269 autoShowType;  // Returns type_automatic if auto show is enabled for the pivot field, or  returns type_manual if auto show is disabled.
@property (readonly) NSInteger autoSortCustomSubtotal;  // Returns the name of the custom subtotal used to sort the specified PivotTable field automatically.
@property (copy, readonly) NSString *autoSortField;  // Returns the name of the data field used to sort the pivot field automatically.
@property (readonly) excelE228 autoSortOrder;  // Returns the order used to sort the pivot field automatically. 
@property (copy, readonly) excelPivotLine *autoSortPivotLine;  // Returns the name of the PivotLine used to sort the specified PivotTable field automatically.
@property (copy) NSString *baseField;  // Returns or sets the base field for a custom calculation.
@property (copy) NSString *baseItem;  // Returns or sets the item in the base field for a custom calculation.
@property excelE209 calculation;  // Returns or sets the type of calculation done by the specified pivot field.
@property (copy) NSString *caption;  // The label text for the pivot field.
@property (copy, readonly) excelPivotField *childField;  // Returns a pivot field object that represents the child pivot field for the specified field, if the field is grouped and has a child field.
@property (copy, readonly) excelCubeField *cubeField;  // Returns the CubeField object from which the specified PivotTable field is descended.
@property (copy) NSString *currentPage;  // Returns or sets the current page showing for the page field, only valid for page fields.
@property (copy) id currentPageList;  // Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report.
@property (copy) NSString *currentPageName;  // Returns or sets the currently displayed page of the specified PivotTable report.
@property (copy, readonly) excelRange *dataRange;  // Returns a range object.  For a data field the range is the data contained in the field, for a row, column or page field it is the items in the field.
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
@property excelE150 function;  // Returns or sets the function used to summarize the pivot field.
@property (readonly) NSInteger groupLevel;  // Returns the placement of the specified field within a group of fields, if the field is a member of a grouped set of fields.
@property BOOL hidden;  // This property is used to hide the individual levels of an OLAP hierarchy.
@property (copy) id hiddenItemsList;  // Returns or sets an Object specifying an array of strings that are hidden items for a PivotTable field.
@property BOOL includeNewItemsInFilter;  // Allows developers to specify whether excluded or included items should be tracked when manual filtering is applied to the PivotField.
@property (readonly) BOOL isCalculated;  // Returns true if the pivot field or item is a calculated field or item.
@property (readonly) BOOL isMemberProperty;  // Returns True when the PivotField contains member properties.
@property (copy, readonly) excelRange *labelRange;  // Returns a range object that represents the cell, or cells, that contain the field label.
@property BOOL layoutBlankLine;  // True if a blank row is inserted after the specified row field in a PivotTable report.
@property BOOL layoutCompactRow;  // Specifies whether or not a PivotField is compacted , items of multiple PivotFields are displayed in a single column, when rows are selected.
@property excelE910 layoutForm;  // Returns or sets the way the specified PivotTable items appear.
@property BOOL layoutPagebreak;  // True if a page break is inserted after each field.
@property excelE902 layoutSubtotalLocation;  // Returns or sets the position of the PivotTable field subtotals in relation to either above or below the specified field.
@property (copy) NSString *memberPropertyCaption;  // Setting the MemberPropertyCaption property controls which member property is used as caption for a given level.
@property (readonly) NSInteger memoryUsed;  // Returns the amount of memory currently being used by the object, in bytes.
@property (copy) NSString *name;  // Returns or sets the name of the pivot field.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the pivot field. 
@property (copy, readonly) excelPivotField *parentField;  // Returns a pivot field object that represents the pivot field that's the group parent of the object. The field must be grouped and have a parent field.
@property (readonly) excelE155 pivotFieldDataType;  // Returns an enumeration describing the type of data in the pivot field.
@property excelE208 pivotFieldOrientation;  // The location of the field in the pivot table.
@property NSInteger position;  // Position of the field, first, second, third, and so on, among all the fields in its orientation, rows, columns, pages, data.
@property NSInteger propertyOrder;  // Valid only for PivotTable fields that are member property fields.
@property (copy, readonly) excelPivotField *propertyParentField;  // Returns a PivotField object representing the field to which the properties in this field pertain.
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
- (void) autoShowType:(excelE269)type range:(excelE270)range count:(NSInteger)count field:(NSString *)field;  // Displays the number of top or bottom items for a row, page, or column field in the specified PivotTable report
- (void) autoSortSortOrder:(excelE228)sortOrder sortField:(NSString *)sortField;  // Establishes automatic pivot table field-sorting rules.
- (void) clearLabelFilters;  // This method deletes all label filters or all date filters in the PivotFilters collection of the PivotField.
- (void) clearValueFilters;  // Calling this method deletes all value filters in the PivotFilters collection of the PivotField.
- (BOOL) getSubtotalsSubtotalIndex:(excelE261)subtotalIndex;  // Returns subtotals displayed with the specified field. Valid only for nondata fields.
- (void) setSubtotalsSubtotalIndex:(excelE268)subtotalIndex value:(BOOL)value;  // Sets subtotals displayed with the specified field. Valid only for nondata fields.

@end

@interface excelCalculatedField : excelPivotField


@end

@interface excelColumnField : excelPivotField


@end

@interface excelDataField : excelPivotField


@end

@interface excelHiddenField : excelPivotField


@end

@interface excelPageField : excelPivotField


@end

// PivotTable Advanced Filter.
@interface excelPivotFilter : excelBaseObject

@property (readonly) BOOL active;  // Returns whether the specified PivotFilter is active.
@property (copy, readonly) excelCubeField *dataCubeField;  // This property is applicable only to non-OLAP PivotTables and provides the Value field ,PivotField in the Values area, being filtered by for a value filter.
@property (copy, readonly) excelPivotField *dataField;  // This property is applicable only to non-OLAP PivotTables and provides the Value field ,PivotField in the Values area, being filtered by for a value filter.
@property (copy, readonly) NSString *objectDescription;  // Provides an optional description for the PivotFilter object.
@property (readonly) excelE911 filterType;  // Specifies the type of filter to be applied.
@property (readonly) BOOL isMemberPropertyFilter;  // Specifies whether the label filter is based on the PivotItem captions of a member property of the field or on the PivotItem captions of the PivotField itself.
@property (copy, readonly) excelPivotField *memberPropertyField;  // This property specifies the member property PivotField on which the label filter is based.
@property (copy, readonly) NSString *name;  // This property provides the option of naming filters for reference.
@property NSInteger order;  // Specifies the evaluation order of the filter among all Value filters applied to the entire PivotTable.
@property (copy, readonly) excelPivotField *pivotField;  // Specifies the PivotField to which the filter is applied.
@property (copy, readonly) id value1;  // This property is a user-supplied parameter to define a filter for a PivotField.
@property (copy, readonly) id value2;  // This property is a user-supplied parameter to define a filter for a PivotField.


@end

@interface excelActiveFilter : excelPivotFilter


@end

// Represents a formula used to calculate results in a pivot table report.
@interface excelPivotFormula : excelBaseObject

@property NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property (copy) NSString *standardFormula;  // Returns or sets a String specifying formulas with standard English formatting.
@property (copy) NSString *value;  // Returns or sets the name of the specified formula in the pivot table formula.


@end

// Represents an item in a pivot field. The items are the individual data entries in a field category.
@interface excelPivotItem : excelBaseObject

- (SBElementArray *) childItems;

@property (copy) NSString *caption;  // The label text for the pivot item.
@property (copy, readonly) excelRange *dataRange;  // Returns a range object specifying the data qualified by the item.
@property BOOL drilledDown;  // True if the flag for the specified PivotTable field or PivotTable item is set to 
@property (copy) NSString *formula;  // Returns or sets the pivot item's formula, in A1-style notation and in the language of the macro.
@property (readonly) BOOL isCalculated;  // Returns true if the pivot  item is a calculated item.
@property (copy, readonly) excelRange *labelRange;  // Returns a range object that represents all the pivot table cells that contain the item.
@property (copy) NSString *name;  // Returns or sets the name of the pivot item.
@property (copy, readonly) excelParentItem *parentItem;  // Returns a pivot item object that represents the parent pivot item in the parent pivot field object, the field must be grouped so that it has a parent.
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

@interface excelCalculatedItem : excelPivotItem


@end

@interface excelChildItem : excelPivotItem


@end

@interface excelColumnItem : excelPivotItem


@end

@interface excelHiddenItem : excelPivotItem


@end

@interface excelParentItem : excelPivotItem


@end

// The PivotLines object is a collection of lines in a PivotTable, containing all lines on rows or columns of the pivot.
@interface excelPivotLine : excelBaseObject

@property (readonly) excelE912 lineType;  // Returns an XlPivotLineType constant that indicates the type of PivotLine.
@property (copy, readonly) id pivotLineCells;  // Returns a collection of PivotCell objects in a PivotLine.
@property (readonly) NSInteger position;  // Returns or sets the position of the PivotLine object.


@end

// Represents a pivot table on a worksheet.
@interface excelPivotTable : excelBaseObject

- (SBElementArray *) columnFields;
- (SBElementArray *) dataFields;
- (SBElementArray *) hiddenFields;
- (SBElementArray *) pageFields;
- (SBElementArray *) pivotFields;
- (SBElementArray *) rowFields;
- (SBElementArray *) calculatedFields;
- (SBElementArray *) pivotFormulas;
- (SBElementArray *) calculatedMembers;
- (SBElementArray *) activeFilters;
- (SBElementArray *) cubeFields;
- (SBElementArray *) slicers;

@property NSInteger CompactRowIndent;  // Returns or sets the indent increment for PivotItems when compact row layout form is turned on.
@property excelE903 allocation;
@property excelE905 allocationMethod;
@property excelE904 allocationValue;
@property (copy) NSString *allocationWeightExpression;
@property BOOL allowMultipleFilters;  // Sets or retrieves a value that indicates whether a PivotField can have multiple filters applied to it at the same time.
@property (copy) NSString *alternativeText;  // Returns or sets the descriptive alternative text string for a ShapeRange object when the object is saved to a Web page.
@property NSInteger cacheIndex;  // Returns or sets the index number of the pivot table cache.
@property BOOL calculatedMembersInFilters;
@property (copy, readonly) id changeList;
@property BOOL columnGrand;  // Returns or sets if the pivot table shows grand totals for columns.
@property (copy, readonly) excelRange *columnRange;  // Returns a range object that represents the range that contains the pivot table column area.
@property (copy) NSString *compactLayoutColumnHeader;  // Specifies the caption that is displayed in the column header of a PivotTable when in compact row layout form.
@property (copy) NSString *compactLayoutRowHeader;  // Specifies the caption that is displayed in the row header of a PivotTable when in compact row layout form.
@property (copy, readonly) excelRange *dataBodyRange;  // Returns a range object that represents the range that contains the pivot table data area.
@property (copy, readonly) excelRange *dataLabelRange;  // Returns a range object that represents the range that contains the labels for the pivot table data fields. 
@property (copy, readonly) excelPivotField *dataPivotField;  // Returns a PivotField object that represents all the data fields in a PivotTable.
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
@property excelE901 layoutRowDefault;  // This property specifies the layout settings for PivotFields when they are added to the PivotTable for the first time.
@property (copy) NSString *location;  // Gets or sets a String that represents the top-left cell in the body of the specified PivotTable.
@property BOOL manualUpdate;  // Returns or sets if the pivot table is recalculated only at the user's request.
@property (copy, readonly) NSString *mdx;  // Returns a String indicating the MDX, Multidimensional-Expression, that would be sent to the provider to populate the current PivotTable view.
@property BOOL mergeLabels;  // Returns or sets if pivot table outer-row item, column item, subtotal, and grand total labels use merged cells.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *nullString;  // Returns or sets the string displayed in cells that contain null values when the display null string property is true. The default value is an empty string.
@property excelE164 pageFieldOrder;  // Returns or sets the order in which page fields are added to the pivot table layout.
@property (copy) NSString *pageFieldStyle;  // Returns or sets the style used in the bound page field area.
@property NSInteger pageFieldWrapCount;  // Returns or sets the number of pivot table page fields in each column or row.
@property (copy, readonly) excelRange *pageRange;  // Returns a range object that represents the range that contains the pivot table page area.
@property (copy, readonly) excelRange *pageRangeCells;  // Returns a range object that represents the cells in the pivot table containing only the page fields and item drop-down lists.
@property (copy, readonly) excelPivotCache *pivotCache;  // Returns a pivot cache object that represents the cache for the specified pivot table
@property (copy, readonly) excelPivotAxis *pivotColumnAxis;  // Returns a PivotAxis object representing the entire column axis 
@property (copy, readonly) excelPivotAxis *pivotRowAxis;  // Returns a PivotAxis object representing the entire row axis 
@property (copy) NSString *pivotSelection;  // Returns or sets the pivot table selection, in standard pivot table selection format.
@property (copy) NSString *pivotSelectionStandard;  // Returns or sets a String indicating the PivotTable selection in standard PivotTable report format using English settings.
@property BOOL preserveFormatting;  // Returns or sets if pivot table formatting is preserved when the pivot table is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.
@property BOOL printDrillIndicators;  // Specifies whether or not drill indicators are printed with the PivotTable.
@property BOOL printTitles;  // set/get whether the print titles for the worksheet are set based on the PivotTable report.
@property (copy, readonly) NSDate *refreshDate;  // Returns the date on which the pivot table was last refreshed.
@property (copy, readonly) NSString *refreshName;  // Returns the name of the person who last refreshed the pivot table data.
@property BOOL repeatItemsOnEachPrintedPage;  // True if row, column, and item labels appear on the first row of each page when the specified PivotTable report is printed. False if labels are printed only on the first page.
@property BOOL rowGrand;  // Returns or sets if the pivot table shows grand totals for rows.
@property (copy, readonly) excelRange *rowRange;  // Returns a range object that represents the range including the pivot table row area.
@property BOOL saveData;  // Returns or sets if data for the pivot table is saved with the workbook. If false only the pivot table definition is saved.
@property excelE214 selectionMode;  // Returns or sets the pivot table structured selection mode.
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
@property excelE279 sourceData;  // Either a range reference as a string or a list of SQL commands
@property BOOL subtotalHiddenPageItems;  // Returns or sets if hidden page field items in the pivot table are included in row and column subtotals, block totals, and grand totals.
@property (copy) NSString *summary;
@property (copy, readonly) excelRange *tableRange1;  // Returns a range object that represents the range containing the entire pivot table, but doesn't include page fields.
@property (copy, readonly) excelRange *tableRange2;  // Returns a range object that represents the range containing the entire pivot table, including page fields.
@property (copy) NSString *tableStyle;  // Returns or sets the style used in the pivot table body.
@property (copy) id tableStyle2;  // The TableStyle2 property specifies the PivotTable style currently applied to the PivotTable.
@property (copy) NSString *tag;  // Returns or sets a string saved with the pivot table.
@property BOOL totalsAnnotation;  // True if an asterisk is displayed next to each subtotal and grand total value in the specified PivotTable report if the report is based on an OLAP data source.
@property (copy) NSString *vacatedStyle;  // Returns or sets the style applied to cells vacated when the pivot table is refreshed.
@property (copy) NSString *value;  // Returns or set the name of the pivot table.
@property (readonly) excelE900 version;
@property BOOL viewCalculatedMembers;  // When set to True, calculated members for Online Analytical Processing, OLAP, PivotTables can be viewed.
@property BOOL visualTotals;  // True to enable Online Analytical Processing, OLAP, PivotTables to retotal after an item has been hidden from view.
@property BOOL visualTotalsForSets;

- (excelPivotField *) addDataFieldField:(id)field caption:(NSString *)caption function:(NSString *)function;  // Adds a data field to a PivotTable report.
- (void) addFieldsToPivotTableRowFields:(NSArray *)rowFields columnFields:(NSArray *)columnFields pageFields:(NSArray *)pageFields addToTable:(BOOL)addToTable;  // Adds row, column, and page fields to a pivot table.
- (void) allocateChanges;
- (void) changeConnectionConnection:(excelWorkbookConnection *)connection;  // Changes the connection of the specified PivotTable.
- (void) changePivotCachePivotCache:(NSString *)pivotCache;  // Changes the PivotCache of the specified PivotTable.
- (void) clearTable;  // The ClearTable method is used for clearing a PivotTable.
- (void) commitChanges;
- (void) convertToFormulasConvertFilters:(BOOL)convertFilters;  // The ConvertToFormulas method is new in 1st_Excel12 and is used for converting a PivotTable to cube formulas.
- (NSString *) createCubeFileFile:(NSString *)file measures:(NSString *)measures levels:(NSString *)levels members:(NSString *)members properties:(NSString *)properties;  // Creates a cube file from a PivotTable report connected to an Online Analytical Processing data source.
- (void) discardChanges;
- (excelRange *) getPivotDataDataField:(NSString *)dataField field1:(NSString *)field1 item1:(NSString *)item1 field2:(NSString *)field2 item2:(NSString *)item2 field3:(NSString *)field3 item3:(NSString *)item3 field4:(NSString *)field4 item4:(NSString *)item4 field5:(NSString *)field5 item5:(NSString *)item5 field6:(NSString *)field6 item6:(NSString *)item6 field7:(NSString *)field7 item7:(NSString *)item7 field8:(NSString *)field8 item8:(NSString *)item8 field9:(NSString *)field9 item9:(NSString *)item9 field10:(NSString *)field10 item10:(NSString *)item10 field11:(NSString *)field11 item11:(NSString *)item11 field12:(NSString *)field12 item12:(NSString *)item12 field13:(NSString *)field13 item13:(NSString *)item13 field14:(NSString *)field14 item14:(NSString *)item14;  // Returns a Range object with information about a data item in a PivotTable report.
- (double) getPivotTableDataName:(NSString *)name;  // Retrieve data from a pivot table
- (NSArray *) getVisibleFields;  // Returns a list of all the visible pivot fields. Visible pivot fields are shown as row, column, page or data fields.
- (void) listFormulas;  // Creates a list of calculated pivot table items and fields on a separate worksheet.
- (void) makeConnection;  // Establishes a connection for the specified PivotTable cache.
- (void) pivotSelectName:(NSString *)name mode:(excelE214)mode;  // Selects part of a pivot table.
- (void) refreshDataSourceValues;
- (BOOL) refreshTable;  // Refreshes the pivot table from the source data. Returns true if it's successful.
- (void) repeatAllLabelsRepeat:(excelE906)repeat;
- (void) rowAxisLayoutLayout:(excelE901)layout;  // This method is used for simultaneously setting layout options for all existing PivotFields.
- (void) saveAsODCODCFileName:(NSString *)ODCFileName objectDescription:(NSString *)objectDescription keywords:(NSString *)keywords;  // Saves the PivotTable cache source as a Microsoft Office Data Connection file.
- (void) showPagesPageField:(NSString *)pageField;  // Creates a new pivot table for each item in the page field. Each new pivot table is created on a new worksheet.
- (void) subtotalLocationLocation:(excelE902)location;  // This method changes the subtotal location for all existing PivotFields.
- (void) update;  // Updates the link or pivot table.

@end

// Represents a worksheet table built from data returned from an external data source, such as an SQL server.
@interface excelQueryTable : excelBaseObject

@property (copy) NSArray *FileMakerFields;
@property NSInteger FileMakerNumCriteria;
@property BOOL FileMakerUseTable;  // Returns or sets if the query uses a table rather than a layout. Only applicable for FileMaker 7 or above
@property BOOL adjustColumnWidth;  // Returns or sets if the column widths are automatically adjusted for the best fit each time you refresh the specified query table.
@property BOOL backgroundQuery;  // Returns or sets if queries for the query table are performed asynchronously.
@property excelE251 commandType;  // Returns or sets one of the XlCmdType constants listed in the following table in the remarks section.
@property (copy) NSString *connection;  // Returns or sets a string that contains one of the following: ODBC settings that enable Microsoft Excel to connect to an ODBC data source; a URL that enables Microsoft Excel to connect to a Web data source; or a file that specifies a database or Web query.
@property (copy, readonly) excelRange *destination;  // Returns the cell in the upper-left corner of the query table destination range, the range where the resulting query table will be placed. The destination range must be on the worksheet that contains the query table object.
@property BOOL enableEditing;  // Returns or sets if the user can edit the specified query table. False if the user can only refresh the query table.
@property BOOL enableRefresh;  // Returns or sets if the query table can be refreshed by the user.
@property (readonly) BOOL fetchedRowOverflow;  // Returns true if the number of rows returned by the last use of the refresh method is greater than the number of rows available on the worksheet.
@property BOOL fieldNames;  // Returns or sets if field names from the data source appear as column headings for the returned data.
@property BOOL fillAdjacentFormulas;  // Returns or sets if formulas to the right of the specified query table are automatically updated whenever the query table is refreshed.
@property BOOL hasAutoformat;  // Returns or sets if the query table is automatically formatted when it's refreshed or when fields are moved.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *postText;  // Returns or sets the string used with the post method of inputting data into a Web server to return data from a Web query.
@property (readonly) excelE250 queryType;  // Returns the type of query used by Microsoft Excel to populate the query table.
@property BOOL refreshOnFileOpen;  // Returns or sets if the query table is automatically updated each time the workbook is opened.
@property excelE191 refreshStyle;  // Returns or sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a recordset returned by a query.
@property (readonly) BOOL refreshing;  // Returns true if there's a background query in progress for the specified query table.
@property (copy, readonly) excelRange *resultRange;  // Returns a range object that represents the area of the worksheet occupied by the specified query table.
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
@property excelE236 textFileParseType;  // Returns or sets the column format for the data in the text file that you're importing into a query table.
@property excelE211 textFilePlatform;  // Returns or sets the origin of the text file you're importing into the query table. This property determines which code page is used during the data import.
@property BOOL textFilePromptOnRefresh;  // Returns or sets if you want to specify the name of the imported text file each time the query table is refreshed.
@property BOOL textFileSemicolonDelimiter;  // Returns or sets if the semicolon character is the delimiter when you import a text file into a query table.
@property BOOL textFileSpaceDelimiter;  // Returns or sets if the space character is the delimiter when you import a text file into a query table.
@property NSInteger textFileStartRow;  // Returns or sets the row number at which text parsing will begin when you import a text file into a query table. Valid values are integers from 1 through 32767. The default value is 1.
@property BOOL textFileTabDelimiter;  // Returns or sets if the tab character is the delimiter when you import a text file into a query table.
@property excelE237 textFileTextQualifier;  // Returns or sets the text qualifier when you import a text file into a query table. The text qualifier specifies that the enclosed data is in text format.
@property (copy) NSString *textFileThousandsSeparator;  // Returns or sets the thousands separator character thatMicrosoft Excel uses when you import a text file into a query table. The default is the system thousands separator character.
@property BOOL useListObject;

- (void) cancelRefresh;  // Cancels all background queries for the specified query table. Use the refreshing property to determine whether a background query is currently in progress.
- (NSDictionary *) getFileMakerCriteriaCriteriaIndex:(NSInteger)criteriaIndex;  // Get the criteria from an existing Excel query that was created against a FileMaker database
- (BOOL) refreshQueryTableBackgroundQuery:(BOOL)backgroundQuery;  // Updates the PivotTable cache or query table.
- (void) setFileMakerCriteriaCriteriaIndex:(NSInteger)criteriaIndex fieldName:(NSString *)fieldName operator:(excelE257)operator_ clauseText:(NSString *)clauseText condition:(excelE258)condition;

@end

// Represents a file in the list of recently used files.
@interface excelRecentFile : excelBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy, readonly) NSString *path;  // Returns the complete path to the file, excluding the final separator and name of the file.


@end

// The RectangularGradient object transitions through a series of colors from the center to one or more sides.
@interface excelRectangularGradient : excelBaseObject

@property (copy, readonly) excelColorstops *colorstops;  // Returns the ColorStops for the LinearGradient object.
@property double rectangularGradientBottom;  // Represents the point or vector that the gradient fill converges to.
@property double rectangularGradientLeft;  // Represents the point or vector that the gradient fill converges to.
@property double rectangularGradientRight;  // Represents the point or vector that the gradient fill converges to.
@property double rectangularGradientTop;  // Represents the point or vector that the gradient fill converges to.


@end

@interface excelRowField : excelPivotField


@end

@interface excelRowItem : excelPivotItem


@end

// Represents a scenario on a worksheet. Represents a scenario on a worksheet. A scenario is a group of input values, called changing cells, that's named and saved.
@interface excelScenario : excelBaseObject

@property (copy) NSString *ExcelComment;  // Returns or sets the comment associated with the scenario. The comment text cannot exceed 255 characters.
@property (copy, readonly) excelRange *changingCells;  // Returns a range object that represents the changing cells for a scenario.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property BOOL hidden;  //  Returns or sets if the scenario is hidden.
@property BOOL locked;  // Returns or sets if the object is locked.
@property (copy) NSString *name;  // Returns or sets the name of the object.

- (void) changeScenarioChangingCells:(excelE267)changingCells values:(NSArray *)values;  // Changes the scenario to have a new set of changing cells and, optionally, scenario values.
- (NSArray *) getValues;  // Returns or sets a list that contains the current values of the changing cells for the scenario.

@end

@interface excelScrollbar : excelBaseObject

@property (copy, readonly) excelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
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
@property excelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property NSInteger smallChange;
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) excelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property NSInteger value;
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents a worksheet.
@interface excelSheet : excelBaseObject

- (SBElementArray *) shapes;
- (SBElementArray *) arcs;
- (SBElementArray *) buttons;
- (SBElementArray *) chartObjects;
- (SBElementArray *) checkboxes;
- (SBElementArray *) dropdowns;
- (SBElementArray *) groupboxes;
- (SBElementArray *) labels;
- (SBElementArray *) lines;
- (SBElementArray *) listboxes;
- (SBElementArray *) namedItems;
- (SBElementArray *) optionButtons;
- (SBElementArray *) ovals;
- (SBElementArray *) pivotTables;
- (SBElementArray *) ranges;
- (SBElementArray *) cells;
- (SBElementArray *) rows;
- (SBElementArray *) columns;
- (SBElementArray *) rectangles;
- (SBElementArray *) scenarios;
- (SBElementArray *) scrollbars;
- (SBElementArray *) spinners;
- (SBElementArray *) textboxes;
- (SBElementArray *) horizontalPageBreaks;
- (SBElementArray *) verticalPageBreaks;
- (SBElementArray *) queryTables;
- (SBElementArray *) ExcelComments;
- (SBElementArray *) hyperlinks;
- (SBElementArray *) listObjects;

@property BOOL autofilterMode;  // Returns or sets if the autofilter drop-down arrows are currently displayed on the sheet. This property is independent of the filter mode property. 
@property (copy, readonly) excelAutofilter *autofilterObject;  // Returns the autofilter object associated with this sheet.
@property (copy, readonly) excelRange *circularReference;  // Returns a range object that represents the range containing the first circular reference on the sheet, or returns missing value if there's no circular reference on the sheet.
@property (readonly) excelE150 consolidationFunction;  // Returns the function code used for the current consolidation.
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
@property (copy, readonly) excelSheet *next;  // Returns a worksheet object that represents the next sheet.
@property (copy, readonly) excelOutline *outlineObject;  // Returns an outline object that represents the outline for the specified worksheet.
@property (copy, readonly) excelPageSetup *pageSetupObject;  // Returns the page setup object associated with this chart.
@property (copy, readonly) excelSheet *previous;  // Returns a worksheet object that represents the previous sheet.
@property (readonly) BOOL protectContents;  // Returns true if the contents of the sheet are protected.
@property (readonly) BOOL protectDrawingObjects;  // Returns true if shapes are protected.
@property (readonly) BOOL protectionMode;  // Returns true if user-interface-only protection is turned on. To turn on user interface protection, use the protect method with the user interface only argument set to true.
@property (copy, readonly) excelProtection *protectionObject;  // Returns a Protection object that represents the protection options of the worksheet.
@property (copy) NSString *scrollArea;  // Returns or sets the range where scrolling is allowed, as an A1-style range reference. Cells outside the scroll area cannot be selected. 
@property (copy, readonly) excelTab *sheetTab;  // Returns the sheet tab of the work sheet
@property (copy, readonly) excelSort *sortObject;  // Returns the sort object in the sheet.
@property (readonly) double standardHeight;  // Returns the standard default height of all the rows in the worksheet, in points. 
@property double standardWidth;  // Returns or sets the standard default width of all the columns in the worksheet.
@property BOOL transitionExpressionEvaluation;  // Returns or sets if Microsoft Excel uses Lotus 1-2-3 expression evaluation rules for the worksheet.
@property (copy, readonly) excelRange *usedRange;  // Returns a range object that represents the used range on the specified worksheet.
@property excelE225 visible;  // Returns or sets if the worksheet is visible.
@property (readonly) excelE151 worksheetType;  // Returns the type of this worksheet.

- (void) circleInvalid;  // Circles invalid entries on the worksheet.
- (void) clearArrows;  // Clears the tracer arrows from the worksheet. Tracer arrows are added by using the auditing feature.
- (void) clearCircles;  // Clears circles from invalid entries on the worksheet.
- (void) copyWorksheetBefore:(excelSheet *)before after:(excelSheet *)after;  // Copies the sheet to another location in the workbook.
- (void) pasteSpecialOnWorksheetFormat:(NSString *)format link:(BOOL)link displayAsIcon:(BOOL)displayAsIcon iconFileName:(NSString *)iconFileName iconIndex:(NSInteger)iconIndex iconLabel:(NSString *)iconLabel noHTMLFormatting:(BOOL)noHTMLFormatting;  // Pastes the contents of the clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.
- (void) pasteWorksheetDestination:(excelE267)destination link:(BOOL)link;  // Pastes the contents of the Clipboard onto the sheet.
- (void) protectWorksheetPassword:(NSString *)password drawingObjects:(BOOL)drawingObjects worksheetContents:(BOOL)worksheetContents scenarios:(BOOL)scenarios userInterfaceOnly:(BOOL)userInterfaceOnly allowFormattingCells:(BOOL)allowFormattingCells allowFormattingColumns:(BOOL)allowFormattingColumns allowFormattingRows:(BOOL)allowFormattingRows allowInsertingColumns:(BOOL)allowInsertingColumns allowInsertingRows:(BOOL)allowInsertingRows allowInsertingHyperlinks:(BOOL)allowInsertingHyperlinks allowDeletingColumns:(BOOL)allowDeletingColumns allowDeletingRows:(BOOL)allowDeletingRows allowSorting:(BOOL)allowSorting allowFiltering:(BOOL)allowFiltering allowUsingPivotTable:(BOOL)allowUsingPivotTable;  // Protects a worksheet so that it cannot be modified.
- (void) resetAllPageBreaks;  // Resets all page breaks on the specified worksheet.
- (void) saveAsFilename:(NSString *)filename fileFormat:(excelE188)fileFormat password:(NSString *)password writeReservationPassword:(NSString *)writeReservationPassword readOnlyRecommended:(BOOL)readOnlyRecommended createBackup:(BOOL)createBackup addToMostRecentlyUsedList:(BOOL)addToMostRecentlyUsedList overwrite:(BOOL)overwrite saveAsLocalLanguage:(BOOL)saveAsLocalLanguage;  // Saves changes into a different file.
- (void) setBackgroundPicturePictureFileName:(NSString *)pictureFileName;  // Sets the background graphic for a worksheet.
- (void) showAllData;  // Makes all rows of the currently filtered list visible. If autofilter is in use, this method changes the arrows to All.
- (void) showDataForm;  // Displays the data form associated with the worksheet.

@end

@interface excelInternationalMacroSheet : excelSheet

@property excelE197 enableSelection;  // Returns or sets what can be selected on the sheet.


@end

@interface excelMacroSheet : excelSheet

@property excelE197 enableSelection;  // Returns or sets what can be selected on the sheet.


@end

// A slicer is a mechanism for filtering data in PivotTable views and cube functions.
@interface excelSlicer : excelBaseObject


@end

// Represents a sort of a range of data.
@interface excelSort : excelBaseObject

@property BOOL matchCase;  // Set to true to perform a case-sensitive sort or set to false to perform non-case sensitive sort. Read/Write.
@property excelE241 sortHeader;  // Specifies whether the first row contains header information. Read/Write.
@property excelE226 sortMethod;  // Specifies the sort method for Chinese/Japanese languages. Read/Write.
@property excelE229 sortOrientation;  // Specifies the orientation for the sort. Read/Write.
@property (copy, readonly) id sortfields;  // Stores the sort state for workbooks, lists, and autofilters. Read-Only.
@property (copy, readonly) excelRange *sortrange;  // Returns a range object that represents the range to which the specified sort applies. Read Only.

- (void) setSortRangeRng:(excelRange *)rng;  // Sets the starting and ending character positions for Sort object.

@end

// The sortfield object contains all the sort information for the worksheet, lists, and autofilter object.
@interface excelSortfield : excelBaseObject

@property (copy) id customOrder;  // Specifies a custom order to sort the fields. Read/Write.
@property excelE302 sortDataOption;  // Specifies how to sort text in the range specified in sortfield object. Read/Write.
@property (copy, readonly) excelRange *sortKey;  // Specifies the range that is currently being sorted on. Read only.
@property excelE301 sortOn;  // Returns or sets what attribute of the cell to sort on. Read/Write.
@property (copy, readonly) id sortOnValues;  // Return the value on which the sort is performed for the specified sortfield object.Read Only.
@property excelE228 sortOrder;  // Determines the sort order for the values specified in the key. Read/Write.
@property NSInteger sortPriority;  // Specifies the priority for the sortfield. Read/Write.

- (void) modifySortKeyRng:(excelRange *)rng;  // Modify the key value by which values are sorted in the field.

@end

// The sortfields collection is a collection of sortfield objects. It allows developers to store a sort state on worksheets, lists, and autofilters.
@interface excelSortfields : excelBaseObject

- (void) addSortfieldKey:(excelRange *)key sorton:(excelE301)sorton order:(excelE228)order customorder:(id)customorder dataoption:(excelE302)dataoption;  // Creates a new sort field and returns a sortfields object.

@end

@interface excelSpinner : excelBaseObject

@property (copy, readonly) excelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
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
@property excelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property NSInteger smallChange;  // Returns or sets the change in value of one click on the spinner control.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) excelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property NSInteger value;  // Returns or sets the current value of the control
@property BOOL visible;  // Returns or sets if the control is visible
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// Represents the sheet tab of a work sheet or chart sheet.
@interface excelTab : excelBaseObject

@property (copy) NSColor *color;
@property excelE103 colorIndex;
@property excelE305 themeColor;
@property double tintAndShade;


@end

// Represents a table style element
@interface excelTableStyleElement : excelBaseObject

@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (readonly) BOOL hasFormat;  // Returns or sets if table style element has format.
@property (copy, readonly) excelInterior *interiorObject;  // Returns a interior object that represents the interior for the specified table style element. 


@end

// Represents a table style
@interface excelTableStyle : excelBaseObject

- (SBElementArray *) tableStyleElements;

@property (readonly) BOOL builtIn;  // if this is a built in table style or not
- (NSString *) default;
@property (copy, readonly) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) NSString *namelocal;  // Returns or sets the local name of the object.
@property BOOL showAsAvailablePivotTableStyle;  // whether to show as avaliable pivot table style
@property BOOL showAsAvailableTableStyle;  // whether to show as avaliable table style


@end

@interface excelTextbox : excelBaseObject

- (SBElementArray *) characters;

@property BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property BOOL autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object. 
@property (copy, readonly) excelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property double height;  // Returns or set the height of the object.
@property excelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property excelE126 orientation;  // May also be a number value from -90 to 90 degrees.
@property excelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL roundedCorners;
@property BOOL shadow;
@property (copy) NSString *stringValue;  // Returns or sets the text of the specified object.
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) excelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property excelE284 verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property BOOL wrapAutoText;  // Returns or sets if the auto text is wrapped.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

@interface excelTop10FormatCondition : excelBaseObject

@property (copy, readonly) excelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property excelE315 calcFor;
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) excelE195 formatConditionType;  // Return the conditional format type.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property excelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.
@property BOOL top10Percentage;
@property NSInteger top10Rank;
@property excelE314 topOrBottom;


@end

// Represents the hierarchical member-selection control of a cube field.
@interface excelTreeviewControl : excelBaseObject

@property (copy) id drilled;
@property (copy) id hidden;


@end

@interface excelUniqueValuesFormatCondition : excelBaseObject

@property (copy, readonly) excelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property excelE317 duplicateOrUnique;
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) excelE195 formatConditionType;  // Return the conditional format type.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property excelE307 pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

// Represents data validation for a worksheet range.
@interface excelValidation : excelBaseObject

@property excelE199 IMEMode;  // Returns or sets the description of the Japanese input rules.
@property (readonly) excelE200 alertStyle;  // Returns the validation alert style.
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
@property (readonly) excelE196 validationOperator;  // Returns the operator for the conditional format or data validation.
@property (readonly) excelE198 validationType;  // Returns the data validation type.
@property (readonly) BOOL value;  // Returns true if all the validation criteria are met, that is, if the range contains valid data.

- (void) addDataValidationType:(excelE198)type alertStyle:(excelE200)alertStyle operator:(excelE196)operator_ formula1:(NSString *)formula1 formula2:(NSString *)formula2;  // Adds data validation to the specified range.
- (void) modifyType:(excelE198)type alertStyle:(excelE200)alertStyle operator:(excelE196)operator_ formula1:(NSString *)formula1 formula2:(NSString *)formula2;  // Modifies data validation for a range.

@end

@interface excelValueChange : excelBaseObject

@property (readonly) excelE905 allocationMethod;
@property (readonly) excelE904 allocationValue;
@property (copy, readonly) NSString *allocationWeightExpression;
@property (readonly) NSInteger order;
@property (copy, readonly) excelPivotCell *pivotCell;
@property (copy, readonly) NSString *tuple;
@property (readonly) double value;
@property (readonly) BOOL visibleInPivotTable;


@end

// Represents a vertical page break. 
@interface excelVerticalPageBreak : excelBaseObject

@property (readonly) excelE190 extent;  // Returns the type of the specified page break: full-screen or only within a print area.
@property (copy) excelRange *location;  // Returns or sets where this vertical page break is located.
@property excelE168 verticalPageBreakType;  // Returns or sets the type of vertical page break.


@end

// Contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.
@interface excelWebOptions : excelBaseObject

@property BOOL allowPng;  // Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save documents as a Web page.
@property excelMtEn encoding;  // Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document.
@property (copy) NSString *locationOfComponents;  // Returns or sets the central URL, on the intranet or Web, or path, local or network, to the location from which authorized users can download Microsoft Office Web components when viewing your saved document.
@property NSInteger pixelsPerInch;  // Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120.
@property (readonly) excelMSsz screenSize;  // Returns the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser.
@property (copy) NSString *webPageKeywords;  // Returns or sets the keywords for the Web page.
@property (copy) NSString *webPageTitle;  // Returns or sets the title for the web page.

- (void) useDefaultFolderSuffix;  // Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.

@end

// Represents a window. Many worksheet characteristics, such as scroll bars and gridlines, are actually properties of the window.
@interface excelWindow : excelBasicWindow

- (SBElementArray *) panes;

@property (copy, readonly) excelCell *activeCell;  // Returns a range object that represents the active cell in the active window, the window on top,  or in the specified window. If the window isn't displaying a worksheet, this property fails.
@property (copy, readonly) excelChart *activeChart;  // Returns a chart object that represents the active chart, either an embedded chart or a chart sheet. An embedded chart is considered active when it's either selected or activated.
@property (copy, readonly) excelPane *activePane;  // Returns a pane object that represents the active pane in the active window, the window on top,  or in the specified window. If the window isn't displaying a worksheet, this property fails.
@property (copy, readonly) excelSheet *activeSheet;  // Returns an object that represents the active sheet, the sheet on top, in the active workbook or in the specified window or workbook.
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
@property excelE103 gridlineColorIndex;  // Returns or sets the gridline color as an index into the current color palette.
@property double height;  // Returns or sets the height of the window. Use the usable height property to determine the maximum size for the window. You cannot set this property if the window is maximized or minimized. Use the window state property to determine the window state.
@property double leftPosition;  // Returns or sets the distance from the left edge of the client area to the left edge of the window.
@property (copy, readonly) excelRange *rangeSelection;  // Returns a range object that represents the selected cells on the worksheet in the specified window even if a graphic object is active or selected on the worksheet.
@property NSInteger scrollColumn;  // Returns or sets the number of the leftmost column in the window. 
@property NSInteger scrollRow;  // Returns or sets the number of the row that appears at the top of the window.
@property (copy, readonly) NSArray *selectedSheets;  // Returns a list of sheet objects that represents all the selected sheets in the specified window.
@property (copy, readonly) excelRange *selection;  // Returns the selected range object in the specified window.
@property BOOL split;  // Returns or sets if the window is split. 
@property NSInteger splitColumn;  // Returns or sets the column number where the window is split into panes, the number of columns to the left of the split line.
@property double splitHorizontal;  // Returns or sets the location of the horizontal window split, in points.
@property NSInteger splitRow;  // Returns or sets the row number where the window is split into panes, the number of rows above the split.
@property double splitVertical;  // Returns or sets the location of the vertical window split, in points.
@property double tabRatio;  // Returns or sets the ratio of the width of the workbook's tab area to the width of the window's horizontal scroll bar, as a number between 0 and 1, the default value is 0.75. 
@property double top;  //  The distance from the top edge of the window to the top edge of the usable area, below the menus, any toolbars docked at the top, and the formula bar. You cannot set this property for a maximized window.
@property (readonly) double usableHeight;  // Returns the maximum height of the space that a window can occupy in points.
@property (readonly) double usableWidth;  // Returns the maximum width of the space that a window can occupy in the application window area, in points.
@property excelE239 view;  // Returns or sets the view showing in the window.
@property BOOL visible;  // Returns or sets if the window is visible.
@property (copy, readonly) excelRange *visibleRange;  // Returns a range object that represents the range of cells that are visible in the window or pane. If a column or row is partially visible, it's included in the range. 
@property double width;  // Returns or sets the width of the window.
@property (readonly) NSInteger windowNumber;  // Returns the window number. For example, a window named Book1.xls:2 has 2 as its window number. Most windows have the window number 1.
@property excelE111 windowState;  // Returns or sets the state of the window.
@property (readonly) excelE154 windowType;  // Returns the window type.
@property excelE278 zoom;  // Returns or sets the display size of the window, as a percentage,100 equals normal size, 200 equals double size, and so on. 

- (void) activateNext;  // Activates the current window, sends it to the back of the window z-order, then activates the next window according to the z-order.
- (void) activatePrevious;  // Activates the specified window and then activates the window at the back of the window z-order.
- (void) scrollWorkbookTabsSheets:(NSInteger)sheets position:(excelE271)position;  // Scrolls through the workbook tabs at the bottom of the window. Doesn't affect the active sheet in the workbook.

@end

// A connection is a set of information needed to obtain data from an external data source other than an 1st_Excel12 workbook.
@interface excelWorkbookConnection : excelBaseObject

@property (copy, readonly) NSString *_default;
@property (copy, readonly) NSString *objectDescription;  // Returns or sets a brief description for a WorkbookConnection object.
@property (copy, readonly) NSString *name;  // Returns or sets the name of the workbook connection object.
@property (copy, readonly) id ranges;  // Returns the range of object for the specified WorkbookConnection object.
@property (readonly) excelE917 type;


@end

// Represents a Microsoft Excel workbook.
@interface excelWorkbook : excelBaseObject

- (SBElementArray *) documentProperties;
- (SBElementArray *) chartSheets;
- (SBElementArray *) commandBars;
- (SBElementArray *) customDocumentProperties;
- (SBElementArray *) namedItems;
- (SBElementArray *) pivotCaches;
- (SBElementArray *) sheets;
- (SBElementArray *) styles;
- (SBElementArray *) customViews;
- (SBElementArray *) windows;
- (SBElementArray *) worksheets;
- (SBElementArray *) internationalMacroSheets;
- (SBElementArray *) macroSheets;
- (SBElementArray *) tableStyles;

@property BOOL acceptLabelsInFormulas;  // Returns or sets if labels can be used in worksheet formulas.
@property NSInteger accuracyVersion;  // Returns or sets the accuracy version for this workbook.
@property (copy, readonly) excelChart *activeChart;  // Returns a chart object that represents the active chart, either an embedded chart or a chart sheet. An embedded chart is considered active when it's either selected or activated.
@property (copy, readonly) excelSheet *activeSheet;  // Returns an object that represents the active sheet, the sheet on top, in the specified window.
@property NSInteger autoUpdateFrequency;  // Returns or sets the number of minutes between automatic updates to the shared workbook. If this property is set to zero  updates occur only when the workbook is saved.
@property BOOL autoUpdateSaveChanges;  // Returns or sets if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. if false changes aren't posted, this workbook is still synchronized with changes made by other users.
@property (readonly) NSInteger calculationVersion;  // Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits, on the left, are the major version of Microsoft Excel.
@property NSInteger changeHistoryDuration;  // Returns or sets the number of days shown in the shared workbook's change history.
@property excelE222 conflictResolution;  // Returns or sets the way conflicts are to be resolved whenever a shared workbook is updated.
@property (readonly) BOOL createBackup;  // Returns true if a backup file is created when this file is saved.
@property BOOL date1904;  // Returns or sets if the workbook uses the 1904 date system.
@property (copy) id defaultPivottableStyle;  // Set or get the default pivot table style for the current workbook
@property (copy) id defaultTableStyle;  // Set or get the default table style for the current workbook
@property excelE242 displayDrawingObjects;  // Returns or sets how shapes are displayed.
@property BOOL enableAutoRecover;  // Gets or sets a value that indicates whether Microsoft Office Excel saves changed files, of all formats, on a timed interval.
@property (readonly) BOOL excel8CompatibilityMode;  // Gets a value that indicates whether the workbook is in compatibility mode.
@property (readonly) excelE188 fileFormat;  // Returns the file format of the workbook.
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
@property (copy, readonly) excelOfficeTheme *theme;
@property BOOL updateRemoteReferences;  // Returns or sets if Microsoft Excel updates remote references in for the workbook.
@property (copy, readonly) NSArray *userStatus;  // Returns a list of lists that provides information about each user who has the workbook open as a shared list. The list is name of user, date and time user opened the workbook, and a number 1 for exclusive, 2 for shared access.
@property (copy, readonly) excelWebOptions *webOptions;  // Returns the web options object, which contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.
@property (copy) NSString *workbookComments;  // Returns or sets the comment string for this workbook.
@property (copy) NSString *writePassword;  // Returns or sets a string for the write password of a workbook.
@property (readonly) BOOL writeReserved;  // Return true if the workbook is write-reserved
@property (copy, readonly) NSString *writeReservedBy;  // Returns the name of the user who currently has write permission for the workbook.

- (void) acceptAllChangesWhen:(NSString *)when who:(NSString *)who where:(NSString *)where;  // Accepts all changes in the specified shared workbook.
- (void) applyThemeFileName:(NSString *)fileName;  // Applies a theme or design template to the specified workbook
- (void) breakLinkName:(NSString *)name type:(excelE165)type;  // Converts formulas linked to other Microsoft Excel sources to values.
- (BOOL) canCheckIn;  // Returns True if Excel can check in a specified workbook to a server.
- (void) changeFileAccessMode:(excelE175)mode writePassword:(NSString *)writePassword notify:(BOOL)notify;  // Changes the access permissions for the workbook. This may require an updated version to be loaded from the disk.
- (void) changeLinkName:(NSString *)name newName:(NSString *)newName type:(excelE165)type;  // Changes a link from one document to another.
- (void) checkInSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic;  // Returns a workbook from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) checkInWithVersionSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic versionType:(excelCivt)versionType;  // Returns a workbook from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) deleteNumberFormatNumberFormat:(NSString *)numberFormat;  // Deletes a custom number format from the workbook.
- (BOOL) exclusiveAccess;  // Assigns the current user exclusive access to the workbook that's open as a shared list.
- (void) followHyperlinkAddress:(NSString *)address subAddress:(NSString *)subAddress newWindow:(BOOL)newWindow extraInfo:(NSString *)extraInfo method:(excelMANT)method headerInfo:(NSString *)headerInfo;  // Displays a cached document, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.
- (void) highlightChangesOptionsWhen:(excelE193)when who:(NSString *)who where:(excelE267)where;  // Controls how changes are shown in a shared workbook.
- (SBObject *) linkInfoName:(NSString *)name linkInfo:(excelE171)linkInfo type:(excelE179)type;  // Returns the link date and update status.
- (NSArray *) linkSourcesType:(excelE165)type;  // Returns a list of links in the workbook. The names in the array are the names of the linked documents. Returns empty if there are no links.
- (void) mergeWorkbookFileName:(NSString *)fileName;  // Merges changes from one workbook into an open workbook.
- (excelWindow *) newWindowOnWorkbook;  // Creates a new window or a copy of the specified workbook window.
- (void) openLinksName:(NSString *)name readOnly:(BOOL)readOnly type:(excelE165)type;  // Opens the supporting documents for a link or links.
- (void) protectSharingFileName:(NSString *)fileName password:(NSString *)password writeReservationPassword:(NSString *)writeReservationPassword readOnlyRecommended:(BOOL)readOnlyRecommended createBackup:(BOOL)createBackup sharingPassword:(NSString *)sharingPassword fileFormat:(excelE188)fileFormat;  // Saves the workbook and protects it for sharing.
- (void) protectWorkbookPassword:(NSString *)password structure:(BOOL)structure windows:(BOOL)windows;  // Protect workbook structure from changes.
- (void) purgeChangeHistoryNowDays:(NSInteger)days sharingPassword:(NSString *)sharingPassword;  // Removes entries from the change log for the specified workbook.
- (void) refreshAll;  // Refreshes all external data ranges and PivotTables in the workbook.
- (void) rejectAllChangesWhen:(NSString *)when who:(NSString *)who where:(NSString *)where;  // Rejects all changes in the specified shared workbook.
- (void) removeUserEntry_index:(NSInteger)entry_index;  // Disconnects the specified user from the shared workbook.
- (void) resetColors;  // Resets the color palette to the default colors.
- (void) saveWorkbookAsFilename:(NSString *)filename fileFormat:(excelE188)fileFormat password:(NSString *)password writeReservationPassword:(NSString *)writeReservationPassword readOnlyRecommended:(BOOL)readOnlyRecommended createBackup:(BOOL)createBackup accessMode:(excelE221)accessMode conflictResolution:(excelE222)conflictResolution addToMostRecentlyUsedList:(BOOL)addToMostRecentlyUsedList overwrite:(BOOL)overwrite;  // Saves changes into a different file.
- (void) sendHtmlMail;  // Opens a message window in Microsoft Outlook for sending the specified document, formatted as html.
- (void) sendMail;  // Opens a message window in your registered mail program for sending the specified document as an attachment.
- (void) toggleFormsDesign;  // Toggles Microsoft Office Excel into and out of design mode.
- (void) unprotectSharingSharingPassword:(NSString *)sharingPassword;  // Turns off protection for sharing and saves the workbook.
- (void) updateFromFile;  // Updates a read-only workbook from the saved disk version of the workbook if the disk version is more recent than the copy of the workbook that is loaded in memory.
- (void) updateLinkName:(NSString *)name type:(excelE165)type;  // Updates a Microsoft Excel  link.
- (void) webPagePreview;  // Displays a preview of the specified workbook as it would look if saved as a Web page.

@end

@interface excelDocument : excelWorkbook


@end

@interface excelWorksheet : excelSheet

@property excelE197 enableSelection;  // Returns or sets what can be selected on the sheet.
@property (readonly) BOOL protectScenarios;  // Returns true if the worksheet scenarios are protected.

- (void) createSummaryForScenariosReportType:(excelE234)reportType resultCells:(excelE267)resultCells;  // Creates a new worksheet that contains a summary report for the scenarios on the specified worksheet.
- (void) mergeScenariosMergeSource:(excelE298)mergeSource;  // Merges the scenarios from the merge source worksheet into this worksheet

@end

// Contains global application-level spelling options
@interface excelXlspellingOptions : excelBaseObject

@property excelE300 dictionaryLang;  // Returns or sets the LCID used by the proofing tools


@end



/*
 * Drawing Suite
 */

@interface excelAdjustment : excelBaseObject

@property double adjustment_value;


@end

// Represents an arc graphic.
@interface excelArc : excelBaseObject

- (SBElementArray *) characters;

@property NSInteger addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property NSInteger autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object. 
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
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property NSInteger stringValue;  // Returns or sets the text of the specified object.
@property NSInteger top;  // Returns the top position of the specified object, in points.
@property (readonly) NSInteger topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property NSInteger verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property NSInteger visible;  // Returns or sets if the object is visible.
@property NSInteger width;  // Returns or sets  the width of the object.
@property BOOL wrapAutoText;  // Returns or sets if the auto text is wrapped.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

@interface excelBulletFormat : excelBaseObject

@property (copy) NSString *bulletCharacter;  // Returns or sets the Unicode character value that is used for bullets in the specified text.
@property (copy, readonly) excelShapeFont *bulletFont;  // Returns a font object that represents character formatting for a bullet format object.
@property (readonly) NSInteger bulletNumber;  // Returns the bullet number of a paragraph.
@property NSInteger bulletStartValue;  // Gets or sets the beginning value of a bulleted list.
@property excelPBtS bulletStyle;  // Returns or sets a constant that represents the style of a bullet.
@property excelPBtT bulletType;  // Returns or sets a constant that represents the type of bullet.
@property double relativeSize;  // Returns or sets the bullet size relative to the size of the first text character in the paragraph.
@property BOOL useTextColor;  // Determines whether the specified bullets are set to the color of the first text character in the paragraph.
@property BOOL useTextFont;  // Determines whether the specified bullets are set to the font of the first text character in the paragraph.
@property BOOL visible;  // Returns or sets a value that specifies whether the bullet is visible.

- (void) setBulletPictureFileName:(NSString *)FileName;  // Sets the graphics file to be used for bullets in a bulleted list.

@end

// Contains properties and methods that apply to line callouts.
@interface excelCalloutFormat : excelBaseObject

@property BOOL accent;  // Returns or sets if a vertical accent bar separates the callout text from the callout line.
@property excelMCAt angle;  // Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box.
@property BOOL autoAttach;  // Returns or sets if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line, where the callout points to, is to the left or right of the callout text box.
@property (readonly) BOOL autoLength;  // Returns if the length of the callout line is automatically set. Use the automatic length method to set this property to true, and use the custom length method to set this property to false.
@property BOOL border;  // Returns or sets whether the text in the specified callout is surrounded by a border.
@property (readonly) double calloutFormatLength;  // When the auto length property of the specified callout is set to false, the length property returns the length in points of the first segment of the callout line, the segment attached to the text callout box.
@property excelMCot calloutFormatType;  // Returns or sets the callout type.
@property (readonly) double drop;  // For callouts with an explicitly set drop value, this property returns the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.
@property (readonly) excelMCDt dropType;  // Returns a value that indicates where the callout line attaches to the callout text box.
@property double gap;  // Returns or sets the horizontal distance in points between the end of the callout line and the text bounding box.

- (void) automaticLength;  // Specifies that the first segment of the callout line, the segment attached to the text callout box, be scaled automatically when the callout is moved.

@end

// Contains properties and methods that apply to connectors. A connector is a line that attaches two other shapes at points called connection sites.
@interface excelConnectorFormat : excelBaseObject

@property (readonly) BOOL beginConnected;  // Returns true if the beginning of the specified connector is connected to a shape.
@property (copy, readonly) excelShape *beginConnectedShape;  // Returns a shape object that represents the shape that the beginning of the specified connector is attached to.
@property (readonly) NSInteger beginConnectionSite;  // Returns an integer that specifies the connection site that the beginning of a connector is connected to.
@property excelMCtT connectorFormatType;  // Returns or sets the connector type.
@property (readonly) BOOL endConnected;  // Returns true if the end of the specified connector is connected to a shape.
@property (copy, readonly) excelShape *endConnectedShape;  // Returns a shape object that represents the shape that the end of the specified connector is attached to.
@property (readonly) NSInteger endConnectionSite;  // Returns an integer that specifies the connection site that the end of a connector is connected to.

- (void) beginConnectConnectedShape:(excelShape *)connectedShape connectionSite:(NSInteger)connectionSite;  // Attaches the beginning of the specified connector to a specified shape. If there's already a connection between the beginning of the connector and another shape, that connection is broken.
- (void) beginDisconnect;  // Detaches the beginning of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the beginning of the connector remains positioned at a connection site but is no longer connected.
- (void) endConnectConnectedShape:(excelShape *)connectedShape connectionSite:(NSInteger)connectionSite;  // Attaches the end of the specified connector to a specified shape. If there's already a connection between the end of the connector and another shape, that connection is broken.
- (void) endDisconnect;  // Detaches the end of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the end of the connector remains positioned at a connection site but is no longer connected.

@end

// Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.
@interface excelFillFormat : excelBaseObject

- (SBElementArray *) gradientStops;

@property (copy) NSColor *backColor;  // Returns or sets a RGB color that represents the background color for the specified fill or patterned line.
@property excelMCoI backColorThemeIndex;  // Returns or sets the specified fill background color.
@property (readonly) excelMFdT fillFormatType;  // Returns the shape fill format type.
@property (copy) NSColor *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property excelMCoI foreColorThemeIndex;  // Returns or sets the specified foreground fill or solid color.
@property (readonly) excelMGCt gradientColorType;  // Returns the gradient color type for the specified fill.
@property (readonly) double gradientDegree;  // Returns a value that indicates how dark or light a one-color gradient fill is. A value of zero means that black is mixed in with the shape's foreground color to form the gradient; a value of 1 means that white is mixed in. Values between 1 and zero blend.
@property (readonly) excelMGdS gradientStyle;  // Returns the gradient style for the specified fill.
@property (readonly) NSInteger gradientVariant;  // Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. If the gradient style is from center gradient, this property returns either 1 or 2.
@property (readonly) excelPpTy pattern;  // Returns a value that represents the pattern applied to the specified fill or line.
@property (readonly) excelMPGb presetGradientType;  // Returns the preset gradient type for the specified fill.
@property (readonly) excelMPzT presetTexture;  // Returns the preset texture for the specified fill.
@property BOOL rotateWithObject;  // Returns or sets whether the fill rotates with the specified shape.
@property excelMTtA textureAlignment;  // Returns or sets the texture alignment for the specified object.
@property double textureHorizontalScale;  // Returns or sets the texture alignment for the specified object.
@property (copy, readonly) NSString *textureName;  // Returns the name of the custom texture file for the specified fill.
@property double textureOffsetX;  // Returns or sets the texture alignment for the specified object.
@property double textureOffsetY;  // Returns or sets the texture alignment for the specified object.
@property BOOL textureTile;  // Returns the texture tile style for the specified fill.
@property (readonly) excelMxtT textureType;  // Returns the texture type for the specified fill.
@property double textureVerticalScale;  // Returns or sets the texture alignment for the specified object.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque, and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.

- (void) deleteGradientStopStopIndex:(NSInteger)stopIndex;  // Removes a gradient stop.
- (void) insertGradientStopCustomColor:(NSColor *)customColor position:(double)position transparency:(double)transparency stopIndex:(NSInteger)stopIndex;  // Adds a stop to a gradient.
- (void) oneColorGradientGradientStyle:(excelMGdS)gradientStyle variant:(NSInteger)variant degree:(double)degree;  // Sets the specified fill to a one-color gradient.
- (void) patternedPattern:(excelPpTy)pattern;  // Sets the specified fill to a pattern.
- (void) presetGradientGradientStyle:(excelMGdS)gradientStyle variant:(NSInteger)variant presetGradientType:(excelMPGb)presetGradientType;  // Sets the specified fill to a preset gradient.
- (void) presetTexturedPresetTexture:(excelMPzT)presetTexture;  // Sets the specified fill to a preset texture.
- (void) solid;  // Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.
- (void) twoColorGradientGradientStyle:(excelMGdS)gradientStyle variant:(NSInteger)variant;  // Sets the specified fill to a two-color gradient.
- (void) userPicturePictureFile:(NSString *)pictureFile;  // Fills the specified shape with one large image.
- (void) userTexturedTextureFile:(NSString *)textureFile;  // Fills the specified shape with small tiles of an image.

@end

// Represents the glow formatting for a shape or range of shapes
@interface excelGlowFormat : excelBaseObject

@property (copy) NSColor *color;  // Returns or sets the color for the specified glow format.
@property excelMCoI colorThemeIndex;  // Returns or sets the color for the specified glow format.
@property double radius;  // Returns or sets the length of the radius for the specified glow format.


@end

// Represents one gradient stop.
@interface excelGradientStop : excelBaseObject

@property (copy) NSColor *color;  // Returns or sets the color for the specified the gradient stop.
@property excelMCoI colorThemeIndex;  // Returns or sets the color for the specified gradient stop.
@property double position;  // Returns or sets the position for the specified gradient stop expressed as a percent.
@property double transparency;  // Returns or sets a value representing the transparency of the gradient fill expressed as a percent.


@end

// Represents line and arrowhead formatting. For a line, the line format object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.
@interface excelLineFormat : excelBaseObject

@property (copy) NSColor *backColor;  // Returns or sets a RGB color that represents the background color for the specified fill or patterned line.
@property excelMCoI backColorThemeIndex;  // Returns or sets the background color for a patterned line.
@property excelMAhL beginArrowheadLength;  // Returns or sets the length of the arrowhead at the beginning of the specified line.
@property excelMAhS beginArrowheadStyle;  // Returns or sets the style of the arrowhead at the beginning of the specified line.
@property excelMAhW beginArrowheadWidth;  // Returns or sets the width of the arrowhead at the beginning of the specified line.
@property excelMlDs dashStyle;  // Returns or sets the dash style for the specified line.
@property excelMAhL endArrowheadLength;  // Returns or sets the length of the arrowhead at the end of the specified line.
@property excelMAhS endArrowheadStyle;  // Returns or sets the style of the arrowhead at the end of the specified line.
@property excelMAhW endArrowheadWidth;  // Returns or sets the width of the arrowhead at the end of the specified line.
@property (copy) NSColor *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property excelMCoI foreColorThemeIndex;  // Returns or sets the foreground color for the line.
@property excelMLnS lineStyle;  // Returns or sets the line format style.
@property excelPpTy pattern;  // Returns or sets a value that represents the pattern applied to the specified fill or line.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.
@property double weight;  // Returns or sets the thickness of the specified line in points.


@end

// Represents a line graphic object.
@interface excelLine : excelBaseObject

@property excelE113 arrowheadLength;  // Returns or sets the length of an arrowhead
@property excelE119 arrowheadStyle;  // Returns or sets the style of an arrowhead.
@property excelE120 arrowheadWidth;  // Returns or sets the width of an arrowhead.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object. 
@property (copy, readonly) excelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or set the height of the object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property excelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) excelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents a Microsoft Office theme.
@interface excelOfficeTheme : excelBaseObject

@property (copy, readonly) excelThemeColorScheme *themeColorScheme;  // Returns the color scheme of a Microsoft Office theme.
@property (copy, readonly) excelThemeEffectScheme *themeEffectScheme;  // Returns the effects scheme of a Microsoft Office theme.
@property (copy, readonly) excelThemeFontScheme *themeFontScheme;  // Returns the font scheme of a Microsoft Office theme.


@end

// Represents an oval graphic.
@interface excelOval : excelBaseObject

- (SBElementArray *) characters;

@property NSInteger addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property NSInteger autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object. 
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
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
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

@interface excelParagraphFormat : excelBaseObject

- (SBElementArray *) tabStops;

@property excelPpgA alignment;  // Returns or sets a value specifying the alignment of the paragraph.
@property excelPBlA baselineAlignment;  // Returns or sets a constant that represents the vertical position of fonts in a paragraph.
@property (copy, readonly) excelBulletFormat *bullet;  // Returns a bullet format object for the paragraph.
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
@interface excelPictureFormat : excelBaseObject

@property double brightness;  // Returns or sets the brightness of the specified picture . The value for this property must be a number from 0.0, dimmest to 1.0, brightest.
@property excelMPc colorType;  // Returns or sets the type of color transformation applied to the specified picture.
@property double contrast;  // Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast to 1.0, the greatest contrast.
@property double cropBottom;  // Returns or sets the number of points that are cropped off the bottom of the specified picture.
@property double cropLeft;  // Returns or sets the number of points that are cropped off the left side of the specified picture.
@property double cropRight;  // Returns or sets the number of points that are cropped off the right side of the specified picture.
@property double cropTop;  // Returns or sets the number of points that are cropped off the top of the specified picture.
@property (copy) NSColor *transparencyColor;  // Returns or sets the transparent color for the specified picture as aRGB color. For this property to take effect, the transparent background property must be set to true.
@property BOOL transparentBackground;  // Returns or sets if the parts of the picture that are defined with a transparent color actually appear transparent.


@end

// Represents a rectangle graphic object.
@interface excelRectangle : excelBaseObject

- (SBElementArray *) characters;

@property NSInteger addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property NSInteger autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object. 
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
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
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
@interface excelReflectionFormat : excelBaseObject

@property excelMRfT reflectionType;  // Returns or sets the type of the reflection format object.


@end

@interface excelRulerLevel : excelBaseObject

@property double firstMargin;  // Returns or sets the first-line indent for the specified outline level, in points.
@property double leftMargin;  // Returns or sets the left indent for the specified outline level, in points.


@end

// Represents the ruler for the text in the specified shape or for all text in the specified text style. Contains tab stops and the indentation settings for text outline levels.
@interface excelRuler : excelBaseObject

- (SBElementArray *) tabStops;
- (SBElementArray *) rulerLevels;


@end

// Represents shadow formatting for a shape.
@interface excelShadowFormat : excelBaseObject

@property double blur;  // Returns or sets the blur, in points, of the specified shadow.
@property (copy) NSColor *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property excelMCoI foreColorThemeIndex;  // Returns or sets the foreground color for the shadow format.
@property BOOL obscured;  // Returns or sets if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. If false the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.
@property double offsetX;  // Returns or sets the horizontal offset in points of the shadow from the specified shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.
@property double offsetY;  // Returns or sets the vertical offset in points of the shadow from the specified shape. A positive value offsets the shadow below the shape; a negative value offsets it above the shape.
@property BOOL rotateWithShape;  // Returns or sets whether to rotate the shadow when rotating the shape.
@property excelMSSt shadowStyle;  // Returns or sets the style of shadow formatting to apply to a shape.
@property excelMSdT shadowType;  // Returns or sets the shape shadow type.
@property double size;  // Returns or sets the width of the shadow.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

// Contains font attributes such as font name, size, and color, for an object.
@interface excelShapeFont : excelBaseObject

@property (copy) NSString *ASCIIName;  // Returns or sets the font used for Latin text; characters with character codes from 0 through 127.
@property BOOL autoRotateNumbers;  // Returns or sets a value that specifies whether the numbers in a numbered list should be rotated when the text is rotated.
@property double baseLineOffset;  // Returns or sets a value specifying the horizontaol offset of the selected font.
@property BOOL bold;  // Returns or sets a value specifying whether the font should be bold.
@property excelMTCa capsType;  // Returns or sets a value specifying how the text should be capitalized.
@property (copy) NSString *eastAsianName;  // Returns or sets the font name used for Asian text.
@property (readonly) BOOL embedable;  // Returns a value indicating whether the font can be embedded in a page.
@property (readonly) BOOL embedded;  // Returns a value specifying whether the font is embedded in a page.
@property BOOL equalizeCharacterHeight;  // Returns or sets a value specifying whether the text should have the same horizontal height.
@property (copy, readonly) excelFillFormat *fillFormat;  // Returns a fill format object that contains fill formatting properties for the specified font.
@property (copy) NSColor *fontColor;
@property excelMCoI fontColorThemeIndex;  // Returns or sets the color for the specified font.
@property (copy) NSString *fontName;  // Returns or sets a value specifying the font to use for a selection.
@property (copy) NSString *fontNameOther;  // Returns or sets the font used for characters whose character set numbers are greater than 127.
@property double fontSize;  // Returns or sets the font size.
@property (copy, readonly) excelGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property (copy) NSColor *highlightColor;  // Returns or sets the text highlight color for object.
@property excelMCoI highlightColorThemeIndex;  // Returns or sets the highlight color for the specified text.
@property BOOL italic;  // Returns or sets a value specifying whether the text for a selection is italic.
@property double kerning;  // Returns or sets a value specifying the amount of spacing between text characters.
@property (copy, readonly) excelLineFormat *lineFormat;  // Returns a value specifiying the format of a line.
@property (copy, readonly) excelReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property (copy, readonly) excelShadowFormat *shadowFormat;  // Returns the value specifying the type of shadow effect for the selection of text.
@property excelMSET softEdgeType;  // Returns or sets the type soft edge format object.
@property double spacing;  // Returns or sets a value specifying the spacing between characters in a selection of text.
@property excelMTSt strikeType;  // Returns or sets a value specifying the strike format used for a selection of text.
@property BOOL subscript;  // Returns or sets a value specifying that the selected text should be displayed a subscript.
@property BOOL superscript;  // Returns or sets a value specifying that the selected text should be displayed a superscript.
@property (copy) NSColor *underlineColor;  // Returns a value specifying the color of the underline for the selected text.
@property excelMCoI underlineColorThemeIndex;  // Returns a value specifying the color of the underline for the selected text.
@property excelMTUl underlineStyle;  // Returns or sets a value specifying the text effect for the selected text.
@property excelMPXF wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.


@end

// Represents the shape text frame in a shape object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.
@interface excelShapeTextFrame : excelBaseObject

@property (readonly) BOOL hasText;  // Returns whether the specified text frame has text.
@property excelMHzA horizontalAnchor;
@property double marginBottom;  // Returns or sets the distance, in points, between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text.
@property double marginLeft;  // Returns or sets the distance, in points, between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text.
@property double marginRight;  // Returns or sets the distance, in points, between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text.
@property double marginTop;  // Returns or sets the distance, in points, between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text.
@property excelTxOr orientation;  // Returns or sets the text orientation.
@property excelMPFo pathFormat;  // Returns or sets the path type for the specified text frame.
@property (copy, readonly) excelRuler *ruler;
@property (copy, readonly) excelTextColumn *textColumn;  // Returns the text column object that represents the columns within the text frame.
@property (copy, readonly) excelTextRange *textRange;
@property (copy, readonly) excelThreeDFormat *threeDFormat;  // Returns the 3-D-effect formatting properties for the specified text.
@property excelMVtA verticalAnchor;
@property excelMWFo warpFormat;  // Returns or sets the warp type for the specified text frame.
@property (readonly) excelMPXF wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.
@property BOOL wordWrap;  // Returns or sets text break lines within or past the boundaries of the shape.
@property excelPAtS wordartAutoSize;  // The size of the specified object that changes automatically to fit text within its boundaries.
@property excelMPXF wordartFormat;  // Returns or sets the WordArt type for the specified text frame.

- (void) ungroup;

@end

// Represents an object in the drawing layer.
@interface excelShape : excelBaseObject

- (SBElementArray *) adjustments;

@property (copy) NSString *alternativeText;  // Returns or sets the descriptive alternative text string for a Shape object when the object is saved to a Web page.
@property excelMAsT autoShapeType;  // Returns or sets the shape type for the specified shape object, which must represent an auto-shape.
@property excelMBSI backgroundStyle;  // Returns or sets the background style.
@property excelMBW blackWhiteMode;  // Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode.
@property (copy, readonly) excelRange *bottomRightCell;  // Returns a range object that represents the cell that lies under the lower-right corner of the object.
@property (copy, readonly) excelChart *chart;  // Returns a chart object that represents the chart contained in the shape.
@property (readonly) BOOL child;  // True if the shape is a child shape.
@property (readonly) NSInteger connectionSiteCount;  // Returns the number of connection sites on the specified shape.
@property (readonly) BOOL connector;  // Returns true if the specified shape is a connector.
@property (copy, readonly) excelConnectorFormat *connectorFormat;  // Returns a connector format object that contains connector formatting properties if this shape is a connector.
@property (readonly) excelMCtT connectorType;  // Returns the connector type if this shape is a connector.
@property (copy, readonly) excelFillFormat *fillFormat;  // Returns a fill format object that contains fill formatting properties for the specified shape.
@property (copy, readonly) excelGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property (readonly) BOOL hasChart;  // True if the specified shape has a chart.
@property double height;  // Returns or sets the height of the object.
@property (readonly) BOOL horizontalFlip;  // Returns true if the specified shape is flipped around the horizontal axis.
@property (copy, readonly) excelHyperlink *hyperlink;  // Returns a hyperlink object that represents the hyperlink for the shape.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) excelLineFormat *lineFormat;  // Returns a line format object for this shape.
@property BOOL lockAspectRatio;  // Returns or sets if the specified shape retains its original proportions when you resize it. If false, you can change the height and width of the shape independently of one another when you resize it.
@property BOOL locked;  // Returns or sets if the object is locked. If false, the object can be modified when the sheet is protected.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) excelShape *parentgroup;  // Returns a Shape object that represents the common parent shape of a child shape or a range of child shapes.
@property excelE210 placement;  // Returns or sets how the object is placed on the worksheet.
@property (copy, readonly) excelReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property double rotation;  // Returns or sets the rotation of the shape, in degrees.
@property (copy, readonly) excelShadowFormat *shadowFormat;  // Returns a shadow format object for this shape.
@property (copy) NSString *shapeOnAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property excelMSSI shapeStyle;  // Returns or sets the shape style corresponding to the Quick Styles.
@property (copy, readonly) excelShapeTextFrame *shapeTextFrame;  // Returns a shape text frame object that contains the alignment and anchoring properties for the specified shape.
@property (readonly) excelMShp shapeType;  // Returns the shape type.
@property (copy, readonly) excelSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) excelTextFrame *textFrame;  // Returns a text frame object that contains the alignment and anchoring properties for the specified shape.
@property (copy, readonly) excelThreeDFormat *threeDFormat;  // Returns a threeD format object that contains 3-D-effect formatting properties for the specified shape.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property (copy, readonly) excelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object.
@property (readonly) BOOL verticalFlip;  // Returns true if the specified shape is flipped around the vertical axis.
@property BOOL visible;  // Returns or sets if the object is visible
@property double width;  // Returns or sets  the width of the object.
@property (copy, readonly) excelWordArtFormat *wordArtFormat;  // Returns the formatting properties for a word art effect.
@property (readonly) NSInteger zOrderPosition;  // Returns the position of the specified shape in the z-order. To set the shape's position in the z-order, use the z order method.

- (void) apply;  // Applies to the specified shape formatting that's been copied by using the pick up method.
- (void) flipFlipCmd:(excelMFlC)flipCmd;  // Flips the specified shape around its horizontal or vertical axis.
- (void) pickUp;  // Copies the formatting of the specified shape. Use the apply method to apply the copied formatting to another shape.
- (void) rerouteConnections;  // Reroutes connectors so that they take the shortest possible path between the shapes they connect. To do this, the reroute connections method may detach the ends of a connector and reattach them to different connecting sites on the connected shapes.
- (void) saveAsPicturePictureType:(excelMPiT)pictureType fileName:(NSString *)fileName;  // Saves the shape in the requested file using the stated graphic format
- (void) setShapesDefaultProperties;  // Applies the formatting for the specified shape to the default shape. Shapes created after this method has been used will have this formatting applied by default.
- (void) zOrderZOrderCommand:(excelMZoC)zOrderCommand;  // Moves the specified shape in front of or behind other shapes in the collection, that is, changes the shape's position in the z-order.

@end

@interface excelCallout : excelShape

@property (copy, readonly) excelCalloutFormat *calloutFormat;  // Returns a connector format object that contains connector formatting properties.
@property (readonly) excelMCot calloutType;  // Returns the type of callout.


@end

@interface excelPicture : excelShape

@property (copy, readonly) NSString *fileName;  // Returns he name of the file that has the picture.
@property (readonly) BOOL linkToFile;  // Returns if the picture is lined to the file.
@property (copy, readonly) excelPictureFormat *pictureFormat;  // Returns a picture format object for this picture.
@property (readonly) BOOL saveWithDocument;  // Returns if the picture should be saved with the document.

- (void) scaleHeightFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(excelMSFr)scale;  // Scales the height of the shape by a specified factor.
- (void) scaleWidthFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(excelMSFr)scale;  // Scales the width of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original size or relative to the current size.

@end

@interface excelShapeConnector : excelShape

@property (copy, readonly) excelConnectorFormat *connectorFormat;  // Returns a connector format object that contains connector formatting properties.
@property (readonly) excelMCtT connectorType;  // Returns the connector type.


@end

// The line shape uses begin line X, begin line Y, end line X, and end line Y when created
@interface excelShapeLine : excelShape

@property double beginLineX;  // Returns or sets the beginning X position of the line.
@property double beginLineY;  // Returns or sets the beginning Y position of the line.
@property double endLineX;  // Returns or sets the ending X position of the line.
@property double endLineY;  // Returns or sets the ending Y position of the line.


@end

@interface excelShapeTextbox : excelShape

@property (readonly) excelTxOr textOrientation;  // Returns the text orientation of the object.


@end

// Represents the soft edge formatting for a shape or range of shapes
@interface excelSoftEdgeFormat : excelBaseObject

@property excelMSET softEdgeType;  // Returns or sets the type soft edge format object.


@end

// Represents a single tab stop.
@interface excelTabStop : excelBaseObject

@property double tabPosition;  // Returns or sets the position of a tab stop relative to the left margin.
@property excelPTSp tabStopType;  // Returns or sets the type of the tab stop object.


@end

// Represents a single text column.
@interface excelTextColumn : excelBaseObject

@property NSInteger columnNumber;  // Returns or sets the index of the text column object.
@property double spacing;  // Returns or sets the spacing between text columns in a text column object.
@property (copy) id textDirection;  // Returns or sets the direction of text in the text column object.


@end

// Represents the text frame in a shape object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.
@interface excelTextFrame : excelBaseObject

- (SBElementArray *) characters;

@property BOOL autoMargins;  // Returns or sets if Microsoft Excel automatically calculates text frame margins.
@property BOOL autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property excelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property excelHzOf horizontalOverflow;  // Returns or sets the horizontal overflow.
@property double marginBottom;  // Returns or sets the distance, in points, between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text.
@property double marginLeft;  // Returns or sets the distance, in points, between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text.
@property double marginRight;  // Returns or sets the distance, in points, between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text.
@property double marginTop;  // Returns or sets the distance, in points, between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text.
@property excelTxOr orientation;  // Returns or sets the text orientation.
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property excelE114 verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property excelVrOf verticalOverflow;  // Returns or sets the vertical overflow.
@property BOOL wrapAutoText;  // Returns or sets if the auto text is wrapped.


@end

// Represents the color scheme of an Office Theme
@interface excelThemeColorScheme : excelBaseObject

- (SBElementArray *) themeColors;

- (NSColor *) getCustomColorName:(NSString *)name;  // Returns the custom color for the specified Microsoft Office theme.
- (void) loadThemeColorSchemeFileName:(NSString *)fileName;  // Loads the color scheme of a Microsoft Office theme from a file
- (void) saveThemeColorSchemeFileName:(NSString *)fileName;  // Saves the color scheme of a Microsoft Office theme to a file.

@end

// Represents a color in the color scheme of a Microsoft Office 2007 theme.
@interface excelThemeColor : excelBaseObject

@property (copy) NSColor *RGB;  // Returns or sets a value of a color in the color scheme of a Microsoft Office theme.
@property (readonly) excelMCSI themeColorSchemeIndex;  // Returns the index value a color scheme of a Microsoft Office theme.


@end

// Represents the effect scheme of a Microsoft Office theme.
@interface excelThemeEffectScheme : excelBaseObject

- (void) loadThemeEffectSchemeFileName:(NSString *)fileName;  // Loads the effects scheme of a Microsoft Office theme from a file

@end

// Represents the font scheme of a Microsoft Office theme.
@interface excelThemeFontScheme : excelBaseObject

- (SBElementArray *) minorThemeFonts;
- (SBElementArray *) majorThemeFonts;

- (void) loadThemeFontSchemeFileName:(NSString *)fileName;  // Loads the font scheme of a Microsoft Office theme from a file.
- (void) saveThemeFontSchemeFileName:(NSString *)fileName;  // Saves the font scheme of a Microsoft Office theme to a file.

@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface excelThemeFont : excelBaseObject

@property (copy) NSString *name;  // Returns or sets a value specifying the font to use for a selection.


@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface excelMajorThemeFont : excelThemeFont


@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface excelMinorThemeFont : excelThemeFont


@end

// Represents a shape's three-dimensional formatting.
@interface excelThreeDFormat : excelBaseObject

@property double ZDistance;  // Returns or sets a Single that represents the distance from the center of an object or text.
@property double bevelBottomDepth;  // Returns or sets the depth/height of the bottom bevel.
@property double bevelBottomInset;  // Returns or sets the inset size/width for the bottom bevel.
@property excelMBlT bevelBottomType;  // Returns or sets the bevel type for the bottom bevel.
@property double bevelTopDepth;  // Returns or sets the depth/height of the top bevel.
@property double bevelTopInset;  // Returns or sets the inset size/width for the top bevel.
@property excelMBlT bevelTopType;  // Returns or sets the bevel type for the top bevel.
@property (copy, readonly) NSColor *contourColor;  // Returns or sets the color of the contour of an object or text.
@property excelMCoI contourColorThemeIndex;  // Returns or sets the color for the specified contour.
@property double contourWidth;  // Returns or sets the width of the contour of an object or text.
@property double depth;  // Returns or sets the depth of the shape's extrusion.
@property (copy, readonly) NSColor *extrusionColor;  // Returns or sets a RGB color that represents the color of the shape's extrusion.
@property excelMCoI extrusionColorThemeIndex;  // Returns or sets the color for the specified extrusion.
@property excelMExC extrusionColorType;  // Returns or sets a value that indicates what will determine the extrusion color.
@property double fieldOfView;  // Returns or sets the amount of perspective for an object or text.
@property double lightAngle;  // Returns or sets the angle of the lighting.
@property BOOL perspective;  // Returns or sets if the extrusion appears in perspective that is, if the walls of the extrusion narrow toward a vanishing point. If false, the extrusion is a parallel, or orthographic, projection that is, if the walls don't narrow toward a vanishing point.
@property (readonly) excelMPzC presetCamera;  // Returns a constant that represents the camera preset.
@property (readonly) excelMExD presetExtrusionDirection;  // Returns or sets the direction taken by the extrusion's sweep path leading away from the extruded shape, the front face of the extrusion.
@property excelMPLd presetLightingDirection;  // Returns or sets the position of the light source relative to the extrusion.
@property excelMLtT presetLightingRig;  // Returns a constant that represents the lighting preset.
@property excelMlSf presetLightingSoftness;  // Returns or sets the intensity of the extrusion lighting.
@property excelMPMt presetMaterial;  // Returns or sets the extrusion surface material.
@property (readonly) excelM3DF presetThreeDFormat;
@property BOOL projectText;  // Returns or sets whether text on a shape rotates with shape.
@property double rotationX;  // Returns or sets the rotation of the extruded shape around the x-axis in degrees. A positive value indicates upward rotation; a negative value indicates downward rotation.
@property double rotationY;  // Returns or sets the rotation of the extruded shape around the y-axis, in degrees. A positive value indicates rotation to the left; a negative value indicates rotation to the right.
@property double rotationZ;  // Returns or sets the rotation of the extruded shape around the z-axis, in degrees. A positive value indicates clockwise rotation; a negative value indicates anti-clockwise rotation.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.

- (void) resetRotation;  // Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.

@end

// Contains properties and methods that apply to WordArt objects.
@interface excelWordArtFormat : excelBaseObject

@property excelMTxA alignment;  // Returns or sets a constant that represents the alignment for the specified text effect.
@property BOOL bold;  // Returns if the text of the word art shape is formatted as bold.
@property (copy) NSString *fontName;  // Returns or sets the font name of the font used by this word art shape.
@property double fontSize;  // Returns or sets the font size of the font used by this word art shape.
@property BOOL italic;  // Returns if the text of the word art shape is formatted as italic.
@property BOOL kernedPairs;  // Returns or sets if character pairs in a WordArt object have been kerned. 
@property BOOL normalizedHeight;  // Returns or sets if all characters, both uppercase and lowercase, in the specified WordArt are the same height.
@property excelMPTs presetShape;  // Returns or sets the shape of the specified WordArt.
@property excelMPXF presetWordArtEffect;  // Returns or sets the style of the specified WordArt.
@property BOOL rotatedChars;  // Returns or sets if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. If false, characters in the specified WordArt retain their original orientation relative to the bounding shape.
@property double tracking;  // Returns or sets the ratio of the horizontal space allotted to each character in the specified WordArt in relation to the width of the character. Can be a value from zero through 5.
@property (copy) NSString *wordArtText;  // Returns or sets the text associated with this word art shape.

- (void) toggleVerticalText;

@end



/*
 * Text Suite
 */

// Represents characters in an object that contains text.
@interface excelCharacter : excelBaseObject

@property (copy) NSString *content;  // Returns or sets the text for the specified object.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *phoneticCharacters;  // Returns or sets the phonetic text in the specified characters object.

- (void) insertIntoString:(NSString *)string;  // Inserts a string preceding the selected characters.

@end

// Contains font attributes, such as font name, size, and color, for an object.
@interface excelFont : excelBaseObject

@property BOOL bold;  // Returns or sets if the font is formatted as bold.
@property (copy) NSColor *color;  // Returns or sets the color for the font.
@property excelE110 fontBackground;  // Returns or sets the text background type.
@property excelE103 fontColorIndex;  // Returns or sets the color index of the font.
@property NSInteger fontSize;  // Returns or sets the font size.
@property (copy) NSString *fontStyle;  // Returns or sets the font style.
@property BOOL italic;  // Returns or sets if the font is formatted as italic.
@property (copy) NSString *name;  // Returns or sets the font name associated with this font object.
@property BOOL outlineFont;  // Returns or sets if the specified font is formatted as outline.
@property BOOL shadow;  // Returns or sets if the specified font is formatted as shadowed.
@property BOOL strikethrough;  // Returns or sets if the font is formatted as strikethrough text.
@property BOOL subscript;  // Returns or sets if the font is formatted as subscript.
@property BOOL superscript;  // Returns or sets if the font is formatted as superscript.
@property excelE305 themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property excelE304 themeFontIndex;  // Returns or sets the theme font in the applied font scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.
@property excelE130 underline;  // Returns or sets the type of underline applied to the font.


@end

// Represents a style description for a range.
@interface excelStyle : excelBaseObject

@property BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically. 
@property (readonly) BOOL builtIn;  // Returns true if the style is a built-in style.
@property (copy, readonly) excelFont *fontObject;  // Returns the font object associated with this style.
@property BOOL formulaHidden;  // Returns or sets if the formula will be hidden when the worksheet is protected.
@property excelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property BOOL includeAlignment;  // Returns or sets if the style includes the add indent, horizontal alignment, vertical alignment, wrap text, and orientation properties.
@property BOOL includeBorder;  // Returns or sets if the style includes the color, color index, line style, and weight border properties.
@property BOOL includeFont;  // Returns or sets if the style includes the background, bold, color, color index, font style, italic, name, outline font, shadow, size, strikethrough, subscript, superscript, and underline font properties.
@property BOOL includeNumber;  // Returns or sets if the style includes the number format property. 
@property BOOL includePatterns;  // Returns or sets if the style includes the color, color index, invert if negative, pattern, pattern color, and pattern color index interior properties.
@property BOOL includeProtection;  // Returns or sets if the style includes the formula hidden and locked protection properties.
@property NSInteger indentLevel;  // Returns or sets the indent level for the style. Can be an integer from 0 to 15.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified style.
@property BOOL locked;  // Returns or sets if the range using this style is locked.
@property BOOL mergedCells;  // Returns true if the style contains merged cells.
@property (copy, readonly) NSString *name;  // Returns the name of the style.
@property (copy, readonly) NSString *nameLocal;  // Returns the name of the style, in the language of the user.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property (copy) NSString *numberFormatLocal;  // Returns or sets the format code for the object as a string in the language of the user.
@property excelE126 orientation;  // May also be a number value from -90 to 90 degrees.
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shrinkToFit;  // Returns or sets if text automatically shrinks to fit in the available column width.
@property (copy, readonly) NSString *value;  // Return the name of the specified style.
@property excelE114 verticalAlignment;  // Returns or sets the vertical alignment of the style.
@property BOOL wrapText;  // Returns or sets if Microsoft Excel wraps the text in the object.

- (excelBorder *) getBorderWhichBorder:(excelE243)whichBorder;  // Returns the specified border object.

@end

// Represents the text frame in a shape object.
@interface excelTextRange : excelBaseObject

- (SBElementArray *) textRangeCharacters;
- (SBElementArray *) words;
- (SBElementArray *) textRangeLines;
- (SBElementArray *) sentences;
- (SBElementArray *) paragraphs;
- (SBElementArray *) textFlows;

@property (readonly) double boundHeight;  // Returns the height, in points, of the text bounding box for the specified text.
@property (readonly) double boundLeft;  // Returns the left coordinate, in points, of the text bounding box for the specified text.
@property (readonly) double boundTop;  // Returns the top coordinate, in points, of the text bounding box for the specified text.
@property (readonly) double boundWidth;  // Returns the width, in points, of the text bounding box for the specified text.
@property (copy) NSString *content;  // Returns or sets the text in a text range.
@property (copy, readonly) excelShapeFont *font;  // Returns the character formatting for the object.
@property (copy, readonly) excelParagraphFormat *paragraphFormat;  // Returns the paragraph formatting for the specified text.
@property (readonly) NSInteger textLength;  // Returns the length of a text range.
@property (readonly) NSInteger textStart;  // Returns the starting point of the specified text range.

- (void) addPeriodsTo;  // Adds period punctuation to the end of the text contained in text range object.
- (void) changeCaseTo:(excelPCgC)to;  // Changes the case of a text range object.
- (void) copyTextRange;  // Copies a text range object.
- (void) cutTextRange;  // Removes a portion or all of the text from a range of text.
- (NSArray *) getRotatedTextBounds;  // Returns back a list containing the four point of the text bounds as follows  {x1, y1}, {x2, y2}, {x3, y3}, {x4, y4} }
- (void) insertTextTextRangeInsertWhere:(excelMTiP)insertWhere newText:(NSString *)newText;  // Adds new text to the text range object.
- (void) pasteTextRange;  // Pastes the contents of the Clipboard into the text range object.
- (void) removePeriodsFrom;  // Removes all period punctuation from the text in the text range object.

@end

@interface excelParagraph : excelTextRange


@end

@interface excelSentence : excelTextRange


@end

@interface excelTextFlow : excelTextRange


@end

@interface excelTextRangeCharacter : excelTextRange


@end

@interface excelTextRangeLine : excelTextRange


@end

@interface excelWord : excelTextRange


@end



/*
 * Table Suite
 */

// Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3-D range.
@interface excelRange : excelBaseObject

- (SBElementArray *) ranges;
- (SBElementArray *) cells;
- (SBElementArray *) rows;
- (SBElementArray *) columns;
- (SBElementArray *) characters;
- (SBElementArray *) formatConditions;
- (SBElementArray *) hyperlinks;

@property (copy, readonly) excelExcelComment *ExcelComment;  //  Returns a comment object that represents the comment associated with the cell in the upper-left corner of the range.
@property BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property (copy, readonly) NSArray *areas;  // Returns a list of  range objects  that represents all the ranges in a multiple-area selection.
@property double columnWidth;  // Returns or sets the width of all columns in the specified range.
@property (copy, readonly) id countLarge;  // Counts the largest value in a given range of values. Read-only Variant.
@property (copy, readonly) excelRange *currentArray;  // If the specified cell is part of an array, returns a range object that represents the entire array.
@property (copy, readonly) excelRange *currentRegion;  // Returns a range object that represents the current region. The current region is a range bounded by any combination of blank rows and blank columns.
@property (copy, readonly) excelRange *dependents;  // Returns a range object that represents the range containing all the dependents of a cell. This can be a multiple selection, a union of Range objects, if there's more than one dependent.
@property (copy, readonly) excelRange *directDependents;  // Returns a range object that represents the range containing all the direct dependents of a cell. This can be a multiple selection, a union of Range objects, if there's more than one dependent.
@property (copy, readonly) excelRange *directPrecedents;  // Returns a range object that represents the range containing all the direct precedents of a cell. This can be a multiple selection, a union of Range objects,  if there's more than one precedent.
@property (copy, readonly) excelDisplayFormat *displayFormat;  // Returns the display format associated with this range.
@property (copy, readonly) excelRange *entireColumn;  // Returns a range object that represents the entire column, or columns,  that contains the specified range.
@property (copy, readonly) excelRange *entireRow;  // Returns a range object that represents the entire row, or rows,  that contains the specified range.
@property (readonly) NSInteger firstColumnIndex;  // Returns the number of the first column in the first area in the specified range.
@property (readonly) NSInteger firstRowIndex;  // Returns the number of the first row of the first area in the range.
@property (copy, readonly) excelFont *fontObject;  // Returns the font object associated with this range.
@property (copy) id formula;  // Returns or sets the object's formula, in A1-style notation.
@property (copy) NSString *formulaArray;  // Returns or sets the array formula of a range.
@property BOOL formulaHidden;  // Returns or sets if the formula will be hidden when the workbook or worksheet is protected.
@property excelE192 formulaLabel;  // Returns or sets the formula label type for the specified range. 
@property (copy) NSString *formulaLocal;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formulaR1c1;  // Returns or sets the formula for the object, using R1C1-style notation.
@property (copy) NSString *formulaR1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property (readonly) BOOL hasArray;  // Returns true if the specified cell is part of an array formula.
@property (readonly) BOOL hasFormula;  // Returns true if all cells in the range contain formulas.
@property (readonly) double height;  // Returns the height of the range.
@property BOOL hidden;  // Returns or sets if the rows or columns are hidden. The specified range must span an entire column or row.
@property excelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the range.
@property NSInteger indentLevel;  // Returns or sets the indent level for the range or style. Can be an integer from 0 to 15.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (readonly) double leftPosition;  // The distance from the left edge of column A to the left edge of the range. If the range is discontinuous, the first area is used. If the range is more than one column wide, the leftmost column in the range is used.
@property (readonly) NSInteger listHeaderRows;  // Returns the number of header rows for the specified range.
@property (copy, readonly) excelListObject *listObject;  // Returns the list object associated with this range.
@property (readonly) excelE152 locationInTable;  // Returns an enumeration that describes the part of the pivot table that contains the upper-left corner of the specified range.
@property BOOL locked;  // Returns or sets  if the range is locked.
@property (copy, readonly) excelRange *mergeArea;  // Returns a range object that represents the merged range containing the specified cell. If the specified cell isn't in a merged range, this property returns the specified cell.
@property BOOL mergeCells;  // Returns or sets if the range contains merged cells. 
@property (copy) NSString *name;  // Returns or sets the name of the range.
@property (copy, readonly) NSString *namedItem;  // Returns the named item associated with this range.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property (copy) NSString *numberFormatLocal;  // Returns or sets the format code for the object as a string in the language of the user.
@property NSInteger outlineLevel;  // Returns or sets the current outline level of the specified row or column
@property excelE168 pageBreak;  // Returns or sets the location of a page break.
@property (copy, readonly) excelPhonetic *phoneticObject;  // Returns the phonetic object, which contains information about a specific phonetic text string in a specified range.
@property (copy, readonly) excelPivotField *pivotField;  // Returns a pivot field object that represents the pivot field containing the upper-left corner of the specified range.
@property (copy, readonly) excelPivotItem *pivotItem;  // Returns a pivot item object that represents the pivot item containing the upper-left corner of the specified range.
@property (copy, readonly) excelPivotTable *pivotTable;  // Returns a pivot table object that represents the pivot table containing the upper-left corner of the specified range.
@property (copy, readonly) excelRange *precedents;  // Returns a range object that represents all the precedents of a cell. This can be a multiple selection, a union of Range objects, if there's more than one precedent.
@property (copy, readonly) NSString *prefixCharacter;  // Returns the prefix character for the cell.
@property (copy, readonly) excelQueryTable *queryTable;  // Returns a query table object that represents the query table that intersects the specified range.
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property double rowHeight;  // Returns  or sets the height of all the rows in the range specified, measured in points.
@property BOOL showDetail;  // Returns or sets if the outline is expanded for the specified range, so that the detail of the column or row is visible. The specified range must be a single summary column or row in an outline.
@property BOOL shrinkToFit;  // Returns or sets if text automatically shrinks to fit in the available column width.
@property (copy, readonly) id stringValue;  // Returns the text for the range.
@property (copy) excelStyle *styleObject;  // Returns or sets a style object that represents the style of the specified range.
@property (readonly) BOOL summary;  // Returns true if the range is an outlining summary row or column. The range should be a row or a column.
@property excelE126 textOrientation;  // The text orientation. Can be a number value from -90 to 90 degrees.
@property (readonly) double top;  // The distance from the top edge of row 1 to the top edge of the range. If the range is discontinuous, the first area is used. If the range is more than one row high, the top, lowest numbered, row in the range is used.
@property BOOL useStandardHeight;  // Returns or sets if the row height of the range object equals the standard height of the sheet.
@property BOOL useStandardWidth;  // Returns or sets if the column width of the range object equals the standard width of the sheet.
@property (copy, readonly) excelValidation *validation;  // Returns the validation object that represents data validation for the specified range.
@property (copy) id value;  // Returns or sets the value of the range.
@property (copy) id value2;  // Returns or sets the value of the range. The difference between this property and Value is that Value2 uses the numerical representation of cells that are formatted as currency and date.
@property excelE284 verticalAlignment;  // Returns or sets the vertical alignment of the range.
@property (readonly) double width;  // Returns the width of the range.
@property (copy, readonly) excelSheet *worksheetObject;  // Returns a worksheet object that represents the worksheet containing the specified range.
@property BOOL wrapText;  // Returns or sets if Microsoft Excel wraps the text in the object.

- (void) activateObject;  // Activates the object.
- (excelExcelComment *) addCommentCommentText:(NSString *)commentText;  // Adds a comment to the range.
- (void) advancedFilterAction:(excelE163)action criteriaRange:(excelE292)criteriaRange copyToRange:(excelE292)copyToRange unique:(BOOL)unique;  // Filters or copies data from a list based on a criteria range. If the initial selection is a single cell, that cell's current region is used.
- (void) applyNamesNames:(NSArray *)names ignoreRelativeAbsolute:(BOOL)ignoreRelativeAbsolute useRowColumnNames:(BOOL)useRowColumnNames omitColumn:(BOOL)omitColumn omitRow:(BOOL)omitRow order:(excelE166)order appendLast:(BOOL)appendLast;  // Applies names to the cells in the specified range.
- (void) applyOutlineStyles;  // Applies outlining styles to the specified range.
- (void) autoOutline;  // Automatically creates an outline for the specified range. If the range is a single cell, Microsoft Excel creates an outline for the entire sheet. The new outline replaces any existing outline.
- (NSString *) autocompleteString:(NSString *)string;  // Returns an AutoComplete match from the list. If there's no AutoComplete match or if more than one entry in the list matches the string to complete, this method returns an empty string.
- (void) autofillDestination:(excelE292)destination type:(excelE185)type;  // Performs an autofill on the cells in the specified range.
- (void) autofilterRangeField:(NSInteger)field criteria1:(id)criteria1 operator:(excelE186)operator_ criteria2:(id)criteria2 visibleDropDown:(BOOL)visibleDropDown;  // Filters a list using the AutoFilter.
- (void) autofit;  // Changes the width of the specified column range or the height of the specified row range to achieve the best fit.
- (void) autoformatFormat:(excelE215)format includeNumber:(BOOL)includeNumber font:(BOOL)font alignment:(BOOL)alignment border:(BOOL)border pattern:(BOOL)pattern width:(BOOL)width;  // Automatically formats a range of cells, using a predefined format.
- (void) borderAroundLineStyle:(excelE133)lineStyle weight:(excelE128)weight colorIndex:(excelE103)colorIndex color:(NSColor *)color;  // Adds a border to a range and sets the color, line style, and weight properties for the new border.
- (void) calculate;  // Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet.
- (void) calculateRowMajorOrder;  // Calculates a specfied range of cells.
- (void) checkSpellingCustomDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase alwaysSuggest:(BOOL)alwaysSuggest;  // Checks the spelling of an object.
- (void) clearExcelComments;  // Clears all cell comments from the specified range.
- (void) clearContents;  // Clears all cell comments from the specified range.
- (void) clearOutline;  // Clears the outline for the specified range.
- (void) clearRange;  // Clears the entire object.
- (void) clearRangeFormats;  // Clears the formatting of the object.
- (excelRange *) columnDifferencesComparison:(excelE292)comparison;  // Returns a range object that represents all the cells whose contents are different from the comparison cell in each column.
- (void) consolidateSources:(NSArray *)sources consolidationFunction:(excelE150)consolidationFunction topRow:(BOOL)topRow leftColumn:(BOOL)leftColumn createLinks:(BOOL)createLinks;  // Consolidates data from multiple ranges on multiple worksheets into a single range on a single worksheet.
- (void) copyPictureAppearance:(excelE207)appearance format:(excelE156)format;  // Copies the selected object to the clipboard as a picture.
- (void) copyRangeDestination:(excelRange *)destination;  //  Copies the range to the specified range.
- (void) createNamesTop:(BOOL)top leftPosition:(BOOL)leftPosition bottom:(BOOL)bottom right:(BOOL)right;  // Creates names in the specified range, based on text labels in the sheet.
- (void) cutRangeDestinationOfCut:(excelE292)destinationOfCut;  // Cuts the range to the clipboard.
- (void) dataSeriesRowcol:(excelE105)rowcol dataSeriesType:(excelE107)dataSeriesType date:(excelE129)date increment:(NSInteger)increment stop:(NSInteger)stop trend:(BOOL)trend;  // Creates a data series in the specified range.
- (void) dataTableRowInput:(excelCell *)rowInput columnInput:(excelCell *)columnInput;  // Creates a data table based on input values and formulas that you define on a worksheet.
- (void) deleteRangeShift:(excelE148)shift;  // Deletes the range
- (void) fillDown;  // Fills down from the top cell or cells in the specified range to the bottom of the range. The contents and formatting of the cell or cells in the top row of a range are copied into the rest of the rows in the range.
- (void) fillLeft;  // Fills left from the rightmost cell or cells in the specified range. The contents and formatting of the cell or cells in the rightmost column of a range are copied into the rest of the columns in the range.
- (void) fillRight;  // Fills right from the leftmost cell or cells in the specified range. The contents and formatting of the cell or cells in the leftmost column of a range are copied into the rest of the columns in the range.
- (void) fillUp;  // Fills up from the bottom cell or cells in the specified range to the top of the range. The contents and formatting of the cell or cells in the bottom row of a range are copied into the rest of the rows in the range.
- (excelRange *) findWhat:(NSString *)what after:(excelE292)after lookIn:(excelE153)lookIn lookAt:(excelE177)lookAt searchOrder:(excelE224)searchOrder searchDirection:(excelE223)searchDirection matchCase:(BOOL)matchCase matchByte:(BOOL)matchByte;  // Finds specific information in a range, and returns a range object that represents the first cell where that information is found. Returns nothing if no match is found. Doesn't affect the selection or the active cell.
- (excelRange *) findNextAfter:(excelE292)after;  // Continues a search that was begun with the find method. Finds the next cell that matches those same conditions and returns a range object that represents that cell. Doesn't affect the selection or the active cell.
- (excelRange *) findPreviousAfter:(excelE292)after;  // Continues a search that was begun with the find method. Finds the previous cell that matches those same conditions and returns a range object that represents that cell. Doesn't affect the selection or the active cell.
- (void) functionWizard;  // Starts the Function Wizard for the upper-left cell of the range.
- (id) getXMLValue;  // Returns the value of the specified range as XML.
- (NSString *) getAddressRowAbsolute:(BOOL)rowAbsolute columnAbsolute:(BOOL)columnAbsolute referenceStyle:(excelE158)referenceStyle external:(BOOL)external relativeTo:(excelE292)relativeTo;  // Returns the range reference in the language of the macro.
- (NSString *) getAddressLocalRowAbsolute:(BOOL)rowAbsolute columnAbsolute:(BOOL)columnAbsolute referenceStyle:(excelE158)referenceStyle external:(BOOL)external relativeTo:(excelE292)relativeTo;  // Returns the range reference in the language of the user.
- (excelBorder *) getBorderWhichBorder:(excelE243)whichBorder;  // Returns the specified border object.
- (excelRange *) getEndDirection:(excelE149)direction;  // Returns a range object that represents the cell at the end of the region that contains the source range.
- (excelRange *) getOffsetRowOffset:(NSInteger)rowOffset columnOffset:(NSInteger)columnOffset;  // Returns a range object that represents a range that's offset from the specified range.
- (excelRange *) getResizeRowSize:(NSInteger)rowSize columnSize:(NSInteger)columnSize;  // Resizes the specified range. Returns a Range object that represents the resized range.
- (BOOL) goalSeekGoal:(double)goal changingCell:(excelRange *)changingCell;  // Calculates the values necessary to achieve a specific goal. If the goal is an amount returned by a formula, this calculates a value that, when supplied to your formula, causes the formula to return the number you want. Returns True if successful.
- (BOOL) groupStart:(NSInteger)start end:(NSInteger)end by:(NSInteger)by periods:(NSArray *)periods;  // When the Range object represents a single cell in a pivot table field's data range, the group method performs numeric or date-based grouping in that field.
- (void) insertIndentInsertAmount:(NSInteger)insertAmount;  // Adds an indent to the specified range.
- (void) insertIntoRangeShift:(excelE147)shift;  // Inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.
- (void) justify;  // Rearranges the text in a range so that it fills the range evenly.
- (void) listNames;  // Pastes a list of all non-hidden names onto the worksheet, beginning with the first cell in the range.
- (void) mergeAcross:(BOOL)across;  // Creates a merged cell from the specified Range object.
- (void) navigateArrowTowardPrecedent:(BOOL)towardPrecedent arrowNumber:(NSInteger)arrowNumber linkNumber:(NSInteger)linkNumber;  // Navigates a tracer arrow for the specified range to the precedent, dependent, or error-causing cell or cells. Selects the precedent, dependent, or error cells and returns a range object that represents the new selection.
- (void) parseParseLine:(NSString *)parseLine destination:(excelE292)destination;  // Parses a range of data and breaks it into multiple cells. Distributes the contents of the range to fill several adjacent columns; the range can be no more than one column wide.
- (void) pasteSpecialWhat:(excelE204)what operation:(excelE203)operation skipBlanks:(BOOL)skipBlanks transpose:(BOOL)transpose;  // Pastes the contents of the Clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.
- (void) printOutFrom:(NSInteger)from to:(NSInteger)to copies:(NSInteger)copies preview:(BOOL)preview activePrinter:(NSString *)activePrinter printToFile:(BOOL)printToFile collate:(BOOL)collate;  // Prints the object
- (void) printPreviewEnableChanges:(BOOL)enableChanges;  // Shows a preview of the object as it would look when printed. This function has been deprecated.
- (void) removeSubtotal;  // Removes subtotals from a list.
- (BOOL) replaceWhat:(NSString *)what replacement:(NSString *)replacement lookAt:(excelE177)lookAt searchOrder:(excelE224)searchOrder matchCase:(BOOL)matchCase matchByte:(BOOL)matchByte matchControlCharacters:(BOOL)matchControlCharacters matchDiacritics:(BOOL)matchDiacritics;  // Finds and replaces characters in cells within a range. Doesn't change the selection or the active cell.
- (excelRange *) rowDifferencesComparison:(excelCell *)comparison;  // Returns a range object that represents all the cells whose contents are different from those of the comparison cell in each row.
- (NSString *) runVBMacroArg1:(id)arg1 arg2:(id)arg2 arg3:(id)arg3 arg4:(id)arg4 arg5:(id)arg5 arg6:(id)arg6 arg7:(id)arg7 arg8:(id)arg8 arg9:(id)arg9 arg10:(id)arg10 arg11:(id)arg11 arg12:(id)arg12 arg13:(id)arg13 arg14:(id)arg14 arg15:(id)arg15 arg16:(id)arg16 arg17:(id)arg17 arg18:(id)arg18 arg19:(id)arg19 arg20:(id)arg20 arg21:(id)arg21 arg22:(id)arg22 arg23:(id)arg23 arg24:(id)arg24 arg25:(id)arg25 arg26:(id)arg26 arg27:(id)arg27 arg28:(id)arg28 arg29:(id)arg29 arg30:(id)arg30;  //  Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.
- (NSString *) runXLMMacroArg1:(id)arg1 arg2:(id)arg2 arg3:(id)arg3 arg4:(id)arg4 arg5:(id)arg5 arg6:(id)arg6 arg7:(id)arg7 arg8:(id)arg8 arg9:(id)arg9 arg10:(id)arg10 arg11:(id)arg11 arg12:(id)arg12 arg13:(id)arg13 arg14:(id)arg14 arg15:(id)arg15 arg16:(id)arg16 arg17:(id)arg17 arg18:(id)arg18 arg19:(id)arg19 arg20:(id)arg20 arg21:(id)arg21 arg22:(id)arg22 arg23:(id)arg23 arg24:(id)arg24 arg25:(id)arg25 arg26:(id)arg26 arg27:(id)arg27 arg28:(id)arg28 arg29:(id)arg29 arg30:(id)arg30;  //  Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.
- (void) setXMLValueRangeValue:(id)rangeValue;  // Set the value of the specified range using XML.
- (void) show;  // Scrolls through the contents of the active window to move the range into view. The range must consist of a single cell in the active document.
- (void) showDependentsRemove:(BOOL)remove;  // Draws tracer arrows to the direct dependents of the range.
- (void) showErrors;  // Draws tracer arrows through the precedents tree to the cell that's the source of the error, and returns the range that contains that cell.
- (void) showPrecedentsRemove:(BOOL)remove;  // Draws tracer arrows to the direct precedents of the range.
- (void) sortKey1:(excelRange *)key1 order1:(excelE228)order1 key2:(excelRange *)key2 sortType:(excelE230)sortType order2:(excelE228)order2 key3:(excelRange *)key3 order3:(excelE228)order3 header:(excelE241)header orderCustom:(NSInteger)orderCustom matchCase:(BOOL)matchCase orientation:(excelE229)orientation sortMethod:(excelE226)sortMethod ignoreControlCharacters:(BOOL)ignoreControlCharacters ignoreDiacritics:(BOOL)ignoreDiacritics dataoption1:(excelE302)dataoption1 dataoption2:(excelE302)dataoption2 dataoption3:(excelE302)dataoption3;  // Sorts a pivot table, a range, or the current region, if the specified range contains only one cell.
- (void) sortSpecialSortMethod:(excelE226)sortMethod key1:(excelRange *)key1 order1:(excelE228)order1 type:(id)type key2:(excelRange *)key2 order2:(excelE228)order2 key3:(excelRange *)key3 order3:(excelE228)order3 header:(excelE241)header orderCustom:(NSInteger)orderCustom matchCase:(BOOL)matchCase orientation:(excelE229)orientation dataoption1:(excelE302)dataoption1 dataoption2:(excelE302)dataoption2 dataoption3:(excelE302)dataoption3;  // Uses East Asian sorting methods to sort the range, or uses the current region if the range contains only one cell.
- (excelRange *) specialCellsType:(excelE182)type value:(excelE231)value;  // Returns a Range object that represents all the cells that match the specified type and value.
- (void) subtotalGroupBy:(NSInteger)groupBy function:(excelE150)function totalList:(NSArray *)totalList replace:(BOOL)replace pageBreaks:(BOOL)pageBreaks summaryBelowData:(excelE232)summaryBelowData;  // Creates subtotals for the range, or the current region, if the range is a single cell.
- (void) textToColumnsDestination:(excelE292)destination dataType:(excelE236)dataType textQualifier:(excelE237)textQualifier consecutiveDelimiter:(BOOL)consecutiveDelimiter tab:(BOOL)tab semicolon:(BOOL)semicolon comma:(BOOL)comma space:(BOOL)space useOther:(BOOL)useOther otherChar:(NSString *)otherChar fieldInfo:(NSArray *)fieldInfo decimalSeparator:(NSString *)decimalSeparator thousandsSeparator:(NSString *)thousandsSeparator;  // Parses a column of cells that contain text into several columns.
- (void) ungroup;  // Promotes a range in an outline, that is, decreases its outline level. The specified range must be a row or column, or a range of rows or columns. If the range is in a pivot table, this method ungroups the items contained in the range.
- (void) unmerge;  // Separates a merged area into individual cells.

@end

@interface excelCell : excelRange


@end

@interface excelColumn : excelRange


@end

@interface excelRow : excelRange


@end



/*
 * Proofing Suite
 */

// Contains Microsoft Excel autocorrect attributes, capitalization of names of days, correction of two initial capital letters, automatic correction list, and so on.
@interface excelAutocorrect : excelBaseObject

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
@interface excelAxisTitle : excelBaseObject

- (SBElementArray *) characters;

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy) NSString *axisTitleText;  // Returns or sets the text associated with this object.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property excelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property BOOL includeInLayout;  // Returns or sets if a axis title will occupy the chart layout space when a chart layout is being determined.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property NSInteger orientation;  // The text orientation.  Can be a number value from -90 to 90 degrees.
@property excelE216 position;  // Returns or sets the position of the axis title on the chart.
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property excelE284 verticalAlignment;  // Returns or sets the vertical alignment of the object.


@end

// Represents a single axis in a chart.
@interface excelAxis : excelBaseObject

@property BOOL axisBetweenCategories;  // Returns or sets if the value axis crosses the category axis between categories.
@property (readonly) excelE109 axisGroup;  // Returns the group for the specified axis.
@property (copy, readonly) excelAxisTitle *axisTitle;  // Returns an axis title object that represents the title of the specified axis.
@property excelE112 axisType;  // Returns or sets the axis type
@property excelE141 baseUnit;  // Returns or sets the base unit for the specified category axis.
@property BOOL baseUnitIsAuto;  // Returns or sets if Microsoft Excel chooses appropriate base units for the specified category axis.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property excelE292 categoryNames;  // Returns or sets all the category names for the specified axis, as a text array. When you set this property, you can set it to either an array or a range object that contains the category names.
@property excelE142 categoryType;  // Returns or sets the category axis type.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the axis.
@property excelE108 crosses;  // Returns or sets the point on the specified axis where the other axis crosses.
@property double crossesAt;  // Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis.
@property excelE139 displayUnit;  // Returns or sets the unit label for the specified axis. 
@property double displayUnitCustom;  // If the value of the display unit property is custom display unit, the display unit custom property returns or sets the value of the displayed units.
@property (copy, readonly) excelDisplayUnitLabel *displayUnitLabel;  // Returns the display unit label object for the specified axis
@property BOOL hasDisplayUnitLabel;  // Returns or sets if the label specified by the display unit or display unit custom property is displayed on the specified axis. The default value is true.
@property BOOL hasMajorGridlines;  // Returns or sets if the axis has major gridlines. Only axes in the primary axis group can have gridlines.
@property BOOL hasMinorGridlines;  // Returns or sets if the axis has minor gridlines. Only axes in the primary axis group can have gridlines.
@property BOOL hasTitle;  // Returns or sets if the axis or chart has a visible title.
@property (readonly) double height;  // Returns the height of the object.
@property (readonly) double leftPosition;  // Returns the left position of the specified object, in points.
@property double logBase;  // Returns or sets the base of the logarithm when you are using log scales.
@property (copy, readonly) excelGridlines *majorGridlines;  // Returns a gridlines object that represents the major gridlines for the specified axis. Only axes in the primary axis group can have gridlines.
@property excelE115 majorTickMark;  // Returns or sets the type of major tick mark for the specified axis. 
@property double majorUnit;  // Returns or sets the major units for the axis.
@property BOOL majorUnitIsAuto;  // Returns or sets if Microsoft Excel calculates the major units for the axis.
@property excelE141 majorUnitScale;  // Returns or sets the major unit scale value for the category axis when the category type property is set to time scale.
@property double maximumScale;  // Returns or sets the maximum value on the axis.
@property BOOL maximumScaleIsAuto;  // Returns or sets if Microsoft Excel calculates the maximum value for the axis.
@property double minimumScale;  // Returns or sets the minimum value on the axis.
@property BOOL minimumScaleIsAuto;  // Returns or sets if Microsoft Excel calculates the minimum value for the axis.
@property (copy, readonly) excelGridlines *minorGridlines;  // Returns a gridlines object that represents the minor gridlines for the specified axis. Only axes in the primary axis group can have gridlines.
@property excelE115 minorTickMark;  // Returns or sets the type of minor tick mark for the specified axis.
@property double minorUnit;  // Returns or sets the minor units on the axis.
@property BOOL minorUnitIsAuto;  // Returns or sets if Microsoft Excel calculates minor units for the axis.
@property excelE141 minorUnitScale;  // Returns or sets the minor unit scale value for the category axis when the category type property is set to time scale. 
@property BOOL reversePlotOrder;  // Returns or sets if Microsoft Excel plots data points from last to first.
@property excelE106 scaleType;  // Returns or sets the value axis scale type.
@property excelE122 tickLabelPosition;  // Describes the position of tick-mark labels on the specified axis.
@property NSInteger tickLabelSpacing;  // Returns or sets the number of categories or series between tick-mark labels. Applies only to category and series axes.
@property BOOL tickLabelSpacingIsAuto;  // Returns or sets whether or not the tick label spacing is automatic.
@property (copy, readonly) excelTickLabels *tickLabels;  // Returns a tick labels object that represents the tick-mark labels for the specified axis.
@property NSInteger tickMarkSpacing;  // Returns or sets the number of categories or series between tick marks. Applies only to category and series axes.
@property (readonly) double top;  // Returns the top position of the specified object, in points.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents the chart area of a chart. The chart area on a 2-D chart contains the axes, the chart title, the axis titles, and the legend. The chart area on a 3-D chart contains the chart title and the legend; it doesn't include the plot area .
@interface excelChartArea : excelBaseObject

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the chart area.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property double height;  // Returns or sets the height of the object.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (readonly) BOOL roundedCorners;  // Returns or sets if the chart area has rounded corners.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property double width;  // Returns or sets  the width of the object.

- (void) clearContents;  // Clears the data from a chart but leaves the formatting. 

@end

// Used only with charts. Represents fill formatting for chart elements.
@interface excelChartFillFormat : excelBaseObject

@property (copy, readonly) NSColor *backColor;  // Returns the background color.
@property NSInteger backgroundSchemeColor;  // Returns or sets the background color as an index in the current color scheme.
@property (readonly) excelMFdT chartFillFormatType;  // The fill type. 
@property (copy, readonly) NSColor *foreColor;  // Returns the foreground color.
@property NSInteger foregroundSchemeColor;  // Returns or sets the foreground color as an index in the current color scheme.
@property (readonly) excelMGCt gradientColorType;  // Returns the gradient color type for the specified fill.
@property (readonly) double gradientDegree;  // Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 dark through 1.0 light.
@property (readonly) excelMGdS gradientStyle;  // Returns the gradient style for the specified fill.
@property (readonly) NSInteger gradientVariant;  // Returns the shade variant for the specified fill as an integer value from 1 through 4.
@property (readonly) excelPpTy pattern;  // Returns the fill pattern.
@property (readonly) excelMPGb presetGradientType;  // Returns the preset gradient type for the specified fill.
@property (readonly) excelMPzT presetTexture;  // Returns the preset texture for the specified fill.
@property (copy, readonly) NSString *textureName;  // Returns the name of the custom texture file for the specified fill.
@property (readonly) excelMxtT textureType;  // Returns the texture type for the specified fill.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque, and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object is visible.


@end

// Provides access to the Office Art formatting for chart elements.
@interface excelChartFormat : excelBaseObject

@property (copy, readonly) excelFillFormat *fillFormat;  // Returns a fill format object for the parent chart element that contains fill formatting properties for the chart element.
@property (copy, readonly) excelGlowFormat *glowFormat;  // Returns a glow format object for a specified chart that contains glow formatting properties for the chart.
@property (copy, readonly) excelLineFormat *lineFormat;  // Returns a line format object that contains line formatting properties for the specified chart element.
@property (copy, readonly) excelShadowFormat *shadowFormat;  // Returns a shadow format object that contains shadow formatting properties for the chart element.
@property (copy, readonly) excelShapeTextFrame *shapeTextFrame;  // Returns a shape text frame object that contains the alignment and anchoring properties for the specified chart.
@property (copy, readonly) excelSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) excelThreeDFormat *threeDFormat;  // Returns a threeD format object that contains 3-D-effect formatting properties for the specified chart.


@end

// Represents one or more series plotted in a chart with the same format. A chart contains one or more chart groups, each chart group contains one or more series, and each series contains one or more points.
@interface excelChartGroup : excelBaseObject

- (SBElementArray *) seriesCollection;

@property excelE109 axisGroup;
@property NSInteger bubbleScale;  // Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from zero to 300, corresponding to a percentage of the default size. Applies only to bubble charts.
@property NSInteger doughnutHoleSize;  // Returns or sets the size of the hole in a doughnut chart group. The hole size is expressed as a percentage of the chart size, between 10 and 90 percent.
@property (copy, readonly) excelDownBars *downBarsObject;  // Returns a down bars object that represents the down bars on a line chart. Applies only to line charts.
@property (copy, readonly) excelDropLines *dropLinesObject;  // Returns a drop lines object that represents the drop lines for a series on a line chart or area chart. Applies only to line charts or area charts.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property NSInteger firstSliceAngle;  // Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees, clockwise from vertical. Applies only to pie, 3-D pie, and doughnut charts.
@property NSInteger gapWidth;  // Returns or sets: For bar and column charts, the space between bar or column clusters, as a percentage of the bar or column width. For pie of pie and bar of pie charts, the space between the primary and secondary sections of the chart.
@property BOOL hasDropLines;  // Returns or sets if the line chart or area chart has drop lines. Applies only to line and area charts.
@property BOOL hasHiLoLines;  // Returns or sets if the line chart has high-low lines. Applies only to line charts.
@property BOOL hasRadarAxisLabels;  // Returns or sets if a radar chart has axis labels. Applies only to radar charts.
@property BOOL hasSeriesLines;  // Returns or sets if a stacked column chart or bar chart has series lines or if a pie of pie chart or bar of pie chart has connector lines between the two sections. Applies only to stacked column charts, bar charts, pie of pie charts, or bar of pie charts.
@property BOOL hasThreeDShading;  // Returns or sets if the chart group has three-dimensional shading.
@property BOOL hasUpDownBars;  // Returns or sets if a line chart has up and down bars. Applies only to line charts.
@property (copy, readonly) excelHiloLines *hiloLinesObject;  // Returns a hilo lines object that represents the high-low lines for a series on a line chart.
@property NSInteger overlap;  // Returns or sets how bars and columns are positioned. Can be a value between -100 and 100. Applies only to 2-D bar and 2-D column charts.
@property (copy, readonly) excelTickLabels *radarAxisLabels;  // Returns a tick labels object that represents the radar axis labels for the specified chart group.
@property NSInteger secondPlotSize;  // Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. 
@property (copy, readonly) excelSeriesLines *seriesLinesObject;  // Returns a series lines object that represents the series lines for a stacked bar chart or a stacked column chart. Applies only to stacked bar and stacked column charts.
@property BOOL showNegativeBubbles;  // Returns or sets if negative bubbles are shown for the chart group. Valid only for bubble charts.
@property excelE146 sizeRepresents;  // Returns or sets what the bubble size represents on a bubble chart.
@property excelE138 splitType;  // Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split.
@property NSInteger splitValue;  // Returns or sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart.
@property (copy, readonly) excelUpBars *upBarsObject;  // Returns an up bars object that represents the up bars on a line chart. Applies only to line charts.
@property BOOL varyByCategories;  // Returns or sets if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series.


@end

@interface excelAreaGroup : excelChartGroup


@end

@interface excelBarGroup : excelChartGroup


@end

// Represents an embedded chart on a worksheet. The chart object object acts as a container for a chart object. Properties and methods for the chart object object control the appearance and size of the embedded chart on the worksheet.
@interface excelChartObject : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelRange *bottomRightCell;  // Returns a range object that represents the cell that lies under the lower-right corner of the object.
@property (copy, readonly) excelChart *chart;  // Returns a chart object that represents the chart contained in the object.
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or set the height of the object.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property excelE210 placement;  // Returns or sets the way the object is attached to the cells below it.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property BOOL protectChartObject;  // Returns or sets if the embedded chart frame cannot be moved, resized, or deleted.
@property BOOL roundedCorners;  // Returns or sets if the embedded chart has rounded corners.
@property BOOL shadow;  // Returns or sets if the if the object has a shadow.
@property double top;  // Returns or sets the distance from the top edge of the object to the top of row 1, on a worksheet, or the top of the chart area, on a chart.
@property (copy, readonly) excelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.

- (void) bringToFront;  // Brings the object to the front of the z-order.
- (void) copyPictureAppearance:(excelE207)appearance format:(excelE156)format;  // Copies the selected object to the clipboard as a picture.
- (void) cut;  // Cuts the object to the clipboard.
- (void) saveAsPicturePictureType:(excelMPiT)pictureType fileName:(NSString *)fileName;  // Saves the shape in the requested file using the stated graphic format
- (void) sendToBack;  // Sends the object to the back of the z-order.

@end

// Represents the chart title.
@interface excelChartTitle : excelBaseObject

- (SBElementArray *) characters;

@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *caption;  // Returns or sets the chart title text.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the chart title.
@property (copy) NSString *chartTitleText;  // Returns or sets the text for the specified object.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property (copy) NSString *formulaLocal;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formulaR1c1;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the macro.
@property (copy) NSString *formulaR1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property (readonly) double height;  // Returns the height of the object.
@property excelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property BOOL includeInLayout;  // Returns or sets if a chart title will occupy the chart layout space when a chart layout is being determined.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property excelE126 orientation;  // May also be a number value from -90 to 90 degrees.
@property excelE216 position;  // Returns or sets the position of the chart title on the chart.
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property excelE284 verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents a chart in a workbook. The chart can be either an embedded chart, contained in a chart object, or a separate chart sheet.
@interface excelChart : excelBaseObject

- (SBElementArray *) shapes;
- (SBElementArray *) arcs;
- (SBElementArray *) areaGroups;
- (SBElementArray *) barGroups;
- (SBElementArray *) buttons;
- (SBElementArray *) chartGroups;
- (SBElementArray *) chartObjects;
- (SBElementArray *) checkboxes;
- (SBElementArray *) columnGroups;
- (SBElementArray *) doughnutGroups;
- (SBElementArray *) dropdowns;
- (SBElementArray *) groupboxes;
- (SBElementArray *) hyperlinks;
- (SBElementArray *) labels;
- (SBElementArray *) lineGroups;
- (SBElementArray *) lines;
- (SBElementArray *) listboxes;
- (SBElementArray *) optionButtons;
- (SBElementArray *) ovals;
- (SBElementArray *) pieGroups;
- (SBElementArray *) radarGroups;
- (SBElementArray *) rectangles;
- (SBElementArray *) scrollbars;
- (SBElementArray *) seriesCollection;
- (SBElementArray *) spinners;
- (SBElementArray *) textboxes;
- (SBElementArray *) xyGroups;

@property (copy, readonly) excelChartGroup *areaThreeDGroup;  // Returns a chart group object that represents the area chart group on a 3-D chart.
@property BOOL autoScaling;  // Returns or sets if Microsoft Excel scales a 3-D chart so that it's closer in size to the equivalent 2-D chart. The chart's right angle axes property must be true.
@property (copy, readonly) excelWalls *backWall;  // Returns a walls object that allows the user to individually format the back wall of a 3-D chart.
@property excelE143 barShape;  // Returns or sets the shape used with the 3-D bar or column chart.
@property (copy, readonly) excelChartGroup *barThreeDGroup;  // Returns a chart group object that represents the bar chart group on a 3-D chart.
@property (copy, readonly) excelChartArea *chartAreaObject;  // Returns a chart area object that represents the complete chart area for the chart.
@property NSInteger chartStyle;  // Returns or sets the chart type.
@property (copy, readonly) excelChartTitle *chartTitle;  // Returns a chart title object that represents the title of the specified chart.
@property excelE144 chartType;  // Returns or sets the chart type.
@property (copy, readonly) excelChartGroup *columnThreeDGroup;  // Returns a chart group object that represents the column chart group on a 3-D chart.
@property (copy, readonly) excelCorners *cornersObject;  // Returns a corners object that represents the corners of a 3-D chart.
@property (copy, readonly) excelDataTable *dataTableObject;  // Returns a data table object that represents the chart data table.
@property NSInteger depthPercent;  // Returns or sets the depth of a 3-D chart as a percentage of the chart width, between 20 and 2000 percent.
@property excelE118 displayBlanksAs;  // Returns or sets the way that blank cells are plotted on a chart.
@property NSInteger elevation;  // Returns or sets the elevation of the 3-D chart view, in degrees.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) excelFloor *floorObject;  // Returns a floor object that represents the floor of the 3-D chart.
@property NSInteger gapDepth;  // Returns or sets the distance between the data series in a 3-D chart, as a percentage of the marker width. The value of this property must be between 0 and 500.
@property BOOL hasDataTable;  // Returns or sets if the chart has a data table.
@property BOOL hasLegend;  // Returns or sets if the chart has a legend.
@property BOOL hasTitle;  // Returns or sets if the chart has a title.
@property NSInteger heightPercent;  // Returns or sets the height of a 3-D chart as a percentage of the chart width, between 5 and 500 percent.
@property (copy, readonly) excelLegend *legendObject;  // Returns a legend object that represents the legend for the chart.
@property (copy, readonly) excelChartGroup *lineThreeDGroup;  // Returns a chart group object that represents the line chart group on a 3-D chart.
@property (copy) NSString *name;  // Returns or sets the name of the chart.
@property (copy, readonly) excelChart *next;  // Returns a worksheet object that represents the next sheet.
@property (copy, readonly) excelPageSetup *pageSetupObject;  // Returns the page setup object associated with this chart.
@property NSInteger perspective;  // Returns or sets the perspective for the 3-D chart view. Must be between 0 and 100. This property is ignored if the right angle axes property is true.
@property (copy, readonly) excelChartGroup *pieThreeDGroup;  // Returns a chart group object that represents the pie chart group on a 3-D chart.
@property (copy, readonly) excelPlotArea *plotAreaObject;  // Returns a plot area object that represents the plot area of a chart.
@property excelE105 plotBy;  // Returns or sets the way columns or rows are used as data series on the chart.
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
@property (copy, readonly) excelTab *sheetTab;  // Returns the sheet tab of the chart sheet
@property BOOL showDataLabelsOverMaximum;  // Returns or sets whether to show the data labels when the value is greater than the maximum value on the value axis.
@property BOOL showWindow;  // Returns or sets if the embedded chart is displayed in a separate window. The Chart object used with this property must refer to an embedded chart.
@property (copy, readonly) excelWalls *sideWall;  // Returns a walls object that allows the user to individually format the side wall of a 3-D chart.
@property BOOL sizeWithWindow;  // Returns or sets if Microsoft Excel resizes the chart to match the size of the chart sheet window. False if the chart size isn't attached to the window size. Applies only to chart sheets, doesn't apply to embedded charts.
@property (copy, readonly) excelChartGroup *surfaceGroup;  // Returns a chart group object that represents the surface chart group of a 3-D chart.
@property excelE225 visible;  // Returns or sets if the chart is visible.
@property BOOL wallsAndGridlinesTwoD;  // Returns or sets if gridlines are drawn two-dimensionally on a 3-D chart.
@property (copy, readonly) excelWalls *wallsObject;  // Returns a walls object that represents the walls of the 3-D chart.

- (void) applyCustomChartTypeChartType:(excelE102)chartType chartName:(NSString *)chartName;  // Applies a standard or custom chart type to a chart or series.
- (void) applyLayoutLayout:(NSInteger)layout chartType:(excelE144)chartType;  // Applies a chart layout.
- (excelChart *) chartLocationWhere:(excelE201)where name:(NSString *)name;  // Moves the chart to a new location.
- (void) chartWizardSource:(excelRange *)source gallery:(excelE144)gallery format:(NSInteger)format plotBy:(excelE105)plotBy categoryLabels:(NSInteger)categoryLabels seriesLabels:(NSInteger)seriesLabels hasLegend:(BOOL)hasLegend title:(NSString *)title categoryTitle:(NSString *)categoryTitle valueTitle:(NSString *)valueTitle extraTitle:(NSString *)extraTitle;  // Modifies the properties of the given chart. You can use this method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties.
- (void) checkSpellingCustomDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase alwaysSuggest:(BOOL)alwaysSuggest;  // Checks the spelling of an object.
- (void) clearToMatchStyle;  // Sets the chart formatting to the last predefined chart style applied.
- (void) copyChartAsPictureAppearance:(excelE207)appearance format:(excelE156)format outputSize:(excelE207)outputSize;  // Copies the selected object to the clipboard as a picture.
- (void) deselect;  // Cancels the selection for the specified chart.
- (excelAxis *) getAxisAxisType:(excelE112)axisType whichAxis:(excelE109)whichAxis;  // Returns an object that represents a specified axis object.
- (NSArray *) getChartElementX:(NSInteger)x y:(NSInteger)y;  // Returns information about the chart element at specified X and Y coordinates. This method returns a list  of three items.  Please refer to the documentation on the meaning of the returned values.
- (BOOL) getHasAxisAxisType:(excelE112)axisType axisGroup:(excelE109)axisGroup;  // Returns which axes exist on the chart.
- (void) pasteChartFormat:(excelE204)format;  // Pastes chart data from the clipboard into the specified chart.
- (void) printOutFrom:(NSInteger)from to:(NSInteger)to copies:(NSInteger)copies preview:(BOOL)preview activePrinter:(NSString *)activePrinter printToFile:(BOOL)printToFile collate:(BOOL)collate;  // Prints the object
- (void) printPreviewEnableChanges:(BOOL)enableChanges;  // Shows a preview of the object as it would look when printed. This function is deprecated.
- (void) protectChartPassword:(NSString *)password drawingObjects:(BOOL)drawingObjects chartContents:(BOOL)chartContents userInterfaceOnly:(BOOL)userInterfaceOnly;  // Protects a chart so that it cannot be modified.
- (void) refresh;  // Updates the cache of the chart object.
- (void) saveAsFilename:(NSString *)filename fileFormat:(excelE188)fileFormat password:(NSString *)password writeReservationPassword:(NSString *)writeReservationPassword readOnlyRecommended:(BOOL)readOnlyRecommended createBackup:(BOOL)createBackup addToMostRecentlyUsedList:(BOOL)addToMostRecentlyUsedList saveAsLocalLanguage:(BOOL)saveAsLocalLanguage;  // Saves changes into a different file.
- (void) setBackgroundPicturePictureFileName:(NSString *)pictureFileName;  // Sets the background graphic for a chart.
- (void) setChartElementChartElement:(id)chartElement;  // Sets chart elements on a chart.
- (void) setHasAxisAxisExists:(BOOL)axisExists axisType:(excelE112)axisType axisGroup:(excelE109)axisGroup;  // Sets which axes exist on the chart.
- (void) setSourceDataSource:(excelRange *)source plotBy:(excelE105)plotBy;  // Sets the source data range for the chart.
- (void) unprotectPassword:(NSString *)password;  // Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.

@end

// A chart sheet is a worksheet that contains only a chart.
@interface excelChartSheet : excelChart

@property (copy, readonly) excelChart *chart;  // Returns the chart for this chart sheet.
@property (readonly) excelE151 worksheetType;  // Returns the type of this worksheet.


@end

@interface excelColumnGroup : excelChartGroup


@end

// Represents the corners of a 3-D chart.
@interface excelCorners : excelBaseObject

@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the data label on a chart point or trendline.
@interface excelDataLabel : excelBaseObject

- (SBElementArray *) characters;

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property BOOL autoText;  // Returns or sets if the object automatically generates appropriate text based on context.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *caption;  // Returns or sets the caption of the object.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the data label.
@property (copy) NSString *dataLabelText;  // Returns or sets the text for this object.
@property excelE134 dataLabelType;  // Returns or sets the type of the data label.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property (copy) NSString *formulaLocal;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formulaR1c1;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the macro.
@property (copy) NSString *formulaR1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property (readonly) double height;  // Returns the height of the object.
@property excelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property BOOL numberFormatLinked;  // Returns or sets if the number format is linked to the cells, so that the number format changes in the labels when it changes in the cells.
@property (copy) NSString *numberFormatLocal;  // Returns or sets the format code for the object as a string in the language of the user.
@property NSInteger orientation;  // The text orientation.  Can be a number value from -90 to 90 degrees.
@property excelE140 position;  // Returns or sets the position of the specified object.
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property BOOL showLegendKey;  // Returns or sets if the data label legend key is visible.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property excelE284 verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents a chart data table.
@interface excelDataTable : excelBaseObject

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the data table.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property BOOL hasBorderHorizontal;  // Returns or sets if the chart data table has horizontal cell borders. 
@property BOOL hasBorderOutline;  // Returns or sets if the chart data table has outline borders.
@property BOOL hasBorderVertical;  // Returns or sets if the chart data table has vertical cell borders. 
@property BOOL showLegendKey;  // Returns or sets if the data label legend key is visible.


@end

// Represents a unit label on an axis in the specified chart. Unit labels are useful for charting large values-- for example, in the millions or billions.
@interface excelDisplayUnitLabel : excelBaseObject

- (SBElementArray *) characters;

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *caption;  // Returns or sets the caption for the object.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the chart title.
@property (copy) NSString *displayLabelUnitText;  // Returns or sets the text for this object.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property excelE121 horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property NSInteger orientation;  // The text orientation.  Can be a number value from -90 to 90 degrees.
@property excelE216 position;  // Returns or sets the position of the chart title on the chart.
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property excelE284 verticalAlignment;  // Returns or sets the vertical alignment of the object.


@end

@interface excelDoughnutGroup : excelChartGroup


@end

// Represents the down bars in a chart group. Down bars connect points on the first series in the chart group with lower values on the last series, the lines go down from the first series.
@interface excelDownBars : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the down bars.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy, readonly) NSString *name;  // Returns the name of the object.


@end

// Represents the drop lines in a chart group. Drop lines connect the points in the chart with the x-axis. Only line and area chart groups can have drop lines.
@interface excelDropLines : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the drop lines.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the error bars on a chart series. Error bars indicate the degree of uncertainty for chart data.
@interface excelErrorBars : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the error bars.
@property excelE104 endStyle;  // Returns or sets the end style for the error bars.
@property (copy, readonly) NSString *name;  // Returns the name of the object.


@end

// Represents the floor of a 3-D chart.
@interface excelFloor : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the floor.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property excelE124 pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property NSInteger thickness;  // Returns or sets  the thickness of the floor.


@end

// Represents major or minor gridlines on a chart axis. Gridlines extend the tick marks on a chart axis to make it easier to see the values associated with the data markers.
@interface excelGridlines : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the gridlines.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the high-low lines in a chart group. High-low lines connect the highest point with the lowest point in every category in the chart group. Only 2-D line groups can have high-low lines.
@interface excelHiloLines : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the hilo lines.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the interior of an object.
@interface excelInterior : excelBaseObject

@property (copy) NSColor *color;  // Returns or sets the primary color of the object.
@property excelE103 colorIndex;  // Returns or sets the color of the interior. The color is specified as an index value into the current color palette.
@property BOOL invertIfNegative;  // Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.
@property (copy, readonly) excelLinearGradient *linearGradient;  // Returns or sets the Gradient property of an Interior object of a selection.
@property excelE137 pattern;  // Returns or sets the interior pattern.
@property (copy) NSColor *patternColor;  // Returns or sets the color of the interior pattern as an RGB value.
@property excelE103 patternColorIndex;  // Returns or sets the color of the interior pattern as an index into the current color palette.
@property excelE305 patternThemeColor;  // Returns or sets a theme color pattern for an Interior object.
@property double patternTintAndShade;  // Returns or sets a tint and shade pattern for an Interior object.
@property (copy, readonly) excelRectangularGradient *rectangularGradient;  // Returns or sets the Gradient property of an Interior object of a selection.
@property excelE305 themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.


@end

// Represents leader lines on a chart. Leader lines connect data labels to data points.
@interface excelLeaderLines : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the leader lines.


@end

// Represents a legend entry in a chart legend.
@interface excelLegendEntry : excelBaseObject

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the legend entry.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (readonly) double height;  // Returns the height of the object.
@property (readonly) double leftPosition;  // Returns the left position of the specified object, in points.
@property (copy, readonly) excelLegendKey *legendKey;  // Returns a legend key object that represents the legend key associated with the entry.
@property (readonly) double top;  // Returns the top position of the specified object, in points.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents a legend key in a chart legend. Each legend key is a graphic that visually links a legend entry with its associated series or trendline in the chart.
@interface excelLegendKey : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the legend key.
@property (readonly) double height;  // Returns the height of the object.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property BOOL invertIfNegative;  // Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.
@property (readonly) double leftPosition;  // Returns the left position of the specified object, in points.
@property (copy) NSColor *markerBackgroundColor;  // Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts.
@property excelE103 markerBackgroundColorIndex;  // Returns or sets the marker background color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property (copy) NSColor *markerForegroundColor;  // Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts.
@property excelE103 markerForegroundColorIndex;  // Returns or sets the marker foreground color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property NSInteger markerSize;  // Returns or sets the data-marker size, in points.
@property excelE135 markerStyle;  // Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart.
@property excelE124 pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property double pictureUnit;  // Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property BOOL smooth;  // Returns or sets if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts.
@property (readonly) double top;  // Returns the top position of the specified object, in points.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents the legend in a chart. Each chart can have only one legend.
@interface excelLegend : excelBaseObject

- (SBElementArray *) legendEntries;

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the legend.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property double height;  // Returns or sets the height of the object.
@property BOOL includeInLayout;  // Returns or sets if a legend will occupy the chart layout space when a chart layout is being determined.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property excelE123 position;  // Returns or sets the position of the legend on the chart.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property double width;  // Returns or sets  the width of the object.


@end

@interface excelLineGroup : excelChartGroup


@end

@interface excelPieGroup : excelChartGroup


@end

// Represents the plot area of a chart. This is the area where your chart data is plotted.
@interface excelPlotArea : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the plot area.
@property double height;  // Returns or sets the height of the object.
@property (readonly) double insideHeight;  // Returns the inside height of the plot area, in points.
@property (readonly) double insideLeft;  // Returns the distance from the chart edge to the inside left edge of the plot area, in points.
@property (readonly) double insideTop;  // Returns the distance from the chart edge to the inside top edge of the plot area, in points.
@property (readonly) double insideWidth;  // Returns the inside width of the plot area, in points.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property excelE216 position;  // Returns or sets the position of the plot area on the chart.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property double width;  // Returns or sets  the width of the object.


@end

@interface excelRadarGroup : excelChartGroup


@end

// Represents series lines in a chart group. Series lines connect the data values from each series. Only 2-D stacked bar or column chart groups can have series lines.
@interface excelSeriesLines : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the series lines.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents a single point in a series in a chart.
@interface excelSeriesPoint : excelBaseObject

@property BOOL applyPictToEnd;  // Returns or sets if a picture is applied to the end of the point or all points in the series.
@property BOOL applyPictToFront;  // Returns or sets if a picture is applied to the front of the point or all points in the series.
@property BOOL applyPictToSides;  // Returns or sets if a picture is applied to the sides of the point or all points in the series.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the point.
@property (copy, readonly) excelDataLabel *dataLabelObject;  // Returns a data label object that represents the data label associated with the point or trendline.
@property NSInteger explosion;  // Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns zero if there's no explosion, the tip of the slice is in the center of the pie.
@property BOOL hasDataLabel;  // Returns or sets if the point has a data label.
@property BOOL hasThreeDEffect;  // Returns or sets if a point has a three-dimensional appearance. Applies only to bubble charts.
@property (readonly) double height;  // Returns the height of the object.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property BOOL invertIfNegative;  // Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.
@property (readonly) double leftPosition;  // Returns the left position of the specified object, in points.
@property (copy) NSColor *markerBackgroundColor;  // Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts.
@property excelE103 markerBackgroundColorIndex;  // Returns or sets the marker background color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property (copy) NSColor *markerForegroundColor;  // Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts.
@property excelE103 markerForegroundColorIndex;  // Returns or sets the marker foreground color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property NSInteger markerSize;  // Returns or sets the data-marker size, in points.
@property excelE135 markerStyle;  // Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property excelE124 pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property double pictureUnit;  // Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored.
@property BOOL secondaryPlot;  // Returns or sets if the point is in the secondary section of either a pie of pie chart or a bar of pie chart. Applies only to points on pie of pie charts or bar of pie charts.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property (readonly) double top;  // Returns the top position of the specified object, in points.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents a series in a chart.
@interface excelSeries : excelBaseObject

- (SBElementArray *) dataLabels;
- (SBElementArray *) seriesPoints;
- (SBElementArray *) trendlines;

@property BOOL applyPictureToEnd;  // Returns or sets if a picture is applied to the end of the point or all points in the series.
@property BOOL applyPictureToFront;  // Returns or sets if a picture is applied to the front of the point or all points in the series.
@property BOOL applyPictureToSides;  // Returns or sets if a picture is applied to the sides of the point or all points in the series.
@property excelE109 axisGroup;  // Returns or sets the axis group for the specified series.
@property excelE143 barShape;  // Returns or sets the shape used with the 3-D bar or column chart.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *bubbleSizes;  // Returns or sets a string in A1-style notation that refers to the worksheet cells containing the size data for the bubble chart. Applies only to bubble charts. 
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the series.
@property excelE144 chartType;  // Returns or sets the chart type.
@property (copy, readonly) excelErrorBars *errorBars;  // Returns an error bars object that represents the error bars for the series.
@property NSInteger explosion;  // Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns zero if there's no explosion, the tip of the slice is in the center of the pie.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property (copy) NSString *formulaLocal;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formulaR1c1;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the macro.
@property (copy) NSString *formulaR1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property BOOL hasDataLabels;  // Returns or sets if the series has data labels.
@property BOOL hasErrorBars;  // Returns or set if the series has error bars. This property isn't available for 3-D charts.
@property BOOL hasLeaderLines;  // Returns or sets if the series has leader lines.
@property BOOL hasThreeDEffect;  // Returns or sets if the series has a three-dimensional appearance. Applies only to bubble charts.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property BOOL invertIfNegative;  // Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.
@property (copy, readonly) excelLeaderLines *leaderLines;  // Returns a leader lines object that represents the leader lines for the series.
@property (copy) NSColor *markerBackgroundColor;  // Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts.
@property excelE103 markerBackgroundColorIndex;  // Returns or sets the marker background color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property (copy) NSColor *markerForegroundColor;  // Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts.
@property excelE103 markerForegroundColorIndex;  // Returns or sets the marker foreground color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property NSInteger markerSize;  // Returns or sets the data-marker size, in points.
@property excelE135 markerStyle;  // Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property excelE124 pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property double pictureUnit;  // Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored.
@property (readonly) NSInteger plotColorIndex;  // Returns the plot color index of the series.
@property NSInteger plotOrder;  // Returns or sets the plot order for the selected series within the chart group.
@property excelE292 seriesValues;  // Returns or sets a list of all the values in the series. This can be a range on a worksheet or a list of constant values.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property BOOL smooth;  // Returns or sets if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts.
@property excelE292 xvalues;  // Returns or sets a list of x values for a chart series. The xvalues property can be set to a range on a worksheet or to a list of values.

- (void) errorBarDirection:(excelE116)direction include:(excelE117)include type:(excelE131)type amount:(double)amount minusValues:(double)minusValues;  // Applies error bars to the series.
- (void) pasteSeries;

@end

// Represents the tick-mark labels associated with tick marks on a chart axis.
@interface excelTickLabels : excelBaseObject

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the tick labels.
@property (readonly) NSInteger depth;  // Returns the number of levels of category tick labels.
@property (copy, readonly) excelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property BOOL multiLevel;  // Returns or sets whether an axis is multilevel or not.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property BOOL numberFormatLinked;  // Returns or sets if the number format is linked to the cells, so that the number format changes in the labels when it changes in the cells.
@property (copy) NSString *numberFormatLocal;  // Returns or sets the format code for the object as a string in the language of the user.
@property NSInteger offset;  // Returns or sets the distance between the levels of labels, and the distance between the first level and the axis line. The value can be an integer percentage from 0 through 1000, relative to the axis label's font size. 
@property excelE127 orientation;  // The text orientation.  Can be a number value from -90 to 90 degrees.
@property excelE265 readingOrder;  // Returns or sets the reading order for the specified object.
@property excelE299 tickAlignment;  // Returns or sets the alignment for the specified tick label.


@end

// Represents a trendline in a chart. A trendline shows the trend, or direction, of data in a series.
@interface excelTrendline : excelBaseObject

@property double backward;  // Returns or sets the number of periods, or units on a scatter chart, that the trendline extends backward.
@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the trendline.
@property (copy, readonly) excelDataLabel *dataLabelObject;  // Returns a data label object that represents the data label associated with the point or trendline.
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
@property excelE132 trendlineType;  // Returns or sets the type of this trend line


@end

// Represents the up bars in a chart group. Up bars connect points on series one with higher values on the last series in the chart group, the lines go up from series one. Only 2-D line groups that contain at least two series can have up bars.
@interface excelUpBars : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the up bars.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the walls of a 3-D chart.
@interface excelWalls : excelBaseObject

@property (copy, readonly) excelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) excelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) excelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the walls.
@property (copy, readonly) excelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy, readonly) NSString *name;  // Returns or sets the name of the object.
@property excelE124 pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property NSInteger pictureUnit;  // Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored.
@property NSInteger thickness;  // Returns or sets  the thickness of the floor.


@end

@interface excelXyGroup : excelChartGroup


@end

