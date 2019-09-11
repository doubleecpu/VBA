Option Compare Database
Option Explicit
' Date: 2019/09/11
' Author: Gilbert Medel
' Current Version: 3.1.0
' Notes: Used to store enumerations so VBA can run without references to other Office 2013/2016 products
' And List custom Enumerations used in VBA
'
' Public Variables
Public Enum MyVBA_Enum
    MyVBA_Date_Time
    MyVBA_Date_Only
    MyVBA_Date_Gov
    MyVBA_Time_24
    MyVBA_Time_12
    MyVBA_File_Name_by_Year
    MyVBA_File_Name_by_Year_Time
    MyVBA_Date_Value
    MyVBA_Local_Time_Value
End Enum
Public Enum MyVBA_SYNC
    MyVBA_SYNC_Push = 1
    MyVBA_SYNC_Pull = 2
End Enum
Public Enum Excel_Constants
    'Name  = Value 'Description
    xlAddIn = 18 'Microsoft Excel 97-2003 Add-In
    xlAddIn8 = 18 'Microsoft Excel 97-2003 Add-In
    xlCSV = 6 'CSV
    xlCSVMac = 22 'Macintosh CSV
    xlCSVMSDOS = 24 'MSDOS CSV
    xlCSVWindows = 23 'Windows CSV
    xlCurrentPlatformText = -4158 'Current Platform Text
    xlDBF2 = 7 'DBF2
    xlDBF3 = 8 'DBF3
    xlDBF4 = 11 'DBF4
    xlDIF = 9 'DIF
    xlExcel12 = 50 'Excel12
    xlExcel2 = 16 'Excel2
    xlExcel2FarEast = 27 'Excel2 FarEast
    xlExcel3 = 29 'Excel3
    xlExcel4 = 33 'Excel4
    xlExcel4Workbook = 35 'Excel4 Workbook
    xlExcel5 = 39 'Excel5
    xlExcel7 = 39 'Excel7
    xlExcel8 = 56 'Excel8
    xlExcel9795 = 43 'Excel9795
    xlHtml = 44 'HTML Format
    xlIntlAddIn = 26 'International Add-In
    xlIntlMacro = 25 'International Macro
    xlOpenDocumentSpreadsheet = 60 'OpenDocument Spreadsheet
    xlOpenXMLAddIn = 55 'Open XML Add-In
    xlOpenXMLTemplate = 54 'Open XML Template
    xlOpenXMLTemplateMacroEnabled = 53 'Open XML Template Macro Enabled
    xlOpenXMLWorkbook = 51 'Open XML Workbook
    xlOpenXMLWorkbookMacroEnabled = 52 'Open XML Workbook Macro Enabled
    xlSYLK = 2 'SYLK
    xlTemplate = 17 'Template
    xlTemplate8 = 17 'Template 8
    xlTextMac = 19 'Macintosh Text
    xlTextMSDOS = 21 'MSDOS Text
    xlTextPrinter = 36 'Printer Text
    xlTextWindows = 20 'Windows Text
    xlUnicodeText = 42 'Unicode Text
    xlWebArchive = 45 'Web Archive
    xlWJ2WD1 = 14 'WJ2WD1
    xlWJ3 = 40 'WJ3
    xlWJ3FJ3 = 41 'WJ3FJ3
    xlWK1 = 5 'WK1
    xlWK1ALL = 31 'WK1ALL
    xlWK1FMT = 30 'WK1FMT
    xlWK3 = 15 'WK3
    xlWK3FM3 = 32 'WK3FM3
    xlWK4 = 38 'WK4
    xlWKS = 4 'Worksheet
    xlWorkbookDefault = 51 'Workbook Default
    xlWorkbookNormal = -4143 'Workbook normal
    xlWorks2FarEast = 28 'Works2 FarEast
    xlWQ1 = 34 'WQ1
    xlXMLSpreadsheet = 46 'XML Spreadsheet
    'XL_Reading_Order
'    xlContext = -5002 'According to context.
'    xlLTR = -5003 'Left-to-right.
'    xlRTL = -5004 'Right-to-left.
    'XL_Line_Style
    xlContinuous = 1 'Continuous line.
    xlDash = -4115 'Dashed line.
    xlDashDot = 4 ' Alternating dashes and dots.
    xlDashDotDot = 5 'Dash followed by two dots.
    xlDot = -4118 'Dotted line.
    xlDouble = -4119 'Double line.
    xlLineStyleNone = -4142 'No line.
    xlSlantDashDot = 13 'Slanted dashes.
    'XL_Borders_Index
    xlDiagonalDown = 5 'Border running from the upper-left corner to the lower-right of each cell in the range.
    xlDiagonalUp = 6 'Border running from the lower-left corner to the upper-right of each cell in the range.
    xlEdgeBottom = 9 'Border at the bottom of the range.
    xlEdgeLeft = 7 'Border at the left edge of the range.
    xlEdgeRight = 10 'Border at the right edge of the range.
    xlEdgeTop = 8 'Border at the top of the range.
    xlInsideHorizontal = 12 'Horizontal borders for all cells in the range except borders on the outside of the range.
    xlInsideVertical = 11 'Vertical borders for all the cells in the range except borders on the outside of the range.
    'XL_Border_Weight
    xlHairline = 1 'Hairline (thinnest border).
    xlMedium = -4138 'Medium.
    xlThick = 4 'Thick (widest border).
    xlThin = 2 'Thin.
    'XL_Pattern
    xlPatternAutomatic = -4105 'Excel controls the pattern.
    xlPatternChecker = 9 'Checkerboard.
    xlPatternCrissCross = 16 'Criss-cross lines.
    xlPatternDown = -4121 'Dark diagonal lines running from the upper-left to the lower-right.
    xlPatternGray16 = 17 '16% gray.
    xlPatternGray25 = -4124 '25% gray.
    xlPatternGray50 = -4125 '50% gray.
    xlPatternGray75 = -4126 '75% gray.
    xlPatternGray8 = 18 '8% gray.
    xlPatternGrid = 15 'Grid.
    xlPatternHorizontal = -4128 'Dark horizontal lines.
    xlPatternLightDown = 13 'Light diagonal lines running from the upper-left to the lower-right.
    xlPatternLightHorizontal = 11 'Light horizontal lines.
    xlPatternLightUp = 14 'Light diagonal lines running from the lower-left to the upper-right.
    xlPatternLightVertical = 12 'Light vertical bars.
    xlPatternNone = -4142 'No pattern.
    xlPatternSemiGray75 = 10 '75% dark gray.
    xlPatternSolid = 1 'Solid color.
    xlPatternUp = -4162 'Dark diagonal lines running from the lower-left to the upper-right.
    xlPatternVertical = -4166 'Dark vertical bars.
    'XL_Constants
    xl3DBar = -4099  '3D Bar
    xl3DEffects1 = 13    '3D Effects1
    xl3DEffects2 = 14    '3D Effects2
    xl3DSurface = -4103  '3D Surface
    xlAbove = 0  'Above
    xlAccounting1 = 4    'Accounting1
    xlAccounting2 = 5    'Accounting2
    xlAccounting4 = 17   'Accounting4
    xlAdd = 2    'Add
    xlAll = -4104    'All
    xlAccounting3 = 6    'Accounting3
    xlAllExceptBorders = 7   'All Except Borders
    xlAutomatic = -4105  'Automatic
    xlBar = 2    'Automatic
    xlBelow = 1  'Below
    xlBidi = -5000   'Bidi
    xlBidiCalendar = 3   'BidiCalendar
    xlBoth = 1   'Both
    xlBottom = -4107     'Bottom
    xlCascade = 7    'Cascade
    xlCenter = -4108     'Center
    xlCenterAcrossSelection = 7  'Center Across Selection
    xlChart4 = 2     'Chart 4
    xlChartSeries = 17   'Chart Series
    xlChartShort = 6     'Chart Short
    xlChartTitles = 18   'Chart Titles
    xlChecker = 9    'Checker
    xlCircle = 8     'Circle
    xlClassic1 = 1   'Classic1
    xlClassic2 = 2   'Classic2
    xlClassic3 = 3   'Classic3
    xlClosed = 3     'Closed
    xlColor1 = 7     'Color1
    xlColor2 = 8     'Color2
    xlColor3 = 9     'Color3
    xlColumn = 3     'Column
    xlCombination = -4111    'Combination
    xlComplete = 4   'Complete
    xlConstants = 2  'Constants
    xlContents = 2   'Contents
    xlContext = -5002    'Context
    xlCorner = 2     'Corner
    xlCrissCross = 16    'CrissCross
    xlCross = 4  'Cross
    xlCustom = -4114     'Custom
    xlDebugCodePane = 13 'Debug Code Pane
    xlDefaultAutoFormat = -1 'Default Auto Format
    xlDesktop = 9    'Desktop
    xlDiamond = 2    'Diamond
    xlDirect = 1     'Direct
    xlDistributed = -4117    'Distributed
    xlDivide = 5     'Divide
    xlDoubleAccounting = 5   'Double Accounting
    xlDoubleClosed = 5   'Double Closed
    xlDoubleOpen = 4     'Double Open
    xlDoubleQuote = 1    'Double Quote
    xlDrawingObject = 14 'Drawing Object
    xlEntireChart = 20   'Entire Chart
    xlExcelMenus = 1     'Excel Menus
    xlExtended = 3   'Extended
    xlFill = 5   'Fill
    xlFirst = 0  'First
    xlFixedValue = 1     'Fixed Value
    xlFloating = 5   'Floating
    xlFormats = -4122    'Formats
    xlFormula = 5    'Formula
    xlFullScript = 1     'Full Script
    xlGeneral = 1    'General
    xlGray16 = 17    'Gray16
    xlGray25 = -4124     'Gray25
    xlGray50 = -4125     'Gray50
    xlGray75 = -4126     'Gray75
    xlGray8 = 18 'Gray8
    xlGregorian = 2  'Gregorian
    xlGrid = 15  'Grid
    xlGridline = 22  'Gridline
    xlHigh = -4127   'High
    xlHindiNumerals = 3  'Hindi Numerals
    xlIcons = 1  'Icons
    xlImmediatePane = 12 'Immediate Pane
    xlInside = 2     'Inside
    xlInteger = 2    'Integer
    xlJustify = -4130    'Justify
    xlLast = 1   'Last
    xlLastCell = 11  'Last Cell
    xlLatin = -5001  'Latin
    xlLeft = -4131   'Left
    xlLeftToRight = 2    'Left To Right
    xlLightDown = 13 'Light Down
    xlLightHorizontal = 11   'Light Horizontal
    xlLightUp = 14   'Light Up
    xlLightVertical = 12 'Light Vertical
    xlList1 = 10 'List1
    xlList2 = 11 'List2
    xlList3 = 12 'List3
    xlLocalFormat1 = 15  'Local Format1
    xlLocalFormat2 = 16  'Local Format2
    xlLogicalCursor = 1  'Logical Cursor
    xlLong = 3   'Long
    xlLotusHelp = 2  'Lotus Help
    xlLow = -4134    'Low
    xlLTR = -5003    'LTR
    xlMacrosheetCell = 7     'MacrosheetCell
    xlManual = -4135     'Manual
    xlMaximum = 2    'Maximum
    xlMinimum = 4    'Minimum
    xlMinusValues = 3    'Minus Values
    xlMixed = 2  'Mixed
    xlMixedAuthorizedScript = 4  'Mixed Authorized Script
    xlMixedScript = 3    'Mixed Script
    xlModule = -4141     'Module
    xlMultiply = 4   'Multiply
    xlNarrow = 1     'Narrow
    xlNextToAxis = 4     'Next To Axis
    xlNoDocuments = 3    'No Documents
    xlNone = -4142   'None
    xlNotes = -4144  'Notes
    xlOff = -4146    'Off
    xlOn = 1     'On
    xlOpaque = 3     'Opaque
    xlOpen = 2   'Open
    xlOutside = 3    'Outside
    xlPartial = 3    'Partial
    xlPartialScript = 2  'Partial Script
    xlPercent = 2    'Percent
    xlPlus = 9   'Plus
    xlPlusValues = 2     'Plus Values
    xlReference = 4  'Reference
    xlRight = -4152  'Right
    xlRTL = -5004    'RTL
    xlScale = 3  'Scale
    xlSemiautomatic = 2  'Semiautomatic
    xlSemiGray75 = 10    'SemiGray75
    xlShort = 1  'Short
    xlShowLabel = 4  'Show Label
    xlShowLabelAndPercent = 5    'Show Label and Percent
    xlShowPercent = 3    'Show Percent
    xlShowValue = 2  'Show Value
    xlSimple = -4154     'Simple
    xlSingle = 2     'Single
    xlSingleAccounting = 4   'Single Accounting
    xlSingleQuote = 2    'Single Quote
    xlSolid = 1  'Solid
    xlSquare = 1     'Square
    xlStar = 5   'Star
    xlStError = 4    'St Error
    xlStrict = 2     'Strict
    xlSubtract = 3   'Subtract
    xlSystem = 1     'System
    xlTextBox = 16   'Text Box
    xlTiled = 1  'Tiled
    xlTitleBar = 8   'Title Bar
    xlToolbar = 1    'Toolbar
    xlToolbarButton = 2  'Toolbar Button
    xlTop = -4160    'Top
    xlTopToBottom = 1    'Top To Bottom
    xlTransparent = 2    'Transparent
    xlTriangle = 3   'Triangle
    xlVeryHidden = 2     'Very Hidden
    xlVisible = 12   'Visible
    xlVisualCursor = 2   'Visual Cursor
    xlWatchPane = 11 'Watch Pane
    xlWide = 3   'Wide
    xlWorkbookTab = 6    'Workbook Tab
    xlWorksheet4 = 1     'Worksheet4
    xlWorksheetCell = 3  'Worksheet Cell
    xlWorksheetShort = 5     'Worksheet Short
End Enum
Public Enum mso_Text_Orientation
    msoTextOrientationDownward = 3 'Downward.
    msoTextOrientationHorizontal = 1 'Horizontal.
    msoTextOrientationHorizontalRotatedFarEast = 6 'Horizontal and rotated as required for Asian language support.
    msoTextOrientationMixed = -2 'Not supported.
    msoTextOrientationUpward = 2 'Upward.
    msoTextOrientationVertical = 5 'Vertical.
    msoTextOrientationVerticalFarEast = 4 'Vertical as required for Asian language support.
End Enum
Public Enum mso_File_Dialog_Type
    msoFileDialogOpen = 1 'Open dialog box.
    msoFileDialogSaveAs = 2 'Save As dialog box.
    msoFileDialogFilePicker = 3 'File picker dialog box.
    msoFileDialogFolderPicker = 4 'Folder picker dialog box.
End Enum
Public Enum Mso_Shape_Type
    msoAutoShape = 1 'AutoShape.
    msoCallout = 2 'Callout.
    msoCanvas = 20 ' Canvas.
    msoChart = 3 'Chart.
    msoComment = 4 'Comment.
    msoContentApp = 27 'Content Office Add-in
    msoDiagram = 21 'Diagram.
    msoEmbeddedOLEObject = 7 'Embedded OLE object.
    msoFormControl = 8 'Form control.
    msoFreeform = 5 'Freeform.
    msoGraphic = 28 'Graphic
    msoGroup = 6 'Group.
    msoIgxGraphic = 24 'SmartArt graphic
    msoInk = 22 'Ink
    msoInkComment = 23 'Ink comment
    msoLine = 9 'Line
    msoLinkedGraphic = 29 'Linked graphic
    msoLinkedOLEObject = 10 'Linked OLE object
    msoLinkedPicture = 11 'Linked picture
    msoMedia = 16 'Media
    msoOLEControlObject = 12 'OLE control object
    msoPicture = 13 'Picture
    msoPlaceholder = 14 'Placeholder
    msoScriptAnchor = 18 'Script anchor
    msoShapeTypeMixed = -2 'Mixed shape type
    msoTable = 19 'Table
    msoTextBox = 17 'Text box
    msoTextEffect = 15 'Text effect
    msoWebVideo = 26 'Web video
End Enum

Public Enum Mso_Text_Effect_Alignment
    msoTextEffectAlignmentCentered = 2 'Centered.
    msoTextEffectAlignmentLeft = 1 'Left-aligned.
    msoTextEffectAlignmentLetterJustify = 4 'Text is justified. Spacing between letters may be adjusted to justify text.
    msoTextEffectAlignmentMixed = -2 'Not used.
    msoTextEffectAlignmentRight = 3 'Right- aligned.
    msoTextEffectAlignmentStretchJustify = 6 'Text is justified. Letters may be stretched to justify text.
    msoTextEffectAlignmentWordJustify = 5 'Text is justified. Spacing between words (but not letters) may be adjusted to justify text.
End Enum
'************xxxxxxxxxxxx''''''''''''''''
'           Default Look Ahead Report
'A TWIP (TWentieth of an Imperial Point) is a 1/20 of a Point (1/72 of an inch).For display of metric units converts using exactly 567 twips per centimetre.
Public Const MyVBA_TWIP As Integer = 1440

Public Enum VBA_TriState
    VBA_CTrue = 1
    VBA_False = 0
    VBA_TriStateMixed = -2
    VBA_TriStateToggle = -3
    VBA_True = -1
End Enum
