unit WPDefs;
{ -----------------------------------------------------------------------------
  WPDefs   - Copyright (C) 2002 by wpcubed GmbH    -   Author: Julian Ziersch
  info: http://www.wptools.de                         mailto:support@wptools.de
  *****************************************************************************
  -----------------------------------------------------------------------------
  WPTools4 RTF Engine
  This unit may not be distributed!
  -----------------------------------------------------------------------------
  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
  ----------------------------------------------------------------------------- }


{$I WPINC.INC}

interface
uses Graphics, Classes, Controls, Dialogs, SysUtils, Forms,
  Printers, Messages, ExtCtrls, Windows, stdctrls, TypInfo, WPLocalize;

{$IFNDEF WPDEBUG}{$D-}{$Y-}{$Q-}{$R-}{$L-}{$ENDIF}

{$R WPFillBits.res}
{$A-} { all records are packed! }
{$IFNDEF T2H}

type
  TDefaultReaderWriter = class
    DefaultWriterANSI: string;
    DefaultWriterRTF: string;
    DefaultWriterHTML: string;
    DefaultReaderANSI: string;
    DefaultReaderRTF: string;
    DefaultReaderHTML: string;
  end;
  {$ENDIF}

var
  WPDefaultCharset: Integer;   //john
  GlobalLanguage: Integer;
  BulletChar: Char;
  BulletFont: string;
  wpClWindowText: TColor;
  wpMisspelledTextColor: TColor;
  WPWordDelimiterArray: array[Char] of Boolean;
  WPTPrinter: TPrinter;
  {$IFNDEF T2H}
  DefaultFontSize: Integer;
  DefaultDecimalTabChar: Char;
  DefaultReaderWriter: TDefaultReaderWriter;
  WPReadFromFiler: Boolean;
  UnitSelection: array[0..5] of Char; { was: string = cm, inch, point,	twips }
  WPPrinterCanvas: TCanvas;
  WPPrinterName: string;
  PrinterCanvasPX, PrinterCanvasPY: Integer;
  WPPrinterHDC: HDC;
  FNoPrinterInstalled: Boolean;
  FWPPrinterCanvasHasHandle: Boolean;
  {$ENDIF}

const
  {$IFNDEF T2H}
  WPCustomTextReaderReadBufferLen = 8192;
  WMWP_UPDATE = WM_USER + 5;
  WMWP_ZOOMUPDATE = WM_USER + 6;
  WMWP_UPDATE_MODE_CHAR = 1; { Char Attributes were changed }
  WMWP_UPDATE_MODE_PAR = 2; { Par Attributes were changed }
  WMWP_UPDATE_MODE_Edit = 4; { Clipboard Attributes were changed }
  WMWP_UPDATE_MODE_RULEROFFSET = 8; { leftoffset was changed }

  WPI_IDX_INTERN = 0;
  WPI_GR_USER0 = $0000;

  WPI_GR_STYLE = $0100; { allow	all up }
  WPI_GR_ALIGN = $0200; { radion button	}
  WPI_GR_EDIT = $0300; { copy,	paste ... }
  WPI_GR_DISK = $0400; { new,open,save	... }
  WPI_GR_PRINT = $0500; { print, print setup }
  WPI_GR_DATA = $0600; { prev,	next ... }
  WPI_GR_PARAGRAPH = $0700;
  WPI_GR_TABLE = $0800; { table	tools }
  WPI_GR_OUTLINE = $0900; { Outline tools	}
  WPI_GR_EXTRA = $0A00; // Extra, like InsertNumber etc

  WPI_GR_DROPDOWN = $0B00; { Comboboxes	}
  WPI_GR_USER1 = $0C00; { User definied	group 2	}
  WPI_GR_USER2 = $0D00; { User definied	group 3	}
  WPI_GR_USER3 = $0E00; { User definied	group 4	}

  { 1. Commands }
  WPI_CO_Normal = 1; {	Group: WPI_GR_STYLE }
  WPI_CO_Bold = 2;
  WPI_CO_Italic = 3;
  WPI_CO_Under = 4;
  WPI_CO_Hyperlink = 5;
  WPI_CO_StrikeOut = 6;
  WPI_CO_SUPER = 7;
  WPI_CO_SUB = 8;
  WPI_CO_HIDDEN = 9;
  WPI_CO_RTFCODE = 10;
  WPI_CO_PROTECTED = 11;
  WPI_CO_UPPERCASE = 12; // afsUpperCaseStyle

  WPI_CO_Left = 1; { Group: WPI_GR_ALIGN }
  WPI_CO_Right = 2;
  WPI_CO_Justified = 3;
  WPI_CO_Center = 4;

  WPI_CO_Copy = 1; { Group WPI_GR_EDIT }
  WPI_CO_Cut = 2;
  WPI_CO_Paste = 3;
  WPI_CO_SelAll = 4;
  WPI_CO_HideSel = 5;
  WPI_CO_Find = 6;
  WPI_CO_Replace = 7;
  WPI_CO_SpellCheck = 8;
  WPI_CO_Undo = 9;
  WPI_CO_FindNext = 10;
  WPI_CO_DeleteText = 11;
  WPI_CO_Redo = 12;
  WPI_CO_SpellAsYouGo = 13;
  WPI_CO_Thesaurus = 14;

  WPI_CO_Exit = 1; { Group: WPI_GR_DISK }
  WPI_CO_New = 2;
  WPI_CO_Open = 3;
  WPI_CO_Save = 4;
  WPI_CO_Close = 5;

  WPI_CO_Print = 1; { Group: WPI_PRINT }
  WPI_CO_PrintSetup = 2;
  WPI_CO_FitWidth = 3; { AutoZoomWidth }
  WPI_CO_FitHeight = 4; { AutoZoomHeight }
  WPI_CO_ZoomIn = 5;
  WPI_CO_ZoomOut = 6;
  WPI_CO_NextPage = 7;
  WPI_CO_PriorPage = 8;
  WPI_CO_PrintDialog = 9;

  WPI_CO_Next = 1; { Group: WPI_DATA }
  WPI_CO_Prev = 2;
  WPI_CO_Add = 3;
  WPI_CO_Del = 4;
  WPI_CO_Edit = 5;
  WPI_CO_Cancel = 6;
  WPI_CO_ToStart = 7;
  WPI_CO_ToEnd = 8;
  WPI_CO_Post = 9;

  WPI_CO_PARPROTECT = 1; { Group WPI_GR_PARAGRAPH }
  WPI_CO_PARKEEP = 2;

  WPI_CO_CreateTable = 1; { Group: WPI_TABLE }
  WPI_CO_BAllOff = 2;
  WPI_CO_BLeft = 3;
  WPI_CO_BTop = 4;
  WPI_CO_BBottom = 5;
  WPI_CO_BRight = 6;
  WPI_CO_SelRow = 7;
  WPI_CO_InsRow = 8;
  WPI_CO_DelRow = 9;
  WPI_CO_DelCol = 10;
  WPI_CO_InsCol = 11;
  WPI_CO_SelCol = 12;
  WPI_CO_SplitCell = 13;
  WPI_CO_CombineCell = 14;
  WPI_CO_BAllOn = 15;
  WPI_CO_BInner = 16;
  WPI_CO_BOuter = 17;


  WPI_CO_Bullets = 1; {Group: WPI_GR_OUTLINE }
  WPI_CO_Numbers = 2;
  WPI_CO_NextLevel = 3; // Only for numbers
  WPI_CO_PriorLevel = 4;

  WPI_CO_IsOutline = 5; // Include in TOC
  WPI_CO_OutlineUp = 6; // Outline Level
  WPI_CO_OutlineDown = 7;



  WPI_CO_InsertNumber = 1; { Group: WPI_GR_EXTRA}
  WPI_CO_InsertField = 2;
  WPI_CO_EditHyperlink = 3;
  {$ENDIF}
  { ++++ Other const }

  NumPaletteEntries = 255;
  FontMaxAnz = 255;
  {$IFNDEF T2H}
  WPPictAspectRatio = -1; { Consts for	PicInsert }
  WPUseNormalValue = 0;
  TSpellCheckHyphenCount = 8;
  KeepOldValue = -1000;
  TABMAX = 64; { how may bits in TTabFlag	}
  TABWORDMAX = 4; { how may words in TTabFlag	}

  TABCODE = #9;
  NEWLINECODE = #10;
  SPACECODE = #32;
  MaxCellAnz = 255; { Max count of columns in tables }

  TabsAllWord = 65535;

  TextObjCode = #1; // Marks the a signular object or code
  TextObjCodeOpening = #2; // Markes a start code
  TextObjCodeClosing = #3; // Markes an end code

  MaxWPPageSettings = 7;
  {$ENDIF}

type
  {$IFNDEF T2H}
  TWPPageDef = record
    ID: Integer;
    UseCM: Boolean;
    Width, Height: Single;
    Name: string;
  end;

  TWPStyleColors = record
    name: string;
    value: TColor;
  end;

const
  WPSSStyleColors: array[0..15] of TWPStyleColors =
  ((name: 'aqua'; value: clAqua),
    (name: 'navy'; value: clNavy),
    (name: 'black'; value: clBlack),
    (name: 'blue'; value: clBlue),
    (name: 'olive'; value: clOlive),
    (name: 'purple'; value: clPurple),
    (name: 'fuchsia'; value: clFuchsia),
    (name: 'red'; value: clRed),
    (name: 'gray'; value: clGray),
    (name: 'silver'; value: clSilver),
    (name: 'green'; value: clGreen),
    (name: 'teal'; value: clTeal),
    (name: 'lime'; value: clLime),
    (name: 'white'; value: clWhite),
    (name: 'maroon'; value: clMaroon),
    (name: 'yellow'; value: clYellow));

  {$IFNDEF DELPHI3ANDUP}DMPAPER_A6 = 1000; {$ENDIF}
  WPPageSettings: array[0..MaxWPPageSettings] of TWPPageDef =
  ((ID: - 1; UseCM: TRUE; Width: 21; Height: 29; Name: 'Custom'),
    (ID: DMPAPER_LETTER; UseCM: FALSE; Width: 8.5; Height: 11.0; Name:
    'Letter 8 1/2 x 11 in'),
    (ID: DMPAPER_LEGAL; UseCM: FALSE; Width: 8.5; Height: 14.0; Name:
    'Legal 8 1/2 x 14 in'),
    (ID: DMPAPER_EXECUTIVE; UseCM: FALSE; Width: 7.25; Height: 10.5; Name:
    'Executive 7.25 x 10.5 in'),
    (ID: DMPAPER_A3; UseCM: TRUE; Width: 29.7; Height: 42.0; Name:
    'DIN A3 297 x 420 mm'),
    (ID: DMPAPER_A4; UseCM: TRUE; Width: 21.0; Height: 29.7; Name:
    'DIN A4 210 x 297 mm'),
    (ID: DMPAPER_A5; UseCM: TRUE; Width: 14.8; Height: 21.0; Name:
    'DIN A5 148 x 210 mm'),
    (ID: DMPAPER_A6; UseCM: TRUE; Width: 10.5; Height: 14.8; Name:
    'DIN A6 105 x 148 mm')
    );
  {$ENDIF}

type
  TWPPageSettings = (wp_Custom, wp_Letter, wp_Legal, wp_Executive,
    wp_DinA3, wp_DinA4, wp_DinA5, wp_DinA6);

  TWPPrintPageMode = (ppmUseBorders, ppmPrintFrame, ppmStretchWidth,
    ppmStretchHeight, ppmIgnoreSelection, ppmDrawSmallAsBlock, ppmUseEvents,
    ppmInternPrintPageOnly);
  TPrintPageMode = set of TWPPrintPageMode;

  TWPInsertPointOptions = set of
    (
    // temporarily only
    wpppStart, wpppEnd, wpppEditor,
    // -----------------------------
    wpppOneLine, // What is allowed?  (No NL, no CR)
    wpppOnePar, // no CR
    wpppMaxLength,
    wpppInteger,
    wpppFloat,
    wpppDate,
    wpppCurrency,
    wpppLowerCase,
    wpppUpperCase,
    wpppCustomCheck,
    wpppNoCharAttrChange, // both means: noRTF !!!
    wpppNoParAttrChange,
    wpppNoEmbeddedTags // can include tags
    );

  TWPInsertPointProps = record
    Options: TWPInsertPointOptions;
    MaxLength: Integer;
    DisableEdit: Boolean;
    DisableMerge: Boolean;
  end;

  TWPUnit = (UnitCm, UnitIn, UnitPt, UnitPix);
  TWPWorkOnOptions = (wpBody, wpHeader, wpFooter);
  {$IFDEF T2H}
  TWPViewOptions = set of
    {$ELSE}
  TWPViewOption =
    {$ENDIF}
  (wpUseDoubleBuffer, wpGrayTables, wpShowGridlines,
    wpShowCR, wpShowNL, wpShowSPC, wpShowHardSPC, wpShowTAB, wpHideObjectFrame,
    wpShowObjectAnchors,
    wpShowBookmarkCodes,
    wpShowEmbeddedCodes,
    wpShowReferenceObjects,
    wpShowInvisibleText,
    wpShowPageNRinGap,
    wpDrawFineUnderlines,
    wpCenterLinesVertically,
    wpUseAbsoluteFontHeight,
    wpDontGrayHeaderFooterInLayout,
    wpThinObjectFrames,
    wpNoTransparentFloatingObjects,
    wpNoEndOfDocumentLine,
    wpHideSelection,
    wpDontIgnoreTabsInRightMargin,
    wpUseWatermark,
    wpDontAutohideBorders,
    wpWriteRightToLeft // always, not only in paprRightToLeft paragraphs!
    );
  TWPPrintHeaderFooter = (wprNever, wprOnAllPages, wprNotOnFirstPage,
    wprOnlyOnFirstPage);
  {$IFNDEF T2H}
  TWPViewOptions = set of TWPViewOption;
  TWPEditOption =
    {$ELSE}
  TWPEditOptions = set of
    {$ENDIF}
  (wpTableResizing, { left right indent }
    wpTableOneCellResizing, { always only one cell, like ssCtrl in	Shift }
    wpTableColumnResizing, { column if possible }
    wpObjectMoving, { only floating objects }
    wpObjectResizingWidth, { all objects }
    wpObjectResizingHeight,
    wpObjectResizingKeepRatio, { only both, width AND height  }
    wpObjectSelecting, { all objects }
    wpObjectDeletion, { DEL when object is selected }
    wpFieldObjectsAsGraphicObjects, { work with TWPOFieldObject as if they were TWPOGraphics }
    wpSpreadsheetCursorMovement, { Cursor up/down in Rows }
    wpAutoInsertRow, { wpAutoInsertRow,  TAB in last cell }
    wpNoEditOutsideTable, { to simulate spreadsheet	}
    wpActivateUndo, { activate UNDO }
    wpActivateUndoHotkey, {	activate ALT + Backspace - requires wpAllowUndo	set too	}
    wpActivateRedo, { makes backup of complete text ! }
    wpActivateRedoHotkey, { makes backup of complete text ! }
    wpAlwaysInsert, { don't switch to	overwrite }
    wpDeactivateCharsetSwitching,
    wpMoveCPOnPageUpDown, { Move Cursor on Page up or Down code }
    wpAutoDetectHyperlinks,
    wpNoHorzScrolling, wpNoVertScrolling,
    wpDontJoinParsWithDifferentStyles, { invented by KJS }
    wpNoAutomaticHangingBulletsAndNumbers,
    wpMDIDragAndDrop,
    wpDontDeleteExtraSpace,
    wpUseHyphenation,
    wpDontSelectCompleteField,
    wpStreamUndoOperation,   // saves also additional info like bands, objects ..
    wpToolBarDisableDifferentFontsInSelections, // Sets Size,Color, Font to blank is not = in a selection
    wpTabToEditFields // used with ProtecteProp: ppAllExceptForEditFields
  );

  {$IFNDEF T2H}
  TWPEditOptions = set of TWPEditOption;
  TWPPrintOption =
    {$ELSE}
  TWPPrintOptions = set of
    {$ENDIF}
  (wpIgnoreBorders,
    wpIgnoreShading,
    wpIgnoreText,
    wpIgnoreGraphics,
    wpDoNotChangePrinterDefaults,
    wpDontPrintWatermark,
    wpDontReadPapernamesFromPrinter
    );

  TWPRTFEngineOption =
  (
    wpAutoContinueAllNumbering
  );

  TWPRTFEngineOptions = set of TWPRTFEngineOption;

  {$IFNDEF T2H}
  TWPPrintOptions = set of TWPPrintOption;
  TWPEditBoxMode =
    {$ELSE}
  TWPEditBoxModes = set of
    {$ENDIF}
  (wpemLimitTextHeight,
    wpemLimitTextWidth, { equivalent to 'WordWrap! }
    wpemAutoSizeWidth,
    wpemAutoSizeHeight,
    wpemCheckCursorMovement,
    wpemCheckTextFlow);

  {$IFNDEF T2H}
  TWPEditBoxModes = set of TWPEditBoxMode;
  {$ENDIF}

  TWPChangeBoxEvent = procedure(Sender: TObject;
    var NewWidth, NewHeight: Integer; var Change: Boolean) of object;

  {$IFDEF T2H}
  TNotifyEvent = procedure(Sender: TObject) of object;
  {$ENDIF}

  TWPOnGetPageGapText = function(Sender: TObject; PageNumber: Integer): string
    of
    object;

  TWPCursorLeavesBoxEvent = procedure(Sender: TObject;
    up: Boolean; var NextBox: TObject; ch: Char) of object;

  TWPFlowTextOutOfBox = procedure(Sender: TObject; Stream: TMemoryStream) of
    object;
  TWPFlowTextIntoBox = procedure(Sender: TObject; var Source: TObject) of
    object;

  TWPPrintMode = (wpPrintCalc, wpNormalPrint, wpFastPrintInit, wpFastPrint,
    wpFastPrintExit);

  TWPPrintEvent = procedure(Sender: TObject;
    prCanvas: TCanvas; { where you shoud draw }
    x, y, w, h: Integer; { Where you may draw    }
    PageNumber: Integer; { Some useful Information }
    LineNumber: Longint
    ) of object;

  THyperLinkEvent = procedure(Sender: TObject;
    text: string; { marked Text }
    stamp: string; { hidden text }
    LineNumber: Longint) of object;

  TWPGetFieldTextEvent = function(Sender: TObject; Field: TObject; var done:
    Boolean): string of object;
  TWPGetFieldBoundsEvent = procedure(Sender: TObject; Field: TObject;
    var width, height, base: Longint) of object;

  TMeasurePageEvent = procedure(Sender: TObject;
    PageNr: Longint;
    var PageH, MargT, MargB: Longint) of object;

  TTextObjectCheckPropertiesEvent = procedure(var Sender: TObject; var CallAgain:
    Boolean) of object;

  TWPTextObjectCheckPropertiesEvent = procedure(Sender: TObject; var CallAgain:
    Boolean) of object;

  TWPStatusItemCont = (stGauge, stHelp, stIndex, stOutline, stStatus, stInsert,
    stCaps, stNum, stXPos, stYPos, stDate, stTime, stPage, stLine, stChanged,
    stFont, stStyle, stUnit, stTxtFmt,
    stName, stFileName, stPathName, stString);

  TWPSetGaugeValueEvent = procedure(Sender: TObject; Value: Integer) of object;
  TWPSetStatusStringEvent = procedure(Sender: TObject; typ: TWPStatusItemCont;
    const Text, Value: string) of object;

  { notify the application that a severe problem occur }
  TWPError = (wpeMemoryCorrupt, wpeOutOfMemory, wpeTooManyTabs);
  TWPErrorEvent = procedure(Sender: TObject; error: TWPError) of object;

  THTMLBitmapRequestEvent = procedure(Sender: TObject; const SRC: string;
    var Bitmap: TGraphic) of object;

  THTMLCreateImageEvent = procedure(Sender, WPObject: TObject;
    const path: string;
    var filename: string;
    Bitmap: TGraphic) of object;

  TStoreHTMLComment = (shcFootnote, shcHiddenFootnote, shcField, shcOFF);
  TWPHTMLCreateImages = (wpciDontCreateImages, wpciCreateImages, wpciCreateImageEvent);


  TObjectClickEvent = procedure(var Sender: TObject;
    x, y: Integer; { Mouse position }
    Tag, AutomaticTag: Integer; { id	in the text }
    var w_twips, h_twips: Longint; { size of box    }
    var changed: Boolean) of object;

  TProtectedChangeEvent = procedure(Sender: TObject;
    const Text: string; { Text in the line	which was about	to be changed }
    Tag: Integer) of object;

  {$IFDEF T2H}
  TWPMailMergeContinueOptions = set of
    (mmStopNow, mmMergeTrim, mmIgnoreNewParAttr, mmIgnoreAllNewParAttr,
    mmMergeAsRTF, mmMergeAsHTML, mmSkipSpacesBehind, mmInsertObject,
    mmDeleteEmptyParagraph, mmDeleteUntilFieldEnd
    );
  {$ELSE}
  TWPMailMergeContinueOption = (mmStopNow,
    { now used insetad of DoContinue := FALSE }
    mmMergeTrim, { trim white spaces from p	}
    mmIgnoreNewParAttr, { when RTF	or HTML	is inserted use	old Par	attr of	the current line }
    mmIgnoreAllNewParAttr, { like mmIgnoreNewParAttr but in the complete text }
    mmMergeAsRTF, { RTF is automatically detected - but you can still force it }
    mmMergeAsHTML, { HTML is only detected when starting with	'<html>'. Better force it }
    mmSkipSpacesBehind, { skip white spaces after this field (usefull if p='' }
    mmInsertObject, { insert object from new obj Reference.
    Can be a	TWPObject or a TCustomRTFEdit, too! }
    mmDeleteEmptyParagraph,
    mmDeleteUntilFieldEnd, mmNoChange
    );
  TWPMailMergeContinueOptions = set of TWPMailMergeContinueOption;
  {$ENDIF}

  TSpellCheckResult = set of (spIgnored, spMisSpelled, spHyphenate);
  TSpellCheckHyphen = array[1..TSpellCheckHyphenCount] of Byte;

  { Obsolte - only provided for compatibility }
  TSpellCheckWordEvent = procedure(Sender: TObject;
    word: string;
    var res: TSpellCheckResult;
    var hypen: TSpellCheckHyphen
    ) of object;

  TWPHyphenateEvent = procedure(Sender: TObject;
    Word        : String;
    CharSet     : Integer;
    var hypen   : TSpellCheckHyphen;
    var hyphencount : Integer
    ) of object;


  TWPStartSpellcheckMode = (wpStartSpellCheck, wpStartThesuarus, wpStartSpellAsYouGo, wpStopSpellAsYouGo);

  TWPStartSpellcheckEvent = procedure(Sender: TObject; Mode: TWPStartSpellcheckMode) of object;

  { -- }

  TWPSpellCheckWordEvent = procedure(Sender: TObject;
    var word: string;
    const ParagraphText: string;
    const PosInPara: Integer;
    var res: TSpellCheckResult;
    var hypen: TSpellCheckHyphen
    ) of object;

  { state of clipboard has changed }
  TEditStateEvent = procedure(Sender: TObject;
    selection_marked: Boolean;
    clipboard_not_empty: Boolean
    ) of object;

  {$IFNDEF T2H}
  TWPVersion2CompMode = set of (wp2NoAutoIndentTabs);

  { Set	the way	how RichText AND Normal	Text is	saved and loaded.
    You	can specify the	old values or you can use the classname
    of the reader/Writer object	 }
  TWPLoadFormat = (AutomaticFormat, LastFormat, FormatRTF, FormatANSI,
    FormatHTML);
  TWPNewLoadFormat = string;

  {------ Ruler	----------------------------- --------------
   All Parameters for indenting	and page size are stored in
   one Record wich is accesible	for the	ruler. This does not
   mean	that you need a	ruler to change	the values! .........}

  TTabFlag = array[0..TABWORDMAX - 1] of Word;

  TWPLayout = record
    paperw: Longint;
    paperh: Longint;
    margl, margr: Longint;
    margt, margb: Longint;
    deftabstop: Longint;
    marg_header: Longint; { Distance Header to page border }
    marg_footer: Longint;
    landscape: Boolean; { If true then first 6 values are swapped }
    hasTitlePg : Boolean; //V4.07 - support for the /titlepg  RTF tag to use a first header
    hasFacingP : Boolean; //V4.07 - support for the /titlepg  RTF tag to use a first header
    marginmirror: Boolean; // Mirror Pages
  end;
  {$ENDIF}

  TWPMoveMode = (wpmLStart, wpmLEnd, wpmPagUp, wpmPagDown,
    wpmHome, wpmEnd, wpmCLeft, wpmCRight, wpmCUp, wpmCDown,
    wpmWLeft, wpmWRight, wpmTopLine, wpmBottomLine,
    wpmEndOfSelection, wpmStartOfSelection,
    wpmFieldStart, wpmFieldEnd);

  TWPWriteBitmapMode = (wbmBMPFile, wbmWBitmap, wbmDIBBitmap);

  TWPObjectWriteRTFModeGlobal = (wobDependingOnObject, wobStandard, wobRTF, wobStandardAndRTF);

  { You	may select these Icons in tree groups }
  {$IFNDEF T2H}
  TWpTbIcon = (SelNormal, SelBold, SelItalic, SelUnder, SelHyperLink,
    SelStrikeOut, SelSuper, SelSub, SelHidden, SelProtected,
    SelRTFCode,
    SelLeft, SelRight, SelBlock, SelCenter);

  TWpTbIcon2 = (SelExit, SelNew, SelOpen, SelSave, SelClose,
    SelPrint, SelPrintSetup, SelFitWidth, SelFitHeight,
    SelZoomIn, SelZoomOut, SelNextPage, SelPriorPage);

  TWPTbIcon3 = (SelToStart, SelNext, SelPrev, SelToEnd, SelEdit, SelAdd, SelDel,
    SelCancel, SelPost);

  TWPTbIcon4 = (SelUndo, SelRedo, SelDeleteText, SelCopy, SelCut, SelPaste, SelSelAll, SelHideSel,
    SelFind, SelReplace, SelSpellCheck);

  TWPTbIcon5 = (SelCreateTable, SelSelRow, SelInsRow, SelDelRow, SelInsCol,
    SelDelCol, SelSelCol,
    SelSplitCell, SelCombineCell,
    SelBAllOff, SelBAllOn, SelBInner, SelBOuter
    , SelBLeft, SelBRight, SelBTop, SelBBottom);

  TWPTbIcon6 = (SelBullets, SelNumbers, SelNextLevel, SelPriorLevel,
    SelParProtect, SelParKeep);

  TwpTbIcons = set of TWpTbIcon;

  TwpTbIcons2 = set of TWpTbIcon2;

  TwpTbIcons3 = set of TWpTbIcon3;

  TwpTbIcons4 = set of TWpTbIcon4;

  TwpTbIcons5 = set of TWpTbIcon5;

  TwpTbIcons6 = set of TWpTbIcon6;

  TWpTbListbox = (SelFontName, SelFontSize, SelFontColor, SelBackgroundColor,
    SelStyle, SelParColor);

  TwpTbListboxen = set of TWpTbListbox;
  {$ELSE}
  TWpTbIcons = set of (SelNormal, SelBold, SelItalic, SelUnder, SelHyperLink,
    SelStrikeOut, SelSuper, SelSub, SelHidden, SelProtected,
    SelRTFCode,
    SelLeft, SelRight, SelBlock, SelCenter);

  TWpTbIcons2 = set of (SelExit, SelNew, SelOpen, SelSave, SelClose,
    SelPrint, SelPrintSetup, SelFitWidth, SelFitHeight,
    SelZoomIn, SelZoomOut, SelNextPage, SelPriorPage);

  TWPTbIcons3 = set of (SelToStart, SelNext, SelPrev, SelToEnd, SelEdit, SelAdd, SelDel,
    SelCancel, SelPost);

  TWPTbIcons4 = set of (SelUndo, SelCopy, SelCut, SelPaste, SelSelAll, SelHideSel,
    SelFind, SelReplace, SelSpellCheck);

  TWPTbIcons5 = set of (SelCreateTable, SelSelRow, SelInsRow, SelDelRow, SelInsCol,
    SelDelCol, SelSelCol,
    SelSplitCell, SelCombineCell,
    SelBAllOff, SelBAllOn, SelBInner, SelBOuter
    , SelBLeft, SelBRight, SelBTop, SelBBottom);

  TWPTbIcons6 = set of (SelBullets, SelNumbers, SelNextLevel, SelPriorLevel,
    SelParProtect, SelParKeep);

  TWpTbListboxen = set of (SelFontName, SelFontSize, SelFontColor, SelBackgroundColor,
    SelStyle, SelParColor);
  {$ENDIF}

  TWpSelNr = (wptNone, wptName, wptSize, wptColor, wptBkColor,
    wptTyp, wptIconSel, wptIconDeSel, wptPage, wptParagraph, wptParColor,
    wptParAlign, wptStyleNames);

  TWPSelectEvent = procedure(Sender: TObject;
    var Typ: TWpSelNr;
    const str: string;
    const num: Integer) of object;

  TWPIconSelectEvent = procedure(Sender: TObject;
    var Typ: TWpSelNr;
    const str: string;
    const group: Integer;
    const num: Integer;
    const index: Integer) of object;

  {$IFNDEF T2H}
  PTColorPalette = ^TPaletteEntry;
  TColorPalette = array[0..NumPaletteEntries] of TPaletteEntry;
  {$ENDIF}

  // TWPNumberStylesList holds the definitions of the numbering and bullet styles
  TWPNumberStyle = (wp_none, wp_bullet, wp_lg_a, wp_a, wp_lg_i, wp_i, wp_1);

  {$IFDEF T2H}
  WrtStyle = set of (afsBold, afsItalic, afsHyperLink, afsSuper, afsSub, afsProtected,
    afsIsFootnote,
    afsIsInsertpoint, afsAutomatic, afsUnderline, afsStrikeOut,
    afsHidden, afsIsRTFCode, afsUppercaseStyle, afsParStyleXXX);
  {$ELSE}
  TOneWrtStyle = (
    afsBold,
    afsItalic,
    afsHyperLink,
    afsSuper,
    afsSub,
    afsProtected,
    afsIsFootnote,
    afsIsInsertpoint,
    afsAutomatic,
    afsUnderline,
    afsStrikeOut,
    afsHidden,
    afsIsRTFCode, //?
    afsDelete,
    afsUserdefined, // used for edit fields
    afsBookmark, // Temporarily assigned to text inbetween bookmark markers (not 100% reliable)
    afsMisSpelled,
    afsIsObject,
    afsHyphen, // May hyphenate here (before this character!)
    afsWasChecked,
    afsUppercaseStyle, // Print all as ABCDEF
    afsHotStyle, // temporarily  assigned
    // In combination with special attributes, such as bookmark or hyperlink
    // Style support - if bit is set the value can be overridden if a
    // style is reassigned to the paragraph. If the user changes attributes
    // using the toolbar the bit is deleted.
    afsParStyleBold,
    afsParStyleItalic,
    afsParStyleUnderline,
    afsParStyleFont,
    afsParStyleSize,
    afsParStyleColor,
    afsParStyleBGColor
    {$IFDEF WPDBCS}, afsDBCSOne, afsDBCSTwo{$ENDIF}
    );
  WrtStyle = set of TOneWrtStyle;
  {$ENDIF}

  TProtectProps = (ppProtected, ppHidden, ppIsInsertpoint,
    ppAutomatic,
    ppIsInvisible, ppParProtected,
    ppProtectSelectedTextToo, ppDontCopyProtectedAttribute,
    ppDontUseAttrOfProtected, ppNoEditAfterProtection,
    ppInsertBetweenProtectedPar,
    ppAllowEditAtTextEnd,
    ppAllExceptForEditFields);
  TProtectProp = set of TProtectProps;

  // ---------------------------------------------------------------------------
  // Style support - if bit is set the value can be overridden if a
  // style is reassignment to the paragraph. If the user changes attributes
  // using the toolbar or ruler the bit is deleted.
  // It is also used in the Style object to specify which item was defined
  TWPStyleElementPar = (wpseIndentLeft, wpseIndentRight, wpseIndentFirst,
    wpseSpaceBefore, wpseMultSpaceBetween, wpseSpaceBetween, wpseSpaceAfter,
    wpseBrdLines, wpseBrdWidth, wpseBrdColor,
    wpseAlign, wpseColor, wpseShading, wpseTabs,
    wpseNumber, wpseNumberLevel, wpseIsOutline, wpseBorderSpacing, wpseKeepTogether, wpseKeepNext);
  TWPStyleElementsPar = set of TWPStyleElementPar;

  TWPStyleElementChar = // attention: See also unit WPStyCol !!!
  (wpseFontBold,
    wpseFontItalic,
    wpseFontUnderline,
    wpseFontFont,
    wpseFontSize,
    wpseFontColor,
    wpseFontBGColor);
  TWPStyleElementsChar = set of TWPStyleElementChar;
  // ---------------------------------------------------------------------------

  {$IFNDEF T2H}
  TDataType = (dtChar, dtObject, dtFootnote);
  {$ENDIF}

  TWPExtraStyleDisplayTypes = (wpstBlock, wpstInline, wpstList, wpstNone);

  PWPExtraStyleProps = ^TWPExtraStyleProps;
  TWPExtraStyleProps = record
    Display: TWPExtraStyleDisplayTypes;
    IsDefault: Boolean;
    BackgrImage: String[127];
    // NumberStyle Props:
    nStyle : TWPNumberStyle;
    nTextA : String[63];
    nTextB : String[63];
    nFont  : String[127];
  end;

  // This are the extra elements which are ignored
  TWPStyleElementExtra = (wpseIgnored, wpseDisplay, wpseIsDefault, wpseBackgroundImage);
  TWPStyleElementsExtra = set of TWPStyleElementExtra;


  PTAttr = ^TAttr;
  TAttr =
    packed record
    Style: WrtStyle;
    Tag: Integer;
    Width: Word;
    Base: Word; // We could calculate this on the fly before paint!
    Height: Word;
    Color: Byte;
    BGColor: Byte;
    Font: Byte;
    Extra: Byte; // UNICode Char with fmRichTextUnicode, otherwise 0 or Charset
    Size: Byte;
    cStyle: Byte; // Character (inline) Style
  end;

  {$IFNDEF T2H}
  ecErrCode = (ecOK, ecUserBreak, ecInvalidFile, (* too many	'}' in RTF *)
    ecInternError, { low memory ... }
    ecNoChangeAllowed, { readonly ... }
    ecEndOfFile, { ecEndOfFile will be created when a	writing	error occurs }
    ecWrongArgument);


  TWPChangeParOneTyp = (wpchIndentLeft, wpchIndentRight, wpchIndentFirst,
    wpchSpaceBefore, wpchMultSpaceBetween, wpchSpaceBetween, wpchSpaceAfter,
    wpchBrdLines,
    wpchBrdLinesVisible, wpchBrdWidth, wpchBrdStyle, wpchAlign, wpchVertAlign, wpchColor,
    wpchShading, wpchTabs,
    wpchNumber, wpchNumberLevel, wpchIncNumberLevel,
    wpchDecNumberLevel, wpchSetOutline, wpchDelOutline,
    wpchParID, wpchHide, wpchUnHide, wpchParWW,
    wpchParKeep, wpchParKeepN, wpchParProtected, wpchStyleNr, wpchBorderSpace,
    wpchOutlineBreak);

  TWPChangeParTyp = set of TWPChangeParOneTyp;

  TWPBrdColorDefChanged = set of (wpbrcColor, wpbrcVisible);
  TWPBrdColorDef = record
    Color: Integer;
    Visible: Boolean;
    Changed: TWPBrdColorDefChanged;
  end;
  TWPBrdLine = (brdLeft, brdRight, brdTop, brdBottom);

  TWPExtParOptions = record
    Border: array[TWPBrdLine] of TWPBrdColorDef;
  end;
  {$ELSE}
  TWPBrdLine = (brdLeft, brdRight, brdTop, brdBottom);
  {$ENDIF}

  TWPDBCSOneMode = (wpdb_Enabled, wpdb_OnlyNotANSIChar, wpdb_CalcCharWidth,
    wpdb_NoWordWrapAtSpace);
  TWPDBCSMode = set of TWPDBCSOneMode;

  {$IFNDEF T2H}
  TWPLayoutMargin = (wpmaTop, wpmaBottom, wpmaLeft,
    wpmaRight, wpmaHeader, wpmaFooter, wpmaPageGap,
    wpmaDrawFrame, wpmaOnlyManualPBreaks);
  TWPLayoutMargins = set of TWPLayoutMargin;
  {$ENDIF}

  TWPFormatOption = (wpfDontBreakTables, wpfDontBreakTableRows,
      wpfIgnoreKeep, wpfIgnoreKeepN,
      wpfAvoidWidows, wpfAvoidOrphans, wpfCenterOnPageVert,
      wpfIgnoreRightMarginAfterTabstop);

  TWPFormatOptions = set of TWPFormatOption;

  TWPLayoutMode = (wplayNormal, wplayPageGap, wplayLayout, wplayFullLayout, wplayShowManualPageBreaks);

  TWPEditUnit = (euTwips, euInch, euCm, euPt, euPercent, euMultiple, euStandard);
  TWPAutoZoom = (wpAutoZoomOff, wpAutoZoomWidth, wpAutoZoomFullPage);

  TTabKind = (tkLeft, tkRight, tkCenter, tkDecimal, tkAll);
  TWPScreenResMode = (rmNormal, rm1440);
  {$IFNDEF T2H}
  TOneStoreOption = (soNoColorTableInRTF, soNoPageFormatInRTF,
    soHardLineBreaks, soNoNULLChar,
    soNoNumberCompatibilityText, soOnlyNumberCompatibilityText,
    soSimpleNumberCompatibilityText,
    soNoOutlineStyleTable,
    soOldStyleOutlineStyleTable, // don't save \list ...
    soNoHeaderText,
    soNoFooterText, soWordwrapInformation, soPagebreakInformation,
    soNewParInTableCell,
    soNoShadingInformation, soNoStyleTable,
    soAllStylesInCollection,
    soANSIFormfeedCode, soWordHyperLinkFormat, soNoRTFVariables,
    soDontRenameToBAKFile, soDontDeleteBAKFile,
    soWriteObjectsAsRTFBinary
    {, soNoStyleRedundantAttribues});
  TStoreOptions = set of TOneStoreOption;

  TOneHTMLStoreOption =
    (soOnlyHTMLBody, soNoBasefontTag, soAllFontInformation,
    soNoSPANBackgroundColor, soOnlyFieldTags, soNoEntities, soCreateImageEventForAllImages, soAllUnicodeEntities
    );
  THTMLStoreOptions = set of TOneHTMLStoreOption;



  TOneLoadOption = (loIgnoreDefaultTabstop, loIgnorePageSize,
    loIgnoreHeaderFooter, loDontAddRTFVariables,
    loIgnorePageMargins, loIgnoreOutlineStyles, loIgnoreFonts, loIgnoreFontSize,
    soIgnoreWordwrapInformation, loIgnoreColorTable,
    loPrintHeaderFooterParameter,
    loApplyParagraphStylesAfterReading,
    loAddReadParagraphStyles, loDontAddReadParagraphStyleNames, loOverwriteParagraphStyleAttr,
    loDontLockAttrDifferentToStyle,
    loANSIFormfeedCode, loOnlyListText);
  TLoadOptions = set of TOneLoadOption;
  PLongint = ^Longint;
  {$ELSE}
  TStoreOptions = set of (soNoColorTableInRTF, soNoPageFormatInRTF,
    soHardLineBreaks, soNoNULLChar,
    soNoNumberCompatibilityText, soOnlyNumberCompatibilityText,
    soNoOutlineStyleTable, soNoHeaderText,
    soNoFooterText, soWordwrapInformation, soNewParInTableCell,
    soNoShadingInformation, soNoStyleTable
    );

  THTMLStoreOptions = set of (soOnlyHTMLBody, soNoBasefontTag, soAllFontInformation,
    soNoSPANBackgroundColor);

  TLoadOptions = set of (loIgnoreDefaultTabstop, loIgnorePageSize,
    loIgnoreHeaderFooter,
    loIgnorePageMargins, loIgnoreOutlineStyles, loIgnoreFonts, loIgnoreFontSize,
    soIgnoreWordwrapInformation, loIgnoreColorTable,
    loPrintHeaderFooterParameter,
    loApplyParagraphStylesAfterReading,
    loAddReadParagraphStyles, loDontAddReadParagraphStyleNames,
    loDontLockAttrDifferentToStyle);
  {$ENDIF}

  TTextHeader = class(TPersistent)
  private
    FDefaultPageWidth: Longint;
    FDefaultPageHeight: Longint;
    FDefaultTopMargin: Longint;
    FDefaultBottomMargin: Longint;
    FDefaultRightMargin: Longint;
    FDefaultLeftMargin: Longint;
    FDefaultLandscape: Boolean;
    FDefaultMarginHeader: Longint;
    FDefaultMarginFooter: Longint;
    FDefaultPageSize: TWPPageSettings;
    FRoundTabs: Boolean;
    FRoundTabsDivForInch: Integer;
    FRoundTabsDivForCm: Integer;
    FUnit: TWPUnit;
    FStoreOptions: TStoreOptions;
    FHTMLStoreOptions: THTMLStoreOptions;
    FLoadOptions: TLoadOptions;
    FLocked: Boolean;
    FPageSize: TWPPageSettings;
  private
    procedure SetPageWidth(x: Longint);
    function GetPageWidth: Longint;
    procedure SetPageHeight(x: Longint);
    function GetPageHeight: Longint;

    procedure SetLogPageWidth(x: Longint);
    function GetLogPageWidth: Longint;
    procedure SetLogPageHeight(x: Longint);
    function GetLogPageHeight: Longint;

    procedure SetMarginMirror(x: Boolean);
    function GetMarginMirror: Boolean;

    procedure SetLeftMargin(x: Longint);
    function GetLeftMargin: Longint;
    procedure SetRightMargin(x: Longint);
    function GetRightMargin: Longint;
    procedure SetTopMargin(x: Longint);
    function GetTopMargin: Longint;
    procedure SetBottomMargin(x: Longint);
    function GetBottomMargin: Longint;
    procedure SetTabDefault(x: Integer);
    function GetTabDefault: Integer;
    procedure SetDefaultPageSize(x: TWPPageSettings);
    procedure SetPageSize(x: TWPPageSettings);
    function GetLandscape: Boolean;
    function GetMarginHeader: Longint;
    function GetMarginFooter: Longint;
    procedure SetLandscape(x: Boolean);
    procedure SetMarginHeader(x: Longint);
    procedure SetMarginFooter(x: Longint);
    procedure SwapLayout;
  public
    destructor Destroy; override;
    constructor Create;
    procedure BeginUpdate;
    procedure EndUpdate;
    function AddFontName(nam: TFontName): Integer;
    function AddFontNameCharset(nam: TFontName; charset: Integer): Integer;
    procedure SetFontCharset(nr, charset: Integer);
    procedure Clear;
    procedure Assign(Source: TPersistent); override;
    {$IFNDEF T2H}
    procedure RecalcLayout;
    procedure CopyLayoutData(p: TTextHeader);
    function FindMatchingTab(pos: Longint; kind: TTabKind; map: TTabFlag):
      Integer;
    function FindTab(var pos: Longint; kind: TTabKind): Integer;
    function AddTab(pos: Longint; kind: TTabKind): Integer;
    function InternAddTab(pos: Longint; kind: TTabKind): Integer;
    function AllTabs: Longint;
    function RemoveTab(pos: Longint; kind: TTabKind): longint;
    procedure UpdatePageInfo(Reformat : Boolean);
    procedure ReCalculateTabs;
    procedure RemoveAllTabs;
    procedure SortTabs;
    function RoundTwips(tw: Longint): Longint;
    function FindMatchingColor(c: TColor; Use256Colors, AddColors: Boolean; var nr: Integer): Boolean;
    function AddColor(Color: TColor): Integer;
    {$ENDIF}
  public
    {$IFNDEF T2H}
    FUpdatePageInfo: procedure(MustReformat: Boolean) of object;
    FPaletteEntries: array[0..NumPaletteEntries] of TPaletteEntry;
    FXPixelsPerInch, FYPixelsPerInch: Integer;
    FFontXPixelsPerInch, FFontYPixelsPerInch: Integer;
    FPageGapWidth: Integer;
    FXLogPixelsPerInch: Integer;
    FXTwipsPerPixel: Integer;
    papWidth, papHeight: Integer;
    { Width and Height depending on current resolution }
    papTabDefault: Integer;
    FOnCheckTabs: procedure of object;
    IsRichText: Boolean;
    Fontname: array[0..FontMaxAnz] of TFontName;
    FontFamily: array[0..FontMaxAnz] of Integer;
    FontCharSet: array[0..FontMaxAnz] of Integer;
    RTFWriteTextOnly: Boolean;
    IXLoadPageNumber, IXPageCount: Integer;
    IXIgnore: Boolean;
    TabPosTw: array[0..TABMAX - 1] of Longint; { Twips	in any order! }
    TabPos: array[0..TABMAX - 1] of Integer; { Pixel }
    TabKind: array[0..TABMAX - 1] of TTabKind; { tkLeft, .... }
    TabIndex: array[0..TABMAX - 1] of Integer; { for sorting }
    ResultTABS: TTabFlag;
    {used	to return the result value. (arrays not	allowd in C++) }
    FDecimalTabChar: Char;
    FPrintHeaderFooterInPageMargins: Boolean;
    _Layout: TWPLayout;
    LayoutPIX: TWPLayout; { margp, as pixel values }
    Format: TWPNewLoadFormat;
    NewTabKind: TTabKind;
    FCurrTabs: TTabFlag;
    FDontRoundNextTab: Boolean;
    {$ENDIF}
  public
    property LogPageWidth: Longint read GetLogPageWidth write SetLogPageWidth;
    property LogPageHeight: Longint read GetLogPageHeight write SetLogPageHeight;
  published
    property StoreOptions: TStoreOptions read FStoreOptions write FStoreOptions;
    property HTMLStoreOptions: THTMLStoreOptions read FHTMLStoreOptions write FHTMLStoreOptions;
    property LoadOptions: TLoadOptions read FLoadOptions write FLoadOptions;
    property UsedUnit: TWPUnit read FUnit write FUnit;
    property RoundTabs: Boolean read FRoundTabs write FRoundTabs;
    property RoundTabsDivForInch: Integer read FRoundTabsDivForInch write
      FRoundTabsDivForInch;
    property RoundTabsDivForCm: Integer read FRoundTabsDivForCm write
      FRoundTabsDivForCm;
    property DecimalTabChar: Char read FDecimalTabChar write FDecimalTabChar;
    property PrintHeaderFooterInPageMargins: Boolean
      read FPrintHeaderFooterInPageMargins
      write FPrintHeaderFooterInPageMargins;

    property DefaultTabstop: Integer read GetTabDefault write SetTabDefault;
    {	Default	values.	Only set when the text is cleared }
    property DefaultPageWidth: Longint read FDefaultPageWidth write
      FDefaultPageWidth;
    property DefaultPageHeight: Longint read FDefaultPageHeight write
      FDefaultPageHeight;
    property DefaultTopMargin: Longint read FDefaultTopMargin write
      FDefaultTopMargin;
    property DefaultBottomMargin: Longint read FDefaultBottomMargin write
      FDefaultBottomMargin;
    property DefaultLeftMargin: Longint read FDefaultLeftMargin write
      FDefaultLeftMargin;
    property DefaultRightMargin: Longint read FDefaultRightMargin write
      FDefaultRightMargin;
    property DefaultMarginHeader: Longint read FDefaultMarginHeader write
      FDefaultMarginHeader;
    property DefaultMarginFooter: Longint read FDefaultMarginFooter write
      FDefaultMarginFooter;
    property DefaultLandscape: Boolean read FDefaultLandscape write
      FDefaultLandscape; { apply FIRST! }

    {	Current	values }
    property MarginMirror : Boolean read GetMarginMirror write SetMarginMirror;
    property PageWidth: Longint read GetPageWidth write SetPageWidth;
    property PageHeight: Longint read GetPageHeight write SetPageHeight;
    property LeftMargin: Longint read GetLeftMargin write SetLeftMargin;
    property RightMargin: Longint read GetRightMargin write SetRightMargin;
    property TopMargin: Longint read GetTopMargin write SetTopMargin;
    property BottomMargin: Longint read GetBottomMargin write SetBottomMargin;
    property MarginHeader: Longint read GetMarginHeader write SetMarginHeader;
    property MarginFooter: Longint read GetMarginFooter write SetMarginFooter;
    property Landscape: Boolean read GetLandscape write SetLandscape;
    { apply FIRST! }

    property DefaultPageSize: TWPPageSettings read FDefaultPageSize write
      SetDefaultPageSize;
    property PageSize: TWPPageSettings read FPageSize write SetPageSize;
  end;
  {$IFNDEF T2H}
  TWPSetOneCellProp = (wpscSetName, wpscSetCom, wpscSetFormat);
  TWPSetCellProp = set of TWPSetOneCellProp;
  {$ENDIF}

  TLinState = set of (listMustPaint, listAutoNewPage, listUse_spcb, listUse_spca,
    listHypenate, listHasObjects, listIgnoreRightIndent,
    listHasWindows, listPreformatted, listIsSelected, listIsInverted, listWasProcessed,
    listHasHotStyle, listWasChanged, listTemporaryHidden, listAutomatic, listPaintHeader, listPaintFooter,
    listIsFirstLine, listIsLastLine {on page});

  PTLine = ^TLine;
  {$IFDEF T2H}
  TLine = packed record
    next: PTLine;
    prev: PTLine;
    pc: PChar;
    pa: PTAttr;
    y_start: LongInt; // Start position in pixels (from beggining of text!)
    plen: Integer;
    Height: Word;
    Base: Word;
    pmax: Word;
  end;
  {$ELSE}
  TLine = packed record
    next: PTLine;
    prev: PTLine;
    pa: PTAttr;
    pc: PChar;
    y_start: Integer;
    plen: Word;
    Height: Word;
    Base: Word;
    lasttab, spacewid, spacerest: Word;
    pmax, pagenum: Word;
    t_xtra, b_xtra : Word;            // Space at top and bottom for PAGE margins
    spc_b, spc_a   : Word;            // Space included in 'Height' calculated uning Space_before, space_between and space_after
    line_top, line_bottom : Smallint; // Space above and after line for dynamic heights (table rows, vert align etc)
                                      // both have to be added to lin^.height !
    _lh: Word; { preformatted	text, twip }
    _rst: Smallint; { preformated - rest in pixels*1000 }
    state: TLinState;
  end;
  {$ENDIF}

  TThreeState = (tsIgnore, tsTRUE, tsFALSE);

  {$IFNDEF T2H}
  TCAttributeChangedEvent = procedure(what: WrtStyle; NeedReformat: Boolean) of
    object;
  {$ENDIF}



  {$IFNDEF T2H}
  TPicAlign = (oalBase, oalBottom, oalTop, oalMiddle); { Ignored in V3.0 }
  {$ENDIF}



  {record to create borders}
  {$IFDEF T2H}
  TBorderType = set of (BlEnabled, BlFinish, BlTop, BLBottom,
    BlLeft, BlRight, BLBox, BLDouble, BLDot, BLHair);
  {$ELSE}
  TOneBorderType = (BlEnabled, BlFinish, BlTop, BLBottom,
    BlLeft, BlRight, BLBox, BLDouble, BLDot, BLHair, BLBar,
    BlInnerEdge, BlOuterEdge, BlCenter, BlTabLines);
  TBorderType = set of TOneBorderType;
  {$ENDIF}
  PTBorder = ^TBorder;
  TBorder = packed record
    LineType: TBorderType;
    Thickness: Byte;
    HColor: Byte; // Color Left Border
    VColor: Byte; // Color Top Border
    HColorR: Byte; // Color Right Border
    VColorB: Byte; // Color Bottom Border
    Space: Byte; // space in pt to the border
  end;

  {$IFDEF T2H}
  TParProp = set of (paprMustReformat, paprNewPage, parIsProtected,
    paprMultSpacing,
    paprApplyJustified,
    paprIsSelected,
    paprKeep, paprKeepN,
    paprNoWordWrap,
    paprIsRightPar, paprIsLeftPar, paprIsTable, paprRightToLeft);
  {$ELSE}

  TOneParProp = (paprNewPage, parIsProtected,
    paprMustReformat,
    paprMultSpacing,
    paprApplyJustified,
    paprIsSelected,
    paprKeep, paprKeepN,
    paprNoWordWrap,
    paprIsRightPar, paprIsLeftPar, paprIsTable,
    paprIsOutline, { Outline, don't copy numberstyle to next line if Number}
    paprIgnoreParagraphLines, { User defined mark, abused for paprAlignToBottom }
    paprPreformat,
    paprHasFields, { Has Text Fields    }
    paprIsMergeCommandStart, { Merge Group Start }
    paprIsMergeCommandEnd, { Merge Group	End }
    paprHidden,
    paprWasProcessed,
    paprRightToLeft, // Write right to left in this par  /new in V4.11d)
    paprHasDescription, //?
    paprHasMarkedChar, // chars with afsDelete
    paprIsHeader, // repeat on top of next page
    paprHasGraphics, // At least one object in the lin/pa is aligned
    paprVertAlCenter, paprVertAlBottom, // alternatively !!
    paprNoRTF, // this paragraphs use always property of first char
    papxHidden, // -dynamic- Hidden by LEVEL COLLAPSE
    papxCollapsed, // Hide sublevel
    // Certain codes in this paragraph
    paprHasBookmark,
    paprOutlineBreak // Start new numbering
    );
  TParProp = set of TOneParProp;
  {$ENDIF}

  TParAlign = (paralLeft, paralCenter, paralRight, paralBlock);

  TParVertAlign = (paralVertTop, paralVertCenter, paralVertBottom); // paprVertAlCenter
  {$IFNDEF T2H}
  TParNumStyle = (pnsFinish, pnsNone, pnsBullet, pnsDigit, pnsLgLetter,
    pnsLetter, pnsLgRoman, pnsRoman);

  TWPControlCodeType = (wpctBookmarkStart, wpctBookmarkEnd, wpctGraphic);

  TWPControlCode = class(TObject)
  public
    typ: TWPControlCodeType;
  end;

  TParObject = class(TObject)
  public
    Owner: Pointer;
    Next: TParObject;
  end;

  PWPCell = ^TWPCell; // par^.celldef
  TWPCell = record
    { Definition of one	CELL }
    CWidthPC: Byte; { Width of Cell  [1/255]	}
    CLeftPC: Byte; { X-Offset  [1/255]		}
    ColSpan: Byte; { used	for HTML tables		}
    CellStyle: Byte; { ????				}
    Format: Integer; { open for extensions ! }
    Value: Double; { Value }
    Changed: Boolean; { Value was changed by 'Command' }
    { Definition of the	current	table }
    Tabletag: Longint;
    // Dynamic Paint Info
    pRight, pLeft, pBorderSpace: Integer;
    { special table properties }
    Cellname: string;
    { the same name can be assigned to multiple cells. This results in
    an area. Areas are automatically added and the cells are counted  (for avarage) }
    Command: string; { function for	number cruncher	}
  end;


  TWPInsertPointRec = record
    MaxLen: Integer;
    Options: Integer;
  end;
  {$ENDIF}

  { C++ Note: all lowercase in tparagraph }
  PTParagraph = ^TParagraph;
  {$IFDEF T2H}
  TParagraph = packed record
    indentfirst, indentleft, indentright: Smallint; // Indentation in twips
    spacebefore, spaceafter, spacebetween: Smallint; // Spacing
    number  : Smallint; // Numbering style (if <>0)
    numlevel: Byte;     // Numbering Level [1..9]  - used to select in group specifeid by 'number'
    color   : Byte;     // Background Color index
    shading : Byte;     // shading in %
    align   : TParAlign; // Alignment
    style   : Integer;   // Style-number  - see TWPRichText.GetStyleNameForNr)
    id      : Integer;   // For any use
    tabs    : TTabFlag;  // BitField: which Tab is used for the paragraph
    border  : TBorder;   // Border
    prop    : TParProp;
  end;
  {$ELSE}
  TParagraph = packed record
    indentfirst, indentleft, indentright: Smallint; // Indentation in twips
    spacebefore, spaceafter, spacebetween: Smallint; // Spacing
    number: Smallint; // Numbering style (if <>0)
    numstart : Smallint; // Numbering Start (if <>0)
    numlevel: Byte;   // Numbering Level [1..9]   }
    color: Byte;      // Background Color index
    shading: Byte;    // shading in %
    align: TParAlign; // Alignment
    style  : Integer; // Style definition number  - reserved
    id: Integer;      // For any use
    tabs: TTabFlag;   // BitField: which Tab is used for the paragraph
    border: TBorder;  // Border
    prop: TParProp;
    {$IFDEF MAXPARLENGTH}maxlen: Integer; {$ENDIF}
    { ---------------------------------------------- }
    line: PTLine;
    next, prev: PTParagraph;
    locked: TWPStyleElementsPar;
    level: Integer;
    autoid: Integer; // assigned in NewParagraph
    { ---------------------------------------------- }
    numcount, bookcount: Word;
    parobject: TParObject;
    celldef: PWPCell;
    attr: TAttr;
    parnum: Integer;
    linenum: Integer; // Number of first line par^.line
    SObj: TObject;
  end;
  {$ENDIF}

  {$IFNDEF T2H}
  PTWPParagraphParProp = ^TWPParagraphParProp;
  TWPParagraphParProp = packed record { only props }
    indentfirst, indentleft, indentright: Smallint; // Indentation in twips
    spacebefore, spaceafter, spacebetween: Smallint; // Spacing
    number: Smallint; // Numbering style (if <>0)
    numlevel: Byte;   // Numbering Level [1..9]   }
    color: Byte;      // HI nibble is Background Color index
    shading: Byte;    // shading in %
    align: TParAlign; // Alignment
    style: Integer;   // Style definition number  - reserved
    id: Integer;      // For any use
    tabs: TTabFlag;   // BitField: which Tab is used for the paragraph
    border: TBorder; // Border
    prop: TParProp;
    {$IFDEF MAXPARLENGTH}maxlen: Integer; {$ENDIF}
  end;
  {$ENDIF}

  TWPUpdateParEvent = procedure(Sender: TObject; Par: PTParagraph) of object;

  TWPEditFieldAlignOpt = set of (wpExtendToNextTab);

    TWPEditFieldGetSize = procedure(
     Sender : TObject;
     Const InspName : String;
     var   EndmarkWidthTW : Integer;
     var   Option : TWPEditFieldAlignOpt;
     CurrentTextWidthTW   : Integer;
     CurrentTextCharCount : Integer;
     par : PTParagraph; lin : PTLine; cp : Integer) of Object;

    TWPEditFieldFocusEvent = procedure(
     Sender : TObject; Const InspName : String; iTag : Integer; Enter : Boolean; var Abort : Boolean) of Object;

  PTCharacterAttr = ^TCharacterAttr;

  TWPTextEffect = (wpeffNone, wpeffPopup, wpeffOutline);

  //: OnGetAttrColor - Event is triggered for CharacterAttr when "UseGetAttrColorEvent = TRUE"
  TWPGetAttrColorEvent = procedure(
     Sender    : TObject;
     CharStyle : TOneWrtStyle;
     Canvas    : TCanvas;   // access to Font, Font.Color, Brush.Color
     var Underline, DoubleUnderline, Background : Boolean;
     HotstyleIsActive : Boolean;
     par : PTParagraph;
     lin : PTLine;
     cp  : Integer) of Object;

  TCharacterAttr = class(TPersistent)
  private
    FCharStyle: TOneWrtStyle;
    FBold: TThreeState;
    FItalic: TThreeState;
    FUnderline: TThreeState;
    FDoubleUnderline: Boolean;
    FStrikeOut: TThreeState;
    FSuperScript: TThreeState;
    FSubScript: TThreeState;
    FHidden: Boolean;
    FNotUsedForSpellAsYouGo : Boolean;
    FUseUnderlineColor: Boolean;
    FUseTextColor: Boolean;
    FUseBackgroundColor: Boolean;
    FUnderlineColor: TColor;
    FTextColor: TColor;
    FBackgroundColor: TColor;
    FHotStyleEffectColor: TColor;
    FHotUnderlineColor: TColor;
    FUseHotUnderlineColor: Boolean;
    FHotBackgroundColor: TColor;
    FUseHotBackgroundColor: Boolean;
    FHotEffect: TWPTextEffect;
    FHotTextColor: TColor;
    FUseHotTextColor: Boolean;
    FHotUnderline: TThreeState;
    FHotStyleIsActive: Boolean;
    FUseGetAttrColorEvent : Boolean;
    FAlsoUseForPrintout: Boolean;
    procedure SetBold(x: TThreeState);
    procedure SetItalic(x: TThreeState);
    procedure SetUnderline(x: TThreeState);
    procedure SetDoubleUnderline(x: Boolean);
    procedure SetStrikeOut(x: TThreeState);
    procedure SetSuperScript(x: TThreeState);
    procedure SetSubScript(x: TThreeState);
    procedure SetHidden(x: Boolean);
    procedure SetUseUnderlineColor(x: Boolean);
    procedure SetUseTextColor(x: Boolean);
    procedure SetUseBackgroundColor(x: Boolean);
    procedure SetUnderlineColor(x: TColor);
    procedure SetTextColor(x: TColor);
    procedure SetBackgroundColor(x: TColor);
    procedure SetHotUnderlineColor(x: TColor);
    procedure SetHotBackgroundColor(x: TColor);
    procedure SetHotTextColor(x: TColor);
    procedure SetHotEffect(x: TWPTextEffect);
  public
    {$IFNDEF T2H}
    FOnAttributeChanged: TCAttributeChangedEvent;
    procedure Assign(Source: TPersistent); override;
    constructor Create(FCharStyle: TOneWrtStyle);
    {$ENDIF}
    property CharStyle: TOneWrtStyle read FCharStyle;
  published
    property Bold: TThreeState read FBold write SetBold;
    property Italic: TThreeState read FItalic write SetItalic;
    property DoubleUnderline: Boolean read FDoubleUnderline write
      SetDoubleUnderline;
    property Underline: TThreeState read FUnderline write SetUnderline;
    property StrikeOut: TThreeState read FStrikeOut write SetStrikeOut;
    property SuperScript: TThreeState read FSuperScript write SetSuperScript;
    property SubScript: TThreeState read FSubScript write SetSubScript;
    property Hidden: Boolean read FHidden write SetHidden;
    property NotUsedForSpellAsYouGo : Boolean read FNotUsedForSpellAsYouGo write FNotUsedForSpellAsYouGo default FALSE;
    property UnderlineColor: TColor read FUnderlineColor write
      SetUnderlineColor;
    property TextColor: TColor read FTextColor write SetTextColor;
    property BackgroundColor: TColor read FBackgroundColor write
      SetBackgroundColor;
    property UseUnderlineColor: Boolean read FUseUnderlineColor write
      SetUseUnderlineColor;
    property UseTextColor: Boolean read FUseTextColor write SetUseTextColor;
    property UseBackgroundColor: Boolean read FUseBackgroundColor write
      SetUseBackgroundColor;
    // These properties are used if the text is 'under the mouse' (afsHotStyle is used)
    property HotUnderlineColor: TColor read FHotUnderlineColor write SetHotUnderlineColor default clBlack;
    property UseHotUnderlineColor: Boolean read FUseHotUnderlineColor write FUseHotUnderlineColor default FALSE;
    property HotBackgroundColor: TColor read FHotBackgroundColor write setHotBackgroundColor default clBlack;
    property UseHotBackgroundColor: Boolean read FUseHotBackgroundColor write FUseHotBackgroundColor default FALSE;
    property HotTextColor: TColor read FHotTextColor write setHotTextColor default clBlack;
    property UseHotTextColor: Boolean read FUseHotTextColor write FUseHotTextColor default FALSE;
    property HotUnderline: TThreeState read FHotUnderline write FHotUnderline default tsIgnore;
    property HotEffect: TWPTextEffect read FHotEffect write SetHotEffect default wpeffNone;
    property HotStyleIsActive: Boolean read FHotStyleIsActive write FHotStyleIsActive default FALSE;
    property AlsoUseForPrintout: Boolean read FAlsoUseForPrintout write FAlsoUseForPrintout default TRUE;
    property HotEffectColor: TColor read FHotStyleEffectColor write FHotStyleEffectColor default clBlack;
    property UseOnGetAttrColorEvent : Boolean read FUseGetAttrColorEvent write FUseGetAttrColorEvent default FALSE;
  end;


  {$IFNDEF T2H}
  TParBandObject = class(TParObject)
  public
    procedure Paint(Memo: TObject; toCanvas: TCanvas;
      par: PTParagraph; var r: TRect); virtual; abstract;
    function GetAsString: string; virtual; abstract;
    procedure SetAsString(const x: string); virtual; abstract;
  end;

  TParRectObject = class(TParObject) { moved by Set_TopOffset }
  public
    { outer bounds }
    pTop: Longint; { relative to top_offset	and left_offset	}
    pLeft: Integer;
    pBottom: Longint; { relative to top_offset	and left_offset	}
    pRight: Integer;
    { inner bounds }
    pInnerLeft: Integer;
    pInnerRight: Integer;
    pBorderSpace: Integer;
  end;

  TWPAssignMergeBandParProc = procedure(Memo: TObject;
    par: PTParagraph; const s: string);
  {$ENDIF}

  TWPPagePropertyRange =
    (wpraOnAllPages, wpraOnOddPages, wpraOnEvenPages, wpraOnFirstPage,
    // these Modes are not compatible to the RTF Standard !
    // They have priority !!!!
    wpraOnLastPage, wpraNotOnFirstAndLastPages,
    wpraNamed,
    wpraIgnored);

  { SetCharAttr }
  TAttrWhat = set of
    (awFontNr,
    awFontSize,
    awFontIncSize,
    awFontDecSize,
    awFontColor,
    awFontBKColor,
    awFontStyleAdd,
    awFontStyleReplace,
    awFontStyleReplaceNormal,
    awSetStandardFontStyle,
    awFontStyleSub,
    awWidth,
    awHeight);

  { only used for ChangeAttr}
  TOneWhatToChange = (wtcFont, wtcSize, wtcColor, wtcBKColor, wtcStyle,
    wtcAddStyle, wtcSubStyle);
  TWhatToChange = set of TOneWhatToChange;

  {$IFNDEF T2H}
  { Get	Info about Character (to save for later	use) }
  TInfoBlock = record
    Metrics: TTextMetric;
    Width: array[0..256] of Integer;
    Height: Integer;
    Base: Integer;
    Fontnam: string;
    {--------  }
    Atr: TAttr;
  end;

  { Changed Parmeter fuer reformat_par und Input }
  TDoRefresh = (refreshNone, refreshSomeLin, refreshPar, refreshAll);
  {$ENDIF}
  TWPMemoryFormat = (fmRichText, fmPlainText, fmRichTextUnicode);

  TParPropsOption = (ppoEnabled,
    ppoIndentfirst,
    ppoIndentleft,
    ppoIndentright,
    ppoSpacebefore,
    ppoSpaceafter,
    ppoSpacebetween,
    ppoMultSpacing,
    ppoBorderType,
    ppoTabs,
    ppoNewPage,
    ppoColor,
    ppoShading,
    ppoAlign,
    ppoAttr,
    ppoTabsList,
    ppoParProtected,
    ppoParID,
    ppoCellAlign);
             
  TParPropsOptions = set of TParPropsOption;

  { Record to pass text +	attributes to the FastAddPar and FastAddTable procedures }
  {$IFNDEF T2H}
  TParagraphProperty = record
    Options: TParPropsOptions;
    Indentfirst, Indentleft, Indentright: Integer;
    { Indentation [Twips] (max	22 inch!) }
    Spacebefore, Spaceafter, Spacebetween: Integer; { Spacing (Twips!)	}
    BorderType: TBorder; { only blBox	within tables }
    Color: Byte; { Paragraph Color }
    Shading: Byte; { shading in	% }
    Align: TParAlign;
    TabsPosList: array[0..5] of Longint;
    TabsKindList: array[0..5] of TTabKind;
    Tab: TTabFlag;
    ParProtected: Boolean;
    CellAlignCenter: Boolean;
    CellAlignBottom: Boolean;
    Id: Integer;
  end;
  PTParagraphProperty = ^TParagraphProperty;
  {$ENDIF}

  PTParProps = ^TParProps;
  TParProps = record
    Options: TParPropsOptions;
    Indentfirst, Indentleft, Indentright: Integer;
    Spacebefore, Spaceafter, Spacebetween: Integer;
    BorderType: TBorder;
    Color: Byte;
    Shading: Byte;
    Align: TParAlign;
    TabsPosList: array[0..5] of Longint;
    TabsKindList: array[0..5] of TTabKind;
    Tab: TTabFlag;
    ParProtected: Boolean;
    CellAlignCenter: Boolean; // Option ppoCellVertAlign
    CellAlignBottom: Boolean; // Option ppoCellVertAlign
    Id: Integer;
    { --- }
    Attr: TAttr;
    Text: string;
    pText: PChar;
    pa: PTAttr;
    TXTObj: Pointer;
    Obj: TObject; // Graphic of type TWPObject
    CWidth: Integer;
    NewPage: Boolean;
    ColSpan: Byte;
    CellName: string;
    CellCommand: string;
    CellFormat: Integer;
  end;

  TWPPageBackgroundBuffer = class(TObject)
  public
    bitmap: Graphics.TBitmap; // The bitmap used for the buffer
    lastuse: Integer; // Lastuse (for the cache)
    pageno_start: Integer; // Used for starting ...
    pageno_end: Integer; // Used for ... until
    MustInitialize: Boolean; // Need Repaint
  end;

  TWPPaintPageEvent = procedure(
    Sender: TObject; // The TWPRichText
    toCanvas: TCanvas; // Drawing Canvas
    PageRect: TRect; // The rectangle
    XRes, YRes, PageNo: Integer; // Res and
    firstpar: PTParagraph; // start par of this page
    firstline: PTLine; // start line of this page
    BufferObj: TWPPageBackgroundBuffer // for various Use (can be NIL)
    ) of object;

  { V2.0 ---- UNDO List }
  {$IFNDEF T2H}
  TWPUndoType = (
    wpuNone,
    wpuReplacePar, { replaces complete paragraph	# with Contents, sets Cursor }
    wpuReplaceParText, {	replaces paragraphtext # with Contents,	sets Cursor }
    wpuSetParIndent, { replaces paragraph indents }
    wpuSetParSpacing, { replaces	paragraph spacing }
    wpuSetParStyle,
    wpuSetParAlignment,
    wpuSetParBorder,
    wpuSetParTabs,
    wpuDeletePar, { deletes Par #	}
    wpuInsertPar, { insert	Par after # }
    wpuSetParProps, { cares about all paragraph properties	}
    wpuSetCellData,
    wpuSetTableData,
    wpuReLoadText, { load text from alldata ! }
    // Object ...
    wpuSetObjectBounds,
    wpuSetObjectPosMode,
    wpuSetStyleCollection,
    wpuSetMargin,
    wpuSetAttr
    );
  {$ENDIF}

  TWPUndoKind =
    (wputNone,
    wputAny,
    wputInput,
    wputDeleteText,
    wputChangeAttributes,
    wputChangeIndent,
    wputChangeSpacing,
    wputChangeAlignment,
    wputChangeTabs,
    wputChangeBorder,
    wputDeleteSelection,
    wputDragAndDrop,
    wputPaste,
    wputInsertObject,
    wputChangeObject,
    wputDeleteObject,
    wputChangeTable,
    wputChangeStyleSheet,
    wputRedo,
    wputMargin
    );

  TWPNotifyUndoChange = procedure(Sender: TObject; kind: TWPUndoKind) of object;


  {$IFNDEF T2H}
  PTWPUndoList = ^TWPUndoList;
  TWPUndoList = record
    next: PTWPUndoList;
    typ: TWPUndoType;
    kind: TWPUndoKind;
    level: Integer; { Current undo level. Cursor movement increases }
    parnum: Longint; { Number of Paragraph }
    partree: PTParagraph; { always only ONE Paragraph. next/prev not	used }
    CWidthPC, CLeftPC: Word;
    TableNumber: Longint;
    { spacing or	indent values }
    par: TParagraph;
    alldata: TMemoryStream;
    alldatatime: Integer;
    { New Cursorpos }
    cppar, cppos: Longint;
    topoffset: Longint;
  end;

  TWPTextColorsIndex = 0..NumPaletteEntries;

  { Event: Move	Cursor Position	 }
  TCursorPosChangedEvent = procedure(Sender: TObject) of object;
  { Event: Set new scrolling range }
  TSetScrollRangeEvent = function: Boolean of object;
  { Event: Check Clipboard }
  TEditStateChangedEvent = procedure(Sender: TObject) of object;
  { Event: Hide	Cursor }
  TPaintStartEvent = procedure of object;
  { Event: Show	Cursor }
  TPaintEndEvent = function: Boolean of object;

  TWPOnePaintMode = (prIntersectClip,
    prOnlyPaintLine,
    prIgnoreFitToWindow,
    prIgnoreZooming,
    prNoCursor,
    prNoScrolling,
    prTransparent,
    prHideAutoNewPage,
    prHideManualPage,
    prShowHidden,
    prWait, { temporary no Reformat	or Paint }
    prIgnoreLeftOffset, { no X Scrolling }
    prIgnoreSelections, { no Selections	 }
    prIgnoreHyperlinks,
    prIgnoreHeight, { no clipping on height	}
    prIgnoreDefaultBKColor,
    prWhiteIsTransparent,
    prHideHalfLines,
    prUsePageBreaks, { for WPLabel. Only print current page	}
    prIgnoreBKColor, { MKColor = White	 }
    prOneLineOnly,
    prLimitHorz,
    prDontChangeClipping,
    prIgnoreParagraphFormat,
    prUsePaper,
    prAllTransparent // even paragraph background
    ); {to another resolution}
  WPPaintMode = set of TWPOnePaintMode;
  {$ENDIF}

  TWPActivateHotStyle = procedure(Sender: TObject;
    pa: PTAttr; par: PTParagraph; lin: PTLine; cp: Integer) of object;
  TWPDeactivateHotStyle = procedure(Sender: TObject) of object;

  TWPCustomStyleCollection = class;

  TWPCustomRtfEditClass = class(TCustomControl)
    {$IFNDEF T2H}
  public
    FOnChangeEditBoxWidth: TWPChangeBoxEvent;
    FOnChangeEditBoxHeight: TWPChangeBoxEvent;
    FOnCursorLeavesEditBox: TWPCursorLeavesBoxEvent;
    FOnFlowTextOutOfEditBox: TWPFlowTextOutOfBox;
    FOnFlowTextIntoEditBox: TWPFlowTextIntoBox;
    FOnGetPageGapText: TWPOnGetPageGapText;
    FEditBoxModes: TWPEditBoxModes;
    FOnActivateHotStyle: TWPActivateHotStyle;
    FOnDeactivateHotStyle: TWPDeactivateHotStyle;
    FStyleCollection: TWPCustomStyleCollection;
  protected
    procedure SetStyleCollection(x: TWPCustomStyleCollection); virtual;
  public
    FRefreshPreview: Boolean;
    procedure _SetFocusAndInput(ToEnd: Boolean; pos: Integer; NewChar: Char;
      Source: TWPCustomRtfEditClass); virtual; abstract;
    procedure _SaveRectToStream(h, w: Integer; mstr: TStream; deleteit:
      Boolean); virtual; abstract;
    procedure OnToolBarIconSelection(Sender: TObject;
      var Typ: TWpSelNr; const str: string; const group, num, index: Integer);
      virtual; abstract;
    procedure DoChangePageSize; virtual; abstract;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState;
      X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState;
      X, Y: Integer); override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure OnToolBarSelection(Sender: TObject; var Typ: TWpSelNr;
      const str: string; const num: Integer); virtual; abstract;
    procedure SaveToStream(s: TStream); virtual; abstract;
    procedure ResizeWin; virtual; abstract;
    procedure Clear; virtual; abstract;
    function GetHeader: TTextHeader; virtual; abstract;
    function GetMemo: TObject; virtual; abstract;
    procedure SetFocusValues(Always: Boolean); virtual;
    property WPStyleCollection: TWPCustomStyleCollection read FStyleCollection write SetStyleCollection;
    {$ENDIF}
  end;

  {$IFNDEF T2H}
  TNotifyRTFEditChangedEvent = procedure(Sender: TObject; NewRTFEdit:
    TWPCustomRtfEditClass) of object;
  {$ENDIF}

  { +++++++++++++++++++ BASIC Reader / Writer Classes +++++++++++++++++++++ }

  TWPCustomTextReader = class(TPersistent) // TPersistent 'cause needs to be registered
    {$IFNDEF T2H}
  public
    par: TParagraph; { actual Values	       }
    preformatted: Boolean;
    getcrcodes: Boolean;
    ignoreheaderfooter: Boolean;
    FCurrentFont: Integer;
    FCurrentFontSize: Integer;
    ixPos, ixSize, ixPagCount: Longint;
  protected
    lin: TLine; { all pointers are nil }
    fpIn: TSTREAM;
    load_line: PTLine;
    load_pos: Integer;
    load_anz: LongInt;
    load_pc: Pchar; { faster access ! }
    load_pa: PTAttr;
    priorload_par: PTParagraph;
    priorload_line: PTLine;
    priorload_pos: Integer;
    priorload_anz: LongInt;
    priorload_pc: Pchar;
    priorload_pa: PTAttr;
    FLastError: ecErrCode;
    ReadBuffer: array[0..WPCustomTextReaderReadBufferLen - 1] of Char;
    ReadBufferLen: Integer;
    ReadBufferPos: Integer;
    ReadBufferStart: Longint;
    FReformatCallback: TNotifyEvent;
    FAddedParCount: Integer;
  protected
    function ParseEscapeCode(ch: Integer): ecErrCode; virtual; abstract;
    function ParseChar(ch: Integer): ecErrCode; virtual; abstract;
    function PrintChar(ch: Integer): ecErrCode; virtual; abstract;
    function PrintAChar(c: Char): ecErrCode; virtual; abstract;
  public
    procedure Init(fp: TStream); virtual; abstract;
    function Parse(fp: TStream; var FirstPar: PTParagraph): ecErrCode; virtual;
      abstract;
    function GetPar(ptag: Pchar): PChar; virtual;
    constructor Create; virtual;
  public { These properties are set befor	the call of PARSE by WPRichText	}
    Attr: TAttr; { actual writing mode	       }
    FHeader: TTextHeader; { for page and font definition }
    FFMemo: TObject; { TWPRTfTXT for	memory allocation	 }
    load_par: PTParagraph; { if nil then create new tree  }
    load_parnr: Longint;
    EscapeCode: Char;
    ReadAutomatic: Boolean; { Used for RTF mailmerging     }
    IgnorePagePar: Boolean; { Ignore pagewidth and colortable }
    IgnoreBorderPar: Boolean;
    { These variables are	optional -->
      a buffer to	the definition of header/footer	in the correct format}
    FHeaderText, FFooterText: PTParagraph;
    { Parameters used by the HTML	Reader }
    SaveGraphicDataInRTF: Boolean;
    StoreComments: TStoreHTMLComment;
    FLoadPath: string;
    DefSize: Integer;
    DefFont: Integer;
    BaseFont: Integer;
    IgnoreNext_P: Boolean;
    property ReformatCallback: TNotifyEvent read FReformatCallback write FReformatCallback;
    property LastError: ecErrCode read FLastError; { ecOK if everything ok }
    {$ENDIF}
  end;

  TWPTextReaderClass = class of TWPCustomTextReader;

  TWPCustomTextWriter = class(TPersistent)
    {$IFNDEF T2H}
  protected
    fpOut: TSTREAM;
    prior_par: PTParagraph; {	prior Values or	nil  }
    prior_lin: PTLine;
    prior_pa: PTAttr;
    save_par: PTParagraph; { actual Values   }
    save_lin: PTLine;
    save_pos: Integer;
    save_anz: LongInt;
    save_pc: Pchar; { faster access ! }
    save_pa: PTAttr;
    LastError: ecErrCode;
    FCalcDefaultFont: Boolean;
    FIgnoreStyles: Boolean;
    procedure IntToBuffer(var p: pchar; val: Longint);
    procedure HexToBuffer(var p: pchar; val: char);
    function WriteHexBuffer(mem: pointer; size: Longint): ecErrCode;
  public
    procedure Init; virtual;
    function WriteHeader(fp: TStream): ecErrCode; virtual;
    function WriteText(fp: TStream): ecErrCode; virtual;
    function WriteFooter(fp: TStream): ecErrCode; virtual;
    property CalcDefaultFont: Boolean read FCalcDefaultFont;
    constructor Create; virtual;
  public
    FFMemo: TObject;
    FHeaderText, FFooterText: PTparagraph;
    FOnlySelectedPar: Boolean;
    par_s, par_e: PTparagraph;
    lin_s, lin_e: PTLine;
    cp_s, cp_e: Integer;
    FHideInsertpoints: Boolean;
    FHideAutomatic: Boolean;
    FIsEmpty: Boolean;
    FDontWriteNULL: Boolean;
    FWriteBitmapMode: TWPWriteBitmapMode;
    { USed for the HTML Writer }
    DefaultFont: Integer;
    DefaultSize: Integer;
    DefaultColor: Integer;
    DefaultFontSize: Integer; { Last Fontsize }
    DontUseFontTag: Boolean;
    FWriteHelpCompilerStyle: Boolean;
    FDontCopyProtectedAttribute: Boolean;
    FDontSaveBorderAttribute: Boolean;
    FWritePath: string; { for HTML }
    {$ENDIF}
  end;

  TWPCustomToolCtrl = class(TCustomControl)
    {$IFNDEF T2H}
  protected
    FNextWPTCtrl: TWPCustomToolCtrl;
    FRtfEdit: TWPCustomRtfEditClass;
    FRTFEditChanged: TNotifyRTFEditChangedEvent;
    FAutoEnabling: Boolean;
    procedure SetRTFedit(x: TWPCustomRtfEditClass); virtual;
    procedure SetFPPaletteEntries(x: PTColorPalette); virtual; abstract;
  public
    FDefaultPalette: TColorPalette;
    FReferenceCount: Integer;
    VFPPaletteEntries: PTColorPalette;
    procedure OnColorDropDown(Sender: TObject);
    function SelectIcon(index, group, num: Integer): Boolean; virtual; abstract;
    procedure UpdateEnabledState; virtual;
    procedure SetPreviewButtons; virtual;
    function DeselectIcon(index, group, num: Integer): Boolean; virtual;
      abstract;
    function EnableIcon(index, group, num: Integer; enabled: Boolean): Boolean;
      virtual; abstract;
    procedure UpdateSelection(Typ: TWpSelNr; const str: string; num: Integer);
      virtual; abstract;
    procedure PerformAll(m: Cardinal; w, l: Longint); virtual; abstract;
    property RtfEdit: TWPCustomRtfEditClass read FRtfEdit write SetRtfEdit;
    property NextToolBar: TWPCustomToolCtrl read FNextWPTCtrl write
      FNextWPTCtrl;
    property FPPaletteEntries: PTColorPalette read VFPPaletteEntries write
      SetFPPaletteEntries;
    property AutoEnabling: Boolean read FAutoEnabling write FAutoEnabling;
  published
    property RTFEditChanged: TNotifyRTFEditChangedEvent read FRTFEditChanged
      write FRTFEditChanged;
    {$ENDIF}
  end;

  // Used as interface for reader and writer
  TWPStyleBinValues = record
    Name            : string[60];
    BasedOnStyle    : string[60];
    NextStyle       : string[60];
    Par             : TParagraph;
    Attr            : TAttr;
    ParElements     : TWPStyleElementsPar;
    AttrElements    : TWPStyleElementsChar;
    Extra           : TWPExtraStyleProps;
  end;

  TWPCustomStyleCollectionItem = class(TCollectionItem)
  public
    LoadedFromDocument: Boolean;
  public
    procedure ReadFromPar(par: PTParagraph; pa: PTAttr;
      RTFEngine: TObject; OnlyModified: Boolean;
      ParOnly: TWPStyleElementsPar; AttrOnly: TWPStyleElementsChar); virtual; abstract;
    procedure Update(RTFEngine: TObject; aPar: PTParagraph;
      var WasInitalized: Boolean); virtual; abstract;
    function _Translate(RTFEngine: TObject; var StyleBinValues: TWPStyleBinValues; InheritedToo: Boolean): Boolean; virtual; abstract;
  end;

  TWPCustomStyleCollection = class(TComponent)
  protected
    FControlledMemos: TList;
  public
    _DontAssignDefault : Boolean;
    constructor Create(aOwner: TComponent); override;
    destructor Destroy; override;
    {$IFNDEF T2H}
    function _InternFind(const StyleName: string): TWPCustomStyleCollectionItem; virtual; abstract;
    procedure _InternModify(RTFEngine: TObject;
      const StyleBinValues: TWPStyleBinValues; add, overwrite: Boolean); virtual; abstract;
    procedure _InternUpdateStart(RTFEngine: TObject); virtual; abstract;
    procedure _InternUpdate(par: PTParagraph; stylenr: Integer); virtual; abstract;
    procedure _InternNewParagraph(RTFEngine: TObject; par: PTParagraph); virtual; abstract;
    {$ENDIF}
    procedure GetStyleList(Strings: TStrings); virtual; abstract;
    procedure SaveToStream(Stream: TStream); virtual; abstract;
    procedure LoadFromStream(Stream: TStream); virtual; abstract;
  end;

  TWPTextWriterClass = class of TWPCustomTextWriter;

  //: If this exception is triggered a necessary unit was not linked to the project
  EWPIOClassError = class(Exception);
  //: This error should never happen. If it does happen save the text and close application
  EWPEngineError = class(Exception);
  //: If this error happens WPTools was not used in the right way
  EWPEngineModeError = class(Exception);

  EWP_Print_Selected_Error = class(Exception) ;

  {$IFNDEF T2H}

  TWPVCLString = (
    meDefaultUnit_INCH_or_CM,
    meUnitInch,
    meUnitCm,
    meDecimalSeperator,
    meDefaultCharSet,

    meFilterRTF,
    meFilterHTML,
    meFilterTXT,
    meFilterXML,
    meFilterCSS,
    meFilterALL,

    meFilter,
    meUserCancel,
    meReady,
    meReading,
    meWriting,
    meFormatA,
    meFormatB,
    meClearMemory,
    meNoSelection,
    meRichText,
    mePlainText,
    meSysError,
    meClearTabs,
    meObjGraphicFilter,
    meObjGraphicFilterExt,
    meObjTextFilter,
    meObjNotSupported,
    meObjGraphicNotLinked,
    meObjDelete,
    meObjInsert,
    meSaveChangedText,
    meClearChangedText,
    meCannotRenameFile,
    meErrorWriteToFile,
    meErrorReadingFile,
    meRecursiveToolbarUsage,
    meWrongScreenmode,
    meTextNotFound,
    meUndefinedInsertpoint,
    meNotSupportedProperty,
    meUseOnlyBitmaps,
    meWrongUsageOfFastAppend,
    mePaletteCompletelyUsed,
    meWriterClassNotFound,
    meReaderClassNotFound,
    meWrongArgumentError,
    meIOClassError,
    meNotSupportedError,
    meutNone,
    meutAny,
    meutInput,
    meutDeleteText,
    meutChangeAttributes,
    meutChangeIndent,
    meutChangeSpacing,
    meutChangeAlignment,
    meutChangeTabs,
    meutChangeBorder,
    meutDeleteSelection,
    meutDragAndDrop,
    meutPaste,
    meutChangeTable,
    meutChangeGraphic,
    meutChangeCode,
    meutUpdateStyle,
    meutChangeTemplate,
    meutInsertBookmark,
    meutInsertHyperlink,
    meutInsertField,
    // Used in unit WPPanel
    meDiaAlLeft, meDiaAlCenter, meDiaAlRight, meDiaAlJustified,
    // Used in WPTblDlg
    meDiaYes, meDiaNo,
    // Specify the TabKind
    meTkLeft, meTkCenter, meTkRight, meTkDecimal,
    // Spacing used in WPParPrp
    meSpacingMultiple, meSpacingAtLeast, meSpacingExact,
    // Used for Page Properties Dialog
    mePaperCustom,
    // Used in WPReporter
    meWrongFormat, meReportTemplateErr,
    meUnicodeModeNotActivated
    );


  TWPCustomRTFFiler = class(TComponent)
  protected
    function  GetCurrentPage: Integer; virtual;
    function  GetPageCount: Integer; virtual;
    procedure SetCurrentPage(x: Integer); virtual;
  public
    procedure ReadStream(s: TStream); virtual;
    procedure ReadFile(const fname: string); virtual;
    function  Find(const text: string; wildcard: Char; CaseSensitive: Boolean):Boolean;virtual;
    procedure Close; virtual;
    procedure Print; virtual;
    property  CurrentPage: Integer read GetCurrentPage write SetCurrentPage;
    property  PageCount: Integer read GetPageCount;
  end;


function WPLoadStr(item: TWPVCLString): string;

//function WPLoadFormat(x: TWPLoadFormat): TWPNewLoadFormat;
//function WPSaveFormat(x: TWPLoadFormat): TWPNewLoadFormat;
{$ENDIF}

function WPCentimeterToTwips(value: Double): Longint;
function WPInchToTwips(value: Double): Longint;
function WPTwipsToCentimeter(value: Longint): Double;
function WPTwipsToInch(value: Longint): Double;
function WPCentimeterToPixel(value: Double; Resolution: Integer): Longint;
function WPInchToPixel(value: Double; Resolution: Integer): Longint;

{$IFNDEF T2H}
procedure WPSubstChar(var S: string; OldC, NewC: Char);
function WPListIntItemCompare(Item1, Item2: Pointer): Integer;

procedure WPAssignStyleAttrValues(Dest: PTAttr; DestLocked: WrtStyle;
  const Source: PTAttr; SourceSel: TWPStyleElementsChar; SetDefault: Boolean);

procedure WPAssignStyleParValues(Dest: PTParagraph; DestLocked: TWPStyleElementsPar;
  const Source: PTParagraph; SourceSel: TWPStyleElementsPar; SetDefault: Boolean);
{$ENDIF}

const
  {$IFNDEF T2H}
  awAll = [awFontNr, awFontSize, awFontColor,
    awFontBKColor, awFontStyleAdd, awFontStyleReplace,
    awFontStyleReplaceNormal, awFontStyleSub, awWidth,
    awHeight];
  WPStyleAttrElementLocking: array[TWPStyleElementChar] of TOneWrtStyle = (afsParStyleBold,
    afsParStyleItalic, afsParStyleUnderline, afsParStyleFont, afsParStyleSize,
    afsParStyleColor, afsParStyleBGColor);
  WPAttrCharStyle = [afsBold, afsItalic, afsHyperLink,
    afsSuper, afsSub, afsProtected, afsIsFootnote, afsUnderline,
    afsStrikeOut, afsHidden, afsIsRTFCode, afsUserdefined, afsBookmark,
    afsIsObject];

  WPAttrCharStyleWhichAffectSize = // those can change the size of a char !
    [afsBold,afsItalic,afsSuper,afsSub,afsUserdefined,afsUppercaseStyle];

  wpse_TheLastItem = High(TWPStyleElementPar);

  // Do not localize this strings - RTF information
  WPPredefinedFields: array[0..9] of string =
  ('Title', 'Subject', 'Author', 'Keywords', 'Operator', 'DocComm', // Standard Info fields
    'Manager', 'Company', 'Category', 'HLinkBase'); // Extra \* fields

var
  FWPStyleParElementName: array[TWPStyleElementPar] of string = (// Don't localize!
    'IndentLeft', 'IndentRight', 'IndentFirst',
    'SpaceBefore', 'MultSpaceBetween', 'SpaceBetween', 'SpaceAfter',
    'BrdLines', 'BrdWidth', 'BrdColor',
    'Align', 'Color', 'Shading', 'Tabs',
    'Number', 'NumLevel', 'IsOutline', 'BrdSpace', 'KeepTogther', 'KeepWithNext');

  FWPStyleAttrElementName: array[TWPStyleElementChar] of string = (// Don't localize!
    'Bold', 'Italic', 'Underline', 'Font', 'Size',
    'FontColor', 'FontBGColor');

  FWPVCLStrings: array[TWPVCLString] of string = (// Localizable via unit WPLanCtr !
    { meDefaultUnit_INCH_or_CM}'CM', // only INCH or CM !
    { meUnitInch }'Inch',
    { meUnitCm }'Cm',
    { meDecimalSeperator }'', // #0, . or ,
    { meDefaultCharSet }'0', // ANSI = 0

    { meFilterRTF }'RTF plik (*.rtf)|*.RTF',
    { meFilterHTML }'HTML plik (*.htm,*.html)|*.HTM;*.HTML',
    { meFilterTXT }'Plik tekstowy (*.txt)|*.TXT',
    { meFilterXML }'XML plik (*.xml)|*.XML',
    { meFilterCSS }'Plik ze stylami (*.CSS)|*.CSS',
    { meFilterALL }'Wszystkie pliki (*.*)|*.*',

    { meFilter }'RTF plik (*.rtf)|*.RTF|HTML plik (*.htm,*.html)|*.HTM;*.HTML|Plik tekstowy (*.txt)|*.TXT|Wszystkie pliki (*.*)|*.*',
    { meUserCancel }'Czy chcesz przerwaæ operacjê?',
    { meReady }'Gotowe',
    { meReading }'odzcyt ...',
    { meWriting }'zapis ...',
    { meFormatA }'formatowanie ...',
    { meFormatB }'formatowanie ...',
    { meClear{ memory }'inicjalizacja',
    { meNoSelection }'nie zaznaczono tekstu',
    { meRichText }'RICH',
    { mePlainText }'PLAIN',
    { meSysError }'Internal Error',
    { meClearTabs }'Usun¹æ wszystkie tabstopy?',
    { meObjGraphicFilter }'Pliki graficzne (*.BMP,*.WMF)|*.BMP;*.WMF',
    { meObjGraphicFilterExt }'Pliki graficzne (*.BMP,*.WMF,*.JPG,*.GIF)|*.BMP;*.WMF;*.JPG;*.GIF',
    { meObjTextFilter }'Pliki tekstowe (*.TXT)|*.TXT',
    { meObjNotSupported }'Cannot use other objects than of type TWPObject.',
    { meObjGraphicNotLinked }'Graphic support was not linked. Please include unit WPEmOBJ to project.',
    { meObjDelete }'Usuniêcie obiektu',
    { meObjInsert }'Wstawienie nowego obiektu',
    { meSaveChangedText }'Zapisaæ zmiany?',
    { meClearChangedText }'Anulowaæ zmiany?',

    { meCannotRenameFile }'Nie mo¿na zmieniæ nazwy pliku %s!',
    { meErrorWriteToFile }'Nie mo¿na zapisaæ pliku %s!',
    { meErrorReadingFile }'B³¹d podczas odczytu z pliku %s!',

    { meRecursiveToolbarUsage }'property NextToolBar was used recursively',
    { meWrongScreenmode }'Wrong ScreenResMode in object ',
    { meTextNotFound }'Nie znaleziono tekstu',
    { meUndefinedInsertpoint }'Insertpoint not defined. Provide tag or fieldname',
    { meNotSupportedProperty }'This property is not yet supported',
    { meUseOnlyBitmaps }'Please use only bitmaps',
    { meWrongUsageOfFastAppend }'cannot append text to <self>',
    { mePaletteCompletelyUsed }'All colors in palette in use',
    { meWriterClassNotFound }'Text-Writer class not found ',
    { meWrongArgumentError }'Wrong Arguments!',
    { meReaderClassNotFound }'Text-Reader class not found ',
    { meIOClassError }'Incorrect IO Class :',
    { meNotSupportedError }'Mode not supported',
    // UNDO Strings
    { meutNone }'',
    { meutAny }'Changes',
    { meutInput }'Input',
    { meutDeleteText }'Delete',
    { meutChangeAttributes }'Attribute Modification',
    { meutChangeIndent }'Indent Modification',
    { meutChangeSpacing }'Spacing Modification',
    { meutChangeAlign{ ment }'Align{ ment Modification',
    { meutChangeTabs }'Tabs Modification',
    { meutChangeBorder }'Border Modification',
    { meutDeleteSelection }'Deletion',
    { meutDragAndDrop }'Drag and Drop',
    { meutPaste }'Paste',
    { meutChangeTable }'Change Table',
    { meutChangeGraphic }'Change Graphic',
    { meutChangeCode }'Change Code',
    { meutUpdateStyle}'Update Style',
    { meutChangeTemplate}'Change Template',
    { meutInsertBookmark }'Insert Bookmark',
    { meutInsertHyperlink }'Insert Hyperlink',
    { meutInsertField }'Insert Field',
    {meDiaAlLeft, meDiaAlCenter,meDiaAlRight,meDiaAlJustified}
    'Lewy', 'rodek', 'Prawy', 'Obustronnie',
    {meDiaYes, meDiaNo }'Tak', 'Nie',
    {meTkLeft, meTkCenter,meTkRight,meTkDecimal}
    'Lewy', 'Prawy', 'rodek', 'Dziesiêtny',
    {meSpacingMultiple, meSpacingAtLeast,meSpacingExact}
    'wielokrotnie', 'przynajmniej', 'dok³adnie',
    {mePaperCustom}'Rozmiar u¿ytkownika',
    { meWrongFormat}'Plik ma z³y format!',
    { meReportTemplateErr }'Error in Report Template!',
    { meUnicodeModeNotActivated } 'Unicode mode was not activated'
    );

function WPUndoString(kind: TWPUndoKind): string;
procedure WPPrinterOpen;
procedure WPPrinterOpenName(pname: string);
procedure WPPrinterClose;

procedure WPRegisterClass(AClass: TPersistentClass);
procedure WPRegisterClasses(AClasses: array of TPersistentClass);

function WPColorToString(Color: TColor): string;
function WPStringToColor(S: string): TColor;
function WPBufferToHex(p: PChar): Byte;
function WPCompareParAttr(pPar1, pPar2: PTParagraph; BorderAttr, SpecialAttr: Boolean): TWPStyleElementsPar;
function WPCompareCharAttr(pa, pa2: PTAttr): TWPStyleElementsChar;
function WPFormatStringForGroupbox(const source: string): string;

const
  WPMaxReformatLineLength = 1000;

var
  TabsRef: array[0..31] of Integer;
  WPDontCheckFontCharsets: Boolean;
  WPAssignMergePar: TWPAssignMergeBandParProc;
  WPDontWriteSpacing: Boolean;
  WPDontWriteIndents: Boolean;
  WPDontPresetBorder: Boolean = TRUE; // if true the borders are not set to 0 through a style !

  WPSelectPaperSize: function(PW, PH: Integer; var ID: Integer): Integer;
  WPONStartSpellcheck: TWPStartSpellcheckEvent; // Global Event!
  WPWriteOnlyFirstColor: Boolean;
  WPLogicalParagraphStartWithSpaces: Boolean; // Paragraph startes with 2 whitesapce
  WPInitFontCacheSize: Integer;
  WPReaderEXOptions : Integer = 0; // Bit1: Border=0->no Border, Bit2: Load Landscape, Bit3: Flip w/h
  GlobalValueUnit: TWPEditUnit;
  WPWYSIWYG_GlobalDisable : Boolean;
  {$IFDEF WPDEMO}
  WPBufferMode, WPVersionMode: Cardinal;
  {$ENDIF}
  WPEditFieldStart, WPEditFieldEnd,
    WPEditFieldStartAlias, WPEditFieldEndAlias: Char;
  WPEditAreaStyle: TOneWrtStyle;
  WPFontSizeTranslation: array[Byte] of Integer;

  FWPT_IsHexChar: array[Byte] of Boolean;
  FWPT_HexCharValue: array[Byte] of Byte;
  FWPT_HexCharValueHI: array[Byte] of Byte;

  {$ELSE}
var
  WPEditFieldStart: Char;
var
  WPEditFieldEnd: Char;

  {$ENDIF}

function WPMax(const a, b: Integer): Integer;
function WPMin(const a, b: Integer): Integer;
{$IFNDEF DELPHI5}
procedure FreeAndNil(var Obj);
{$ENDIF}

function WPHasNonHiddenTarget(CheckPar: PTParagraph; MovingDown: Boolean): boolean;
procedure WPCopyCelldef(var source, dest: TWPCell);

implementation { ///////////////////////////////////////////////// }

{$IFDEF WPDEMO}
uses WPNagFrm;
{$ENDIF}

function WPHasNonHiddenTarget(CheckPar: PTParagraph; MovingDown: Boolean): boolean;
begin
  {$IFDEF WPOLDSTYLE}
  Result := CheckPar <> nil;
  {$ELSE}
  while (CheckPar <> nil) and ((paprHidden in checkPar^.prop)
    or (papxHidden in checkPar^.prop)) do
    if MovingDown then
      CheckPar := CheckPar^.Next
    else
      CheckPar := CheckPar^.Prev;
  Result := CheckPar <> nil;
  {$ENDIF WPOLDSTYLE}
end;

procedure WPCopyCelldef(var source, dest: TWPCell);
begin
  dest.CWidthPC := source.CWidthPC;
  dest.CLeftPC := source.CLeftPC;
  dest.ColSpan := source.ColSpan;
  dest.CellStyle := source.CellStyle;
  dest.Format := source.Format;
  dest.Value := source.Value;
  dest.Changed := source.Changed;
  dest.Tabletag := source.Tabletag;
  dest.pRight := source.pRight;
  dest.pLeft := source.pLeft;
  dest.pBorderSpace := source.pBorderSpace;
  dest.Cellname := source.Cellname;
  dest.Command := source.Command;
end;

function WPMax(const a, b: Integer): Integer;
begin
  if a > b then
    Result := a
  else
    Result := b;
end;

function WPMin(const a, b: Integer): Integer;
begin
  if a < b then
    Result := a
  else
    Result := b;
end;

{$IFNDEF DELPHI5}

procedure FreeAndNil(var Obj);
var
  P: TObject;
begin
  P := TObject(Obj);
  TObject(Obj) := nil; // clear the reference before destroying the object
  P.Free;
end;
{$ENDIF}

function WPFormatStringForGroupbox(const source: string): string;
begin
  Result := source;
  while (Result <> '') and (Result[1] = #32) do
    Result := Copy(Result, 2, 255);
  while (Result <> '') and (Result[Length(Result)] = #32) do
    Result := Copy(Result, 1, Length(Result) - 1);
  while (Result <> '') and (Result[Length(Result)] = ':') do
    Result := Copy(Result, 1, Length(Result) - 1);
  Result := #32 + Result + #32;
end;

// -----------------------------------------------------------------------------
// We could use the borland ColorToIdent but this is quicker and exception save

//TODO: Use procedure GetColorValues(Proc: TGetStrProc);

function WPColorToString(Color: TColor): string;
var
  i: Integer;
  srgb: packed record
    peRed: Byte;
    peGreen: Byte;
    peBlue: Byte;
    peFlags: Byte;
  end;
begin
  for i := 0 to 15 do
    if WPSSStyleColors[i].Value = Color then
    begin
      Result := WPSSStyleColors[i].Name;
      exit;
    end;
  PINTEGER(@srgb)^ := ColorToRGB(Color and $00FFFFFF);
  FmtStr(Result, '#%.2x%.2x%.2x', [srgb.peRed, srgb.peGreen, srgb.peBlue]);
end;

function WPBufferToHex(p: PChar): Byte;
begin
  Result := 0;
  if p <> nil then
  begin
    if FWPT_IsHexChar[Byte(p^)] then
    begin
      Result := FWPT_HexCharValueHi[Byte(p^)]; inc(p);
    end;
    if FWPT_IsHexChar[Byte(p^)] then
    begin
      Result := Result + FWPT_HexCharValue[Byte(p^)];
    end;
  end;
end;

function WPStringToColor(S: string): TColor;
var
  i: Integer;
  peRed, peGreen, peBlue: Byte;
begin
  if s = '' then
    Result := clNone
  else
  begin
    // Accept the 'clXXXX syntax, too
    if (Length(s) > 2) and (s[1] = 'c') and (s[2] = 'l') then
    begin
      Result := StringToColor(s);
      exit;
    end;

    for i := 0 to 15 do
      if CompareText(WPSSStyleColors[i].Name, S) = 0 then
      begin
        Result := WPSSStyleColors[i].Value;
        exit;
      end;

    if (s[1] = '#') and (Length(s) = 7) then // #RRGGBB
    begin
      peRed := Byte(WPBufferToHex(@(s[2])));
      peGreen := Byte(WPBufferToHex(@(s[4])));
      peBlue := Byte(WPBufferToHex(@(s[6])));
      Result := TColor(RGB(peRed, peGreen, peBlue));
    end
    else if ((s[1] = '#') or (s[1] = '$')
      {$IFNDEF VER100} or (s[1] = HexDisplayPrefix){$ENDIF})
      and (Length(s) = 9) then // Delphi: #00BBGGRR   $00BBGGRR
    begin
      peBlue := Byte(WPBufferToHex(@(s[4])));
      peGreen := Byte(WPBufferToHex(@(s[6])));
      peRed := Byte(WPBufferToHex(@(s[8])));
      Result := TColor(RGB(peRed, peGreen, peBlue));
    end
    else
      Result := clNone;
  end;
end;


function WPCompareParAttr(pPar1, pPar2: PTParagraph; BorderAttr, SpecialAttr: Boolean): TWPStyleElementsPar;
var
  i: Integer;
begin
  Result := [];
  if pPar1^.spacebefore <> pPar2^.spacebefore then
    include(Result, wpseSpaceBefore);
  if not (paprIsTable in pPar1^.prop) and (pPar1^.spaceafter <> pPar2^.spaceafter) then
    include(Result, wpseSpaceBefore);
  if pPar1^.indentleft <> pPar2^.indentleft then
    include(Result, wpseIndentLeft);
  if pPar1^.indentfirst <> pPar2^.indentfirst then
    include(Result, wpseIndentFirst);
  if pPar1^.indentright <> pPar2^.indentright then
    include(Result, wpseIndentRight);
  if pPar1^.spacebetween <> pPar2^.spacebetween then
    include(Result, wpseSpaceBetween);
  if (paprMultSpacing in pPar1^.prop) and not (paprMultSpacing in pPar2^.prop) then
    include(Result, wpseMultSpaceBetween);
  if pPar1^.Align <> pPar2^.Align then
    include(Result, wpseAlign);
  if pPar2^.Shading = 0 then pPar2^.Shading := 100;
  if pPar1^.Shading = 0 then pPar1^.Shading := 100;
  if (pPar1^.Color <> pPar2^.Color) or (pPar1^.Shading <> pPar2^.Shading) then
  begin
    include(Result, wpseShading);
    include(Result, wpseColor);
  end;

  if BorderAttr then
  begin
    if pPar1^.Border.LineType <> pPar2^.Border.LineType then
      include(Result, wpseBrdLines);
    if pPar1^.Border.Thickness <> pPar2^.Border.Thickness then
      include(Result, wpseBrdWidth);
    if pPar1^.Border.HColor <> pPar2^.Border.HColor then
      include(Result, wpseBrdColor);
  end;

  if SpecialAttr then
  begin
    for i := 0 to TABWORDMAX do
      if pPar1^.Tabs[i] <> pPar2^.Tabs[i] then
      begin
        include(Result, wpseTabs);
        break;
      end;
    if pPar1^.Number <> pPar2^.Number then
      include(Result, wpseNumber);
    if pPar1^.NumLevel <> pPar2^.NumLevel then
      include(Result, wpseNumberLevel);
    if pPar1^.Border.Space <> pPar2^.Border.Space then
      include(Result, wpseBorderSpacing);
    if (paprIsOutline in pPar1^.prop) <> (paprIsOutline in pPar2^.prop) then
      include(Result, wpseIsOutline);
    if (paprKeep in pPar1^.prop) <> (paprKeep in pPar2^.prop) then
      include(Result, wpseKeepTogether);
    if (paprKeepN in pPar1^.prop) <> (paprKeepN in pPar2^.prop) then
      include(Result, wpseKeepNext);
  end;
end;

function WPCompareCharAttr(pa, pa2: PTAttr): TWPStyleElementsChar;
begin
  Result := [];
  if ((pa^.style * [afsBold]) <> (pa2^.Style * [afsBold])) then
    include(Result, wpseFontBold);
  if ((pa^.style * [afsItalic]) <> (pa2^.Style * [afsItalic])) then
    include(Result, wpseFontItalic);
  if ((pa^.style * [afsUnderline]) <> (pa2^.Style * [afsUnderline])) then
    include(Result, wpseFontUnderline);
  if (pa^.Font <> pa2^.Font) then
    include(Result, wpseFontFont);
  if (pa^.Size <> pa2^.Size) then
    include(Result, wpseFontSize);
  if (pa^.Color <> pa2^.Color) then
    include(Result, wpseFontColor);
  if (pa^.BGColor <> pa2^.BGColor) then
    include(Result, wpseFontBGColor);
end;


{ These	procedures check if a class was	already	registered }
var FLockClasses : TStringList;

procedure WPRegisterClass(AClass: TPersistentClass);
var i : Integer;
begin
  if FLockClasses<>nil then
  begin
    i:=FLockClasses.IndexOf(AClass.ClassName);
    if i>=0 then FLockClasses.Delete(i);
  end;

  if Classes.GetClass(AClass.ClassName) = nil then
    Classes.RegisterClass(AClass);
end;

procedure WPRegisterClasses(AClasses: array of TPersistentClass);
var
  I: Integer;
begin
  for I := Low(AClasses) to High(AClasses) do
    WPRegisterClass(AClasses[I]);
end;

procedure TWPCustomRtfEditClass.SetFocusValues(Always: Boolean);
begin
end;

procedure TWPCustomRtfEditClass.SetStyleCollection(x: TWPCustomStyleCollection);
var
  i: Integer;
begin
  if FStyleCollection <> x then
  begin
    if FStyleCollection <> nil then
    begin
      i := FStyleCollection.FControlledMemos.IndexOf(Self);
      if i >= 0 then FStyleCollection.FControlledMemos.Delete(i);
    end;
    FStyleCollection := x;
    if FStyleCollection <> nil then
    begin
      FStyleCollection.FreeNotification(Self);
      FStyleCollection.FControlledMemos.Add(Self);
    end;
  end;
end;

constructor TWPCustomStyleCollection.Create(aOwner: TComponent);
begin
  inherited Create(aOwner);
  FControlledMemos := TList.Create;
end;

destructor TWPCustomStyleCollection.Destroy;
begin
  FControlledMemos.Free;
  inherited Destroy;
end;


procedure WPAssignStyleParValues(Dest: PTParagraph; DestLocked: TWPStyleElementsPar;
  const Source: PTParagraph; SourceSel: TWPStyleElementsPar; SetDefault: Boolean);
var
  p: TWPStyleElementPar;
begin
  for p := wpseIndentLeft to wpse_TheLastItem do
    if not (p in DestLocked) then
    begin
      if p in SourceSel then
        case p of
          wpseIndentLeft: Dest^.indentleft := Source^.indentleft;
          wpseIndentRight: Dest^.indentright := Source^.indentright;
          wpseIndentFirst: Dest^.indentfirst := Source^.indentfirst;
          wpseSpaceBefore: Dest^.spacebefore := Source^.spacebefore;
          wpseMultSpaceBetween:
            Dest^.prop := (Dest^.prop - [paprMultSpacing]) + (Source^.prop * [paprMultSpacing]);
          wpseSpaceBetween: Dest^.spacebetween := Source^.spacebetween;
          wpseSpaceAfter: Dest^.spaceafter := Source^.spaceafter;
          wpseBrdLines: Dest^.border.LineType := Source^.border.LineType;
          wpseBorderSpacing: Dest^.border.space := Source^.border.space;
          wpseKeepTogether:
            Dest^.prop := (Dest^.prop - [paprKeep]) + (Source^.prop * [paprKeep]);
          wpseKeepNext:
            Dest^.prop := (Dest^.prop - [paprKeepN]) + (Source^.prop * [paprKeepN]);
          wpseBrdWidth: Dest^.border.Thickness := Source^.border.Thickness;
          wpseAlign: Dest^.Align := Source^.Align;
          wpseColor:
            begin
              Dest^.Color := Source^.Color;
              if Source^.Shading = 0 then
                Dest^.Shading := 100
              else
                Dest^.Shading := Source^.Shading;
            end;
          wpseBrdColor:
            begin
              Dest^.Border.HColor := Source^.Border.HColor;
              Dest^.Border.VColor := Source^.Border.VColor;
              Dest^.Border.VColorB := Source^.Border.VColorB;
              Dest^.Border.HColorR := Source^.Border.HColorR;
            end;
          wpseShading: Dest^.Shading := Source^.Shading;
          wpseTabs: Dest^.Tabs := Source^.Tabs;
          wpseNumber: begin Dest^.Number := Source^.Number; Dest^.NumLevel := Source^.NumLevel; end;
          wpseNumberLevel: begin Dest^.NumLevel := Source^.NumLevel; Dest^.Number := Source^.Number;
                                  if (Dest^.NumLevel=0) and (Dest^.Number<>0) then
                                     Dest^.Number := 0;
                           end; 
          wpseIsOutline: Dest^.prop := (Dest^.prop - [paprIsOutline]) + (Source^.prop * [paprIsOutline]);
        end
      else if SetDefault then // Default Setting
        case p of
          wpseIndentLeft: Dest^.indentleft := 0;
          wpseIndentRight: Dest^.indentright := 0;
          wpseIndentFirst: Dest^.indentfirst := 0;
          wpseSpaceBefore: Dest^.spacebefore := 0;
          wpseMultSpaceBetween: exclude(Dest^.prop, paprMultSpacing);
          wpseSpaceBetween: Dest^.spacebetween := 0;
          wpseSpaceAfter: Dest^.spaceafter := 0;
          wpseBrdLines:
            if not WPDontPresetBorder {or not (paprIsTable in Dest^.prop)} then
              Dest^.border.LineType := [];
          wpseBrdWidth:
            if not WPDontPresetBorder or not (paprIsTable in Dest^.prop) then
              Dest^.border.Thickness := 0;
          wpseAlign: Dest^.Align := paralLeft;
          wpseColor:
            begin
              Dest^.Color := 0;
            end;
          wpseBrdColor:
            if not WPDontPresetBorder or not (paprIsTable in Dest^.prop) then
            begin
              Dest^.Border.HColor := 0;
              Dest^.Border.VColor := 0;
              Dest^.Border.VColorB := 0;
              Dest^.Border.HColorR := 0;
            end;
          wpseShading:
            if not (wpseColor in SourceSel) then
              Dest^.Shading := 0;
          wpseTabs: FillChar(Dest^.Tabs, SizeOf(Dest^.Tabs), 0);
          wpseNumber: if not (wpseNumberLevel in SourceSel) then Dest^.Number := 0;
          wpseNumberLevel: if not (wpseNumber in SourceSel) then
          begin Dest^.Number := 0; Dest^.NumLevel := 0; end;
          wpseIsOutline: exclude(Dest^.prop, paprIsOutline);
        end;
    end;
end;

procedure WPAssignStyleAttrValues(Dest: PTAttr; DestLocked: WrtStyle;
  const Source: PTAttr; SourceSel: TWPStyleElementsChar; SetDefault: Boolean);
var
  a: TWPStyleElementChar;
begin
  for a := wpseFontBold to wpseFontBGColor do
    if not (WPStyleAttrElementLocking[a] in DestLocked) then
    begin
      if a in SourceSel then
        case a of
          wpseFontBold:
            Dest^.style := (Dest^.style - [afsBold]) + (Source^.Style * [afsBold]);
          wpseFontItalic:
            Dest^.style := (Dest^.style - [afsItalic]) + (Source^.Style * [afsItalic]);
          wpseFontUnderline:
            Dest^.style := (Dest^.style - [afsUnderline]) + (Source^.Style * [afsUnderline]);
          wpseFontFont: Dest^.font := Source^.Font;
          wpseFontSize: Dest^.Size := Source^.Size;
          wpseFontColor: Dest^.Color := Source^.Color;
          wpseFontBGColor: Dest^.BGColor := Source^.BGColor;
        end
      else if SetDefault then
        case a of
          wpseFontBold: Exclude(Dest^.Style, afsBold);
          wpseFontItalic: Exclude(Dest^.Style, afsItalic);
          wpseFontUnderline: Exclude(Dest^.Style, afsUnderline);
          //  wpseFontFont: Dest^.font := 0;
          //  wpseFontSize: Dest^.Size := 10;
          wpseFontColor: Dest^.Color := 0;
          wpseFontBGColor: Dest^.BGColor := 0;
        end;
    end;
end;

procedure TWPCustomRtfEditClass.MouseDown(Button: TMouseButton; Shift: TShiftState;
  X, Y: Integer);
begin
  inherited MouseDown(Button, Shift, x, y);
end;

procedure TWPCustomRtfEditClass.MouseUp(Button: TMouseButton; Shift: TShiftState;
  X, Y: Integer);
begin
  inherited MouseUp(Button, Shift, x, y);
end;

procedure TWPCustomRtfEditClass.MouseMove(Shift: TShiftState; X, Y: Integer);
begin
  inherited MouseMove(Shift, x, y);
end;

procedure TWPCustomToolCtrl.SetRTFedit(x: TWPCustomRtfEditClass);
begin
  if assigned(FRTFEditChanged) then FRTFEditChanged(Self, x);
end;

procedure TWPCustomToolCtrl.UpdateEnabledState;
begin
end;

// for easier code update
// Obsoletein V4
{ function WPLoadFormat(x: TWPLoadFormat): TWPNewLoadFormat;
begin
  case x of
    FormatRTF: Result := DefaultReaderWriter.DefaultReaderRTF;
    FormatANSI: Result := DefaultReaderWriter.DefaultReaderANSI;
    FormatHTML: Result := DefaultReaderWriter.DefaultReaderHTML;
  else
    Result := 'AUTO';
  end;
end;

function WPSaveFormat(x: TWPLoadFormat): TWPNewLoadFormat;
begin
  case x of
    FormatRTF: Result := DefaultReaderWriter.DefaultWriterRTF;
    FormatANSI: Result := DefaultReaderWriter.DefaultWriterANSI;
    FormatHTML: Result := DefaultReaderWriter.DefaultWriterHTML;
  else
    Result := 'AUTO';
  end;
end; }

{ ++++++++++++++++++++ WPTxtDef+++++++++++++++++++++++++++ }

{ TCharacterAttr }

constructor TCharacterAttr.Create(FCharStyle: TOneWrtStyle);
begin
  inherited Create;
  Self.FCharStyle := FCharStyle;
  FAlsoUseForPrintout := TRUE;
end;

procedure TCharacterAttr.SetBold(x: TThreeState);
begin
  if x <> FBold then
  begin
    FBold := x;
    if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], TRUE);
  end;
end;

procedure TCharacterAttr.SetItalic(x: TThreeState);
begin
  if x <> FItalic then
  begin
    FItalic := x; if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], TRUE);
  end;
end;

procedure TCharacterAttr.SetUnderline(x: TThreeState);
begin
  if x <> FUnderline then
  begin
    FUnderline := x; if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], FALSE);
  end;
end;

procedure TCharacterAttr.SetDoubleUnderline(x: Boolean);
begin
  if x <> FDoubleUnderline then
  begin
    FDoubleUnderline := x;
    if FDoubleUnderline then FUnderline := tsTRUE;
    if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], FALSE);
  end;
end;

procedure TCharacterAttr.SetStrikeOut(x: TThreeState);
begin
  if x <> FStrikeOut then
  begin
    FStrikeOut := x; if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], FALSE);
  end;
end;

procedure TCharacterAttr.SetSubScript(x: TThreeState);
begin
  if x <> FSubScript then
  begin
    FSubScript := x;
    if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], TRUE);
  end;
end;

procedure TCharacterAttr.SetSuperScript(x: TThreeState);
begin
  if x <> FSuperScript then
  begin
    FSuperScript := x;
    if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], TRUE);
  end;
end;

procedure TCharacterAttr.SetHidden(x: Boolean);
begin
  if x <> FHidden then
  begin
    FHidden := x;
    if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], TRUE);
  end;
end;

procedure TCharacterAttr.SetUseUnderlineColor(x: Boolean);
begin
  if x <> FUseUnderlineColor then
  begin
    FUseUnderlineColor := x;
    if (FUnderline = tsTRUE) and assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], FALSE);
  end;
end;

procedure TCharacterAttr.SetUseTextColor(x: Boolean);
begin
  if x <> FUseTextColor then
  begin
    FUseTextColor := x;
    if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], FALSE);
  end;
end;

procedure TCharacterAttr.SetUseBackgroundColor(x: Boolean);
begin
  if x <> FUseBackgroundColor then
  begin
    FUseBackgroundColor := x;
    if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], FALSE);
  end;
end;

procedure TCharacterAttr.SetUnderlineColor(x: TColor);
begin
  if x <> FUnderlineColor then
  begin
    FUnderlineColor := x;
    FUseUnderlineColor := TRUE;
    if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], FALSE);
  end;
end;

procedure TCharacterAttr.SetTextColor(x: TColor);
begin
  if x <> FTextColor then
  begin
    FTextColor := x;
    FUseTextColor := TRUE;
    if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], FALSE);
  end;
end;

procedure TCharacterAttr.SetBackgroundColor(x: TColor);
begin
  if x <> FBackgroundColor then
  begin
    FBackgroundColor := x;
    FUseBackgroundColor := TRUE;
    if assigned(FOnAttributeChanged) then
      FOnAttributeChanged([FCharStyle], FALSE);
  end;
end;

procedure TCharacterAttr.SetHotUnderlineColor(x: TColor);
begin
  FHotUnderlineColor := x;
  FUseHotUnderlineColor := TRUE;
  FHotStyleIsActive := TRUE;
end;

procedure TCharacterAttr.SetHotBackgroundColor(x: TColor);
begin
  FHotBackgroundColor := x;
  FUseHotBackgroundColor := TRUE;
  FHotStyleIsActive := TRUE;
end;

procedure TCharacterAttr.SetHotTextColor(x: TColor);
begin
  FHotTextColor := x;
  FUseHotTextColor := TRUE;
  FHotStyleIsActive := TRUE;
end;

procedure TCharacterAttr.SetHotEffect(x: TWPTextEffect);
begin
  FHotEffect := x;
  if x <> wpeffNone then
  begin
    if x = wpeffPopup then
      FHotStyleEffectColor := clBtnFace
    else if x = wpeffOutline then
      FHotStyleEffectColor := clYellow;
    FHotStyleIsActive := TRUE;
  end;
end;

procedure TCharacterAttr.Assign(Source: TPersistent);
begin
  if Source is TCharacterAttr then
  begin
    FBold := TCharacterAttr(Source).FBold;
    FItalic := TCharacterAttr(Source).FItalic;
    FUnderline := TCharacterAttr(Source).FUnderline;
    FStrikeOut := TCharacterAttr(Source).FStrikeOut;
    FSuperScript := TCharacterAttr(Source).FSuperScript;
    FSubScript := TCharacterAttr(Source).FSubScript;
    FHidden := TCharacterAttr(Source).FHidden;
    FUseUnderlineColor := TCharacterAttr(Source).FUseUnderlineColor;
    FUseTextColor := TCharacterAttr(Source).FUseTextColor;
    FUseBackgroundColor := TCharacterAttr(Source).FUseBackgroundColor;
    FUnderlineColor := TCharacterAttr(Source).FUnderlineColor;
    FTextColor := TCharacterAttr(Source).FTextColor;
    FBackgroundColor := TCharacterAttr(Source).FBackgroundColor;
    FHotUnderlineColor := TCharacterAttr(Source).FHotUnderlineColor;
    FUseHotUnderlineColor := TCharacterAttr(Source).FUseHotUnderlineColor;
    FHotBackgroundColor := TCharacterAttr(Source).FHotBackgroundColor;
    FUseHotBackgroundColor := TCharacterAttr(Source).FUseHotBackgroundColor;
    FHotTextColor := TCharacterAttr(Source).FHotTextColor;
    FUseHotTextColor := TCharacterAttr(Source).FUseHotTextColor;
    FHotUnderline := TCharacterAttr(Source).FHotUnderline;
    FHotStyleIsActive := TCharacterAttr(Source).FHotStyleIsActive;
    FAlsoUseForPrintout := TCharacterAttr(Source).FAlsoUseForPrintout;
  end
  else
    inherited Assign(Source);
end;

{ TTextHeader }

constructor TTextHeader.Create;
begin
  inherited Create;
  FPageGapWidth := 282;
  DefaultPageWidth := 12240;
  DefaultPageHeight := 15840;
  DefaultTopMargin := 1440;
  DefaultBottomMargin := 1440;
  DefaultRightMargin := 1880;
  DefaultLeftMargin := 1880;
  DefaultMarginHeader := 720;
  DefaultMarginFooter := 720;

  FPrintHeaderFooterInPageMargins := TRUE;
  FRoundTabs := TRUE;
  RoundTabsDivForInch := 90;
  RoundTabsDivForCm := 144;
  Clear;
  Format := 'AUTO';
  FLoadOptions := [loAddReadParagraphStyles,loOverwriteParagraphStyleAttr];
  FStoreOptions := [soAllStylesInCollection,soWordHyperLinkFormat,soWriteObjectsAsRTFBinary];
end;

destructor TTextHeader.Destroy;
begin
  RemoveAllTabs;
  inherited Destroy;
end;

function TTextHeader.AddFontName(nam: TFontName): Integer;
begin
  Result := AddFontNameCharset(nam, -1);
end;

function TTextHeader.AddFontNameCharset(nam: TFontName; charset: Integer): Integer;
var
  i: Integer;
begin
  i := 0;
  while i < FontMaxAnz do
  begin
    if (FontName[i] <> '') and
       (CompareText(FontName[i], nam) = 0) and
       (WPDontCheckFontCharsets or
       (charset = FontCharset[i]) or
       ((FontCharset[i] = -1)  and
        (charset=-1))
      ) then break;
    inc(i);
  end;

  if i<FontMaxAnz then
     Result := i  // V4.06
  else
  begin
    i := 0;
    while (i < FontMaxAnz) and (FontName[i]<>'') do inc(i);

    // We found an entry so we are adding the font name here
    if i < FontMaxAnz then
    begin
      FontName[i] := nam;
      if (CompareText(nam, 'wingdings') = 0) or
         (CompareText(nam, 'webdings') = 0) or
         (CompareText(nam, 'symbol') = 0) then
           FontCharset[i] := 2
      else
           FontCharset[i] := charset;
      Result := i;
    end
    else
      Result := 0;
  end;
end;

procedure TTextHeader.SetFontCharset(nr, charset: Integer);
begin
  if nr<FontMaxAnz then
  begin
    FontCharset[nr] := charset;
  end;
end;

{  deftab RTF tag }

procedure TTextHeader.SetTabDefault(x: Integer);
begin
  _Layout.deftabstop := x;
  papTabDefault := MulDiv(x, FFontXPixelsPerInch, 1440);
  UpdatePageInfo(true);
end;

function TTextHeader.GetTabDefault: Integer;
begin
  Result := _Layout.deftabstop;
end;

function TTextHeader.GetLandscape: Boolean;
begin
  Result := _Layout.landscape;
end;

function TTextHeader.GetMarginHeader: Longint;
begin
  Result := _Layout.marg_header; { Never swapped ! }
end;

function TTextHeader.GetMarginFooter: Longint;
begin
  Result := _Layout.marg_footer; { Never swapped ! }
end;

procedure TTextHeader.SetLandscape(x: Boolean);
begin
  if _Layout.Landscape <> x then
  begin
    SwapLayout;
    UpdatePageInfo(true);
  end;
end;

procedure TTextHeader.SetDefaultPageSize(x: TWPPageSettings);
var
  i, w, h: Integer;
begin
  if x <> wp_Custom then
  begin
    i := Integer(x);
    if WPPageSettings[i].UseCM then
    begin
      w := WPCentimeterToTwips(WPPageSettings[i].Width);
      h := WPCentimeterToTwips(WPPageSettings[i].Height);
    end
    else
    begin
      w := WPInchToTwips(WPPageSettings[i].Width);
      h := WPInchToTwips(WPPageSettings[i].Height);
    end;
    DefaultPageWidth := w;
    DefaultPageHeight := h;
  end;
  FDefaultPageSize := x;
end;


procedure TTextHeader.SetPageSize(x: TWPPageSettings);
var
  i, w, h: Integer;
begin
  if x <> wp_Custom then
  begin
    i := Integer(x);
    if WPPageSettings[i].UseCM then
    begin
      w := WPCentimeterToTwips(WPPageSettings[i].Width);
      h := WPCentimeterToTwips(WPPageSettings[i].Height);
    end
    else
    begin
      w := WPInchToTwips(WPPageSettings[i].Width);
      h := WPInchToTwips(WPPageSettings[i].Height);
    end;
    PageWidth := w;
    PageHeight := h;
  end;
  FPageSize := x;
end;

procedure TTextHeader.SwapLayout;
var
  c: Longint;
begin
  c := _Layout.paperw;
  _Layout.paperw := _Layout.paperh;
  _Layout.paperh := c;
  _Layout.landscape := not _Layout.landscape;
  RecalcLayout;
end;

procedure TTextHeader.SetMarginHeader(x: Longint);
begin
  _Layout.marg_header := x;
end;

procedure TTextHeader.SetMarginFooter(x: Longint);
begin
  _Layout.marg_footer := x;
end;

procedure TTextHeader.Clear; {	Set default }
var
  i: Integer;
const
  Col: array[0..15] of TColor =
  (clBlack, clRed, clGreen, clBlue, clYellow,
    clFuchsia, clPurple, clMaroon, clLime, clAqua, clTeal, clNavy,
    clWhite, clLtGray, clGray, clBlack);
begin
  for i := 1 to FontMaxAnz - 1 do
    FontName[i] := ''; { dont delete Font #0 }
  for i := 0 to TABMAX - 1 do
    TabIndex[i] := i;

  { _Layout.psz	:= 1;  }
  if DefaultPageWidth < DefaultLeftMargin + DefaultRightMargin + 100 then
    DefaultPageWidth := DefaultLeftMargin + DefaultRightMargin + 100;
  if DefaultPageHeight < DefaultTopMargin + DefaultBottomMargin + 100 then
    DefaultPageHeight := DefaultTopMargin + DefaultBottomMargin + 100;
  _Layout.paperw := DefaultPageWidth;
  _Layout.paperh := DefaultPageHeight;
  _Layout.margl := DefaultLeftMargin;
  _Layout.margr := DefaultRightMargin;
  _Layout.margt := DefaultTopMargin;
  _Layout.margb := DefaultBottomMargin;
  _Layout.marg_header := FDefaultMarginHeader;
  _Layout.marg_footer := FDefaultMarginFooter;
  if DefaultLandscape then SwapLayout;
  RecalcLayout;
  GetPaletteEntries(GetStockObject(DEFAULT_PALETTE), 0, NumPaletteEntries,
    FPaletteEntries);
  fillchar(FPaletteEntries[0], SizeOf(FPaletteEntries), 0);
  for i := 0 to 15 do
  begin
    Longint(FPaletteEntries[i]) := ColorToRGB(Col[i]);
    FPaletteEntries[i].peFlags := 1;
  end;
  if _Layout.deftabstop = 0 then
  begin
    DefaultTabstop := 720; // RTF Default !
   { if FUnit = UnitCm then
      DefaultTabstop := 254
    else
      DefaultTabstop := 360; 	0,25 Inch }
  end;
  NewTabKind := tkLeft;
end;

function TTextHeader.AddColor(Color: TColor): Integer;
begin
  if not FindMatchingColor(Color, TRUE, TRUE, Result) then Result := 0;
end;

function TTextHeader.FindMatchingColor(c: TColor;
  Use256Colors, AddColors: Boolean; var nr: Integer): Boolean;
var
  isfree, r, g, b, i, diff, f, t: Integer;
  rgb: TPaletteEntry;
begin
  Result := FALSE;
  nr := -1;
  isfree := 0;
  if Use256Colors then
  begin
    f := 16; t := 255;
  end
  else
  begin
    f := 0; t := 15;
  end;
  diff := $7FFFFFFF;
  rgb := TPaletteEntry(ColorToRGB(c));
  if (c = clBlack) or (c = wpClWindowText) then
  begin
    nr := 0; Result := TRUE; exit;
  end;

  for i := 0 to t do
  begin
    // We protect the first 15 Colors. They are fixed !!!
    if (not Use256Colors or (i > 15)) and
      (FPaletteEntries[i].peFlags = 0) then
    begin
      inc(isfree);
      continue;
    end;
    r := FPaletteEntries[i].peRed - rgb.peRed;
    g := FPaletteEntries[i].peGreen - rgb.peGreen;
    b := FPaletteEntries[i].peBlue - rgb.peBlue;
    r := r * r + g * g + b * b;
    if r < diff then
    begin
      nr := i;
      diff := r;
      if diff = 0 then break;
    end;
  end;
  // Distance small enough or full then use it
  if (isfree = 0) or (diff <= 256) then
  begin
    Result := TRUE;
    FPaletteEntries[nr].peFlags := 1;
    exit;
  end;
  // Add this Color
  for i := f + 1 to t do
    if FPaletteEntries[i].peFlags = 0 then
    begin
      nr := i;
      Result := TRUE;
      // FPaletteUsed[i] := TRUE;
      FPaletteEntries[i] := rgb;
      FPaletteEntries[i].peFlags := 1;
      break;
    end;
end;

procedure TTextHeader.RecalcLayout;
var
  rx, ry: Integer;
begin
  if _Layout.Landscape then
  begin
    rx := FFontYPixelsPerInch;
    ry := FFontXPixelsPerInch;
  end
  else
  begin
    rx := FFontXPixelsPerInch;
    ry := FFontYPixelsPerInch;
  end;
  LayoutPIX.paperw := MulDiv(_Layout.paperw, rx, 1440);
  LayoutPIX.paperh := MulDiv(_Layout.paperh, ry, 1440);
  LayoutPIX.margr := MulDiv(_Layout.margr, rx, 1440);
  LayoutPIX.margl := MulDiv(_Layout.margL, rx, 1440);
  LayoutPIX.margt := MulDiv(_Layout.margT, ry, 1440);
  LayoutPIX.margb := MulDiv(_Layout.margB, ry, 1440);
  if _Layout.marg_header < 254 then _Layout.marg_header := 254;
  if _Layout.marg_footer < 254 then _Layout.marg_footer := 254;
  LayoutPIX.marg_header := MulDiv(_Layout.marg_header, ry, 1440);
  LayoutPIX.marg_footer := MulDiv(_Layout.marg_footer, ry, 1440);
  papWidth := LayoutPIX.paperw - LayoutPIX.margl - LayoutPIX.margr;
  papHeight := LayoutPIX.paperh - LayoutPIX.margt - LayoutPIX.margb;
end;

procedure TTextHeader.BeginUpdate;
begin
  FLocked := TRUE;
end;

procedure TTextHeader.EndUpdate;
begin
  FLocked := FALSE;
  UpdatePageInfo(true);
end;

procedure TTextHeader.UpdatePageInfo(Reformat : Boolean);
begin
  if not FLocked then
  begin
    RecalcLayout;
    if assigned(FUpdatePageInfo) then FUpdatePageInfo(Reformat); // Reformat
  end;
end;

procedure TTextHeader.SetPageWidth(x: Longint);
begin
  if _Layout.landscape then
    _Layout.paperh := x
  else
    _Layout.paperw := x;
  FDefaultPageSize := wp_Custom;
  UpdatePageInfo(true);
end;

function TTextHeader.GetPageWidth: Longint;
begin
  if _Layout.landscape then
    Result := _Layout.paperh
  else
    Result := _Layout.paperw;
end;

procedure TTextHeader.SetPageHeight(x: Longint);
begin
  if _Layout.landscape then
    _Layout.paperw := x
  else
    _Layout.paperh := x;
  FDefaultPageSize := wp_Custom;
  UpdatePageInfo(true);
end;

function TTextHeader.GetPageHeight: Longint;
begin
  if _Layout.landscape then
    Result := _Layout.paperw
  else
    Result := _Layout.paperh;
end;

procedure TTextHeader.SetLogPageWidth(x: Longint);
begin
  _Layout.paperw := x;
  UpdatePageInfo(true);
end;

function TTextHeader.GetLogPageWidth: Longint;
begin
  Result := _Layout.paperw;
end;

procedure TTextHeader.SetLogPageHeight(x: Longint);
begin
  _Layout.paperh := x;
  UpdatePageInfo(true);
end;

function TTextHeader.GetLogPageHeight: Longint;
begin
  Result := _Layout.paperh;
end;

procedure TTextHeader.SetLeftMargin(x: Longint);
begin
  _Layout.margl := x;
  UpdatePageInfo(true);
end;

procedure TTextHeader.SetMarginMirror(x: Boolean);
begin
  _Layout.marginmirror := x;
  UpdatePageInfo(true);
end;

function TTextHeader.GetMarginMirror: Boolean;
begin
  Result := _Layout.marginmirror;
end;

function TTextHeader.GetLeftMargin: Longint;
begin
  Result := _Layout.margl;
end;

procedure TTextHeader.SetRightMargin(x: Longint);
begin
  _Layout.margr := x;
  UpdatePageInfo(true);
end;

function TTextHeader.GetRightMargin: Longint;
begin
  Result := _Layout.margr;
end;

procedure TTextHeader.SetTopMargin(x: Longint);
begin
  _Layout.margt := x;
  UpdatePageInfo(true);
end;

function TTextHeader.GetTopMargin: Longint;
begin
  Result := _Layout.margt;
end;

procedure TTextHeader.SetBottomMargin(x: Longint);
begin
  _Layout.margb := x;
  UpdatePageInfo(true);
end;

function TTextHeader.GetBottomMargin: Longint;
begin
  Result := _Layout.margb;
end;

function TTextHeader.RoundTwips(tw: Longint): Longint;
var
  w: Integer;
begin
  w := PageWidth - LeftMargin - RightMargin;
  if FRoundTabs and (FFontXPixelsPerInch > 0) then
  begin
    if (FUnit <> UnitCm) and (FRoundTabsDivForInch <> 0) then
    begin // V3.08 -- very small values -> 1st step
      if (tw < FRoundTabsDivForInch) and (tw > 4) then
        tw := FRoundTabsDivForInch + 1
      else
        inc(tw, FRoundTabsDivForInch div 2);
      Result := tw - (tw mod FRoundTabsDivForInch);
      if abs(Result - w) <= FRoundTabsDivForInch then
        Result := w;
    end
    else if (FUnit = UnitCm) and (FRoundTabsDivForCm <> 0) then
    begin // V3.08 -- very small values -> 1st step
      if (tw < FRoundTabsDivForCm) and (tw > 4) then
        tw := FRoundTabsDivForCm + 1
      else
        inc(tw, FRoundTabsDivForCm div 2);
      Result := tw - (tw mod FRoundTabsDivForCm);
      if abs(Result - w) <= FRoundTabsDivForCm then
        Result := w;
    end
    else
      Result := tw;
  end
  else
    Result := tw;
end;

{ Transfer of the varaibles }

procedure TTextHeader.CopyLayoutData(p: TTextHeader);
var
  i: Integer;
begin
  for i := 0 to NumPaletteEntries do
  begin
    FPaletteEntries[i] := p.FPaletteEntries[i];
  end;
  _Layout := p._Layout;
  FDefaultPageWidth:= p.FDefaultPageWidth;
  FDefaultPageHeight:= p.FDefaultPageHeight;
  FDefaultTopMargin:= p.FDefaultTopMargin;
  FDefaultBottomMargin:= p.FDefaultBottomMargin;
  FDefaultRightMargin:= p.FDefaultRightMargin;
  FDefaultLeftMargin:= p.FDefaultLeftMargin;
  FDefaultLandscape:= p.FDefaultLandscape;
  FDefaultMarginHeader:= p.FDefaultMarginHeader;
  FDefaultMarginFooter:= p.FDefaultMarginFooter;
  FDefaultPageSize:= p.FDefaultPageSize;
  papTabDefault := MulDiv(_Layout.deftabstop, FFontXPixelsPerInch, 1440);
  FPrintHeaderFooterInPageMargins := p.FPrintHeaderFooterInPageMargins;
  UpdatePageInfo(true);
end;

procedure TTextHeader.Assign(Source: TPersistent);
var
  i: Integer;
  p: TTextHeader;
begin
  if Source is TTextHeader then
  begin
    HTMLStoreOptions := TTextHeader(Source).HTMLStoreOptions;
    StoreOptions := TTextHeader(Source).StoreOptions;
    LoadOptions := TTextHeader(Source).LoadOptions;
    p := TTextHeader(Source);
    Clear;
    if p <> nil then
    begin
      for i := 0 to FontMaxAnz - 1 do
        if p.FontName[i] <> '' then { dont't make text inusable }
        begin
          FontName[i] := p.FontName[i];
          FontFamily[i] := p.FontFamily[i];
          FontCharSet[i] := p.FontCharSet[i];
        end;

      for i := 0 to TABMAX - 1 do
      begin
        TabPosTw[i] := p.TabPosTw[i];
        TabPos[i] := p.TabPos[i];
        TabKind[i] := p.TabKind[i];
        TabIndex[i] := p.TabIndex[i];
      end;

      CopyLayoutData(p);
    end;
  end
  else
    inherited Assign(Source);
end;

function TTextHeader.FindMatchingTab(pos: Longint; kind: TTabKind; map:
  TTabFlag): Integer;
var
  i: Integer;
begin
  Result := -1;
  i := 0;
  while i < TABMAX do
  begin
    if (TabPosTw[i] = pos) and ((kind = tkAll) or (TabKind[i] = kind))
      and ((map[i div 16] and TabsRef[i mod 16]) <> 0) then
    begin
      Result := i;
      exit;
    end;
    inc(i);
  end;
  if not FDontRoundNextTab then
  begin
    pos := RoundTwips(pos);
    i := 0;
    while i < TABMAX do
    begin
      if (TabPosTw[i] = pos) and ((kind = tkAll) or (TabKind[i] = kind))
        and ((map[i div 16] and TabsRef[i mod 16]) <> 0) then
      begin
        Result := i;
        exit;
      end;
      inc(i);
    end;
  end;
end;

// Alternative: Compare with abs() formula and do not round the position !

function TTextHeader.FindTab(var pos: Longint; kind: TTabKind): Integer;
var
  i: Integer;
begin
  Result := -1;
  i := 0;
  while i < TABMAX do
  begin
    if (TabPosTw[i] = pos) and ((kind = tkAll) or (TabKind[i] = kind)) then
    begin
      Result := i;
      exit;
    end;
    inc(i);
  end;
  if not FDontRoundNextTab then
  begin
    pos := RoundTwips(pos);
    i := 0;
    while i < TABMAX do
    begin
      if (TabPosTw[i] = pos) and ((kind = tkAll) or (TabKind[i] = kind)) then
      begin
        Result := i;
        exit;
      end;
      inc(i);
    end;
  end;
end;

{ Result is index of new tab or	-1 }

function TTextHeader.AddTab(pos: Longint; kind: TTabKind): Integer;
begin
  fillchar(ResultTABS, Sizeof(ResultTABS), 0);
  Result := InternAddTab(pos, kind);
end;

function TTextHeader.InternAddTab(pos: Longint; kind: TTabKind): Integer;
{ pos = Position in twips }
var
  i: Integer;
begin
  Result := -1;
  try
    if pos <= 0 then exit;
    i := FindTab(pos, kind);
    if i >= 0 then
    begin
      ResultTABS[i div 16] := ResultTABS[i div 16] or TabsRef[i mod 16];
      Result := i;
      exit;
    end;
    { search for	free place and add new tab }
    i := 0;
    while i < TABMAX do
    begin
      if TabPosTw[i] = 0 then
      begin
        TabPosTw[i] := pos;
        TabPos[i] := MulDiv(pos, FFontXPixelsPerInch, 1440);
        if kind = tkAll then
          TabKind[i] := NewTabKind
        else
          TabKind[i] := Kind;
        ResultTABS[i div 16] := ResultTABS[i div 16] or TabsRef[i mod 16];
        SortTabs;
        Result := i;
        exit;
      end;
      inc(i);
    end;
    { tries to delete positions not needed }
    if assigned(FOnCheckTabs) then
    begin
      FOnCheckTabs;
      i := 0;
      while i < TABMAX do
      begin
        if TabPosTw[i] = 0 then
        begin
          TabPosTw[i] := pos;
          TabPos[i] := MulDiv(pos, FFontXPixelsPerInch, 1440);
          if kind = tkAll then
            TabKind[i] := NewTabKind
          else
            TabKind[i] := Kind;
          ResultTABS[i div 16] := ResultTABS[i div 16] or TabsRef[i mod 16];
          SortTabs;
          Result := i;
          exit;
        end;
        inc(i);
      end;
    end;
  finally
    FDontRoundNextTab := FALSE;
  end;
end;

procedure TTextHeader.ReCalculateTabs;
var
  i: Integer;
begin
  papTabDefault := MulDiv(_Layout.deftabstop, FFontXPixelsPerInch, 1440);
  for i := 0 to TABMAX - 1 do
    TabPos[i] := MulDiv(TabPosTw[i], FFontXPixelsPerInch, 1440);
  papWidth := { <<<<< }
  MulDiv(_Layout.paperw - _Layout.margl - _Layout.margr, FFontXPixelsPerInch,
    1440);
  papHeight :=
    MulDiv(_Layout.paperh - _Layout.margt - _Layout.margb, FFontYPixelsPerInch,
    1440);
end;

procedure TTextHeader.SortTabs;
var
  cc: Integer;
  function A(IND: INTEGER): LONGINT;
  begin
    if ind >= TABMAX then
    begin
      Result := $FFFFFF; exit;
    end
    else if ind < 0 then
    begin
      Result := 0; exit;
    end;
    Result := TabPosTw[TabIndex[ind]];
    if Result = 0 then Result := $FFFFFF;
  end;
  procedure SORT;
  var
    i, ww: Integer;
    ch: Boolean;
  begin
    repeat
      i := 0;
      ch := FALSE;
      while i < TABMAX - 1 do
      begin
        if a(i) > a(i + 1) then
        begin
          ww := TabIndex[i];
          TabIndex[i] := TabIndex[i + 1];
          TabIndex[i + 1] := ww;
          ch := TRUE;
        end;
        inc(i);
      end;
    until not ch;
  end; { Sort }
begin
  for cc := 0 to TABMAX - 1 do
    TabIndex[cc] := cc;
  Sort; { DON NOT SORT TABS! Sort only the	index }
end;


function TTextHeader.RemoveTab(pos: Longint; kind: TTabKind): Longint;
{	pos in twips }
var
  i: Integer;
begin
  fillchar(ResultTABS, Sizeof(ResultTABS), 0);
  i := FindTab(pos, kind);
  if i >= 0 then
  begin
    TabPosTw[i] := 0;
    TabPos[i] := 0;
    TabKind[i] := tkLeft;
    ResultTABS[i div 16] := TabsRef[i mod 16];
  end;
  Result := (PLongint(@ResultTABS[0]))^;
end;

procedure TTextHeader.RemoveAllTabs;
var
  i: Integer;
begin
  i := 0;
  while i < TABMAX do
  begin
    TabPosTw[i] := 0;
    TabPos[i] := 0;
    TabKind[i] := tkLeft;
    inc(i);
  end;
  for i := 0 to TABMAX - 1 do
    TabIndex[i] := i;
end;

function TTextHeader.AllTabs: Longint;
var
  i: Integer;
begin
  fillchar(ResultTABS, Sizeof(ResultTABS), 0);
  i := 0;
  while i < TABMAX do
  begin
    if TabPosTw[i] <> 0 then
      ResultTABS[i div 16] := ResultTABS[i div 16] or TabsRef[i mod 16];
    inc(i);
  end;
  Result := (PLongint(@ResultTABS[0]))^;
end;

procedure TWPCustomToolCtrl.SetPreviewButtons;
begin
end;

function TWPCustomTextReader.GetPar(ptag: Pchar): PChar;
begin
  Result := nil;
end;

constructor TWPCustomTextReader.Create;
begin
  inherited Create;
end;

constructor TWPCustomTextWriter.Create;
begin
  inherited Create;
end;

procedure TWPCustomTextWriter.Init;
begin
  LastError := ecOK;
end;

function TWPCustomTextWriter.WriteHeader(fp: TStream): ecErrCode;
begin
  Result := ecOK;
end;

function TWPCustomTextWriter.WriteText(fp: TStream): ecErrCode;
begin
  Result := ecOK;
end;

function TWPCustomTextWriter.WriteFooter(fp: TStream): ecErrCode;
begin
  Result := ecOK;
end;

{ Useful Procedure; Faster then the Libary
  Both write to buffer and change the pointer }

procedure TWPCustomTextWriter.IntToBuffer(var p: pchar; val: Longint); {  puts Integer to Buffer }
var
  oval: Longint;
begin
  oval := val;
  if oval > 99999 then
  begin
    p^ := Char((val div 100000) + Integer('0'));
    inc(p);
    val := val mod 100000;
  end;
  if oval > 9999 then
  begin
    p^ := Char((val div 10000) + Integer('0'));
    inc(p);
    val := val mod 10000;
  end;
  if oval > 999 then
  begin
    p^ := Char((val div 1000) + Integer('0'));
    inc(p);
    val := val mod 1000;
  end;
  if oval > 99 then
  begin
    p^ := Char((val div 100) + Integer('0'));
    inc(p);
    val := val mod 100;
  end;
  if oval > 9 then
  begin
    p^ := Char((val div 10) + Integer('0'));
    inc(p);
    val := val mod 10;
  end;
  p^ := Char(val + Integer('0'));
  inc(p);
end;

procedure TWPCustomTextWriter.HexToBuffer(var p: pchar; val: char);
const
  Hex: PChar = '0123456789abcdef';
begin
  p^ := Hex[Ord(val) shr 4]; inc(p);
  p^ := Hex[Ord(val) and 15]; inc(p);
end;

function TWPCustomTextWriter.WriteHexBuffer(mem: pointer; size: Longint): ecErrCode;
var
  pp_to, pp: PChar;
  l, hl: Longint;
  buf: PChar;
  temp: Pointer;
const
  Hex: PChar = '0123456789abcdef';
begin
  Result := ecOk;
  l := size;
  size := size * 2;
  hl := size;
  buf := nil;
  if l > 0 then
  try
    GetMem(temp, Size);
    buf := temp;
    pp_to := buf;
    pp := mem;
    inc(pp_to, hl - 1);
    inc(pp, l - 1);
    while TRUE do
    begin
      pp_to^ := Hex[Ord(pp^) and 15]; dec(pp_to);
      pp_to^ := Hex[Ord(pp^) shr 4];
      dec(l);
      if l <= 0 then break;
      dec(pp_to);
      dec(pp);
    end;
    if fpOut.Write(buf^, size) <> size then Result := ecEndOfFile;
  finally
    if buf <> nil then
    begin
      System.FreeMem(buf, size)
    end;
  end;
end;

// -----------------------------------------------------------------------------
// TWPCustomRTFFiler -----------------------------------------------------------
// -----------------------------------------------------------------------------

 function  TWPCustomRTFFiler.GetCurrentPage: Integer;
 begin Result := -1; end;
 function  TWPCustomRTFFiler.GetPageCount: Integer;
 begin Result := -1; end;
 procedure TWPCustomRTFFiler.SetCurrentPage(x: Integer); begin end;
 procedure TWPCustomRTFFiler.ReadStream(s: TStream); begin end;
 procedure TWPCustomRTFFiler.ReadFile(const fname: string); begin end;
 function  TWPCustomRTFFiler.Find(const text: string; wildcard: Char; CaseSensitive: Boolean):Boolean;
 begin Result := FALSE; end;
 procedure TWPCustomRTFFiler.Close; begin end;
 procedure TWPCustomRTFFiler.Print; begin end;

// -----------------------------------------------------------------------------
// Global Routines -------------------------------------------------------------
// -----------------------------------------------------------------------------

function WPUndoString(kind: TWPUndoKind): string;
begin   
  case Kind of
    wputNone: Result := WPLoadStr(meutNone);
    wputAny: Result := WPLoadStr(meutAny);
    wputInput: Result := WPLoadStr(meutInput);
    wputDeleteText: Result := WPLoadStr(meutDeleteText);
    wputChangeAttributes: Result := WPLoadStr(meutChangeAttributes);
    wputChangeIndent: Result := WPLoadStr(meutChangeIndent);
    wputChangeSpacing: Result := WPLoadStr(meutChangeSpacing);
    wputChangeAlignment: Result := WPLoadStr(meutChangeAlignment);
    wputChangeTabs: Result := WPLoadStr(meutChangeTabs);
    wputChangeBorder: Result := WPLoadStr(meutChangeBorder);
    wputDeleteSelection: Result := WPLoadStr(meutDeleteSelection);
    wputDragAndDrop: Result := WPLoadStr(meutDragAndDrop);
    wputPaste: Result := WPLoadStr(meutPaste);
    wputInsertObject : Result := WPLoadStr(meutChangeGraphic);
    wputChangeObject: Result := WPLoadStr(meutChangeGraphic);
    wputDeleteObject: Result := WPLoadStr(meutChangeGraphic);
    wputChangeTable: Result := WPLoadStr(meutChangeTable);
    wputChangeStyleSheet: Result := WPLoadStr(meutUpdateStyle);
    wputRedo: Result := WPLoadStr(meutAny);
  else
    Result := WPLoadStr(meutAny);
  end;
end;

function WPLoadStr(item: TWPVCLString): string;
begin
  Result := FWPVCLStrings[item];
end;

// -----------------------------------------------------------------------------

procedure WPPrinterOpen;
begin
  WPPrinterOpenName('');
end;

procedure WPPrinterOpenName(pname: string);
var
  ADevice, ADriver, APort: array[0..128] of char;
  ADeviceMode: THandle;
  hd: HDC;
  i: Integer;
begin
  FNoPrinterInstalled := (WPTPrinter.Printers.Count <= 0);
  if (WPPrinterCanvas = nil) or FNoPrinterInstalled then exit;
  if FWPPrinterCanvasHasHandle then WPPrinterClose;
  try
    if pname <> '' then
    begin
      if CompareText(pname, 'none') = 0 then exit;
      i := WPTPrinter.Printers.IndexOf(pname);
      if i >= 0 then WPTPrinter.PrinterIndex := i;
    end;
    WPTPrinter.GetPrinter(ADevice, ADriver, APort, ADeviceMode);

    if ADevice = '' then
    begin
      FNoPrinterInstalled := TRUE;
      exit;
    end;

    // WAS:  hd := CreateDC(ADriver, ADevice, NIL, nil);
    if ADevice <> '' then
      hd := CreateDC(ADriver, ADevice, nil, nil)
    else
      hd := CreateDC(ADriver, nil, nil, nil);

    if hd <> 0 then
    begin
      i := GetDeviceCaps(hd, TECHNOLOGY);
      PrinterCanvasPX := GetDeviceCaps(hd, LOGPIXELSX);
      PrinterCanvasPY := GetDeviceCaps(hd, LOGPIXELSY);
      if (i = DT_CHARSTREAM) then
      begin
        { ShowMessage('DT_CHARSTREAM ->	No WYSIWYG'); }
        DeleteDC(hd);
        exit;
      end;
      WPPrinterHDC := hd;
      WPPrinterName := StrPas(ADevice);
      WPPrinterCanvas.Handle := hd;
      FWPPrinterCanvasHasHandle := TRUE;
    end
    else
      WPPrinterHDC := 0;
  except
    FNoPrinterInstalled := TRUE;
  end;
end;

procedure WPPrinterClose;
begin
  try
    if WPPrinterCanvas = nil then exit;
    if FWPPrinterCanvasHasHandle then
    begin
      WPPrinterCanvas.Handle := 0;
      DeleteDC(WPPrinterHDC);
      FWPPrinterCanvasHasHandle := FALSE;
      WPPrinterName := '';
      WPPrinterHDC := 0;
    end;
  except
  end;
end;

function iWPSelectPaperSize(PW, PH: Integer; var ID: Integer): Integer; { Get DM Code }
var
  i, w, h: Integer;
begin
  i := MaxWPPageSettings;
  // Fix wrong DinA4 width: 29.6 cm -------------------------
  if abs(PH - WPCentimeterToTwips(29.6)) < 10 then
    PH := WPCentimeterToTwips(29.7);
  if abs(PW - WPCentimeterToTwips(29.6)) < 10 then
    PW := WPCentimeterToTwips(29.7);
  // -------------------------------------------------------
  while i > 0 do
  begin
    if WPPageSettings[i].UseCM then
    begin
      w := WPCentimeterToTwips(WPPageSettings[i].Width);
      h := WPCentimeterToTwips(WPPageSettings[i].Height);
    end
    else
    begin
      w := WPInchToTwips(WPPageSettings[i].Width);
      h := WPInchToTwips(WPPageSettings[i].Height);
    end;
    if (abs(pw - w) < 30) and (abs(ph - h) < 30) then break;
    dec(i);
  end;
  ID := i;
  Result := WPPageSettings[i].ID;
end;

procedure TWPCustomToolCtrl.OnColorDropDown(Sender: TObject);
var
  i: Integer;
  pal: PTColorPalette;
begin
  if (VFPPaletteEntries <> nil) and (Sender is TComboBox) then
  begin
    TComboBox(Sender).Items.Clear;
    pal := VFPPaletteEntries;
    for i := 0 to NumPaletteEntries do
    begin
      if (i <= 15) or (pal^.peFlags <> 0) then
      begin
        TComboBox(Sender).Items.AddObject(IntToStr(i), TObject(i));
      end;
      inc(pal);
    end;
  end;
end;

{ Global conversion  functions }

function WPCentimeterToTwips(value: Double): Longint;
begin
  Result := Round(value * 566.929);
end;

function WPCentimeterToPixel(value: Double; Resolution: Integer): Longint;
begin
  Result := Round(value * 566.929 / 1440 * Resolution);
end;

function WPInchToTwips(value: Double): Longint;
begin
  Result := Round(value * 1440);
end;

function WPInchToPixel(value: Double; Resolution: Integer): Longint;
begin
  Result := Round(value * Resolution);
end;

function WPTwipsToCentimeter(value: Longint): Double;
begin
  Result := value / 566.929;
end;

function WPTwipsToInch(value: Longint): Double;
begin
  Result := value / 1440;
end;

function WPListIntItemCompare(Item1, Item2: Pointer): Integer;
begin
  Result := Integer(Item1) - Integer(Item2);
end;

procedure WPSubstChar(var S: string; OldC, NewC: Char);
var
  p: PChar;
begin
  p := Pchar(s);
  while p^ <> #0 do
  begin
    if p^ = OldC then p^ := NewC;
    inc(p);
  end;
end;



// -----------------------------------------------------------------------------
var
  prevWPLocalSaveVCLStrings: procedure;
  prevWPLocalLoadVCLStrings: procedure;
  // -----------------------------------------------------------------------------


procedure WPSaveActionCaptions;
var
  i: TWPVCLString;
  typ: Pointer;
begin
  typ := TypeInfo(TWPVCLString);
  if (typ <> nil) and Assigned(WPLangSaveString) then
    for i := Low(TWPVCLString) to High(TWPVCLString) do
      WPLangSaveString('WPString/' + GetEnumName(typ, Integer(i)), FWPVCLStrings[i], 0);
  if assigned(prevWPLocalSaveVCLStrings) then prevWPLocalSaveVCLStrings;
end;

procedure WPLoadActionCaptions;
var
  i: TWPVCLString;
  a: Integer;
  s, aa, st: string;
  typ: Pointer;
begin
  s := '';
  aa := FWPVCLStrings[meDefaultUnit_INCH_or_CM];
  typ := TypeInfo(TWPVCLString);
  if (typ <> nil) and Assigned(WPLangLoadString) then
    for i := Low(TWPVCLString) to High(TWPVCLString) do
    begin
      st := GetEnumName(typ, Integer(i));
      if WPLangLoadString('WPString/' + st, s, a) then
         FWPVCLStrings[i] := s;
    end;
  // Update global language, too
  if CompareText(FWPVCLStrings[meDefaultUnit_INCH_or_CM], 'INCH') = 0 then
      GlobalValueUnit := euInch
  else if CompareText(FWPVCLStrings[meDefaultUnit_INCH_or_CM], 'CM') = 0 then
      GlobalValueUnit := euCm;

  if assigned(prevWPLocalLoadVCLStrings) then prevWPLocalLoadVCLStrings;
end;

var
  ib: Byte;
  val: Cardinal;
  loopb, loopi: Integer;

initialization
  GlobalValueUnit := euCM;
  prevWPLocalSaveVCLStrings := WPLocalSaveVCLStrings;
  WPLocalSaveVCLStrings := WPSaveActionCaptions;
  prevWPLocalLoadVCLStrings := WPLocalLoadVCLStrings;
  WPLocalLoadVCLStrings := WPLoadActionCaptions;
  wpClWindowText := clBlack;
  WPEditFieldStart := #171; // <<
  WPEditFieldEnd := #187; // >>
  WPEditAreaStyle := afsIsInsertpoint;
  wpMisspelledTextColor := clRed;
  WPSelectPaperSize := iWPSelectPaperSize;
  DefaultReaderWriter := TDefaultReaderWriter.Create;
  DefaultDecimalTabChar := DecimalSeparator;
  loopi := 0;
  val := 1; // abused variables ...
  while loopi < 32 do
  begin
    TabsRef[loopi] := val;
    val := val shl 1;
    inc(loopi);
  end;
  WPInitFontCacheSize := 50;
  GlobalLanguage := 0;
  DefaultFontSize := 10;
  BulletChar := #0;
  while TRUE do
  begin
    WPWordDelimiterArray[BulletChar] :=
      (BulletChar < 'A') and not (BulletChar in ['0'..'9']);
    if BulletChar = #255 then
      break
    else
      inc(BulletChar);
  end;
  { Fontsize translation }
  ib := 0;
  while TRUE do
  begin
    WPFontSizeTranslation[ib] := ib;
    if ib = 255 then break;
    inc(ib);
  end;

  BulletChar := #108;
  BulletFont := 'WingDings';
  WPTPrinter := Printer;
  DefaultReaderWriter.DefaultWriterANSI := 'TWPTextWriter';
  DefaultReaderWriter.DefaultWriterRTF := 'TWPRTFWriter';
  DefaultReaderWriter.DefaultWriterHTML := 'TWPHTMLWriter';
  DefaultReaderWriter.DefaultReaderANSI := 'TWPTextReader';
  DefaultReaderWriter.DefaultReaderRTF := 'TWPRTFReader';
  DefaultReaderWriter.DefaultReaderHTML := 'TWPHTMLReader';
  WPPrinterCanvas := TCanvas.Create;

  loopb := Integer('0');
  loopi := 0;
  while loopb <= Integer('9') do
  begin
    FWPT_IsHexChar[loopb] := TRUE;
    FWPT_HexCharValue[loopb] := loopi;
    FWPT_HexCharValueHI[loopb] := loopi * 16;
    inc(loopi);
    inc(loopb);
  end;
  loopb := Integer('A');
  loopi := 10;
  while loopb <= Integer('F') do
  begin
    FWPT_IsHexChar[loopb] := TRUE;
    FWPT_HexCharValue[loopb] := loopi;
    FWPT_HexCharValueHI[loopb] := loopi * 16;
    inc(loopi);
    inc(loopb);
  end;
  loopb := Integer('a');
  loopi := 10;
  while loopb <= Integer('f') do
  begin
    FWPT_IsHexChar[loopb] := TRUE;
    FWPT_HexCharValue[loopb] := loopi;
    FWPT_HexCharValueHI[loopb] := loopi * 16;
    inc(loopi);
    inc(loopb);
  end;



finalization
  WPLocalSaveVCLStrings := prevWPLocalSaveVCLStrings;
  WPLocalLoadVCLStrings := prevWPLocalLoadVCLStrings;
  DefaultReaderWriter.Free;
  if FWPPrinterCanvasHasHandle then
  begin
    WPPrinterCanvas.Handle := 0;
    DeleteDC(WPPrinterHDC);
    FWPPrinterCanvasHasHandle := FALSE;
  end;
  WPPrinterCanvas.Free;
  WPPrinterName := '';
  elo elo 3 2 0

end.



