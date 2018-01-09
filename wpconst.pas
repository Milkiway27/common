{ Stringconst -	WPDEFS.RES
  Include file,	no pascal source! }


{ ReportBuilder	uses numbers  43000 - 48599
  Ace uses 56400 - 56499
  WPTools uses 52000 - 53800 and 56600 - 57800
  WP Form&Report uses 54500 - 54600 }

const
    meWPLabels = 52000;	{ basis	of ALL TWPResLabel }

    meFilter	  = 56501;
    meUserCancel  = 56502;
    meReady	  = 56503;
    meReading	  = 56504;
    meWriting	  = 56505;
    meFormatA	  = 56506;
    meFormatB	  = 56507;
    meClearMemory = 56508;
    meNoSelection = 56509;
    meRichText	  = 56510;
    mePlainText	  = 56511;
    meSysError	  = 56512;
    meClearTabs	  = 56513;
    { new in V2	}
    meObjGraphicFilter	   = 56518;
    meObjGraphicFilterExt  = 56519;
    meObjTextFilter	   = 56520;
    meObjNotSupported	   = 56521;
    meObjGraphicNotLinked  = 56522;
    meObjDelete		   = 56523;
    meObjInsert		   = 56524;
    meSaveChangedText	   = 56525;
    meClearChangedText	   = 56526;
    meRecursiveToolbarUsage= 56527;
    meWrongScreenmode	   = 56528;
    meTextNotFound	   = 56529;
    meUndefinedInsertpoint = 56530;
    meNotSupportedProperty = 56531;
    meUseOnlyBitmaps	   = 56532;
    meWrongUsageOfFastAppend = 56533;
    mePaletteCompletelyUsed = 56534;
    meWriterClassNotFound   = 56535;
    meReaderClassNotFound   = 56536;
    meIOClassError	    = 56537;
    meNotSupportedError	    = 56538;

    meutNone = 56540;
    meutAny = 56541;
    meutInput =	56542;
    meutDeleteText = 56543;
    meutChangeAttributes = 56544;
    meutChangeIndent = 56545;
    meutChangeSpacing =	56546;
    meutChangeAlignment	= 56547;
    meutChangeTabs = 56548;
    meutChangeBorder = 56549;
    meutDeleteSelection	= 56550;
    meutDragAndDrop = 56551;
    meutPaste =	56552;

    { for the propertiy	dialogs	}
    meDiaAlLeft		   = 56560;
    meDiaAlRight	   = 56561;
    meDiaAlCenter	   = 56562;
    meDiaAlJustified	   = 56563;
    meDiaYes		   = 56564;
    meDiaNo		   = 56565;

    { 30 strings for WPReporter	Addon }
    wpmeRep_BandError	     = 56800;
    wpmeRep_NoBandInGrp	     = 56801;
    wpmeRep_TooManyClosings  = 56802;
    wpmeRep_WrongFormat	     = 56803;



    glDiff    =	0;
    glEnglish =	glDiff+0;
    glGerman  =	glDiff+200;
    glFrench  =	glDiff+400;
    glSpanish =	glDiff+600;
    glDutch   =	glDiff+800;
    glItalian =	glDiff+1000;
    glBrPortuguese = glDiff+1200; { Brazilian Portuguese }
    glDanish =	glDiff+1400;
    glCzech =	glDiff+1600;

