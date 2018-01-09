unit WP1Style;

{$I WPINC.INC}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, WPUtil, Dialogs,
  StdCtrls, Menus, WPBltDlg, WPTabdlg, WpParBrd, WpParPrp, WPRtfTxt, WPRich, WPDefs,
  WPPrint, WpWinCtr, WPStyCol, ExtCtrls;

type
  {$IFNDEF T2H}
  TWPOneStyleDlg = class;

  TWPOneStyleDefinition = class(TWPShadedForm)
    WPParagraphPropDlg1: TWPParagraphPropDlg;
    WPParagraphBorderDlg1: TWPParagraphBorderDlg;
    WPTabDlg1: TWPTabDlg;
    WPBulletDlg1: TWPBulletDlg;
    PopupMenu1: TPopupMenu;
    ParagraphProperties1: TMenuItem;
    Borders1: TMenuItem;
    Tabstops1: TMenuItem;
    Numbers1: TMenuItem;
    ModifyEntry: TButton;
    LevelSel: TComboBox;
    labNumberLevel: TLabel;
    labBasedOn: TLabel;
    BasedOnSel: TComboBox;
    NextSel: TComboBox;
    labNextStyle: TLabel;
    labName: TLabel;
    NameEdit: TEdit;
    TemplateWP: TWPRichText;
    WPStyleCollection1: TWPStyleCollection;
    Bevel1: TBevel;
    Bevel2: TBevel;
    labFont: TLabel;
    Bevel3: TBevel;
    BoldCheck: TCheckBox;
    UnderCheck: TCheckBox;
    labSize: TLabel;
    ItalicCheck: TCheckBox;
    FontCombo: TComboBox;
    SizeCombo: TComboBox;
    FontColor: TComboBox;
    BKColor: TComboBox;
    labColors: TLabel;
    btnOk: TButton;
    btnCancel: TButton;
    labBold: TLabel;
    labItalic: TLabel;
    labUnderline: TLabel;
    Memo1: TMemo;
    IsOutline: TCheckBox;
    labIsOutline: TLabel;
    labKeepTextToether: TLabel;
    labKeepWithNext: TLabel;
    labIsCharacterStyle: TLabel;
    NONE_str: TLabel;
    procedure ModifyEntryClick(Sender: TObject);
    procedure Numbers1Click(Sender: TObject);
    procedure SizeComboKeyPress(Sender: TObject; var Key: Char);
    procedure FormCreate(Sender: TObject);
    procedure FontComboChange(Sender: TObject);
    procedure SizeComboChange(Sender: TObject);
    procedure BoldCheckClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FontColorClick(Sender: TObject);
    procedure FontColorDrawItem(Control: TWinControl; Index: Integer;
      Rect: TRect; State: TOwnerDrawState);
    procedure NextSelChange(Sender: TObject);
    procedure BasedOnSelChange(Sender: TObject);
    procedure NameEditChange(Sender: TObject);
    procedure LevelSelChange(Sender: TObject);
    procedure IsOutlineClick(Sender: TObject);
    procedure NextSelDropDown(Sender: TObject);
    procedure BasedOnSelDropDown(Sender: TObject);
  protected
    Controller: TWPOneStyleDlg;
  public
    procedure ReadStyleInfo;
    procedure WriteStyleInfo;
  end;
  {$ENDIF}

  TWPOneStyleDlg = class(TWPCustomAttrDlg)
  private
    dia: TWPOneStyleDefinition;
    FStyleCollection: TWPStyleCollection;
    FStyleName: string;
    FStyleBasedOn: string;
    FStyleNext: string;
    FStyleValues: TWPParStyleValues;
    FCollection: TWPStyleCollection; // can be nil !
    FProcessUpdate: Boolean;
    procedure SetStyleCollection(x: TWPStyleCollection);
    procedure SetStyleValues(x: TWPParStyleValues);
  protected
    procedure SetEditBox(x: TWPCustomRichText); override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    FWindowLeft, FWindowTop: Integer;
    FNumberStyles   : TWPRTFNumberingStyleCollection; 
    function Execute: Boolean; override;
    constructor Create(aOwner: TComponent); override;
    destructor Destroy; override;
  published
    property EditBox;
    property ProcessUpdate: Boolean read FProcessUpdate write FProcessUpdate;
    property OnShowDialog;
    property StyleCollection: TWPStyleCollection read FStyleCollection write SetStyleCollection;
    property StyleName: string read FStyleName write FStyleName;
    property StyleBasedOn: string read FStyleBasedOn write FStyleBasedOn;
    property StyleNext: string read FStyleNext write FStyleNext;
    property StyleValues: TWPParStyleValues read FStyleValues write SetStyleValues;
  end;

implementation

{$R *.DFM}

{ -------------------------------------------------------------------------- }

constructor TWPOneStyleDlg.Create(aOwner: TComponent);
begin
  inherited Create(aOwner);
  FStyleValues := TWPParStyleValues.Create;
end;

destructor TWPOneStyleDlg.Destroy;
begin
  FStyleValues.Free;
  inherited Destroy;
end;

procedure TWPOneStyleDlg.SetStyleCollection(x: TWPStyleCollection);
begin
  FStyleCollection := x;
  if FStyleCollection <> nil then
  begin
    // FStyleCollection.FreeNotification(nil);
    FEditBox := nil;
  end;
end;

procedure TWPOneStyleDlg.SetStyleValues(x: TWPParStyleValues);
begin
  FStyleValues.Assign(x);
end;

procedure TWPOneStyleDlg.SetEditBox(x: TWPCustomRichText);
begin
  inherited SetEditBox(x);
  if EditBox <> nil then FStyleCollection := nil;
end;

procedure TWPOneStyleDlg.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  if (Operation = opRemove) and (AComponent = FStyleCollection) then FStyleCollection := nil;
  inherited Notification(AComponent, Operation);
end;

function TWPOneStyleDlg.Execute: Boolean;
var
  aStyle, aStyle2: TWPParStyle;
  FCurrentStyle: string;
begin
  Result := FALSE;
  if FStyleCollection <> nil then
    FCollection := FStyleCollection
  else if FEditBox <> nil then
    FCollection := FEditBox.WPStyleCollection as TWPStyleCollection
  else
    FCollection := nil;

  try
    dia := TWPOneStyleDefinition.Create(Self);
    dia.Controller := Self;

    if FNumberStyles<>nil then
       dia.TemplateWP.NumberStyles.Assign(FNumberStyles);

    if (FWindowLeft <> 0) and (FWindowTop <> 0) then
    begin
      dia.Left := FWindowLeft;
      dia.Top := FWindowTop;
    end
    else
      dia.Position := poScreenCenter;
    if FStyleName <> '' then
      FCurrentStyle := FStyleName
    else if (FEditBox <> nil) then
    begin
      FCurrentStyle := FEditBox.CurrAttr.StyleName;
    end
    else
      FCurrentStyle := FStyleName;
    if FCollection <> nil then
    begin
      // Dia.NextSel.Style := csDropDownList;
      // Dia.BasedOnSel.Style := csDropDownList;
    end
    else
    begin
      Dia.NextSel.Items.Add(FCurrentStyle);
    end;
    // Prepare the single style in the dialogs StyleCollection
    aStyle := dia.WPStyleCollection1.Styles[0];
    if FCollection <> nil then
      aStyle2 := FCollection.Find(FCurrentStyle)
    else
      aStyle2 := nil;

    if aStyle2 = nil then
    begin
      aStyle.Name := FCurrentStyle;
      aStyle.BasedOnStyle := FStyleBasedOn;
      aStyle.NextStyle := FStyleNext;
      aStyle.Values := FStyleValues;
    end
    else
      aStyle.Assign(aStyle2);

    // Continue to prepare the dialog
    dia.BasedOnSel.Text := aStyle.BasedOnStyle;
    dia.NextSel.Text := aStyle.NextStyle;
    dia.NameEdit.Text := FCurrentStyle;

    // Update the preview TWPRichText
    dia.WriteStyleInfo;

    // Show the dialog
    if not FCreateAndFreeDialog
       and MayOpenDialog(dia)
       then
    begin
      if dia.ShowModal = mrOK then
      begin
        Result := TRUE;
        FCurrentStyle := aStyle.Name;
        if FStyleName <> '' then
        begin
          FStyleName := FCurrentStyle;
          FStyleBasedOn := aStyle.BasedOnStyle;
          FStyleNext := aStyle.NextStyle;
          FStyleValues.Assign(aStyle.Values);
        end;

        if aStyle2 <> nil then
        begin
          aStyle2.Assign(aStyle);
        end
        else if FCollection <> nil then
        begin
          aStyle2 := FCollection.Add(FCurrentStyle);
          aStyle2.Assign(aStyle);
        end;

        aStyle2.Modified := TRUE;

        if (FStyleName = '') and (FEditBox <> nil) then
        begin
          FEditBox.CurrAttr.StyleName := FCurrentStyle;
        end;
        if FProcessUpdate and (FCollection <> nil) then
          FCollection.UpdateAll(nil, true);
      end;
    end;
  finally
    dia.Free;
  end;
end;

{ -------------------------------------------------------------------------- }

procedure TWPOneStyleDefinition.ModifyEntryClick(Sender: TObject);
var
  p: TPoint;
begin
  p := ModifyEntry.ClientToScreen(Point(0, ModifyEntry.Height));
  PopupMenu1.Popup(p.x, p.y);
end;

procedure TWPOneStyleDefinition.Numbers1Click(Sender: TObject);
var
  b: Boolean;
begin
  b := FALSE;
  // TemplateWP.FirstPar^.Locked := WPStyleCollection1.Styles[0].ParElements;
  case (Sender as TMenuItem).Tag of
    1: b := WPParagraphPropDlg1.Execute;
    3: b := WPParagraphBorderDlg1.Execute;
    4: b := WPTabDlg1.Execute;
    5: begin
         b := WPBulletDlg1.Execute;
         if b then
         begin
            if TemplateWP.FirstPar^.numlevel=0 then
            begin
               if LevelSel.ItemIndex>=0 then LevelSel.ItemIndex := -1;
               exclude(TemplateWP.FirstPar^.locked,wpseNumberLevel);
            end else
               LevelSel.ItemIndex := TemplateWP.FirstPar^.numlevel;
               
            if TemplateWP.FirstPar^.number=0 then
               exclude(TemplateWP.FirstPar^.locked,wpseNumber);
            WPStyleCollection1.Styles[0].Values.RemoveValue(
               FWPStyleParElementName[wpseNumberLevel] );
            WPStyleCollection1.Styles[0].Values.RemoveValue(
               FWPStyleParElementName[wpseNumber] );
         end;
    end;
  end;
  if b then ReadStyleInfo;
end;

procedure TWPOneStyleDefinition.SizeComboKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not (Key in ['0'..'9']) and not (Key = #8) then Key := #0;
end;

procedure TWPOneStyleDefinition.FormCreate(Sender: TObject);
begin
  FontCombo.Items.Assign(Screen.Fonts);
  WPBulletDlg1.DontEditOutline := TRUE;
end;

procedure TWPOneStyleDefinition.FontComboChange(Sender: TObject);
begin
  TemplateWP.CPPosition := 0;
  if FontCombo.Text = '' then
  begin
    TemplateWP.CurrAttr.FontName := 'Arial';
    TemplateWP.CurrAttr.DeleteStyle([afsParStyleFont]);
  end
  else
  begin
    TemplateWP.CurrAttr.FontName := FontCombo.Text;
    TemplateWP.CurrAttr.AddStyle([afsParStyleFont]);
  end;
  ReadStyleInfo;
end;

procedure TWPOneStyleDefinition.LevelSelChange(Sender: TObject);
begin
  TemplateWP.CPPosition := 0;
  if LevelSel.ItemIndex <= 0 then
    TemplateWP.CurrAttr.ExNumberStyle := nil
  else
    TemplateWP.CurrAttr.ExNumberStyle :=
       TemplateWP.NumberStyles.AddOutlineStyle(1,LevelSel.ItemIndex);
  ReadStyleInfo;
end;

procedure TWPOneStyleDefinition.IsOutlineClick(Sender: TObject);
begin
  TemplateWP.CurrAttr.OutlineMode := IsOutline.Checked;
  ReadStyleInfo;
end;

procedure TWPOneStyleDefinition.SizeComboChange(Sender: TObject);
begin
  TemplateWP.CPPosition := 0;
  if SizeCombo.Text = '' then
  begin
    TemplateWP.CurrAttr.Size := 11;
    TemplateWP.CurrAttr.DeleteStyle([afsParStyleSize]);
  end
  else
    TemplateWP.CurrAttr.Size := StrToIntDef(SizeCombo.Text, 1);
  ReadStyleInfo;
end;

procedure TWPOneStyleDefinition.BoldCheckClick(Sender: TObject);
const
  AttrLock: array[0..2] of TOneWrtStyle = (afsParStyleBold, afsParStyleItalic, afsParStyleUnderline);
  AttrCheck: array[0..2] of TOneWrtStyle = (afsBold, afsItalic, afsUnderline);
begin
  case (Sender as TCheckBox).State of
    cbGrayed:
      begin
        TemplateWP.CurrAttr.DeleteStyle([AttrCheck[(Sender as TCheckBox).Tag]]);
        TemplateWP.CurrAttr.DeleteStyle([AttrLock[(Sender as TCheckBox).Tag]]);
      end;
    cbChecked:
      begin
        TemplateWP.CurrAttr.AddStyle([AttrCheck[(Sender as TCheckBox).Tag]]);
        TemplateWP.CurrAttr.AddStyle([AttrLock[(Sender as TCheckBox).Tag]]);
      end;
    cbUnchecked:
      begin
        TemplateWP.CurrAttr.DeleteStyle([AttrCheck[(Sender as TCheckBox).Tag]]);
        TemplateWP.CurrAttr.AddStyle([AttrLock[(Sender as TCheckBox).Tag]]);
      end;
  end;
  ReadStyleInfo;
end;

procedure TWPOneStyleDefinition.FormShow(Sender: TObject);
var
  pa: PTAttr;
const
  state: array[Boolean] of TCheckBoxState = (cbUnchecked, cbChecked);
begin
    LevelSel.Items[0] := NONE_str.Caption;
  TemplateWP.SelectAll;
  WPStyleCollection1.AssignStyle(TemplateWP, NameEdit.Text, true);
  pa := TemplateWP.FirstPar^.line^.pa;
  if afsParStyleFont in pa^.style then
    FontCombo.ItemIndex := FontCombo.Items.IndexOf(TemplateWP.GetFontName(pa^.Font));
  if afsParStyleSize in pa^.style then
    SizeCombo.Text := IntToStr(pa^.Size);
  if afsParStyleBold in pa^.style then
    BoldCheck.State := state[afsBold in pa^.Style];
  if afsParStyleItalic in pa^.style then
    ItalicCheck.State := state[afsItalic in pa^.Style];
  if afsParStyleUnderline in pa^.style then
    UnderCheck.State := state[afsUnderline in pa^.Style];
  if afsParStyleColor in pa^.style then
    FontColor.ItemIndex := FontColor.Items.Add(ColorToString(TemplateWP.CurrAttr.NrToColor(pa^.Color)));
  if afsParStyleBGColor in pa^.style then
    BKColor.ItemIndex := BKColor.Items.Add(ColorToString(TemplateWP.CurrAttr.NrToColor(pa^.BGColor)));
  IsOutline.Checked := paprIsOutline in TemplateWP.FirstPar^.prop;
  if TemplateWP.FirstPar^.numlevel = 0 then
    LevelSel.ItemIndex := -1
  else
    LevelSel.ItemIndex := TemplateWP.FirstPar^.numlevel mod 9;
end;

procedure TWPOneStyleDefinition.FontColorClick(Sender: TObject);
var
  col: TColor; s: string;
begin
  s := (Sender as TComboBox).Text;
  if s = '' then
  begin
    if Sender = FontColor then
    begin
      TemplateWP.CurrAttr.Color := TemplateWP.CurrAttr.ColorToNr(clWindow, true);
      TemplateWP.CurrAttr.DeleteStyle([afsParStyleColor]);
    end
    else
    begin
      TemplateWP.CurrAttr.BKColor := TemplateWP.CurrAttr.ColorToNr(clWindow, true);
      TemplateWP.CurrAttr.DeleteStyle([afsParStyleBGColor]);
    end;
  end
  else
  begin
    col := StringToColor(s);
    if Sender = FontColor then
      TemplateWP.CurrAttr.Color := TemplateWP.CurrAttr.ColorToNr(col, true)
    else
      TemplateWP.CurrAttr.BKColor := TemplateWP.CurrAttr.ColorToNr(col, true);
  end;
  ReadStyleInfo;
end;

procedure TWPOneStyleDefinition.FontColorDrawItem(Control: TWinControl;
  Index: Integer; Rect: TRect; State: TOwnerDrawState);
var
  col: TColor; i: Integer;
begin
  try
    col := StringToColor((Control as TComboBox).Items.Strings[Index]);
  except
    col := clWindow;
  end;
  if Index = 0 then
  begin
  end
  else
    TComboBox(Control).Canvas.Brush.Color := col;
  TComboBox(Control).Canvas.Pen.Color := clBlack;
  TComboBox(Control).Canvas.Pen.Style := psSolid;
  TComboBox(Control).Canvas.Pen.Width := 1;
  TComboBox(Control).Canvas.FillRect(Rect);
  TComboBox(Control).Canvas.Rectangle(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom);
  if Control = BKColor then
  begin
    i := (Rect.Bottom - Rect.Top) div 3;
    TComboBox(Control).Canvas.MoveTo(Rect.Left, Rect.Top + i - 1);
    TComboBox(Control).Canvas.LineTo(Rect.Right, Rect.Top + i - 1);
    TComboBox(Control).Canvas.MoveTo(Rect.Left, Rect.Top + i * 2);
    TComboBox(Control).Canvas.LineTo(Rect.Right, Rect.Top + i * 2);
  end;
end;

procedure TWPOneStyleDefinition.ReadStyleInfo;
var AttrOnly: TWPStyleElementsChar;  pa : PTAttr;
begin
  pa := TemplateWP.FirstPar^.line^.pa;
  AttrOnly := [];
  if afsParStyleBold in pa^.style then include(AttrOnly, wpseFontBold);
  if afsParStyleItalic in pa^.style then include(AttrOnly, wpseFontItalic);
  if afsParStyleUnderline in pa^.style then include(AttrOnly, wpseFontUnderline);
  if afsParStyleFont in pa^.style then include(AttrOnly, wpseFontFont);
  if afsParStyleSize in pa^.style then include(AttrOnly, wpseFontSize);
  if afsParStyleColor in pa^.style then include(AttrOnly, wpseFontColor);
  if afsParStyleBGColor in pa^.style then include(AttrOnly, wpseFontBGColor);
  WPStyleCollection1.Styles[0].ReadFromPar(TemplateWP.FirstPar, pa,
    TemplateWP.Memo, TRUE, TemplateWP.FirstPar.locked, AttrOnly);
  Memo1.text := WPStyleCollection1.Styles[0].Values.CommaText;
end;

procedure TWPOneStyleDefinition.WriteStyleInfo;
var
  i: Integer;
begin
  TemplateWP.BeginUpdate;
  try
    i := TemplateWP.GetStyleNrForName(WPStyleCollection1.Styles[0].Name);
    TemplateWP.FirstPar^.Style := i;
    TemplateWP.CurrAttr.ClearAttributes(0, 12, -1);
    WPStyleCollection1.UpdateAll(nil, false);
  finally
    TemplateWP.EndUpdate;
    Memo1.text := WPStyleCollection1.Styles[0].Values.CommaText;
  end;
end;

procedure TWPOneStyleDefinition.NextSelChange(Sender: TObject);
begin
  WPStyleCollection1.Styles[0].NextStyle := NextSel.Text;
end;

procedure TWPOneStyleDefinition.BasedOnSelChange(Sender: TObject);
begin
  WPStyleCollection1.Styles[0].BasedOnStyle := BasedOnSel.Text;
  WriteStyleInfo;
end;

procedure TWPOneStyleDefinition.NameEditChange(Sender: TObject);
begin
  if NextSel.Text = WPStyleCollection1.Styles[0].Name then
    NextSel.Text := NameEdit.Text;
  WPStyleCollection1.Styles[0].Name := NameEdit.Text;
end;


procedure TWPOneStyleDefinition.NextSelDropDown(Sender: TObject);
begin
  if Controller.FCollection <> nil then
  begin
    Controller.FCollection.GetStyleList(NextSel.Items);
    if NextSel.Items.IndexOf(NameEdit.Text) < 0 then
      NextSel.Items.Insert(0, NameEdit.Text);
  end;
end;

procedure TWPOneStyleDefinition.BasedOnSelDropDown(Sender: TObject);
var
  i: Integer;
begin
  if Controller.FCollection <> nil then
  begin
    Controller.FCollection.GetStyleList(BasedOnSel.Items);
    i := BasedOnSel.Items.IndexOf(NameEdit.Text);
    if i >= 0 then BasedOnSel.Items.Delete(i);
  end;
end;

end.

