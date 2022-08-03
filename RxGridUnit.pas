unit RxGridUnit;

interface

uses Windows, Classes, Math, Grids, dbGrids, DB, Graphics, Controls, Messages, Types, Forms;

const
  DefRxGridOptions = [ { dgEditing, } dgTitles, dgIndicator, dgColumnResize,
    dgColLines, dgRowLines, dgConfirmDelete, dgCancelOnExit];

type
  TVertAlignment = (vaTopJustify, vaCenter, vaBottomJustify);

{$DEFINE RX_D4}
{$DEFINE RX_D5}
{$IFDEF RX_V110}
{$IFDEF CBUILDER}
{$NODEFINE DefRxGridOptions}
{$ENDIF}
{$ENDIF}
type
  TTitleClickEvent = procedure(Sender: TObject; ACol: Longint; Field: TField)
    of object;
  TCheckTitleBtnEvent = procedure(Sender: TObject; ACol: Longint; Field: TField;
    var Enabled: Boolean) of object;
  TGetCellParamsEvent = procedure(Sender: TObject; Field: TField; AFont: TFont;
    var Background: TColor; Highlight: Boolean) of object;
  TSortMarker = (smNone, smDown, smUp);
  TGetBtnParamsEvent = procedure(Sender: TObject; Field: TField; AFont: TFont;
    var Background: TColor; var SortMarker: TSortMarker; IsDown: Boolean)
    of object;
  TGetCellPropsEvent = procedure(Sender: TObject; Field: TField; AFont: TFont;
    var Background: TColor) of object; { obsolete }
  TDBEditShowEvent = procedure(Sender: TObject; Field: TField;
    var AllowEdit: Boolean) of object;

  TRxDBGrid = class(TDBGrid)
  private
    { FShowGlyphs: Boolean; }
    FDefaultDrawing: Boolean;
    FTitleButtons: Boolean;
    FPressedCol: TColumn;
    FPressed: Boolean;
    FTracking: Boolean;
    FSwapButtons: Boolean;
    FDisableCount: Integer;
    FFixedCols: Integer;
    FMsIndicators: TImageList;
    FOnCheckButton: TCheckTitleBtnEvent;
    FOnGetCellProps: TGetCellPropsEvent;
    FOnGetCellParams: TGetCellParamsEvent;
    FOnGetBtnParams: TGetBtnParamsEvent;
    FOnEditChange: TNotifyEvent;
    FOnKeyPress: TKeyPressEvent;
    FOnTitleBtnClick: TTitleClickEvent;
    FOnTopLeftChanged: TNotifyEvent;
    { function GetImageIndex(Field: TField): Integer;
      procedure SetShowGlyphs(Value: Boolean); }
    procedure SetRowsHeight(Value: Integer);
    function GetRowsHeight: Integer;
    procedure SetTitleButtons(Value: Boolean);
    procedure StopTracking;
    procedure TrackButton(X, Y: Integer);
    function GetMasterColumn(ACol, ARow: Longint): TColumn;
    function GetTitleOffset: Byte;
    procedure SetFixedCols(Value: Integer);
    function GetFixedCols: Integer;
    function CalcLeftColumn: Integer;
    procedure WMChar(var Msg: TWMChar); message WM_CHAR;
    procedure WMCancelMode(var Message: TMessage); message WM_CANCELMODE;
    procedure WMRButtonUp(var Message: TWMMouse); message WM_RBUTTONUP;
  protected
    function AcquireFocus: Boolean;
    procedure DoTitleClick(ACol: Longint; AField: TField); dynamic;
    procedure CheckTitleButton(ACol, ARow: Longint;
      var Enabled: Boolean); dynamic;
    procedure DrawCell(ACol, ARow: Longint; ARect: TRect;
      AState: TGridDrawState); override;
    procedure DrawDataCell(const Rect: TRect; Field: TField;
      State: TGridDrawState); override; { obsolete from Delphi 2.0 }
    procedure EditChanged(Sender: TObject); dynamic;
    procedure GetCellProps(Field: TField; AFont: TFont; var Background: TColor;
      Highlight: Boolean); virtual;
    procedure KeyPress(var Key: Char); override;
    procedure SetColumnAttributes; override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState;
      X, Y: Integer); override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState;
      X, Y: Integer); override;
    function DoMouseWheelDown(Shift: TShiftState; MousePos: TPoint)
      : Boolean; override;
    function DoMouseWheelUp(Shift: TShiftState; MousePos: TPoint)
      : Boolean; override;
    procedure LayoutChanged; override;
    procedure TopLeftChanged; override;
    procedure DrawColumnCell(const Rect: TRect; DataCol: Integer;
      Column: TColumn; State: TGridDrawState); override;
    procedure ColWidthsChanged; override;
    procedure Paint; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure DefaultDataCellDraw(const Rect: TRect; Field: TField;
      State: TGridDrawState);
    procedure DisableScroll;
    procedure EnableScroll;
    function ScrollDisabled: Boolean;
    procedure MouseToCell(X, Y: Integer; var ACol, ARow: Longint);
    property Canvas;
    property Col;
    property InplaceEditor;
    property LeftCol;
    property Row;
    property VisibleRowCount;
    property VisibleColCount;
    property IndicatorOffset;
    property TitleOffset: Byte read GetTitleOffset;
  published
    property FixedCols: Integer read GetFixedCols write SetFixedCols default 0;
    property DefaultDrawing: Boolean read FDefaultDrawing write FDefaultDrawing
      default True;
    { property ShowGlyphs: Boolean read FShowGlyphs write SetShowGlyphs
      default True; }
    property TitleButtons: Boolean read FTitleButtons write SetTitleButtons
      default False;
    property RowsHeight: Integer read GetRowsHeight write SetRowsHeight
      stored False; { obsolete, for backward compatibility only }
    property OnCheckButton: TCheckTitleBtnEvent read FOnCheckButton
      write FOnCheckButton;
    property OnGetCellProps: TGetCellPropsEvent read FOnGetCellProps
      write FOnGetCellProps; { obsolete }
    property OnGetCellParams: TGetCellParamsEvent read FOnGetCellParams
      write FOnGetCellParams;
    property OnGetBtnParams: TGetBtnParamsEvent read FOnGetBtnParams
      write FOnGetBtnParams;
    property OnEditChange: TNotifyEvent read FOnEditChange write FOnEditChange;
    property OnTitleBtnClick: TTitleClickEvent read FOnTitleBtnClick
      write FOnTitleBtnClick;
    property OnKeyPress: TKeyPressEvent read FOnKeyPress write FOnKeyPress;
    property OnTopLeftChanged: TNotifyEvent read FOnTopLeftChanged
      write FOnTopLeftChanged;
    property OnContextPopup;
    property OnMouseDown;
    property OnMouseMove;
    property OnMouseUp;
    property OnMouseWheelDown;
    property OnMouseWheelUp;
  end;

type
  THack = class(TCustomControl);


  function WidthOf(R: TRect): Integer;
  function PaletteColor(Color: TColor): Longint;
  procedure StretchBltTransparent(DstDC: HDC; DstX, DstY, DstW, DstH: Integer;
    SrcDC: HDC; SrcX, SrcY, SrcW, SrcH: Integer; Palette: HPalette;
    TransparentColor: TColorRef);
  procedure StretchBitmapTransparent(Dest: TCanvas; Bitmap: TBitmap;
    TransparentColor: TColor; DstX, DstY, DstW, DstH, SrcX, SrcY, SrcW,
    SrcH: Integer);
  procedure DrawBitmapTransparent(Dest: TCanvas; DstX, DstY: Integer;
    Bitmap: TBitmap; TransparentColor: TColor);
  procedure UsesBitmap;
  procedure WriteText(ACanvas: TCanvas; ARect: TRect; DX, DY: Integer;
    const Text: string; Alignment: TAlignment; WordWrap: Boolean
{$IFDEF RX_D4}; ARightToLeft: Boolean = False {$ENDIF});
  function MinimizeText(const Text: string; Canvas: TCanvas;
    MaxWidth: Integer): string;
  procedure DrawCellTextEx(Control: TCustomControl; ACol, ARow: Longint;
    const S: string; const ARect: TRect; Align: TAlignment;
    VertAlign: TVertAlignment; WordWrap: Boolean; ARightToLeft: Boolean);
  procedure DrawCellText(Control: TCustomControl; ACol, ARow: Longint;
    const S: string; const ARect: TRect; Align: TAlignment;
    VertAlign: TVertAlignment; ARightToLeft: Boolean);
  procedure GridInvalidateRow(Grid: TRxDBGrid; Row: Longint);



implementation
uses SysUtils;
{ TRxDBGrid }

function WidthOf(R: TRect): Integer;
begin
  Result := R.Right - R.Left;
end;

var
  DrawBitmap: TBitmap = nil;

const
  { TBitmap.GetTransparentColor from GRAPHICS.PAS use this value }
  PaletteMask = $02000000;

function PaletteColor(Color: TColor): Longint;
begin
  Result := ColorToRGB(Color) or PaletteMask;
end;

procedure StretchBltTransparent(DstDC: HDC; DstX, DstY, DstW, DstH: Integer;
  SrcDC: HDC; SrcX, SrcY, SrcW, SrcH: Integer; Palette: HPalette;
  TransparentColor: TColorRef);
var
  Color: TColorRef;
  bmAndBack, bmAndObject, bmAndMem, bmSave: HBitmap;
  bmBackOld, bmObjectOld, bmMemOld, bmSaveOld: HBitmap;
  MemDC, BackDC, ObjectDC, SaveDC: HDC;
  palDst, palMem, palSave, palObj: HPalette;
begin
  { Create some DCs to hold temporary data }
  BackDC := CreateCompatibleDC(DstDC);
  ObjectDC := CreateCompatibleDC(DstDC);
  MemDC := CreateCompatibleDC(DstDC);
  SaveDC := CreateCompatibleDC(DstDC);
  { Create a bitmap for each DC }
  bmAndObject := CreateBitmap(SrcW, SrcH, 1, 1, nil);
  bmAndBack := CreateBitmap(SrcW, SrcH, 1, 1, nil);
  bmAndMem := CreateCompatibleBitmap(DstDC, DstW, DstH);
  bmSave := CreateCompatibleBitmap(DstDC, SrcW, SrcH);
  { Each DC must select a bitmap object to store pixel data }
  bmBackOld := SelectObject(BackDC, bmAndBack);
  bmObjectOld := SelectObject(ObjectDC, bmAndObject);
  bmMemOld := SelectObject(MemDC, bmAndMem);
  bmSaveOld := SelectObject(SaveDC, bmSave);
  { Select palette }
  palDst := 0;
  palMem := 0;
  palSave := 0;
  palObj := 0;
  if Palette <> 0 then
  begin
    palDst := SelectPalette(DstDC, Palette, True);
    RealizePalette(DstDC);
    palSave := SelectPalette(SaveDC, Palette, False);
    RealizePalette(SaveDC);
    palObj := SelectPalette(ObjectDC, Palette, False);
    RealizePalette(ObjectDC);
    palMem := SelectPalette(MemDC, Palette, True);
    RealizePalette(MemDC);
  end;
  { Set proper mapping mode }
  SetMapMode(SrcDC, GetMapMode(DstDC));
  SetMapMode(SaveDC, GetMapMode(DstDC));
  { Save the bitmap sent here }
  BitBlt(SaveDC, 0, 0, SrcW, SrcH, SrcDC, SrcX, SrcY, SRCCOPY);
  { Set the background color of the source DC to the color, }
  { contained in the parts of the bitmap that should be transparent }
  Color := SetBkColor(SaveDC, PaletteColor(TransparentColor));
  { Create the object mask for the bitmap by performing a BitBlt() }
  { from the source bitmap to a monochrome bitmap }
  BitBlt(ObjectDC, 0, 0, SrcW, SrcH, SaveDC, 0, 0, SRCCOPY);
  { Set the background color of the source DC back to the original }
  SetBkColor(SaveDC, Color);
  { Create the inverse of the object mask }
  BitBlt(BackDC, 0, 0, SrcW, SrcH, ObjectDC, 0, 0, NOTSRCCOPY);
  { Copy the background of the main DC to the destination }
  BitBlt(MemDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, SRCCOPY);
  { Mask out the places where the bitmap will be placed }
  StretchBlt(MemDC, 0, 0, DstW, DstH, ObjectDC, 0, 0, SrcW, SrcH, SRCAND);
  { Mask out the transparent colored pixels on the bitmap }
  BitBlt(SaveDC, 0, 0, SrcW, SrcH, BackDC, 0, 0, SRCAND);
  { XOR the bitmap with the background on the destination DC }
  StretchBlt(MemDC, 0, 0, DstW, DstH, SaveDC, 0, 0, SrcW, SrcH, SRCPAINT);
  { Copy the destination to the screen }
  BitBlt(DstDC, DstX, DstY, DstW, DstH, MemDC, 0, 0, SRCCOPY);
  { Restore palette }
  if Palette <> 0 then
  begin
    SelectPalette(MemDC, palMem, False);
    SelectPalette(ObjectDC, palObj, False);
    SelectPalette(SaveDC, palSave, False);
    SelectPalette(DstDC, palDst, True);
  end;
  { Delete the memory bitmaps }
  DeleteObject(SelectObject(BackDC, bmBackOld));
  DeleteObject(SelectObject(ObjectDC, bmObjectOld));
  DeleteObject(SelectObject(MemDC, bmMemOld));
  DeleteObject(SelectObject(SaveDC, bmSaveOld));
  { Delete the memory DCs }
  DeleteDC(MemDC);
  DeleteDC(BackDC);
  DeleteDC(ObjectDC);
  DeleteDC(SaveDC);
end;

procedure StretchBitmapTransparent(Dest: TCanvas; Bitmap: TBitmap;
  TransparentColor: TColor; DstX, DstY, DstW, DstH, SrcX, SrcY, SrcW,
  SrcH: Integer);
var
  CanvasChanging: TNotifyEvent;
begin
  if DstW <= 0 then
    DstW := Bitmap.Width;
  if DstH <= 0 then
    DstH := Bitmap.Height;
  if (SrcW <= 0) or (SrcH <= 0) then
  begin
    SrcX := 0;
    SrcY := 0;
    SrcW := Bitmap.Width;
    SrcH := Bitmap.Height;
  end;
  if not Bitmap.Monochrome then
    SetStretchBltMode(Dest.Handle, STRETCH_DELETESCANS);
  CanvasChanging := Bitmap.Canvas.OnChanging;
  Bitmap.Canvas.Lock;
  try
    Bitmap.Canvas.OnChanging := nil;
    if TransparentColor = clNone then
    begin
      StretchBlt(Dest.Handle, DstX, DstY, DstW, DstH, Bitmap.Canvas.Handle,
        SrcX, SrcY, SrcW, SrcH, Dest.CopyMode);
    end
    else
    begin
      if TransparentColor = clDefault then
        TransparentColor := Bitmap.Canvas.Pixels[0, Bitmap.Height - 1];
      if Bitmap.Monochrome then
        TransparentColor := clWhite
      else
        TransparentColor := ColorToRGB(TransparentColor);
      StretchBltTransparent(Dest.Handle, DstX, DstY, DstW, DstH,
        Bitmap.Canvas.Handle, SrcX, SrcY, SrcW, SrcH, Bitmap.Palette,
        TransparentColor);
    end;
  finally
    Bitmap.Canvas.OnChanging := CanvasChanging;
    Bitmap.Canvas.Unlock;
  end;
end;

procedure DrawBitmapTransparent(Dest: TCanvas; DstX, DstY: Integer;
  Bitmap: TBitmap; TransparentColor: TColor);
begin
  StretchBitmapTransparent(Dest, Bitmap, TransparentColor, DstX, DstY,
    Bitmap.Width, Bitmap.Height, 0, 0, Bitmap.Width, Bitmap.Height);
end;

procedure UsesBitmap;
begin
  if DrawBitmap = nil then
    DrawBitmap := TBitmap.Create;
end;

procedure WriteText(ACanvas: TCanvas; ARect: TRect; DX, DY: Integer;
  const Text: string; Alignment: TAlignment; WordWrap: Boolean;
  ARightToLeft: Boolean = False);
const
  AlignFlags: array [TAlignment] of Integer = (DT_LEFT or DT_EXPANDTABS or
    DT_NOPREFIX, DT_RIGHT or DT_EXPANDTABS or DT_NOPREFIX,
    DT_CENTER or DT_EXPANDTABS or DT_NOPREFIX);
  WrapFlags: array [Boolean] of Integer = (0, DT_WORDBREAK);
  RTL: array [Boolean] of Integer = (0, DT_RTLREADING);
var
  B, R: TRect;
  I, Left: Integer;
begin
  UsesBitmap;
  I := ColorToRGB(ACanvas.Brush.Color);
  if not WordWrap and (Integer(GetNearestColor(ACanvas.Handle, I)) = I) and
    (Pos(#13, Text) = 0) then
  begin { Use ExtTextOut for solid colors }
    { In BiDi, because we changed the window origin, the text that does not
      change alignment, actually gets its alignment changed. }
    if (ACanvas.CanvasOrientation = coRightToLeft) and (not ARightToLeft) then
      ChangeBiDiModeAlignment(Alignment);
    case Alignment of
      taLeftJustify:
        Left := ARect.Left + DX;
      taRightJustify:
        Left := ARect.Right - ACanvas.TextWidth(Text) - 3;
    else { taCenter }
      Left := ARect.Left + (ARect.Right - ARect.Left) shr 1 -
        (ACanvas.TextWidth(Text) shr 1);
    end;
    ACanvas.TextRect(ARect, Left, ARect.Top + DY, Text);
    ExtTextOut(ACanvas.Handle, Left, ARect.Top + DY, ETO_OPAQUE or ETO_CLIPPED,
      @ARect, PChar(Text), Length(Text), nil);
  end
  else
  begin { Use FillRect and DrawText for dithered colors }
    DrawBitmap.Canvas.Lock;
    try
      with DrawBitmap,
        ARect do { Use offscreen bitmap to eliminate flicker and }
      begin { brush origin tics in painting / scrolling. }
        Width := Max(Width, Right - Left);
        Height := Max(Height, Bottom - Top);
        R := Rect(DX, DY, Right - Left - {$IFDEF WIN32} 1 {$ELSE} 2
{$ENDIF}, Bottom - Top - 1);
        B := Rect(0, 0, Right - Left, Bottom - Top);
      end;
      with DrawBitmap.Canvas do
      begin
        Font := ACanvas.Font;
        Font.Color := ACanvas.Font.Color;
        Brush := ACanvas.Brush;
        Brush.Style := bsSolid;
        FillRect(B);
        SetBkMode(Handle, TRANSPARENT);
        if (ACanvas.CanvasOrientation = coRightToLeft) then
          ChangeBiDiModeAlignment(Alignment);
        DrawText(Handle, PChar(Text), Length(Text), R, AlignFlags[Alignment] or
          RTL[ARightToLeft] or WrapFlags[WordWrap]);
      end;
      ACanvas.CopyRect(ARect, DrawBitmap.Canvas, B);
    finally
      DrawBitmap.Canvas.Unlock;
    end;
  end;
end;

function MinimizeText(const Text: string; Canvas: TCanvas;
  MaxWidth: Integer): string;
var
  I: Integer;
begin
  Result := Text;
  I := 1;
  while (I <= Length(Text)) and (Canvas.TextWidth(Result) > MaxWidth) do
  begin
    Inc(I);
    Result := Copy(Text, 1, Max(0, Length(Text) - I)) + '...';
  end;
end;

procedure DrawCellTextEx(Control: TCustomControl; ACol, ARow: Longint;
  const S: string; const ARect: TRect; Align: TAlignment;
  VertAlign: TVertAlignment; WordWrap: Boolean; ARightToLeft: Boolean);
const
  MinOffs = 2;
var
  H: Integer;
begin
  case VertAlign of
    vaTopJustify:
      H := MinOffs;
    vaCenter:
      with THack(Control) do
        H := Max(1, (ARect.Bottom - ARect.Top - Canvas.TextHeight('W')) div 2);
  else { vaBottomJustify }
    begin
      with THack(Control) do
        H := Max(MinOffs, ARect.Bottom - ARect.Top - Canvas.TextHeight('W'));
    end;
  end;
  WriteText(THack(Control).Canvas, ARect, MinOffs, H, S, Align, WordWrap { ,
      ARightToLeft } );
end;

procedure DrawCellText(Control: TCustomControl; ACol, ARow: Longint;
  const S: string; const ARect: TRect; Align: TAlignment;
  VertAlign: TVertAlignment; ARightToLeft: Boolean);
begin
  DrawCellTextEx(Control, ACol, ARow, S, ARect, Align, VertAlign,
    Align = taCenter, ARightToLeft);
end;

type
  TGridPicture = (gpBlob, gpMemo, gpPicture, gpOle, gpObject, gpData,
    gpNotEmpty, gpMarkDown, gpMarkUp);

var
  GridBmpNames: array [TGridPicture] of PChar = (
    'DBG_BLOB',
    'DBG_MEMO',
    'DBG_PICT',
    'DBG_OLE',
    'DBG_OBJECT',
    'DBG_DATA',
    'DBG_NOTEMPTY',
    'DBG_SMDOWN',
    'DBG_SMUP'
  );
  GridBitmaps: array [TGridPicture] of TBitmap = (
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil
  );

const
  bmMultiDot = 'DBG_MSDOT';
  bmMultiArrow = 'DBG_MSARROW';

constructor TRxDBGrid.Create(AOwner: TComponent);
var
  Bmp: TBitmap;
begin
  inherited Create(AOwner);
  inherited DefaultDrawing := False;
  Options := DefRxGridOptions;
  Bmp := TBitmap.Create;
  try
    Bmp.Handle := LoadBitmap(hInstance, bmMultiDot);
    FMsIndicators := TImageList.CreateSize(Bmp.Width, Bmp.Height);
    FMsIndicators.AddMasked(Bmp, clWhite);
    Bmp.Handle := LoadBitmap(hInstance, bmMultiArrow);
    FMsIndicators.AddMasked(Bmp, clWhite);
  finally
    Bmp.Free;
  end;
  { FShowGlyphs := True; }
  FDefaultDrawing := True;
end;

destructor TRxDBGrid.Destroy;
begin
  FMsIndicators.Free;
  inherited Destroy;
end;
(*
  function TRxDBGrid.GetImageIndex(Field: TField): Integer;
  var
  AOnGetText: TFieldGetTextEvent;
  AOnSetText: TFieldSetTextEvent;
  begin
  Result := -1;
  if FShowGlyphs and Assigned(Field) then begin
  if (not ReadOnly) and Field.CanModify then begin
  { Allow editing of memo fields if OnSetText and OnGetText
  events are assigned }
  AOnGetText := Field.OnGetText;
  AOnSetText := Field.OnSetText;
  if Assigned(AOnSetText) and Assigned(AOnGetText) then Exit;
  end;
  case Field.DataType of
  ftBytes, ftVarBytes, ftBlob: Result := Ord(gpBlob);
  ftMemo: Result := Ord(gpMemo);
  ftGraphic: Result := Ord(gpPicture);
  ftTypedBinary: Result := Ord(gpBlob);
  ftFmtMemo: Result := Ord(gpMemo);
  ftParadoxOle, ftDBaseOle: Result := Ord(gpOle);
  ftCursor: Result := Ord(gpData);
  ftReference, ftDataSet: Result := Ord(gpData);
  ftOraClob: Result := Ord(gpMemo);
  ftOraBlob: Result := Ord(gpBlob);
  end;
  end;
  end;
*)

procedure TRxDBGrid.LayoutChanged;
var
  ACol: Longint;
begin
  ACol := Col;
  inherited LayoutChanged;
  if Datalink.Active and (FixedCols > 0) then
    Col := Min(Max(CalcLeftColumn, ACol), ColCount - 1);
end;

function TRxDBGrid.CalcLeftColumn: Integer;
begin
  Result := FixedCols + IndicatorOffset;
  while (Result < ColCount) and (ColWidths[Result] <= 0) do
    Inc(Result);
end;

procedure TRxDBGrid.ColWidthsChanged;
var
  ACol: Longint;
begin
  ACol := Col;
  inherited ColWidthsChanged;
  if Datalink.Active and (FixedCols > 0) then
    Col := Min(Max(CalcLeftColumn, ACol), ColCount - 1);
end;

function TRxDBGrid.GetTitleOffset: Byte;
{$IFDEF RX_D4}
var
  I, J: Integer;
{$ENDIF}
begin
  Result := 0;
  if dgTitles in Options then
  begin
    Result := 1;

{$IFDEF RX_D4}
    if (Datalink <> nil) and (Datalink.DataSet <> nil) and Datalink.DataSet.ObjectView
    then
    begin
      for I := 0 to Columns.Count - 1 do
      begin
        if Columns[I].Showing then
        begin
          J := Columns[I].Depth;
          if J >= Result then
            Result := J + 1;
        end;
      end;
    end;
{$ENDIF}
  end;
end;

procedure TRxDBGrid.SetColumnAttributes;
begin
  inherited SetColumnAttributes;
  SetFixedCols(FFixedCols);
end;

procedure TRxDBGrid.SetFixedCols(Value: Integer);
var
  FixCount, I: Integer;
begin
  FixCount := Max(Value, 0) + IndicatorOffset;
  if Datalink.Active and not(csLoading in ComponentState) and
    (ColCount > IndicatorOffset + 1) then
  begin
    FixCount := Min(FixCount, ColCount - 1);
    inherited FixedCols := FixCount;
    for I := 1 to Min(FixedCols, ColCount - 1) do
      TabStops[I] := False;
  end;
  FFixedCols := FixCount - IndicatorOffset;
end;

function TRxDBGrid.GetFixedCols: Integer;
begin
  if Datalink.Active then
    Result := inherited FixedCols - IndicatorOffset
  else
    Result := FFixedCols;
end;

{
  procedure TRxDBGrid.SetShowGlyphs(Value: Boolean);
  begin
  if FShowGlyphs <> Value then begin
  FShowGlyphs := Value;
  Invalidate;
  end;
  end;
}
procedure TRxDBGrid.SetRowsHeight(Value: Integer);
begin
  if not(csDesigning in ComponentState) and (DefaultRowHeight <> Value) then
  begin
    DefaultRowHeight := Value;
    if dgTitles in Options then
      RowHeights[0] := Value + 2;
    if HandleAllocated then
      Perform(WM_SIZE, SIZE_RESTORED, MakeLong(ClientWidth, ClientHeight));
  end;
end;

function TRxDBGrid.GetRowsHeight: Integer;
begin
  Result := DefaultRowHeight;
end;

procedure TRxDBGrid.Paint;
begin
  inherited Paint;
  if not(csDesigning in ComponentState) and (dgRowSelect in Options) and
    DefaultDrawing and Focused then
  begin
    Canvas.Font.Color := clWindowText;
    with Selection do
      DrawFocusRect(Canvas.Handle, BoxRect(Left, Top, Right, Bottom));
  end;
end;

procedure TRxDBGrid.SetTitleButtons(Value: Boolean);
begin
  if FTitleButtons <> Value then
  begin
    FTitleButtons := Value;
    Invalidate;
  end;
end;

function TRxDBGrid.AcquireFocus: Boolean;
begin
  Result := True;
  if FAcquireFocus and CanFocus and not(csDesigning in ComponentState) then
  begin
    SetFocus;
    Result := Focused or (InplaceEditor <> nil) and InplaceEditor.Focused;
  end;
end;

procedure TRxDBGrid.GetCellProps(Field: TField; AFont: TFont;
  var Background: TColor; Highlight: Boolean);
var
  AColor, ABack: TColor;
begin
  if Assigned(FOnGetCellParams) then
    FOnGetCellParams(Self, Field, AFont, Background, Highlight)
  else if Assigned(FOnGetCellProps) then
  begin
    if Highlight then
    begin
      AColor := AFont.Color;
      FOnGetCellProps(Self, Field, AFont, ABack);
      AFont.Color := AColor;
    end
    else
      FOnGetCellProps(Self, Field, AFont, Background);
  end;
end;

procedure TRxDBGrid.DoTitleClick(ACol: Longint; AField: TField);
begin
  if Assigned(FOnTitleBtnClick) then
    FOnTitleBtnClick(Self, ACol, AField);
end;

procedure TRxDBGrid.CheckTitleButton(ACol, ARow: Longint; var Enabled: Boolean);
var
  Field: TField;
begin
  if (ACol >= 0) and (ACol < Columns.Count) then
  begin
    if Assigned(FOnCheckButton) then
    begin
      Field := Columns[ACol].Field;
      if ColumnAtDepth(Columns[ACol], ARow) <> nil then
        Field := ColumnAtDepth(Columns[ACol], ARow).Field;
      FOnCheckButton(Self, ACol, Field, Enabled);
    end;
  end
  else
    Enabled := False;
end;

procedure TRxDBGrid.DisableScroll;
begin
  Inc(FDisableCount);
end;

type
  THackLink = class(TGridDataLink);

procedure TRxDBGrid.EnableScroll;
begin
  if FDisableCount <> 0 then
  begin
    Dec(FDisableCount);
    if FDisableCount = 0 then
      THackLink(Datalink).DataSetScrolled(0);
  end;
end;

function TRxDBGrid.ScrollDisabled: Boolean;
begin
  Result := FDisableCount <> 0;
end;

function TRxDBGrid.DoMouseWheelDown(Shift: TShiftState;
  MousePos: TPoint): Boolean;
begin
  Result := False;
  if Assigned(OnMouseWheelDown) then
    OnMouseWheelDown(Self, Shift, MousePos, Result);
  if not Result then
  begin
    if not AcquireFocus then
      Exit;
    if Datalink.Active then
    begin
      Result := Datalink.DataSet.MoveBy(1) <> 0;
    end;
  end;
end;

function TRxDBGrid.DoMouseWheelUp(Shift: TShiftState; MousePos: TPoint)
  : Boolean;
begin
  Result := False;
  if Assigned(OnMouseWheelUp) then
    OnMouseWheelUp(Self, Shift, MousePos, Result);
  if not Result then
  begin
    if not AcquireFocus then
      Exit;
    if Datalink.Active then
    begin
      Result := Datalink.DataSet.MoveBy(-1) <> 0;
    end;
  end;
end;

procedure TRxDBGrid.EditChanged(Sender: TObject);
begin
  if Assigned(FOnEditChange) then
    FOnEditChange(Self);
end;

procedure GridInvalidateRow(Grid: TRxDBGrid; Row: Longint);
var
  I: Longint;
begin
  for I := 0 to Grid.ColCount - 1 do
    Grid.InvalidateCell(I, Row);
end;

procedure TRxDBGrid.TopLeftChanged;
begin
  if (dgRowSelect in Options) and DefaultDrawing then
    GridInvalidateRow(Self, Self.Row);
  inherited TopLeftChanged;
  if FTracking then
    StopTracking;
  if Assigned(FOnTopLeftChanged) then
    FOnTopLeftChanged(Self);
end;

procedure TRxDBGrid.StopTracking;
begin
  if FTracking then
  begin
    TrackButton(-1, -1);
    FTracking := False;
    MouseCapture := False;
  end;
end;

procedure TRxDBGrid.TrackButton(X, Y: Integer);
var
  Cell: TGridCoord;
  NewPressed: Boolean;
  I, Offset: Integer;
begin
  Cell := MouseCoord(X, Y);
  Offset := TitleOffset;
  NewPressed := PtInRect(Rect(0, 0, ClientWidth, ClientHeight), Point(X, Y)) and
    (FPressedCol = GetMasterColumn(Cell.X, Cell.Y)) and (Cell.Y < Offset);
  if FPressed <> NewPressed then
  begin
    FPressed := NewPressed;
    for I := 0 to Offset - 1 do
      GridInvalidateRow(Self, I);
  end;
end;

procedure TRxDBGrid.MouseDown(Button: TMouseButton; Shift: TShiftState;
  X, Y: Integer);
var
  Cell: TGridCoord;
  MouseDownEvent: TMouseEvent;
  EnableClick: Boolean;
begin
  if not AcquireFocus then
    Exit;
  if (ssDouble in Shift) and (Button = mbLeft) then
  begin
    DblClick;
    Exit;
  end;
  if Sizing(X, Y) then
    inherited MouseDown(Button, Shift, X, Y)
  else
  begin
    Cell := MouseCoord(X, Y);
    if (DragKind = dkDock) and (Cell.X < IndicatorOffset) and
      (Cell.Y < TitleOffset) and (not(csDesigning in ComponentState)) then
    begin
      BeginDrag(False);
      Exit;
    end;
    if FTitleButtons and (Datalink <> nil) and Datalink.Active and
      (Cell.Y < TitleOffset) and (Cell.X >= IndicatorOffset) and
      not(csDesigning in ComponentState) then
    begin
      if (dgColumnResize in Options) and (Button = mbRight) then
      begin
        Button := mbLeft;
        FSwapButtons := True;
        MouseCapture := True;
      end
      else if Button = mbLeft then
      begin
        EnableClick := True;
        CheckTitleButton(Cell.X - IndicatorOffset, Cell.Y, EnableClick);
        if EnableClick then
        begin
          MouseCapture := True;
          FTracking := True;
          FPressedCol := GetMasterColumn(Cell.X, Cell.Y);
          TrackButton(X, Y);
        end
        else
          Beep;
        Exit;
      end;
    end;
    if (Cell.X < FixedCols + IndicatorOffset) and Datalink.Active then
    begin
      if (dgIndicator in Options) then
        inherited MouseDown(Button, Shift, 1, Y)
      else if Cell.Y >= TitleOffset then
        if Cell.Y - Row <> 0 then
          Datalink.DataSet.MoveBy(Cell.Y - Row);
    end
    else
      inherited MouseDown(Button, Shift, X, Y);
    MouseDownEvent := OnMouseDown;
    if Assigned(MouseDownEvent) then
      MouseDownEvent(Self, Button, Shift, X, Y);
  end;
end;

procedure TRxDBGrid.MouseMove(Shift: TShiftState; X, Y: Integer);
begin
  if FTracking then
    TrackButton(X, Y);
  inherited MouseMove(Shift, X, Y);
end;

procedure TRxDBGrid.MouseUp(Button: TMouseButton; Shift: TShiftState;
  X, Y: Integer);
var
  Cell: TGridCoord;
  ACol: Longint;
  DoClick: Boolean;
begin
  if FTracking and (FPressedCol <> nil) then
  begin
    Cell := MouseCoord(X, Y);
    DoClick := PtInRect(Rect(0, 0, ClientWidth, ClientHeight), Point(X, Y)) and
      (Cell.Y < TitleOffset) and
      (FPressedCol = GetMasterColumn(Cell.X, Cell.Y));
    StopTracking;
    if DoClick then
    begin
      ACol := Cell.X;
      if (dgIndicator in Options) then
        Dec(ACol);
      if (Datalink <> nil) and Datalink.Active and (ACol >= 0) and (ACol <
{$IFDEF WIN32} Columns.Count {$ELSE} FieldCount {$ENDIF}) then
      begin
        DoTitleClick(FPressedCol.Index, FPressedCol.Field);
      end;
    end;
  end
  else if FSwapButtons then
  begin
    FSwapButtons := False;
    MouseCapture := False;
    if Button = mbRight then
      Button := mbLeft;
  end;
  inherited MouseUp(Button, Shift, X, Y);
end;

procedure TRxDBGrid.WMRButtonUp(var Message: TWMMouse);
begin
  if not(FGridState in [gsColMoving, gsRowMoving]) then
    inherited
  else if not(csNoStdEvents in ControlStyle) then
    with Message do
      MouseUp(mbRight, KeysToShiftState(Keys), XPos, YPos);
end;

procedure TRxDBGrid.WMCancelMode(var Message: TMessage);
begin
  StopTracking;
  inherited;
end;

procedure TRxDBGrid.WMChar(var Msg: TWMChar);

  function DoKeyPress(var Msg: TWMChar): Boolean;
  var
    Form: TCustomForm;
    Ch: Char;
  begin
    Result := True;
    Form := GetParentForm(Self);
    if (Form <> nil) and TForm(Form).KeyPreview and THack(Form).DoKeyPress(Msg)
    then
      Exit;
    with Msg do
    begin
      if Assigned(FOnKeyPress) then
      begin
        Ch := Char(CharCode);
        FOnKeyPress(Self, Ch);
        CharCode := word(Ch);
      end;
      if Char(CharCode) = #0 then
        Exit;
    end;
    Result := False;
  end;

begin
  if EditorMode or not DoKeyPress(Msg) then
    inherited;
end;

function GetGridBitmap(BmpType: TGridPicture): TBitmap;
begin
  if GridBitmaps[BmpType] = nil then
  begin
    GridBitmaps[BmpType] := TBitmap.Create;
    GridBitmaps[BmpType].Handle := LoadBitmap(hInstance, GridBmpNames[BmpType]);
  end;
  Result := GridBitmaps[BmpType];
end;

procedure TRxDBGrid.KeyPress(var Key: Char);
begin
  if EditorMode then
    inherited OnKeyPress := FOnKeyPress;
  try
    inherited KeyPress(Key);
  finally
    inherited OnKeyPress := nil;
  end;
end;

procedure TRxDBGrid.DefaultDataCellDraw(const Rect: TRect; Field: TField;
  State: TGridDrawState);
begin
  DefaultDrawDataCell(Rect, Field, State);
end;

function TRxDBGrid.GetMasterColumn(ACol, ARow: Longint): TColumn;
begin
  if (dgIndicator in Options) then
    Dec(ACol, IndicatorOffset);
  if (Datalink <> nil) and Datalink.Active and (ACol >= 0) and
    (ACol < Columns.Count) then
  begin
    Result := Columns[ACol];
    Result := ColumnAtDepth(Result, ARow);
  end
  else
    Result := nil;
end;

procedure TRxDBGrid.DrawCell(ACol, ARow: Longint; ARect: TRect;
  AState: TGridDrawState);

  function CalcTitleRect(Col: TColumn; ARow: Integer;
    var MasterCol: TColumn): TRect;
  { copied from Inprise's DbGrids.pas }
  var
    I, J: Integer;
    InBiDiMode: Boolean;
    DrawInfo: TGridDrawInfo;
  begin
    MasterCol := ColumnAtDepth(Col, ARow);
    if MasterCol = nil then
      Exit;
    I := DataToRawColumn(MasterCol.Index);
    if I >= LeftCol then
      J := MasterCol.Depth
    else
    begin
      if (FixedCols > 0) and (MasterCol.Index < FixedCols) then
      begin
        J := MasterCol.Depth;
      end
      else
      begin
        I := LeftCol;
        if Col.Depth > ARow then
          J := ARow
        else
          J := Col.Depth;
      end;
    end;
    Result := CellRect(I, J);
    InBiDiMode := UseRightToLeftAlignment and
      (Canvas.CanvasOrientation = coLeftToRight);
    for I := Col.Index to Columns.Count - 1 do
    begin
      if ColumnAtDepth(Columns[I], ARow) <> MasterCol then
        Break;
      if not InBiDiMode then
      begin
        J := CellRect(DataToRawColumn(I), ARow).Right;
        if J = 0 then
          Break;
        Result.Right := Max(Result.Right, J);
      end
      else
      begin
        J := CellRect(DataToRawColumn(I), ARow).Left;
        if J >= ClientWidth then
          Break;
        Result.Left := J;
      end;
    end;
    J := Col.Depth;
    if (J <= ARow) and (J < FixedRows - 1) then
    begin
      CalcFixedInfo(DrawInfo);
      Result.Bottom := DrawInfo.Vert.FixedBoundary -
        DrawInfo.Vert.EffectiveLineWidth;
    end;
  end;

  procedure DrawExpandBtn(var TitleRect, TextRect: TRect; InBiDiMode: Boolean;
    Expanded: Boolean); { copied from Inprise's DbGrids.pas }
  const
    ScrollArrows: array [Boolean, Boolean] of Integer =
      ((DFCS_SCROLLRIGHT, DFCS_SCROLLLEFT),
      (DFCS_SCROLLLEFT, DFCS_SCROLLRIGHT));
  var
    ButtonRect: TRect;
    I: Integer;
  begin
    I := GetSystemMetrics(SM_CXHSCROLL);
    if ((TextRect.Right - TextRect.Left) > I) then
    begin
      Dec(TextRect.Right, I);
      ButtonRect := TitleRect;
      ButtonRect.Left := TextRect.Right;
      I := SaveDC(Canvas.Handle);
      try
        Canvas.FillRect(ButtonRect);
        InflateRect(ButtonRect, -1, -1);
        with ButtonRect do
          IntersectClipRect(Canvas.Handle, Left, Top, Right, Bottom);
        InflateRect(ButtonRect, 1, 1);
        { DrawFrameControl doesn't draw properly when orienatation has changed.
          It draws as ExtTextOut does. }
        if InBiDiMode then { stretch the arrows box }
          Inc(ButtonRect.Right, GetSystemMetrics(SM_CXHSCROLL) + 4);
        DrawFrameControl(Canvas.Handle, ButtonRect, DFC_SCROLL,
          ScrollArrows[InBiDiMode, Expanded] or DFCS_FLAT);
      finally
        RestoreDC(Canvas.Handle, I);
      end;
      TitleRect.Right := ButtonRect.Left;
    end;
  end;

var
  { FrameOffs: Byte; }
  BackColor: TColor;
  SortMarker: TSortMarker;
  Indicator, ALeft: Integer;
  Down: Boolean;
  Bmp: TBitmap;
  SavePen: TColor;
  { OldActive: Longint;
    MultiSelected: Boolean;
    FixRect: TRect; }
  TitleRect, TextRect: TRect;
  AField: TField;
  MasterCol: TColumn;
  InBiDiMode: Boolean;
  DrawColumn: TColumn;
const
  EdgeFlag: array [Boolean] of UINT = (BDR_RAISEDINNER, BDR_SUNKENINNER);
begin
  inherited DrawCell(ACol, ARow, ARect, AState);
  InBiDiMode := Canvas.CanvasOrientation = coRightToLeft;
  if not(csLoading in ComponentState) and (FTitleButtons or (FixedCols > 0)) and
    (gdFixed in AState) and (dgTitles in Options) and (ARow < TitleOffset) then
  begin
    SavePen := Canvas.Pen.Color;
    try
      Canvas.Pen.Color := clWindowFrame;
      if (dgIndicator in Options) then
        Dec(ACol, IndicatorOffset);
      AField := nil;
      SortMarker := smNone;
      if (Datalink <> nil) and Datalink.Active and (ACol >= 0) and
        (ACol < Columns.Count) then
      begin
        DrawColumn := Columns[ACol];
        AField := DrawColumn.Field;
      end
      else
        DrawColumn := nil;
      if Assigned(DrawColumn) and not DrawColumn.Showing then
        Exit;
      TitleRect := CalcTitleRect(DrawColumn, ARow, MasterCol);
      if TitleRect.Right < ARect.Right then
        TitleRect.Right := ARect.Right;
      if MasterCol = nil then
        Exit
      else if MasterCol <> DrawColumn then
        AField := MasterCol.Field;
      DrawColumn := MasterCol;
      if ((dgColLines in Options) or FTitleButtons) and (ACol = FixedCols - 1)
      then
      begin
        if (ACol < Columns.Count - 1) and not(Columns[ACol + 1].Showing) then
        begin
          Canvas.MoveTo(TitleRect.Right, TitleRect.Top);
          Canvas.LineTo(TitleRect.Right, TitleRect.Bottom);
        end;
      end;
      if ((dgRowLines in Options) or FTitleButtons) and not MasterCol.Showing
      then
      begin
        Canvas.MoveTo(TitleRect.Left, TitleRect.Bottom);
        Canvas.LineTo(TitleRect.Right, TitleRect.Bottom);
      end;
      Down := FPressed and FTitleButtons and (FPressedCol = DrawColumn);
      if FTitleButtons or ([dgRowLines, dgColLines] * Options = [dgRowLines,
        dgColLines]) then
      begin
        DrawEdge(Canvas.Handle, TitleRect, EdgeFlag[Down], BF_BOTTOMRIGHT);
        DrawEdge(Canvas.Handle, TitleRect, EdgeFlag[Down], BF_TOPLEFT);
        InflateRect(TitleRect, -1, -1);
      end;
      Canvas.Font := TitleFont;
      Canvas.Brush.Color := FixedColor;
      if (DrawColumn <> nil) then
      begin
        Canvas.Font := DrawColumn.Title.Font;
        Canvas.Brush.Color := DrawColumn.Title.Color;
      end;
      if FTitleButtons and (AField <> nil) and Assigned(FOnGetBtnParams) then
      begin
        BackColor := Canvas.Brush.Color;
        FOnGetBtnParams(Self, AField, Canvas.Font, BackColor, SortMarker, Down);
        Canvas.Brush.Color := BackColor;
      end;
      if Down then
      begin
        Inc(TitleRect.Left);
        Inc(TitleRect.Top);
      end;
      ARect := TitleRect;
      if (Datalink = nil) or not Datalink.Active then
        Canvas.FillRect(TitleRect)
      else if (DrawColumn <> nil) then
      begin
        case SortMarker of
          smDown:
            Bmp := GetGridBitmap(gpMarkDown);
          smUp:
            Bmp := GetGridBitmap(gpMarkUp);
        else
          Bmp := nil;
        end;
        if Bmp <> nil then
          Indicator := Bmp.Width + 6
        else
          Indicator := 1;
        TextRect := TitleRect;
        if DrawColumn.Expandable then
          DrawExpandBtn(TitleRect, TextRect, InBiDiMode, DrawColumn.Expanded);
        with DrawColumn.Title do
          DrawCellText(Self, ACol, ARow, MinimizeText(Caption, Canvas,
            WidthOf(TextRect) - Indicator), TextRect, Alignment, vaCenter,
            IsRightToLeft);
        if Bmp <> nil then
        begin
          ALeft := TitleRect.Right - Bmp.Width - 3;
          if Down then
            Inc(ALeft);
          if IsRightToLeft then
            ALeft := TitleRect.Left + 3;
          if (ALeft > TitleRect.Left) and (ALeft + Bmp.Width < TitleRect.Right)
          then
            DrawBitmapTransparent(Canvas, ALeft,
              (TitleRect.Bottom + TitleRect.Top - Bmp.Height) div 2, Bmp,
              clFuchsia);
        end;
      end
      else
        DrawCellText(Self, ACol, ARow, '', ARect, taLeftJustify,
          vaCenter, False);
    finally
      Canvas.Pen.Color := SavePen;
    end;
  end
  else
  begin
    Canvas.Font := Self.Font;
    if (Datalink <> nil) and Datalink.Active and (ACol >= 0) and
      (ACol < Columns.Count) then
    begin
      DrawColumn := Columns[ACol];
      if DrawColumn <> nil then
        Canvas.Font := DrawColumn.Font;
    end;
  end;
end;

procedure TRxDBGrid.DrawColumnCell(const Rect: TRect; DataCol: Integer;
  Column: TColumn; State: TGridDrawState);
var
  { I: Integer; }
  NewBackgrnd: TColor;
  Highlight: Boolean;
  { Bmp: TBitmap; }
  Field: TField;
begin
  Field := Column.Field;
  NewBackgrnd := Canvas.Brush.Color;
  Highlight := (gdSelected in State) and
    ((dgAlwaysShowSelection in Options) or Focused);
  GetCellProps(Field, Canvas.Font, NewBackgrnd,
    Highlight { or ActiveRowSelected } );
  Canvas.Brush.Color := NewBackgrnd;
  if FDefaultDrawing then
  begin
    { I := GetImageIndex(Field);
      if I >= 0 then begin
      Bmp := GetGridBitmap(TGridPicture(I));
      Canvas.FillRect(Rect);
      DrawBitmapTransparent(Canvas, (Rect.Left + Rect.Right - Bmp.Width) div 2,
      (Rect.Top + Rect.Bottom - Bmp.Height) div 2, Bmp, clOlive);
      end else }
    DefaultDrawColumnCell(Rect, DataCol, Column, State);
  end;
  if Columns.State = csDefault then
    inherited DrawDataCell(Rect, Field, State);
  inherited DrawColumnCell(Rect, DataCol, Column, State);
  if FDefaultDrawing and Highlight and not(csDesigning in ComponentState) and
    not(dgRowSelect in Options) and (ValidParentForm(Self).ActiveControl = Self)
  then
    Canvas.DrawFocusRect(Rect);
end;

procedure TRxDBGrid.DrawDataCell(const Rect: TRect; Field: TField;
  State: TGridDrawState);
begin
end;

procedure TRxDBGrid.MouseToCell(X, Y: Integer; var ACol, ARow: Longint);
var
  Coord: TGridCoord;
begin
  Coord := MouseCoord(X, Y);
  ACol := Coord.X;
  ARow := Coord.Y;
end;


end.
