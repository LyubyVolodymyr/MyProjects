unit FS;

interface
{$DEFINE USE_RX_GRID}


uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DB, DBGrids, Math, Menus, ExtCtrls, ComObj, getPeriodFrm,
  {$IFDEF USE_RX_GRID}
  RxGridUnit
  {$ENDIF}
  ;

const
  MaxSelectRows = 10000;

{$R *.R32}


type
  TFilteringMode = (fmOff, fmText, fmDates);

const
  dbsgMaxWords = 100;
  dbsgMaxOrWords = 10;

type
  TsgFilterType = (ftNone, ftKeyWords, ftSelected, ftDates);

type
  TDBSearchGrid = class({$IFDEF USE_RX_GRID} TRxDBGrid {$ELSE} TDBGrid {$ENDIF} )
  private
    { Private declarations }
    Selection: record
      SelectCount: Integer;
      Selected: array [1 .. MaxSelectRows] of variant;
    end;

    SearchStr: string;
    HintW: THintWindow;
    FNeedSearch: Boolean;
    HintLocked: Boolean;
    ExportMenu: TMenuItem;
    FontChangeMenu: TMenuItem;
    FExportCaption: string;
    FFontChangeCaption: string;
    FLocateNotSupport: Boolean;
    FDisableExportMenu: Boolean;
    FDisableFontChangemenu: Boolean;
    FSavedTitleCaption: String;
    FSavedTitleColor: TColor;
    // FFilteringActive:TFilteringMode;
    // FilterStr:string;
    FStartPeriod, FEndPeriod: TDate;
    Words: array [1 .. dbsgMaxWords] of record
          wc: Integer;
          wds: array [1 .. dbsgMaxOrWords] of String;
      end;

    WordCount:Integer;
    SFI:Integer;
    FDefaultFiltering:Boolean;
    // FStoredFilerRec:TFilterRecordEvent;
    FSearchField:string;
    FGroupSelectEnable:Boolean;
    FilteringActive:TsgFilterType;
    SavedOnFilterRecord:TFilterRecordEvent;
    SavedFiltered:Boolean;
    PredFilterStr:string;

    procedure RefreshView;
    procedure WMKillFocus(var Message: TMessage); message WM_KillFocus;
  protected
    { Protected declarations }
    procedure KeyPress(var Key: Char);override;
    procedure KeyDown(var Key: word; Shift: TShiftState);override;
    procedure DoExit;override;
    procedure ColEnter;override;
    procedure Scroll(Distance: Integer);override;
    {$IFNDEF USE_RX_GRID}
    procedure DrawColumnCell(const Rect: TRect; DataCol: Integer;
      Column: TColumn; State: TGridDrawState);override;
    {$ENDIF}
    procedure Loaded;override;
    procedure DecodeFilterString(FiltStr: string);
    {$IFDEF USE_RX_GRID}
    procedure GetCellProps(Field: TField; AFont: TFont; var Background: TColor;
      Highlight: Boolean);  override;
    {$ELSE}
    procedure GetCellProps(Field: TField; AFont: TFont; var Background: TColor;
      Highlight: Boolean);  virtual;
    {$ENDIF}
    function GetSelectedKey(Index: Integer): variant;
  public
    { Public declarations }
    function CellRect(ACol, ARow: Longint): TRect;
    procedure LeaveHint;
    Procedure DoHint;
    Constructor Create(AOwner: TComponent);
  override;
    Destructor Destroy;Override;
    function DoSearch(Text: string; fgReturn: Boolean): Boolean;virtual;
    function NeedHandleReturn: Boolean;
    procedure ExportClick(Sender: TObject);
    procedure FontChangeClick(Sender: TObject);
    procedure StartFilter;
    procedure StopFilterA(fgLocateCurrent: Boolean);
    procedure StartFilterExt(FiltString, FieldName: string);
    procedure OnFilterRec(DataSet: TDataSet; var Accept: Boolean);
    procedure PasteCurrentCell;
    procedure ClearSelection;
    procedure ToggleSelection;
    function IsSelected(Key: variant): Boolean;
    function IsCurrentSelected: Boolean;
    procedure SelectAllTo;
    procedure StartFilterSelected;
    property SelectedKeys[Index: Integer]: variant read GetSelectedKey;
    // Zero-based array
    property SelectedRecords: Integer read Selection.SelectCount;

  published
    { Published declarations }
    property NeedSearch: Boolean read FNeedSearch write FNeedSearch default True;
    property HintWindow: THintWindow read HintW;
    property ExportCaption: string read FExportCaption write FExportCaption;
    property FonChangeCaption: string read FFontChangeCaption
        write FFontChangeCaption;
    property LocateNotSupport: Boolean read FLocateNotSupport
        write FLocateNotSupport default False;
    property DisableExportMenu: Boolean read FDisableExportMenu
        write FDisableExportMenu default False;
    property DisableFontChangeMenu: Boolean read FDisableFontChangemenu
        write FDisableFontChangemenu default True;
    property DefaultFiltering: Boolean read FDefaultFiltering
        write FDefaultFiltering default True;
    property SearchField: string read FSearchField write FSearchField;
    property GroupSelectEnable: Boolean read FGroupSelectEnable
        write FGroupSelectEnable default False;
  end;

procedure Register;

implementation

uses {bde;} clipbrd, RefreshSQL, Variants;

const
  DefaultExportMenuCaption = 'Експортувати в Excel';

const
  DefaultFontChangeMenuCaption = 'Налаштувати шрифт списку';
  MagicGroup = 253;

procedure Register;
begin
  RegisterComponents('Samples', [TDBSearchGrid]);
end;


{ --------------------------- TDBSearchGrid ----------------------- }

procedure TDBSearchGrid.Loaded;
var
  m, newMenu: TPopupMenu;
  Own: TControl;
  Divide: TMenuItem;
  c: Integer;
  OwnForm: TForm;
begin
  inherited;
  if (csDesigning in ComponentState) then
    Exit;
  m := nil;
  OwnForm := nil;
  if Assigned(PopupMenu) then
    m := PopupMenu
  else
  begin
    Own := Parent;
    repeat
      if Own is TPanel then
      begin
        if Assigned((Own as TPanel).PopupMenu) then
          m := (Own as TPanel).PopupMenu;
      end
      else if Own is TForm then
      begin
        if Assigned((Own as TForm).PopupMenu) then
          m := (Own as TForm).PopupMenu;
        OwnForm := Own as TForm;
      end;
      if not Assigned(m) then
      begin
        if Own is TForm then
          Own := nil
        else if Assigned(Own.Parent) then
          Own := Own.Parent
        else
          Own := nil;
      end;
    until Assigned(m) or not Assigned(Own);
  end;

  if Assigned(m) then
  begin
    c := m.Items.Count;
    while (c > 0) and (m.Items[c - 1].GroupIndex <> MagicGroup) do
      Dec(c);
    if c = 0 then
    begin
      Divide := TMenuItem.Create(m);
      Divide.Caption := '-';
      Divide.GroupIndex := MagicGroup;
      m.Items.Add(Divide);
    end;
    newMenu := m;
  end
  else
  begin
    if Assigned(OwnForm) then
    begin
      newMenu := TPopupMenu.Create(OwnForm);
      OwnForm.PopupMenu := newMenu;
    end
    else
      newMenu := nil;
  end;
  if Assigned(newMenu) then
  begin
    if not Assigned(ExportMenu) and not DisableExportMenu then
    begin
      ExportMenu := TMenuItem.Create(newMenu);
      if ExportCaption = '' then
        ExportMenu.Caption := DefaultExportMenuCaption
      else
        ExportMenu.Caption := ExportCaption;
      ExportMenu.OnClick := ExportClick;
      ExportMenu.GroupIndex := MagicGroup;
      newMenu.Items.Add(ExportMenu);
    end;
    if not Assigned(FontChangeMenu) and not DisableFontChangeMenu then
    begin
      FontChangeMenu := TMenuItem.Create(newMenu);
      if FFontChangeCaption = '' then
        FontChangeMenu.Caption := DefaultFontChangeMenuCaption
      else
        FontChangeMenu.Caption := FFontChangeCaption;
      FontChangeMenu.OnClick := FontChangeClick;
      FontChangeMenu.GroupIndex := MagicGroup + 1;
      newMenu.Items.Add(FontChangeMenu);
    end;
  end;
end;

procedure TDBSearchGrid.FontChangeClick(Sender: TObject);
var
  FD: TFontDialog;
begin
  if DisableFontChangeMenu then
    Exit;
  FD := TFontDialog.Create(Self);
  try
    If Font.Name <> '' then
      FD.Font.Name := Font.Name;
    If Font.Size > 0 then
      FD.Font.Size := Font.Size;
    FD.Font.Style := Font.Style;
    if FD.Execute then
    begin
      Font.Name := FD.Font.Name;
      Font.Size := FD.Font.Size;
      Font.Style := FD.Font.Style;
    end;
  finally
    FD.Free;
  end;
end;

procedure TDBSearchGrid.ExportClick(Sender: TObject);
var
  Excel: variant;
  I, fldCount, Pos: Integer;
  ds: TDataSet;
begin
  if not Assigned(DataSource) then
    Exit;
  if not Assigned(DataSource.DataSet) then
    Exit;
  ds := DataSource.DataSet;
  if not ds.Active then
    Exit;
  Excel := CreateOLEObject('Excel.Application');
  Excel.WorkBooks.Add;
  fldCount := Columns.Count;
  for I := 1 to fldCount do
  begin
    Excel.Cells[1, I].Value := Columns[I - 1].Title.Caption;
    Excel.Cells[1, I].Font.Bold := True;
    Excel.Columns[I].ColumnWidth := Columns[I - 1].Field.DisplayWidth;
    case Columns[I - 1].Alignment of
      taLeftJustify:
        Excel.Columns[I].HorizontalAlignment := 2;
      taCenter:
        Excel.Columns[I].HorizontalAlignment := 3;
      taRightJustify:
        Excel.Columns[I].HorizontalAlignment := 4;
    end;
    Excel.Columns[I].NumberFormat := '@';
  end;
  ds.DisableControls;
  ds.First;
  Pos := 2;
  while not ds.Eof do
  begin
    for I := 1 to fldCount do
      Excel.Cells[Pos, I].Value := Columns[I - 1].Field.DisplayText;
    ds.Next;
    Inc(Pos);
  end;
  Excel.Application.Visible := True;
  ds.EnableControls;
end;

function TDBSearchGrid.CellRect(ACol, ARow: Longint): TRect;
begin
  Result := inherited CellRect(ACol, ARow);
end;

destructor TDBSearchGrid.Destroy;
begin
  LeaveHint;
  Inherited Destroy;
end;

{$IFNDEF USE_RX_GRID}
procedure TDBSearchGrid.DrawColumnCell(const Rect: TRect; DataCol: Integer;
  Column: TColumn; State: TGridDrawState);
var
  NewBackgrnd: TColor;
  Highlight: Boolean;
  Field: TField;
begin
  Field := Column.Field;
  NewBackgrnd := Canvas.Brush.Color;
  Highlight := (gdSelected in State) and
    ((dgAlwaysShowSelection in Options) or Focused);
  GetCellProps(Field, Canvas.Font, NewBackgrnd,
    Highlight { or ActiveRowSelected } );
  Canvas.Brush.Color := NewBackgrnd;
  DefaultDrawColumnCell(Rect, DataCol, Column, State);
  if Columns.State = csDefault then
    inherited DrawDataCell(Rect, Field, State);
  inherited DrawColumnCell(Rect, DataCol, Column, State);
  if Highlight and not(csDesigning in ComponentState) and
    not(dgRowSelect in Options) and (ValidParentForm(Self).ActiveControl = Self)
  then
    Canvas.DrawFocusRect(Rect);
end;
{$ENDIF}

function TDBSearchGrid.NeedHandleReturn;
begin
  NeedHandleReturn := (DataSource <> nil) and (DataSource.DataSet <> nil) and
    (DataSource.Enabled) and (DataSource.DataSet.Active) and
    (SelectedField.DataType <> ftString) and (HintWindow <> nil);
end;

constructor TDBSearchGrid.Create;
begin
  inherited Create(AOwner);
  SearchStr := '';
  HintW := nil;
  FNeedSearch := True;
  HintLocked := False;
  ExportMenu := nil;
  FontChangeMenu := nil;
  FLocateNotSupport := False;
  FDefaultFiltering := True;
  FilteringActive := ftNone;
  // FilterStr:='';
  SFI := 0;
  FDisableExportMenu := False;
  FDisableFontChangemenu := True;
  FGroupSelectEnable := False;
end;

procedure TDBSearchGrid.ColEnter;
begin
  inherited ColEnter;
  LeaveHint;
  SearchStr := '';
end;

procedure TDBSearchGrid.Scroll;
begin
  Inherited Scroll(Distance);
  if not HintLocked then
  begin
    LeaveHint;
    SearchStr := '';
  end;
end;

procedure TDBSearchGrid.DoExit;
begin
  LeaveHint;
  SearchStr := '';
  inherited DoExit;
end;

procedure TDBSearchGrid.LeaveHint;
begin
  if HintW <> nil then
  begin
    HintW.ReleaseHandle;
    HintW.Free;
    HintW := nil;
    { SearchStr:=''; }
  end;
end;

procedure TDBSearchGrid.DoHint;
var
  Rec: TRect;
begin
  LeaveHint;
  HintW := THintWindow.Create(Self);
  Rec := CellRect(Col, Row);
  HintW.ActivateHint(Bounds(Rec.Left + ClientOrigin.X,
    Rec.Bottom + ClientOrigin.Y, HintW.Canvas.TextWidth(SearchStr) + 5, 18),
    SearchStr);
  { HintW.Refresh; }
end;

procedure TDBSearchGrid.RefreshView;
begin
  if Assigned(DataSource) and Assigned(DataSource.DataSet) and
    (DataSource.DataSet is TRefreshedQuery) then
    (DataSource.DataSet as TRefreshedQuery).Refresh;
end;

procedure TDBSearchGrid.KeyDown;
var
  SS: string;
  fgHandled: Boolean;
begin
  if FGroupSelectEnable then
  begin
    fgHandled := True;
    if (Key = VK_Right) and (Shift = [ssCtrl]) and FDefaultFiltering then
      StartFilterSelected
    else if (Key = vk_return) and (Shift = [ssCtrl, ssAlt]) then
      ClearSelection
    else if (Key = vk_return) and (Shift = [ssCtrl]) then
      ToggleSelection
    else if (Key = vk_return) and (Shift = [ssCtrl, ssShift]) then
      SelectAllTo
    else
      fgHandled := False;
  end
  else
    fgHandled := False;

  if not fgHandled then
  begin
    if (Key = VK_DOWN) and (Shift = [ssCtrl]) and FDefaultFiltering then
      StartFilter
    else if (Key = VK_INSERT) and (Shift = [ssCtrl]) then
      PasteCurrentCell
    else if (Key = VK_UP) and (Shift = [ssCtrl]) and FDefaultFiltering then
      StopFilterA(True)
    else if (Key = VK_F5) and (Shift = [ssCtrl]) then
      RefreshView
    {$IFDEF USE_RX_GRID}
    else if (Key = VK_DOWN) and (Shift = [ssAlt]) then
    begin
      DoTitleClick(Col, SelectedField);
    end
    {$ENDIF}
    else if (not FNeedSearch) or (SearchStr = '') then
      inherited
    else
    begin
      if Key = VK_F3 then
        with DataSource.DataSet do
        begin
          DisableControls;
          LeaveHint;
          SS := AnsiUpperCase(SearchStr);
          Next;
          while not Eof and
            (AnsiUpperCase(Copy(SelectedField.DisplayText, 1, Length(SS))
            ) <> SS) do
            Next;
          EnableControls;
          if not Eof then
            DoHint;
        end;
      inherited;
    end;
  end;
end;

procedure TDBSearchGrid.KeyPress;
begin
  if not FNeedSearch then
    inherited KeyPress(Key)
  else
  begin
    if Key in [' ' .. #255] then
    begin
      SearchStr := SearchStr + Key;
      LeaveHint;
      if not DoSearch(SearchStr, False) then
        SearchStr := Copy(SearchStr, 1, Length(SearchStr) - 1);
      if Length(SearchStr) <> 0 then
        DoHint;
    end
    else if (Key = Char(vk_return)) and (Length(SearchStr) > 0) then
    begin
      LeaveHint;
      if DoSearch(SearchStr, True) then
        SearchStr := ''
      else
      begin
        Key := #0;
        DoHint;
      end;
    end
    else if (Key = Char(VK_ESCAPE)) then
    begin
      if HintW <> nil then
      begin
        LeaveHint;
        SearchStr := '';
      end;
    end
    else if Key = Char(VK_BACK) then
    begin
      if Length(SearchStr) > 1 then
      begin
        SearchStr := Copy(SearchStr, 1, Length(SearchStr) - 1);
        DoHint;
      end
      else
        LeaveHint;
    end;
    Inherited KeyPress(Key);
  end;
end;

function TDBSearchGrid.DoSearch;
var
  BM: TBookmark;
  SF: TField;
begin
  Result := True;
  HintLocked := True;
  if (DataSource <> nil) and (DataSource.DataSet <> nil) and
    (DataSource.Enabled) and (DataSource.DataSet.Active) then
    with DataSource.DataSet do
    begin
      if Trim(FSearchField) <> '' then
        SF := FieldByName(FSearchField)
      else
        SF := SelectedField;
      if (SF.DataType = ftString) or Assigned(SF.OnGetText) then
      begin
        DisableControls;
        if (Assigned(SF.OnGetText)) or LocateNotSupport then
        begin
          BM := GetBookmark;
          try
            { First; }
            while not Eof and
              (AnsiUpperCase(Copy(SF.DisplayText, 1, Length(Text))) <>
              AnsiUpperCase(TrimRight(Text))) do
              Next;
            if Eof then
            begin
              First;
              While not Eof and
                (AnsiUpperCase(Copy(SF.DisplayText, 1, Length(Text))) <>
                AnsiUpperCase(TrimRight(Text))) do
                Next;
              if Eof then
              begin
                Result := False;
                MessageBEEP(mb_iconquestion);
                GotoBookmark(BM);
              end;
            end;
          finally
            FreeBookmark(BM);
          end;
        end
        else // not (Assigned(SF.OnGetText)) or LocateNotSupport
        begin
          if not Locate(SF.FieldName, Text, [loPartialKey, loCaseInsensitive])
          then
          begin
            MessageBEEP(mb_iconquestion);
            Result := False;
          end;
        end;
        EnableControls;
      end
      else // not ((SF.DataType=ftString) or Assigned(SF.OnGetText))
        if fgReturn then
        begin
          if not Locate(SF.FieldName, Text, [loCaseInsensitive]) then
          begin
            MessageBEEP(mb_iconquestion);
            Result := False;
          end;
        end;
    end;
  HintLocked := False;
end;

procedure TDBSearchGrid.StartFilter;
var
  I: Integer;
  FilterStr: string;
begin
  if FilteringActive = ftSelected then
    Exit;
  // if FFilteringActive<>fmOff then exit;
  if not Assigned(DataSource) then
    Exit;
  if not Assigned(DataSource.DataSet) then
    Exit;
  if FilteringActive <> ftNone then
    StopFilterA(True);
  { if Assigned(DataSource.DataSet.OnFilterRecord) and DataSource.DataSet.Filtered then
    raise Exception.Create('Встановлено інший фільтр. Фільтрування неможливе.'); }
  if (SelectedField is TDateField) or (SelectedField is TDateTimeField) then
  begin
    if InputPeriod(FStartPeriod, FEndPeriod) then
    begin
      SFI := SelectedField.Index;

      SavedOnFilterRecord := DataSource.DataSet.OnFilterRecord;
      SavedFiltered := DataSource.DataSet.Filtered;
      DataSource.DataSet.Filtered := False;

      // FStoredFilerRec:=DataSource.DataSet.OnFilterRecord;
      DataSource.DataSet.OnFilterRecord := Self.OnFilterRec;
      FilteringActive := ftDates;
      DataSource.DataSet.Filtered := True;

      for I := 0 to Columns.Count - 1 do
        if Columns[I].FieldName = SelectedField.FieldName then
        begin
          FSavedTitleCaption := Columns[I].Title.Caption;
          Columns[I].Title.Caption := 'Фільтр:(період дат)';
          FSavedTitleColor := Columns[I].Title.Color;
          Columns[I].Title.Color := clYellow;
        end;

    end;
  end
  else
  begin
    FilterStr := PredFilterStr;
    if InputQuery('Фільтрування', 'Введіль ключові слова фільтру', FilterStr)
    then
    begin
      PredFilterStr := FilterStr;
      DecodeFilterString(AnsiUpperCase(FilterStr));
      SFI := SelectedField.Index;

      SavedOnFilterRecord := DataSource.DataSet.OnFilterRecord;
      SavedFiltered := DataSource.DataSet.Filtered;
      DataSource.DataSet.Filtered := False;
      // FStoredFilerRec:=DataSource.DataSet.OnFilterRecord;
      DataSource.DataSet.OnFilterRecord := Self.OnFilterRec;
      FilteringActive := ftKeyWords;
      DataSource.DataSet.Filtered := True;

      for I := 0 to Columns.Count - 1 do
        if Columns[I].FieldName = SelectedField.FieldName then
        begin
          FSavedTitleCaption := Columns[I].Title.Caption;
          Columns[I].Title.Caption := 'Фільтр:"' + FilterStr + '"';
          FSavedTitleColor := Columns[I].Title.Color;
          Columns[I].Title.Color := clYellow;
        end;
    end;
  end;
end;

procedure TDBSearchGrid.StopFilterA(fgLocateCurrent: Boolean);
var
  BM: TBookmark;
  I: Integer;
begin
  if FilteringActive = ftNone then
    Exit;
  if not Assigned(DataSource) then
    Exit;
  if not Assigned(DataSource.DataSet) then
    Exit;
  if fgLocateCurrent then
    BM := DataSource.DataSet.GetBookmark;
  try
    for I := 0 to Columns.Count - 1 do
      if Columns[I].FieldName = DataSource.DataSet.Fields[SFI].FieldName then
      begin
        Columns[I].Title.Caption := FSavedTitleCaption;
        Columns[I].Title.Color := FSavedTitleColor;
      end;
    DataSource.DataSet.Filtered := False;
    DataSource.DataSet.OnFilterRecord := SavedOnFilterRecord;
    DataSource.DataSet.Filtered := SavedFiltered;
    FilteringActive := ftNone;
    if fgLocateCurrent then
      DataSource.DataSet.GotoBookmark(BM);
  finally
    if fgLocateCurrent then
      DataSource.DataSet.FreeBookmark(BM);

  end

end;

procedure TDBSearchGrid.OnFilterRec(DataSet: TDataSet; var Accept: Boolean);
var
  S: string;
  Ac, Ac2, Ac3: Boolean;
  I, J: Integer;
begin

  if Assigned(SavedOnFilterRecord) then
    SavedOnFilterRecord(DataSet, Accept);

  Ac := True;

  case FilteringActive of
    ftNone:
      Ac := True; { Accept all of filter off }
    ftDates:
      begin
        if (DataSet.Fields[SFI] is TDateField) then
          with (DataSet.Fields[SFI] as TDateField) do
            Ac := (asDateTime >= FStartPeriod) and (asDateTime <= FEndPeriod)
        else if (DataSet.Fields[SFI] is TDateTimeField) then
          with (DataSet.Fields[SFI] as TDateTimeField) do
            Ac := (asDateTime >= FStartPeriod) and (asDateTime <= FEndPeriod);
      end;
    ftKeyWords:
      begin
        S := AnsiUpperCase(DataSet.Fields[SFI].DisplayText);
        Ac3 := True;
        for I := 1 to WordCount do
        begin
          Ac2 := False;
          for J := 1 to Words[I].wc do
            if Pos(Words[I].wds[J], S) <> 0 then
              Ac2 := True;

          if not Ac2 then
            Ac3 := False;
        end;
        Ac := Ac3;
      end;
    ftSelected:
      Ac := IsCurrentSelected;

  end; { case }
  Accept := Ac and Accept;

end;

procedure TDBSearchGrid.DecodeFilterString(FiltStr: string);
var
  S, s1: string;
  p: Integer;

  procedure SetWord;
  var
    s2: string;
    wc: Integer;
  begin
    wc := 0;
    while s1 <> '' do
    begin
      p := Pos('!', s1);
      if p > 0 then
      begin
        s2 := Trim(Copy(s1, 1, p - 1));
        s1 := Trim(Copy(s1, p + 1, Length(s1) - p));
      end
      else
      begin
        s2 := Trim(s1);
        s1 := '';
      end;
      Inc(wc);
      Words[WordCount].wc := wc;
      Words[WordCount].wds[wc] := s2;
    end;
  end;

begin
  S := Trim(FiltStr);
  WordCount := 0;
  while S <> '' do
  begin
    p := Pos(' ', S);
    if p > 0 then
    begin
      s1 := Trim(Copy(S, 1, p - 1));
      S := Trim(Copy(S, p + 1, Length(S) - p));
    end
    else
    begin
      s1 := Trim(S);
      S := '';
    end;
    Inc(WordCount);
    SetWord;
  end;
end;

procedure TDBSearchGrid.PasteCurrentCell;
var
  S: string;
begin
  S := SelectedField.DisplayText;
  Clipboard.asText := S;
end;

procedure TDBSearchGrid.WMKillFocus(var Message: TMessage);
begin
  LeaveHint;
  inherited;
end;

procedure TDBSearchGrid.StartFilterExt(FiltString, FieldName: string);
var
  I: Integer;
begin
  if FiltString = '' then
  begin
    StopFilterA(True);
    Exit;
  end;

  if not Assigned(DataSource) then
    Exit;
  if not Assigned(DataSource.DataSet) then
    Exit;

  DataSource.DataSet.DisableControls;
  try
    if FilteringActive <> ftNone then
      StopFilterA(False);

    if Assigned(DataSource.DataSet.OnFilterRecord) or DataSource.DataSet.Filtered
    then
      raise Exception.Create
        ('Встановлено інший фільтр. Фільтрування неможливе.');
    begin
      DecodeFilterString(AnsiUpperCase(FiltString));
      SFI := DataSource.DataSet.FieldByName(FieldName).Index;

      SavedOnFilterRecord := DataSource.DataSet.OnFilterRecord;
      SavedFiltered := DataSource.DataSet.Filtered;
      DataSource.DataSet.Filtered := False;
      // FStoredFilerRec:=DataSource.DataSet.OnFilterRecord;
      DataSource.DataSet.OnFilterRecord := Self.OnFilterRec;
      FilteringActive := ftKeyWords;
      DataSource.DataSet.Filtered := True;

      for I := 0 to Columns.Count - 1 do
        if Columns[I].FieldName = FieldName then
        begin
          FSavedTitleCaption := Columns[I].Title.Caption;
          Columns[I].Title.Caption := 'Фільтр:"' + FiltString + '"';
          FSavedTitleColor := Columns[I].Title.Color;
          Columns[I].Title.Color := clYellow;
        end;

    end;
  finally
    DataSource.DataSet.EnableControls;
  end;
end;

procedure TDBSearchGrid.StartFilterSelected;
begin
  if FilteringActive = ftNone then
  begin
    // FilterCount:=1; {!!!}
    SavedOnFilterRecord := DataSource.DataSet.OnFilterRecord;
    SavedFiltered := DataSource.DataSet.Filtered;
    DataSource.DataSet.Filtered := False;
    DataSource.DataSet.OnFilterRecord := Self.OnFilterRec;
    FilteringActive := ftSelected;
    DataSource.DataSet.Filtered := True;
  end
  else
    Exit;
end;

procedure TDBSearchGrid.ClearSelection;
begin
  if not FGroupSelectEnable then
    Exit;
  Selection.SelectCount := 0;
  Refresh;
end;

function TDBSearchGrid.IsSelected(Key: variant): Boolean;
var
  I: Integer;
begin
  I := 1;
  Result := False;
  if not FGroupSelectEnable then
    Exit;
  while (I <= Selection.SelectCount) and not Result do
  begin
    Result := Key = Selection.Selected[I];
    Inc(I);
  end;
end;

function TDBSearchGrid.IsCurrentSelected: Boolean;
var
  CurrentKey: variant;
  FUniqueField: TField;
begin
  Result := False;
  if not FGroupSelectEnable then
    Exit;

  if not Assigned(DataSource) or not Assigned(DataSource.DataSet) or
    not(DataSource.DataSet is TRefreshedQuery) then
    raise Exception.Create('Не задано джерело даних або не вірний тип.');

  FUniqueField := (DataSource.DataSet as TRefreshedQuery).UniqueField;
  if not Assigned(FUniqueField) then
    raise Exception.Create('Не задано унікальне поле джерела даних');

  CurrentKey := FUniqueField.Value;

  Result := IsSelected(CurrentKey);
end;

procedure TDBSearchGrid.ToggleSelection;
var
  CurrentKey: variant;
  FUniqueField: TField;
  I: Integer;
begin
  if not FGroupSelectEnable then
    Exit;

  if not Assigned(DataSource) or not Assigned(DataSource.DataSet) or
    not(DataSource.DataSet is TRefreshedQuery) then
    raise Exception.Create('Не задано джерело даних або не вірний тип.');

  FUniqueField := (DataSource.DataSet as TRefreshedQuery).UniqueField;
  if not Assigned(FUniqueField) then
    raise Exception.Create('Не задано унікальне поле джерела даних');

  CurrentKey := FUniqueField.Value;;
  if IsSelected(CurrentKey) then
  begin
    I := 1;
    while (I <= Selection.SelectCount) and
      (CurrentKey <> Selection.Selected[I]) do
      Inc(I);
    while (I < Selection.SelectCount) do
    begin
      Selection.Selected[I] := Selection.Selected[I + 1];
      Inc(I);
    end;
    Dec(Selection.SelectCount);
  end
  else
  begin
    if Selection.SelectCount >= MaxSelectRows then
      raise Exception.Create
        ('Перевищено максимально допустиму кількість виділених рядків');
    Inc(Selection.SelectCount);
    Selection.Selected[Selection.SelectCount] := CurrentKey;
  end;
  Refresh;
end;

procedure TDBSearchGrid.SelectAllTo;
var
  CurrentKey: variant;
  FUniqueField: TField;
  fgExit: Boolean;
  FirstRecord, lastRecord: variant;

type
  TMovingMode = (mmStart, mmStop, mmStartByCurrent, mmStartBySelected);
var
  Mode: TMovingMode;

begin
  if not FGroupSelectEnable then
    Exit;

  if not Assigned(DataSource) or not Assigned(DataSource.DataSet) or
    not(DataSource.DataSet is TRefreshedQuery) then
    raise Exception.Create('Не задано джерело даних або не вірний тип.');

  FUniqueField := (DataSource.DataSet as TRefreshedQuery).UniqueField;
  if not Assigned(FUniqueField) then
    raise Exception.Create('Не задано унікальне поле джерела даних');

  CurrentKey := FUniqueField.Value;;

  DataSource.DataSet.DisableControls;

  if not IsSelected(CurrentKey) then
  begin
    DataSource.DataSet.First;
    Mode := mmStart;
    FirstRecord := Unassigned;
    lastRecord := Unassigned;
    while not DataSource.DataSet.Eof and (Mode <> mmStop) do
    begin
      case Mode of
        mmStart:
          begin
            if FUniqueField.Value = CurrentKey then
            begin
              FirstRecord := CurrentKey;
              Mode := mmStartByCurrent;
            end
            else if IsCurrentSelected then
            begin
              Mode := mmStartBySelected;
              FirstRecord := FUniqueField.Value;
            end;
          end;
        mmStartBySelected:
          begin
            if FUniqueField.Value = CurrentKey then
            begin
              lastRecord := CurrentKey;
              Mode := mmStop;
            end
            else if IsCurrentSelected then
              FirstRecord := FUniqueField.Value;
          end;
        mmStartByCurrent:
          begin
            if IsCurrentSelected then
            begin
              lastRecord := FUniqueField.Value;
              Mode := mmStop;
            end;
          end;
      end; { case }
      DataSource.DataSet.Next;
    end;
    if not varIsEmpty(FirstRecord) and not varIsEmpty(lastRecord) then
    begin
      fgExit := False;
      if DataSource.DataSet.Locate(FUniqueField.FieldName, FirstRecord, []) then
        while not DataSource.DataSet.Eof and not fgExit do
        begin
          if not IsCurrentSelected then
          begin
            if Selection.SelectCount >= MaxSelectRows then
              raise Exception.Create
                ('Перевищено максимально допустиму кількість виділених рядків');
            Inc(Selection.SelectCount);
            Selection.Selected[Selection.SelectCount] := FUniqueField.Value;
          end;
          if FUniqueField.Value = lastRecord then
            fgExit := True;
          DataSource.DataSet.Next;
        end;
    end;
    DataSource.DataSet.Locate(FUniqueField.FieldName, CurrentKey, []);
  end;
  DataSource.DataSet.EnableControls;
end;

procedure TDBSearchGrid.GetCellProps(Field: TField; AFont: TFont;
  var Background: TColor; Highlight: Boolean);
begin
  {$IFDEF USE_RX_GRID}
  inherited GetCellProps(Field, AFont, Background, Highlight);
  {$ENDIF}
  if FGroupSelectEnable and not Highlight then
    if IsCurrentSelected then
      Background := $00FFA0A0;
end;

function TDBSearchGrid.GetSelectedKey(Index: Integer): variant;
begin
  Result := Selection.Selected[Index + 1];
end;

end.

