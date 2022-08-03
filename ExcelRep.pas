unit ExcelRep;
{$I ExcelRep.INC}

interface

{$DEFINE VER100}
{$R excelrep.res}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, DB,
{$IFDEF IBDIRECT} IbQuery, IbDatabase {$ELSE} dbTables{$ENDIF};

type
  TCustomExcelReport = class(TComponent)
  public
    procedure RecurseExec(Excel: variant); virtual; abstract;
  end;

type
  TExcelDoAfterReady = (edNone, edPrint, edShow);

type
  TExcelReport = class;

  (* TColumnSelectEvent= procedure (Sender:TWordReport; Word:Variant;Column:integer) of object; *)
  TFieldGetTextEvent = procedure(Sender: TExcelReport; Var Text: string;
    FieldNo: integer) of Object;
  TOnCreateDocumentEvent = procedure(Sender: TExcelReport; var Excel: variant)
    of Object;
{$IFDEF IBDIRECT}
  TOnNextRecordEvent = procedure(Sender: TExcelReport; Query: TibQuery;
    var TableRow: integer; var MoreData: boolean) of object;
{$ELSE}
  TOnNextRecordEvent = procedure(Sender: TExcelReport; Query: TQuery;
    var TableRow: integer; var MoreData: boolean) of object;
{$ENDIF}
  TMainGetPos = procedure(Sender: TExcelReport; FieldNo: integer;
    var Col, Row: integer) of object;
{$IFDEF IBDIRECT}
  TGetProps = procedure(Sender: TExcelReport; Query: TibQuery; FieldNo: integer;
    var AAlign: TAlign; var FontStyle: TFontStyles; var FontSize: integer)
    of object;
{$ELSE}
  TGetProps = procedure(Sender: TExcelReport; Query: TQuery; FieldNo: integer;
    var AAlign: TAlign; var FontStyle: TFontStyles; var FontSize: integer)
    of object;
{$ENDIF}

  TExcelReport = class(TCustomExcelReport)
  private
    { Private declarations }
    FNumering: boolean;
    FTableSQL: TStrings;
    FMainSQL: TStrings;
{$IFDEF IBDIRECT}
    FBase: TIBBase;
{$ELSE}
    FDataBaseName: string;
{$ENDIF}
    FTemplateName: string;
    FParams: TParams;
    FOnMainGetText: TFieldGetTextEvent;
    FOnTableGetText: TFieldGetTextEvent;
    FOnCreateDocument: TOnCreateDocumentEvent;
    FOnCompleteDocument: TOnCreateDocumentEvent;
    FMainFieldCount: integer;
    FTableFieldCount: integer;
    FCascade: TCustomExcelReport;
    FOnNextRecord: TOnNextRecordEvent;
    FExcel: variant;
    FDoAfterReady: TExcelDoAfterReady;
    FTableBookmark: string;
    FStartCount: integer;
    FSelectBookmarkPrefix: string;
    FFirstTableColumn: integer;
    FFirstTableRow: integer;
    FMainGetPos: TMainGetPos;
    FMainGetProps: TGetProps;
    FTableGetProps: TGetProps;
    FBookmarksCount: integer;
    procedure SetTableSQL(NewSQL: TStrings);
    procedure SetMainSQL(NewSQL: TStrings);
{$IFDEF IBDIRECT}
    procedure SetDatabase(NewDatabase: TIbDataBase);
    function GetDatabase: TIbDataBase;
{$ELSE}
    procedure SetDataBaseName(NewDataBaseName: string);
{$ENDIF}
  protected
    { Protected declarations }
    procedure DoIt;
    Procedure DocumentComplete; virtual;
    procedure TableGetText(Var Text: string; FieldNo: integer); virtual;
  public
    { Public declarations }
{$IFDEF IBDIRECT}
    procedure SetParams(SQL: TibQuery);
{$ELSE}
    procedure SetParams(SQL: TQuery);
{$ENDIF}
    procedure CreateParams;
    procedure SetCell(Col, Row: integer; Text: string; Align: TAlign;
      FontStyle: TFontStyles; FontSize: integer);
    procedure Exec;
    { Запуск SQL запросів і створення текста (ГОЛОВНА ПРОЦЕДУРА) }
    procedure RecurseExec(Excel: variant); override;
    constructor Create(AOwner: TComponent); override;
    Destructor Destroy; override;
    property Excel: variant read FExcel;
  published
    { Published declarations }

    { Флажок необхідності нумерації рядків. Якщо TRUE то перший стовпець - номер по порядку }
    property Numering: boolean read FNumering write FNumering default false;

    { Назва файла з шаблоном. Це повинен бути файл-шаблон (*.XLS) Може бути вказано шлях,
      у противному випадку файл береться з каталога, звідки було запущено програму }
    property TemplateName: string read FTemplateName write FTemplateName;

    { Бази даних. Теж саме що і "DatabaseName" в компонентів TTable,TQuery }
{$IFDEF IBDIRECT}
    property DataBase: TIbDataBase read GetDatabase write SetDatabase;
{$ELSE}
    property DatabaseName: string read FDataBaseName write SetDataBaseName;
{$ENDIF}
    { Текст SQL запроса для формування таблиці. Якщо порожній - таблична частина не формується }
    property TableSQL: TStrings read FTableSQL write SetTableSQL;

    { Текст SQL запроса для формування додаткових полів. Імена полів в
      результаті цього запроса повинні співпадати з іменами закладок в документі.
      Зайві поля, неіснуючі закладки - ігноруються.Якщо назва поля закінчується на
      символ "_" (підкреслення) то значення поля конвертується в текстовий вигляд як сума
      (1234.55 - "одна тисяча двісті тридцять пять гривень 55 коп.") }
    property MainSQL: TStrings read FMainSQL write SetMainSQL;

    { Параметри для запросу. Див. Params компонентів TQuery,TStoredProc. Включає в
      себе параметри обох запросів TableSQL i MainSQL }
    Property Params: TParams read FParams write FParams;

    property FirstTableColumn: integer read FFirstTableColumn
      write FFirstTableColumn;
    property FirstTableRow: integer read FFirstTableRow write FFirstTableRow;

    { Процедури, визначені користувачем, які дають змогу змінити текст, що вставляється в таблицю.
      Наприклад якщо значення третьго поля може бути 1 або 0 а в таблицю мовинно вставлятися "YES" або "NO"
      то текст процедури буде слідуючим:
      begin
      if FieldNo=3 then
      if Text='1' then Text:='Yes' else
      if Text='0' then Text:='No' else Text:='Error';
      end; }

    Property OnMainGetText: TFieldGetTextEvent read FOnMainGetText
      write FOnMainGetText;
    property OnMainGetPos: TMainGetPos read FMainGetPos write FMainGetPos;
    property OnTableGetText: TFieldGetTextEvent read FOnTableGetText
      write FOnTableGetText;

    { Викликається після створення документа. Може використовуватися для ініціалізації
      ручної вставки даних або додаткового форматування документа. }
    Property OnCreateDocument: TOnCreateDocumentEvent read FOnCreateDocument
      write FOnCreateDocument;
    Property OnCompleteDocument: TOnCreateDocumentEvent read FOnCompleteDocument
      write FOnCompleteDocument;

    { Наступні дві властивості використовуються для визначення кількості полів
      при використанні ручної вставки данних (через MainGetText,TableGetText).
      У випадку не порожніх SQL-запитів ігноруються }
    property MainFieldCount: integer read FMainFieldCount write FMainFieldCount;
    property TableFieldCount: integer read FTableFieldCount
      write FTableFieldCount;

    { Викликається при ручній вставці даних для переходу на наступний запис. Тільки при TableSQL.SQL='' }
    property OnNextRecord: TOnNextRecordEvent read FOnNextRecord write FOnNextRecord;
    
    { Вказівник на такий самий компонент.Використовується для каскадного об'єднання компонентів,
      при необхідності вставити в документ більше одної таблиці.Всі каскадно-підлеглі компоненти використовують
      документ, створений головним компонентом }
    property Cascade: TCustomExcelReport read FCascade write FCascade;
    { Що необхідно робити після завершення формування документу.Можливі варіанти:
      wdNone - нічого, wdPrint - роздрукувати і закрити, wdShow - показати для подальшого редагування }
    property DoAfterReady: TExcelDoAfterReady read FDoAfterReady
      write FDoAfterReady default edShow;
    property StartCount: integer read FStartCount write FStartCount default 1;
    property OnMainGetProps: TGetProps read FMainGetProps write FMainGetProps;
    property OnTableGetProps: TGetProps read FTableGetProps
      write FTableGetProps;
  end;

function ConvertToText(value: double): string;

implementation

uses
  variants, ComObj;

procedure TExcelReport.RecurseExec;
begin
  FExcel := Excel;
  DoIt;
  if Assigned(FCascade) then
    FCascade.RecurseExec(FExcel);
  FExcel := UnAssigned;
end;

procedure TExcelReport.Exec;
var
  e: variant;
  s: string;
begin
  try
    e := GetActiveOLEObject('Excel.Application');
  except
    e := UnAssigned;
  end;
  if VarIsEmpty(e) then
    e := CreateOLEObject('Excel.Application');
  if FTemplateName <> '' then
  begin
    if ExtractFilePath(FTemplateName) = '' then
    begin
      if csDesigning in ComponentState then
        s := getCurrentdir + '/' + FTemplateName
      else
        s := ExtractFilePath(ParamStr(0)) + FTemplateName;
    end
    else
      s := FTemplateName;
    e.Workbooks.Add(s);
  end
  else
    e.Workbooks.Add;
  if Assigned(FOnCreateDocument) then
    FOnCreateDocument(self, e);

  RecurseExec(e);

  case FDoAfterReady of
    edShow:
      e.Visible := true;
    edPrint:
      begin
        e.ActiveWorkBook.PrintOut;
        e.ActiveWorkBook.Close(0);
        e.Quit;
        { W.ActiveDocument.Close(SaveChanges:=2); }
      end;
  end;

  e := UnAssigned;
end;

procedure TExcelReport.CreateParams;
var {$IFDEF IBDIRECT} SQL1: TibQuery; {$ELSE}SQL1: TQuery; {$ENDIF}
  List, List1, List2: TParams;
  TempParam: TParam;
begin
  { if not (csLoading in ComponentState) then }
  begin
    if FMainSQL.Count <> 0 then
    begin
      List1 := TParams.Create;
{$IFDEF IBDIRECT}
      SQL1 := TibQuery.Create(self);
{$ELSE}
      SQL1 := TQuery.Create(self);
{$ENDIF}
      SQL1.SQL.Assign(FMainSQL);
      List1.Assign(SQL1.Params);
      SQL1.free;
    end
    else
      List1 := nil;

    if FTableSQL.Count <> 0 then
    begin
      List2 := TParams.Create;
{$IFDEF IBDIRECT}
      SQL1 := TibQuery.Create(self);
{$ELSE}
      SQL1 := TQuery.Create(self);
{$ENDIF}
      SQL1.SQL.Assign(FTableSQL);
      List2.Assign(SQL1.Params);
      SQL1.free;
    end
    else
      List2 := nil;

    List := TParams.Create;
    if List1 <> nil then
    begin
      List.Assign(List1);
      if List2 <> nil then
      begin
        while List2.Count <> 0 do
        begin
          TempParam := List2.Items[0];
          List2.RemoveParam(TempParam);
          try
            List.ParamByName(TempParam.Name);
          except
            List.AddParam(TempParam);
          end;
        end;
      end;
    end
    else
    begin
      if List2 <> nil then
        List.Assign(List2);
    end;

    if FParams <> nil then
    begin
      List.AssignValues(FParams);
      FParams.free;
    end;
    FParams := List;
    if List1 <> nil then
      List1.free;
    if List2 <> nil then
      List2.free;
  end { else FParams.Clear; }
end;

destructor TExcelReport.Destroy;
begin
  if FTableSQL <> nil then
  begin
    FTableSQL.free;
    FTableSQL := nil;
  end;
  if FMainSQL <> nil then
  begin
    FMainSQL.free;
    FMainSQL := nil;
  end;
  if FParams <> nil then
  begin
    FParams.free;
    FParams := nil;
  end;
{$IFDEF IBDIRECT}
  IF FBase <> nil then
  begin
    FBase.free;
    FBase := nil;
  end;
{$ENDIF}
  inherited;
end;
{$IFDEF IBDIRECT}

procedure TExcelReport.SetDatabase;
begin
  FBase.DataBase := NewDatabase;
end;

function TExcelReport.GetDatabase;
begin
  Result := FBase.DataBase;
end;
{$ELSE}

procedure TExcelReport.SetDataBaseName;
begin
  FDataBaseName := NewDataBaseName;
end;
{$ENDIF}

procedure TExcelReport.SetMainSQL;
begin
  FMainSQL.Assign(NewSQL);
  CreateParams;
end;

procedure TExcelReport.SetTableSQL;
begin
  FTableSQL.Assign(NewSQL);
  CreateParams;
end;

constructor TExcelReport.Create;
begin
  inherited Create(AOwner);
  { FColumnsCount:=1; }
  FStartCount := 1;
  FTableSQL := TStringList.Create;
  FMainSQL := TStringList.Create;
  FParams := TParams.Create;
  FCascade := nil;
  FDoAfterReady := edShow;
{$IFDEF IBDIRECT}
  FBase := TIBBase.Create(self);
{$ENDIF}
end;

procedure TExcelReport.DoIt;
var // v:variant;
  s3, s2, s: string;
  TableColumn, TableRow, index: integer;
  FieldCount: integer;
  { ColumnC:integer; }
  RecCount: integer;
{$IFDEF IBDIRECT}
  MainQuery, TableQuery: TibQuery;
{$ELSE}
  MainQuery, TableQuery: TQuery;
{$ENDIF}
  fgDataExist: boolean;
  FontStyle: TFontStyles;
  FontSize: integer;
  AAlign: TAlign;
begin
  // v:=FExcel.Application.WordBasic; {Need for fastest call WordBasic}

  { --------- Таблична частина ------- }
  if (FTableSQL.Count <> 0) or
    ((FTableFieldCount <> 0) and Assigned(FOnTableGetText) and
    Assigned(FOnNextRecord)) then
  begin
    TableQuery := nil;
    try
      if FTableSQL.Count <> 0 then
      begin
{$IFDEF IBDIRECT}
        TableQuery := TibQuery.Create(self);
        TableQuery.SQL.Assign(FTableSQL);
        TableQuery.DataBase := DataBase;
        TableQuery.Transaction := DataBase.DefaultTransaction;
{$ELSE}
        TableQuery := TQuery.Create(self);
        TableQuery.SQL.Assign(FTableSQL);
        TableQuery.DatabaseName := FDataBaseName;
{$ENDIF}
        SetParams(TableQuery);
        TableQuery.Open;
        if FTableFieldCount = 0 then
          FTableFieldCount := TableQuery.FieldCount;
        TableQuery.First;
      end
      else
      begin
        TableQuery := nil;
      end;
      FieldCount := FTableFieldCount;
      s := '';
      RecCount := FStartCount;

      TableRow := FirstTableRow - 1;

      { Формування таблиці }
      repeat
        TableColumn := FirstTableColumn;
        inc(TableRow);

        if Assigned(TableQuery) then
          fgDataExist := not TableQuery.EOF;
        if Assigned(FOnNextRecord) then
          FOnNextRecord(self, TableQuery, TableRow, fgDataExist);
        if fgDataExist then
        begin

          If FNumering then
          begin
            str(RecCount, s3);
            SetCell(TableColumn, TableRow, s3, alRight, [], 0);
            inc(TableColumn);
          end;

          for index := 1 to FieldCount do
          begin
            if TableQuery <> nil then
            begin
              if (TableQuery.Fields[index - 1].IsNull) then
                s2 := ''
              else if (TableQuery.Fields[index - 1] is TFloatField) then
              begin
                if Frac(TableQuery.Fields[index - 1].asFloat * 100) < 0.0001
                then
                begin
                  s2 := FloatToStrF(TableQuery.Fields[index - 1].asFloat,
                    ffNumber, 15, 2);
                  // if pos(',',s2)<>0 then s2[pos(',',s2)]:='.';
                end
                else
                  s2 := TableQuery.Fields[index - 1].asString;
              end
              else
                s2 := Trim(TableQuery.Fields[index - 1].asString);
            end;
            TableGetText(s2, index - 1);
            FontStyle := [];
            FontSize := 0;
            AAlign := alNone;
            if Assigned(FTableGetProps) then
              FTableGetProps(self, TableQuery, index - 1, AAlign, FontStyle,
                FontSize);
            SetCell(TableColumn, TableRow, s2, AAlign, FontStyle, FontSize);
            inc(TableColumn);
          end; { FOR }
          if TableQuery <> nil then
            TableQuery.Next;
          inc(RecCount);
        end; { if fgDataExist }
      until not fgDataExist; { while }
    finally
      if TableQuery <> nil then
        TableQuery.free;
    end;
  end; { IF FTableDS<>nil }

  { ------------- Шапка ------------- }
  if (FMainSQL.Count <> 0) or
    ((FMainFieldCount <> 0) and Assigned(FOnMainGetText)) then
  begin
    MainQuery := nil;
    try
      if FMainSQL.Count <> 0 then
      begin
{$IFDEF IBDIRECT}
        MainQuery := TibQuery.Create(self);
        MainQuery.SQL.Assign(FMainSQL);
        MainQuery.DataBase := DataBase;
        MainQuery.Transaction := DataBase.DefaultTransaction;
{$ELSE}
        MainQuery := TQuery.Create(self);
        MainQuery.SQL.Assign(FMainSQL);
        MainQuery.DatabaseName := FDataBaseName;
{$ENDIF}
        SetParams(MainQuery);
        MainQuery.Open;
        if (FMainFieldCount = 0) then
          FMainFieldCount := MainQuery.FieldCount;
      end
      else
        MainQuery := nil;

      FieldCount := FMainFieldCount;

      for index := 1 to FieldCount do
      begin
        try
          if MainQuery <> nil then
          begin
            s2 := MainQuery.Fields[index - 1].FieldName;
            if MainQuery.Fields[index - 1] is TFloatField then
            begin
              if Frac(MainQuery.Fields[index - 1].asFloat * 100) < 0.0001 then
                s := FloatToStrF(MainQuery.Fields[index - 1].asFloat,
                  ffNumber, 15, 2)
              else
                s := MainQuery.Fields[index - 1].asString;
            end
            else
              s := Trim(MainQuery.Fields[index - 1].asString);
          end
          else
          begin
            s2 := '';
            s := '';
          end;
          if (Length(s2) > 0) and (s2[Length(s2)] = '_') then
            s := ConvertToText(MainQuery.Fields[Index - 1].asFloat);
          TableColumn := index;
          TableRow := 1;
          if Assigned(FOnMainGetText) then
            FOnMainGetText(self, s, index - 1);
          if Assigned(FMainGetPos) then
            FMainGetPos(self, index - 1, TableColumn, TableRow);
          FontStyle := [];
          FontSize := 0;
          AAlign := alNone;
          if Assigned(FMainGetProps) then
            FMainGetProps(self, MainQuery, index - 1, AAlign, FontStyle,
              FontSize);
          SetCell(TableColumn, TableRow, s, AAlign, FontStyle, FontSize);

        except
          { Отсуствующие закладки игнорируєм }
          s := '';
        end;
      end;
    finally
      if MainQuery <> nil then
        MainQuery.free;
    end;
  end;
  DocumentComplete;
end;

procedure TExcelReport.DocumentComplete;
begin
  if Assigned(FOnCompleteDocument) then
    FOnCompleteDocument(self, FExcel);
end;

procedure TExcelReport.TableGetText;
begin
  if Assigned(FOnTableGetText) then
    FOnTableGetText(self, Text, FieldNo);
end;

procedure TExcelReport.SetParams;
begin
  if SQL <> nil then
    SQL.Params.AssignValues(FParams);
end;

procedure TExcelReport.SetCell(Col, Row: integer; Text: string; Align: TAlign;
  FontStyle: TFontStyles; FontSize: integer);
begin
  case Align of
    alNone:
      Excel.Cells[Row, Col].HorizontalAlignment := 1;
    alLeft:
      Excel.Cells[Row, Col].HorizontalAlignment := 2;
    alClient:
      begin
        Excel.Cells[Row, Col].HorizontalAlignment := 3;
        Excel.Cells[Row, Col].VerticalAlignment := 2;
      end;
    alRight:
      Excel.Cells[Row, Col].HorizontalAlignment := 4;
    alTop:
      Excel.Cells[Row, Col].VerticalAlignment := 1;
    alBottom:
      Excel.Cells[Row, Col].VerticalAlignment := 3;
  end;
  if FontSize <> 0 then
    Excel.Cells[Row, Col].Font.Size := FontSize;
  if FontStyle <> [] then
  begin
    if fsBold in FontStyle then
      Excel.Cells[Row, Col].Font.Bold := 1;
    if fsItalic in FontStyle then
      Excel.Cells[Row, Col].Font.Italic := 1;
    if fsUnderline in FontStyle then
      Excel.Cells[Row, Col].Font.Underline := 2;
    if fsStrikeout in FontStyle then
      Excel.Cells[Row, Col].Font.Strikethrough := 1;
  end;
  Excel.Cells[Row, Col].NumberFormat := '@';
  Excel.Cells[Row, Col].value := Text;
end;

{ Конвертація числа в текст }
function ConvertToText(value: double): string;
var
  s: string;
  s2: string[5];
  d: array [1 .. 15] of byte;
  NR, j, i: byte;
  r, v: double;

const
  rozr3: array [1 .. 3, 1 .. 4] of string[10]
    = (('гривень ', 'тисяч ', 'мільйонів ', 'мільярдів '),
    ('гривня ', 'тисяча  ', 'мільйон ', 'мільярд '),
    ('гривні ', 'тисячі ', 'мільйони ', 'мільярди '));
  odin: array [0 .. 9] of string[8] = ('', 'один ', 'два  ', 'три ', 'чотири ',
    'п''ять ', 'шість ', 'сім ', 'вісім ', 'дев''ять ');
  odin2: array [0 .. 9] of string[8] = ('', 'одна  ', 'дві ', 'три ', 'чотири ',
    'п''ять ', 'шість ', 'сім ', 'вісім ', 'дев''ять ');
  des: array [0 .. 9] of string[11] = ('', '', 'двадцять ', 'тридцять ',
    'сорок ', 'п''ятдесят ', 'шістдесят ', 'сімдесят ', 'вісімдесят ',
    'девяносто ');
  sotni: array [0 .. 9] of string[10] = ('', 'сто ', 'двісті ', 'тристо ',
    'чотиристо ', 'п''ятсот ', 'шістсот ', 'сімсот ', 'вісімсот ',
    'дев''ятсот ');
  nadts: array [0 .. 9] of string[15] = ('десять ', 'одинадцять ',
    'дванадцять ', 'триналцять ', 'чотирнадцять ', 'п''ятнадцять ',
    'шістнадцять ', 'сімнадцять ', 'вісімнадцять ', 'дев''ятнадцять ');
const
  eps = 0.001;
begin
  s := '';
  v := value;
  r := 1E+14;
  for i := 15 downto 1 do
  begin
    d[i] := round(int(v / r) + eps);
    v := v - r * d[i];
    r := int(r / 10);
  end;
  v := round(v * 100);
  i := 15;
  while (i >= 1) and (d[i] = 0) do
    dec(i);
  while i > 0 do
  begin
    j := (i - 1) mod 3;
    NR := 0;
    case j of
      0:
        begin
          case d[i] of
            0, 5, 6, 7, 8, 9:
              NR := 1; { вЁбпз,¬i«м©®­iў,ЈpЁўҐ­м }
            1:
              NR := 2; { вЁбпз ,¬i«м©®­,ЈpЁў­п }
            2, 3, 4:
              NR := 3; { вЁбпзi, ¬i«м©®­Ё, ЈpЁў­i }
          end;
          if (i - 1) div 3 > 1 then
            s := s + odin[d[i]]
          else
            s := s + odin2[d[i]];
          if (d[i] <> 0) or (d[i + 1] <> 0) or (d[i + 2] <> 0) or (i = 1) then
            s := s + rozr3[NR, (i - 1) div 3 + 1];
        end;
      1:
        if d[i] = 1 then
        begin
          s := s + nadts[d[i - 1]] + rozr3[1, (i - 2) div 3 + 1];
          dec(i);
        end
        else
          s := s + des[d[i]];
      2:
        s := s + sotni[d[i]];
    end;
    dec(i);
  end;
  str(v: - 2: 0, s2);
  IF Length(s2) = 1 then
    s2 := '0' + s2;
  s := s + s2 + ' копійок.';
  ConvertToText := s;
end;

end.
