unit RefreshSQL;

interface

uses
  Windows, Messages, SysUtils, Classes,
  IBQuery, DB, IBEvents, Contnrs, IBDatabase;

type
  TIBEventRefresh = class;

  TRefreshedQuery = class(TIBQuery)
  private
    { Private declarations }
    fgInHandle: boolean;
    FSortedField: TField;
    FUniqueField: TField;
    FEventName: string;
    ValKey, ValSort: Variant;
    RefreshStarted: boolean;
    FBeforeManualRefresh, FOnManualRefresh: TNotifyEvent;
    FEvent: TIBEventRefresh;
  protected
    { Protected declarations }
    procedure SetSortedField(NewField: TField);
    procedure EventHandler;
    procedure SetUniqueField(NewField: TField);
    procedure SetEventName(NewEventName: string);
    procedure DoAfterOpen; override;
    procedure DoBeforeClose; override;
    procedure DoAfterClose; override;
    procedure DoBeforeOpen; override;
    procedure Loaded; override;
    procedure SetEventRefresh(Value: TIBEventRefresh);
    function GetEventRefresh: TIBEventRefresh;
    procedure SetDatabase(Value: TIBDatabase);
    procedure InternalPrepare; override;
  public
    { Public declarations }
    CanRefresh: boolean;
    procedure Refresh; virtual;
    procedure RefreshCurrentRecord; virtual;
    procedure StartRefresh;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    { Published declarations }
    property SortedField: TField read FSortedField write SetSortedField;
    property UniqueField: TField read FUniqueField write SetUniqueField;
    property BeforeManualRefresh: TNotifyEvent read FBeforeManualRefresh
      write FBeforeManualRefresh;
    property OnManualRefresh: TNotifyEvent read FOnManualRefresh
      write FOnManualRefresh;
    property EventName: string read FEventName write SetEventName;
    property EventAlerter: TIBEventRefresh read GetEventRefresh
      write SetEventRefresh;
    property Transaction stored false;
    property Database write SetDatabase;
  end;

  TRefreshItem = class
  public
    EventName: string;
    RefreshedQuery: TRefreshedQuery;
  end;

  TIBEventRefresh = class(TIBEvents)
    FQueryList: TObjectList;
  protected
    procedure DoRefresh(Sender: TObject; EventName: String; EventCount: Integer;
      var CancelAlerts: boolean);
    procedure RefreshEventList;
    procedure SetRefreshRegistered(Value: boolean);
    function GetRefreshRegistered: boolean;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure AddEvent(EventName: string; RefreshObject: TRefreshedQuery);
    procedure DeleteEvent(pos: Integer);
    procedure DeleteEventByObject(RefreshObject: TRefreshedQuery);
    function EventPos(EventName: string): Integer;
    property Events stored false;
    property RefreshRegistered: boolean read GetRefreshRegistered
      write SetRefreshRegistered stored false default false;
  end;

procedure Register;

implementation

uses IBCustomDataSet, forms;

procedure Register;
begin
  RegisterComponents('Samples', [TRefreshedQuery]);
  RegisterComponents('Samples', [TIBEventRefresh]);
end;

{ ------------------------IBEventRefresh------------------------------ }

constructor TIBEventRefresh.Create;
begin
  inherited;
  FQueryList := TObjectList.Create;
  FQueryList.OwnsObjects := true;
  OnEventAlert := DoRefresh;
end;

destructor TIBEventRefresh.Destroy;
begin
  RefreshRegistered := false;
  Events.Clear;
  FreeAndNil(FQueryList);
  if Assigned(Database) and Database.Connected then
    Database.Close;
  Application.ProcessMessages;
  Sleep(100);
  Application.ProcessMessages;
  inherited;
end;

procedure TIBEventRefresh.SetRefreshRegistered;
begin
  if ((Value and Assigned(Database) and Database.Connected) or not Value) and
    (RefreshRegistered <> Value) then
  begin
    Registered := Value;
    Sleep(10);
    Application.ProcessMessages;
  end;
end;

function TIBEventRefresh.GetRefreshRegistered;
begin
  Result := Registered;
end;

procedure TIBEventRefresh.AddEvent(EventName: string;
  RefreshObject: TRefreshedQuery);
var
  i: Integer;
  fgFound: boolean;
  RefreshItem: TRefreshItem;
begin
  i := 0;
  fgFound := false;
  while (not fgFound) and (i < FQueryList.Count) do
  begin
    fgFound := ((FQueryList.Items[i] as TRefreshItem).EventName = EventName) and
      ((FQueryList.Items[i] as TRefreshItem).RefreshedQuery = RefreshObject);
    inc(i);
  end;
  if not fgFound then
  begin
    if RefreshRegistered then
      RefreshRegistered := false;
    RefreshItem := TRefreshItem.Create;
    RefreshItem.EventName := EventName;
    RefreshItem.RefreshedQuery := RefreshObject;
    FQueryList.Add(RefreshItem);
  end;
  RefreshEventList;
end;

procedure TIBEventRefresh.DeleteEvent(pos: Integer);
begin
  if (pos < FQueryList.Count) and (pos >= 0) then
  begin
    if RefreshRegistered then
      RefreshRegistered := false;
    FQueryList.Delete(pos);
    RefreshEventList;
  end;
end;

procedure TIBEventRefresh.DeleteEventByObject(RefreshObject: TRefreshedQuery);
var
  fgNeedRefresh: boolean;
  i: Integer;
begin
  fgNeedRefresh := false;
  i := 0;
  while (i < FQueryList.Count) do
  begin
    if (FQueryList.Items[i] as TRefreshItem).RefreshedQuery = RefreshObject then
    begin
      if RefreshRegistered then
        RefreshRegistered := false;
      FQueryList.Delete(i);
      fgNeedRefresh := true;
    end
    else
      inc(i);
  end;
  if fgNeedRefresh then
    RefreshEventList;
end;

function TIBEventRefresh.EventPos(EventName: string): Integer;
var
  i: Integer;
  fgFound: boolean;
begin
  i := 0;
  fgFound := false;
  while (not fgFound) and (i < FQueryList.Count) do
  begin
    fgFound := (FQueryList.Items[i] as TRefreshItem).EventName = EventName;
    if not fgFound then
      inc(i);
  end;
  if fgFound then
    Result := i
  else
    Result := -1;
end;

procedure TIBEventRefresh.RefreshEventList;
var
  i: Integer;
  s: TStringList;
begin
  s := TStringList.Create;
  try
    for i := 0 to FQueryList.Count - 1 do
    begin
      if s.IndexOf((FQueryList.Items[i] as TRefreshItem).EventName) = -1 then
        s.Add((FQueryList.Items[i] as TRefreshItem).EventName);
    end;
    if s.Text <> Events.Text then
    begin
      RefreshRegistered := false;
      Events.Assign(s);
    end;
  finally
    s.Free;
  end;
  if Events.Text <> '' then
  begin
    if not RefreshRegistered and Assigned(Database) and Database.Connected then
      RefreshRegistered := true;
  end
end;

procedure TIBEventRefresh.DoRefresh(Sender: TObject; EventName: String;
  EventCount: Integer; var CancelAlerts: boolean);
var
  i: Integer;
begin
  for i := 0 to FQueryList.Count - 1 do
  begin
    if (FQueryList.Items[i] as TRefreshItem).EventName = EventName then
    begin
      (FQueryList.Items[i] as TRefreshItem).RefreshedQuery.EventHandler;
    end;
  end;
  CancelAlerts := false;
end;

{ ------------------------RefreshedQuery------------------------------ }
constructor TRefreshedQuery.Create;
begin
  inherited;
  if csDesigning in ComponentState then
  begin
    Transaction := TIBTransaction.Create(Self);
  end
  else
    Transaction := nil;
end;

destructor TRefreshedQuery.Destroy;
var
  FTransaction: TIBTransaction;
begin
  Active := false;
  If Assigned(Transaction) then
  begin
    if Transaction.Active then
      Transaction.Commit;
    FTransaction := Transaction;
    Transaction := nil;
    FreeAndNil(FTransaction);
  end;
  Inherited;
end;

procedure TRefreshedQuery.SetDatabase(Value: TIBDatabase);
var
  FTransaction: TIBTransaction;
begin
  if Assigned(Transaction) then
  begin
    if Transaction.Active then
      Transaction.Commit;
    FTransaction := Transaction;
    Transaction := nil;
    FreeAndNil(FTransaction);
  end;
  inherited Database := Value;
  if Assigned(Database) then
  begin
    Transaction := TIBTransaction.Create(Self);
    Transaction.DefaultDatabase := Database;
    if Database.Connected then
      Transaction.StartTransaction;
  end;
end;

procedure TRefreshedQuery.Loaded;
var
  AActive: boolean;
begin
  AActive := Active;
  Close;
  inherited;
  if csDesigning in ComponentState then
  begin
    if not Assigned(Transaction) then
      Transaction := TIBTransaction.Create(Self);
    Transaction.DefaultDatabase := Database;
    if (Assigned(Database)) and (Database.Connected) and (not Transaction.Active)
    then
      Transaction.StartTransaction;
  end;
  Active := AActive;
end;

procedure TRefreshedQuery.StartRefresh;
begin
  if Assigned(FBeforeManualRefresh) then
  begin
    FBeforeManualRefresh(Self);
    RefreshStarted := true;
  end
  else
  begin
    if Assigned(FUniqueField) then
      ValKey := FUniqueField.Value;
    if Assigned(FSortedField) then
      ValSort := FSortedField.Value;
    RefreshStarted := true;
  end;
end;

procedure TRefreshedQuery.Refresh;
begin
  DisableControls;
  if not Assigned(FOnManualRefresh) and Assigned(FUniqueField) and
    (FUniqueField.DataType = ftInteger) and (FUniqueField.asInteger = 0) then
  begin
    fgInHandle := true;
    Close;
    Open;
    fgInHandle := false;
    Last;
    CanRefresh := true;
    EnableControls;
    RefreshStarted := false;
    Exit;
  end;
  if not RefreshStarted then
    StartRefresh;
  fgInHandle := true;
  Close;
  Open;
  fgInHandle := false;
  CanRefresh := true;
  if Assigned(FOnManualRefresh) then
    FOnManualRefresh(Self)
  else if Assigned(FUniqueField) then
  begin
    if not Locate(FUniqueField.FieldName, ValKey, []) then
    begin
      First;
      if Assigned(FSortedField) then
        while not Eof and (ValSort > FSortedField.Value) do
          Next
      else
        Last;
    end;
  end
  else
  begin
    First;
    if Assigned(FSortedField) then
      while not Eof and (ValSort > FSortedField.Value) do
        Next
    else
      Last
  end;
  RefreshStarted := false;
  EnableControls;
end;

procedure TRefreshedQuery.SetSortedField;
begin
  if not Assigned(NewField) then
    FSortedField := nil
  else if NewField.dataSet <> Self then
    raise Exception.Create('Field must be of this DataSet')
  else if not(NewField.DataType in [ftString, ftSmallint, ftInteger, ftWord,
    ftAutoInc, ftBoolean, ftFloat, ftCurrency, ftBCD, ftDate, ftDateTime,
    ftTime]) then
    raise Exception.Create('Unsupported field data type.')
  else
    FSortedField := NewField;
end;

procedure TRefreshedQuery.SetUniqueField;
begin
  if not Assigned(NewField) then
    FUniqueField := nil
  else if NewField.dataSet <> Self then
    raise Exception.Create('Field must be of this DataSet')
  else if not(NewField.DataType in [ftString, ftSmallint, ftInteger, ftWord,
    ftAutoInc, ftBoolean, ftFloat, ftCurrency, ftBCD, ftDate, ftDateTime,
    ftTime]) then
    raise Exception.Create('Unsupported field data type.')
  else
    FUniqueField := NewField;
end;

procedure TRefreshedQuery.SetEventName;
begin
  if FEventName <> NewEventName then
  begin
    if (ComponentState = []) and not fgInHandle then
    begin
      if Assigned(FEvent) then
      begin
        if FEventName <> '' then
        begin
          If FEvent.Events.IndexOf(FEventName) <> -1 then
            FEvent.DeleteEvent(FEvent.Events.IndexOf(FEventName));
        end;
        if NewEventName <> '' then
        begin
          FEvent.AddEvent(NewEventName, Self);
        end;
      end;
    end;
  end;
  FEventName := NewEventName;
end;

procedure TRefreshedQuery.EventHandler;
begin
  if (State = dsBrowse) and
    ((not CachedUpdates) or (UpdateStatus = usUnmodified)) then
    Refresh;
end;

procedure TRefreshedQuery.DoBeforeOpen;
var
  FTransaction: TIBTransaction;
begin
  if not(csDesigning in ComponentState) and not Assigned(Transaction) then
  begin
    FTransaction := TIBTransaction.Create(Self);
    FTransaction.DefaultDatabase := Database;
    Transaction := FTransaction;
    if Assigned(Database) and (Database.Connected) then
      Transaction.StartTransaction;
  end;
  inherited;
end;

procedure TRefreshedQuery.InternalPrepare;
var
  FTransaction: TIBTransaction;
begin
  if not Assigned(Transaction) then
  begin
    FTransaction := TIBTransaction.Create(Self);
    FTransaction.DefaultDatabase := Database;
    Transaction := FTransaction;
    if Assigned(Database) and (Database.Connected) then
      Transaction.StartTransaction;
  end;
  inherited;
end;

procedure TRefreshedQuery.DoAfterClose;
var
  FTransaction: TIBTransaction;
begin
  inherited;
  if not(csDesigning in ComponentState) then
  begin
    if Transaction.Active then
      Transaction.Commit;
    FTransaction := Transaction;
    Transaction := nil;
    FTransaction.Free;
  end;
end;

procedure TRefreshedQuery.DoAfterOpen;
begin
  Inherited;
  if (not fgInHandle) and (not RefreshStarted) and (FEventName <> '') and
    Assigned(FEvent) then
  begin
    FEvent.AddEvent(FEventName, Self);
  end;
end;

procedure TRefreshedQuery.DoBeforeClose;
begin
  if (not fgInHandle) and (not RefreshStarted) and Assigned(FEvent) then
  begin
    FEvent.DeleteEventByObject(Self);
  end;
  Inherited;
end;

procedure TRefreshedQuery.SetEventRefresh;
begin
  if (Assigned(FEvent)) and (FEventName <> '') then
  begin
    If FEvent.Events.IndexOf(FEventName) <> -1 then
      FEvent.DeleteEvent(FEvent.Events.IndexOf(FEventName));
  end;
  FEvent := Value;
  if Assigned(FEvent) and (FEventName <> '') and Active then
  begin
    FEvent.AddEvent(FEventName, Self);
  end;
end;

function TRefreshedQuery.GetEventRefresh: TIBEventRefresh;
begin
  Result := FEvent;
end;

procedure TRefreshedQuery.RefreshCurrentRecord;
begin
  if Assigned(UpdateObject) and (UpdateObject.RefreshSQL.Text <> '') then
    InternalRefreshRow
  else
    Refresh;
end;

end.
