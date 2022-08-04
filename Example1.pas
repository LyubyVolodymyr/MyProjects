
procedure TMainFrm.FormShow(Sender: TObject);
begin
  FCanClose := false;
  if not Assigned(Setups) then
    DoCreateSetups;

  ProgramState := psLoading;

  RESTClient1.BaseURL := Setups.ServerAddress;

  if not DoLogin then
    ShowMessage('Сервер не доступний. Помилка авторизації')
  else
    TThread.CreateAnonymousThread(
      procedure ()
      begin
        ConnectBluetoothPrinter;
        LoadBaseData;
      end
    ).Resume;

end;


procedure TMainFrm.LoadBaseData;

  procedure LoadName;
  begin
    (* Інформація про постачальника *)
    RESTRequest1.Params.Clear;
    // RESTClient1.SetHTTPHeader('TOKEN',AccToken);
    RESTRequest1.AddParameter('TOKEN', AccToken,
      TRESTRequestParameterKind.pkHTTPHEADER);
    RESTRequest1.AddParameter('id', '1', TRESTRequestParameterKind.pkGETorPOST);
    RESTRequest1.Execute;
    if RESTResponse1.StatusCode = 200 then
    begin
      Bakery.BakeryName := RESTResponse1.JSONValue.GetValue<string>('NAME');
      Bakery.Id := RESTResponse1.JSONValue.GetValue<Integer>('ID');
    end
      else ShowError();
  end;

  procedure LoadDestinations;
  var JV : TJSONValue;
      I : Integer;
  begin
    (* Інформація про отримувачів *)
    RESTRequest1.Params.Clear;
    RESTRequest1.AddParameter('id', '2', TRESTRequestParameterKind.pkGETorPOST);
    RESTRequest1.AddParameter('TOKEN', AccToken,
      TRESTRequestParameterKind.pkHTTPHEADER);
    RESTRequest1.Execute;
    if RESTResponse1.StatusCode = 200 then
    begin
        DestinationCount := 0;
        JV := RESTResponse1.JSONValue.FindValue('destinations');
        if Assigned(JV) and (JV is TJSONArray) then
        with (JV as TJSONArray) do
        begin
                DestinationCount := Count;
                SetLength(Destinations, DestinationCount);
                for I := 0 to DestinationCount - 1 do
                begin
                    Destinations[I].Id := JV.A[I].GetValue<Integer>('id');
                    Destinations[I].Name := JV.A[I].GetValue<String>('name');
                    // Destinations[I].IsForeign:=JV.A[I].GetValue<Boolean>('is_foreign');
                end;
        end
    end else ShowError;
  end;

  procedure LoadProducts;
  var   I: Integer;
        JV: TJSONValue;
  begin
    (* Інформація про ассортимент *)
    RESTRequest1.Params.Clear;
    RESTRequest1.AddParameter('id', '3', TRESTRequestParameterKind.pkGETorPOST);
    RESTRequest1.AddParameter('TOKEN', AccToken,
      TRESTRequestParameterKind.pkHTTPHEADER);
    RESTRequest1.Execute;
    if RESTResponse1.StatusCode = 200 then
    begin
        JV := RESTResponse1.JSONValue.FindValue('products');
        if Assigned(JV) and (JV is TJSONArray) then
        with (JV as TJSONArray) do
        begin
          ProductCount := Count;
          SetLength(Products, ProductCount);
          for I := 0 to ProductCount - 1 do
          begin
            Products[I].Id := JV.A[I].GetValue<Integer>('id');
            Products[I].Name := JV.A[I].GetValue<String>('name');
            Products[I].Price := JV.A[I].GetValue<double>('price');
            // Products[I].WholePrice:=JV.A[I].GetValue<Double>('whole_price');
          end;
        end
    end
        else ShowError;
  end;

begin
  try
    RESTRequest1.Resource := 'get_bkr_table.php';
    RESTRequest1.Method := rmGET;

    SetLength(Destinations, 0);
    DestinationCount := 0;
    Bakery.BakeryName := '';
    SetLength(Products, 0);
    ProductCount := 0;
    RESTClient1.HandleRedirects := false;
    RESTRequest1.HandleRedirects := false;

    LoadName;
    LoadDestinations;
    LoadProducts;

    if DestinationCount>0 then LoadDestinationsToList;
    if ProductCount>0 then LoadProductsToList;

    ProgramState := psMainMenu;

  except
    if (RESTResponse1.StatusCode >= 300) and (RESTResponse1.StatusCode < 400)
        then // moved, not authorized
            AccToken := '';
  end;
end;


