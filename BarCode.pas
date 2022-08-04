unit BarCode;
{$INCLUDE SWITCHES.INC}

interface
uses windows,messages,SysUtils, Classes, Controls,StdCtrls,ExtCtrls;

var BCEnable:boolean=true;
var MinBCLength:integer=5;

const
  WM_BARCODE=WM_USER+133;
{$IFDEF COM_BC_SCAN}
type TCOMBarThread=class(TThread)
         PortHandle:THandle;
      public
        procedure Execute;override;
        constructor Create;
        destructor Destroy;override;
     end;
var  ComBarThread:TComBarThread;
{$ENDIF}


{$IFDEF KBD_BC_PREFIX}
type TKbdBarPrefix=class(TComponent)
     public
        FHook:HHook;
        constructor Create(AOwner:TComponent);override;
        destructor Destroy;override;
     end;
var KbdScanPrefix:TKbdBarPrefix;
{$ENDIF}

{$IFDEF KBD_BC_TIME}
type TKbdBarTime=class(TComponent)
         Timer:TTimer;
         procedure DoTimer(Sender: TObject);
     public
        FHook:HHook;
        constructor Create(AOwner:TComponent);override;
        destructor Destroy;override;
     end;
var  KbdScanTime:TKbdBarTime;
{$ENDIF}

{$IFDEF BC_ALL_SYMBOLS}
const BCChars:set of char=['0','1'..'9','a'..'z','A'..'Z'];
{$ELSE}
const BCChars:set of char=['0','1'..'9'];
{$ENDIF}

implementation
uses forms, ConstFrm, Dialogs;



{ -------------------- Сканер штрих-кодів з підключенням до COM-порта ---------------------}
{$IFDEF COM_BC_SCAN}

procedure TCOMBarThread.Execute;
var Readed:DWORD;
    BarBuffer,Buffer:array [1..10240] of char;
    PartialReaded:string;
    i:integer;
begin
  PartialReaded:='';
  repeat
    sleep(100);
    if PortHandle>32 then
    begin
      ReadFile(PortHandle,Buffer,SizeOf(Buffer),Readed,nil);
      for i:=1 to Readed do
      begin
        if Buffer[i] in [#13,#10,#0] then
        begin
          if (Length(PartialReaded)>0) then
          begin
             if Application.Active  and Assigned(Screen.ActiveCustomForm) then
             begin
                if MainConsts.BarCode.fgUTF8Decode then PartialReaded:=UTF8ToAnsi(PartialReaded);
                StrPCopy(@BarBuffer,PartialReaded);
                SendMessage(Screen.ActiveCustomForm.Handle,WM_BARCODE,Integer(@BarBuffer),0);
             end;
          end;
          PartialReaded:='';
        end else PartialReaded:=PartialReaded+Buffer[i];
      end;

    end else DoTerminate;
  until Terminated;
end;

destructor TCOMBarThread.Destroy;
begin
    if PortHandle>32 then
    begin
      CloseHandle(PortHandle);
      PortHandle:=0;
    end;
    inherited Destroy;
    ComBarThread:=nil;
end;

constructor TCOMBarThread.Create;
var
    PN:string;
    DCB:_DCB;
    commtimeouts:_COMMTIMEOUTS;
begin
  inherited Create(false);
  if Assigned(ComBarThread)
     then raise exception.Create('Вже існує екземпляр драйвера сканера штрих-кодів');

  try
      PN:='\\.\COM'+intToStr(MainConsts.BarCode.BarPortNumber);
      PortHandle:=CreateFile(PChar(PN),Generic_read,0,nil,OPEN_EXISTING,0,0);
      if PortHandle<32 then
          raise Exception.Create('Cannot open port');
      EscapeCommFunction(PortHandle,SETDTR);
      EscapeCommFunction(PortHandle,SETRTS);

      if not GetCommState(PortHandle,DCB) then
         raise Exception.Create('Cannot get ComState');

      case MainConsts.BarCode.PortSpeed of
         1200: DCB.baudRate:=CBR_1200;
         2400: DCB.baudRate:=CBR_2400;
         4800: DCB.baudRate:=CBR_4800;
         9600: DCB.baudRate:=CBR_9600;
        14400: DCB.baudRate:=CBR_14400;
        19200: DCB.baudRate:=CBR_19200;
        38400: DCB.baudRate:=CBR_38400;
        57600: DCB.baudRate:=CBR_57600;
        else raise Exception.Create('Bad baudrate');
      end;

      case MainConsts.BarCode.PortParity of
        0: DCB.Parity:=EVENPARITY;
        1: DCB.Parity:=ODDPARITY;
        2: DCB.Parity:=NOPARITY;
        3: DCB.Parity:=MARKPARITY;
        else raise Exception.Create('Bad parity');
      end;

      DCB.ByteSize:=MainConsts.BarCode.PortByteSize;
      case MainConsts.BarCode.PortStopBits of
         0: DCB.StopBits:=ONESTOPBIT;
         2: DCB.StopBits:=TWOSTOPBITS;
         1: DCB.StopBits:=ONE5STOPBITS;
         else Raise Exception.Create('Bad stop bits value.It must be 1,2 or 3(1.5)');
      end;
      DCB.DCBlength:=SizeOf(DCB);
      DCB.Flags:=0;
      DCB.XonLim:=0;
      DCB.XoffLim:=0;
      DCB.XonChar:=#0;
      IF not SetCommState(PortHandle, DCB) then
         raise Exception.Create('Cannot Set ComState');
      if not GetCommTimeOuts(PortHandle,COMMTIMEOUTS) then
         raise Exception.Create('Cannot get COMMTIMEOUTS');
      CommTimeOuts.ReadIntervalTimeout:=200;
      CommTimeOuts.ReadTotalTimeoutMultiplier:=0;
      CommTimeOuts.ReadTotalTimeoutConstant:=20;
      if not SetCommTimeOuts(PortHandle,COMMTIMEOUTS) then
         raise Exception.Create('Cannot set COMMTIMEOUTS');
   except
       on e : exception do
       begin
           if PortHandle>32 then
           begin
              CloseHandle(PortHandle);
              PortHandle:=0;
           end;
           MessageUA(e.Message,mtError,[mbCancel],0);
           Terminate;
       end;
   end;
end;
{$ENDIF}

{$IFDEF KBD_BC_PREFIX}

 {Сканер штрих-кодів з підключенням через клавіатуру, посилає префікси/суфікси}
{Сканер повинен посилати коди так #$1F,#$20-ШТРИХ-КОД-#$20 без термінатора CR або LF}

const FirstPrefixByte=$1F;
      SecondPrefixByte=$20;
      SuffixByte=$20;

var fgScanState:integer=0;
    ScannedString:string='';

function KbdProcPrefix(code:integer;wParam:DWORD;lParam:DWORD):lResult;far;stdcall;
var
    BarBuffer:array [1..2048] of char;
begin
    {$IFDEF DEBUG}
    LOG('code='+IntToStr(Code)+Format(',wParam=%x,lParam=%x',[wParam,lParam]));
    {$ENDIF}
    Result:=0;
    if (code<0) or not BCEnable then result:=CallNextHookEx(KbdScanPrefix.FHook,code,wParam,lParam) else
    begin
       if (lParam and KF_UP)=0 then
       begin
          Result:=1;
          case fgScanState of
          0: if wParam=FirstPrefixByte then fgScanState:=1 else Result:=0;
          1: if wParam=SecondPrefixByte then begin fgScanState:=2;ScannedString:='';end
             else
               begin
                 fgScanState:=0;
                 result:=0;
               end;
          2: if wParam=SuffixByte then
             begin
               fgScanState:=0;
               if ScannedString<>'' then
               begin
                 StrPCopy(@BarBuffer,ScannedString);
                 if Application.Active  and Assigned(Screen.ActiveCustomForm) then
                    SendMessage(Screen.ActiveCustomForm.Handle,WM_BARCODE,Integer(@BarBuffer),0);
               end;
             end else ScannedString:=ScannedString+char(wParam);
          end
       end else
       if fgScanState=0 then Result:=0 else result:=1;
    end;
end;

constructor TKbdBarPrefix.Create;
begin
    inherited ;
    if Assigned(KbdScanPrefix) {$IFDEF KBD_BC_TIME} or Assigned(KbdScanTime) {$ENDIF}
        then raise exception.Create('Вже існує екземпляр драйвера сканера штрих-кодів');
    KbdScanPrefix:=Self;
    FHook:=SetWindowsHookEx(WH_KEYBOARD,@KbdProcPrefix,0,GetCurrentThreadID);
end;

destructor TKbdBarPrefix.Destroy;
begin
   UnhookWindowsHookEx(FHook);
   inherited;
   KbdScanPrefix:=nil;
end;
{$ENDIF}


{$IFDEF KBD_BC_TIME}
 {Сканер штрих-кодів з підключенням через клавіатуру, без префіксів/суфіксів,
 розпізнавання через швидкість надходження цифр}

var   Scanned:array [1..1024] of record wParam,lParam:dword;end;
      ScannedCount:integer;
      TimeCount:integer;

const BC_TimeOut=20; {Тайм-аут для сканера штрих-кодів в мс.}

function BreakSequence:boolean;
var i:integer;
    BarBuffer:array [1..2048] of char;
    ScannedString:string;
    Msg:TagMsg;
    hwnd:THandle;
begin
    Result:=false;
    IF ScannedCount=0 then exit;
    if ScannedCount>=MinBCLength then
    begin
       ScannedString:='';
       for i:=1 to ScannedCount do
          ScannedString:=ScannedString+Char(Scanned[i].wParam);
       StrPCopy(@BarBuffer,ScannedString);
       ScannedCount:=0;
       if Application.Active  and Assigned(Screen.ActiveCustomForm) then
         SendMessage(Screen.ActiveCustomForm.Handle,WM_BARCODE,Integer(@BarBuffer),0);
       Result:=true;
    end else
    for i:=1 to ScannedCount do
    begin
       if Application.Active and Assigned(Screen.ActiveControl) then
          {
          PostMessage(Screen.ActiveControl.Handle, WM_CHAR, Scanned[i].wParam,Scanned[i].lParam);
          }
          PostMessage(Screen.ActiveControl.Handle, WM_KEYDOWN, Scanned[i].wParam,Scanned[i].lParam);

    end;
    ScannedCount:=0;
    TimeCount:=0;
end;

procedure SaveChar(WPar,LPar:DWORD);
begin
   inc(ScannedCount);
   with Scanned[ScannedCount] do
   begin
     lParam:=LPar;
     wParam:=WPar;
   end;
   TimeCount:=0;
end;

function KbdProcTime(code:integer;wParam:DWORD;lParam:DWORD):lResult;far;stdcall;
   {1-Prevent next call, 0- allow next call & target procedure}
var
    ch:char;
    fgDown,Res,fgAlt:boolean;
begin
    Result:=0;
    if (code<0) or (not BCEnable) then
        result:=CallNextHookEx(KbdScanTime.FHook,code,wParam,lParam) else
    begin
       fgDown:=((lParam shr 16) and KF_UP)=0;
       fgAlt:=((lParam shr 16) and KF_ALTDOWN)<>0;
       ch:=char(wParam);
       if not (ch in BCChars) or fgAlt then
       begin
            if fgDown then
            begin
                Res:=BreakSequence;
                if (ch=#13) and Res then Result:=1;
            end
       end  else
       begin
            if fgDown then
            begin
               SaveChar(wParam,lParam);
               Result:=1;
            end;
       end;
    end;
end;

procedure TKbdBarTime.DoTimer;
begin
   inc(TimeCount);
   if TimeCount>2 then
   begin
     BreakSequence;
   end;
end;



constructor TKbdBarTime.Create;
begin
    inherited ;
    if Assigned(KbdScanTime) {$IFDEF KBD_BC_PREFIX}or Assigned(KbdScanPrefix){$ENDIF}
      then raise exception.Create('Вже існує екземпляр драйвера сканера штрих-кодів');
    KbdScanTime:=Self;
    FHook:=SetWindowsHookEx(WH_KEYBOARD,@KbdProcTime,0,GetCurrentThreadID);
    Timer:=Ttimer.Create(Self);
    Timer.Interval:=BC_timeOut;
    Timer.OnTimer:=DoTimer;
    Timer.Enabled:=true;
end;

destructor TKbdBarTime.Destroy;
begin
   UnhookWindowsHookEx(FHook);
   inherited;
   KbdScanTime:=nil;
end;

{$ENDIF}


end.
