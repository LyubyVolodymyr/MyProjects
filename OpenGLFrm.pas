unit OpenGLFrm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, BorderInterface, dglOpenGL, DrawInterface,
  Vcl.Menus, ControlThreads, FountainStageUnit;

type
  TOpenGLThread = class(TControlThread)
  private
    FText: string;
    procedure Render;
    procedure SetupGL;
    procedure BuildFont;
    procedure KillFont;
    procedure glPrint(const s: string);
  protected
    dc: HDC;
    base: GLuint;
    Font: HFONT;
    hrc: HGLRC;
    Angle1, Angle2: Double;
    StartMouseX, StartMouseY: integer;
    StartA1, StartA2, StartA3: Double;
    PosViewX, PosViewY, PosViewZ: Double;
    fgDrag: boolean;
    DrawerOGL: TDrawerOpenGL;
    Width, Height: integer;
    PredTimeFPS, FpsCounter: Cardinal;
    Showed: boolean;
  public
    procedure MainThreadProc; override;
    procedure Execute; override;
  end;

type
  TOpenGLForm = class(TForm)
    PopupMenu1: TPopupMenu;
    DayMode1: TMenuItem;
    Stayontop1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure FormMouseMove(Sender: TObject; Shift: TShiftState; X, Y: integer);
    procedure FormDestroy(Sender: TObject);
    procedure FormMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: integer);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormHide(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure DayMode1Click(Sender: TObject);
    function  GetShowed:boolean;
    procedure Stayontop1Click(Sender: TObject);
  private
    { Private declarations }
    RenderThread: TOpenGLThread;
    procedure SaveLayoutsThisForm;
  public
    { Public declarations }
    fgDayMode: boolean;
    property Showed:boolean read GetShowed;

  end;

var
  OpenGLForm: TOpenGLForm;

implementation

uses Bulbes, NozzleInterface, LayoutIntf, System.DateUtils, NozzleStates,
  mmsystem, FountainConsts;

{$R *.dfm}

var
  l_model: DArr = (0.5, 0.5, 0.5, 1);

var
  Coord: DArr = (0, 0, 30, 0);

const
  NearClipping = 0; // Ѕлижн€€ плоскость отсечени€
  FarClipping = 200; // ƒальн€€ плоскость отсечени€
  delta: Double = 0.1;
  BOUBLE_COUNT = 100;
  // JETS_COUNT=20;

var
  pos: DArr = (4, 1.5, -0.1, 0.1);
  dir: DArr = (0, 1, 1, 1);
  Black: DArr = (0, 0, 0, 1);
  amb: DArr = (1, 1, 1, 0.1);

procedure TOpenGLForm.FormCreate(Sender: TObject);
begin
  RenderThread := TOpenGLThread.Create;
  with RenderThread do
  begin

    PosViewX := -0.1;
    PosViewY := -0.7;
    PosViewZ := -5;

    Angle1 := 0;
    Angle2 := -13;

    DrawerOGL := TDrawerOpenGL.Create;
    dc := GetDC(Self.Handle); // получаем контекст устройства по форме Form1
  end;
  RenderThread.Showed := false;
  RenderThread.Start;
end;

procedure TOpenGLThread.Render;
var
  i: integer;
  AtTime: TNozzleTime;
  B: TBorder;
begin
  if not Showed then
    exit;

  glClear(GL_COLOR_BUFFER_BIT or GL_DEPTH_BUFFER_BIT);

  glViewport(0, 0, Width, Height);
  glPrint(FText);

  glMatrixMode(GL_PROJECTION);
  glLoadIdentity;

  // glOrtho(-1,1,-1.5,1,-35,250);
  // gluLookAt( Coord[0],Coord[1],Coord[2], 0,1,0, 0,1,0 );

  gluPerspective(50.0, Self.Width / Self.Height, 0.1, 100);
  // glTranslatef(0.0, 0, -3.0);

  glRotateF(15, 1, 0, 0);
  glRotateF(Angle1, 0, 1, 0);
  glRotateF(Angle2, 1, 0, 0);
  glTranslateF(PosViewX, PosViewY, PosViewZ);
  glMatrixMode(GL_MODELVIEW);
  glLoadIdentity;

  glEnable(GL_LIGHTING);
  glEnable(GL_LIGHT0);
  glLightfv(GL_LIGHT0, GL_DIFFUSE, @amb);
  glLightfv(GL_LIGHT0, GL_POSITION, @pos);
  glLightfv(GL_LIGHT0, GL_SPOT_EXPONENT, @dir);

  // glTranslatef(0,0,-5);
  AtTime := NowNozzleTime;

  for i := 0 to vNozzleStates.Count - 1 do
  begin
    with vNozzleStates[i].OwnerNozzle do
      DrawerOGL.DrawNozzle(X, Y, orient, false);

    vNozzleStates[i].RefreshBulbes(AtTime);

    with vNozzleStates[i] do
      DrawerOGL.DrawBulbes(Bulbes, Color, OpenGLForm.fgDayMode);
  end;

  { Ѕасейн }
  B := FountainStage.Border;
  if Assigned(B) then
    B.Draw(DrawerOGL);
  SwapBuffers(dc);
end;

procedure TOpenGLThread.BuildFont();
begin
  base := glGenLists(96);
  Font := CreateFont(-16, 0, 0, 0, FW_BOLD, 0, 0, 0, ANSI_CHARSET,
    OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, FF_DONTCARE or
    DEFAULT_PITCH, 'Courier NEW');
  SelectObject(dc, Font);
  wglUseFontBitmaps(dc, 32, 96, base);
end;

procedure TOpenGLThread.KillFont();
begin
  glDeleteLists(base, 96);
end;

procedure TOpenGLThread.glPrint(const s: string);
var
  Text: array [0 .. 255] of byte;
  i: integer;

const
  txt_color: DArr = (1.5, 1.5, 1.5, 1);
begin
  if s = '' then
    exit;
  for i := 1 to Length(s) do
    Text[i - 1] := byte(s[i]);
  Text[Length(s)] := 0;
  glMatrixMode(GL_PROJECTION);
  glLoadIdentity;
  gluOrtho2D(0, 0, 0, 0);
  glMatrixMode(GL_MODELVIEW);
  glLoadIdentity;
  glTranslateF(0, 0, -1);

  glMaterialfv(GL_FRONT_AND_BACK, GL_DIFFUSE, @txt_color);
  glPushAttrib(GL_LIST_BIT);
  glRasterPos2f(-1, 0.8);

  glListBase(base - 32);
  glCallLists(Length(s), GL_UNSIGNED_BYTE, @Text);
  glPopAttrib;
end;

procedure TOpenGLThread.SetupGL;
begin
  if not InitOpenGL then
    Application.Terminate;

  // эта строка создаЄт контекст рендеринга
  hrc := CreateRenderingContext(dc, [opDoubleBuffered], 32, 24, 0, 0, 0, 0);


  // paintbox1.Canvas.Brush.Style:=bsclear;

  ActivateRenderingContext(dc, hrc); // активируем контекст рендеринга

  glClearColor(0.0, 0.2, 0.3, 0.5); // цвет фона голубой
  glClearDepth(1.0);
  glEnable(GL_DEPTH_TEST); // включить режим тест глубины
  glEnable(GL_ALPHA_TEST);
  // glEnable(GL_BLEND);
  glBlendFunc(GL_SRC_ALPHA, GL_ONE);
  glLightModelfv(GL_LIGHT_MODEL_AMBIENT, @l_model);
  glLightModeli(GL_LIGHT_MODEL_TWO_SIDE, GL_TRUE);
  // glEnable(GL_CULL_FACE); //включить режим отображени€ только передних поверхностей

  BuildFont();

end;

procedure TOpenGLThread.MainThreadProc;
var
  s: Cardinal;
begin
  s := timeGetTime;
  inc(FpsCounter);
  if s > PredTimeFPS + 1000 then // >1s
  begin
    FText := 'FPS:' + FormatFloat('#0.00', FpsCounter * 1000 /
      (s - PredTimeFPS));
    FpsCounter := 0;
    PredTimeFPS := s;
  end;
  vNozzleStates.UpdateByTimerTick;
  Render;
end;

procedure TOpenGLThread.Execute;
begin
  SetupGL;
  PredTimeFPS := timeGetTime;
  FpsCounter := 0;
  ExecPeriod := VisualizationDiscretization;

  inherited Execute;

  KillFont;
  DeactivateRenderingContext;
  DestroyRenderingContext(hrc);

  ReleaseDC(Handle, dc);
  if Assigned(DrawerOGL) then
    DrawerOGL.Free;
end;

procedure TOpenGLForm.FormMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: integer);
begin
  if ssLeft in Shift then
    with RenderThread do
    begin
      Angle1 := StartA1 + (X - StartMouseX) * delta;
      Angle2 := StartA2 + (Y - StartMouseY) * delta;
    end
  else if ssRight in Shift then
    with RenderThread do
    begin
      if ssCtrl in Shift then
      begin
        PosViewZ := StartA3 + (Y - StartMouseY) * delta;
      end
      else
      begin
        PosViewX := StartA1 + (X - StartMouseX) * delta;
        PosViewY := StartA2 - (Y - StartMouseY) * delta;
      end;
    end;
end;

procedure TOpenGLForm.FormResize(Sender: TObject);
begin
  RenderThread.Width := Self.Width;
  RenderThread.Height := Self.Height;
end;

procedure TOpenGLForm.FormDestroy(Sender: TObject);
begin
  RenderThread.Terminate;
  RenderThread.WaitFor;
  RenderThread.Free;
end;

procedure TOpenGLForm.FormMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: integer);
begin
  if ssLeft in Shift then
    with RenderThread do
    begin
      StartMouseX := X;
      StartMouseY := Y;
      StartA1 := Angle1;
      StartA2 := Angle2;
    end
  else if ssRight in Shift then
    with RenderThread do
    begin
      StartMouseX := X;
      StartMouseY := Y;
      if ssCtrl in Shift then
      begin
        StartA3 := PosViewZ;
      end
      else
      begin
        StartA1 := PosViewX;
        StartA2 := PosViewY;
      end;
    end;
end;

procedure TOpenGLForm.FormShow(Sender: TObject);
var
  SL: TStringList;
begin
  SL := TStringList.Create;
  try
    SL.Values['PosViewX'] := '-0,1';
    SL.Values['PosViewY'] := '-0,7';
    SL.Values['PosViewZ'] := '-5';
    SL.Values['Angle1'] := '0';
    SL.Values['Angle2'] := '-13';
    LoadLayouts(Self, SL);
    if Assigned(RenderThread) then
    begin
      RenderThread.PosViewX := StrToFloat(SL.Values['PosViewX']);
      RenderThread.PosViewY := StrToFloat(SL.Values['PosViewY']);
      RenderThread.PosViewZ := StrToFloat(SL.Values['PosViewZ']);
      RenderThread.Angle1 := StrToFloat(SL.Values['Angle1']);
      RenderThread.Angle2 := StrToFloat(SL.Values['Angle2']);
    end;
  finally
    SL.Free;
  end;
  RenderThread.Showed := true;
end;

procedure TOpenGLForm.SaveLayoutsThisForm;
var
  SL: TStringList;
begin
  SL := TStringList.Create;
  try
    if Assigned(RenderThread) then
    begin
      SL.Values['PosViewX'] := FloatToStr(RenderThread.PosViewX);
      SL.Values['PosViewY'] := FloatToStr(RenderThread.PosViewY);
      SL.Values['PosViewZ'] := FloatToStr(RenderThread.PosViewZ);
      SL.Values['Angle1'] := FloatToStr(RenderThread.Angle1);
      SL.Values['Angle2'] := FloatToStr(RenderThread.Angle2);
    end;
    SaveLayouts(Self, SL);
  finally
    SL.Free;
  end;
end;

procedure TOpenGLForm.Stayontop1Click(Sender: TObject);
begin
   if Stayontop1.Checked then
       FormStyle:=fsStayOnTop else FormStyle:=fsNormal;
end;

procedure TOpenGLForm.DayMode1Click(Sender: TObject);
begin
  fgDayMode := DayMode1.Checked;
end;

procedure TOpenGLForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if Action = caFree then
    SaveLayoutsThisForm;
end;

procedure TOpenGLForm.FormHide(Sender: TObject);
begin
  SaveLayoutsThisForm;
  RenderThread.Showed := false;
end;

function  TOpenGLForm.GetShowed:boolean;
begin
   Result:=RenderThread.Showed;
end;


initialization

OpenGLForm := nil;

end.
