unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls,mmsystem,types, AviWriter_2, StdCtrls;

type
  TForm2 = class(TForm)
    Panel1: TPanel;
    Start_IMG: TImage;
    Stop_IMG: TImage;
    time_lb: TLabel;
    Comand: TLabel;
    show_timer: TTimer;
    close_timer: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure CloseXvidWindow(Sender: TObject);
    procedure showTime(Sender: TObject);
    procedure start_timers(Sender: TObject);
    procedure onClose(Sender: TObject; var Action: TCloseAction);
    procedure ButtonStartStop(Sender: TObject);
    procedure ButtonStop(Sender: TObject);
    procedure Button_start(Sender: TObject);
    { Private declarations }
  public
    { Public declarations }
    procedure on_tmr(Sender: TObject);
    procedure AviWriterProgress(Sender: TObject; FrameCount: Integer;  var abort: Boolean);
    procedure wdm;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;
  //
  ftimer, close_timer, show_timer: ttimer;
  AviWriter: TAviWriter_2;
  on_avi, on_xvid_wnd: boolean;

  gl_wdm: string;
  gl_rec,gl_cnt: integer;
  b: tbitmap;
  contTime : Integer;


implementation

{$R *.dfm}
function CheckFormat(SDate:string):string;
var IDateChar: string;
    x,y      : integer;
begin
 IDateChar:= ':';
 for y:=1 to length(IDateChar) do begin
  x:=pos(IDateChar[y],SDate);
  while x>0 do begin
   Delete(SDate,x,1);
   Insert('-',SDate,x);
   x:= pos(IDateChar[y],SDate);
  end
 end;
 CheckFormat:= SDate
end;

// ����� ����������� ========================
procedure tform2.wdm;
var
    path : string;
begin
  on_avi:= not on_avi;

  if on_avi then begin //�������������-
   //
   with AviWriter do begin
    path := GetEnvironmentVariable('Temp')+'\VideoForRequest';
    if not(directoryexists(path)) then
    MkDir(path);

    AviWriter.filename    := path + '\Movie.avi';
    AviWriter.TempFileName:= path + '\' + ExtractFilePath(AviWriter.filename) + '~AWTemp' + ExtractFileName(AviWriter.filename);
    frameTime             := ftimer.Interval;
    OnTheFlyCompression   := true;
    width:= GetSystemMetrics(0);
    height:= GetSystemMetrics(1);
    //
    SetCompression('XVID XVID');
    SetCompressionQuality(2000)
   end;
   AviWriter.InitVideo;
   gl_rec:= gettickcount; //����� ������-
  end else begin //��������� ������-
   AviWriter.FinalizeVideo;
   AviWriter.WriteAvi;
  end;
end;
procedure Tform2.AviWriterProgress(Sender: TObject; FrameCount: Integer;
  var abort: Boolean);
begin
 gl_cnt:= FrameCount;
end;

procedure Tform2.on_tmr(Sender: TObject);
var
  CurInfo: tagCURSORINFO;
  IcoInfo: _ICONINFO;
  ACursor: HICON;
  Pt: TPoint;
begin
 //������ �� �������-
  if on_avi then
  begin
    BitBlt(b.Canvas.Handle, 0, 0, Screen.Width, Screen.Height,
      GetDC(0), 0, 0, SRCCopy);

   // -- ������
    CurInfo.cbSize := SizeOf(CurInfo);
    GetCursorInfo(CurInfo);
    ACursor := CurInfo.hCursor;
    Pt := CurInfo.ptScreenPos;
    GetIconInfo(ACursor, IcoInfo);
    DrawIcon(b.Canvas.Handle, Pt.X - IcoInfo.xHotspot, Pt.Y - IcoInfo.yHotspot, ACursor);

    AviWriter.AddFrame(b);

  end

end;

procedure TForm2.FormCreate(Sender: TObject);
begin
 SetWindowLong(application.Handle,GWL_EXSTYLE,GetWindowLong(application.Handle, GWL_EXSTYLE) or
  not WS_EX_APPWINDOW);

 //
 b:= tbitmap.Create;
 b.Width:= GetSystemMetrics(0);
 b.Height:= GetSystemMetrics(1);
 //
 AviWriter:= TAviWriter_2.Create(nil);
 AviWriter.OnProgress:= AviWriterProgress;
 //
 ftimer:= ttimer.Create(nil);
 ftimer.interval:= Round(1000/{FPS}6{������/���});   //  FPS
 ftimer.ontimer := on_tmr;
 ftimer.Enabled := true;


end;

procedure TForm2.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
 freeandnil(ftimer);
 freeandnil(AviWriter);
 b.Free
end;
procedure TForm2.ButtonStartStop(Sender: TObject);
begin
  if on_avi then
  begin
    Comand.Caption := '�����';
    show_timer.Enabled := False;
    wdm;
    Start_IMG.Visible := True;
    Stop_IMG.Visible := False;
    ModalResult := 1;
  end
  else
  begin
    Comand.Caption := '����';
    show_timer.Enabled := true;
    close_timer.Enabled := true;
    //Form2.BorderIcons := [];
    wdm;
    Start_IMG.Visible := False;
    Stop_IMG.Visible := True;
  end;

end;

procedure TForm2.CloseXvidWindow(Sender: TObject);
var    HwndM : HWND;
begin

   //��������//���� � ����������
   HwndM := findwindow(nil, 'Xvid Status');
   if HwndM <>0  then
    begin
      SendMessage (HwndM, WM_CLOSE, 0, 0);
      close_timer.Enabled := False;
      freeandnil(close_timer);
    end;
end;

procedure TForm2.showTime(Sender: TObject);
begin
if contTime <= 0 then
begin
 ButtonStartStop(nil);
end;
  Dec(contTime);
  time_lb.Caption := Format('%.2d:%.2d', [contTime div 60, contTime mod 60]);
  if contTime < 11 then
     if contTime mod 2 = 0 then time_lb.Font.Color := clRed
     else time_lb.Font.Color := clBlack;
end;

procedure TForm2.start_timers(Sender: TObject);
begin
    contTime := 120; // 2 min
    time_lb.Font.Color := clBlack;
    time_lb.Caption := Format('%.2d:%.2d', [contTime div 60, contTime mod 60]);


end;



procedure TForm2.onClose(Sender: TObject; var Action: TCloseAction);
begin
    if on_avi then ButtonStartStop(nil);
end;

procedure TForm2.ButtonStop(Sender: TObject);
begin
  ButtonStartStop(nil);

end;

procedure TForm2.Button_start(Sender: TObject);
begin 
  ButtonStartStop(nil);

end;

end.
