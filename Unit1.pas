unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, S4_TLB, DateUtils, DB, ADODB, ComObj, Menus,
  ActnMan, ActnColorMaps, ToolWin, ActnCtrls, ActnMenus, Registry, jpeg,
  ComCtrls, Unit2;

type
  TForm1 = class(TForm)
    mainFrame: TPanel;
    caption: TPanel;
    closeBtn: TPanel;
    maximizeBtn: TPanel;
    minimizeBtn: TPanel;
    iconPanel: TPanel;
    imgicon: TImage;
    closeIMG: TImage;
    maxIMG: TImage;
    minIMG: TImage;
    normIMG: TImage;
    mainBorder: TShape;
    caption_title: TLabel;
    footer: TPanel;
    resizeIMG: TImage;
    shpBCKGND: TShape;
    ResizeLBL: TLabel;
    resizePNL: TPanel;
    closeLB: TLabel;
    max_NormLB: TLabel;
    minimLB: TLabel;
    typeLB: TLabel;
    departLB: TLabel;
    FIOLB: TLabel;
    TabNoLB: TLabel;
    telLB: TLabel;
    invNoLB: TLabel;
    descriptLB: TLabel;
    PDSPNL: TPanel;
    ispolLB: TLabel;
    category: TLabel;
    categLB: TLabel;
    countText: TLabel;
    symbolsLBL: TLabel;
    browseLB: TLabel;
    invNoName: TLabel;
    ADOConnection_ACCESS: TADOConnection;
    ADOQuery_getInvNo1: TADOQuery;
    ADOQueryGETINVNO: TADOQuery;
    ADOQueryInvNoChange: TADOQuery;
    pmColorChange: TPopupMenu;
    BorderColor1: TMenuItem;
    dlgColor1: TColorDialog;
    ResetColors1: TMenuItem;
    Cancel_BTNPNL: TPanel;
    Cancel_BTNLB: TLabel;
    Cancel_BTNSHP: TShape;
    Report_BTNPNL: TPanel;
    Report_BTNSHP: TShape;
    Report_BTNLB: TLabel;
    REGISTER_BTNPNL: TPanel;
    REGISTER_BTNSHP: TShape;
    REGISTER_BTNLB: TLabel;
    Browse_BTNPNL: TPanel;
    Browse_BTNSHP: TShape;
    Browse_BTNLB: TLabel;
    EDIT_PNL_DEPART: TPanel;
    EDIT_EDIT_PNL_DEPART: TPanel;
    SH_DEPART: TShape;
    EDIT_EDIT_DEPART: TEdit;
    EDIT_LB_PNL_DEPART: TPanel;
    EDIT_LB_DEPART: TLabel;
    EDIT_PNL_FIO: TPanel;
    EDIT_EDIT_PNL_FIO: TPanel;
    SH_FIO: TShape;
    EDIT_EDIT_FIO: TEdit;
    EDIT_LB_PNL_FIO: TPanel;
    EDIT_LB_FIO: TLabel;
    EDIT_PNL_TABNO: TPanel;
    EDIT_EDIT_PNL_TABNO: TPanel;
    SH_TABNO: TShape;
    EDIT_EDIT_TABNO: TEdit;
    EDIT_LB_PNL_TABNO: TPanel;
    EDIT_LB_TABNO: TLabel;
    EDIT_PNL_TEL: TPanel;
    EDIT_EDIT_PNL_TEL: TPanel;
    SH_TEL: TShape;
    EDIT_EDIT_TEL: TEdit;
    EDIT_LB_PNL_TEL: TPanel;
    EDIT_LB_TEL: TLabel;
    EDIT_PNL_BROWSE: TPanel;
    EDIT_EDIT_PNL_BROWSE: TPanel;
    SH_BROWSE: TShape;
    EDIT_LB_PNL_BROWSE: TPanel;
    EDIT_LB_BROWSE: TLabel;
    EDIT_EDIT_BROWSE: TEdit;
    MEMO_PNL: TPanel;
    MEMO_MEMO_PNL: TPanel;
    SH_MEMO: TShape;
    MEMO_MEMO: TMemo;
    MEMO_LB_PNL: TPanel;
    LB_MEMO: TLabel;
    COMBO_PNL_INVNO: TPanel;
    COMBO_CMB_PNL_INVNO: TPanel;
    SH_INVNO: TShape;
    CMB_INVNO: TComboBox;
    COMBO_LB_PNL_INVNO: TPanel;
    LB_INVNO: TLabel;
    COMBO_PNL_ISPOL: TPanel;
    COMBO_CMB_PNL_ISPOL: TPanel;
    SH_ISPOL: TShape;
    CMB_ISPOL: TComboBox;
    COMBO_LB_PNL_ISPOL: TPanel;
    LB_ISPOL: TLabel;
    BackgroundColor1: TMenuItem;
    TECHMACH_BTN: TRadioButton;
    OASU_RBTN: TRadioButton;
    BTN_SCR: TPanel;
    SH_SCR: TShape;
    LB_SCR: TLabel;
    BTN_VIDEO: TPanel;
    SH_VIDEO: TShape;
    LB_VIDEO: TLabel;



    procedure FormCreate(Sender: TObject);
    procedure INITIALIZE(Sender: TObject);
    procedure main();
    procedure CreateObj();
    procedure InfoUser();
    procedure Config();
    procedure Register();

    procedure DescriptionTextChanged(Sender: TObject);
    procedure BrowseBTNClick(Sender: TObject);
    procedure sendBTNClick(Sender: TObject);
    procedure ReportBTNClick(Sender: TObject);
    procedure INvNoChange(Sender: TObject);
    procedure Change_TYPE_REQ(Sender: TObject);
    procedure controlOfFile(path:string);
//=====================
// FORM COLOR EDITOR
//=====================
    procedure changeBorderColor(BorderColor : TColor);
    procedure changeBackgroungColor(BackgroungColor : TColor);
    function increaseBrightness(inColor : TColor; valInc : Integer): TColor;
    function RGB(r, g, b: Integer): TColor;
    procedure ColorPicker(Sender: TObject);
//======================
// SETTINGS
//======================
    procedure SaveSettings();
    procedure LoadSettings();
//======================
//CAPTION BUTTONS
//======================
    procedure CLOSE_WINDOW(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure PRESSED_BTN(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure MINIM_WINDOW(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure MAX_NORM_CLICK(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure MOUSE_ENTER(Sender: TObject);
    procedure MOUSE_LEAVE(Sender: TObject);
    procedure DBL_MAXIM_WINDOW(Sender: TObject);
    procedure MoveIconMainFrameIcon(moveDown: BOOL);
    procedure tesd(Sender: TObject);
//======================
//WINDOW MOVE
//======================
    procedure DRAG_WIMDOW_UP(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure DRAG_WDN_DOWN(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure DRAG_WIMDOW_MOVE(Sender: TObject; Shift: TShiftState; X, Y: Integer);

//======================
// RESIZE
//======================
    procedure RESIZE_ENTER(Sender: TObject);
    procedure RESIZE_LEAVE(Sender: TObject);
    procedure RESIZE_MOVE(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure RESIZE_WND_DOWN(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure RESIZE_WND_UP(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

//======================
// EDIT
//======================
    procedure EDIT_TEXT_CHANGE(Sender: TObject);
    procedure EDIT_LB_CLICK(Sender: TObject);
    procedure EDIT_TEXT_EXIT(Sender: TObject);
    procedure EDIT_LB_ENTER(Sender: TObject);
    procedure EDIT_LB_LEAVE(Sender: TObject);
    function getText(Sender: TObject) : string;
    procedure setText(Sender: TObject; text : string);

//=====================
// BUTTON
//=====================
    procedure BUTTON_ENTER(Sender: TObject);
    procedure BUTTON_LEAVE(Sender: TObject);
    procedure BUTTON_CLICK_DOWN(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BUTTON_CLICK_CLOSE(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GetPrintScreen(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BUTTON_REPORT_UP(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BUTTON_REGISTER_UP(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BUTTON_CLICK_UP(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GetVideo(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

type TConfig = record
     DocTypeName, ClassifFolder : String;  //document name, classifficator folder
     DocTypeID, ArchID, WFlowID : Integer; //id document type, id archive, id workflow
   end;
var
  Form1: TForm1;
  lastPanel : TPanel;
  normal_height, normal_width, XPos, YPos, R, G, B, Y, Cb, Cr: Integer;
  Moving : Boolean;
  BorderColor, BackgroungColor, FontColor : TColor;

  S4 : TS4App;
  OASU , PDS : TConfig;
  Reg : TRegistry;

implementation

{$R *.dfm}
Function convertDepart_forSQL() : String;
var t : String;
begin
  t:= Form1.EDIT_EDIT_DEPART.Text;
  if 'ЦЕХ №1' = t then Result := 'Ц.1'
  else if 'ЦЕХ №2' = t then Result := 'Ц.2'')) OR (((Книга.[Место эксплуатации])=''Ц.2/Ц.3'
  else if 'ЦЕХ №3' = t then Result := 'Ц.3'')) OR (((Книга.[Место эксплуатации])=''Ц.2/Ц.3'')) OR (((Книга.[Место эксплуатации])=''Ц.3(4)'
  else if 'ЦЕХ №4' = t then Result := 'Ц.4'')) OR (((Книга.[Место эксплуатации])=''Ц.3(4)'
  else if 'ЦЕХ №5' = t then Result := 'Ц.5'
  else if 'ЦЕХ №6' = t then Result := 'Ц.6'
  else if 'ЦЕХ №7' = t then Result := 'Ц.7'')) OR (((Книга.[Место эксплуатации])=''Ц.9/Ц.7'
  else if 'ЦЕХ №9' = t then Result := 'Ц.9'')) OR (((Книга.[Место эксплуатации])=''Ц.9/Ц.11'')) OR (((Книга.[Место эксплуатации])=''Ц.9/Ц.7'
  else if 'ЦЕХ №10' = t then Result := 'Ц.10'')) OR (((Книга.[Место эксплуатации])=''Ц.10 (СЗОС)'')) OR (((Книга.[Место эксплуатации])=''Ц.10 Сморгонь'
  else if 'ЦЕХ №11' = t then Result := 'Ц.11'')) OR (((Книга.[Место эксплуатации])=''Ц.11 Лида'')) OR (((Книга.[Место эксплуатации])=''Ц.9/Ц.11'
  else if 'ЦЕХ №12' = t then Result := 'Ц.12'
  else if 'ЦЕХ №15' = t then Result := 'Ц.15'
  else if 'ЦЕХ №16' = t then Result := 'Ц.16'
  else if 'ЦЕХ №18' = t then Result := 'Ц.18'
  else if 'БСАПР ОАСУ' = t then Result := 'Ц.9'')) OR (((Книга.[Место эксплуатации])=''Ц.9/Ц.11'')) OR (((Книга.[Место эксплуатации])=''Ц.9/Ц.7'  //===========TEST [REMOVE AFTER TEST]
  else Result := '';
end;

Function getUserIspolnitelList() : TStringList;
var
   op : TStringList;
   t,p : String;
   n : Integer;
begin
  op := TStringList.Create;
  S4.OpenQuery( 'SELECT NAME_GROUP'+
                  ' FROM Search.dbo.GROUPS'+
                  ' WHERE NAME_GROUP LIKE ''ITIL PDS%''' );

  Repeat
    p := S4.QueryFieldByName( 'NAME_GROUP' );
    if Pos('-', p) = 0 then
    begin
      t:='';
      n:= Length(p) - 9;
      op.Add(Copy(p,10,n));// like function Substring(9);
    end;
  until  S4.QueryGoNext = 0 ;
  S4.CloseQuery;
  Result := op;
end;
//get info about user and fill in form
procedure TForm1.InfoUser();
var
  id : Integer;
  queryNmae : string;
  buffer: array[0..MAX_COMPUTERNAME_LENGTH + 1] of Char;
  nSize : Cardinal;
begin
      id := S4.GetUserID();
      queryNmae := 'select g.name_group'+
      ' from grpingrp l inner join grpingrp l2 on l2.group_id = l.ingroup_id and l2.ingroup_id <> 999999998 inner join groups g on  g.group_id = l.ingroup_id and g.group_code <> ""'+
      ' inner join groups g2 on g2.group_id = l.group_id and g2.user_id =' + IntToStr(id);
      S4.OpenQuery(queryNmae );
      Form1.EDIT_EDIT_DEPART.Text := S4.QueryFieldByName('name_group');
      S4.CloseQuery();
      Form1.EDIT_EDIT_FIO.Text := S4.GetUserFullName_ByUserID(id);
      nSize := 1024;
      GetComputerName(@buffer, nSize);
      Form1.CMB_INVNO.Text := StrPas(buffer);
      S4.OpenQuery('select WORKPHONE from USERS_INFO where USER_ID =' + IntToStr(id));
      Form1.EDIT_EDIT_TEL.Text := S4.QueryFieldByName('WORKPHONE');
      S4.CloseQuery();
      nSize := 255;
      Windows.GetUserName(@buffer, nSize);
      Form1.EDIT_EDIT_TABNO.Text := StrPas(buffer);
      EDIT_LB_LEAVE(EDIT_LB_DEPART);
      EDIT_LB_LEAVE(EDIT_LB_TEL);
      EDIT_LB_LEAVE(EDIT_LB_TABNO);
      EDIT_LB_LEAVE(EDIT_LB_FIO);
      EDIT_LB_LEAVE(LB_MEMO);
      Change_TYPE_REQ(nil);
      EDIT_LB_LEAVE(LB_INVNO);
      INvNoChange(CMB_INVNO);
end;



//start search workflow
procedure Tform1.Register;
var
  extantion, folder, folderKey, designation , fileName, fullFileName, otmetka, nameProcess, group, sqlQuery, grourID: string;
  res : TStrings;
  t, objType, docID : Integer;
  conf : TConfig;
  Router, Process, Varrs :OLEVariant;
  perfomance,myDate : TDateTime;
begin
  perfomance := Now;
  extantion := 'doc';
  objType := 0;
  if OASU_RBTN.Checked then
        conf := OASU
  else if t = 1 then
        conf := PDS;
  folder := 'D:\SEARCHWORK\';
  Router := S4.GetSbServer.GetRouter;
  Process := Router.CreateProcess(conf.WFlowID);
  Varrs:= Process.StartActivity.Variables;
  folderKey := S4.GetClassificatorInterface.OpenFolderByName(conf.ClassifFolder);
  designation := S4.GetClassificatorInterface.GetDesignationByKey(folderKey, '');
  fileName := S4.GenerateFileName('', extantion);
  fullfilename := folder + filename;
  if Not (DirectoryExists(folder)) then
        CreateDir(folder);

  docID := S4.CreateFileDocumentWithDocType2(fullfilename, conf.DocTypeID, conf.ArchID, filename, '', objType);
  S4.OpenDocument(docID);
  if Length(EDIT_EDIT_BROWSE.Text) > 2 then
        S4.AppendAdvanFile2(EDIT_EDIT_BROWSE.Text, 0);
  S4.SetFieldValue('DESIGNATIO', designation);
  S4.SetFieldValue('NAME', conf.DocTypeName);
  S4.SetFieldValue('FIO', EDIT_EDIT_FIO.Text);
  S4.SetFieldValue('PROBLEM', MEMO_MEMO.Text);
  S4.SetFieldValue('OTDEL', EDIT_EDIT_DEPART.Text);
  S4.SetFieldValue('OTMETKA' , 'Не выполнено' );

  if t = 1 then
  begin
        S4.SetFieldValue('NAZV_OBORUD' , invNoName.Caption);
        S4.SetFieldValue('LAST_WHEN_WHOM' , CMB_ISPOL.Text );
        otmetka := 'Конфликт';
        nameProcess := 'Заявка. инв№ "' + CMB_INVNO.Text + '". сервис № ' + designation + '  ' + EDIT_EDIT_DEPART.Text;
        group := CMB_ISPOL.Text;
        sqlQuery := 'SELECT GROUP_ID   FROM DBO.GROUPS     WHERE NAME_GROUP = ''ITIL PDS ' + group + #39;
        S4.OpenQuery( sqlQuery );
        grourID := S4.QueryFieldByName( 'GROUP_ID' );
        Varrs.GetVariableByName( 'ISPOLNITEL' ).asString := grourID + '=1';
        Varrs.GetVariableByName( 'PRIORITET' ).asString := category.Caption;
        res := TStringList.Create;
        res.Add( #13#10 );
        res.AddStrings( CMB_ISPOL.Items);
        Varrs.GetVariableByName( 'SQL_group' ).asString := res.Text;
        ShowMessage(#39 + res.Text + #39);
  end
  else
        begin
          otmetka := 'Выполнено частично';
          nameProcess := conf.DocTypeName + ' в сервиc № ' + designation;
        end;

  S4.CheckIn();
  S4.CloseDocument( );
  S4.GetClassificatorInterface.IncludeDocument(docID);
  Process.StartActivity.Attachments.AddLink(docID);

  Varrs.GetVariableByName( 'OTMETKA' ).asString := 'Не выполнено' + #10#13 + 'Не выполнено' + #13 + 'Выполнено' + #13 + otmetka + #13;
  Varrs.GetVariableByName( 'ZAYAVKA' ).value := designation;
  Varrs.GetVariableByName( 'TITLE' ).value := 'Заявка № ' + designation + ' от ' + S4.GetFieldValue( 'CREATEDATE' );
  Varrs.GetVariableByName( 'CREATEDATE' ).value := S4.GetFieldValue( 'CREATEDATE' );
  Varrs.GetVariableByName( 'OTDEL' ).value := EDIT_EDIT_DEPART.Text;
  Varrs.GetVariableByName( 'PROBLEM' ).value := MEMO_MEMO.Text;
  Varrs.GetVariableByName( 'LOGIN' ).value := EDIT_EDIT_TABNO.Text;
  myDate := Now;
  Varrs.GetVariableByName( 'KONTRDATE' ).value := formatdatetime('dd.mm.yyyy', myDate);
  myDate := IncHour(myDate, 2);
  Varrs.GetVariableByName( 'KONTRTIME' ).value := formatdatetime('hh:nn:ss', myDate);
  Varrs.GetVariableByName( 'FIO' ).value := EDIT_EDIT_FIO.Text;
  Varrs.GetVariableByName( 'INVNOMER' ).value := CMB_INVNO.Text;
  Varrs.GetVariableByName( 'TEL' ).value := EDIT_EDIT_TEL.Text;
  Process.Name := nameProcess;
  Process.Start;
  Process := Unassigned;

  perfomance := perfomance - Now;
  ShowMessage('OK!' + #13 + '(' + formatdatetime('ss.zzz', perfomance) + ' sec)' );
  CLOSE_WINDOW(nil, mbLeft, [ssLeft], 0, 0);
end;

procedure TForm1.Config();
begin
   OASU.DocTypeName := 'Заявка в ОАСУ';
   OASU.ClassifFolder := 'Журнал работ БРОВТ';
   OASU.DocTypeID := 1000121;
   OASU.ArchID := 448;
   OASU.WFlowID := 138349;
   PDS.DocTypeName := 'Заявка в ЭДС';
   PDS.ClassifFolder := 'ПДС';
   PDS.DocTypeID := 1000536;
   PDS.ArchID := 621;
   PDS.WFlowID := 253829;
end;  
//simple S4 connect and created obj
procedure TForm1.CreateObj();
begin
   S4 := coTS4App.Create();
   if S4.Login <> 1 then
   Application.Terminate;
end;

//**************************************
//======================================
//                    MAIN
//======================================
//**************************************
procedure TForm1.main();
begin
      Config();
      CreateObj();
      InfoUser();
end;

//************************************
//           for window  style
//************************************
{$IFDEF undef}{$REGION 'for window  style'}{$ENDIF}


procedure TForm1.INITIALIZE(Sender: TObject);
begin
  BackgroungColor := TColor($f2f2f2);
  BorderColor := {R}22 + {G}52 * 256 + {B}122 * 256 * 256;
  main();
  LoadSettings;
  changeBorderColor(BorderColor);
  changeBackgroungColor(BackgroungColor);
end;






procedure TForm1.SaveSettings;
begin
  Reg:=TRegistry.Create;
  Reg.RootKey := HKEY_CURRENT_USER;
  Reg.OpenKey('\SOFTWARE\ITIL_1211', true);

  Reg.WriteInteger('BorderColor', BorderColor);
  Reg.WriteInteger('BackgroungColor', BackgroungColor);
  Reg.WriteInteger('Left', Form1.Left);
  Reg.WriteInteger('Top', Form1.Top);
  Reg.WriteInteger('Width', Form1.Width);
  Reg.WriteInteger('Height', Form1.Height);
  Reg.WriteInteger('TypeReq', Integer(OASU_RBTN.Checked));

  Reg.Free;
end;
procedure TForm1.LoadSettings;
var Reg: TRegistry;
b:BOOL;
i:Integer;
begin
  Reg:=TRegistry.Create;
  Reg.RootKey := HKEY_CURRENT_USER;
  Reg.OpenKey('\SOFTWARE\ITIL_1211', true);

  if reg.ValueExists('BorderColor') then BorderColor := Reg.ReadInteger('BorderColor');
  if reg.ValueExists('BackgroungColor') then BackgroungColor := Reg.ReadInteger('BackgroungColor');
  if reg.ValueExists('Left') then Form1.Left := Reg.ReadInteger('Left');
  if reg.ValueExists('Top') then Form1.Top := Reg.ReadInteger('Top');
  if reg.ValueExists('Width') then Form1.Width := Reg.ReadInteger('Width');
  if reg.ValueExists('Height') then Form1.Height := Reg.ReadInteger('Height');
  if reg.ValueExists('TypeReq') then OASU_RBTN.Checked := Reg.ReadInteger('TypeReq') = 1;
  TECHMACH_BTN.Checked := not(OASU_RBTN.Checked);

  Reg.free;
end;
procedure TForm1.changeBorderColor(BorderColor : TColor);
begin

  mainBorder.Pen.Color := BorderColor;
  footer.Color := BorderColor;
  shpBCKGND.Pen.Color := BorderColor;
end;
procedure TForm1.changeBackgroungColor(BackgroungColor : TColor);
begin
  mainFrame.Color := BackgroungColor;
  shpBCKGND.Brush.Color := BackgroungColor;
  Cancel_BTNPNL.Color := BackgroungColor;
  //ReColor window button
  closeBtn.Color := BackgroungColor;
  minimizeBtn.Color := BackgroungColor;
  maximizeBtn.Color := BackgroungColor;
  //to get new Y
  increaseBrightness(BackgroungColor, 0);

//  FontColor := clGray;
FontColor := clBlack;
  if Y < 100 then FontColor := clWhite;

  //ReColor Border Custom style Buttons
  Cancel_BTNSHP.Pen.Color := FontColor;
  SH_SCR.Pen.Color := FontColor;
  SH_VIDEO.Pen.Color := FontColor;
  Report_BTNSHP.Pen.Color := FontColor;
  REGISTER_BTNSHP.Pen.Color := FontColor;
  Browse_BTNSHP.Pen.Color := FontColor;

  //ReColor Label Custom style Buttons
  Cancel_BTNLB.Font.Color := FontColor;
  LB_SCR.Font.Color := FontColor;
  Report_BTNLB.Font.Color := FontColor;
  REGISTER_BTNLB.Font.Color := FontColor;
  Browse_BTNLB.Font.Color := FontColor;

  //ReColor Label Custom style Buttons
  Cancel_BTNPNL.Color := BackgroungColor;
  Report_BTNPNL.Color := BackgroungColor;
  REGISTER_BTNPNL.Color := BackgroungColor;
  Browse_BTNPNL.Color := BackgroungColor;

  //Recolor Font all children elements
  mainFrame.Font.Color := FontColor;

  //ReColor all label
  caption_title.Font.Color := FontColor;
  typeLB.Font.Color := FontColor;
  departLB.Font.Color := FontColor;
  FIOLB.Font.Color := FontColor;
  TabNoLB.Font.Color := FontColor;
  telLB.Font.Color := FontColor;
  invNoLB.Font.Color := FontColor;
  categLB.Font.Color := FontColor;
  category.Font.Color := FontColor;
  invNoName.Font.Color := FontColor;
  ispolLB.Font.Color := FontColor;
  descriptLB.Font.Color := FontColor;
  countText.Font.Color := FontColor;
  symbolsLBL.Font.Color := FontColor;
  browseLB.Font.Color := FontColor;

  //Recolor Edit border
  SH_BROWSE.Pen.Color := FontColor;
  SH_ISPOL.Pen.Color := FontColor;
  if SH_DEPART.Pen.Color <> clRed then SH_DEPART.Pen.Color := FontColor;
  if SH_FIO.Pen.Color <> clRed then SH_FIO.Pen.Color := FontColor;
  if SH_TABNO.Pen.Color <> clRed then SH_TABNO.Pen.Color := FontColor;
  if SH_TEL.Pen.Color <> clRed then SH_TEL.Pen.Color := FontColor;
  if SH_INVNO.Pen.Color <> clRed then SH_INVNO.Pen.Color := FontColor;
  if SH_MEMO.Pen.Color <> clRed then SH_MEMO.Pen.Color := FontColor;


end;
//=====================================================================
// EDIT
//=====================================================================
procedure TForm1.EDIT_TEXT_CHANGE(Sender: TObject);
var
pn1, pn2, pn3 : TPanel;
sh1 : TShape;
lb1 : TLabel;
ed1 : TEdit;
text : string;

begin
        pn1 := (sender as TControl).Parent as TPanel;
        pn2 := pn1.Parent as TPanel;
        pn3 := pn2.Controls[1] as TPanel;
        lb1 := pn3.Controls[0] as TLabel;

        text := getText(sender);
        if Length(text) > 0 then
                setText(lb1,text)
        else
                setText(lb1, 'Заполните поле');
        if Sender is TMemo then DescriptionTextChanged(sender);
end;

procedure TForm1.EDIT_TEXT_EXIT(Sender: TObject);
var
pn1, pn2, pn3 : TPanel;
sh1 : TShape;
lb1 : TLabel;
ed1 : TEdit;
cnt1 : TWinControl;
begin
        cnt1 := sender as TWinControl;
        pn1 := cnt1.Parent as TPanel;
        sh1 := pn1.Controls[0] as TShape;
        pn2 := pn1.Parent as TPanel;
        pn3 := pn2.Controls[1] as TPanel;
        lb1 := pn3.Controls[0] as TLabel;

        pn3.Visible := True;
        if sh1.Pen.Color <> clRed then
                sh1.Pen.Color := clMedGray;
        EDIT_LB_LEAVE(lb1);
        if Pos('BROWSE', cnt1.Name) <> 0 then
                controlOfFile(getText(cnt1));
end;

procedure TForm1.EDIT_LB_CLICK(Sender: TObject);
var
pn1, pn2, pn3 : TPanel;
sh1 : TShape;
lb1 : TLabel;
ed1 : TEdit;
cnt1 : TWinControl;
//tcol : TColor;
begin
        lb1 := sender as TLabel;
        pn1 := lb1.Parent as TPanel;
        pn2 := pn1.Parent as TPanel;
        pn3 := pn2.Controls[0] as TPanel;
        cnt1 := pn3.Controls[1] as TWinControl;
        sh1 := pn3.Controls[0] as TShape;

        cnt1.SetFocus;
        if cnt1 is TComboBox then (cnt1 as TComboBox).DroppedDown := True;
        pn1.Visible := False;
        sh1.Pen.Color := clMedGray;

        sh1.Brush.Color := clWhite;
        pn1.Color := clWhite;
end;

procedure TForm1.EDIT_LB_ENTER(Sender: TObject);
var
pn1, pn2, pn3 : TPanel;
sh1 : TShape;
lb1 : TLabel;
ed1 : TEdit;
begin
        lb1 := sender as TLabel;
        pn1 := lb1.Parent as TPanel;
        pn2 := pn1.Parent as TPanel;
        pn3 := pn2.Controls[0] as TPanel;
        sh1 := pn3.Controls[0] as TShape;

        sh1.Pen.Color := clBlack;
        sh1.Pen.Color := clMedGray;
end;
procedure TForm1.EDIT_LB_LEAVE(Sender: TObject);
var
  pn1, pn2, pn3: TPanel;
  sh1: TShape;
  lb1: TLabel;
  ed1: TEdit;
  text: string;
begin
  lb1 := sender as TLabel;
  pn1 := lb1.Parent as TPanel;
  pn2 := pn1.Parent as TPanel;
  pn3 := pn2.Controls[0] as TPanel;
  sh1 := pn3.Controls[0] as TShape;
  text := getText(pn3.Controls[1]);

  if pn1.Visible then
  begin
    sh1.Pen.Color := FontColor;

    if (Length(text) > 0) or (Pos('BROWSE', lb1.Name) <> 0) then
    begin
      lb1.Font.Color := clBlack;
      sh1.Pen.Color := FontColor;
      sh1.Brush.Color := clWhite;
      pn1.Color := clWhite;

    end
    else
    begin
      lb1.Font.Color := clGray;
      sh1.Pen.Color := clRed;
      sh1.Pen.Color := clRed;
      sh1.Brush.Color := $ccccF7;
      pn1.Color := $ccccF7;
    end;
  end;
end;
function TForm1.getText(Sender: TObject) : string;
begin
 if Sender is TEdit then Result := (Sender as TEdit).Text
 else   if Sender is TMemo then Result := (Sender as TMemo).Text
 else   if Sender is TComboBox then Result := (Sender as TComboBox).Text
 else   if Sender is TLabel then Result := (Sender as TLabel).Caption;
end;

procedure TForm1.setText(Sender: TObject; text : string);
begin
 if Sender is TEdit then (Sender as TEdit).Text := text
 else   if Sender is TMemo then (Sender as TMemo).Text := text
 else   if Sender is TComboBox then (Sender as TComboBox).Text := text
 else   if Sender is TLabel then (Sender as TLabel).Caption := text
end;


//=====================MOVE=WINDOW=====================================
procedure TForm1.DRAG_WIMDOW_UP(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  Moving:=False;
end;

procedure TForm1.DRAG_WDN_DOWN(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if WindowState <> wsMaximized then
  begin
    XPos:=X;
    YPos:=Y;
    Moving:=True;
  end;

end;

procedure TForm1.DRAG_WIMDOW_MOVE(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
    If Moving then Form1.Left:=Form1.Left+X-XPos;
    If Moving then Form1.Top:=Form1.Top+Y-YPos;
end;

//===================WINDOW=BUTTONS=======================================
procedure TForm1.MINIM_WINDOW(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
         Application.Minimize();
end;

procedure TForm1.MAX_NORM_CLICK(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  DBL_MAXIM_WINDOW(Sender);
end;

procedure TForm1.DBL_MAXIM_WINDOW(Sender: TObject);
  var
    hTaskbar: HWND;
    T: TRect;
begin
  if WindowState = wsNormal then
  //=============MAXIMAZE===============================================
  begin
      normal_height := Form1.Top;
      normal_width := Form1.Left;
      maxIMG.Visible := False;
      normIMG.Visible := True;

      footer.Visible := False;

    hTaskBar := FindWindow('Shell_TrayWnd', nil);
     if hTaskbar <> 0 then
     begin
       GetWindowRect(hTaskBar, T);
     end;

    WindowState := wsMaximized;
    Form1.Top := 0 ;
    Form1.Height := T.Top + 11;
    MoveIconMainFrameIcon(True);
  end
  else
  //===============NORMALIZE==========================================
  begin
   maxIMG.Visible := True;
   normIMG.Visible := False;

   footer.Visible := True;

   Form1.Top := normal_height;
   Form1.Left := normal_width;
   WindowState := wsNormal;
   if normal_height < 30 then
   begin
     normal_height :=300;
   end;
   MoveIconMainFrameIcon(False);
  end;
end;

procedure TForm1.MoveIconMainFrameIcon(moveDown: BOOL);
begin
  if moveDown then
  begin// DOWN ICON
    mainBorder.Top := 0;
    mainFrame.Top := 1;
    shpBCKGND.Top := 0;
  end
  else
  begin// UP ICON
    mainBorder.Top := 8;
    mainFrame.Top := 9;
    shpBCKGND.Top := 8;

  end;
end;








//=====================CLOSE WINDOW+++++++++++++++++++++++++++++
procedure TForm1.CLOSE_WINDOW(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin  
  SaveSettings;
  Release;
  Application.Terminate;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  lastPanel := nil;
end;

procedure TForm1.PRESSED_BTN(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
      var panel :TPanel;
          lb : TLabel;
begin
    lb := sender as TLabel;
    panel :=  lb.Parent as TPanel;
    panel.Color := clRed;
end;

procedure TForm1.tesd(Sender: TObject);
begin
     if WindowState = wsMaximized then
   // MessageDlg('Вы открыли форму во весь рост', mtInformation, [mbOK], 0)
    else if WindowState = wsMinimized then
    //MessageDlg('Вы скрыли форму', mtInformation, [mbOK], 0)
  else if WindowState = wsNormal then
   // MessageDlg('Форма в нормальм состоянии', mtInformation, [mbOK], 0)
end;

//_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+
//
// MOUSE LEAVE-ENTER EVENTS
//_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+_+


procedure TForm1.MOUSE_ENTER(Sender: TObject);
 var panel:TPanel;
      lb : TLabel;
begin
    lb := sender as TLabel;
    panel :=  lb.Parent as TPanel;
    if (panel.Color = BackgroungColor)  then
      begin
        panel.Color := clWhite;
        if(lastPanel <> nil) and (panel <> lastPanel) then
          begin
            lastPanel.Color := BackgroungColor;
          end
      end;
      lastPanel := panel;
  end;

procedure TForm1.MOUSE_LEAVE(Sender: TObject);
begin
  if lastPanel <> nil then begin
      lastPanel.Color := BackgroungColor;
      lastPanel := nil;
  end;
end;

//=========BUTTON=========
procedure TForm1.BUTTON_ENTER(Sender: TObject);
var     lb :TLabel;
        panel : TPanel;
        shape : TShape;
        col: TColor;
begin
    lb := sender as TLabel;
    panel :=  lb.Parent as TPanel;
    shape := panel.Controls[0] as TShape;

    shape.Pen.Color := FontColor;
    lb.Font.Color := FontColor;
    panel.Color := increaseBrightness(BackgroungColor, -10);
    lb.Font.Style := [fsbold];
end;

procedure TForm1.BUTTON_LEAVE(Sender: TObject);
var     lb :TLabel;
        panel : TPanel;
        shape : TShape;
begin
    lb := sender as TLabel;
    panel :=  lb.Parent as TPanel;
    shape := panel.Controls[0] as TShape;
        panel.Color := BackgroungColor;
        shape.Pen.Color := FontColor;
        lb.Font.Color := FontColor;
        lb.Font.Style := [];

end;

// ************************************
// RGB
// ************************************

{Conver to TColor}
function TForm1.RGB(r, g, b: Integer): TColor;
begin
  Result := r + g * 256 + b * 256 * 256;
end;

{}
function TForm1.increaseBrightness(inColor : TColor; valInc : Integer): TColor;
  var v: Integer;
begin
  R := Byte(inColor);
  G := Byte(inColor shr 8);
  B := Byte(inColor shr 16);

  Y := Trunc( 0.299  * R + 0.587  * G + 0.114 * B); // Канал яркости
  if valInc < 100 then
        Y := Round((Y * valInc)/100 + Y);
  if Y > 255 then
        Y := 255
  else if Y < 0 then
        Y := 0;
  Cb := Trunc(-0.1687 * R - 0.3313 * G + 0.5   * B + 128.0);
  Cr := Trunc( 0.5    * R - 0.4187 * G - 0.0813* B + 128.0);

  v := Trunc(Y + 1.772 * (Cb - 128.0));
  if v > 255 then v := 255 else if v < 0 then v := 0;
  B := v;

  v := Trunc(Y - 0.34414 * (Cb - 128.0) - 0.71414 * (Cr - 128.0));
  if v > 255 then v := 255 else if v < 0 then v := 0;
  G := v;

  v := Trunc(Y + 1.402 * (Cr - 128.0));
  if v > 255 then v := 255 else if v < 0 then v := 0;
  R := v;
  Result := RGB(R,G,B);
end;


//===================RESIZE==========================
procedure TForm1.RESIZE_ENTER(Sender: TObject);
begin
    Screen.Cursor := crSizeNWSE;
    GetRValue(Color);
end;

procedure TForm1.RESIZE_LEAVE(Sender: TObject);
begin
    Screen.Cursor := crArrow;
end;

procedure TForm1.RESIZE_MOVE(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
    If Moving then Form1.Width:=Form1.Width+X-XPos;
    If Moving then Form1.Height:=Form1.Height+Y-YPos;
end;


procedure TForm1.RESIZE_WND_DOWN(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
    begin
    if WindowState <> wsMaximized then
    begin
      XPos:=X;
      YPos:=Y;
      Moving := True;
      Form1.Update;
    end;
  end;
end;

procedure TForm1.RESIZE_WND_UP(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  Moving := False;
end;
{$IFDEF undef}{$ENDREGION}{$ENDIF}
//**********************************************
//                  Form changed control
//**********************************************
//================COLOR==CHANGE=================


procedure TForm1.DescriptionTextChanged(Sender: TObject);
var
  memo : TMemo;
  count : Integer;
begin
  memo := Sender as TMemo;
  count := Length(memo.Text);
  countText.Caption := 'Введено ' + IntToStr(count);
end;

procedure TForm1.BrowseBTNClick(Sender: TObject);
var
  openDialog : topendialog;    // Open dialog variable
  MaxSize : Real;

begin
  // Create the open dialog object - assign to our open dialog variable
  openDialog := TOpenDialog.Create(self);
  // Set up the starting directory to be the current one
  openDialog.InitialDir := GetCurrentDir;
  if (Length(EDIT_EDIT_BROWSE.Text) > 0) and FileExists(EDIT_EDIT_BROWSE.Text) then
     openDialog.InitialDir := ExtractFileDir(EDIT_EDIT_BROWSE.Text);
  // Only allow existing files to be selected
  openDialog.Options := [ofFileMustExist];
  // Allow only .dpr and .pas files to be selected
  // Select pascal files as the starting filter type
  openDialog.FilterIndex := 2;
  // Display the open file dialog
  if openDialog.Execute then controlOfFile(openDialog.FileName);
  openDialog.Free;
end;

procedure TForm1.controlOfFile(path:string);
 var
 MaxSize :Real;
 sr : TSearchRec;
begin

  MaxSize := 20971520;//20MB
//   MaxSize := 15728640;//15MB
//   MaxSize := 10485760;//10MB
//   MaxSize := 5242880;//5MB
   if FindFirst(path, faAnyFile, sr ) = 0 then
   begin
      if (Int64(sr.FindData.nFileSizeHigh) shl Int64(32) + Int64(sr.FindData.nFileSizeLow)) < MaxSize then
      begin
        EDIT_EDIT_BROWSE.Text:= path;
        EDIT_LB_BROWSE.Caption :=path;
      end
      else
      begin
        EDIT_EDIT_BROWSE.Text:= '';
        EDIT_LB_BROWSE.Caption := '';
        ShowMessage('Вы не можете прикрепить файл больше чем ' + FloatToStr(MaxSize/1048576) + 'МБ');
      end;  
      FindClose(sr) ;
   end;
end;  

procedure TForm1.sendBTNClick(Sender: TObject);
var
  count :Integer;
begin
    count := Length(Form1.EDIT_EDIT_DEPART.Text) * Length(Form1.EDIT_EDIT_FIO.Text) * Length(Form1.EDIT_EDIT_TABNO.Text) * Length(Form1.EDIT_EDIT_TEL.Text) * Length(Form1.CMB_INVNO.Text) * Length(MEMO_MEMO.Text);
    if TECHMACH_BTN.Checked then
        count := count * Length(Form1.CMB_ISPOL.Text);
    if count = 0 then
        ShowMessage('Заполните все поля')
  else
    Register();
end;

procedure TForm1.Change_TYPE_REQ(Sender: TObject);
var
buffer: array[0..MAX_COMPUTERNAME_LENGTH + 1] of Char;
nSize : Cardinal;
departSQL_Depend,usingAtPlace,spisDB : String;

begin
     spisDB := '';
     departSQL_Depend := convertDepart_forSQL();  //return string for SQL query to determinate list of machine tools depend of current user depart

     if (TECHMACH_BTN.Checked) and (Length(departSQL_Depend)>1) then
     begin // to activate PDS type request
       CMB_INVNO.Text := 'Загрузка...';
       EDIT_LB_LEAVE(LB_INVNO);
       EDIT_TEXT_CHANGE(CMB_INVNO);

       PDSPNL.Visible := true;
       CMB_ISPOL.Items.Clear;
       CMB_ISPOL.Items := getUserIspolnitelList();
       CMB_ISPOL.ItemIndex :=0;
       CMB_INVNO.Style := csOwnerDrawFixed;
       EDIT_LB_LEAVE(LB_ISPOL);
       EDIT_TEXT_CHANGE(CMB_ISPOL);

       if Length(departSQL_Depend) < 1 then //if depert can't determinate. wrote name of depart manually
       begin
         ADOQuery_getInvNo1.Active := True;  // injection start
         ADOQuery_getInvNo1.Recordset.MoveFirst;
         while Not ADOQuery_getInvNo1.Recordset.EOF do
         begin
           usingAtPlace := ADOQuery_getInvNo1.Recordset.Fields['Место эксплуатации'].Value;
           spisDB := spisDB + usingAtPlace + ' ';
           if usingAtPlace = EDIT_EDIT_DEPART.Text then
           begin
                departSQL_Depend := usingAtPlace;
           end;
           ADOQuery_getInvNo1.Recordset.MoveNext;

         end;
         ADOQuery_getInvNo1.Active := false; // injection stop
       end;


       ADOQueryGETINVNO.SQL.Clear;
       ADOQueryGETINVNO.SQL.Add( 'SELECT Книга.[Инвентарный №] FROM Книга WHERE (((Книга.[Место эксплуатации])=''' +
                            departSQL_Depend + ''')) ORDER BY Книга.[Инвентарный №];');
       ADOQueryGETINVNO.Active := True;  // injection start
       if ADOQueryGETINVNO.Recordset.RecordCount <> 0 then
       begin
        ADOQueryGETINVNO.Recordset.MoveFirst;
        CMB_INVNO.Items.Clear;

         while Not ADOQueryGETINVNO.Recordset.EOF do
         begin
           CMB_INVNO.Items.Add(ADOQueryGETINVNO.Recordset.Fields['Инвентарный №'].Value);
           ADOQueryGETINVNO.Recordset.MoveNext;
         end;
        ADOQueryGETINVNO.Active := False;  // injection start
        CMB_INVNO.ItemIndex := 0;
        INvNoChange(nil);
       end
     end
     else
     begin // to activate OASU type reuest
       TECHMACH_BTN.Checked := False;
       OASU_RBTN.Checked := True;
       TECHMACH_BTN.ShowHint := False;
       if Length(departSQL_Depend)<1 then
       begin
        TECHMACH_BTN.Hint := 'Вы не можете создавать данный тип заявки';
        TECHMACH_BTN.ShowHint := True;
       end;  
       PDSPNL.Visible := false;
       CMB_INVNO.Style := csSimple;
       nSize :=1024;
       GetComputerName(@buffer, nSize);
       Form1.CMB_INVNO.Text := StrPas(buffer);
       EDIT_TEXT_CHANGE(CMB_INVNO);
     end;
end;

procedure TForm1.INvNoChange(Sender: TObject);
        var val : String;
begin
        EDIT_TEXT_CHANGE(CMB_INVNO);
        if TECHMACH_BTN.Checked then begin
          ADOQueryInvNoChange.SQL.Clear;
          ADOQueryInvNoChange.SQL.Add('SELECT Книга.[Наименование оборудования] FROM Книга WHERE (((Книга.[Инвентарный №])='''+
          CMB_INVNO.Text + '''));');
          ADOQueryInvNoChange.Active := True;
          val := ADOQueryInvNoChange.Recordset.Fields['Наименование оборудования'].Value;
          invNoName.Caption := val;
          invNoName.Hint := val;
          ADOQueryInvNoChange.Active := False;
        end;
end;

procedure TForm1.ReportBTNClick(Sender: TObject);
var
  MSWord, Doc, Tables, Table: Variant;
  filename, alias, problem, probWord: String;
  j, i : Integer;
begin
    try
      MsWord := CreateOleObject('Word.Application');
      MsWord.Visible := False;
    except
      Exception.Create('Error');
    end;
  S4.OpenQuery('select ALIAS from ARCHIVES where ARCHIVE_ID = ' + IntToStr(OASU.ArchID));
  alias := S4.QueryFieldByName('ALIAS');
  S4.CloseQuery;
  S4.OpenQuery('select p. doc_id, d.DESIGNATIO, d.CREATEDATE, p.PROBLEM, p.OTMETKA, p.ISPOLNITEL, p.DATEISP, p.TIMEISP, p.USERNOTE from ' +
  alias + ' p left join  doclist d on p.doc_id = d.doc_id where FIO like ''' + EDIT_EDIT_FIO.Text +
  '%'' order by d.CREATEDATE');

  filename := '\\sql-main\IM\Application\Report Templates\Отчет по заявкам в сервис ОАСУ.dot';

  MSWord.Documents.Open(filename);
  Doc := MSWord.Documents.Application.ActiveDocument;
  Tables := Doc.Tables;
  Table := Tables.Item(1);
  Table.Cell(1{Column}, 2 {Row}).Range.Text := EDIT_EDIT_FIO.Text;
  S4.QueryGoFirst;
  j  := 4;
  while S4.QueryEOF = 0 do
  begin
        if j > 4 then
        Begin
          {Append row to table}
          MsWord.Selection.Tables.Item(1).Rows.Item(MsWord.Selection.Tables.Item(1).Rows.Count).Select;
          MSWord.Selection.InsertRowsBelow;
        end;
        Table.Cell(j, 1).Range.Text := S4.QueryFieldByName( 'DESIGNATIO' );
        Table.Cell(j, 2).Range.Text := S4.QueryFieldByName( 'CREATEDATE' );
        problem := S4.QueryFieldByName( 'PROBLEM' );
        if Length(problem) > 94 then
        begin
          For i := 0 to Length(problem) do
                probWord := probWord + problem[i];
                if ((i Mod 94) = 0) and (i <> 0) then
                begin
                   probWord := probWord + #13;
                end;
        problem := probWord;
        end;

        Table.Cell(j, 3).Range.Text := problem;
        Table.Cell(j, 4).Range.Text := S4.QueryFieldByName( 'ISPOLNITEL' );
        Table.Cell(j, 5).Range.Text := S4.QueryFieldByName( 'DATEISP' ) + ' ' + S4.QueryFieldByName( 'TIMEISP' );
        Table.Cell(j, 6).Range.Text := S4.QueryFieldByName( 'USERNOTE' );
        S4.QueryGoNext();
        Inc(j);
  end;
  S4.CloseQuery;
  MsWord.Visible := True;
  end;

procedure TForm1.ColorPicker(Sender: TObject);
begin

  if (sender as TMenuItem).Name = 'BackgroundColor1'  then
  begin
    if dlgColor1.Execute then
    begin
      BackgroungColor := dlgColor1.Color;
      changeBackgroungColor(BackgroungColor);
    end
  end
  else if (sender as TMenuItem).Name = 'ResetColors1' then
  begin
        BackgroungColor := TColor($f2f2f2);
        BorderColor := {R}22 + {G}52 * 256 + {B}122 * 256 * 256;

        changeBackgroungColor(BackgroungColor);
        changeBorderColor(BorderColor);
        Form1.Height := Form1.Constraints.MinHeight;
        Form1.Width := Form1.Constraints.MinWidth;
        OASU_RBTN.Checked := True;
        TECHMACH_BTN.Checked := not(OASU_RBTN.Checked);
  end
  else
  begin
      if dlgColor1.Execute then
      begin
        BorderColor := dlgColor1.Color;
        changeBorderColor(BorderColor);
      end;
  end;
end;


procedure TForm1.BUTTON_CLICK_DOWN(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var     lb :TLabel;
        panel : TPanel;
        shape : TShape;
begin
    lb := sender as TLabel;
    panel :=  lb.Parent as TPanel;
    shape := panel.Controls[0] as TShape;

    shape.Pen.Color := BackgroungColor;
    
    panel.Color := increaseBrightness(BackgroungColor, -70);
    lb.Font.Color := BackgroungColor;
end;

procedure TForm1.BUTTON_CLICK_CLOSE(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  BUTTON_ENTER(Sender);
  CLOSE_WINDOW(Sender, Button, Shift, X, Y);
end;

procedure TForm1.BUTTON_REPORT_UP(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  BUTTON_ENTER(Sender);
  ReportBTNClick(Sender);
end;

procedure TForm1.BUTTON_REGISTER_UP(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  sendBTNClick(Sender);
  BUTTON_ENTER(Sender);
end;

procedure TForm1.BUTTON_CLICK_UP(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  BrowseBTNClick(Sender);
  BUTTON_ENTER(Sender);
end;

procedure TForm1.GetPrintScreen(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
const
  CAPTUREBLT = $40000000;
var  ScreenBM : TBitMap;
      SendJPG : TJPEGImage;
      W, H ,Offset: Integer;
      path: string;
begin
  Form1.Visible := False;
  ScreenBM := TBitMap.Create;
  SendJPG := TJPEGImage.Create;

  Offset := 0;
  if Form1.Left > Screen.Width then
  begin
    W := Screen.Monitors[1].Width;
    H := Screen.Monitors[1].Height;
    Offset := Screen.Width;
  end
  else
  begin
    W := Screen.Width;
    H := Screen.Height;
  end;

  ScreenBM.Width := W;
  ScreenBM.Height := H;

  SendJPG.CompressionQuality := 100;  //степень сжатия от 1 до 100
  SendJPG.Compress;
  BitBlt(ScreenBM.Canvas.Handle, 0, 0, W, H,
GetDC(0), Offset, 0, SRCCopy or CAPTUREBLT);
  SendJPG.Assign(ScreenBM);
  CreateDir(GetEnvironmentVariable('Temp')+'\ScrenShotForRequest');
  path:=GetEnvironmentVariable('Temp')+'\ScrenShotForRequest\'+FormatDateTime('ddmmyyyy_hhnn', Now) + '.jpg';

  SendJPG.SaveToFile(path);
  EDIT_EDIT_BROWSE.Text := path;
  EDIT_LB_BROWSE.Caption := path;
  Form1.Visible := True;

  SendJPG.Free;
  ScreenBM.Free;
end;

procedure TForm1.GetVideo(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
  var frm : TForm2;
result : Integer;
begin

  Form1.Visible := False;
  Application.CreateForm(TForm2, Form2);
  frm := TForm2.Create(Self);
  result := frm.ShowModal;
  frm.Free;
  if result = 1 then
    controlOfFile( GetEnvironmentVariable('Temp')+'\VideoForRequest\Movie.avi');
  Form1.Visible := True;
end;


end.

