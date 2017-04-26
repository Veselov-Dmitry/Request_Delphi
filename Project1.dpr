program Project1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  S4_TLB in 'C:\Program Files (x86)\Borland\Delphi7\Imports\S4_TLB.pas',
  Unit2 in 'Unit2.pas' {Form2};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Заявка';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
