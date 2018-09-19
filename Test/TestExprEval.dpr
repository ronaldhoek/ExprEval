program TestExprEval;

uses
  Vcl.Forms,
  MainU in 'MainU.pas' {Form1},
  prExprFuncDB in '..\Source\prExprFuncDB.pas',
  prExprFunc in '..\Source\prExprFunc.pas',
  prExpr in '..\Source\prExpr.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
