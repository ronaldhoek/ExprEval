unit MainU;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ActnList,
  JvAppStorage, JvAppIniStorage, JvComponentBase, JvFormPlacement, prExpr;

type
  TForm1 = class(TForm)
    mmExpression: TMemo;
    lblExpression: TLabel;
    mmResult: TMemo;
    lblResult: TLabel;
    btnRun: TButton;
    ActionList1: TActionList;
    actnExecute: TAction;
    FormStorage: TJvFormStorage;
    appStorage: TJvAppIniFileStorage;
    procedure FormCreate(Sender: TObject);
    procedure actnExecuteExecute(Sender: TObject);
  private
    function IdentifierFunction(const Identifier: String; ParameterList:
        IParameterList): IValue;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

uses
  prExprFunc, Data.DB, System.TypInfo;

procedure TForm1.FormCreate(Sender: TObject);
begin
  appStorage.FileName := ChangeFileExt(ExtractFileName(Application.ExeName), '.ini');
end;

procedure TForm1.actnExecuteExecute(Sender: TObject);
var
  s: String;
  _Result: IValue;
begin
  mmResult.Lines.Clear;
  for s in mmExpression.Lines do
    if Trim(s) > '' then
  begin
    _Result := ExprFunc.Create(s, IdentifierFunction);
    mmResult.Lines.Add( VarToStr(_Result.AsVariant) + ' : ' + GetEnumName(TypeInfo(TExprType), Ord(_Result.ExprType)));
  end;
end;

function TForm1.IdentifierFunction(const Identifier: String; ParameterList:
    IParameterList): IValue;
begin
  // Database typen
  if not CheckEnumeratedVal(TypeInfo(TFieldType), Identifier, Result) then
    Result := nil;
end;

end.
