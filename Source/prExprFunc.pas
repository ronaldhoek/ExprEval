unit prExprFunc;

interface

uses
  System.Classes, System.Contnrs, System.TypInfo, prExpr;

type
  //  ICalcBoolExpr = interface(IInterface)
  //    function AsBoolean: Boolean;
  //  end;
  //
  //  TCalcBoolVWFunc = function (const Voorwaarde: string): Boolean of object;

  TExprResult = record
    Valid: Boolean;
    Value: IValue;
  end;

  // IValue uitbreiding met 'IsNull' routine
  INullableValue = interface(IValue)
    ['{F1CE3074-A7E7-4F11-8B30-7AD95E7484A1}']
    function IsNull: Boolean;
  end;

  ICalcExpression = interface(IInterface)
    ['{0E75D69A-8DC2-4058-96A8-2DD752E3ECE0}']
    procedure AddParam(const Name: string; const Value: Variant);
    procedure ClearCalculated;
    procedure ClearParams;
    function GetExpression: string;
    function GetOutput: TExprResult;
    function GetOutputBool: Boolean;
    function GetResult: IValue;
    function GetValid: Boolean;
    function RemoveParam(const Name: string): Boolean;
    procedure SetExpression(const Expr: string);
    function UpdateParam(const Name: string; const Value: Variant): Boolean;
    property Expression: string read GetExpression;
    property Output: TExprResult read GetOutput;
    property OutputBool: Boolean read GetOutputBool;
    property Result: IValue read GetResult;
    property Valid: Boolean read GetValid;
  end;

  ExprFunc = class(TObject)
  public
    class function Create(const Expression: string; IdentifierFunction:
            TIdentifierFunction): IValue;
    class function ExecBool(const Expression: string; IdentifierFunction:
            TIdentifierFunction): Boolean;
    class function Execute(const Expression: string; IdentifierFunction:
            TIdentifierFunction): TExprResult;
    class function GetIdentifier(const Identifier: String; ParameterList:
            IParameterList): IValue; overload;
    class function GetIdentifier(const Identifier: String; ParameterList:
            IParameterList; out Value: IValue): Boolean; overload;
    class function NewExpr(const Expression: string; IdentifierFunction:
            TIdentifierFunction): ICalcExpression;
  end;

  ///  <summary>
  ///  Basis class om zelf nieuwe functies te maken en toe te voegen aan de
  ///  interne lijst van ICalcExpression.
  ///  Indien dit niet gewenst is, kan altijd o.b.v. 'IdentifierFunction' zelf
  ///  een callback worden gemaakt, met een implementatie voor de functie.
  ///  </summary>
  TprFunction = class(TFunction)
  protected
    class function IdentifierName: string; virtual; abstract;
    class procedure Register;
    class procedure Unregister;
  end;

  (*
    Speciale 'losse' waarde
  *)

  // Basis implementatie voor 'INullableValue'
  TNullableExpression = class(TExpression, INullableValue)
  public
    function IsNull: Boolean; virtual; abstract;
  end;

  TNullValue = class(TNullableExpression)
  private
    FExprType: TExprType;
  public
    constructor Create(aExprType: TExprType);
    function AsBoolean: TprBoolean; override;
    function AsFloat: TprFloat; override;
    function AsInteger: TprInt; override;
    function AsObject: TObject; override;
    function AsString: TprString; override;
    function AsVariant: Variant; override;
    function ExprType: TExprType; override;
    function IsNull: Boolean; override;
  end;

  TVariantValue = class(TNullableExpression)
  private
    FExprType: TExprType;
    FExprTypeInit: Boolean;
    FValue: Variant;
  public
    constructor Create(aValue: Variant);
    function AsBoolean: TprBoolean; override;
    function AsFloat: TprFloat; override;
    function AsInteger: TprInt; override;
    function AsString: TprString; override;
    function AsVariant: Variant; override;
    function ExprType: TExprType; override;
    function IsNull: Boolean; override;
  end;

  (*
    Callback for casting 'enum' values to TExprType
  *)

  TEnumToExprTypeCastProc = function(aValue: IValue): TExprType;

  procedure RegisterExprTypeCastProc(aTypeInfo: PTypeInfo; aProc:
      TEnumToExprTypeCastProc);
  procedure UnregisterExprTypeCastProc(aTypeInfo: PTypeInfo; aProc:
      TEnumToExprTypeCastProc);

  (*
    Other helper functions
  *)

  function TryValueToExprType(V: IValue;  out aType: TExprType): Boolean;
  function ValueToExprType(V: IValue): TExprType;

  function TryVarTypeToExprType(const VType: TVarType; out aType: TExprType): Boolean; overload;
  function TryVarTypeToExprType(const V: Variant; out aType: TExprType): Boolean; overload;

  function VarTypeToExprType(const VType: TVarType): TExprType; overload;
  function VarTypeToExprType(const V: Variant): TExprType; overload;

implementation

uses
  System.SysUtils, System.Variants, System.Math, System.StrUtils;

const
  sFmtCalcTypeError = 'Cannot add varianttype %d as Calc-parameter';

type
  TExprTypeCastProc = record
    TypeInfo: PTypeInfo;
    Proc: TEnumToExprTypeCastProc;
  end;

var
  _ExprTypeCastProcs: array of TExprTypeCastProc;


  (*
    Hulproutines om te bepalen of variant 'leeg' is
  *)

function VarIsEmptyOrNull(const V: Variant): Boolean;
begin
  Result := FindVarData(V)^.VType in [varEmpty, varNull];
end;

function ValueIsNull(V: IValue): Boolean;
var
  Obj: INullableValue;
begin
  // if Supports(V, INullableValue, Obj) then
  if V = nil then
    Result := True
  else if V.QueryInterface(INullableValue, Obj) = S_OK then
    Result := Obj.IsNull
  else
    Result := VarIsEmptyOrNull(V.AsVariant);
end;

(*
  ExprType casting procedures
*)

function IndexOfExprTypeCastProc(aTypeInfo: PTypeInfo): Integer;
var
  I: Integer;
begin
  for I := 0 to Length(_ExprTypeCastProcs) - 1 do
    if _ExprTypeCastProcs[I].TypeInfo = aTypeInfo then
      Exit(I);
  Result := -1;
end;

procedure RegisterExprTypeCastProc(aTypeInfo: PTypeInfo; aProc:
    TEnumToExprTypeCastProc);
var
  I: Integer;
begin
  if aTypeInfo = nil then Exit;

  if IndexOfExprTypeCastProc(aTypeInfo) <> -1 then
    raise EExpression.CreateFmt('TypeInfo %s already registered for ExprType castings', [aTypeInfo.Name]);

  // Find empty spot
  I := IndexOfExprTypeCastProc(nil);
  if I = -1 then
  begin
    I := Length(_ExprTypeCastProcs);
    SetLength(_ExprTypeCastProcs, I + 1);
  end;
  _ExprTypeCastProcs[I].TypeInfo := aTypeInfo;
  _ExprTypeCastProcs[I].Proc := aProc;
end;

procedure UnregisterExprTypeCastProc(aTypeInfo: PTypeInfo; aProc:
    TEnumToExprTypeCastProc);
var
  I: Integer;
begin
  I := IndexOfExprTypeCastProc(aTypeInfo);
  if I >= 0 then
  begin
    _ExprTypeCastProcs[I].TypeInfo := nil;
    _ExprTypeCastProcs[I].Proc := nil;
  end;
end;

function TryValueToExprType(V: IValue;  out aType: TExprType): Boolean;
var
  I: Integer;
begin
  if V.ExprType in [ttInteger, ttEnumerated] then
  begin
    // Parameter is 'TExprType'
    if V.TypeInfo = System.TypeInfo(TExprType) then
    begin
      aType := TExprType(V.AsInteger);
      Result := True;
    end else
    begin
      // Find cast procedure for type
      I := IndexOfExprTypeCastProc(V.TypeInfo);
      if I >= 0 then
      begin
        aType := _ExprTypeCastProcs[I].Proc(V);
        Result := True;
      end else
        Result := False;
    end;
  end else
    Result := False;
end;

function ValueToExprType(V: IValue): TExprType;
const
  sErrCannotCastToExprType = 'Cannot cast %s to TExprType';
begin
  if not TryValueToExprType(V, Result) then
    raise EExpression.CreateFmt(sErrCannotCastToExprType, [V.TypeName]);
end;

(*
  Publieke routines
*)

function TryVarTypeToExprType(const VType: TVarType; out aType: TExprType): Boolean;
begin
  Result := True;
  case VType and varTypeMask of
    varOleStr, varString, varUString, varStrArg, varUStrArg:
      aType := ttString;
    varByte, varShortInt, varWord, varSmallInt, varInteger, varLongWord, varInt64, varUInt64:
      aType := ttInteger;
    varSingle, varDouble, varCurrency, varDate:
      aType := ttFloat;
    varBoolean:
      aType := ttBoolean;
  else
    Result := False;
  end;
end;

function TryVarTypeToExprType(const V: Variant; out aType: TExprType): Boolean;
begin
  Result := TryVarTypeToExprType(VarType(V), aType);
end;

function VarTypeToExprType(const VType: TVarType): TExprType;
begin
  if not TryVarTypeToExprType(VType, Result) then
    raise EExpression.CreateFmt('Variant type %d cannot be mached to an expression type', [VType]);
end;

function VarTypeToExprType(const V: Variant): TExprType; overload;
begin
  Result := VarTypeToExprType(VarType(V));
end;

type
  (*
    Expression evaluatie object
  *)

  TParamObject = class(TObject)
  private
    FValue: IValue;
  public
    constructor Create(aValue: IValue);
    property Value: IValue read FValue;
  end;

  TICalcExpression = class(TInterfacedObject, ICalcExpression)
  private
    FExpression: string;
    FExternalCalc: TIdentifierFunction;
    FOutput: TExprResult;
    FParameterList: TStringList;
    procedure AddParam(const Name: string; const Value: Variant);
    procedure ClearCalculated;
    procedure ClearParams;
    function GetExpression: string;
    function GetInternalParameter(const Identifier: string): IValue;
    function GetOutput: TExprResult;
    function GetOutputBool: Boolean;
    function GetParamObjectValue(const Value: Variant): IValue;
    function GetResult: IValue;
    function GetValid: Boolean;
    function RemoveParam(const Name: string): Boolean;
    procedure SetExpression(const Expr: string);
    function UpdateParam(const Name: string; const Value: Variant): Boolean;
    function _InternalGetIdentifier(const Identifier: String; ParameterList:
            IParameterList): IValue;
  public
    constructor Create(const Expr:string; IdentifierFunction:
            TIdentifierFunction);
    destructor Destroy; override;
  end;

  (*
    Functie registratie objecten
  *)

  TprFunctionClass = class of TprFunction;

  TprFunctionList = class(TClassList)
  private
    FAddingName: Boolean;
    FNameList: TStrings;
    procedure AddClassName(aClass: TClass);
    procedure RemoveClassName(aClass: TClass);
  protected
    function GetClassName(aClass: TClass): string;
    procedure Notify(Instance: Pointer; Action: TListNotification); override;
  public
    constructor Create;
    destructor Destroy; override;
    function CreateFunction(const Identifier: string; const ParameterList:
            IParameterList): IValue; overload;
    function CreateFunction(const Identifier: string; const ParameterList:
            IParameterList; out Value: IValue): Boolean; overload;
  end;

  (*
    Basis klassen
  *)

  TBaseMinMax = class(TprFunction)
  private
    FCommonType: TExprType;
    function IgnoreCase: Boolean;
  protected
    function TestParameters: Boolean; override; final;
  public
    function ExprType: TExprType; override; final;
  end;

  (*
    Interne functies
  *)

  TMin = class sealed(TBaseMinMax)
  protected
    class function IdentifierName: string; override;
  public
    function AsBoolean: Boolean; override;
    function AsFloat: Extended; override;
    function AsInteger: Integer; override;
    function AsString: string; override;
  end;

  TMax = class sealed(TBaseMinMax)
  protected
    class function IdentifierName: string; override;
  public
    function AsBoolean: Boolean; override;
    function AsFloat: Extended; override;
    function AsInteger: Integer; override;
    function AsString: string; override;
  end;

  TArray = class(TprFunction)
  protected
    class function IdentifierName: string; override;
  public
    function AsObject: TObject; override;
    function ExprType: TExprType; override;
    function IndexOf(Value: IValue): Integer;
    function InList(Value: IValue): Boolean;
    function TypeInfo: PTypeInfo; override;
  end;

  TInArray = class(TprFunction)
  protected
    class function IdentifierName: string; override;
    function TestParameters: Boolean; override;
  public
    function AsBoolean: TprBoolean; override;
    function ExprType: TExprType; override;
  end;

  TCaseOf = class(TprFunction)
  protected
    class function IdentifierName: string; override;
    function TestParameters: Boolean; override;
  public
    function AsInteger: TprInt; override;
    function ExprType: TExprType; override;
  end;

  TBaseSelect = class(TprFunction, INullableValue)
  private
    function Selected(ReturnDefault: Boolean = True): IValue;
  protected
    function GetDefault: IValue; virtual; abstract;
    function GetIndex: Integer; virtual; abstract;
    function SelectArr: TArray; virtual; abstract;
  public
    function AsBoolean: TprBoolean; override;
    function AsFloat: TprFloat; override;
    function AsInteger: TprInt; override;
    function AsObject: TObject; override;
    function AsString: TprString; override;
    function AsVariant: Variant; override;
    function IsNull: Boolean;
  end;

  TSelect = class(TBaseSelect)
  protected
    function GetDefault: IValue; override;
    function GetIndex: Integer; override;
    class function IdentifierName: string; override;
    function SelectArr: TArray; override;
    function TestParameters: Boolean; override;
  public
    function ExprType: TExprType; override;
  end;

  TCaseSelect = class(TBaseSelect)
  private
    function CaseArr: TArray;
  protected
    function GetDefault: IValue; override;
    function GetIndex: Integer; override;
    class function IdentifierName: string; override;
    function SelectArr: TArray; override;
    function TestParameters: Boolean; override;
  public
    function ExprType: TExprType; override;
  end;

  TIsNull = class(TprFunction)
  protected
    class function IdentifierName: string; override;
    function TestParameters: Boolean; override;
  public
    function AsBoolean: Boolean; override;
    function ExprType: TExprType; override;
  end;

  TIfNull = class(TprFunction)
  private
    function Rex: IValue;
  protected
    class function IdentifierName: string; override;
    function TestParameters: Boolean; override;
  public
    function AsBoolean: Boolean; override;
    function ExprType: TExprType; override;
    function AsFloat: Extended; override;
    function AsInteger: Integer; override;
    function AsString: string; override;
    function AsVariant: Variant; override;
  end;

  TNull = class(TprFunction, INullableValue)
  protected
    class function IdentifierName: string; override;
    function TestParameters: Boolean; override;
  public
    function AsBoolean: TprBoolean; override;
    function AsFloat: TprFloat; override;
    function AsInteger: TprInt; override;
    function AsObject: TObject; override;
    function AsString: TprString; override;
    function AsVariant: Variant; override;
    function ExprType: TExprType; override;
    function IsNull: Boolean;
  end;

  (*
    Extra wiskundige routine routines
  *)

  TInRange = class(TprFunction)
  protected
    class function IdentifierName: string; override;
    function TestParameters: Boolean; override;
  public
    function AsBoolean: Boolean; override;
    function ExprType: TExprType; override;
  end;

  (*
    Datum/tijd routines
  *)

  TDateInfo = record
    Year, Month, Day: Word;
  end;

  TTimeInfo = record
    Hour, Minutes, Secondes, Miliseconds: Word;
  end;

  TBaseDateTime = class(TprFunction)
  private
    function GetDateInfo: TDateInfo;
    function GetTimeInfo: TTimeInfo;
    function GetValue: TDateTime;
  protected
    class function IdentifierName: string; override;
    function TestParameters: Boolean; override;
    property DateInfo: TDateInfo read GetDateInfo;
    property TimeInfo: TTimeInfo read GetTimeInfo;
    property Value: TDateTime read GetValue;
  end;

  TDateYear = class(TBaseDateTime)
  public
    function AsInteger: TprInt; override;
    function ExprType: TExprType; override;
  end;

  TDateMonth = class(TBaseDateTime)
  public
    function AsInteger: TprInt; override;
    function ExprType: TExprType; override;
  end;

  TDateDay = class(TBaseDateTime)
  public
    function AsInteger: TprInt; override;
    function ExprType: TExprType; override;
  end;

  TTimeHour = class(TBaseDateTime)
  public
    function AsInteger: TprInt; override;
    function ExprType: TExprType; override;
  end;

  TTimeMinute = class(TBaseDateTime)
  public
    function AsInteger: TprInt; override;
    function ExprType: TExprType; override;
  end;

  TTimeSecond = class(TBaseDateTime)
  public
    function AsInteger: TprInt; override;
    function ExprType: TExprType; override;
  end;

  TTimeMilisecond = class(TBaseDateTime)
  public
    function AsInteger: TprInt; override;
    function ExprType: TExprType; override;
  end;

  TEncodeDate = class(TprFunction)
  protected
    class function IdentifierName: string; override;
    function TestParameters: Boolean; override;
  public
    function AsFloat: TprFloat; override;
    function AsString: TprString; override;
    function ExprType: TExprType; override;
  end;

  TEncodeTime = class(TprFunction)
  protected
    class function IdentifierName: string; override;
    function TestParameters: Boolean; override;
  public
    function AsFloat: TprFloat; override;
    function AsString: TprString; override;
    function ExprType: TExprType; override;
  end;

var
  _FunctionList: TprFunctionList;

{ ExprFunc }
{
*********************************** ExprFunc ***********************************
}
class function ExprFunc.Create(const Expression: string; IdentifierFunction:
        TIdentifierFunction): IValue;
begin
  with TICalcExpression.Create(Expression, IdentifierFunction) do
  try
    Result := CreateExpression(Expression, _InternalGetIdentifier);
  finally
    Free;
  end;
end;

class function ExprFunc.ExecBool(const Expression: string; IdentifierFunction:
        TIdentifierFunction): Boolean;
begin
  Result := NewExpr(Expression, IdentifierFunction).OutputBool;
end;

class function ExprFunc.Execute(const Expression: string; IdentifierFunction:
        TIdentifierFunction): TExprResult;
begin
  Result := NewExpr(Expression, IdentifierFunction).Output;
end;

class function ExprFunc.GetIdentifier(const Identifier: String; ParameterList:
        IParameterList): IValue;
begin
  Result := _FunctionList.CreateFunction(Identifier,  ParameterList);
end;

class function ExprFunc.GetIdentifier(const Identifier: String; ParameterList:
        IParameterList; out Value: IValue): Boolean;
begin
  Result := _FunctionList.CreateFunction(Identifier,  ParameterList, Value);
end;

class function ExprFunc.NewExpr(const Expression: string; IdentifierFunction:
        TIdentifierFunction): ICalcExpression;
begin
  Result := TICalcExpression.Create(Expression, IdentifierFunction);
end;

{ TprFunction }
{
********************************* TprFunction **********************************
}
class procedure TprFunction.Register;
begin
  if Assigned(_FunctionList) then
    _FunctionList.Add(Self);
end;

class procedure TprFunction.Unregister;
begin
  if Assigned(_FunctionList) then
    _FunctionList.Remove(Self);
end;

{ TParamObject }
{
********************************* TParamObject *********************************
}
constructor TParamObject.Create(aValue: IValue);
begin
  FValue := aValue;
end;

{ TICalcExpression }
{
******************************* TICalcExpression *******************************
}
procedure TICalcExpression.AddParam(const Name: string; const Value: Variant);
begin
  FParameterList.AddObject(Name, TParamObject.Create(GetParamObjectValue(Value)));
  ClearCalculated; // When modifying parameterlist, rebuiding is required
end;

procedure TICalcExpression.ClearCalculated;
begin
  FOutput.Valid := False;
  FOutput.Value := nil;
end;

procedure TICalcExpression.ClearParams;
begin
  FParameterList.Clear;
  ClearCalculated;
end;

constructor TICalcExpression.Create(const Expr:string; IdentifierFunction:
        TIdentifierFunction);
begin
  inherited Create;
  FExternalCalc := IdentifierFunction;
  FExpression := Trim(Expr);
  FParameterList := TStringList.Create(True);
  FParameterList.Sorted := True;
  FParameterList.Duplicates := dupError;
  //  FParameterList.CaseSensitive := False;
end;

destructor TICalcExpression.Destroy;
begin
  FParameterList.Free;
  inherited;
end;

function TICalcExpression.GetExpression: string;
begin
  Result := FExpression;
end;

function TICalcExpression.GetInternalParameter(const Identifier: string):
        IValue;
var
  i: Integer;
begin
  // Eerst True/False omzetten
  if Identifier='TRUE' then
    Result:=TBooleanLiteral.Create(True)
  else if Identifier='FALSE' then
    Result:=TBooleanLiteral.Create(False)
  else If FParameterList.Find(Identifier, i) then // Dan Parameterlijst doorlopen
    Result := TParamObject(FParameterList.Objects[i]).Value
  else
    Result := nil;
end;

function TICalcExpression.GetOutput: TExprResult;
begin
  //  Execute;
  if FOutput.Value = nil then
  begin
    FOutput.Valid := FExpression <> '';
    if FOutput.Valid then
    begin
      FOutput.Value := CreateExpression(FExpression, _InternalGetIdentifier);
      FOutput.Valid := Assigned(FOutput.Value);
    end;
  end;
  Result := FOutput;
end;

function TICalcExpression.GetOutputBool: Boolean;
begin
  // Function for calling expressions which need to return boolean
  // - When expression is empty the result is True
  // - When the result can't be read as boolean the return value is False
  Result := Length(FExpression) = 0;
  if not Result then
    with GetOutput do
      Result := Valid and
                Value.CanReadAs(ttBoolean) and
                Value.AsBoolean;
end;

function TICalcExpression.GetParamObjectValue(const Value: Variant): IValue;
begin
  Result := TVariantValue.Create(Value);
  if Result.ExprType = ttObject then
    Raise EExpression.CreateFmt(sFmtCalcTypeError, [VarType(Value)]);
end;

function TICalcExpression.GetResult: IValue;
begin
  Result := GetOutput.Value;
end;

function TICalcExpression.GetValid: Boolean;
begin
  Result := GetOutput.Valid;
end;

function TICalcExpression.RemoveParam(const Name: string): Boolean;
var
  i: Integer;
begin
  Result := FParameterList.Find(Name, i);
  if Result then
  begin
    FParameterList.Delete(i);
    ClearCalculated;
  end;
end;

procedure TICalcExpression.SetExpression(const Expr: string);
begin
  // Zet nieuwe expressie
  FExpression := Trim(Expr);
  // "Berekend" vrijgeven
  ClearCalculated;
end;

function TICalcExpression.UpdateParam(const Name: string; const Value:
    Variant): Boolean;
var
  i: Integer;
  _NewType: TExprType;
  _Obj: TParamObject;
begin
  Result := FParameterList.Find(Name, i);
  if Result then
  begin
    // Bepaal type nieuwe waarde
    if not TryVarTypeToExprType(Value, _NewType) or (_NewType = ttObject) then
      Raise EExpression.CreateFmt(sFmtCalcTypeError, [VarType(Value)]);

    // 'Waarde' aanpassen
    _Obj := TParamObject(FParameterList.Objects[i]);
    (_Obj.Value as TVariantValue).FValue := Value;
    (_Obj.Value as TVariantValue).FExprType := _NewType;
    // Bij ander type MOET de expressie opnieuw worden opgebouwd
    if (_Obj.Value as TVariantValue).FExprType <> _NewType then
      ClearCalculated;
  end;
end;

function TICalcExpression._InternalGetIdentifier(const Identifier: String;
        ParameterList: IParameterList): IValue;

  const
    sIdentName: array[boolean] of string = ('property','methode');

begin
  if ExprFunc.GetIdentifier(Identifier, ParameterList, Result) then Exit;

  if not Assigned(ParameterList) then
  begin
    // Check internal parameterlist
    Result := GetInternalParameter(Identifier);
    if Assigned(Result) then Exit;
  end;

  // Then use external call (if existing)
  If Assigned(FExternalCalc) then
    Result := FExternalCalc(Identifier, ParameterList);
  if Assigned(Result) then Exit;

  // When external call doesn't exist or doesn't return a IValue then
  // you wil get an exception (in the prExpr unit)
  Raise EExpression.CreateFmt('Missing %s %s',
    [sIdentName[Assigned(ParameterList)], Identifier]);
end;

{ TprFunctionList }
{
******************************* TprFunctionList ********************************
}
procedure TprFunctionList.AddClassName(aClass: TClass);
var
  i: Integer;
begin
  FAddingName := True;
  try
    // Index van Class in lijst opvragen
    i := IndexOf(aClass);
    if i = -1 then
      Raise EListError.CreateFmt('Class "%s" not listed', [aClass.ClassName]);

    try
      FNameList.AddObject( GetClassName(aClass), TObject(aClass) );
    except
      Delete(i);
      Raise; // Fout tonen - waarschijnlijk dubbel toegevoegd!
    end;
  finally
    FAddingName := False;
  end;
end;

constructor TprFunctionList.Create;
begin
  inherited;
  FNameList := TStringList.Create;
  TStringList(FNameList).Sorted := True; // Speed
end;

function TprFunctionList.CreateFunction(const Identifier: string; const
        ParameterList: IParameterList): IValue;
var
  i: Integer;
begin
  i := FNameList.IndexOf(Identifier);
  if i >= 0 then
    Result := TprFunctionClass(FNameList.Objects[i]).Create(ParameterList)
  else if (Assigned(ParameterList) and (ParameterList.Count > 0)) or // Has parameters?
          not CheckEnumeratedVal(System.TypeInfo(TExprType), Identifier, Result) then
    Result := nil;
end;

function TprFunctionList.CreateFunction(const Identifier: string; const
        ParameterList: IParameterList; out Value: IValue): Boolean;
var
  i: Integer;
begin
  i := FNameList.IndexOf(Identifier);
  if i >= 0 then
  begin
    Value := TprFunctionClass(FNameList.Objects[i]).Create(ParameterList);
    Result := True;
  end else
    Result :=
      ((ParameterList = nil) or (ParameterList.Count = 0)) and // No parameters?
      CheckEnumeratedVal(System.TypeInfo(TExprType), Identifier, Value);
end;

destructor TprFunctionList.Destroy;
begin
  FreeAndNil(FNameList);
  inherited;
end;

function TprFunctionList.GetClassName(aClass: TClass): string;
begin
  if not aClass.InheritsFrom(TprFunction) then
    Raise EExpression.Create('Invalid class type ');

  Result := TprFunctionClass(aClass).IdentifierName
end;

procedure TprFunctionList.Notify(Instance: Pointer; Action: TListNotification);
begin
  { -------------------------------------------------------------------------
    Bij het toevoegen van een class o.b.v. de naam, kan een fout optreden,
    wanneer de toe te voegen naam bepaald wordt.
    In dat geval, zal de reeds toegevoegde Class, verwijderd worden.
    Op dat moment is de naam van de Class dus nog niet toegevoegd en hoeft
    deze dus ook niet verwijderd te worden (controle FAddingName).
    ------------------------------------------------------------------------- }
  case Action of
    lnAdded:
      AddClassName( TClass(Instance) );
    lnDeleted:
      if not FAddingName then
        RemoveClassName( TClass(Instance) );
  end;

  inherited Notify(Instance, Action);
end;

procedure TprFunctionList.RemoveClassName(aClass: TClass);
var
  i: Integer;
begin
  if Assigned(FNameList) then
  begin
    i := FNameList.IndexOf( GetClassName(aClass) );
    if i <> -1 then
      FNameList.Delete(i);
  end;
end;

{ TNullValue }
{
********************************** TNullValue **********************************
}
function TNullValue.AsBoolean: TprBoolean;
begin
  Result := False;
end;

function TNullValue.AsFloat: TprFloat;
begin
  Result := 0;
end;

function TNullValue.AsInteger: TprInt;
begin
  Result := 0;
end;

function TNullValue.AsObject: TObject;
begin
  Result := nil;
end;

function TNullValue.AsString: TprString;
begin
  Result := '';
end;

function TNullValue.AsVariant: Variant;
begin
  Result := Unassigned;
end;

constructor TNullValue.Create(aExprType: TExprType);
begin
  FExprType := aExprType;
end;

function TNullValue.ExprType: TExprType;
begin
  Result := FExprType;
end;

function TNullValue.IsNull: Boolean;
begin
  Result := True;
end;

{ TVariantValue }
{
******************************** TVariantValue *********************************
}
function TVariantValue.AsBoolean: TprBoolean;
begin
  if IsNull then
    Result := False
  else
    Result := VarAsType(FValue, varBoolean);
end;

function TVariantValue.AsFloat: TprFloat;
begin
  if IsNull then
    Result := 0
  else
    Result := VarAsType(FValue, varDouble);
end;

function TVariantValue.AsInteger: TprInt;
begin
  if IsNull then
    Result := 0
  else
    Result := VarAsType(FValue, varInteger);
end;

function TVariantValue.AsString: TprString;
begin
  Result := VarToStr(FValue);
end;

function TVariantValue.AsVariant: Variant;
begin
  Result := FValue;
end;

constructor TVariantValue.Create(aValue: Variant);
begin
  inherited Create;
  FValue := aValue;
end;

function TVariantValue.ExprType: TExprType;
begin
  if not FExprTypeInit then
  begin
    // Probeer type te bepalen
    if not TryVarTypeToExprType(FValue, FExprType) then // Basis typen...
    begin
      if VarIsEmptyOrNull(FValue) then
        FExprType := ttObject // 'Leeg' en alleen 'TNullableObject' to raadplegen
      else
        Raise EExpression.CreateFmt('Variant type not support by %s', [Self.ClassName]);
    end;
    // Type is geinitialiseerd
    FExprTypeInit := True;
  end;
  Result := FExprType;
end;

function TVariantValue.IsNull: Boolean;
begin
  // Als het type 'Object' is, dan is de waarde leeg/null (zie ExprType)
  Result := (ExprType = ttObject);
end;

function TBaseMinMax.IgnoreCase: Boolean;
begin
  if ParameterCount > 2 then
    Result := Param[2].AsBoolean
  else
    Result := False;
end;

{ TBaseMinMax }

function TBaseMinMax.ExprType: TExprType;
begin
  Result := FCommonType;
end;

function TBaseMinMax.TestParameters: Boolean;
begin
  Result := (ParameterCount in [2, 3]);
  FCommonType := CommonType(Param[0].ExprType, Param[1].ExprType);
end;

{ TMin }

function TMin.AsBoolean: Boolean;
begin
  Result := Param[0].AsBoolean and Param[1].AsBoolean;
end;

function TMin.AsFloat: Extended;
begin
  Result := Min(Param[0].AsFloat, Param[1].AsFloat);
end;

function TMin.AsInteger: Integer;
begin
  Result := Min(Param[0].AsInteger, Param[1].AsInteger);
end;

function TMin.AsString: string;
var
  iCompResult: Integer;
begin
  if IgnoreCase then
    iCompResult := CompareText(Param[0].AsString, Param[1].AsString)
  else
    iCompResult := CompareStr(Param[0].AsString, Param[1].AsString);

  // Param[0] waarde is 'kleiner'
  if iCompResult < 0 then
    Result := Param[0].AsString
  else
    Result := Param[1].AsString;
end;

class function TMin.IdentifierName: string;
begin
  Result := 'MIN';
end;

{ TMax }

function TMax.AsBoolean: Boolean;
begin
  Result := Param[0].AsBoolean or Param[1].AsBoolean;
end;

function TMax.AsFloat: Extended;
begin
  Result := Max(Param[0].AsFloat, Param[1].AsFloat);
end;

function TMax.AsInteger: Integer;
begin
  Result := Max(Param[0].AsInteger, Param[1].AsInteger);
end;

function TMax.AsString: string;
var
  iCompResult: Integer;
begin
  if IgnoreCase then
    iCompResult := CompareText(Param[0].AsString, Param[1].AsString)
  else
    iCompResult := CompareStr(Param[0].AsString, Param[1].AsString);

  // Param[0] waarde is 'groter'
  if iCompResult > 0 then
    Result := Param[0].AsString
  else
    Result := Param[1].AsString;
end;

class function TMax.IdentifierName: string;
begin
  Result := 'MAX';
end;

{ TArray }
{
************************************ TArray ************************************
}
function TArray.AsObject: TObject;
begin
  Result := Self;
end;

function TArray.ExprType: TExprType;
begin
  Result := ttObject;
end;

class function TArray.IdentifierName: string;
begin
  Result := 'ARR';
end;

function TArray.IndexOf(Value: IValue): Integer;
var
  OK: Boolean;
begin
  for Result:=0 to ParameterCount-1 do
  begin
    OK := Param[Result].CanReadAs(Value.ExprType);
    if OK then
    Case Value.ExprType of
      ttString : OK := Param[Result].AsString  = Value.AsString;
      ttFloat  : OK := Param[Result].AsFloat   = Value.AsFloat;
      ttInteger: OK := Param[Result].AsInteger = Value.AsInteger;
      ttBoolean: OK := Param[Result].AsBoolean = Value.AsBoolean;
    end;
    if OK then Exit;
  end;
  // Niet gevonden
  Result := -1;
end;

function TArray.InList(Value: IValue): Boolean;
begin
  Result := IndexOf(Value) >= 0;
end;

function TArray.TypeInfo: PTypeInfo;
begin
  Result := Self.ClassInfo;
end;

{ TInArray }
{
*********************************** TInArray ***********************************
}
function TInArray.AsBoolean: TprBoolean;
begin
  Result := TArray(Param[1].AsObject).InList(Param[0]);
end;

function TInArray.ExprType: TExprType;
begin
  Result := ttBoolean;
end;

class function TInArray.IdentifierName: string;
begin
  Result := 'IN';
end;

function TInArray.TestParameters: Boolean;
begin
  Result :=
    (ParameterCount = 2) and
    (Param[1].TypeInfo = TArray.ClassInfo);
end;

{ TCaseOf }
{
*********************************** TCaseOf ************************************
}
function TCaseOf.AsInteger: TprInt;
begin
  Result := TArray(Param[1].AsObject).IndexOf(Param[0]);
end;

function TCaseOf.ExprType: TExprType;
begin
  Result := ttInteger; // Geeft nummer terug ( >= -1)
end;

class function TCaseOf.IdentifierName: string;
begin
  Result := 'CASEOF';
end;

function TCaseOf.TestParameters: Boolean;
begin
  // Param 0 = Zoekwaarde, Param 1 = Waarden
  Result :=
    (ParameterCount = 2) and
    (Param[1].TypeInfo = TArray.ClassInfo);
end;

{ TBaseSelect }
{
********************************* TBaseSelect **********************************
}
function TBaseSelect.AsBoolean: TprBoolean;
begin
  Result := Selected.AsBoolean;
end;

function TBaseSelect.AsFloat: TprFloat;
begin
  Result := Selected.AsFloat;
end;

function TBaseSelect.AsInteger: TprInt;
begin
  Result := Selected.AsInteger;
end;

function TBaseSelect.AsObject: TObject;
begin
  Result := Selected.AsObject;
end;

function TBaseSelect.AsString: TprString;
begin
  Result := Selected.AsString;
end;

function TBaseSelect.AsVariant: Variant;
begin
  Result := Selected.AsVariant;
end;

function TBaseSelect.IsNull: Boolean;
begin
  Result := ValueIsNull(Selected(False));
end;

function TBaseSelect.Selected(ReturnDefault: Boolean = True): IValue;
var
  n: Integer;
begin
  n := GetIndex;
  if (n < 0) or (n >= SelectArr.ParameterCount) then
  begin
    if ReturnDefault then
      Result := GetDefault
    else
      Result := nil;
  end else
    Result := SelectArr.Param[n];
end;

{ TSelect }
{
*********************************** TSelect ************************************
}
function TSelect.ExprType: TExprType;
begin
  if ParameterCount = 3 then
    Result := Param[2].ExprType // Van default waarde nemen
  else
    Result := SelectArr.Param[0].ExprType; // Anders eerste waarde in de lijst
end;

function TSelect.GetDefault: IValue;
begin
  if ParameterCount = 3 then
    Result := Param[2]
  else
    Result := TNullValue.Create(ExprType);
end;

function TSelect.GetIndex: Integer;
begin
  Result := Param[0].AsInteger;
end;

class function TSelect.IdentifierName: string;
begin
  Result := 'SELECT';
end;

function TSelect.SelectArr: TArray;
begin
  Result := TArray(Param[1].AsObject);
end;

function TSelect.TestParameters: Boolean;
begin
  // Param 0 = Index, Param 1 = Waarden, [Param 2 = Default (not in list)]
  Result :=
    (ParameterCount in [2, 3]) and
    (Param[0].CanReadAs(ttInteger)) and
    (Param[1].TypeInfo = TArray.ClassInfo);
end;

{ TCaseSelect }
{
********************************* TCaseSelect **********************************
}
function TCaseSelect.CaseArr: TArray;
begin
  Result := TArray(Param[1].AsObject);
end;

function TCaseSelect.ExprType: TExprType;
begin
  if ParameterCount = 4 then
    Result := Param[3].ExprType // Van default waarde nemen
  else
    Result := SelectArr.Param[0].ExprType; // Anders eerste waarde in de resultaatlijst
end;

function TCaseSelect.GetDefault: IValue;
begin
  if ParameterCount = 4 then
    Result := Param[3]
  else
    Result := TNullValue.Create(ExprType);
end;

function TCaseSelect.GetIndex: Integer;
begin
  Result := CaseArr.IndexOf(Param[0]);
end;

class function TCaseSelect.IdentifierName: string;
begin
  Result := 'CASESELECT';
end;

function TCaseSelect.SelectArr: TArray;
begin
  Result := TArray(Param[2].AsObject);
end;

function TCaseSelect.TestParameters: Boolean;
begin
  Result :=
    (ParameterCount in [3, 4]) and
    (Param[1].TypeInfo = TArray.ClassInfo) and
    (Param[2].TypeInfo = TArray.ClassInfo);
end;

{ TIsNull }
{
*********************************** TIsNull ************************************
}
function TIsNull.AsBoolean: Boolean;
begin
  Result := ValueIsNull(Param[0]);
end;

function TIsNull.ExprType: TExprType;
begin
  Result := ttBoolean;
end;

class function TIsNull.IdentifierName: string;
begin
  Result := 'ISNULL';
end;

function TIsNull.TestParameters: Boolean;
begin
  Result := (ParameterCount = 1);
end;

{ TIfNull }
{
*********************************** TIfNull ************************************
}
function TIfNull.AsBoolean: Boolean;
begin
  Result := Rex.AsBoolean;
end;

function TIfNull.AsFloat: Extended;
begin
  Result := Rex.AsFloat;
end;

function TIfNull.AsInteger: Integer;
begin
  Result := Rex.AsInteger;
end;

function TIfNull.AsString: string;
begin
  Result := Rex.AsString;
end;

function TIfNull.AsVariant: Variant;
begin
  Result := Rex.AsVariant;
end;

function TIfNull.ExprType: TExprType;
begin
  if Param[0].ExprType = ttObject then
    Result := Param[1].ExprType
  else
    Result := CommonType(Param[0].ExprType, Param[1].ExprType);
end;

class function TIfNull.IdentifierName: string;
begin
  Result := 'IFNULL';
end;

function TIfNull.Rex: IValue;
begin
  if ValueIsNull(Param[0]) then
    Result := Param[1]
  else
    Result := Param[0]
end;

function TIfNull.TestParameters: Boolean;
begin
  Result := (ParameterCount = 2);
end;

{ TNull }
{
********************************** TNull **********************************
}
function TNull.AsBoolean: TprBoolean;
begin
  Result := False;
end;

function TNull.AsFloat: TprFloat;
begin
  Result := 0;
end;

function TNull.AsInteger: TprInt;
begin
  Result := 0;
end;

function TNull.AsObject: TObject;
begin
  Result := nil;
end;

function TNull.AsString: TprString;
begin
  Result := '';
end;

function TNull.AsVariant: Variant;
begin
  Result := Unassigned;
end;

function TNull.ExprType: TExprType;
begin
  if ParameterCount = 1 then
    Result := ValueToExprType(Param[0])
  else
    Result := ttObject;
end;

class function TNull.IdentifierName: string;
begin
  Result := 'NULL';
end;

function TNull.IsNull: Boolean;
begin
  Result := True;
end;

function TNull.TestParameters: Boolean;
begin
  Result := (ParameterCount in [0, 1]);
end;

{ TInRange }

function TInRange.AsBoolean: Boolean;
var
  _Type: TExprType;
begin
  _Type := CommonType(Param[0].ExprType, Param[1].ExprType);
  _Type := CommonType(_Type, Param[2].ExprType);
  if _Type = ttInteger then
    Result := InRange(Param[0].AsInteger, Param[1].AsInteger, Param[2].AsInteger)
  else
    Result := InRange(Param[0].AsFloat, Param[1].AsFloat, Param[2].AsFloat)
end;

function TInRange.ExprType: TExprType;
begin
  Result := ttBoolean;
end;

class function TInRange.IdentifierName: string;
begin
  Result := 'INRANGE';
end;

function TInRange.TestParameters: Boolean;
begin
  Result := (ParameterCount = 3) and Param[0].CanReadAs(ttFloat);
end;

{ TBaseDateTime }

function TBaseDateTime.GetDateInfo: TDateInfo;
begin
  DecodeDate(Value, Result.Year, Result.Month, Result.Day);
end;

function TBaseDateTime.GetTimeInfo: TTimeInfo;
begin
  DecodeTime(Value, Result.Hour, Result.Minutes, Result.Secondes, Result.Miliseconds);
end;

function TBaseDateTime.GetValue: TDateTime;
begin
  Result := Param[0].AsInteger;
end;

class function TBaseDateTime.IdentifierName: string;
begin
  Result := Copy(Self.ClassName, 6, MaxInt); // Eerst 5 karakters overslaan 'TDate/TTime'
end;

function TBaseDateTime.TestParameters: Boolean;
begin
  Result := ParameterCount = 1;
end;

{ TDateYear }

function TDateYear.AsInteger: Integer;
begin
  Result := DateInfo.Year;
end;

function TDateYear.ExprType: TExprType;
begin
  Result := ttInteger;
end;

{ TDateMonth }

function TDateMonth.AsInteger: Integer;
begin
  Result := DateInfo.Month;
end;

function TDateMonth.ExprType: TExprType;
begin
  Result := ttInteger;
end;

{ TDateDay }

function TDateDay.AsInteger: Integer;
begin
  Result := DateInfo.Day;
end;

function TDateDay.ExprType: TExprType;
begin
  Result := ttInteger;
end;

{ TTimeHour }

function TTimeHour.AsInteger: Integer;
begin
  Result := TimeInfo.Hour;
end;

function TTimeHour.ExprType: TExprType;
begin
  Result := ttInteger;
end;

{ TTimeMinute }

function TTimeMinute.AsInteger: Integer;
begin
  Result := TimeInfo.Minutes;
end;

function TTimeMinute.ExprType: TExprType;
begin
  Result := ttInteger;
end;

{ TTimeSecond }

function TTimeSecond.AsInteger: Integer;
begin
  Result := TimeInfo.Secondes;
end;

function TTimeSecond.ExprType: TExprType;
begin
  Result := ttInteger;
end;

{ TTimeMilisecond }

function TTimeMilisecond.AsInteger: Integer;
begin
  Result := TimeInfo.Miliseconds;
end;

function TTimeMilisecond.ExprType: TExprType;
begin
  Result := ttInteger;
end;

{ TEncodeDate }

function TEncodeDate.AsFloat: TprFloat;
begin
  // Year, month, day
  Result := EncodeDate(Param[0].AsInteger, Param[1].AsInteger, Param[2].AsInteger)
end;

function TEncodeDate.AsString: TprString;
begin
  Result := DateToStr(AsFloat);
end;

function TEncodeDate.ExprType: TExprType;
begin
  Result := ttFloat;
end;

class function TEncodeDate.IdentifierName: string;
begin
  Result := 'DATE';
end;

function TEncodeDate.TestParameters: Boolean;
begin
  Result := ParameterCount = 3;
end;

{ TEncodeTime }

function TEncodeTime.AsFloat: TprFloat;
begin
  // Hout, minute, second, milisecond
  Result := EncodeTime(Param[0].AsInteger, Param[1].AsInteger, Param[2].AsInteger, Param[3].AsInteger);
end;

function TEncodeTime.AsString: TprString;
begin
  Result := TimeToStr(AsFloat);
end;

function TEncodeTime.ExprType: TExprType;
begin
  Result := ttFloat;
end;

class function TEncodeTime.IdentifierName: string;
begin
  Result := 'TIME';
end;

function TEncodeTime.TestParameters: Boolean;
begin
  Result := ParameterCount = 4;
end;

initialization

  _ExprTypeCastProcs := nil;
  _FunctionList := TprFunctionList.Create;

  // Registreer routines
  TMin.Register;
  TMax.Register;

  TArray.Register;
  TCaseOf.Register;
  TSelect.Register;
  TCaseSelect.Register;
  TInArray.Register;
  TIsNull.Register;
  TIfNull.Register;
  TNull.Register;

  TInRange.Register;

  TDateYear.Register;
  TDateMonth.Register;
  TDateDay.Register;
  TTimeHour.Register;
  TTimeMinute.Register;
  TTimeSecond.Register;
  TTimeMilisecond.Register;
  TEncodeDate.Register;
  TEncodeTime.Register;

finalization

  FreeAndNil(_FunctionList);

end.
