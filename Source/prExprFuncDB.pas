unit prExprFuncDB;

interface

uses
  DB, TypInfo, prExpr, prExprFunc;

type
  TFieldValue = class(TNullableExpression)
  private
    FField: TField;
  public
    constructor Create(aField: TField);
    function AsBoolean: Boolean; override;
    function AsFloat: Extended; override;
    function AsInteger: Integer; override;
    function AsObject: TObject; override;
    function AsString: string; override;
    function AsVariant: Variant; override;
    function ExprType: TExprType; override;
    function IsNull: Boolean; override;
  end;

  TFieldTypeLiteral = class(TExpression)
  private
    FFieldType: TFieldType;
  public
    constructor Create(const aFieldType: string); overload;
    constructor Create(aFieldType: TFieldType); overload;
    function AsInteger: Integer; override;
    function ExprType: TExprType; override;
    function TypeInfo: PTypeInfo; override;
  end;

  function FieldTypeToExprType(aFieldType: TFieldType): TExprType;

implementation

(*
  Private routine
*)

// TFieldType to TExprType cast procedure
function CastFieldTypeToExprType(aValue: IValue): TExprType;
begin
  Result := FieldTypeToExprType(TFieldType(aValue.AsInteger));
end;

(*
  Publieke hulproutines
*)

function FieldTypeToExprType(aFieldType: TFieldType): TExprType;
begin
  case aFieldType of
    ftString, ftMemo, ftFmtMemo, ftFixedChar,
    ftWideString, ftFixedWideChar, ftWideMemo,
    ftGuid:
      Result := ttString;

    ftFloat, ftCurrency, ftBCD, ftTime, ftDateTime, ftTimeStamp:
      Result := ttFloat;

    ftSmallint, ftInteger, ftWord, ftDate, ftAutoInc, ftLargeint:
      Result := ttInteger;

    ftBoolean:
      Result := ttBoolean;
  else
    Result := ttObject;
  end;
end;

{ TFieldValue }
{
******************************** TFieldValue *********************************
}
function TFieldValue.AsBoolean: Boolean;
begin
  Result := FField.AsBoolean;
end;

function TFieldValue.AsFloat: Extended;
begin
  Result := FField.AsFloat;
end;

function TFieldValue.AsInteger: Integer;
begin
  // Datum veld kan niet o.b.v. AsInteger worden uitgelezen
  // - dus uitlezen als 'AsFloat en waarde truncen!
  if FField.DataType = ftDate then
    Result := Trunc(FField.AsFloat)
  else
    Result := FField.AsInteger;
end;

function TFieldValue.AsObject: TObject;
begin
  Result := FField;
end;

function TFieldValue.AsString: string;
begin
  Result := FField.AsString;
end;

function TFieldValue.AsVariant: Variant;
begin
  Result := FField.AsVariant;
end;

constructor TFieldValue.Create(aField: TField);
begin
  inherited Create;
  FField := aField;
end;

function TFieldValue.ExprType: TExprType;
begin
  Result := FieldTypeToExprType(FField.DataType);
end;

function TFieldValue.IsNull: Boolean;
begin
  Result := FField.IsNull;
end;

{ TFieldTypeLiteral }

function TFieldTypeLiteral.AsInteger: Integer;
begin
  Result := Integer(FFieldType);
end;

constructor TFieldTypeLiteral.Create(const aFieldType: string);
begin
  inherited Create;
  FFieldType := TFieldType(GetEnumValue(TypeInfo, aFieldType));
end;

constructor TFieldTypeLiteral.Create(aFieldType: TFieldType);
begin
  inherited Create;
  FFieldType := aFieldType;
end;

function TFieldTypeLiteral.ExprType: TExprType;
begin
  Result := ttEnumerated;
end;

function TFieldTypeLiteral.TypeInfo: PTypeInfo;
begin
  Result := System.TypeInfo(TFieldType);
end;

initialization

  RegisterExprTypeCastProc(System.TypeInfo(TFieldType), CastFieldTypeToExprType);

finalization

  UnregisterExprTypeCastProc(System.TypeInfo(TFieldType), CastFieldTypeToExprType);

end.
