unit prExpr;

interface
{main documentation block just before implementation
Copyright:   1997-1999 Production Robots Engineering Ltd, all rights reserved.
Version:     1.04 3/7/99
Status:      Free for private or commercial use subject to following restrictions:
             * Use entirely at your own risk
             * Do not resdistribute without this note
             * Any redistribution to be free of charges
any questions to Martin Lafferty  martinl@prel.co.uk (Old: robots@enterprise.net)

Homepage: http://www.pre.demon.co.uk/delphi.htm }

// Use routines from Math unit for comparing float values
{$DEFINE USEMATHFLOATCOMPARE}

uses
{$IFDEF MSWINDOWS}
  Windows, // For MessageBox function
{$ENDIF}
{$IFDEF LINUX}
  Types,
  Libc,
{$ENDIF}
  TypInfo, Classes, SysUtils;

type
  TExprType = (ttObject, ttString, ttFloat, ttInteger, ttEnumerated, ttBoolean);

  TprBoolean = boolean;
  TprInt = Integer;
  TprFloat = Extended;
  TprString = string;
  TprVariant = Variant;

  IValue = interface(IInterface)
    ['{286DCC4A-83AA-41E9-9D5B-B122ED6F5A55}']
    function AsBoolean: TprBoolean;
    function AsFloat: TprFloat;
    function AsInteger: TprInt;
    function AsObject: TObject;
    function AsString: TprString;
    function AsVariant: TprVariant;
    function CanReadAs(aType: TExprType): Boolean;
    function ExprType: TExprType;
    function TestParameters: Boolean;
    function TypeInfo: PTypeInfo;
    function TypeName: string;
  end;

  TExpression = class(TInterfacedObject, IValue)
  protected
    function TestParameters: Boolean; virtual;
  public
    constructor Create;
    destructor Destroy; override;
    function AsBoolean: TprBoolean; virtual;
    function AsFloat: TprFloat; virtual;
    function AsInteger: TprInt; virtual;
    function AsObject: TObject; virtual;
    function AsString: TprString; virtual;
    function AsVariant: TprVariant; virtual;
    function CanReadAs(aType: TExprType): Boolean;
    function ExprType: TExprType; virtual; abstract;
    function TypeInfo: PTypeInfo; virtual;
    function TypeName: string; virtual;
  end;

  TStringLiteral = class(TExpression)
  private
    FAsString: TprString;
  public
    constructor Create(aAsString: TprString);
    function AsString: TprString; override;
    function ExprType: TExprType; override;
  end;

  TFloatLiteral = class(TExpression)
  private
    FAsFloat: TprFloat;
  public
    constructor Create(aAsFloat: TprFloat);
    function AsFloat: TprFloat; override;
    function ExprType: TExprType; override;
  end;

  TIntegerLiteral = class(TExpression)
  private
    FAsInteger: TprInt;
  public
    constructor Create(aAsInteger: TprInt);
    function AsInteger: TprInt; override;
    function ExprType: TExprType; override;
  end;

  TEnumeratedLiteral = class(TIntegerLiteral)
  private
    Rtti: Pointer;
  public
    constructor Create(aRtti: Pointer; aAsInteger: Integer);
    constructor StrCreate(aRtti: Pointer; const aVal: String);
    function ExprType: TExprType; override;
    function TypeInfo: PTypeInfo; override;
  end;

  TBooleanLiteral = class(TExpression)
  private
    FAsBoolean: TprBoolean;
  public
    constructor Create(aAsBoolean: TprBoolean);
    function AsBoolean: Boolean; override;
    function ExprType: TExprType; override;
  end;

  TVariantLiteral = class(TExpression)
  private
    FVariant: TprVariant;
  public
    constructor Create(const aAsVariant: TprVariant);
    function AsBoolean: TprBoolean; override;
    function AsFloat: TprFloat; override;
    function AsInteger: TprInt; override;
    function AsString: TprString; override;
    function AsVariant: TprVariant; override;
    function ExprType: TExprType; override;
    function TypeInfo: PTypeInfo; override;
  end;

  TObjectRef = class(TExpression)
  private
    FObject: TObject;
  public
    constructor Create(aObject: TObject);
    function AsObject: TObject; override;
    function ExprType: TExprType; override;
    function TypeInfo: PTypeInfo; override;
  end;

  IParameterList = interface(IInterface)
    function AddExpression(e: IValue): Integer;
    procedure CheckCount(const Count: Byte); overload;
    procedure CheckCount(const MinCount, MaxCount: Byte); overload;
    procedure CheckCount(const Counts: array of Byte); overload;
    function GetAsBoolean(i: Integer): TprBoolean;
    function GetAsFloat(i: Integer): TprFloat;
    function GetAsInteger(i: Integer): TprInt;
    function GetAsObject(i: Integer): TObject;
    function GetAsString(i: Integer): TprString;
    function GetCount: Integer;
    function GetExprType(i: Integer): TExprType;
    function GetParam(i: Integer): IValue;
    property AsBoolean[i: Integer]: TprBoolean read GetAsBoolean;
    property AsFloat[i: Integer]: TprFloat read GetAsFloat;
    property AsInteger[i: Integer]: TprInt read GetAsInteger;
    property AsObject[i: Integer]: TObject read GetAsObject;
    property AsString[i: Integer]: TprString read GetAsString;
    property Count: Integer read GetCount;
    property ExprType[i: Integer]: TExprType read GetExprType;
    property Param[i: Integer]: IValue read GetParam; default;
  end;

  TFunction = class(TExpression)
  private
    FParameterList: IParameterList;
    function GetParam(n: Integer): IValue;
    function GetParameterCount: Integer;
  public
    constructor Create(aParameterList: IParameterList);
    property Param[n: Integer]: IValue read GetParam;
    property ParameterCount: Integer read GetParameterCount;
  end;

  EExpression = class(Exception)
  end;

  TIdentifierFunction = function (const Identifier: String; ParameterList:
          IParameterList): IValue of object;
  function CheckEnumeratedVal(Rtti: Pointer; const aVal: String): IValue; overload;
  function CheckEnumeratedVal(Rtti: Pointer; const aVal: String; out oValue:
      IValue): Boolean; overload;
  function CreateExpression(const S: String; IdentifierFunction:
      TIdentifierFunction; const DisableStandardFunc: boolean = False): IValue;

  function FoldConstant( Value: IValue): IValue;
  {replace complex constant expression with literal, hence reducing
   evaluation time. This function does not release Value - caller
   should do that if it is appropriate. This is usually, but not
   necessarily always, the case.}

  ///  <summary>
  ///  Get commonly readable datatype of two types
  ///  </summary>
  function CommonType( Op1Type, Op2Type: TExprType): TExprType;

type
  TICType = (icExpression, icParameterList);
const
  prExprMajorVersion = 1;
  prExprMinorVersion = 5;
var
  WarnForInstanceCountGreaterZeroOnUninitialize: boolean = False;

  function InstanceCount(const aType: TICType): integer;

implementation

uses
{$IFDEF USEMATHFLOATCOMPARE}
  Math,
{$ENDIF}
  Variants;
var
  _ExprInstanceCount: Integer = 0;
  _ParamListInstanceCount: Integer = 0;
ThreadVar
  DisableStandardFunctionProcessing: boolean;

function InstanceCount(const aType: TICType): integer;
begin
  Case aType of
    icExpression   : Result := _ExprInstanceCount;
    icParameterList: Result := _ParamListInstanceCount;
  else
    Result := -1;
  end;
end;

{This unit comprises a mixed type expression evaluator which follows pascal
syntax (reasonably accurately) and approximates standard pascal types.

Feedback
--------
I shall be pleased to hear from anyone
- with any questions or comments
- who finds or suspects any bugs
- who wants me to quote for implementing extensions or applications

in any event, my address is:   robots@enterprise.net
I am sometimes very busy and cannot always enter into protracted or complex correspondence
but I do really like to hear from users of this code

I have found this code very useful and surprisingly robust. I sincerely hope you do too.

For detailed explanation as to how to effectively use this unit please refer to prExpr.txt.

Compatibility
-------------
Version 1.04 of this unit is not compatible with previous versions and will break
existing applications.

This code developed with Delphi 4 and 3. It won't work with Delphi 2 because it uses
interfaces. If you want a Delphi 2 version then try v1.03 which also supports Delphi 1
(16 bit)

Additional Resources
--------------------
This archive includes a help file prExpr.hlp, which you can incorporate into your
help system to provide your users with a definition of expression syntax. If you want
the rtf file from which this is compiled then download

http://homepages.enterprise.net/robots/downloads/exprhelp.zip

That package includes
  prExpr.rtf - The source file (inline graphics)
  prExpr.hpj - The help project file.

Next Steps (ideas not implemented)
----------
TypeRegistry for enumerated types could
(a) allow enumerates to be efficiently and automatically parsed
    (Build a binary list of all the names)
(b) allow typecasts to enumerated types


Version History
---------------
Homepage:
  http://www.pre.demon.co.uk/delphi.htm

latest version should be available from:
  http://www.pre.demon.co.uk/downloads/expreval.zip

version 1.04 is not backwardly compatible with previous versions so
for historical reasons version 1.03 should be available from:
  http://www.pre.demon.co.uk/downloads/expr103.zip

prExpr version 1.05 - 2004/01/02

From this verion on, the ParameterList is based on an InterfacedObject,
thus not needing to free it manually.
As soon as there are no more references to the List, it is free.
(This will make the notification on "16/9/97 - Compromise solution:"
 not valid anymore)

Added then "prExprMajorVersion" and "prExprMinorVersion" constant.

Code modifications:
- new ParamterList by: R. Hoek - ComponentAgro B.V. - ronald.hoek@componentagro.nl
- Float comparison using routine 'CompareValue'/'SameValue' from Delphi 'Math' unit
  (To use old binary compare disable the define USEMATHFLOATCOMPARE)

prExpr version 1.04
  For Version 1.04 the help file is out of date: it is not actually wrong - it just misses out
  a load of stuff, like class references and enumerated types.

  v104 is not backwardly compatible with v1.03 but it might be worth converting your applications
  because there is some quite neat stuff in here. You really need to understand interfaces though.

  1. No longer supports 16 bit
  2. Value is now an interface. This was an idea I nicked from Clayton Collie.
     This makes handling object disposals really easy when there are lots of random references
     to a bunch of expressions stored in a haphazard way. The compiler handles it for you.
  3. New rule: it must be possible to determine the type of an expression at parse-time. Therefore the
     two parameters to IF must be the same type. This prevents lots of awkward situations.
  4. Typecasts should now work both to a more general type (e.g Integer to string) and to a more
     specific type (String to integer). Implausible casts e.g Integer('four') will raise an exception
     when evaluated, but still return a valid type at parse time (see rule 3 above).
  5. Objects now supported. If an ID Function returns an instance of an object, then all its
     published properties are available in the expression without further ado. Don't forget a class
     has to be compiled under $M+ to have rtti.
  6. Enumerated types now supported, using rtti.
  7. Rtti means 'Run-Time-Type-Information'. I thought you would know that.
  8. While I am about breaking backward compatibility, I have decided to get rid
     of most read-only properties. Surely it is silly to use a complicated syntax
     when there is an equivalent, much simpler option
     (Maybe I thought write specifiers would have some meaning in some other
     context of IValue)

31/1/98 v1.03
(a) Unit Name changed from 'Expressions' to 'prExpr'.
    reasons:
     1. Merging 16/32 versions into one unit means name must be
        8.3 compliant
     2. Name should be 8.3 compliant anyway. Long filenames are
        still a pain in the neck.
     3. 'Expressions' is a term with too many meanings. Better to
        use an arbitary, mostly meaningless name.

(b) Incorporation of 16 & 32 bit versions in one unit.

(c) Modification by Markus Stephany
    (http://home.t-online.de/home/mirbir.st)
    Support for Hex literals added. Marked (mst) in source.

(d) Reverse 'Decimal Separator' mod made in v1.02.

(e) Significant structural changes to rationalise by
    eliminating repeated code. Introduced concept of 'Expression
    chain' which means that functions Factor, Term, Simple,
    and Expression now have a common implementation (Chain)
    The source is now a lot shorter, and, I hope, clearer. These
    changes should have eliminated Ken Friesen's bug in a more
    structured way.

(f) I have added another 'tier' to the syntax hierachy. The basic
    syntax element is now the 'SimpleFactor' - (was Factor). A
    factor now consists of a string of SimpleFactors linked by
    ^ the exponention operator. This change allows the ^ operator
    to be supported. prExpr.hlp updated.

(g) Archive structure changes:
      1. Expr.hlp has been renamed prExpr.hlp and is included with the
         issue archive.

      2. Tutorial documentation removed from this file to a separate file
         prExpr.txt. (hint - right click on filename then choose 'Open file
         at cursor')

      3. Form unit name changed to Main.pas

      4. Now includes 16 bit example files (tester16.dpr, Main16.pas, Main16.dfm)


20/1/98 v1.02
companion help material issued.

Structure of comment blocks rationalised. Or derationalised,
depending on your point of view.

9/1/98 v1.01
Bugs reported by Ken Friesen

1) (1+2))-1=3
(this is a bug, but known. See comment right at end
of source. Function EoE (EndOfExpression) returns true for all of
 ')' ',' or #0. This is necessary for handing functions and
 parameters but irritating if your expression has an extra ).
 Fixed.

2) 1+( )= Access Violation
Oversight. Parser did not check for null subexpresssion (fixed)

3) 0-2+2=-4
4) 1-0+1=0
5) -2-2+2 = -6
I cannot believe that this has not been picked up before now!
There is an awful lot of recursion about and this was caused
by the fact that the function SIMPLE called itself to in order
to generate a string of TERMS. The result of this is that any
simple expressions containing more than two terms were constructed
as if they were bracketed from the end of the expression.

i.e a+b+c+d was evaluated as a + (b + (c + d))

This was an elegant construct and I fell for it regardless of the fact
that it was completely wrong. This problem also affected the function
Term, but because a*b*c*d = a*(b*(c*d)) I got away with it.

I have made a (rough) fix. Which works but may have introduced other
problems. The structure of simple is now (approximately)

function Simple: TExpression;
Result:= Term
while NextOperator in [+, -, or, xor] do
  Result:= Binary(NextOperator, Result, Term)

As opposed to the previous (incorrect) way of doing things which was:

function Simple: TExpression;
Result:= Term
if NextOperator in [+, -, or, xor] then
  Result:= Binary(NextOperator, Result, Simple)


I have also made this modification to  TERM, in order to be consistent,
and because it was a  fluke that it worked before.

The unit now passes Ken's tests. I cannot be sure I have not introduced
other problems. I should devise a proper test routine, when I have some
time.

30/12/97 v1.00
Released to http://homepages.enterprise.net/robots/downloads/expreval.zip
Some slight restructing. Added more comprehensive documentation. Removed
a few calls to StrPas which are redundant under D2/D3

11/11/97
Bug caused mishandling of function lists. Fixed.

5/11/97
Slight modifications for first issue of Troxler.exe

16/9/97
Realised that it should be possible to pass the parameter stack
to the identifier function. The only problem with this approach is
how to handle disposal of the stack.

We could require that the identifier function disposes of the stack...
I don't really like this (I can't think why at the moment). Another
approach would be to define a 'placeholder' expression which does nothing
but hold the parameter list and the <clients> expression.

Compromise solution:
  The parser constructs an instance of TParameter list and passes it to
  the 'user' via a call to IdentifierFunction. There are four possible
  mechanisms for disposal of the parameter list.
     a) If the Identifier function returns NIL the parser disposes
        of the parameter list then raises 'Unknown identifier'.
     b) If the Identifier function raises an exception then the parser
        catches this exception (in a 'finally' clause) and disposes
        of the parameter list.
     c) If the Identifier function returns an expression then it must
        dispose of the parameter list if it does not wish to keep it.
     d) If the Identifier function returns an expression which is
        derived from TFunction, then it may pass the parameter list to
        its result. The result frees the parameter list when it is freed.
        (i.e. ParameterList passed to TFunction.Create is freed by
        TFunction.Destroy)

Simple rule - if IdentFunction returns Non-nil then parameters are
responsiblity of the object returned. Otherwise caller will handle. OK?

7/9/97
function handling completely changed.

added support for Integers including support for the following operators
  bitwise not
  bitwise and
  bitwise or
  bitwise xor
  shl
  shr
  div

now support std functions:

arithmetic...
  TRUNC, ROUND, ABS, ARCTAN, COS, EXP, FRAC, INT,
     LN, PI, SIN, SQR, SQRT, POWER

string...
  UPPER, LOWER, COPY, POS, LENGTH

Fixed a couple of minor bugs. Forgotten what they are.


18/6/97
Written for Mark Page's troxler thing - as part of the report definition language,
but might be needed for Robot application framework. Not tested much.
Loosely based on syntax diagrams in BP7 Language Guide pages 66 to 79.
This is where the nomenclature Term, Factor, SimpleExpression, Expression is
derived.
}

type
  TParameterList = class(TInterfacedObject, IParameterList)
  private
    FList: TList;
    function AddExpression(e: IValue): Integer;
    function GetAsBoolean(i: Integer): TprBoolean;
    function GetAsFloat(i: Integer): TprFloat;
    function GetAsInteger(i: Integer): TprInt;
    function GetAsObject(i: Integer): TObject;
    function GetAsString(i: Integer): TprString;
    function GetCount: Integer;
    function GetExprType(i: Integer): TExprType;
    function GetParam(i: Integer): IValue;
  protected
    procedure CheckCount(const Count: Byte); overload;
    procedure CheckCount(const MinCount, MaxCount: Byte); overload;
    procedure CheckCount(const Counts: array of Byte); overload;
  public
    constructor Create;
    destructor Destroy; override;
  end;

  TOperator = ( opNot,
                opExp,
                opMult, opDivide, opDiv, opMod, opAnd, opShl, opShr,
                opPlus, opMinus, opOr, opXor,
                opEq, opNEq, opLT, opGT, opLTE, opGTE);

  TOperators = set of TOperator;

  TUnaryOp = class(TExpression)
  private
    FOperand: IValue;
    FOperandType: TExprType;
    FOperator: TOperator;
  public
    constructor Create(aOperator: TOperator; aOperand: IValue);
    function AsBoolean: TprBoolean; override;
    function AsFloat: TprFloat; override;
    function AsInteger: TprInt; override;
    function ExprType: TExprType; override;
  end;

  TBinaryOp = class(TExpression)
  private
    FOperand1: IValue;
    FOperand2: IValue;
    FOperandType: TExprType;
    FOperator: TOperator;
  public
    constructor Create(aOperator: TOperator; aOperand1, aOperand2: IValue);
    function AsBoolean: TprBoolean; override;
    function AsFloat: TprFloat; override;
    function AsInteger: TprInt; override;
    function AsString: TprString; override;
    function ExprType: TExprType; override;
  end;

  TRelationalOp = class(TExpression)
  private
    FOperand1: IValue;
    FOperand2: IValue;
    FOperandType: TExprType;
    FOperator: TOperator;
  public
    constructor Create(aOperator: TOperator; aOperand1, aOperand2: IValue);
    function AsBoolean: TprBoolean; override;
    function AsFloat: TprFloat; override;
    function AsInteger: TprInt; override;
    function AsString: TprString; override;
    function ExprType: TExprType; override;
  end;

  TObjectProperty = class(TExpression)
  private
    FObj: TObject;
    FPropInfo: PPropInfo;
    FPropType: TExprType;
  public
    constructor Create(aObj: IValue; const PropName: String);
    function AsBoolean: TprBoolean; override;
    function AsFloat: TprFloat; override;
    function AsInteger: TprInt; override;
    function AsObject: TObject; override;
    function AsString: TprString; override;
    function ExprType: TExprType; override;
    function TypeInfo: PTypeInfo; override;
  end;


const
  MaxStringLength = 255; {why?}
  Digits = ['0'..'9'];
  PrimaryIdentChars = ['a'..'z', 'A'..'Z', '_'];
  IdentChars = PrimaryIdentChars + Digits;


  NExprType: array[TExprType] of String =
      ('Object', 'String', 'Float', 'Integer', 'Enumerated', 'Boolean');

  NOperator: array[TOperator] of String =
              ( 'opNot',
                'opExp',
                'opMult', 'opDivide', 'opDiv', 'opMod', 'opAnd', 'opShl', 'opShr',
                'opPlus', 'opMinus', 'opOr', 'opXor',
                'opEq', 'opNEq', 'opLT', 'opGT', 'opLTE', 'opGTE');

  UnaryOperators = [opNot];
  ExpOperator = [opExp];
  MultiplyingOperators = [opMult, opDivide, opDiv, opMod, opAnd, opShl, opShr];
  AddingOperators = [opPlus, opMinus, opOr, opXor];
  RelationalOperators = [opEq, opNEq, opLT, opGT, opLTE, opGTE];

  NBoolean: array[Boolean] of String = ('FALSE', 'TRUE');

function ResultType( Operator: TOperator; OperandType: TExprType): TExprType;

  procedure NotAppropriate;
  begin
    Result:= ttString;
    raise EExpression.CreateFmt( 'Operator %s incompatible with %s',
                                 [NOperator[Operator], NExprType[OperandType]])
  end;

begin
  case OperandType of
    ttString:
    case Operator of
      opPlus: Result:= ttString;
      opEq..opGTE: Result:= ttBoolean;
    else
      NotAppropriate;
    end;
    ttFloat:
    case Operator of
      opExp, opMult, opDivide, opPlus, opMinus: Result:= ttFloat;
      opEq..opGTE: Result:= ttBoolean;
    else
      NotAppropriate;
    end;
    ttInteger:
    case Operator of
      opNot, opMult, opDiv, opMod, opAnd, opShl, opShr, opPlus, opMinus,
      opOr, opXor: Result:= ttInteger;
      opExp, opDivide: Result:= ttFloat;
      opEq..opGTE: Result:= ttBoolean;
    else
      NotAppropriate;
    end;
    ttBoolean:
    case Operator of
      opNot, opAnd, opOr, opXor, opEq, opNEq: Result:= ttBoolean;
    else
      NotAppropriate;
    end;
    ttObject:
    case Operator of
      opEq, opNEq: Result:= ttBoolean;
    else
      NotAppropriate;
    end;
  else
    NotAppropriate
  end
end;

function IncompatibleTypes(T1, T2: TExprType): TExprType;
{result is not defined... Do this to avoid warning}
begin
  raise EExpression.CreateFmt('Type %s is incompatible with type %s',
                                  [NExprType[T1], NExprType[T2]])
end;

function CommonType( Op1Type, Op2Type: TExprType): TExprType;
begin
  if (Op1Type = ttObject) or (Op2Type = ttObject) then
  begin
    if Op1Type <> Op2Type then
      Result:= IncompatibleTypes(Op1Type, Op2Type)
    else
      Result:= ttObject
  end else
  begin
    if Op1Type < Op2Type then
      Result:= Op1Type else
      Result:= Op2Type
  end;
end;

procedure Internal( Code: Integer);
begin
  raise EExpression.CreateFmt('Internal parser error. Code %d', [Code])
end;

{
********************************* TExpression **********************************
}
function TExpression.AsBoolean: TprBoolean;
begin
  raise EExpression.CreateFmt('Cannot read %s as boolean',
                               [NExprType[ExprType]])
end;

function TExpression.AsFloat: TprFloat;
begin
  case ExprType of
    ttInteger, ttEnumerated, ttBoolean: Result:= AsInteger;
  else
    raise EExpression.CreateFmt('Cannot read %s as Float',
                                   [NExprType[ExprType]]);
  end
end;

function TExpression.AsInteger: TprInt;
begin
  case ExprType of
    ttBoolean: Result:= Integer(AsBoolean);
  else
    raise EExpression.CreateFmt('Cannot read %s as Integer',
                               [NExprType[ExprType]]);
  end;
end;

function TExpression.AsObject: TObject;
begin
  raise EExpression.CreateFmt('Cannot read %s as object',
                               [NExprType[ExprType]])
end;

function TExpression.AsString: TprString;

  {too scary to deal with Enumerated types here?}

begin
  case ExprType of
    ttObject    : Result:= AsObject.ClassName;
    ttFloat     : Result:= FloatToStr(AsFloat);
    ttInteger   : Result:= IntToStr(AsInteger);
    ttEnumerated: Result:= GetEnumName(TypeInfo, AsInteger);
    ttBoolean   : Result:= NBoolean[AsBoolean];
  else
    EExpression.CreateFmt('Cannot read %s as String',
                              [NExprType[ExprType]]);
  end
end;

function TExpression.AsVariant: TprVariant;
begin
  case ExprType of
  //    ttObject    : Result := AsObject.ClassName;
    ttString    : Result := AsString;
    ttFloat     : Result := AsFloat;
    ttInteger,
    ttEnumerated: Result:= AsInteger;
    ttBoolean   : Result:= AsBoolean;
  else
    EExpression.CreateFmt('Cannot read %s as Variant',
                              [NExprType[ExprType]]);
  end
end;

function TExpression.CanReadAs(aType: TExprType): Boolean;
var
  et: TExprType;
begin
  et:= ExprType;
  if (et = ttObject) or
     (aType = ttObject) then
    Result:= aType = et
  else
    Result:= aType <= et
end;

constructor TExpression.Create;
begin
  inherited;
  Inc(_ExprInstanceCount);
end;

destructor TExpression.Destroy;
begin
  Dec(_ExprInstanceCount);
  inherited;
end;

function TExpression.TestParameters: Boolean;
begin
  Result:= true
end;

function TExpression.TypeInfo: PTypeInfo;
begin
  case ExprType of
    ttString: Result:= System.TypeInfo(String);
    ttFloat: Result:= System.TypeInfo(Double);
    ttInteger: Result:= System.TypeInfo(Integer);
    ttBoolean: Result:= System.TypeInfo(Boolean);
  else
    raise EExpression.CreateFmt('Cannot provide TypeInfo for %s', [ClassName])
  end
end;

function TExpression.TypeName: string;
begin
  Result:= TypeInfo^.Name;
end;

{
******************************** TStringLiteral ********************************
}
function TStringLiteral.AsString: TprString;
begin
  Result:= FAsString
end;

constructor TStringLiteral.Create(aAsString: TprString);
begin
  inherited Create;
  FAsString:= aAsString
end;

function TStringLiteral.ExprType: TExprType;
begin
  Result:= ttString
end;

{
******************************** TFloatLiteral *********************************
}
function TFloatLiteral.AsFloat: TprFloat;
begin
  Result:= FAsFloat
end;

constructor TFloatLiteral.Create(aAsFloat: TprFloat);
begin
  inherited Create;
  FAsFloat:= aAsFloat
end;

function TFloatLiteral.ExprType: TExprType;
begin
  Result:= ttFloat
end;

{
******************************* TIntegerLiteral ********************************
}
function TIntegerLiteral.AsInteger: TprInt;
begin
  Result:= FAsInteger
end;

constructor TIntegerLiteral.Create(aAsInteger: TprInt);
begin
  inherited Create;
  FAsInteger:= aAsInteger
end;

function TIntegerLiteral.ExprType: TExprType;
begin
  Result:= ttInteger
end;

{
******************************* TBooleanLiteral ********************************
}
function TBooleanLiteral.AsBoolean: Boolean;
begin
  Result:= FAsBoolean
end;

constructor TBooleanLiteral.Create(aAsBoolean: TprBoolean);
begin
  inherited Create;
  FAsBoolean:= aAsBoolean
end;

function TBooleanLiteral.ExprType: TExprType;
begin
  Result:= ttBoolean
end;

{
****************************** TEnumeratedLiteral ******************************
}
constructor TEnumeratedLiteral.Create(aRtti: Pointer; aAsInteger: Integer);
begin
  inherited Create(aAsInteger);
  Rtti:= aRtti
end;

function TEnumeratedLiteral.ExprType: TExprType;
begin
  Result := ttEnumerated;
end;

constructor TEnumeratedLiteral.StrCreate(aRtti: Pointer; const aVal: String);
var
  i: Integer;
begin
  i:= GetEnumValue(PTypeInfo(aRtti), aVal);
  if i = -1 then
    raise EExpression.CreateFmt('%s is not a valid value for %s',
                [aVal, PTypeInfo(aRtti)^.Name]);
  Create(aRtti, i)
end;

function TEnumeratedLiteral.TypeInfo: PTypeInfo;
begin
  Result:= Rtti
end;

function CheckEnumeratedVal(Rtti: Pointer; const aVal: String): IValue;
begin
  try
    Result:= TEnumeratedLiteral.StrCreate(Rtti, aVal)
  except
    on EExpression do
      Result:= nil
  end
end;

function CheckEnumeratedVal(Rtti: Pointer; const aVal: String; out oValue:
    IValue): Boolean;
var
  i: Integer;
begin
  i:= GetEnumValue(PTypeInfo(Rtti), aVal);
  if i = -1 then
    Result := False
  else begin
    oValue := TEnumeratedLiteral.Create(Rtti, i);
    Result := True;
  end;
end;

{
******************************* TVariantLiteral ********************************
}
function TVariantLiteral.AsBoolean: TprBoolean;
begin
  Result := VarAsType(FVariant, varBoolean);
end;

function TVariantLiteral.AsFloat: TprFloat;
begin
  Result := VarAsType(FVariant, varDouble)
end;

function TVariantLiteral.AsInteger: TprInt;
begin
  Result := VarAsType(FVariant, varInteger)
end;

function TVariantLiteral.AsString: TprString;
begin
  Result := VarToStr(FVariant);
end;

function TVariantLiteral.AsVariant: TprVariant;
begin
  Result := FVariant;
end;

constructor TVariantLiteral.Create(const aAsVariant: TprVariant);
begin
  inherited Create;
  FVariant := aAsVariant;
end;

function TVariantLiteral.ExprType: TExprType;
begin
  if VarIsType(FVariant, varBoolean) then
    Result := ttBoolean
  else if VarIsOrdinal(FVariant) then
    Result := ttInteger
  else if VarIsFloat(FVariant) then
    Result := ttFloat
  else
    Result := ttString;
end;

function TVariantLiteral.TypeInfo: PTypeInfo;
begin
  Result := System.TypeInfo(Variant);
end;

{
********************************** TObjectRef **********************************
}
function TObjectRef.AsObject: TObject;
begin
  Result:= FObject
end;

constructor TObjectRef.Create(aObject: TObject);
begin
  inherited Create;
  FObject:= aObject
end;

function TObjectRef.ExprType: TExprType;
begin
  Result:= ttObject
end;

function TObjectRef.TypeInfo: PTypeInfo;
begin
  if Assigned(FObject) then
    Result:= FObject.ClassInfo
  else
    Result:= TObject.ClassInfo
end;

{
*********************************** TUnaryOp ***********************************
}
function TUnaryOp.AsBoolean: TprBoolean;
begin
  case FOperator of
    opNot: Result:= not(FOperand.AsBoolean)
  else
    Result:= inherited AsBoolean;
  end
end;

function TUnaryOp.AsFloat: TprFloat;
begin
  case FOperator of
    opMinus: Result:= -FOperand.AsFloat;
    opPlus: Result:= FOperand.AsFloat;
  else
    Result:= inherited AsFloat;
  end
end;

function TUnaryOp.AsInteger: TprInt;
begin
  Result:= 0;
  case FOperator of
    opMinus: Result:= -FOperand.AsInteger;
    opPlus: Result:= FOperand.AsInteger;
    opNot:
    case FOperandType of
      ttInteger: Result:= not FOperand.AsInteger;
      ttBoolean: Result:= Integer(AsBoolean);
    else
      Internal(6);
    end;
  else
    Result:= inherited AsInteger;
  end
end;

constructor TUnaryOp.Create(aOperator: TOperator; aOperand: IValue);
begin
  inherited Create;
  FOperand:= aOperand;
  FOperator:= aOperator;
  FOperandType:= FOperand.ExprType;
  if not (FOperator in [opNot, opPlus, opMinus]) then
    raise EExpression.CreateFmt('%s is not simple unary FOperator',
                                [NOperator[FOperator]])
end;

function TUnaryOp.ExprType: TExprType;
begin
  Result:= ResultType(FOperator, FOperandType)
end;

{
********************************** TBinaryOp ***********************************
}
function TBinaryOp.AsBoolean: TprBoolean;
begin
  Result:= false;
  case FOperator of
    opAnd: Result:= FOperand1.AsBoolean and FOperand2.AsBoolean;
    opOr: Result:= FOperand1.AsBoolean or FOperand2.AsBoolean;
    opXor: Result:= FOperand1.AsBoolean xor FOperand2.AsBoolean;
  else
    Internal(13);
  end
end;

function TBinaryOp.AsFloat: TprFloat;
begin
  Result:= 0;
  case ExprType of
    ttFloat:
      case FOperator of
        opExp: Result:= Exp(FOperand2.AsFloat * Ln(FOperand1.AsFloat));
        opPlus: Result:= FOperand1.AsFloat + FOperand2.AsFloat;
        opMinus: Result:= FOperand1.AsFloat - FOperand2.AsFloat;
        opMult: Result:= FOperand1.AsFloat * FOperand2.AsFloat;
        opDivide: Result:= FOperand1.AsFloat / FOperand2.AsFloat;
      else
        Internal(11);
      end;
    ttInteger:
        Result:= AsInteger;
    ttBoolean:
       Result:= Integer(AsBoolean);
  end
end;

function TBinaryOp.AsInteger: TprInt;
begin
  Result:= 0;
  case ExprType of
    ttInteger:
    case FOperator of
      opPlus: Result:= FOperand1.AsInteger + FOperand2.AsInteger;
      opMinus: Result:= FOperand1.AsInteger - FOperand2.AsInteger;
      opMult: Result:= FOperand1.AsInteger * FOperand2.AsInteger;
      opDiv: Result:= FOperand1.AsInteger div FOperand2.AsInteger;
      opMod: Result:= FOperand1.AsInteger mod FOperand2.AsInteger;
      opShl: Result:= FOperand1.AsInteger shl FOperand2.AsInteger;
      opShr: Result:= FOperand1.AsInteger shr FOperand2.AsInteger;
      opAnd: Result:= FOperand1.AsInteger and FOperand2.AsInteger;
      opOr: Result:= FOperand1.AsInteger or FOperand2.AsInteger;
      opXor: Result:= FOperand1.AsInteger xor FOperand2.AsInteger;
    else
      Internal(12);
    end;
    ttBoolean:
      Result:= Integer(AsBoolean);
  end
end;

function TBinaryOp.AsString: TprString;
begin
  Result:= '';
  case ExprType of
    ttString:
      case FOperator of
        opPlus: Result:= FOperand1.AsString + FOperand2.AsString;
      else
        Internal(10);
      end;
    ttFloat:
      Result:= FloatToStr(AsFloat);
    ttInteger:
      Result:= IntToStr(AsInteger);
    ttBoolean:
      Result:= NBoolean[AsBoolean];
  end
end;

constructor TBinaryOp.Create(aOperator: TOperator; aOperand1, aOperand2:
        IValue);
begin
  inherited Create;
  {what if type changes? Operands might be IF expressions!}
  FOperator:= aOperator;
  FOperand1:= aOperand1;
  FOperand2:= aOperand2;
  FOperandType:= CommonType(FOperand1.ExprType, FOperand2.ExprType);
  if not (FOperator in [opExp, opMult..opXor]) then
    raise EExpression.CreateFmt('%s is not a simple binary FOperator',
              [NOperator[FOperator]])
end;

function TBinaryOp.ExprType: TExprType;
begin
  Result:= ResultType(FOperator, FOperandType)
end;

{
******************************** TRelationalOp *********************************
}
function TRelationalOp.AsBoolean: TprBoolean;
begin
  Result:= false;
  case FOperandType of
    ttBoolean:
    case FOperator of
      opEq: Result:= FOperand1.AsBoolean = FOperand2.AsBoolean;
      opNEq: Result:= FOperand1.AsBoolean <> FOperand2.AsBoolean;
    else
      raise EExpression.CreateFmt('cannot apply %s to boolean operands',
                                  [NOperator[FOperator]]);
    end;

    ttInteger:
    case FOperator of
      opLT: Result:= FOperand1.AsInteger < FOperand2.AsInteger;
      opLTE: Result:= FOperand1.AsInteger <= FOperand2.AsInteger;
      opGT: Result:= FOperand1.AsInteger > FOperand2.AsInteger;
      opGTE: Result:= FOperand1.AsInteger >= FOperand2.AsInteger;
      opEq: Result:= FOperand1.AsInteger = FOperand2.AsInteger;
      opNEq: Result:= FOperand1.AsInteger <> FOperand2.AsInteger;
    end;

    // CARH 2014-01-02:
    // Compare float using 'CompareValue' and 'SameValue' routines!
    // These routine are slower, but make sure the 'equal' comparison is done
    // with a minimal tollerance!!!
    ttFloat:
    case FOperator of
    {$IFNDEF USEMATHFLOATCOMPARE}
      opLT: Result:= FOperand1.AsFloat < FOperand2.AsFloat;
      opLTE: Result:= FOperand1.AsFloat <= FOperand2.AsFloat;
      opGT: Result:= FOperand1.AsFloat > FOperand2.AsFloat;
      opGTE: Result:= FOperand1.AsFloat >= FOperand2.AsFloat;
      opEq: Result:= FOperand1.AsFloat = FOperand2.AsFloat;
      opNEq: Result:= FOperand1.AsFloat <> FOperand2.AsFloat;
    {$ELSE}
      opLT: Result := CompareValue(FOperand1.AsFloat, FOperand2.AsFloat) < 0;
      opLTE: Result := CompareValue(FOperand1.AsFloat, FOperand2.AsFloat) <= 0;
      opGT: Result := CompareValue(FOperand1.AsFloat, FOperand2.AsFloat) > 0;
      opGTE: Result := CompareValue(FOperand1.AsFloat, FOperand2.AsFloat) >= 0;
      opEq: Result := SameValue(FOperand1.AsFloat, FOperand2.AsFloat);
      opNEq: Result:= not(SameValue(FOperand1.AsFloat, FOperand2.AsFloat));
    {$ENDIF}
    end;

    ttString:
    case FOperator of
      opLT: Result:= FOperand1.AsString < FOperand2.AsString;
      opLTE: Result:= FOperand1.AsString <= FOperand2.AsString;
      opGT: Result:= FOperand1.AsString > FOperand2.AsString;
      opGTE: Result:= FOperand1.AsString >= FOperand2.AsString;
      opEq: Result:= FOperand1.AsString = FOperand2.AsString;
      opNEq: Result:= FOperand1.AsString <> FOperand2.AsString;
    end;
  end
end;

function TRelationalOp.AsFloat: TprFloat;
begin
  Result:= Integer(AsBoolean)
end;

function TRelationalOp.AsInteger: TprInt;
begin
  Result:= Integer(AsBoolean)
end;

function TRelationalOp.AsString: TprString;
begin
  Result:= NBoolean[AsBoolean]
end;

constructor TRelationalOp.Create(aOperator: TOperator; aOperand1, aOperand2:
        IValue);
begin
  inherited Create;
  FOperator:= aOperator;
  FOperand1:= aOperand1;
  FOperand2:= aOperand2;
  FOperandType:= CommonType(FOperand1.ExprType, FOperand2.ExprType);
  if not (FOperator in RelationalOperators) then
    raise EExpression.CreateFmt('%s is not relational FOperator',
                                 [NOperator[FOperator]])
end;

function TRelationalOp.ExprType: TExprType;
begin
  Result:= ttBoolean
end;

{
******************************* TObjectProperty ********************************
}
function TObjectProperty.AsBoolean: TprBoolean;
begin
  if FPropType = ttBoolean then
    Result:= LongBool(GetOrdProp(FObj, FPropInfo))
  else
    Result:= inherited AsBoolean
end;

function TObjectProperty.AsFloat: TprFloat;
begin
  if FPropType = ttFloat then
    Result:= GetFloatProp(FObj, FPropInfo)
  else
    Result:= inherited AsFloat
end;

function TObjectProperty.AsInteger: TprInt;
begin
  case FPropType of
    ttInteger, ttEnumerated:
      Result:= GetOrdProp(FObj, FPropInfo)
  else
    Result:= inherited AsInteger;
  end
end;

function TObjectProperty.AsObject: TObject;
begin
  if FPropType = ttObject then
    Result:= TObject(GetOrdProp(FObj, FPropInfo))
  else
    Result:= inherited AsObject
end;

function TObjectProperty.AsString: TprString;
begin
  case FPropType of
    ttString: Result:= GetStrProp(FObj, FPropInfo);
    ttEnumerated: Result:= GetEnumName(FPropInfo.PropType^, AsInteger);
  else
    Result:= inherited AsString
  end
end;

constructor TObjectProperty.Create(aObj: IValue; const PropName: String);
begin
  inherited Create;
  FObj:= aObj.AsObject;
  FPropInfo:= GetPropInfo(PTypeInfo(FObj.ClassInfo), PropName);
  if not Assigned(FPropInfo) then
    raise EExpression.CreateFmt('%s is not published property of %s',
                   [PropName, aObj.AsObject.ClassName]);
  case FPropInfo.PropType^^.Kind of
    tkClass: FPropType:= ttObject;
    tkEnumeration:
    if FPropInfo.PropType^^.Name = 'Boolean' then {special case}
      FPropType:= ttBoolean
    else
      FPropType:= ttEnumerated; {not boolean}
    tkInteger, tkChar: FPropType:= ttInteger;
    tkFloat: FPropType:= ttFloat;
    tkString, tkLString, tkWString: FPropType:= ttString;
  else
    raise EExpression.CreateFmt('Property %s unsupported type', [PropName]);
  end
end;

function TObjectProperty.ExprType: TExprType;
begin
  Result:= FPropType
end;

function TObjectProperty.TypeInfo: PTypeInfo;
begin
  Result:= FPropInfo.PropType^
end;

{
******************************** TParameterList ********************************
}
function TParameterList.AddExpression(e: IValue): Integer;
begin
  Result:= FList.Add(Pointer(e));
  e._AddRef;
end;

procedure TParameterList.CheckCount(const Count: Byte);
begin
  // Aantal parameters moet exact zijn
  if GetCount <> Count then
    raise EExpression.CreateFmt('Number of parameters (%d) must be %d', [GetCount, Count]);
end;

procedure TParameterList.CheckCount(const MinCount, MaxCount: Byte);
begin
  // Aantal parameters moet een waarde heben MinCount t/m MaxCount
  // Als MaxCount = 0 dan wordt alleen MinCount gecontroleerd !
  if GetCount < MinCount then
    raise EExpression.CreateFmt('Number of parameters (%d) must be at least %d', [GetCount, MinCount]);

  if (MaxCount > 0) and (GetCount > MaxCount) then
    raise EExpression.CreateFmt('Number of parameters (%d) can be a maximum of %d', [GetCount, MaxCount]);
end;

procedure TParameterList.CheckCount(const Counts: array of Byte);
var
  n, iCount: Integer;
  sValCounts: string;
begin
  // Aantal parameters moet exact aan van de waarden zijn
  iCount := GetCount;
  for n := 0 to High(Counts) do
  begin
    if iCount = Counts[n] then
      Exit; // OK

    if n = 0 then
      sValCounts := IntToStr(Counts[n])
    else
      sValCounts := Format('%s, %d', [sValCounts, Counts[n]]);
  end;

  raise EExpression.CreateFmt('Number of paramters (%d) mut be one of the following'#13'%s', [iCount, sValCounts]);
end;

constructor TParameterList.Create;
begin
  inherited;
  Inc(_ParamListInstanceCount);
  FList := TList.Create;
end;

destructor TParameterList.Destroy;
var
  i: Integer;
begin
  for i:= 0 to (FList.Count - 1) do
    IValue(FList[i])._Release;
  FreeAndNil(FList);
  Dec(_ParamListInstanceCount);
  inherited;
end;

function TParameterList.GetAsBoolean(i: Integer): TprBoolean;
begin
  Result:= GetParam(i).AsBoolean;
end;

function TParameterList.GetAsFloat(i: Integer): TprFloat;
begin
  Result:= GetParam(i).AsFloat;
end;

function TParameterList.GetAsInteger(i: Integer): TprInt;
begin
  Result:= GetParam(i).AsInteger;
end;

function TParameterList.GetAsObject(i: Integer): TObject;
begin
  Result:= GetParam(i).AsObject;
end;

function TParameterList.GetAsString(i: Integer): TprString;
begin
  Result:= GetParam(i).AsString;
end;

function TParameterList.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TParameterList.GetExprType(i: Integer): TExprType;
begin
  Result:= GetParam(i).ExprType;
end;

function TParameterList.GetParam(i: Integer): IValue;
begin
  Result:= IValue(FList[i]);
end;

{
********************************** TFunction ***********************************
}
constructor TFunction.Create(aParameterList: IParameterList);
begin
  inherited Create;
  FParameterList:= aParameterList
end;

function TFunction.GetParam(n: Integer): IValue;
begin
  Result:= FParameterList.Param[n];
end;

function TFunction.GetParameterCount: Integer;
begin
  if Assigned(FParameterList) then
    Result := FParameterList.Count
  else
    Result := 0;
end;

type
  TConditional = class(TFunction)
  private
    FCommonType: TExprType;
    function Rex: IValue;
  protected
    function TestParameters: Boolean; override;
  public
    function AsBoolean: TprBoolean; override;
    function AsFloat: TprFloat; override;
    function AsInteger: TprInt; override;
    function AsString: TprString; override;
    function AsVariant: Variant; override;
    function ExprType: TExprType; override;
    function TypeInfo: PTypeInfo; override;
  end;

  TTypeCast = class(TFunction)
  private
    FOperandType: TExprType;
    FOperator: TExprType;
  protected
    function TestParameters: Boolean; override;
  public
    constructor Create(aParameterList: IParameterList; aOperator: TExprType);
    function AsBoolean: TprBoolean; override;
    function AsFloat: TprFloat; override;
    function AsInteger: TprInt; override;
    function AsObject: TObject; override;
    function AsString: TprString; override;
    function ExprType: TExprType; override;
  end;

  TMF =
    (mfTrunc, mfRound, mfAbs, mfArcTan, mfCos, mfExp, mfFrac, mfInt,
     mfLn, mfPi, mfSin, mfSqr, mfSqrt, mfPower);

  TMathExpression = class(TFunction)
  private
    FOperator: TMF;
  protected
    function TestParameters: Boolean; override;
  public
    constructor Create(aParameterList: IParameterList; aOperator: TMF);
    function AsFloat: TprFloat; override;
    function AsInteger: TprInt; override;
    function ExprType: TExprType; override;
  end;

  TSF =
    (sfUpper, sfLower, sfCopy, sfTrim, sfLTrim, sfRTrim, sfPos, sfLength, sfReplace);

  TStringExpression = class(TFunction)
  private
    FOperator: TSF;
  protected
    function TestParameters: Boolean; override;
  public
    constructor Create(aParameterList: IParameterList; aOperator: TSF);
    function AsInteger: TprInt; override;
    function AsString: TprString; override;
    function ExprType: TExprType; override;
  end;

const
  NTypeCast: array[TExprType] of PChar =
    ('OBJECT', 'STRING', 'FLOAT', 'INTEGER', 'ENUMERATED', 'BOOLEAN');
  NMF: array[TMF] of PChar =
    ('TRUNC', 'ROUND', 'ABS', 'ARCTAN', 'COS', 'EXP', 'FRAC', 'INT',
     'LN', 'PI', 'SIN', 'SQR', 'SQRT', 'POWER');
  NSF: array[TSF] of PChar =
    ('UPPER', 'LOWER', 'COPY', 'TRIM', 'LTRIM', 'RTRIM', 'POS', 'LENGTH', 'REPLACE');

{
********************************* TConditional *********************************
}
function TConditional.AsBoolean: TprBoolean;
begin
  Result:= Rex.AsBoolean
end;

function TConditional.AsFloat: TprFloat;
begin
  Result:= Rex.AsFloat
end;

function TConditional.AsInteger: TprInt;
begin
  Result:= Rex.AsInteger
end;

function TConditional.AsString: TprString;
begin
  Result:= Rex.AsString
end;

function TConditional.AsVariant: Variant;
begin
  Result:= Rex.AsVariant
end;

function TConditional.ExprType: TExprType;
begin
  Result := FCommonType
end;

function TConditional.Rex: IValue;
begin
  if Param[0].AsBoolean then
    Result:= Param[1] else
    Result:= Param[2]
end;

function TConditional.TestParameters: Boolean;

  procedure CheckCommonType(aType1, aType2: TExprType);
  begin
    // Set 'FCommonType'
    if aType1 = aType2 then
      FCommonType := aType1
    // RH: 2014116 - TEnumeratedLiteral.ExprType is now ttEnumerated
    // - to keep this working, set common type to integer (allows exchanging data)
    else if ((aType1 = ttEnumerated) and (aType2 = ttInteger)) or
            ((aType1 = ttInteger) and (aType2 = ttEnumerated)) then
      FCommonType := ttInteger
    else
      raise EExpression.Create('IF options must be the same type');
  end;

begin
  // Basic checks
  if not (ParameterCount = 3) then
    raise EExpression.Create('IF must have 3 parameters');
  if not (Param[0].ExprType = ttBoolean) then
    raise EExpression.Create('First parameter to If must be Boolean');

  // RH: 2014116 - TEnumeratedLiteral.ExprType is now ttEnumerated
  // - to keep this working set common type to integer (allows exchanging data)
  CheckCommonType(Param[1].ExprType, Param[2].ExprType);

  Result:= true
end;

function TConditional.TypeInfo: PTypeInfo;
begin
  // RH: 2014116 - TEnumeratedLiteral.ExprType is now ttEnumerated
  if (Param[1].TypeInfo = Param[2].TypeInfo) then // TypeInfo gelijk - dan een van beide teruggeven
    Result := Param[1].TypeInfo
  else if (FCommonType = ttEnumerated) then // Beide items zijn een 'Enumerated' type
    Result := Rex.TypeInfo // Dan expressie maar uitvoeren :(
  else if (Param[1].ExprType = ttEnumerated) then // Een van beide is 'Enum' - dan integer variant teruggeven
    Result := Param[2].TypeInfo
  else
    Result := Param[1].TypeInfo; // Anders typeinfo van eerste teruggeven
end;

{
********************************** TTypeCast ***********************************
}
function TTypeCast.AsBoolean: TprBoolean;
var
  s: string;

  const
    Eps30 = 1e-30;

begin
  if FOperator = ttBoolean then
  begin
    case FOperandType of
      ttString:
      begin
         s:= Uppercase(Param[0].AsString);
         if s =  NBoolean[false] then
           Result:= False
         else
         if s = NBoolean[true] then
           Result:= True
         else
           raise EExpression.CreateFmt('Cannot convert %s to Boolean', [s])
      end;
      ttFloat:
        Result:= Abs(Param[0].AsFloat) > Eps30;
      ttInteger:
        Result:= Param[0].AsInteger <> 0
    else
      Result:= Param[0].AsBoolean;
    end
  end else
  begin
    Result:= inherited AsBoolean
  end
end;

function TTypeCast.AsFloat: TprFloat;
var
  Code: Integer;
  s: string;
begin
  if FOperator = ttFloat then
  begin
    case FOperandType of
      ttString:
      begin
         s:= Param[0].AsString;
         Val(s, Result, Code);
         if Code <> 0 then
           raise EExpression.CreateFmt('Cannot convert %s to float', [s])
      end;
    else
      Result:= Param[0].AsFloat
    end
  end else
  begin
    Result:= inherited AsFloat
  end
end;

function TTypeCast.AsInteger: TprInt;
var
  Code: Integer;
  s: string;
begin
  if FOperator = ttInteger then
  begin
    case FOperandType of
      ttString:
      begin
         s:= Param[0].AsString;
         Val(s, Result, Code);
         if Code <> 0 then
           raise EExpression.CreateFmt('Cannot convert %s to integer', [s])
      end;
      ttFloat:
      begin
        Result:= Trunc(Param[0].AsFloat);
      end
    else
      Result:= Param[0].AsInteger
    end
  end else
  begin
    Result:= inherited AsInteger
  end
end;

function TTypeCast.AsObject: TObject;
begin
  if FOperator = ttObject then
    Result:= Param[0].AsObject
  else
    Result:= inherited AsObject {almost certainly bomb}
end;

function TTypeCast.AsString: TprString;
begin
  if FOperator = ttString then
  begin
    Result:= Param[0].AsString
  end else
  begin
    Result:= inherited AsString
  end
end;

constructor TTypeCast.Create(aParameterList: IParameterList; aOperator:
        TExprType);
begin
  if aOperator = ttEnumerated then
    raise EExpression.Create('Cannot cast to enumerated');
  if aParameterList.Count = 1 then
    FOperandType:= aParameterList.Param[0].ExprType
  else
    raise EExpression.Create('Invalid parameters to typecast');
  {allow futile cast Object(ObjVar) }
  if (aOperator = ttObject) and
     (FOperandType <> ttObject) then
    IncompatibleTypes(aOperator, FOperandType);

  {objects may be cast to string or object only
   casting to string helplessly returns class name}
  if (FOperandType = ttObject) and
     not ((aOperator = ttObject) or
          (aOperator = ttString)) then
      IncompatibleTypes(aOperator, FOperandType);
  inherited Create(aParameterList);
  FOperator:= aOperator
end;

function TTypeCast.ExprType: TExprType;
begin
  Result:= FOperator
end;

function TTypeCast.TestParameters: Boolean;
begin
  Result:= ParameterCount = 1
end;

{
******************************* TMathExpression ********************************
}
function TMathExpression.AsFloat: TprFloat;
begin
  case FOperator of
    mfAbs: Result:= Abs(Param[0].AsFloat);
    mfArcTan: Result:= ArcTan(Param[0].AsFloat);
    mfCos: Result:= Cos(Param[0].AsFloat);
    mfExp: Result:= Exp(Param[0].AsFloat);
    mfFrac: Result:= Frac(Param[0].AsFloat);
    mfInt: Result:= Int(Param[0].AsFloat);
    mfLn: Result:= Ln(Param[0].AsFloat);
    mfPi: Result:= Pi;
    mfSin: Result:= Sin(Param[0].AsFloat);
    mfSqr: Result:= Sqr(Param[0].AsFloat);
    mfSqrt: Result:= Sqrt(Param[0].AsFloat);
    mfPower: Result:= Exp(Param[1].AsFloat * Ln(Param[0].AsFloat))
  else
    Result:= inherited AsFloat;
  end
end;

function TMathExpression.AsInteger: TprInt;
begin
  case FOperator of
    mfTrunc: Result:= Trunc(Param[0].AsFloat);
    mfRound: Result:= Round(Param[0].AsFloat);
    mfAbs: Result:= Abs(Param[0].AsInteger);
  else
    Result:= inherited AsInteger;
  end
end;

constructor TMathExpression.Create(aParameterList: IParameterList; aOperator:
        TMF);
begin
  inherited Create(aParameterList);
  FOperator:= aOperator
end;

function TMathExpression.ExprType: TExprType;
begin
  case FOperator of
    mfTrunc, mfRound: Result:= ttInteger;
  else
    Result:= ttFloat;
  end
end;

function TMathExpression.TestParameters: Boolean;
begin
  Result:= True;
  case FOperator of
    mfTrunc, mfRound, mfArcTan, mfCos, mfExp, mfFrac, mfInt,
    mfLn, mfSin, mfSqr, mfSqrt, mfAbs:
    begin
      Result:= (ParameterCount = 1) and
           Param[0].CanReadAs(ttFloat);
    end;
    mfPower:
    begin
      Result:= (ParameterCount = 2) and
           Param[0].CanReadAs(ttFloat) and
           Param[1].CanReadAs(ttFloat);
    end;
  end
end;

{
****************************** TStringExpression *******************************
}
function TStringExpression.AsInteger: TprInt;
begin
  case FOperator of
    sfPos: Result:= Pos(Param[0].AsString, Param[1].AsString);
    sfLength: Result:= Length(Param[0].AsString);
  else
    Result:= inherited AsInteger
  end
end;

function TStringExpression.AsString: TprString;

  function GetReplaceFlags: TReplaceFlags;
  var
    I: Byte;
  begin
    if ParameterCount > 3 then
    begin
      I := Param[3].AsInteger;
      Result := TReplaceFlags((@I)^); // Byte-set !!!
    end else
      Result := [];
  end;

begin
  case FOperator of
    sfUpper  : Result := UpperCase(Param[0].AsString);
    sfLower  : Result := LowerCase(Param[0].AsString);
    sfCopy   : Result := Copy(Param[0].AsString, Param[1].AsInteger, Param[2].AsInteger);
    sfTrim   : Result := Trim(Param[0].AsString);
    sfLTrim  : Result := TrimLeft(Param[0].AsString);
    sfRTrim  : Result := TrimRight(Param[0].AsString);
    sfReplace: Result := StringReplace(Param[0].AsString, Param[1].AsString, Param[2].AsString, GetReplaceFlags);
  else
    Result:= inherited AsString;
  end
end;

constructor TStringExpression.Create(aParameterList: IParameterList; aOperator:
        TSF);
begin
  inherited Create(aParameterList);
  FOperator:= aOperator
end;

function TStringExpression.ExprType: TExprType;
begin
  case FOperator of
    sfPos, sfLength: Result:= ttInteger;
  else
    Result:= ttString;
  end
end;

function TStringExpression.TestParameters: Boolean;
begin
  case FOperator of
    sfUpper, sfLower, sfLength, sfTrim, sfLTrim, sfRTrim:
      Result:= (ParameterCount = 1) and
               Param[0].CanReadAs(ttString);
    sfCopy:
      Result:= (ParameterCount = 3) and
            Param[0].CanReadAs(ttString) and
            (Param[1].ExprType = ttInteger) and
            (Param[2].ExprType = ttInteger);
    sfPos:
      Result:= (ParameterCount = 2) and
            Param[0].CanReadAs(ttString) and
            Param[1].CanReadAs(ttString);

    sfReplace:
        Result :=
          (ParameterCount >= 3) and
          (ParameterCount <= 4) and
          Param[0].CanReadAs(ttString) and
          Param[1].CanReadAs(ttString) and
          Param[2].CanReadAs(ttString) and
          ( (ParameterCount = 3) or
            ( Param[3].CanReadAs(ttInteger) and // Flags
              (Param[3].AsInteger >= 0) and
              (Param[3].AsInteger <= 3) ) );
  else
    Result:= False;
  end;
end;

function StandardFunctions (const Ident: String; PL: IParameterList): IValue;
var
  i: TExprType;
  j: TMF;
  k: TSF;
begin
  // IF operator-function
  if Ident = 'IF' then
  begin
    Result := TConditional.Create(PL);
    Exit;
  end;

  // Check for TypeCasts
  for i:= Low(TExprType) to High(TExprType) do
    if Ident = NTypeCast[i] then
  begin
    Result := TTypeCast.Create(PL, i);
    Exit;
  end;

  // Check for Math expression
  for j:= Low(TMF) to High(TMF) do
    if Ident = NMF[j] then
  begin
    Result := TMathExpression.Create(PL, j);
    Exit;
  end;

  // Check for String Expression
  for k:= Low(TSF) to High(TSF) do
    if Ident = NSF[k] then
  begin
    Result := TStringExpression.Create(PL, k);
    Exit;
  end;

  // Not a standard function
  Result:= nil
end;

{parser...}
const
{note: These two cannot be the same}
  DecSeparator = '.';
  ParamDelimiter = ',';

  {$IF CompilerVersion >= 22.0}  // XE and later (float formatsettings)
  FS: TFormatSettings = (DecimalSeparator: DecSeparator);
  {$IFEND}

  OpTokens: array[TOperator] of PChar =
              ( 'NOT',
                '^',
                '*', '/', 'DIV', 'MOD', 'AND', 'SHL', 'SHR',
                '+', '-', 'OR', 'XOR',
                '=', '<>', '<', '>', '<=', '>=');

  Whitespace = [#$1..#$20];
  SignChars = ['+', '-'];
  RelationalChars = ['<', '>', '='];
  OpChars = SignChars + ['^', '/', '*'] + RelationalChars;

  OpenSub = '(';
  CloseSub = ')';
  SQuote = '''';

  ExprDelimiters = [#0, CloseSub, ParamDelimiter];

  {mst}
  SHex = '$';
  HexDigs = Digits+['a'..'f','A'..'F'];
  {mst}

procedure SwallowWhitespace( var P: PChar);
begin
  while P^ in Whitespace do inc(P)
end;

function EoE( var P: PChar): Boolean;
begin
  Result:= (P^ in ExprDelimiters)
end;


function GetOperator( var P: PChar; var Operator: TOperator): Boolean;
{this leaves p pointing to next char after operator}
var
  Buf: array[0..3] of Char;
  lp: PChar;
  i: Integer;

function tt( op: TOperator): Boolean;
begin
  if StrLComp(Buf, OpTokens[Op], i) = 0 then
  begin
    Operator:= op;
    Result:= true
  end else
  begin
    Result:= false
  end
end;

begin
  Result:= false;
  if P^ in OpChars then
  begin
    Result:= true;
    Buf[0]:= P^;
    Inc(P);
    case Buf[0] of
      '*': Operator:= opMult;
      '+': Operator:= opPlus;
      '-': Operator:= opMinus;
      '/': Operator:= opDivide;
      '<': if P^ = '=' then
           begin
             Operator:= opLTE;
             Inc(P)
           end else
           if P^ = '>' then
           begin
             Operator:= opNEq;
             Inc(P)
           end else
           begin
             Operator:= opLT
           end;
       '=': Operator:= opEq;
       '>': if P^ = '=' then
            begin
              Operator:= opGTE;
              Inc(P)
            end else
            begin
              Operator:= opGT
            end;
      '^': Operator:= opExp;
    end
  end else
  if UpCase(P^) in ['A', 'D', 'M', 'N', 'O', 'S', 'X'] then
  begin  {check for 'identifer' style operators. We can ignore NOT}
    lp:= P;
    i:= 0;
    while (i <= 3) and (lp^ in IdentChars) do
    begin
      Buf[i]:= UpCase(lp^);
      inc(lp);
      inc(i)
    end;
    if i in [2,3] then
    begin
      if tt(opNot) then
        Result:= true
      else
      if tt(opDiv) then
        Result:= true
      else
      if tt(opMod) then
        Result:= true
      else
      if tt(opAnd) then
        Result:= true
      else
      if tt(opShl) then
        Result:= true
      else
      if tt(opShr) then
        Result:= true
      else
      if tt(opOr) then
        Result:= true
      else
      if tt(opXor) then
        Result:= true
    end;
    if Result then
      inc(P, i)
  end
end;

type
  TExprFunc = function( var P: PChar; IDF: TIdentifierFunction): IValue;

function Chain(var P: PChar; IDF: TIdentifierFunction;
                   NextFunc: TExprFunc; Ops: TOperators): IValue;
{this function is used to construct a chain of expressions}
var
  NextOpr: TOperator;
  StopF: Boolean;
  lp: PChar;
begin
  StopF:= false;
  Result:= NextFunc(P, IDF);
  try
    repeat
      SwallowWhitespace(P);
      lp:= P;
      if not EoE(P) and GetOperator(lp, NextOpr) and (NextOpr in Ops) then
      begin
        P:= lp;
        if NextOpr in RelationalOperators then
          Result:= TRelationalOp.Create(NextOpr, Result, NextFunc(P, IDF))
        else
          Result:= TBinaryOp.Create(NextOpr, Result, NextFunc(P, IDF))
      end else
      begin
        StopF:= true
      end
    until StopF
  except
    Result:= nil;
    raise
  end
end;

function Expression( var P: PChar; IDF: TIdentifierFunction): IValue; forward;


function SimpleFactor( var P: PChar; IDF: TIdentifierFunction): IValue;

  function UnsignedNumber: IValue;
  type
    TNScan = (nsMantissa, nsDPFound, nsExpFound, nsFound);
  var
    S: String;
    State: TNScan;
    Int: Boolean;
    {$IF CompilerVersion < 22.0}  // Pre XE
    SaveSep: Char;
    {$IFEND}

    procedure Bomb;
    begin
      raise EExpression.Create('Bad numeric format')
    end;

  begin
    S:= '';
    Int:= false;
    State:= nsMantissa;
    repeat
      if P^ in Digits then
      begin
        S:= S + P^;
        inc(P)
      end else
      if P^ = DecSeparator then
      begin
        if State = nsMantissa then
        begin
          S:= S + P^;
          inc(P);
          State:= nsDPFound
        end else
        begin
          Bomb
        end;
      end else
      if (P^ = 'e') or (P^ = 'E') then
      begin
        if (State = nsMantissa) or
           (State = nsDPFound) then
        begin
          S:= S + 'E';
          inc(P);
          if P^ = '-' then
          begin
            S:= S + P^;
            inc(P)
          end;
          State:= nsExpFound;
          if not (P^ in Digits) then
            Bomb
        end else
        begin
          Bomb
        end
      end else
      begin
        Int:= (State = nsMantissa);
        State:= nsFound
      end;
      if Length(S) > 28 then
        Bomb
    until State = nsFound;
    if Int then
    begin
      Result:= TIntegerLiteral.Create(StrToInt(S))
    end else
    begin
      {$IF CompilerVersion >= 22.0}  // XE and later...
      Result:= TFloatLiteral.Create(StrToFloat(S, FS));
      {$ELSE}
      {WATCH OUT if you are running another thread
       which might refer to DecimalSeparator &&&}
      SaveSep:= SysUtils.DecimalSeparator;
      SysUtils.DecimalSeparator:= DecSeparator;
      try
        Result:= TFloatLiteral.Create(StrToFloat(S))
      finally
        SysUtils.DecimalSeparator:= SaveSep
      end
      {$IFEND}
    end
  end;

  function CharacterString: IValue;
  var
    SR: String;
  begin
    SR:= '';
    repeat
      inc(P);
      if P^ = SQuote then
      begin
       inc(P);
        if P^ <> SQuote then
          break;
       end;
       if P^ = #0 then
         raise EExpression.Create('Unterminated string');
       if Length(SR) > MaxStringLength then
         raise EExpression.Create('String too long');
       SR:= SR + P^;
    until false;
    Result:= TStringLiteral.Create(SR)
  end;

{mst}
  function HexValue : IValue;
  var
    SR: String;
  begin
    SR:= '';
    repeat
      inc(P);
      if Length(SR) > MaxStringLength then
        raise EExpression.Create('Hex string too long');
      if not (P^ in HexDigs) then break;
        SR:= SR + P^
    until False;
    try
      Result:= TintegerLiteral.Create(StrToInt(SHex+SR))
    except
      raise EExpression.Create('Invalid char in hex number')
    end;
  end;
{mst}

var
  Identifier: String;
  Operator: TOperator;
  PList: IParameterList;
  MoreParameters: Boolean;

begin {simple factor}
  Result:= nil;
  try
    SwallowWhitespace(P);
    if GetOperator(P, Operator) then
    begin
      case Operator of
        opPlus:
          Result:= TUnaryOp.Create(opPlus, SimpleFactor(P, IDF));
        opMinus:
          Result:= TUnaryOp.Create(opMinus, SimpleFactor(P, IDF));
        opNot:
          Result:= TUnaryOp.Create(opNot, SimpleFactor(P, IDF));
      else
        raise EExpression.CreateFmt('%s not allowed here', [NOperator[Operator]]);
      end;
    end else
    if P^ = SQuote then
    begin
      Result:= CharacterString;
    end else
    {mst}
    if P^ = SHex then
    begin
      Result:= HexValue;
    end else
    {mst}
    if P^ in Digits then
    begin
      Result:= UnsignedNumber;
    end else
    if P^ = OpenSub then
    begin
      Inc(P);
      Result:= Expression(P, IDF);
      {K Friesen's bug 2. Expression may be nil if
      factor = (). Note: this may also apply to
      parameters i.e. Func(x ,, y)}
      if Result = nil then
        raise EExpression.Create('invalid sub-expression');
      if P^ = CloseSub then
        inc(P)
      else
        raise EExpression.Create(' ) expected')
    end else
    if P^ in PrimaryIdentChars then
    begin
      Identifier:= '';
      while P^ in IdentChars do
      begin
        Identifier:= Identifier + UpCase(P^);
        inc(P)
      end;
      if Identifier = 'TRUE' then
      begin
        Result:= TBooleanLiteral.Create(true)
      end else
      if Identifier = 'FALSE' then
      begin
        Result:= TBooleanLiteral.Create(false)
      end else
      begin
        PList:= nil;
        try
          SwallowWhitespace(P);
          MoreParameters:= P^ = OpenSub;
          if MoreParameters then
          begin
            PList:= TParameterList.Create;
            while MoreParameters do
            begin
              inc(P);
              PList.AddExpression(Expression(P, IDF));
              MoreParameters:= P^ = ParamDelimiter
            end;
            {bug fix 11/11/97}
            if P^ = CloseSub then
              Inc(P)
            else
              raise EExpression.Create('Incorrectly formed parameters')
          end;
          if not DisableStandardFunctionProcessing then
            Result:= StandardFunctions(Identifier, PList);
          if (Result = nil) and Assigned(IDF) then
            Result:= IDF(Identifier, PList);
          if Result = nil then
            raise EExpression.CreateFmt('Unknown Identifier %s', [Identifier])
          else
          if not Result.TestParameters then
            raise EExpression.CreateFmt('Invalid parameters to %s', [Identifier])
        finally
          // Remove reference directly
          PList := nil;
        end
      end
    end else
    if EoE(P) then
    begin
      raise EExpression.Create('Unexpected end of factor')
    end else
    begin
      raise EExpression.Create('Syntax error') {leak here ?}
    end
  except
    Result:= nil;
    raise
  end
end;  {Simplefactor}

function ObjectProperty( var P: PChar; IDF: TIdentifierFunction): IValue;
var
  PropName: String;
begin
  SwallowWhitespace(P);
  Result:= SimpleFactor(P, IDF);
  SwallowWhitespace(P);
  while (Result.ExprType = ttObject) and (P^ = '.') do
  begin
    Inc(P);
    SwallowWhitespace(P);
    if P^ in PrimaryIdentChars then
    begin
      PropName:= '';
      while P^ in IdentChars do
      begin
        PropName:= PropName + P^;
        inc(P)
      end;
      Result:= TObjectProperty.Create(Result, PropName)
    end else
    begin
      raise EExpression.CreateFmt('Invalid property of object %s', [Result.AsObject.ClassName])
    end
  end;
end;

function Factor( var P: PChar; IDF: TIdentifierFunction): IValue;
begin
  Result:= Chain(P, IDF, ObjectProperty, [opExp])
end;

function Term( var P: PChar; IDF: TIdentifierFunction): IValue;
begin
  Result:= Chain(P, IDF, Factor, [opMult, opDivide, opDiv, opMod, opAnd, opShl, opShr])
end;

function Simple( var P: PChar; IDF: TIdentifierFunction): IValue;
begin
  Result:= Chain(P, IDF, Term, [opPlus, opMinus, opOr, opXor])
end;

function Expression( var P: PChar; IDF: TIdentifierFunction): IValue;
begin
  Result:= Chain(P, IDF, Simple, RelationalOperators)
end;

function CreateExpression(const S: String; IdentifierFunction:
    TIdentifierFunction; const DisableStandardFunc: boolean = False): IValue;
var
  P:PChar;
begin
  DisableStandardFunctionProcessing := DisableStandardFunc;

  P:= PChar(S);
  Result:= Expression(P, IdentifierFunction);
  if P^ <> #0 then
  begin
    Result:= nil;
    raise EExpression.CreateFmt('%s not appropriate', [P^])
  end
end;


function FoldConstant( Value: IValue): IValue;
begin
  if Assigned(Value) then
  case Value.ExprType of
    ttObject: Result:= TObjectRef.Create(Value.AsObject);
    ttString: Result:= TStringLiteral.Create(Value.AsString);
    ttFloat: Result:= TFloatLiteral.Create(Value.AsFloat);
    ttInteger: Result:= TIntegerLiteral.Create(Value.AsInteger);
    ttEnumerated: Result:= TEnumeratedLiteral.Create(Value.TypeInfo, Value.AsInteger);
    ttBoolean: Result:= TBooleanLiteral.Create(Value.AsBoolean);
  else
    Result:= nil
  end
end;

procedure ShowError;
const
  sFmtICountErr =
    'Not al interfaces/objects have been released.'#13 +
    '- Number of TExpression (IValue) references: %d'#13 +
    '- Number of TParameterList references: %d';
var
  s: string;
begin
  s := Format(sFmtICountErr, [_ExprInstanceCount, _ParamListInstanceCount]);
{$IFDEF MSWINDOWS}
  MessageBox(GetDesktopWindow, PChar(s), 'Warning', MB_OK or MB_ICONWARNING);
{$ENDIF}
{$IFDEF LINUX}
  if TTextRec(ErrOutput).Mode = fmOutput then
    Flush(ErrOutput);
  __write(STDERR_FILENO, PChar(s), Lenght(s));
{$ENDIF}
end;

initialization

finalization

  if WarnForInstanceCountGreaterZeroOnUninitialize and
     ((_ExprInstanceCount <> 0) or (_ParamListInstanceCount <> 0)) then
    ShowError;
end.
