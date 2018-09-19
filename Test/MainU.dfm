object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Test ExprEval'
  ClientHeight = 434
  ClientWidth = 483
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  DesignSize = (
    483
    434)
  PixelsPerInch = 96
  TextHeight = 13
  object lblExpression: TLabel
    Left = 8
    Top = 8
    Width = 52
    Height = 13
    Caption = 'Expression'
    FocusControl = mmExpression
  end
  object lblResult: TLabel
    Left = 8
    Top = 325
    Width = 30
    Height = 13
    Anchors = [akLeft, akBottom]
    Caption = 'Result'
    FocusControl = mmResult
    ExplicitTop = 160
  end
  object mmExpression: TMemo
    Left = 8
    Top = 24
    Width = 467
    Height = 280
    Anchors = [akLeft, akTop, akRight, akBottom]
    ScrollBars = ssBoth
    TabOrder = 0
    WordWrap = False
  end
  object mmResult: TMemo
    Left = 8
    Top = 341
    Width = 467
    Height = 85
    Anchors = [akLeft, akRight, akBottom]
    Lines.Strings = (
      '')
    ScrollBars = ssVertical
    TabOrder = 1
  end
  object btnRun: TButton
    Left = 400
    Top = 310
    Width = 75
    Height = 25
    Action = actnExecute
    Anchors = [akRight, akBottom]
    TabOrder = 2
  end
  object ActionList1: TActionList
    Left = 176
    Top = 112
    object actnExecute: TAction
      Caption = 'Execute'
      ShortCut = 120
      OnExecute = actnExecuteExecute
    end
  end
  object FormStorage: TJvFormStorage
    AppStorage = appStorage
    AppStoragePath = '%FORM_NAME%\'
    StoredProps.Strings = (
      'mmExpression.Lines')
    StoredValues = <>
    Left = 248
    Top = 112
  end
  object appStorage: TJvAppIniFileStorage
    StorageOptions.BooleanStringTrueValues = 'TRUE, YES, Y'
    StorageOptions.BooleanStringFalseValues = 'FALSE, NO, N'
    SubStorages = <>
    Left = 320
    Top = 112
  end
end
