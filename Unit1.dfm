object Form1: TForm1
  Left = 320
  Top = 118
  BorderStyle = bsSingle
  Caption = 'ExcelToSQL'
  ClientHeight = 257
  ClientWidth = 354
  Color = clWindow
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 24
    Width = 84
    Height = 13
    Caption = #1054#1090#1082#1088#1099#1090#1099#1081' '#1092#1072#1081#1083':'
  end
  object Label2: TLabel
    Left = 224
    Top = 88
    Width = 106
    Height = 13
    Caption = #1057#1087#1080#1089#1086#1082' '#1089#1086#1074#1087#1072#1076#1077#1085#1080#1081': '
  end
  object Label3: TLabel
    Left = 8
    Top = 88
    Width = 124
    Height = 13
    Caption = #1042#1089#1077#1075#1086' '#1085#1086#1084#1077#1088#1086#1074' '#1074' '#1092#1072#1081#1083#1077':'
  end
  object Label4: TLabel
    Left = 8
    Top = 112
    Width = 81
    Height = 13
    Caption = #1042' '#1073#1072#1079#1091' '#1074#1085#1077#1089#1077#1085#1086':'
  end
  object Label5: TLabel
    Left = 8
    Top = 136
    Width = 106
    Height = 13
    Caption = #1057#1086#1074#1087#1072#1076#1077#1085#1080#1081' '#1089' '#1073#1072#1079#1086#1081':'
  end
  object StringGrid1: TStringGrid
    Left = 8
    Top = 240
    Width = 617
    Height = 217
    TabStop = False
    RowCount = 3
    FixedRows = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSizing, goColSizing, goRowMoving, goColMoving]
    TabOrder = 0
    Visible = False
    ColWidths = (
      64
      64
      97
      64
      64)
    RowHeights = (
      28
      24
      24)
  end
  object Button1: TButton
    Left = 176
    Top = 336
    Width = 75
    Height = 25
    Caption = 'Open File'
    TabOrder = 1
    Visible = False
    OnClick = Button1Click
  end
  object Edit1: TEdit
    Left = 96
    Top = 24
    Width = 137
    Height = 21
    BorderStyle = bsNone
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlue
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ParentShowHint = False
    ReadOnly = True
    ShowHint = True
    TabOrder = 2
  end
  object Button2: TButton
    Left = 240
    Top = 16
    Width = 105
    Height = 33
    Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1074' '#1073#1072#1079#1091
    Enabled = False
    TabOrder = 3
    OnClick = Button2Click
  end
  object Memo1: TMemo
    Left = 224
    Top = 112
    Width = 121
    Height = 121
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 4
  end
  object Edit4: TEdit
    Left = 136
    Top = 88
    Width = 57
    Height = 21
    BorderStyle = bsNone
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlue
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ParentShowHint = False
    ReadOnly = True
    ShowHint = True
    TabOrder = 5
  end
  object Edit5: TEdit
    Left = 96
    Top = 112
    Width = 57
    Height = 21
    BorderStyle = bsNone
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlue
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ParentShowHint = False
    ReadOnly = True
    ShowHint = True
    TabOrder = 6
  end
  object Edit6: TEdit
    Left = 120
    Top = 136
    Width = 57
    Height = 21
    BorderStyle = bsNone
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlue
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ParentShowHint = False
    ReadOnly = True
    ShowHint = True
    TabOrder = 7
  end
  object Edit2: TEdit
    Left = 8
    Top = 208
    Width = 161
    Height = 21
    TabOrder = 8
    Visible = False
  end
  object ProgressBar1: TProgressBar
    Left = 8
    Top = 56
    Width = 337
    Height = 17
    TabOrder = 9
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 238
    Width = 354
    Height = 19
    Panels = <>
  end
  object Button3: TButton
    Left = 8
    Top = 176
    Width = 75
    Height = 25
    Caption = #1054#1073#1088#1072#1073#1086#1090#1082#1072
    TabOrder = 11
    Visible = False
    OnClick = Button3Click
  end
  object OpenDialog1: TOpenDialog
    Filter = '*.xls'
    Left = 136
    Top = 336
  end
  object ADODataSet1: TADODataSet
    ConnectionString = 
      'Provider=MSDASQL.1;Password=as-12;Persist Security Info=True;Use' +
      'r ID=kontroler;Data Source=mymagn;Mode=Read;Initial Catalog=obje' +
      'cts'
    Parameters = <>
    Left = 104
    Top = 336
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=MSDASQL.1;Password=as-12;Persist Security Info=True;Use' +
      'r ID=kontroler;Data Source=mymagn;Initial Catalog=objects'
    Provider = 'MSDASQL.1'
    Left = 40
    Top = 336
  end
  object ADOCommand1: TADOCommand
    Connection = ADOConnection1
    Parameters = <>
    Left = 72
    Top = 336
  end
  object ADOQuery1: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 8
    Top = 336
  end
  object MainMenu1: TMainMenu
    Left = 256
    Top = 336
    object File1: TMenuItem
      Caption = #1060#1072#1081#1083
      object OpenFile1: TMenuItem
        Caption = #1054#1090#1082#1088#1099#1090#1100
        OnClick = OpenFile1Click
      end
      object N1: TMenuItem
        Caption = '-'
      end
      object Exit1: TMenuItem
        Caption = #1042#1099#1093#1086#1076
        OnClick = Exit1Click
      end
    end
    object Edit3: TMenuItem
      Caption = #1055#1088#1072#1074#1082#1072
      object N2: TMenuItem
        Caption = #1054#1095#1080#1089#1090#1080#1090#1100
        OnClick = N2Click
      end
    end
  end
end
