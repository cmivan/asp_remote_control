object FormBox: TFormBox
  Left = 869
  Top = 272
  Width = 243
  Height = 116
  AlphaBlend = True
  Anchors = [akLeft, akTop, akRight, akBottom]
  Caption = 'FormBox'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  ScreenSnap = True
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Memo1: TMemo
    Left = 8
    Top = 32
    Width = 217
    Height = 41
    BiDiMode = bdRightToLeftNoAlign
    BorderStyle = bsNone
    ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
    ParentBiDiMode = False
    ScrollBars = ssVertical
    TabOrder = 0
  end
  object Edit1: TEdit
    Left = 8
    Top = 8
    Width = 153
    Height = 21
    ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
    TabOrder = 1
    Text = 'C:\'
  end
  object RUN_TIMER: TTimer
    Interval = 2000
    OnTimer = RUN_TIMERTimer
    Left = 200
  end
  object TimerLine: TTimer
    Enabled = False
    Interval = 10
    OnTimer = TimerLineTimer
    Left = 168
  end
end
