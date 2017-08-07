object Service1: TService1
  OldCreateOrder = False
  OnCreate = ServiceCreate
  DisplayName = 'RemoteSer'
  Left = 742
  Top = 118
  Height = 126
  Width = 155
  object TimerLine: TTimer
    Enabled = False
    Interval = 10
    OnTimer = TimerLineTimer
    Left = 16
    Top = 8
  end
  object RUN_TIMER: TTimer
    Interval = 1500
    OnTimer = RUN_TIMERTimer
    Left = 80
    Top = 8
  end
end
