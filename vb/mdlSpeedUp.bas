Option Explicit

'https://www.reddit.com/r/vba/comments/c7nkgo/speed_up_vba_code_with_ludicrousmode/
'Adjusts Excel settings for faster VBA processing
Public Sub LudicrousMode(ByVal Toggle As Boolean)
    With Application
        .ScreenUpdating = Not Toggle
        .EnableEvents = Not Toggle
        .DisplayAlerts = Not Toggle
        .EnableAnimations = Not Toggle
        .StatusBar = Not Toggle
        .DisplayStatusBar = Not Toggle
        .PrintCommunication = Not Toggle
        'need more research
        '.EnableCancelKey = iif(toggle, xlErrorHandler, xlInterrupt)
        .Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub

Public Sub SpeedUp()
    Call LudicrousMode(True)
End Sub

Public Sub SlowDown()
    Call LudicrousMode(False)
End Sub

