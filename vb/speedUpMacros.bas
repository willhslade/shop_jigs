Attribute VB_Name = "speedUpMacros"
Option Explicit
 
Public glb_origCalculationMode As Integer
 
'http://vbaexpress.com/kb/getarticle.php?kb_id=1035
Sub SpeedOn(Optional StatusBarMsg As String = "Running macro...")
    glb_origCalculationMode = Application.Calculation
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Cursor = xlWait
        .StatusBar = StatusBarMsg
        .EnableCancelKey = xlErrorHandler
    End With
End Sub
 
Sub SpeedOff()
    With Application
        .Calculation = glb_origCalculationMode
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .CalculateBeforeSave = True
        .Cursor = xlDefault
        .StatusBar = False
        .EnableCancelKey = xlInterrupt
    End With
End Sub

