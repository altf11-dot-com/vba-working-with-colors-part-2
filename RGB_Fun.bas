Attribute VB_Name = "RGB_Fun"
Option Explicit
Private r As Integer, g As Integer, b As Integer
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Private Sub Init()

    'random beginning numbers
'    r = Int(Rnd() * 256)
'    g = Int(Rnd() * 256)
'    b = Int(Rnd() * 256)

    'begin black
    r = 255
    g = 255
    b = 0

End Sub


Sub RGB_Fun()
    Dim n0 As Long, upperLimit As Long
    ActiveWorkbook.Save
    [a1].Activate
    Init
    Cells(3, 1).Value = r
    Cells(3, 3).Value = g
    Cells(3, 5).Value = b
    ColorMeWorld r, g, b
    DoEvents
    Application.Calculate
    Application.Wait DateAdd("s", 3, Now)
    upperLimit = 256
    n0 = 0
    Do
        n0 = n0 + 1
        r = r + IIf(r < 255, 1, 0)
        g = g + IIf(g < 255, 1, 0)
        b = b + IIf(b < 255, 1, 0)
        If r < 0 Then r = 0
        If g < 0 Then g = 0
        If b < 0 Then b = 0
        ColorMeWorld r, g, b
        If n0 Mod 10 = 0 Then
            Application.Calculate
            DoEvents
        End If
        [A4].Value = n0
        DoEvents
        Sleep 20
        If r = 255 And g = 255 And b = 255 Then n0 = upperLimit
    Loop While n0 < upperLimit
    Beep
End Sub

Private Sub ColorMeWorld(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer)
    Cells(1, 1).Interior.Color = RGB(r, 0, 0)
    Cells(3, 1).Value = r
    Cells(1, 3).Interior.Color = RGB(0, g, 0)
    Cells(3, 3).Value = g
    Cells(1, 5).Interior.Color = RGB(0, 0, b)
    Cells(3, 5).Value = b
    Cells(1, 7).Interior.Color = RGB(r, g, b)
    Cells(3, 7).Formula = r & ", " & g & ", " & b
End Sub

