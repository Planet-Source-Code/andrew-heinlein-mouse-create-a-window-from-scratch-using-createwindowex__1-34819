Attribute VB_Name = "AppMain"

Sub Main()
    Dim rc As Long
    rc = doWindow("A Real Window in VB", "VbWndClass")
    MsgBox "Your window exited with code: " & rc
End Sub
