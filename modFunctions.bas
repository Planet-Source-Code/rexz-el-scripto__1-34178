Attribute VB_Name = "modFunctions"
Declare Function GetTickCount Lib "kernel32" () As Long
Global Compiler As New Process

Public Sub Pause(HowLong As Long) ' HowLong interval i ms.
 Dim lngEnd As Long
 
lngEnd = GetTickCount() + HowLong
Do

DoEvents

Loop Until GetTickCount() >= lngEnd
End Sub
Function FindPart(lzStr As String, mPart As String) As Integer
Dim TPos As Integer
    TPos = InStr(lzStr, mPart)
    If TPos Then
        FindPart = 1
    Else
        FindPart = 0
    End If
    
End Function
Function FindPoint(lzStr As String, mPart As String) As Integer
Dim Xpos As Integer
    Xpos = InStr(lzStr, mPart)
    If Xpos > 0 Then
        FindPoint = Xpos
    Else
        FindPoint = 0
    End If
    
End Function
Function SetColor(lzColor As String) As Long
Select Case lzColor
Case "Green"
SetColor = vbGreen
Case "Red"
SetColor = vbRed
Case "Blue"
SetColor = vbBlue
Case "Yellow"
SetColor = vbYellow
Case "White"
SetColor = vbWhite
Case "Black"
SetColor = vbBlack
Case "Grey"
SetColor = RGB(125, 125, 125)
Case Else
SetColor = lzColor
End Select
End Function

Public Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FileExists = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
    Exit Function
End Function
