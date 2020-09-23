VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim sLine As String, nCurrentLine As Integer
Public VarNames As New Collection
Private Variables As New Collection
Dim inputvalue As Boolean
Dim Feed() As String, Command As String
Dim arTemp() As String, arTemp2() As String
Dim ScriptForm(1 To 10) As frmWindow
Public Script As String
Dim Skip As Boolean
Dim ControlCode As String, filename As String
Public Event LineChange(nLine As Integer, sContents As String)

Public Function LoopApp()
If FindPart(Script, "noend") = 1 Then
theend = True
Else
theend = False
End If
End Function

Public Function Execute()
On Error Resume Next
    Dim iTemp As Integer, iTemp2 As Integer, i As Integer
    Dim sTemp As String, sTemp2 As String
    
    Dim bloop As Boolean
    Dim loopplace As Integer, loopname As String
    Dim IfSkip As Boolean
    Dim StopLabel As String

    Dim StrBuff As String, Keyword As String, SingleWord As String, statement As String
    Dim Xpos As String, YPos As String, color As String, value As Long
    
    Dim InsideLoop As New Collection
    
    arTemp = Split(Script, vbCrLf)
    frmMain.txtDebug.Text = ""
    Errors = 0
    
        For nCurrentLine = LBound(arTemp) To UBound(arTemp)
        sLine = arTemp(nCurrentLine)
        If sLine = "" Then GoTo skipcurrentdo:
        RaiseEvent LineChange(nCurrentLine, sLine)
        
        StrBuff = sLine
        StrBuff = Trim(StrBuff)
        StrBuff = Replace(StrBuff, Chr(9), "")
        statement = "": Xpos = "": YPos = "": color = "": value = 0
        
        iTemp = InStr(1, sLine, "'")
        
        If iTemp <> 0 Then
            sLine = Mid(sLine, 1, iTemp - 1)
            If sLine = "" Then GoTo skipcurrentdo:
        End If
        
                If StopLabel <> "" Then
            If sLine = StopLabel Then
                StopLabel = ""
            End If
            GoTo skipcurrentdo:
        End If
        
        ' Check for end of a block
        If Skip = True Then
            If SpaceTrim(LCase(sLine)) = "exit" Then
                Skip = False
            End If
            GoTo skipcurrentdo:
        End If
        
                If IfSkip = True Then
            Select Case SpaceTrim(LCase(sLine))
                Case "endif"
                    IfSkip = False
                Case "else"
                    IfSkip = False
            End Select
            GoTo skipcurrentdo:
        Else
            If SpaceTrim(LCase(sLine)) = "else" Then
                Skip = True
                GoTo skipcurrentdo:
            End If
        End If
        
                If LCase(sLine) = "exit" And InsideLoop.Count > 0 Then
            If Eval(Mid(InsideLoop(InsideLoop.Count), InStr(1, InsideLoop(InsideLoop.Count), ":") + 1)) Then
                nCurrentLine = CInt(Mid(InsideLoop(InsideLoop.Count), 1, InStr(1, InsideLoop(InsideLoop.Count), ":") - 1)) - 1
            Else
                InsideLoop.Remove InsideLoop.Count
            End If
            GoTo skipcurrentdo:
        End If
        
        iTemp = InStr(1, sLine, " ")
        
                    Select Case LCase(Mid(sLine, 1, iTemp - 1))
                Case "if" ' Need I explain?
                    If Not Eval(Mid(sLine, iTemp + 1)) Then
                        IfSkip = True
                    End If
                    
                    GoTo skipcurrentdo:
                Case "goto" ' Goto a label
                    StopLabel = Trim(Mid(sLine, iTemp + 1))
                    
                    GoTo skipcurrentdo:
                Case "dim" ' Create variable
                    arTemp2 = Split(Trim(Mid(sLine, iTemp + 1)), ",")
                    For i = LBound(arTemp2) To UBound(arTemp2)
                        VarNames.Add Trim(arTemp2(i))
                        Variables.Add ""
                    Next i
                    GoTo skipcurrentdo:
                Case "while" ' Loop
                    If Eval(Mid(sLine, iTemp + 1)) Then
                        InsideLoop.Add nCurrentLine + 1 & ":" & Trim(Mid(sLine, iTemp + 1))
                    Else
                        Skip = True
                    End If
                    GoTo skipcurrentdo:
            End Select
        
                For i = 1 To Len(sLine)
            Select Case Mid(sLine, i, 1)
                Case "=" ' Assignment
                    ChangeVar Trim(Mid(sLine, 1, InStr(1, sLine, "=") - 1)), Equation(Mid(sLine, InStr(1, sLine, "=") + 1))
                    GoTo skipcurrentdo:
                Case "~" ' Assignment from Function
                    ChangeVar Trim(Mid(sLine, 1, InStr(1, sLine, "~") - 1)), DoFunction(Mid(sLine, InStr(1, sLine, "~") + 1))
                    GoTo skipcurrentdo:
                Case "(" ' Procedure Call
                    DoFunction sLine
                    GoTo skipcurrentdo:
            End Select
        Next i
        
skipcurrentdo:
Next nCurrentLine
Dim a
If theend = False Then
For a = 1 To UBound(ScriptForm())
ScriptForm(a).CurrentX = 10
ScriptForm(a).CurrentY = ScriptForm(a).Height - 20
ScriptForm(a).ForeColor = Default
ScriptForm(a).BackColor = Default
ScriptForm(a).Cls
ScriptForm(a).Print "Program has ended, press any key to continue"
Next a
Else
End If
For a = 1 To VarNames.Count
VarNames.Remove a
Variables.Remove a
Next a

If Errors > 1 Or Errors = 0 Then
frmMain.Add vbCrLf & "Totally " & Errors & " Errors"
Else
frmMain.Add vbCrLf & "Totally " & Errors & " Error"
End If
End Function

Function SpaceTrim(Code As String) As String
SpaceTrim = Replace(Code, " ", "")
End Function

Function Convert2Color(str As String) As Long
Select Case LCase(str)
Case "black"
Convert2Color = vbBlack
Case "white"
Convert2Color = vbWhite
Case "green"
Convert2Color = vbGreen
Case "blue"
Convert2Color = vbBlue
Case "red"
Convert2Color = vbRed
Case "yellow"
Convert2Color = vbYellow
Case Else
Convert2Color = CLng(str)
End Select
End Function

Function ExitExecute()
nCurrentLine = UBound(arTemp)
End Function

Function GetLine(LineNumber As Integer) As String
Dim temp() As String
temp = Split(Script, vbCrLf)
GetLine = temp(LineNumber)
End Function

Private Function DoFunction(scall As String) As Variant
    Dim sname As String, realsname() As String
    Dim argList() As String
    Dim i As Integer

    sname = Trim(Mid(scall, 1, InStr(1, scall, "(") - 1))
    argList = Split(Mid(scall, InStr(1, scall, "(") + 1, InStr(1, scall, ")") - (InStr(1, scall, "(") + 1)), ",")
    
    For i = LBound(argList) To UBound(argList)
        argList(i) = Equation(Trim(argList(i)))
    Next i
    
    If FindPart(sname, " ") = 1 Then
    
    realsname = Split(sname, " ")
        
    If realsname(0) = "on" Or realsname(0) = "return:" Then
    DoFunction = ExecuteFunction(realsname(0), argList(), Mid(scall, InStr(scall, realsname(0)) + Len(realsname(0)) + 1))
    Else
    DoFunction = ExecuteFunction(realsname(0), argList())
    End If
    Else
    DoFunction = ExecuteFunction(sname, argList())
    End If
    
    sname = ""
    
End Function
Function IsAVar(sn As String) As Boolean
Dim i
For i = 0 To VarNames.Count - 1
If sn = VarNames.Item(i) Then
IsAVar = True
Exit Function
End If
Next i
End Function
Function ExecuteFunction(fName As String, argList() As String, Optional subcall As String)
Select Case LCase(fName)
Case "mout"
If ShowDebug = True Then
If UBound(argList()) < 2 Then
Errors = Errors + 1
ShowError 6, nCurrentLine + 1
Exit Function
End If
Else
End If
ExecuteFunction = MsgBox(argList(0), argList(1), argList(2))
Case "min"
If ShowDebug = True Then
If UBound(argList()) < 2 Then
Errors = Errors + 1
ShowError 6, nCurrentLine + 1
Exit Function
End If
Else
End If
ExecuteFunction = InputBox(argList(0), argList(1), argList(2))
Case "open"
If ShowDebug = True Then
If UBound(argList()) < 1 Then
Errors = Errors + 1
ShowError 6, nCurrentLine + 1
Exit Function
ElseIf FindPart(sLine, ")") = 0 Then

Errors = Errors + 1
ShowError 3, nCurrentLine + 1
Exit Function
End If
Else
End If
Set ScriptForm(CInt(argList(0))) = New frmWindow
ScriptForm(CInt(argList(0))).Caption = argList(1)
ScriptForm(CInt(argList(0))).Show
Case "close"
Unload ScriptForm(CInt(argList(0)))
Set ScriptForm(CInt(argList(0))) = Nothing
Case "hide"
ScriptForm(CInt(argList(0))).Hide
Case "show"
ScriptForm(CInt(argList(0))).Show
Case "print"
ScriptForm(CInt(argList(0))).Print argList(1)
Case "set"
If argList(0) = "backcolor" Then
ScriptForm(CInt(argList(1))).BackColor = Convert2Color(argList(2))
ElseIf argList(0) = "forecolor" Then
ScriptForm(CInt(argList(1))).ForeColor = Convert2Color(argList(2))
ElseIf argList(0) = "caption" Then
ScriptForm(CInt(argList(1))).Caption = argList(2)
ElseIf argList(0) = "cx" Then
ScriptForm(CInt(argList(1))).CurrentX = argList(2)
ElseIf argList(0) = "cy" Then
ScriptForm(CInt(argList(1))).CurrentY = argList(2)
ElseIf argList(0) = "caption" Then
ScriptForm(CInt(argList(1))).Caption = argList(2)
ElseIf argList(0) = "windowstate" Then
ScriptForm(CInt(argList(1))).WindowState = argList(2)
ElseIf argList(0) = "font" Then
ScriptForm(CInt(argList(1))).Font = argList(2)
ElseIf argList(0) = "fontsize" Then
ScriptForm(CInt(argList(1))).FontSize = CInt(argList(2))
End If
Case "sleep"
Pause CLng(argList(0))
Case "screen"
If argList(0) = "pixel" Then
If UBound(argList()) < 2 Then
ScriptForm(CInt(argList(1))).PSet (ScriptForm(CInt(argList(1))).CurrentX, ScriptForm(CInt(argList(1))).CurrentY), ScriptForm(CInt(argList(1))).ForeColor
ElseIf UBound(argList()) > 2 Then
ScriptForm(CInt(argList(1))).PSet (CInt(argList(2)), CInt(argList(3))), ScriptForm(CInt(argList(1))).ForeColor
End If
ElseIf argList(0) = "clear" Then
ScriptForm(CInt(argList(1))).Cls
ElseIf argList(0) = "refresh" Then
ScriptForm(CInt(argList(1))).Refresh
ElseIf argList(0) = "picture" Then
If FileExists(argList(2)) = True Then
ScriptForm(CInt(argList(1))).Picture = LoadPicture(argList(2))
Else
ShowError 7, nCurrentLine - 1
theend = True
Exit Function
End If
End If
Case "option"
If argList(1) = "mode" Then
ScriptForm(CInt(argList(0))).ScaleMode = argList(2)
ElseIf argList(1) = "focus" Then
ScriptForm(CInt(argList(0))).SetFocus
End If
Case "random"
ExecuteFunction = Int(Rnd * CInt(argList(0)))
Case "rgb"
ExecuteFunction = RGB(CInt(argList(0)), CInt(argList(1)), CInt(argList(2)))
Case "mid"
ExecuteFunction = Mid(argList(0), argList(1), argList(2))
Case "left"
ExecuteFunction = Left(argList(0), argList(1))
Case "right"
ExecuteFunction = Right(argList(0), argList(1))
Case "instr"
ExecuteFunction = InStr(argList(0), argList(1))
Case "split"
ExecuteFunction = Split(argList(0), argList(1))
Case "len"
ExecuteFunction = CInt(Len(argList(0)))
Case "end"
If argList(0) = "app" Then
Dim i
For i = 1 To UBound(ScriptForm())
Unload ScriptForm(i)
Next i
ElseIf argList(0) = "loop" Then
Skip = True
End If
Case "doevents"
DoEvents
Case "noend"
theend = True
Case "theend"
theend = False
Case "on"
If GetStatement(subcall) = True Then
inputvalue = True
Else
inputvalue = False
End If
Case "return:"
If inputvalue = True Then
DoFunction subcall
ElseIf inputvalue = False Then
End If
Case "keydown"
DeclareKeydowns (argList(0))
Case "with"

If FileExists(argList(0)) = True Then
Dim addcode As String, total As String
Open argList(0) For Input As #1
Line Input #1, addcode
Close #1
total = addcode
Script = Replace(Script, GetLine(nCurrentLine), vbCrLf & vbCrLf & total)
MsgBox Script
Me.Script = Script
Else
ShowError 7, nCurrentLine + 1
Exit Function
End If
Case "network"
If argList(0) = "open" Then
Load ScriptForm(argList(1)).Winsock(argList(2))
ElseIf argList(0) = "bind" Then
ScriptForm(argList(1)).Winsock(argList(2)).Bind argList(3), argList(4)
ElseIf argList(0) = "listen" Then
ScriptForm(argList(1)).Winsock(argList(2)).Listen
ElseIf argList(0) = "connect" Then
ScriptForm(argList(1)).Winsock(argList(2)).Connect argList(2), argList(3)
ElseIf argList(0) = "close" Then
ScriptForm(argList(1)).Winsock(argList(2)).Close
End If

Case "netstate"
ExecuteFunction = ScriptForm(argList(0)).Winsock(argList(1)).State

Case "make"
Select Case argList(0)
Case "button"
Load ScriptForm(argList(1)).Command(argList(2))
ScriptForm(argList(1)).Command(argList(2)).Caption = argList(3)
ScriptForm(argList(1)).Command(argList(2)).Visible = True
Case "textbox"
Load ScriptForm(argList(1)).Text(argList(2))
ScriptForm(argList(1)).Text(argList(2)).Text = argList(3)
ScriptForm(argList(1)).Text(argList(2)).Visible = True
End Select

Case "place"
Select Case Left(argList(0), 6)
Case "button"
ScriptForm(argList(1)).Command(argList(Right(argList(0), 1))).Left = argList(2)
ScriptForm(argList(1)).Command(argList(Right(argList(0), 1))).Top = argList(3)

Case "textbo"
ScriptForm(argList(1)).Text(argList(Right(argList(0), 1))).Left = argList(2)
ScriptForm(argList(1)).Text(argList(Right(argList(0), 1))).Top = argList(3)
End Select

Case "line"
ScriptForm(argList(0)).Line (argList(1), argList(2))-(argList(3), argList(4)), ScriptForm(argList(0)).ForeColor

Case "width"
ExecuteFunction = ScriptForm(argList(0)).ScaleWidth
Case "height"
ExecuteFunction = ScriptForm(argList(0)).ScaleHeight

Case "textwidth"
ExecuteFunction = frmMain.TextWidth(argList(0))
Case "textheight"
ExecuteFunction = frmMain.TextHeight(argList(0))

Case Else
If ShowDebug = True Then
ShowError 5, nCurrentLine + 1
Errors = Errors + 1
Else
End If
End Select
End Function

Function GetStatement(statement As String) As Boolean
GetStatement = False
Dim callName As String, scall As String, statement2 As String
callName = Mid(statement, 1, InStr(statement, "(") - 1)
scall = Mid(statement, InStr(statement, "(") + 1, InStr(statement, ")") - InStr(statement, "(") - 1)
scall = Replace(scall, Chr(34), "")

Select Case LCase(SpaceTrim(callName))
Case "key"

If scall = "return" Then
GetStatement = GetAsyncKeyState(VK_RETURN)
ElseIf scall = "space" Then
GetStatement = GetAsyncKeyState(VK_SPACE)
ElseIf scall = "escape" Then
GetStatement = GetAsyncKeyState(VK_ESCAPE)
ElseIf scall = "down" Then
GetStatement = GetAsyncKeyState(VK_DOWN)
ElseIf scall = "up" Then
GetStatement = GetAsyncKeyState(VK_UP)
ElseIf scall = "left" Then
GetStatement = GetAsyncKeyState(VK_LEFT)
ElseIf scall = "right" Then
GetStatement = GetAsyncKeyState(VK_RIGHT)
ElseIf scall = "f1" Then
GetStatement = GetAsyncKeyState(VK_F1)
ElseIf scall = "f2" Then
GetStatement = GetAsyncKeyState(VK_F2)
ElseIf scall = "f3" Then
GetStatement = GetAsyncKeyState(VK_F3)
ElseIf scall = "f4" Then
GetStatement = GetAsyncKeyState(VK_F4)
ElseIf scall = "f5" Then
GetStatement = GetAsyncKeyState(VK_F5)
ElseIf scall = "f6" Then
GetStatement = GetAsyncKeyState(VK_F6)
ElseIf scall = "f7" Then
GetStatement = GetAsyncKeyState(VK_F7)
ElseIf scall = "f8" Then
GetStatement = GetAsyncKeyState(VK_F8)
ElseIf scall = "f9" Then
GetStatement = GetAsyncKeyState(VK_F9)
ElseIf scall = "f10" Then
GetStatement = GetAsyncKeyState(VK_F10)
ElseIf scall = "f11" Then
GetStatement = GetAsyncKeyState(VK_F11)
ElseIf scall = "f12" Then
GetStatement = GetAsyncKeyState(VK_F12)
ElseIf scall = "shift" Then
GetStatement = GetAsyncKeyState(VK_SHIFT)
End If

Case "net:getdata"
GetStatement = WSKN(scall).getdata
Case "net:request"
GetStatement = WSKN(scall).request
End Select

End Function

Function DeclareKeydowns(sform As Integer)
If GetAsyncKeyState(VK_RETURN) Then ScriptForm(sform).Print ("KeyDown: Return")
If GetAsyncKeyState(VK_SPACE) Then ScriptForm(sform).Print ("KeyDown: Space")
If GetAsyncKeyState(VK_SHIFT) Then ScriptForm(sform).Print ("KeyDown: Shift")
If GetAsyncKeyState(VK_LEFT) Then ScriptForm(sform).Print ("KeyDown: Left")
If GetAsyncKeyState(VK_RIGHT) Then ScriptForm(sform).Print ("KeyDown: Right")
If GetAsyncKeyState(VK_UP) Then ScriptForm(sform).Print ("KeyDown: Up")
If GetAsyncKeyState(VK_DOWN) Then ScriptForm(sform).Print ("KeyDown: Down")
If GetAsyncKeyState(VK_F1) Then ScriptForm(sform).Print ("KeyDown: F1")
If GetAsyncKeyState(VK_F2) Then ScriptForm(sform).Print ("KeyDown: F2")
If GetAsyncKeyState(VK_F3) Then ScriptForm(sform).Print ("KeyDown: F3")
If GetAsyncKeyState(VK_F4) Then ScriptForm(sform).Print ("KeyDown: F4")
If GetAsyncKeyState(VK_F5) Then ScriptForm(sform).Print ("KeyDown: F5")
If GetAsyncKeyState(VK_F6) Then ScriptForm(sform).Print ("KeyDown: F6")
If GetAsyncKeyState(VK_F7) Then ScriptForm(sform).Print ("KeyDown: F7")
If GetAsyncKeyState(VK_F8) Then ScriptForm(sform).Print ("KeyDown: F8")
If GetAsyncKeyState(VK_F9) Then ScriptForm(sform).Print ("KeyDown: F9")
If GetAsyncKeyState(VK_F10) Then ScriptForm(sform).Print ("KeyDown: F10")
If GetAsyncKeyState(VK_F11) Then ScriptForm(sform).Print ("KeyDown: F11")
If GetAsyncKeyState(VK_F12) Then ScriptForm(sform).Print ("KeyDown: F12")
Exit Function
End Function

Private Function Equation(ByVal sEquation As String) As Variant
    ' For solving full equations like: iTemp + ("3" * ("5" / "4"))
    Dim iTemp As Integer, iTemp2 As Integer, i As Integer
    Dim sTemp As String
    
    ' First lets change anything inside ()'s to a constant value
    Do
        For i = 1 To Len(sEquation)
            sTemp = Mid(sEquation, i, 1)
            If sTemp = Chr(34) Then
                i = InStr(i + 1, sEquation, Chr(34))
            ElseIf sTemp = "(" Then
                iTemp = i
            ElseIf sTemp = ")" Then
                iTemp2 = i
                Exit For
            End If
        Next i
        If iTemp <> 0 Then
            sEquation = Mid(sEquation, 1, iTemp - 1) & " " & Chr(34) & Solve(Mid(sEquation, iTemp + 1, iTemp2 - (iTemp + 1))) & Chr(34) & " " & Mid(sEquation, iTemp2 + 1)
            iTemp = 0: iTemp2 = 0
        Else
            Exit Do
        End If
    Loop
    
    ' Now that all the ()'s are constant values, lets finish it up
    Equation = Solve(sEquation)
End Function

Private Function Solve(sEquation As String) As Variant
    ' For solving equations without ()'s like: itemp + "3" / "2"
    Dim i As Integer, iTemp As Integer
    Dim sTemp As String, sTemp2 As String
    Dim WaitVal As Variant
    Dim WaitOp As String
    Dim WaitVar
    
    For i = 1 To Len(sEquation)
        sTemp = Mid(sEquation, i, 1)
        If sTemp = Chr(34) Then
            iTemp = InStr(i + 1, sEquation, Chr(34))
            WaitVal = Mid(sEquation, i + 1, iTemp - (i + 1))
            i = iTemp
            If WaitOp <> "" Then
                Select Case WaitOp
                    Case "+"
                        Solve = Val(Solve) + Val(WaitVal)
                    Case "-"
                        Solve = Val(Solve) - Val(WaitVal)
                    Case "/"
                        Solve = Val(Solve) / Val(WaitVal)
                    Case "\"
                        Solve = Val(Solve) \ Val(WaitVal)
                    Case "^"
                        Solve = Val(Solve) ^ Val(WaitVal)
                    Case "*"
                        Solve = Val(Solve) * Val(WaitVal)
                    Case "&" ' String Concentration
                        Solve = Solve & WaitVal
                End Select
                WaitOp = ""
            Else
                Solve = WaitVal
            End If
        ElseIf sTemp = " " Then
            ' Make sure it skips this
        ElseIf sTemp = "+" Or sTemp = "-" Or sTemp = "/" Or sTemp = "\" Or sTemp = "^" Or sTemp = "&" Or sTemp = "*" Then
            If WaitVar <> 0 Then
                WaitVal = GetVar(Mid(sEquation, WaitVar, i - (WaitVar + 1)))
                If WaitOp <> "" Then
                    Select Case WaitOp
                        Case "+"
                        Solve = Val(Solve) + Val(WaitVal)
                        Case "-"
                            Solve = Val(Solve) - Val(WaitVal)
                        Case "/"
                            Solve = Val(Solve) / Val(WaitVal)
                        Case "\"
                            Solve = Val(Solve) \ Val(WaitVal)
                        Case "^"
                            Solve = Val(Solve) ^ Val(WaitVal)
                        Case "*"
                            Solve = Val(Solve) * Val(WaitVal)
                        Case "&" ' String Concentration
                            Solve = Solve & WaitVal
                    End Select
                    WaitOp = ""
                Else
                    Solve = WaitVal
                End If
                
                WaitVar = 0
            End If
            
            WaitOp = sTemp
        Else
            If WaitVar = 0 Then WaitVar = i
            If i >= Len(sEquation) Then
                WaitVal = GetVar(Mid(sEquation, WaitVar, i))
                If WaitOp <> "" Then
                    Select Case WaitOp
                        Case "+"
                            Solve = Val(Solve) + Val(WaitVal)
                        Case "-"
                            Solve = Val(Solve) - Val(WaitVal)
                        Case "/"
                            Solve = Val(Solve) / Val(WaitVal)
                        Case "\"
                            Solve = Val(Solve) \ Val(WaitVal)
                        Case "^"
                            Solve = Val(Solve) ^ Val(WaitVal)
                        Case "*"
                            Solve = Val(Solve) * Val(WaitVal)
                        Case "&" ' String Concentration
                            Solve = Solve & WaitVal
                    End Select
                    WaitOp = ""
                Else
                    Solve = WaitVal
                End If
            End If
        End If
    Next i
End Function
Private Sub ChangeVar(TheVar As String, NewVal As Variant)
    Dim i As Integer
    For i = VarNames.Count To 1 Step -1
        If VarNames(i) = TheVar Then
            Variables.Remove i
            VarNames.Remove i
            
            Variables.Add NewVal
            VarNames.Add TheVar
            Exit Sub
        End If
    Next i
End Sub

Private Function GetVar(TheVar As String) As Variant
    Dim i As Integer
    For i = 1 To VarNames.Count
        If VarNames(i) = TheVar Then
            GetVar = Variables(i)
            Exit Function
        End If
    Next i
End Function

Private Function Eval(ByVal sFormula As String) As Boolean
    Dim i As Integer, iWait As Integer
    Dim LeftVal As String, RightVal As String, Operator As String
    Dim sTemp As String
    
    ' Get the left value
    For i = 1 To Len(sFormula)
        sTemp = Mid(sFormula, i, 1)
        Select Case sTemp
            Case "("
                iWait = iWait + 1
            Case ")"
                iWait = iWait - 1
                If iWait = 0 Then
                    LeftVal = Mid(sFormula, 1, i)
                    sFormula = Trim(Mid(sFormula, i + 1))
                    Exit For
                End If
            Case Chr(34)
                i = InStr(i + 1, sFormula, Chr(34))
                If iWait = 0 Then
                    LeftVal = Mid(sFormula, 1, i)
                    sFormula = Trim(Mid(sFormula, i + 1))
                    Exit For
                End If
            Case Else
                If sTemp = ">" Or sTemp = "<" Or sTemp = "=" Then
                    If iWait = 0 Then
                        LeftVal = Trim(Mid(sFormula, 1, i - 1))
                        sFormula = Trim(Mid(sFormula, i))
                        Exit For
                    End If
                End If
        End Select
    Next i
    
    ' Get the comparison operator
    sTemp = Mid(sFormula, 2, 1)
    If sTemp = ">" Or sTemp = "<" Or sTemp = "=" Then
        Operator = Left(sFormula, 2)
        sFormula = Mid(sFormula, 3)
    Else
        Operator = Left(sFormula, 1)
        sFormula = Mid(sFormula, 2)
    End If
    
    ' Set the rest to the right side operator
    RightVal = sFormula
    sFormula = ""
    
    ' Solve these sides
    LeftVal = Equation(LeftVal)
    RightVal = Equation(RightVal)
    
    Select Case Operator
        Case ">"
            If Val(LeftVal) > Val(RightVal) Then Eval = True
        Case "<"
            If Val(LeftVal) < Val(RightVal) Then Eval = True
        Case "<>"
            If Val(LeftVal) <> Val(RightVal) Then Eval = True
        Case ">="
            If Val(LeftVal) >= Val(RightVal) Then Eval = True
        Case "<="
            If Val(LeftVal) <= Val(RightVal) Then Eval = True
        Case "="
            If LeftVal = RightVal Then Eval = True
        Case "=="
            If LCase(LeftVal) = LCase(RightVal) Then Eval = True
    End Select
End Function
