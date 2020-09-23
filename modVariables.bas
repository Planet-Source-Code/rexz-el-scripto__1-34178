Attribute VB_Name = "modVariables"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const VK_RETURN = &HD
Public Const VK_CONTROL = &H11
Public Const VK_DOWN = &H28
Public Const VK_ESCAPE = &H1B
Public Const VK_F1 = &H70
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const VK_SPACE = &H20
Public Const VK_SHIFT = &H10
Public Const VK_UP = &H26

Public totalusers As Integer

Public Type WSK
    getdata As Boolean
    request As Boolean
End Type

Public WSKN(10) As WSK

Global ShowDebug As Boolean
Global DebugExtern As Boolean
Global ShowCompStats As Boolean

Global KWColor As Long
Global IDColor As Long
Global EFColor As Long
Global FIColor As Long
Global OTColor As Long

Global ColorKeyWords As Boolean
Global theend As Boolean

Public Type DLLContent
    sparent As String
    classes As New Collection
End Type

Public Type tpPlugin
    plObject As Object
    Index As Integer
    sname As String
    content As DLLContent
End Type

Public PluginCount As Integer
Public Plugin(10000) As tpPlugin

Public Const EM_LINEFROMCHAR = &HC9

Function EditMenu(txtBox As RichTextBox, Cmd As String)
Dim StrFind As String
Dim Xpos As Integer

    Select Case Cmd
        Case "CUT"
            Clipboard.SetText txtBox.SelText
            txtBox.SelText = ""
        Case "COPY"
            Clipboard.SetText txtBox.SelText
        Case "PASTE"
            txtBox.SelText = Clipboard.GetText
        Case "SELALL"
            txtBox.SelStart = 0
            txtBox.SelLength = Len(txtBox.text)
            
        Case "FIND"
            StrFind = frmFind.txtFind.text
            If Len(StrFind) = 0 Then
                Exit Function
            Else
                Xpos = InStr(txtBox.text, StrFind)
                If Xpos > 0 Then
                    txtBox.SetFocus
                    txtBox.SelStart = Xpos - 1
                    txtBox.SelLength = Len(StrFind)
                Else
                    Beep
                    MsgBox "Search text " & Chr(34) & StrFind & Chr(34) & " was not found", vbExclamation
                End If
            End If
            Xpos = 0
            Cmd = ""
            StrFind = ""
    End Select
    
End Function
