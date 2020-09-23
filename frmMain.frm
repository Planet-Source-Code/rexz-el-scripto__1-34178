VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "El Scripto - New.esc"
   ClientHeight    =   8895
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   11190
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   7095
      Left            =   2160
      TabIndex        =   7
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12515
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7065
      ScaleWidth      =   2025
      TabIndex        =   3
      Top             =   0
      Width           =   2055
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":03CF
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   0
         TabIndex        =   6
         Top             =   840
         Width           =   1965
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "El Scripto"
         BeginProperty Font 
            Name            =   "Pine Casual"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Left            =   140
         TabIndex        =   5
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "El Scripto"
         BeginProperty Font 
            Name            =   "Pine Casual"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   480
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.TextBox txtDebug 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7440
      Width           =   11175
   End
   Begin MSComDlg.CommonDialog Cm1 
      Left            =   6840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   8550
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2804
            MinWidth        =   2804
            Picture         =   "frmMain.frx":0459
            Text            =   "Press F1 for help"
            TextSave        =   "Press F1 for help"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Ln = 0"
            TextSave        =   "Ln = 0"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Col = 0"
            TextSave        =   "Col = 0"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2813
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C55
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3097
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":391B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":419F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A23
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E65
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52A7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debug:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   7200
      Width           =   630
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Project"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Project"
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "Save Project as"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Find in text"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuSupOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuCompileEXE 
         Caption         =   "Compile"
         Begin VB.Menu mnuCompile 
            Caption         =   "Make EXE"
         End
         Begin VB.Menu mnuRun 
            Caption         =   "Run"
            Shortcut        =   {F5}
         End
      End
      Begin VB.Menu mnuComps 
         Caption         =   "Compabilities"
      End
      Begin VB.Menu mnuAddins 
         Caption         =   "Make/Edit Add-ins"
      End
      Begin VB.Menu mnuPlugin 
         Caption         =   "Add Plugins"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuViewHelp 
         Caption         =   "View Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuTool 
         Caption         =   "[Empty]"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuTool 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTool 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTool 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTool 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTool 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sChange
Private Declare Function fCreateShellLink Lib "VB6STKIT.DLL" (ByVal _
        lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
        lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
        
        Private Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, _
                        ByVal uFlags As Long, ByVal dwItem1 As Long, _
                        ByVal dwItem2 As Long)
        
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0

Private Sub Form_Load()

KWColor = RGB(0, 0, 195)
IDColor = RGB(0, 0, 255)
EFColor = RGB(0, 125, 255)
FIColor = RGB(0, 100, 0)
OTColor = RGB(185, 0, 0)
ShowDebug = True

cm1.InitDir = App.path
Dim strString As String
    Dim lngDword As Long
    Dim Record As String
    
    If Command$ <> "%1" And Command$ <> "" Then
        'Command$ is the file you need To open!
        'Load the file
        Open Command$ For Input As #1
        Do While Not EOF(1)
            Line Input #1, Record
            txtfile = txtfile & Record & vbCrLf
            setcolors
        Loop
        If Right(Command$, 2) = "SC" Then
        txtCode.Text = TrimCode(txtfile)
        ElseIf Right(Command$, 2) = ".A" Then
        frmEditor.txtCode.Text = TrimCode(txtfile)
        frmEditor.Show
        frmEditor.SetFocus
        Me.WindowState = vbMinimized
        End If
        'Add your file to the Recent file folder:
        lReturn = fCreateShellLink("..\..\Recent", _
                Command$, Command$, "")
        
    End If
        If GetString(HKEY_CLASSES_ROOT, ".a", "Content Type") = "" Then
        'Nope - not added yet. Register the file type:
        
        'create an entry in the class key
        Call SaveString(HKEY_CLASSES_ROOT, ".a", "", "ELS Add-In")
        'content type
        Call SaveString(HKEY_CLASSES_ROOT, ".a", "Content Type", "text/plain")
        'name
        Call SaveString(HKEY_CLASSES_ROOT, "ELS Add-In", "", "El Scripto Add-In")
        'edit flags
        Call SaveDWord(HKEY_CLASSES_ROOT, "ELS Add-In", "EditFlags", "0000")
        'file's icon (can be an icon file, or an icon located within a dll file)
        'in this example, I am using a resource icon in this exe, 0 (app icon).
        Call SaveString(HKEY_CLASSES_ROOT, "ELS Add-In\DefaultIcon", "", App.path & "\" & App.EXEName & ".exe,0")
        'Shell
        Call SaveString(HKEY_CLASSES_ROOT, "ELS Add-In\Shell", "", "")
        'Shell Open
        Call SaveString(HKEY_CLASSES_ROOT, "ELS Add-In\Shell\Open", "", "")
        'Shell open command
        Call SaveString(HKEY_CLASSES_ROOT, "ELS Add-In\Shell\Open\Command", "", App.path & "\" & App.EXEName & ".exe %1")
        'Update the Windows Icon Cache to see our icon right away:
        SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
        End If
        
        If GetString(HKEY_CLASSES_ROOT, ".esc", "Content Type") = "" Then
        'Nope - not added yet. Register the file type:
        
        'create an entry in the class key
        Call SaveString(HKEY_CLASSES_ROOT, ".esc", "", "ELS Project")
        'content type
        Call SaveString(HKEY_CLASSES_ROOT, ".esc", "Content Type", "text/plain")
        'name
        Call SaveString(HKEY_CLASSES_ROOT, "ELS Project", "", "El Scripto Add-In")
        'edit flags
        Call SaveDWord(HKEY_CLASSES_ROOT, "ELS Project", "EditFlags", "0000")
        'file's icon (can be an icon file, or an icon located within a dll file)
        'in this example, I am using a resource icon in this exe, 0 (app icon).
        Call SaveString(HKEY_CLASSES_ROOT, "ELS Project\DefaultIcon", "", App.path & "\" & App.EXEName & ".exe,0")
        'Shell
        Call SaveString(HKEY_CLASSES_ROOT, "ELS Project\Shell", "", "")
        'Shell Open
        Call SaveString(HKEY_CLASSES_ROOT, "ELS Project\Shell\Open", "", "")
        'Shell open command
        Call SaveString(HKEY_CLASSES_ROOT, "ELS Project\Shell\Open\Command", "", App.path & "\" & App.EXEName & ".exe %1")
        'Update the Windows Icon Cache to see our icon right away:
        SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
        End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Change = 0 Then
End
Else
Dim Msg
Msg = MsgBox("Do you want to save the changes before you exit?", vbYesNoCancel + vbQuestion, "Save changes?")
If Msg = vbYes Then
mnuSave_Click
ElseIf Msg = vbNo Then
End
Else
Exit Sub
End If
End If
End Sub

Private Sub mnuAbout_Click()
Dim Msg
Msg = "El Scripto by Hans Bjerndell 2001" & vbCrLf & vbCrLf & "Copyright©2001 All rights reserved."
MsgBox Msg, vbInformation + vbOKOnly, "About El Scripto"
End Sub

Private Sub mnuAddins_Click()
frmEditor.Show
End Sub

Private Sub mnuCompile_Click()
On Error GoTo ErrorCatch:
    
    If Not FileExists(App.path & "\ide.exe") Then
        MsgBox "Could not find El Scripto IDE, aborting.", vbCritical, "Error"
        Exit Sub
    End If
    
    cm1.filename = ""
    cm1.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
    cm1.ShowSave
    If cm1.filename <> "" Then
        If FileExists(cm1.filename) Then
            If MsgBox("Overwrite existing file?", vbQuestion + vbYesNo, "JEl") = vbNo Then
                Exit Sub
            Else
                Kill cm1.filename
            End If
        End If
        
        Dim nFile As Integer
        nFile = FreeFile
        FileCopy App.path & "\ide.exe", cm1.filename
        Open cm1.filename For Output As #nFile
        Print #nFile, "|*ESC*|" & txtCode.Text
        Close #nFile
        MsgBox "File Compiled!", vbInformation, "El Scripto"
        
        Dim sTemp As String, sTemp2 As String
        
        Open cm1.filename For Output As #1
        Open App.path & "\ide.exe" For Binary As #2
        
        ' Copy data from jelexe into new exe
        While Not EOF(2)
            sTemp = Input$(2000, #2)
            sTemp2 = sTemp2 & sTemp
            Print #1, sTemp2;
            sTemp2 = ""
            If Len(sTemp) > 2000 Then
                sTemp = ""
            End If
        Wend
        
        ' Append the script
        Print #1, "|*ESC*|" & txtCode.Text
        
        Close #2
        Close #1
        
        
    End If
    
    Exit Sub
ErrorCatch:
    MsgBox "Error has occured: " & Err.Description, vbCritical, "Error"
    Resume Next
End Sub

Private Sub mnuComps_Click()
frmCompabilities.Show
End Sub

Private Sub mnuCopy_Click()
EditMenu txtCode, "COPY"
End Sub

Private Sub mnuCut_Click()
EditMenu txtCode, "CUT"
End Sub

Private Sub mnuExit_Click()
If Change = 0 Then
End
Else
Dim Msg
Msg = MsgBox("Do you want to save the changes before you exit?", vbYesNoCancel + vbQuestion, "Save changes?")
If Msg = vbYes Then
mnuSave_Click
ElseIf Msg = vbNo Then
End
Else
Exit Sub
End If
End If
End Sub

Private Sub mnuNew_Click()
frmNew.Show
End Sub

Private Sub mnuOpen_Click()
On Error GoTo exitit:
Dim Code As String
cm1.Filter = "El Scripto Project |*.esc;*.a|"
cm1.ShowOpen
Open cm1.filename For Input As #1
If cm1.filename <> "" Then
txtCode.Text = ""
Do Until EOF(1)
DoEvents
Line Input #1, Code
txtCode.Text = txtCode.Text & Code & vbCrLf
Loop
If ColorKeyWords = True Then
setcolors
Else
End If
Else
Exit Sub
End If
Close #1

Me.Caption = "El Scripto - " & cm1.FileTitle
exitit:
End Sub

Private Sub mnuPaste_Click()
EditMenu txtCode, "PASTE"
End Sub

Private Sub mnuPlugin_Click()
frmPlugInHandler.Show
End Sub

Private Sub mnuRun_Click()
Compiler.Script = txtCode.Text
Compiler.LoopApp
Compiler.Execute
End Sub

Private Sub mnuSave_Click()
cm1.Filter = "El Scripto Project |*.esc|"
cm1.ShowSave
On Error GoTo exitit
Open cm1.filename For Output As #1
Print #1, txtCode.Text
Close #1
Change = 0
Me.Caption = "El Scripto - " & cm1.FileTitle
exitit:
End Sub

Private Sub mnuSaveas_Click()
cm1.Filter = "El Scripto Project |*.esc|"
cm1.ShowSave
On Error GoTo exitit
Open cm1.filename For Output As #1
Print #1, txtCode.Text
Close #1
Change = 0
Me.Caption = "El Scripto - " & cm1.FileTitle
exitit:
End Sub

Private Sub mnuSearch_Click()
frmFind.Show
End Sub

Private Sub mnuSelAll_Click()
EditMenu txtCode, "SELALL"
End Sub

Private Sub mnuSupOptions_Click()
frmOptions.Show
End Sub

Private Sub mnuTool_Click(Index As Integer)
For i& = 0 To PluginCount
If mnuTool(Index).Caption & ".dll" = Plugin(i).sname Then
Plugin(i).plObject.run
Exit Sub
End If
Next i
End Sub

Private Sub mnuViewHelp_Click()
frmHelp.Show
End Sub

Private Sub txtCode_Change()
Dim Cur_Line As Long
Dim Cur_Col As Long

    On Local Error Resume Next
    Cur_Line = SendMessage(txtCode.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
    Cur_Col = txtCode.SelStart
    StatusBar1.Panels(4).Text = "Ln = " & Format(Cur_Line, "##,###")
    StatusBar1.Panels(5).Text = "Col = " & Cur_Col
    If txtCode <> "" Then
    Change = 1
    Else
    Change = 0
    End If
End Sub


Function Add(Text As String)
txtDebug.SelStart = Len(txtDebug)
txtDebug.Text = txtDebug.Text & Text & vbCrLf
txtDebug.SelLength = Len(txtDebug)
End Function

Private Sub txtCode_Click()
txtCode_Change
End Sub

Function TrimCode(lzCode) As String
TrimCode = lzCode
'TrimCode = Replace(lzCode, Chr(34), "")
End Function

Public Sub setcolors()
commentchar = "'"
If KeyCode = 13 Then Exit Sub
LockWindowUpdate Me.hwnd
clearwordcolors txtCode

ColorizeWord txtCode, commentchar, &H8000& 'This char is for comments like this
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
ColorizeWord txtCode, "open", KWColor
ColorizeWord txtCode, "set", KWColor
ColorizeWord txtCode, "dim", KWColor
ColorizeWord txtCode, "print", KWColor
ColorizeWord txtCode, "mout", KWColor
ColorizeWord txtCode, "min", KWColor
ColorizeWord txtCode, "show", KWColor
ColorizeWord txtCode, "hide", KWColor
ColorizeWord txtCode, "close", KWColor
ColorizeWord txtCode, "sleep", KWColor
ColorizeWord txtCode, "close", KWColor
ColorizeWord txtCode, "screen", KWColor
ColorizeWord txtCode, "on", KWColor
ColorizeWord txtCode, "end", KWColor
ColorizeWord txtCode, "if", KWColor
ColorizeWord txtCode, "else", KWColor
ColorizeWord txtCode, "option", KWColor
ColorizeWord txtCode, "random", KWColor
ColorizeWord txtCode, "with", KWColor
ColorizeWord txtCode, "code", KWColor
ColorizeWord txtCode, "while", KWColor
ColorizeWord txtCode, "goto", KWColor
ColorizeWord txtCode, "exit", KWColor
ColorizeWord txtCode, "endif", KWColor
ColorizeWord txtCode, "call", KWColor

ColorizeWord txtCode, "rgb", IDColor
ColorizeWord txtCode, "doevents", IDColor
ColorizeWord txtCode, "noend", IDColor
ColorizeWord txtCode, "return:", IDColor
ColorizeWord txtCode, "doevents()", IDColor
ColorizeWord txtCode, "noend()", IDColor

ColorizeWord txtCode, "mid", EFColor
ColorizeWord txtCode, "left", EFColor
ColorizeWord txtCode, "right", EFColor
ColorizeWord txtCode, "instr", EFColor
ColorizeWord txtCode, "split", EFColor
ColorizeWord txtCode, "len", EFColor
ColorizeWord txtCode, "key", EFColor

ColorizeWord txtCode, "end function", FIColor
ColorizeWord txtCode, "function", FIColor
ColorizeWord txtCode, "public", FIColor
ColorizeWord txtCode, "private", FIColor
ColorizeWord txtCode, "public function", FIColor
ColorizeWord txtCode, "private function", FIColor

ColorizeWord txtCode, "(", OTColor
ColorizeWord txtCode, "[", OTColor
ColorizeWord txtCode, "]", OTColor
ColorizeWord txtCode, ")", OTColor

ColorizeWord txtCode, "1", EFColor
ColorizeWord txtCode, "2", EFColor
ColorizeWord txtCode, "3", EFColor
ColorizeWord txtCode, "4", EFColor
ColorizeWord txtCode, "5", EFColor
ColorizeWord txtCode, "6", EFColor
ColorizeWord txtCode, "7", EFColor
ColorizeWord txtCode, "8", EFColor
ColorizeWord txtCode, "9", EFColor
ColorizeWord txtCode, "0", EFColor

LockWindowUpdate 0&
txtCode.Enabled = True
If txtCode.Visible = True Then
txtCode.SetFocus
End If
End Sub

Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeySpace Then
If ColorKeyWords = True Then
setcolors
Else
End If
'End If
End Sub

Private Sub txtCode_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Cur_Line As Long
Dim Cur_Col As Long

    On Local Error Resume Next
    Cur_Line = SendMessage(txtCode.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
    Cur_Col = txtCode.SelStart
    StatusBar1.Panels(4).Text = "Ln = " & Format(Cur_Line, "##,###")
    StatusBar1.Panels(5).Text = "Col = " & Cur_Col
End Sub

Private Sub txtCode_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If ColorKeyWords = True Then
setcolors
Else
End If
End Sub

Function SwapSideBar(value As Boolean)
If value = True Then
picLeft.Visible = True
txtCode.Width = 9015
txtCode.Left = 2160
ElseIf value = False Then
picLeft.Visible = False
txtCode.Left = 30
txtCode.Width = 11145
End If
End Function

Function SwapStatusBar(value As Boolean)
If value = True Then
StatusBar1.Visible = True
Me.Height = 9510
ElseIf value = False Then
StatusBar1.Visible = False
Me.Height = 9135
End If
End Function
