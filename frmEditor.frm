VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "El Script - Make new Add-in"
   ClientHeight    =   3150
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
   Begin MSComDlg.CommonDialog Cm1 
      Left            =   4680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Quit Editor"
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuExit_Click()
frmMain.WindowState = vbNormal
Unload Me
End Sub

Private Sub mnuNew_Click()
txtCode.text = ""
End Sub

Private Sub mnuOpen_Click()
Dim Code As String
cm1.Filter = "El Scripto Add-In |*.a|"
cm1.ShowOpen
On Error Resume Next
Open cm1.filename For Input As #1
Line Input #1, Code
Close #1
txtCode.text = Code
End Sub

Private Sub mnuSave_Click()
cm1.Filter = "El Scripto Add-In |*.a|"
cm1.ShowSave
On Error Resume Next
Open cm1.filename For Output As #1
Print #1, txtCode.text
Close #1
End Sub
