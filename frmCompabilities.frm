VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompabilities 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compabilities"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Info 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Info 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Info 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompabilities.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LView 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3201
      View            =   3
      Arrange         =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Types"
         Object.Width           =   4057
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon ID:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2520
      TabIndex        =   6
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Asociate Program:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Extension Type:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCompabilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
LView.ListItems.Add 1, "esc", "Project.esc", 0, 0
LView.ListItems.Add 2, "a", "Add-In.a", 0, 0
'LView.ListItems.Add 3, "ll", "Link-Library.ll", 0, 0 'Future versions:
'LView.ListItems.Add 4, "r", "Resource.r", 0, 0
'LView.ListItems.Add 5, "s", "Header.s", 0, 0
End Sub
Private Sub LView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim temp As String
On Error Resume Next
temp = LView.HitTest(x, y)
GetInfo Mid(temp, InStr(temp, ".") + 1)
End Sub

Function GetInfo(extension As String)
Select Case extension
Case "esc"
Info(0).Text = "ESC"
Info(1).Text = "El Scripto (Main)"
Info(2).Text = frmMain.Icon
Case "a"
Info(0).Text = "A"
Info(1).Text = "El Scripto Add-In Editor"
Info(2).Text = frmMain.Icon
Case "ll"
Info(0).Text = "LL"
Info(1).Text = "El Scripto (Main, with intercourse)"
Info(2).Text = frmMain.Icon
Case "r"
Info(0).Text = "R"
Info(1).Text = "El Scripto Resource Viewer"
Info(2).Text = frmMain.Icon
Case "s"
Info(0).Text = "S"
Info(1).Text = "El Scripto (Main)"
Info(2).Text = frmMain.Icon
End Select
End Function

Private Sub LView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim temp As String
On Error Resume Next
temp = LView.HitTest(x, y)
GetInfo Mid(temp, InStr(temp, ".") + 1)
End Sub
