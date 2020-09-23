VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cm1 
      Left            =   2160
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3836
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Interface"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Options(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Options(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmColoring"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Options(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tools"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Options(7)"
      Tab(1).Control(1)=   "Options(6)"
      Tab(1).Control(2)=   "Options(5)"
      Tab(1).Control(3)=   "Options(4)"
      Tab(1).Control(4)=   "Options(3)"
      Tab(1).Control(5)=   "Label2"
      Tab(1).Control(6)=   "Label1"
      Tab(1).ControlCount=   7
      Begin VB.CheckBox Options 
         Caption         =   "Show compilation stats"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   14
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox Options 
         Caption         =   "Debug extern controls"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   13
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox Options 
         Caption         =   "Show debug-syntax"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   12
         Top             =   1320
         Value           =   2  'Grayed
         Width           =   1695
      End
      Begin VB.CheckBox Options 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   10
         Top             =   840
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox Options 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox Options 
         Caption         =   "Show statusbar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame frmColoring 
         Caption         =   "Keyword Coloring"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2415
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1440
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   55
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cmoKeys 
            Height          =   315
            ItemData        =   "frmOptions.frx":0038
            Left            =   120
            List            =   "frmOptions.frx":004B
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CheckBox Options 
         Caption         =   "Show sidebar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Options 
         Caption         =   "Highlight keywords"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   400
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Debug:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   -74880
         TabIndex        =   11
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Menu:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmoKeys_Change()
cmoKeys_Scroll
End Sub

Private Sub cmoKeys_Click()
Select Case cmoKeys.List(cmoKeys.ListIndex)
Case "Keywords"
picColor.BackColor = KWColor
Case "Identifiers"
picColor.BackColor = IDColor
Case "Embedded Functions"
picColor.BackColor = EFColor
Case "Function Identifiers"
picColor.BackColor = FIColor
Case "Other"
picColor.BackColor = OTColor
End Select
End Sub

Private Sub cmoKeys_Scroll()
Select Case cmoKeys.List(cmoKeys.ListIndex)
Case "Keywords"
picColor.BackColor = KWColor
Case "Identifiers"
picColor.BackColor = IDColor
Case "Embedded Functions"
picColor.BackColor = EFColor
Case "Function Identifiers"
picColor.BackColor = FIColor
Case "Other"
picColor.BackColor = OTColor
End Select
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
SetSettings
End Sub

Private Sub Options_Click(Index As Integer)
Select Case Index
Case 0
ColorKeyWords = Options(0).value
frmColoring.Enabled = Options(0).value
Case 1
frmMain.SwapSideBar Options(1).value
Case 2
frmMain.SwapStatusBar Options(2).value
Case 3
frmMain.mnuHelp.Visible = Options(3).value
Case 4
frmMain.mnuEdit.Visible = Options(4).value
Case 5
ShowDebug = Options(5).value
Case 6
DebugExtern = Options(6).value
Case 7
ShowCompStats = Options(7).value
End Select
End Sub

Private Sub picColor_Click()
On Error GoTo exitit:
cm1.ShowColor

Select Case cmoKeys.Text
Case "Keywords"
KWColor = cm1.color
Case "Identifiers"
IDColor = cm1.color
Case "Embedded Functions"
EFColor = cm1.color
Case "Function Identifiers"
FIColor = cm1.color
Case "Other"
OTColor = cm1.color
End Select
picColor.BackColor = cm1.color
exitit:
End Sub

Function SetSettings()
If frmMain.picLeft.Visible = True Then
Options(1).value = 1
Else
Options(1).value = 0
End If
If frmMain.StatusBar1.Visible = True Then
Options(2).value = 1
Else
Options(2).value = 0
End If
If ColorKeyWords = True Then
Options(0).value = 1
frmColoring.Enabled = True
Else
Options(0).value = 0
frmColoring.Enabled = False
End If
If frmMain.mnuHelp.Visible = True Then
Options(3).value = 1
Else
Options(3).value = 0
End If
If frmMain.mnuEdit.Visible = True Then
Options(4).value = 1
Else
Options(4).value = 0
End If
If ShowDebug = True Then
Options(5).value = 1
Else
Options(5).value = 0
End If
End Function
