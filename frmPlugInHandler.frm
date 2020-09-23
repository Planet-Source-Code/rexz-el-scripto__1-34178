VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlugInHandler 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Plugin Handler"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdActivate 
      Caption         =   "Activate"
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.FileListBox File 
      Height          =   255
      Left            =   4920
      Pattern         =   "*.dll"
      TabIndex        =   4
      Top             =   2760
      Width           =   255
      Visible         =   0   'False
   End
   Begin MSComDlg.CommonDialog cm1 
      Left            =   4800
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin MSComctlLib.ListView lstView 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7011
      View            =   3
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmPlugInHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActivate_Click()
'// As you see, this pluginhandler is only interacted with my plugins :)
On Error Resume Next
Dim responce
For i& = 0 To PluginCount
If Plugin(i).sname = lstView.ListItems.Item(lstView.SelectedItem.Index).Text Then
responce = Plugin(i).plObject.Activate
If responce = "create_tool" Then
Dim closest As Integer
closest = FindClosestMnu()
frmMain.mnuTools.Visible = True
frmMain.mnuTool(closest).Caption = Replace(Plugin(i).sname, ".dll", "")
frmMain.mnuTool(closest).Visible = True
frmMain.mnuTool(closest).Enabled = True
Else
Plugin(i).plObject.run
End If
End If
Next i
End Sub
Function FindClosestMnu() As Integer
For i& = 0 To frmMain.mnuTool.Count
If frmMain.mnuTool(i).Caption = "" Or frmMain.mnuTool(i).Caption = "[Empty]" Then
FindClosestMnu = i
Exit Function
End If
Next i
End Function
Private Sub cmdBrowse_Click()
On Error Resume Next
cm1.Filter = "Dynamic Link Libraries |*.dll|"
cm1.ShowOpen
If cm1.filename <> "" Then
lstView.ListItems.Add lstView.ListItems.Count + 1, cm1.filename, Replace(cm1.FileTitle, ".dll", "")
AddToList
MakeLists
Else: Exit Sub
End If
error:
End Sub

Private Sub cmdOK_Click()
Me.Hide
End Sub

Private Sub cmdView_Click()
Dim i As Integer
For i = 0 To PluginCount
If Plugin(i).sname = lstView.ListItems.Item(lstView.SelectedItem.Index).Text Then
frmViewPlugin.Show
frmViewPlugin.OpenPlugin Plugin(i).sname
frmViewPlugin.GetProperties
Exit Sub
End If
Next i
End Sub

Private Sub Form_Load()
File.path = App.path & "\Add-ins\"
AddToList
MakeLists
End Sub

Function AddToList()
For i& = 0 To File.ListCount - 1
lstView.ListItems.Add lstView.ListItems.Count + 1, App.path & "\Add-ins\" & File.List(i), File.List(i)
lstView.ListItems(lstView.ListItems.Count).SubItems(1) = "Dynamic Link Library"
Next i
End Function

Function MakeLists()
For i& = 1 To lstView.ListItems.Count

If lstView.ListItems(i).Checked = True Then
Else

Set Plugin(PluginCount).plObject = aGetObject(lstView.ListItems.Item(i).Key)
Plugin(PluginCount).sname = lstView.ListItems.Item(i).Text
SetProperties lstView.ListItems.Item(i).Key, PluginCount
Plugin(PluginCount).Index = PluginCount
lstView.ListItems.Item(i).Checked = True
PluginCount = PluginCount + 1

End If
Next i
End Function

Function SetProperties(path As String, plugindex As Integer)
Dim c As TypeLibInfo
Set c = TLI.TypeLibInfoFromFile(path)
Plugin(plugindex).content.sparent = c.CoClasses.Item(1).Parent
For i& = 1 To c.CoClasses.Count
Plugin(plugindex).content.classes.Add c.CoClasses.Item(i).Name
Next i
End Function
