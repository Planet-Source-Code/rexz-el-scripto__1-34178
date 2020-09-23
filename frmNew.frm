VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   0
      TabIndex        =   1
      Top             =   1870
      Width           =   2175
   End
   Begin VB.ListBox lstOptions 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      ItemData        =   "frmNew.frx":0000
      Left            =   0
      List            =   "frmNew.frx":000A
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Select Case LCase(lstOptions.Text)
Case "new (empty)"
frmMain.txtCode.Text = ""
Unload Me
Case "new (window code)"
frmMain.txtCode.Text = "open (" & Chr(34) & "1" & Chr(34) & "," & Chr(34) & "Window" & Chr(34) & ")"
Unload Me
Case "custom control"
frmMain.txtCode.Text = ""
Unload Me
Case "runtime library"
frmMain.txtCode.Text = ""
Unload Me
Case "ell - extern link library"
Case ""
MsgBox "Please choose an option", vbExclamation, "El Scripto - New"
End Select
End Sub
