VERSION 5.00
Begin VB.Form frmViewPlugin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "View Plugin"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3270
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
   ScaleHeight     =   2265
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstClasses 
      Appearance      =   0  'Flat
      Height          =   2010
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblClass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class Count: "
      Height          =   165
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   840
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name: "
      Height          =   165
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enum Classes:"
      Height          =   165
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   930
   End
End
Attribute VB_Name = "frmViewPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim plindex As Integer
Function OpenPlugin(sname As String)
Dim i
For i = 0 To PluginCount
If Plugin(i).sname = sname Then
plindex = i
Exit Function
End If
Next i
End Function

Function GetProperties()
Dim i
For i = 1 To Plugin(plindex).content.classes.Count
lstClasses.AddItem Plugin(plindex).content.classes.Item(i)
lblName.Caption = "Name: " & Plugin(plindex).sname
lblClass.Caption = "Class Count: " & Plugin(plindex).content.classes.Count
Next i
End Function
