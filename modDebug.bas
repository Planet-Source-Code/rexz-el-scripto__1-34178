Attribute VB_Name = "modDebug"
Public Errors As Long
Function ShowError(ErrNum As Integer, LineNum As Integer)
Dim Errormsg As String
Select Case ErrNum
Case 1
Errormsg = "Could not find ';' at the end of line " & LineNum
Case 2
Errormsg = "Invalid value in statement at line " & LineNum
Case 3
Errormsg = "Missing characters in statement at line " & LineNum
Case 4
Errormsg = "Requested variable at line " & LineNum & " does not exist."
Case 5
Errormsg = "Unknown command at line " & LineNum & "."
Case 6
Errormsg = "Missing parameters in statement at line " & LineNum & "."
Case 7
Errormsg = "Error opening file at line " & LineNum & "."
End Select
frmMain.Add "-Error: " & Errormsg
End Function
