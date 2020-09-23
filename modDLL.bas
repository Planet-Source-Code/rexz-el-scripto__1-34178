Attribute VB_Name = "modDLL"
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long


Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long


Private Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long

Public Function aGetObject(ByVal path As String, Optional ByVal ClassName As String) As Object
    On Error GoTo error:
    Dim c As TypeLibInfo
    Dim sGuid As String
    Dim fName As String
    'TLBINF32.dll
    'By querying a registered ActiveX Server

    '     for its GUID,
        'we can then look in the registry and fi
        '     nd its name,
        'which then allows us to use VB's Create
        '     Object to
        'load the class. Kind of strange, but ma
        '     y be useful.
        'Its surprising VB's GetObject doesn't
        'work on VB's own classes, so here's thi
        '     s one:
        '
        'Any object created in this manner I thi
        '     nk is on a different
        'thread from this one. This may or may n
        '     ot be true or pose
        'a problem in some cases.
        Set c = TLI.TypeLibInfoFromFile(path)
        If c.CoClasses.Count <= 0 Then

            Exit Function
        End If

        
        If ClassName = "" Then 'If the user didn't specify a class, then take the first one
            sGuid = c.CoClasses.Item(1).Guid
        Else
            sGuid = c.CoClasses.NamedItem(ClassName).Guid  'If they did, try and Get that class by name
        End If
        fName = c.CoClasses.Item(1).Parent & "." & c.CoClasses.Item(1).Name
        If fName > "" Then
        Set aGetObject = CreateObject(fName)
        End If
        Exit Function
error:
        MsgBox "This DLL is not registered.", vbCritical + vbOKOnly, "Not Registered"
    End Function


Private Function GetClassString(ByVal sGuid As String) As String
    Const HKEY_CLASSES_ROOT = &H80000000
    Dim lpSubKey As String
    Dim cData As Long, sData As String, ordType As Long, e As Long
    Dim hKey As Long
    lpSubKey = "CLSID\" & sGuid
    e = RegOpenKeyEx(HKEY_CLASSES_ROOT, lpSubKey, 0, 1, hKey)
    e = RegQueryValueExLong(hKey, "", 0&, ordType, 0&, cData)
    sData = String(cData, 0)
    e = RegQueryValueExStr(hKey, "", 0&, ordType, sData, cData)
    RegCloseKey (hKey)
    GetClassString = sData
End Function
