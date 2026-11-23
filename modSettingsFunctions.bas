Attribute VB_Name = "modSettingsFunctions"
Option Explicit

Public Sub CopyFromSettingToRegistry(strProgName As String)
    '*************************************************************************
    '   Created:    12/21/00    By: Ray Raf
    '   Purpose: This module will copy all settings entered into the registry
    '                 using SaveSetting.  SaveSetting puts everything under the
    '                 current user so that when a new user logs into windows,
    '                 all of the entries made previously are not accessable.  This routine
    '                 will move the settings to a different area of the registry that wil
    '                be accessible whomever is logged in.
    '*************************************************************************

     Dim Res, NewLine As String
     Dim resValue As Variant
     Dim iSubKey As Long
     Dim iValue As Long
     Dim strValName As String
     Dim strMainKey As String, strSubKey As String
     Dim strBaseTarget As String
     
     strBaseTarget = "Software\AIC\" & strProgName & "\"
     iSubKey = 0
     strMainKey = "Software\VB and VBA Program Settings\" & strProgName & "\"
     Res = ReadRegistryGetSubkey(HKEY_CURRENT_USER, strMainKey, iSubKey)
     NewLine = ""
     Do Until Res = "Not Found"
        'loop thru all sub keys
        'Debug.Print Res
        strSubKey = Res
        iValue = 0
         resValue = ReadRegistryGetAll(HKEY_CURRENT_USER, strMainKey & strSubKey & "\", iValue)
        Do Until resValue(2) = "Not Found"
            'loop thru all values. (1) = value name, (2) = value
            WriteRegistry HKEY_LOCAL_MACHINE, strBaseTarget & strSubKey & "\", resValue(1), ValString, resValue(2)
' WriteRegistry HKEY_CURRENT_USER, "SOFTWARE\My Name\My App\", "NewSubKey", ValString, "NewValueHere"

            'Debug.Print strSubKey & " - " & resValue(1)
            iValue = iValue + 1
             resValue = ReadRegistryGetAll(HKEY_CURRENT_USER, strMainKey & strSubKey & "\", iValue)
        Loop
        iSubKey = iSubKey + 1
        Res = ReadRegistryGetSubkey(HKEY_CURRENT_USER, strMainKey, iSubKey)
        NewLine = Chr(13) & Chr(10)
     Loop

End Sub

Public Function GetSet(AppName As String, Section As String, Key As String, Optional Default As String = "")
    '*************************************************************************
    '   Created:    12/21/00    By: Ray Raf
    '   Purpose: This function emulates the GetSetting VB Function but
    '                   gets the settings in the HKEY_LOCAL_MACHINE Main Key
    '                   instead of the Current User Main Key.  It assumes that
    '                   all of the settings will be saved under the program name
    '                   in "Software\AIC".
    '   Inputs:     AppName - The name of the application (program).  This is the first
    '                                    section level under "Software\AIC"
    '                 Section - The individual settings can be grouped.  This is
    '                               the group name
    '                 Key - The key name for the setting desired.
    '                 Default - If supplied, will be returned by the function when the
    '                             Key name passed to this function does not exist.  Otherwise
    '                             a blank string will be returned.
    '   Returns:    the setting value, if the key exists in the registry.  If the key does not
    '                   exist, the default value (if supplied) will be returned, otherwise a blank
    '                   string.
    '*************************************************************************
    Dim strL As String
    Dim strSection As String, strKey As String
    Dim strBaseTarget As String
    
    'all settings for this program will have the same base target string
    strBaseTarget = "Software\AIC\" & AppName & "\"
    strSection = strBaseTarget & Section & "\"
    strKey = Key
    
    strL = ReadRegistry(HKEY_LOCAL_MACHINE, strSection, strKey)
    If strL = "Not Found" Then
        GetSet = Default
    Else
        GetSet = strL
    End If
    
End Function


Public Sub SaveSet(AppName As String, Section As String, Key As String, Setting As Variant)
    '*************************************************************************
    '   Created:    12/21/00    By: Ray Raf
    '   Purpose: This function emulates the SaveSetting VB Function but
    '                   saves the settings in the HKEY_LOCAL_MACHINE Main Key
    '                   instead of the Current User Main Key.  It assumes that
    '                   all of the settings will be saved under the program name
    '                   in "Software\AIC".
    '   Inputs:     AppName - The name of the application (program).  This is the first
    '                                    section level under "Software\AIC"
    '                 Section - The individual settings can be grouped.  This is
    '                               the group name.
    '                 Key - The key name for the setting desired.
    '                 Setting - The actual value to be saved
    '*************************************************************************
    Dim strBaseTarget As String
    Dim strSection As String, strKey As String
    Dim resValue As Variant
    Dim strSetting As String
    
    strSetting = CStr(Setting)
    strBaseTarget = "Software\AIC\" & AppName & "\"
    strSection = strBaseTarget & Section & "\"
    strKey = Key
    
    'HKEY_LOCAL_MACHINE and ValString are defined in modRegistryFunctions
    WriteRegistry HKEY_LOCAL_MACHINE, strSection, strKey, ValString, Setting
          
End Sub

Public Sub CopyFromRegistryToINI(strProgName As String)
    '*************************************************************************
    '   Created:    02/07/01    By: Ray Raf
    '   Purpose: This module will copy all of the settings stored in
    '                 the registry for the passed program to an INI formatted
    '                 file.  This is done to preserve registry setting in case
    '                 of a machine failure or other need of re-initializing
    '                 all machine settings to known values.
    '   Inputs:     strProgName - the program to save the settings for.
    '*************************************************************************

    Dim Res, NewLine As String
    Dim resValue As Variant
    Dim iSubKey As Long
    Dim iValue As Long
    Dim strValName As String
    Dim strMainKey As String, strSubKey As String
    'Dim strBaseTarget As String
    Dim strFileName As String
    Dim iFileHandle As Integer
    
    strFileName = InputBox("Enter the full path and file name (minus extention)" & vbCrLf & _
                                   "or leave blank to save in program directory as " & vbCrLf & _
                                   " [progname].ini ")
    
    If Trim(strFileName) = "" Then
        'use default storage location
        strFileName = App.Path & "\" & App.Title & ".ini"
    Else
        If InStr(1, strFileName, ".") = 0 Then
            'add exention if one not included
            strFileName = strFileName & ".ini"
        End If
    End If
    
    'prepare Ini File
    If Dir(strFileName) <> "" Then
        'fill already exists, delete it
        Kill strFileName
    End If
        
    iFileHandle = FreeFile
    On Error Resume Next
    Open strFileName For Output As iFileHandle
    If Err.Number <> 0 Then
        MsgBox "Error opening file for output.  Aborting operation!" & vbCrLf & _
                    Err.Number & "-" & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    
     'strBaseTarget = "Software\AIC\" & strProgName & "\"
     iSubKey = 0
     strMainKey = "Software\AIC\" & strProgName & "\"
     Res = ReadRegistryGetSubkey(HKEY_LOCAL_MACHINE, strMainKey, iSubKey)
     NewLine = ""
     Do Until Res = "Not Found"
        'loop thru all sub keys
        'Debug.Print Res
        Print #iFileHandle, "[" & Res & "]"
        strSubKey = Res
        iValue = 0
         resValue = ReadRegistryGetAll(HKEY_LOCAL_MACHINE, strMainKey & strSubKey & "\", iValue)
        Do Until resValue(2) = "Not Found"
            'loop thru all values. (1) = value name, (2) = value
            'WriteRegistry HKEY_LOCAL_MACHINE, strBaseTarget & strSubKey & "\", resValue(1), ValString, resValue(2)
            Print #iFileHandle, resValue(1) & "=" & resValue(2)
            iValue = iValue + 1
             resValue = ReadRegistryGetAll(HKEY_LOCAL_MACHINE, strMainKey & strSubKey & "\", iValue)
        Loop
        iSubKey = iSubKey + 1
        Res = ReadRegistryGetSubkey(HKEY_LOCAL_MACHINE, strMainKey, iSubKey)
        NewLine = Chr(13) & Chr(10)
     Loop

CloseAndExit:

    Close #iFileHandle
End Sub

