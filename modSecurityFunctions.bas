Attribute VB_Name = "modSecurityFunctions"
Option Explicit

'---------------------------------MODULE COMMENT----------------------------------
'   The functions in this module are used in conjunction with the
'  Security Option control set addin on an Options screen.
'--------------------------------------------------------------------------------------------------

'gets the currently logged in windows user
Declare Function WNetGetUserA Lib "mpr" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

Public Enum SecurityLogin
    slogWindows = 0         'use windows logins for security check
    slogProgram = 1         'use a login specific to this program.
End Enum

Public Enum SecurityLoc
    slocForm = 0            'logins are all listed on security tab of options form
    slocAccess = 1          'logins are listed in a MS Access database table
    slocSQLServer = 2       'logins are listed in a SQL Server database table
End Enum

Global globCurrentUser As String                 'Name of the current user for security purposes.
Global globCurrentAccessLevel As Integer        '-1 = no access, 0 = operator level, 1 = admin level
Global globLastActivity As Single
Public SecurityINIFileName As String

Public Const conLoginExpireTime = 600    'amount of seconds (obtained from Timer object) to
                                                         'wait before expiring login (in seconds)
Public Function VerifyUser() As Integer
    '*************************************************************************
    '   Created:    03/15/01    By: Ray Raf
    '   Purpose:    This function will get a user name and password and
    '                    verify it against a user access list.  Where the user name
    '                    and password comes from is determined by the settings
    '                    on the Security tab of the options form.
    '   Returns:    -1 for bad login, 0 for operator level access, 1 for Admin level access
    '*************************************************************************
    Dim slLogin As SecurityLogin
    Dim slLoc As SecurityLoc
    Dim strSet As String
    Dim strUserNameBuff As String * 255
    
    
    slLogin = slogProgram   'this program has no other login posibilities
    slLoc = slocSQLServer   'this program has no other location for login data
    
    Select Case slLogin
        Case slogProgram
            VerifyUser = VerifyUsingProgramLogin(slLoc)
        Case slogWindows
            VerifyUser = VerifyUsingWinUser(slLoc)
    End Select
    
    If VerifyUser >= 0 Then
        'sucessful login.  start auto logout timing
        globLastActivity = Timer
    End If

End Function

Private Function VerifyUsingWinUser(slLoc As SecurityLoc) As Integer
    '*************************************************************************
    '   Created:    03/15/01    By: Ray Raf
    '   Purpose:    The function will check access levels allowed based
    '                    on the currently logged in Windows user
    '   Inputs:     slLoc - indicates if the user name should be checked
    '                   against names for this program, an Access database,
    '                   or a SQL server database.
    '   Returns:    -1 for bad login, 0 for operator level access, 1 for Admin level access
    '*************************************************************************
    Dim iReturn As Boolean
    Dim strUserNameBuff As String * 255
    Dim strDBName As String
    
    Call WNetGetUserA(vbNullString, strUserNameBuff, 255&)
    globCurrentUser = Left$(strUserNameBuff, InStr(strUserNameBuff, vbNullChar) - 1)
    globCurrentUser = LCase(globCurrentUser)
    Select Case slLoc
        Case slocForm
            'use data stored in the registry for this program
            iReturn = VerifyUserAgainstForm(globCurrentUser)
        Case slocAccess
            iReturn = VerifyUserAgainstAccess(globCurrentUser)
        Case slocSQLServer
            iReturn = VerifyUserAgainstSQLServer(globCurrentUser, strDBName, "")
    End Select
    
    VerifyUsingWinUser = iReturn
    
End Function

Private Function VerifyUsingProgramLogin(slLoc As SecurityLoc) As Integer
    '*************************************************************************
    '   Created:    03/15/01    By: Ray Raf
    '   Purpose:    The function will check access levels allowed based
    '                    on an entered user name (and password if required)
    '   Inputs:     slLoc - indicates if the user name should be checked
    '                   against names for this program, an Access database,
    '                   or a SQL server database.
    '   Returns:    -1 for bad login, 0 for operator level access, 1 for Admin level access
    '*************************************************************************
    Dim iReturn As Integer
    Dim strPassword As String
    Dim strDBName As String
    
    frmLogin.Show vbModal
    
    If frmLogin.CheckLogin = False Then
        'user pressed cancel
        VerifyUsingProgramLogin = -1
        Unload frmLogin
        Exit Function
    End If
    
    globCurrentUser = frmLogin.txtUserName.Text
    strPassword = frmLogin.txtPassword.Text
    Unload frmLogin
    
    Select Case slLoc
        Case slocForm
            'use data stored in the registry for this program
            iReturn = VerifyUserAgainstForm(globCurrentUser, strPassword)
        Case slocAccess
            iReturn = VerifyUserAgainstAccess(globCurrentUser, strPassword)
        Case slocSQLServer
            iReturn = VerifyUserAgainstSQLServer(globCurrentUser, strDBName, strPassword)
            If iReturn >= 0 Then
                With globFormMain.sbStatusBar.Panels(1)
                    .Text = "User: " & strDBName
                    If iReturn = 0 Then
                        'show operator access
                        .Picture = globFormMain.imlIcons.ListImages("Unsecure").Picture
                    Else
                        'show admin access
                        .Picture = globFormMain.imlIcons.ListImages("Secure").Picture
                    End If
                End With
            Else
                'use was not validated
                With globFormMain.sbStatusBar.Panels(1)
                    .Text = "User: "
                    .Picture = globFormMain.imlIcons.ListImages("Stop").Picture
                End With
            End If
    End Select

    VerifyUsingProgramLogin = iReturn
    
End Function

Private Function VerifyUserAgainstForm(strUser As String, Optional strPassword As String = "") As Integer
    '*************************************************************************
    '   Created:    03/15/01    By: Ray Raf
    '   Purpose:    The function will check access levels allowed as determined by the User Table
    '                       on the security options tab.
    '   Inputs:     strUser - the user string to check access for
    '                  strPassword - the password for the user.  If no password is given, then only the user
    '                       name will be checked for, and the access level for that user returned.
    '   Returns:    -1 for bad login, 0 for operator level access, 1 for Admin level access
    '*************************************************************************
    
    VerifyUserAgainstForm = 1       'temp for debug
End Function

Private Function VerifyUserAgainstAccess(strUser As String, Optional strPassword As String = "") As Integer
    '*************************************************************************
    '   Created:    03/15/01    By: Ray Raf
    '   Purpose:    The function will check access levels allowed as determined by the
    '                       Access database selected.
    '   Inputs:     strUser - the user string to check access for
    '                  strPassword - the password for the user.  If no password is given, then only the user
    '                       name will be checked for, and the access level for that user returned.
    '   Returns:    -1 for bad login, 0 for operator level access, 1 for Admin level access
    '*************************************************************************
    
    VerifyUserAgainstAccess = 1       'temp for debug
End Function

Private Function VerifyUserAgainstSQLServer(strUser As String, strName As String, strPassword As String) As Integer
    '*************************************************************************
    '   Created:    03/15/01    By: Ray Raf
    '   Purpose:    The function will check access levels allowed as determined by the User Table
    '                       on the security options tab.
    '   Inputs:     strUser - the user string to check access for
    '                  strPassword - the password for the user.
    '   Returns:    -1 for bad login, 0 for operator level access, 1 for Admin level access
    '*************************************************************************
    Dim strServer As String
    Dim strDatabase As String
    Dim strUID As String
    Dim strPWD As String
    Dim strTable As String
    Dim ConnectionString As String
    Dim dbcon As New ADODB.Connection
    Dim rstSecurity As New ADODB.Recordset
    Dim SQL As String
    Dim strAL As String
    
    strName = ""
    
    'first get setup info for connecting to the security table
    strUID = GetSet(App.Title, "Security", "UID", "")
    strPWD = GetSet(App.Title, "Security", "Password", "")
    strDatabase = GetSet(App.Title, "Security", "Database", "")
    strServer = GetSet(App.Title, "Security", "Server", "")
    strTable = GetSet(App.Title, "Security", "Table", "")
    
    'check connection to SQL Database
    ConnectionString = "driver={SQL Server};" & _
        "server=" & strServer & ";" & _
        "uid=" & strUID & ";" & _
        "pwd=" & strPWD & ";" & _
        "database=" & strDatabase & ";"
    
    'make sure connection isn't already open
    If dbcon.State = 1 Then
        dbcon.Close
    End If
    
    'open the connection to the database
    With dbcon
        .ConnectionString = ConnectionString
        On Error Resume Next
        .Open
        If Err.Number <> 0 Then
            'had a problem making a connection.
            MsgBox "Could not make a connection to the SQL Server database specified. " & vbCrLf & _
                        "Returned message was: " & Err.Description
            ErrorLogger "VerifyUserAgainstSQLServer", Err.Description
            VerifyUserAgainstSQLServer = -1
            Err.Clear
            Load frmOptions
            Set frmOptions.tbsOptions.SelectedItem = frmOptions.tbsOptions.Tabs(3)
            frmOptions.Show vbModal
            Exit Function
        End If
    End With
    
    'generate SQL string to return a record containing the desired user number
    SQL = "SELECT * FROM " & strTable & " WHERE [EmployeeNumber] = '" & strUser & "' "
    rstSecurity.Open SQL, dbcon, adOpenDynamic, adLockReadOnly
    If Err.Number <> 0 Then
        'could not open the security table
        MsgBox "Error accessing security table in the database. " & vbCrLf & _
                        "Returned message was: " & Err.Description
        ErrorLogger "VerifyUserAgainstSQLServer", Err.Description
        VerifyUserAgainstSQLServer = -1
        Err.Clear
        Load frmOptions
        Set frmOptions.tbsOptions.SelectedItem = frmOptions.tbsOptions.Tabs(3)
        frmOptions.Show vbModal
        GoTo CloseAndExit
    End If
                
    If rstSecurity.EOF = True Then
        'if EOF is true at this time, it means that no records were returned matching the
        'generated SQL statement.  The user could not be verified against the users
        'existing in the database.  The login failed.
        strName = ""
        VerifyUserAgainstSQLServer = -1
        GoTo CloseAndExit
    End If
    
    'at this point we know that the user exists in the user data base
    strAL = Trim(LCase(rstSecurity!Access_Level))
    strName = rstSecurity!FirstName & " " & rstSecurity!LastName
    Select Case strAL
        Case "administrator"
            If Trim(strPassword) = Trim(rstSecurity!Password) Then
                VerifyUserAgainstSQLServer = 1  'has higher level access
            Else
                'invalid password
                VerifyUserAgainstSQLServer = -1
            End If
        Case "operator"
            If Trim(strPassword) = Trim(rstSecurity!Password) Then
                VerifyUserAgainstSQLServer = 0   'operator access only
            Else
                'invalid password
                VerifyUserAgainstSQLServer = -1
            End If
        Case Else
            'user exists in the table but has no privileges
            VerifyUserAgainstSQLServer = -1
    End Select

    
CloseAndExit:
    rstSecurity.Close
    dbcon.Close

End Function

Public Function CheckForValidLogin() As Boolean
    '*************************************************************************
    '   Created:    06/03/03    By: Ray Raf
    '   Purpose:    The function will check to see if the current login has expired.
    '                    The Login will expired if there has not be any activity for a given
    '                    period of time.
    '   Returns:    false if login has expired, true if it is still good
    '*************************************************************************
    Dim sglDiff As Single
    Dim iResult As Integer
    
    sglDiff = Timer - globLastActivity
    If sglDiff > conLoginExpireTime Then
        'login has expired
        MsgBox "Current login has expired!"
        iResult = VerifyUser
        globCurrentAccessLevel = iResult
        If iResult < 0 Then
            'login was not validated
            MsgBox "User login Failed.  Program is locked until valid user logs in!"
            globCurrentUser = ""
            CheckForValidLogin = False
        Else
            'new/existing user logged in sucessfully.
            globLastActivity = Timer
            CheckForValidLogin = True
            SetAccess
        End If
    Else
        'login still valid.  assume check is due to operator action
        'reset timeout start time to current time.
        globLastActivity = Timer
        CheckForValidLogin = True
    End If
        
End Function

Public Function ReadInI(strFile As String, strSectionHeader As String, _
                            strVariableName As String) As String
        MsgBox "need to fix ReadIni Function"
        
End Function

