Attribute VB_Name = "modDebugBuild"
Option Explicit

'--main database connection
'Global adosetupdata As ADODB.Connection
'--adoRS = temp recordset, adoSerials = temp recordset for retrieving sn's
Global adoRS, adoSerials As ADODB.Recordset
Global dbconSetupData As New ADODB.Connection
Global dbconResultsData As New ADODB.Connection

Public Function ConnectDatabases() As Boolean
    '------------------------------------------------------------------------------------
    '   Created:    03/21/01    By:Ray Raf
    '   Purpose:    Establish connection to DistrTester databases:
    '   Returns:    True if connection was established, false otherwise
    '------------------------------------------------------------------------------------
    
    Dim UserName As String
    Dim Password As String
    Dim ServerName As String
    Dim sqlDatabase As String
    Dim ConnectionString As String
    Set adoRS = New ADODB.Recordset
    Set adoSerials = New ADODB.Recordset
    
    On Error GoTo ErrorHandler
    'Get Setup database connection parameters
    ServerName = GetSet(AppName:="Arrester Build Program", Section:="DB", _
            Key:="SetupReadOnly_Server", Default:="")
    sqlDatabase = GetSet(AppName:="Arrester Build Program", Section:="DB", _
            Key:="SetupReadOnly_Database", Default:="")
    UserName = GetSet(AppName:="Arrester Build Program", Section:="DB", _
            Key:="SetupReadOnly_UID", Default:="")
    Password = GetSet(AppName:="Arrester Build Program", Section:="DB", _
            Key:="SetupReadOnly_PWD", Default:="")

        'use SQL connection
        ConnectionString = "driver={SQL Server};" & _
            "server=" & ServerName & ";" & _
            "uid=" & UserName & ";" & _
            "pwd=" & Password & ";" & _
            "database=" & sqlDatabase & ";"
        'make sure connection is closed
        If dbconSetupData.State = 1 Then
            dbconSetupData.Close
        End If
        
        With dbconSetupData
            .ConnectionString = ConnectionString
            .Open
        End With
    
        
        
    'Get Results database connection parameters
    ServerName = GetSet(AppName:="Arrester Build Program", Section:="DB", _
            Key:="Results_Server", Default:="")
    sqlDatabase = GetSet(AppName:="Arrester Build Program", Section:="DB", _
            Key:="Results_Database", Default:="")
    UserName = GetSet(AppName:="Arrester Build Program", Section:="DB", _
            Key:="Results_UID", Default:="")
    Password = GetSet(AppName:="Arrester Build Program", Section:="DB", _
            Key:="Results_PWD", Default:="")

    ConnectionString = "driver={SQL Server};" & _
        "server=" & ServerName & ";" & _
        "uid=" & UserName & ";" & _
        "pwd=" & Password & ";" & _
        "database=" & sqlDatabase & ";"
    If dbconResultsData.State = 1 Then
        dbconResultsData.Close
    End If

    With dbconResultsData
        .ConnectionString = ConnectionString
        .Open
    End With

    ConnectDatabases = True
    Exit Function
ErrorHandler:
    If Err.Number = -2147467259 Then
        'can't connect to database server
        MsgBox "Cannot connect to Database server!"
    Else
    End If
    Err.Clear
    ConnectDatabases = False
End Function



