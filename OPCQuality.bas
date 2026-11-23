Attribute VB_Name = "OPCQuality"
Option Explicit

Global Const OPC_QUALITY_BAD As Long = &H0
Global Const OPC_QUALITY_UNCERTAIN As Long = &H40
Global Const OPC_QUALITY_GOOD      As Long = &HC0
Global Const OPC_QUALITY_CONFIG_ERROR As Long = &H4
Global Const OPC_QUALITY_NOT_CONNECTED As Long = &H8
Global Const OPC_QUALITY_DEVICE_FAILURE As Long = &HC
Global Const OPC_QUALITY_SENSOR_FAILURE As Long = &H10
Global Const OPC_QUALITY_LAST_KNOWN   As Long = &H14
Global Const OPC_QUALITY_COMM_FAILURE As Long = &H18
Global Const OPC_QUALITY_OUT_OF_SERVICE As Long = &H1C
Global Const OPC_QUALITY_LAST_USABLE  As Long = &H44
Global Const OPC_QUALITY_SENSOR_CAL   As Long = &H50
Global Const OPC_QUALITY_EGU_EXCEEDED As Long = &H54
Global Const OPC_QUALITY_SUB_NORMAL   As Long = &H58
Global Const OPC_QUALITY_LOCAL_OVERRIDE As Long = &HD8
Global Const OPC_STATUS_RUNNING As Long = &H1
Global Const OPC_STATUS_FAILED As Long = &H2
Global Const OPC_STATUS_NOCONFIG As Long = &H3
Global Const OPC_STATUS_SUSPENDED As Long = &H4
Global Const OPC_STATUS_TEST As Long = &H5

Global Const OPC_DS_CACHE As Long = 1
Global Const OPC_DS_DEVICE As Long = 2

Public Function GetQualityString(lQuality As Long) As String

    Dim sStatus As String
    
    On Error Resume Next
    
    sStatus = "??"
    
    Select Case lQuality
        Case OPC_QUALITY_BAD
            sStatus = "Bad"
        Case OPC_QUALITY_UNCERTAIN
            sStatus = "Uncertain"
        Case OPC_QUALITY_GOOD
            sStatus = "Good"
        Case OPC_QUALITY_CONFIG_ERROR
            sStatus = "Bad - Configuration error"
        Case OPC_QUALITY_NOT_CONNECTED
            sStatus = "Bad - Not connected"
        Case OPC_QUALITY_DEVICE_FAILURE
            sStatus = "Bad - Device failure"
        Case OPC_QUALITY_SENSOR_FAILURE
            sStatus = "Bad - Sensor failure"
        Case OPC_QUALITY_LAST_KNOWN
            sStatus = "Bad - Last known value"
        Case OPC_QUALITY_COMM_FAILURE
            sStatus = "Bad - Communication failure"
        Case OPC_QUALITY_OUT_OF_SERVICE
            sStatus = "Bad - Out of service"
        Case OPC_QUALITY_LAST_USABLE
            sStatus = "Uncertain - Last usable value"
        Case OPC_QUALITY_SENSOR_CAL
            sStatus = "Uncertain - Sensor not accurate"
        Case OPC_QUALITY_EGU_EXCEEDED
            sStatus = "Uncertain - EGU exceeded"
        Case OPC_QUALITY_SUB_NORMAL
            sStatus = "Uncertain - Sub-normal"
        Case OPC_QUALITY_LOCAL_OVERRIDE
            sStatus = "Good - Local override"
    End Select
    GetQualityString = sStatus
End Function
