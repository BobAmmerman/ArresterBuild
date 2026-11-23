Attribute VB_Name = "modPolymerBuild"
Option Explicit

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function CreateModBuild(NoSheds As Integer, NoHeatSinks As Integer, NoSpacers As Integer, NoDisks As Integer, NoOf5kVDisks As Integer, ORingCnt As Integer) As Variant()
' Created:  9/10/03     By: DLimburg
' Purpose:  Populate array placing required components in proper positon within module.
' Input:    NoSheds, NoHeatSinks, NoSpacers, NoDisks, ORingCnt
' Output:   array of components
' Revise:   12/10/03    By: DLimburg
' Revision: 12/10/03 chg'd test "If (ORingReqd = True And DiskCounter <= NoOfDiskInGroup " from
'           "If (ORingReqd = True And DiskCounter < NoOfDiskInGroup" because application would miss
'           the first o'ring when there was only one disk in the group.
'           10/31/03 Added code to handle case of 1 disk and 1 spacer with orings by placing spacer first.
'           10/30/03 To allow for the need to have 8 heat sinks added case 8 to Select Case NoHeatSinks.
'           Modified test for placement of oring to insure first oring is placed when only one disk is required.
'           Added code to handle case of 1 disk and 1 spacer w/o orings by placing spacer first.
'           10/23/03 Changed input ORingReqd to ORingCnt as Integer.  Procedure does not work with
'           actual number of orings, but insure that one is against each disk surface.
'           10/15/03 Added input of ORingReqd As Boolean.  Function will include the o'ring assembly.
'           in the build if set to true, otherwise it is left out.  This modification made to allow
'           the previous and the new style arresters to be build at the same time.
'
'           Change 2009_1 DLimburg 11Feb2009
'           Add code to insure 5kV disks are placed in the top of the arreter module,
'           once 5kV disk are placed the remainer are 3kV by default
'*****************************************************************************************

Dim Build() As Variant
Dim BuildCnt As Integer
Dim DiskCounter As Integer
Dim DiskHolder As Integer
Dim GroupCounter As Integer
Dim GroupHolder As Integer
Dim HeatSinkCount As Integer
Dim HeatSinkBottomCnt As Integer
Dim HeatSinkTopCnt As Integer
Dim NoOfDiskInGroup As Integer
Dim NoOfGroups As Integer
Dim ORingReqd As Boolean
Dim ReqDisk As Integer
Dim ReqSpacers As Integer
Dim TotalDisk As Integer
Dim TotalSpacer As Integer
Dim X As Double
Dim SQL As String
Dim Counter4_5kVDisks As Integer    ' change 2009_1
Dim DiskName As String              ' change 2009_1

        TotalDisk = 0
        ReqDisk = NoDisks
        DiskHolder = NoDisks
' Change 2009_1  start
        Counter4_5kVDisks = NoOf5kVDisks
        
        If Counter4_5kVDisks > 0 Then
            DiskName = "5kV MOV Disk"
        Else
            DiskName = "3kV MOV Disk"
        End If
'  Change 2009_1   end

        TotalSpacer = 0
        GroupCounter = 0
        BuildCnt = 0
        
        If ORingCnt > 0 Then
            ORingReqd = True
        Else
            ORingReqd = False
        End If
        
        ReDim Build(NoHeatSinks + NoSpacers + NoDisks + 2 + ORingCnt)
        
        Select Case NoHeatSinks
        'define number of heat sinks to be in the top and bottom of module.
        'there are never more than 7 head sinks
            Case 0
                HeatSinkBottomCnt = 0
                HeatSinkTopCnt = 0
            Case 1
                HeatSinkBottomCnt = 1
                HeatSinkTopCnt = 0
            Case 2
                HeatSinkBottomCnt = 2
                HeatSinkTopCnt = 0
            Case 3
                HeatSinkBottomCnt = 3
                HeatSinkTopCnt = 0
            Case 4
                HeatSinkBottomCnt = 4
                HeatSinkTopCnt = 0
            Case 5
                HeatSinkBottomCnt = 3
                HeatSinkTopCnt = 2
            Case 6
                HeatSinkBottomCnt = 3
                HeatSinkTopCnt = 3
            Case 7
                HeatSinkBottomCnt = 4
                HeatSinkTopCnt = 3
            Case 8
                HeatSinkBottomCnt = 4
                HeatSinkTopCnt = 4
        End Select
    '*****************************************************************************************
        
            Build(BuildCnt) = "Electrode - module top"
            BuildCnt = BuildCnt + 1
    ' ****************************************************************************************
        HeatSinkCount = 0
        ' up to four heat sinks are always placed in the bottom of the stack with remainer going in top
            Do Until HeatSinkCount = HeatSinkTopCnt   ' loop 3 - start
                Build(BuildCnt) = "Heat Sink"
                BuildCnt = BuildCnt + 1
                HeatSinkCount = HeatSinkCount + 1
            Loop
    
    
    ' ****************************************************************************************
        ' add groups of disks
            If (ORingReqd = False And NoDisks = 1 And NoSpacers = 1) Then
            ' this is a unique case on one disk and one spacer w/o orings.
                Build(BuildCnt) = "Spacer"
                BuildCnt = BuildCnt + 1
                Build(BuildCnt) = DiskName
                Counter4_5kVDisks = Counter4_5kVDisks - 1
                If Counter4_5kVDisks > 0 Then
                    DiskName = "5kV MOV Disk"
                Else
                    DiskName = "3kV MOV Disk"
                End If
                BuildCnt = BuildCnt + 1
            Else
                If (ORingReqd = True And NoDisks = 1 And NoSpacers = 1) Then
                ' this is a unique case on one disk and one spacer with orings.
                    Build(BuildCnt) = "Spacer"
                    BuildCnt = BuildCnt + 1
                    Build(BuildCnt) = "O'Ring Assembly"
                    BuildCnt = BuildCnt + 1
                    Build(BuildCnt) = DiskName
                    Counter4_5kVDisks = Counter4_5kVDisks - 1
                    If Counter4_5kVDisks > 0 Then
                        DiskName = "5kV MOV Disk"
                    Else
                        DiskName = "3kV MOV Disk"
                    End If
                    BuildCnt = BuildCnt + 1
                    Build(BuildCnt) = "O'Ring Assembly"
                    BuildCnt = BuildCnt + 1
                Else
                    NoOfGroups = NoSpacers + 1
                    GroupHolder = NoOfGroups
                    Do Until GroupCounter = NoOfGroups                                                      ' loop 4 - start
                        DiskCounter = 1
                        If NoOfGroups = 1 Then                                                              ' if 6 - start
                            NoOfDiskInGroup = NoDisks
                        Else                                                                                ' if 6 - else
                            If (DiskHolder / GroupHolder) > 1 And DiskHolder Mod GroupHolder <> 0 Then      ' if 7 - start
                                X = (DiskHolder / GroupHolder) + 0.5
                                NoOfDiskInGroup = X
                            Else                                                                            ' if 7 - else
                                If DiskHolder / GroupHolder < 1 Then
                                    NoOfDiskInGroup = 1
                                Else
                                    NoOfDiskInGroup = (DiskHolder / GroupHolder)
                                End If
                           End If                                                                           ' if 7 - end
                        End If                                                                              ' if 6 - end
                        
                        If (ORingReqd = True And DiskCounter <= NoOfDiskInGroup And _
                            TotalDisk < NoDisks) Or (ORingReqd = True And NoDisks = 1) Then
                            Build(BuildCnt) = "O'Ring Assembly"
                            BuildCnt = BuildCnt + 1
                        End If
                        
                        Do Until DiskCounter > NoOfDiskInGroup Or TotalDisk = NoDisks                   ' loop 5 - start
                        ' add disks if called for by build
            
                            Build(BuildCnt) = DiskName
                            Counter4_5kVDisks = Counter4_5kVDisks - 1
                            If Counter4_5kVDisks > 0 Then
                                DiskName = "5kV MOV Disk"
                            Else
                                DiskName = "3kV MOV Disk"
                            End If
                            TotalDisk = TotalDisk + 1
                            BuildCnt = BuildCnt + 1
                            ReqDisk = ReqDisk - 1
                            DiskCounter = DiskCounter + 1
                            
                            If ORingReqd = True Then
                                Build(BuildCnt) = "O'Ring Assembly"
                                BuildCnt = BuildCnt + 1
                            End If
            
                        Loop                                                                            ' loop 5 - end
                        
                        GroupCounter = GroupCounter + 1
                                    
                        ' add correct spacer
                        If TotalSpacer < NoSpacers Then            ' if 1 - start
                        'add spacer unless all required spacers have already been added.
                            Build(BuildCnt) = "Spacer"
                            BuildCnt = BuildCnt + 1
                            TotalSpacer = TotalSpacer + 1
                            ReqSpacers = ReqSpacers - 1
                        End If
                        
                        DiskHolder = DiskHolder - NoOfDiskInGroup
                        GroupHolder = GroupHolder - 1
                 
                    Loop
                    ' loop 4 - end
                End If
            End If
    '*****************************************************************************************
    '   Heat Sink to be placed at the Bottom of the arrester
        HeatSinkCount = 0
        Do Until HeatSinkCount = HeatSinkBottomCnt    ' loop 3 - start
            Build(BuildCnt) = "Heat Sink"
            BuildCnt = BuildCnt + 1
            HeatSinkCount = HeatSinkCount + 1
        Loop                                                ' loop 3 -end
        Build(BuildCnt) = "Electrode - module bottom"
    
    '*****************************************************************************************

                

    
    CreateModBuild = Build
    
    Exit Function
errorinmod:
    MsgBox "error in CreateModBuild"
End Function
'*****************************************************************************************




