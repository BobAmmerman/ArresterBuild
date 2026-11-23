Attribute VB_Name = "Module1"
Option Explicit

Dim appRSWork As rslogix500.Application
Dim prjLadder As rslogix500.Project

Public Sub main()
    Dim pfile As ProgramFile
    
    'appRSWork = RSLogix500.Application
    
    Set prjLadder = rslogix500.FileOpen("d:\vb projects\project template\resources\1973mc20.rss", False, False, True)
    
    For Each pfile In prjLadder.ProgramFiles
        
     Next pfile
    
End Sub

