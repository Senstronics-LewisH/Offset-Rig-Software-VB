Attribute VB_Name = "Vision"
Option Explicit

Public Sub SendvisionCommand(ByVal Command As String)
    
      MainForm.VisionComm.Output = Command & vbCr
    
End Sub
Public Sub InitialiseVision()

        MainForm.VisionComm.PortOpen = True
        SendvisionCommand "R0"
        SendvisionCommand "T1"

End Sub
Public Sub ChangeVisionProgram(ProgramNumber)

    SendvisionCommand "PW,1," & ProgramNumber
    
End Sub
Public Sub RecieveVision()
  
    Dim Reply As String
    Dim T1Reply As String
    Dim T2Reply As String
    Dim T3Reply As String
    Dim T4Reply As String

'T101 = oring present if 0 not present if 1
'T100 = restrictor present if 0 not present if 1
'T104 = restrictorwelded if 0 yes if 1 no
'T102 = uniontype if 0 yes if 1 no

    Reply = MainForm.VisionComm.Input
    SendvisionCommand "T1"
    
    Sleep 500
    Reply = MainForm.VisionComm.Input
    
    If Len(Reply) > 12 Then
        SendvisionCommand "T1"
        Sleep 500
        Reply = MainForm.VisionComm.Input
        Sleep 500
    End If
    
    T1Reply = Mid$(Reply, 4, 1)
    T2Reply = Mid$(Reply, 6, 1)
    T3Reply = Mid$(Reply, 8, 1)
    T4Reply = Mid$(Reply, 10, 1)
        
'T101 = oring present if 0 not present if 1

    If T1Reply = "0" Then
        If ORingRequired = True Then
            MainForm.OringPass.Visible = True
            ORingResult = True
        Else
            MainForm.OringFail.Visible = True
            SensorStatus(MainForm.SensorID) = FailedOring
            ORingResult = False
        End If
    Else
        If ORingRequired = True Then
            MainForm.OringFail.Visible = True
            SensorStatus(MainForm.SensorID) = FailedOring
            ORingResult = False
        Else
            MainForm.OringPass.Visible = True
            ORingResult = True
        End If
    End If
    
'T100 = restrictor present if 0 not present if 1

    If T2Reply = "0" Then
        If RestrictorRequired = True Then
            MainForm.RestrictorPass.Visible = True
            RestrictorResult = True
        Else
            MainForm.RestrictorFail.Visible = True
            SensorStatus(MainForm.SensorID) = FailedRestPresent
            RestrictorResult = False
        End If
    Else
        If RestrictorRequired = True Then
            MainForm.RestrictorFail.Visible = True
            SensorStatus(MainForm.SensorID) = FailedRestPresent
            RestrictorResult = False
        Else
            MainForm.RestrictorPass.Visible = True
            RestrictorResult = True
        End If
    End If
    
'T104 = restrictorwelded if 0 yes if 1 no


    If T3Reply = "0" Then
        If RestrictorRequired = True Then
            MainForm.RestrictorWeldedPass.Visible = True
            RestrictorWelded = True
        Else
            MainForm.RestrictorWeldedFail.Visible = True
            SensorStatus(MainForm.SensorID) = FailedRestweld
            RestrictorWelded = False
        End If
    Else
        If RestrictorRequired = True Then
            MainForm.RestrictorWeldedFail.Visible = True
            SensorStatus(MainForm.SensorID) = FailedRestweld
            RestrictorWelded = False
        Else
            MainForm.RestrictorWeldedPass.Visible = True
            RestrictorWelded = True
        End If
    End If
    
'T102 = uniontype if 0 yes if 1 no
    
        If T4Reply = "0" Then
            MainForm.UnionPass.Visible = True
            CorrectUnion = True
        Else
            MainForm.UnionFail.Visible = True
            SensorStatus(MainForm.SensorID) = FailedWrongUnion
            CorrectUnion = False
        End If
    
End Sub

