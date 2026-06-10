Attribute VB_Name = "Barcode"
Option Explicit

Public Function SECONDMCSBARCODE()

    Dim i As Long
   
    
  
    
    If IsValidAnaloguePartNumber Then
        Dim ElementType As String
        ElementType = Mid$(PARTNUMBER, ELEMENT_TYPE_INDEX, ELEMENT_TYPE_WIDTH)
        
        Select Case ElementType
            Case "W", "F", "H", "M"
            Case Else
                IsValidAnaloguePartNumber = False
        End Select
    End If
    
    ' Check Pressure Rating
    If IsValidAnaloguePartNumber Then
        Dim PressureRating As Long
        Dim PressureRatingSubString As String
        PressureRatingSubString = Mid$(PARTNUMBER, PRESSURE_RATING_INDEX, PRESSURE_RATING_WIDTH)
        
        Found = False
        If (Mid$(PressureRatingSubString, 1, 1) = "B") Then
            PressureRating = Val(Mid$(PressureRatingSubString, 2, 4))
            For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
                If AnalogueSensorCombinations.PressureRating(i).BarPressure = PressureRating Then
                    Found = True
                End If
            Next
        Else
            PressureRating = Val(PressureRatingSubString)
            For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
                If AnalogueSensorCombinations.PressureRating(i).PsiPressure = PressureRating Then
                    Found = True
                End If
            Next
        End If
        
        If Found = False Then
            IsValidAnaloguePartNumber = False
        End If
    End If
    
'    If IsValidAnaloguePartNumber Then
'        Dim ElementConnection As String
'        ElementConnection = Mid$(PartNumber, ELEMENT_CONNECTION_INDEX, ELEMENT_CONNECTION_WIDTH)
'        Found = False
'        For i = 0 To AnalogueSensorCombinations.NumElementConnections - 1
'            If AnalogueSensorCombinations.ElementConnection(i) = ElementConnection Then
'                Found = True
'            End If
'        Next
'        If Found = False Then
'            IsValidAnaloguePartNumber = False
'        End If
'    End If
    
    ' The pressure union is not a vital part of calibration
    ' We do not care if it is entered correctly or not
'    If IsValidAnaloguePartNumber Then
'        Dim PressureUnion As String
'        PressureUnion = Mid$(PartNumber, UNION_INDEX, UNION_WIDTH)
'        Found = False
'        For i = 0 To AnalogueSensorCombinations.NumPressureUnions - 1
'            If AnalogueSensorCombinations.PressureUnion(i) = PressureUnion Then
'                Found = True
'            End If
'        Next
'        If Found = False Then
'            IsValidAnaloguePartNumber = False
'        End If
'    End If
    

    ' Check spike protector
'    If IsValidAnaloguePartNumber Then
'        Dim SpikeProtection As String
'        SpikeProtection = Mid$(PartNumber, PRESSURE_SPIKE_PROTECTION_INDEX, PRESSURE_SPIKE_PROTECTION_WIDTH)
'
'        If SpikeProtection <> "0" And SpikeProtection <> "1" Then
'            IsValidAnaloguePartNumber = False
'        End If
'    End If
    
'    If IsValidAnaloguePartNumber Then
'        Dim ConnectorAndPcbAssembly As String
'        ConnectorAndPcbAssembly = Mid$(PartNumber, CONNECTOR_AND_PCB_ASEMBLY_INDEX, CONNECTOR_AND_PCB_ASEMBLY_WIDTH)
'        Found = False
'        For i = 0 To AnalogueSensorCombinations.NumConnectorAndPcbAssemblies - 1
'            If AnalogueSensorCombinations.ConnectorAndPcbAssembly(i) = ConnectorAndPcbAssembly Then
'                Found = True
'            End If
'        Next
'        If Found = False Then
'            IsValidAnaloguePartNumber = False
'        End If
'    End If
    
    If IsValidAnaloguePartNumber Then
        Dim OutputType As String
        OutputType = Mid$(PARTNUMBER, OUTPUT_P1_TYPE_INDEX, OUTPUT_P1_TYPE_WIDTH)
        If Not (OutputType = "A" Or OutputType = "B" Or OutputType = "C" Or OutputType = "D" Or OutputType = "") Then
            IsValidAnaloguePartNumber = False
        End If
    End If
    
' need to check outputs
End Function


Public Function IsStringNumericOnly(ByVal InputString As String) As Boolean
    Dim i As Integer
    Dim Length As Integer
    
    Length = Len(InputString)
    
    If Length > 0 Then
        IsStringNumericOnly = True
        For i = 1 To Length
            Select Case Mid$(InputString, i, 1)
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-", "+"
            Case Else
                IsStringNumericOnly = False
            End Select
        Next i
    Else
        IsStringNumericOnly = False
    End If
End Function


' This function converts a string holding a pressure in the form stored in an MCS string
' into a pressure in Bar
Public Function GetPressureFromMcsSubString(ByVal PartNumberSubString As String, ByRef Pressure As Double) As Boolean

    Dim SignMultiplier As Double
    Dim PressureDivider As Double
    GetPressureFromMcsSubString = True
    
    Select Case Left$(PartNumberSubString, 1)
    Case "-"
        SignMultiplier = -1
        PartNumberSubString = Right$(PartNumberSubString, Len(PartNumberSubString) - 1)
    Case "+"
        SignMultiplier = 1
        PartNumberSubString = Right$(PartNumberSubString, Len(PartNumberSubString) - 1)
    Case Else
        ' also support no leading sign identifier and interpret as positive
        SignMultiplier = 1
    End Select
    
    Select Case Left$(PartNumberSubString, 1)
    Case "B"
        PressureDivider = 1
        PartNumberSubString = Right$(PartNumberSubString, Len(PartNumberSubString) - 1)
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
        PressureDivider = BAR_TO_PSI
    Case Else
        GetPressureFromMcsSubString = False
    End Select
    
    If Not IsStringNumericOnly(PartNumberSubString) Then
        GetPressureFromMcsSubString = False
    End If
    
    If GetPressureFromMcsSubString Then
        Pressure = SignMultiplier * Val(PartNumberSubString) / PressureDivider
    End If
End Function

' Many values used in calibration are looked up depending on the pressure rating
' of the pack.
' This function finds the index within the AnalogueSensorCombinations.PressureRating()
' array of the supplied MCS part number
' If found, PressureRatingIndex is set and the function returns true,
' otherwise the function returns false
Public Function GetPressureRatingIndexFromAnaloguePartNumber(ByVal PARTNUMBER As String, ByRef PressureRatingIndex As Integer) As Boolean
    
    Dim i As Integer
    Dim PressureRatingSubString As String
    Dim PressureRating As Double
    
    GetPressureRatingIndexFromAnaloguePartNumber = False
    
    PressureRatingSubString = Mid$(PARTNUMBER, PRESSURE_RATING_INDEX, PRESSURE_RATING_WIDTH)
    
    If (Mid$(PressureRatingSubString, 1, 1) = "B") Then
        PressureRating = Val(Mid$(PressureRatingSubString, 2, 4))
        For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
            If AnalogueSensorCombinations.PressureRating(i).BarPressure = PressureRating Then
                PressureRatingIndex = i
                GetPressureRatingIndexFromAnaloguePartNumber = True
            End If
        Next
    Else
        PressureRating = Val(PressureRatingSubString)
        For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
            If AnalogueSensorCombinations.PressureRating(i).PsiPressure = PressureRating Then
                PressureRatingIndex = i
                GetPressureRatingIndexFromAnaloguePartNumber = True
            End If
        Next
    End If
  
End Function





' This function extracts the pressure rating information from an MCS part number
' The return value is always in bar regardless of whether originally specified
' in psi or bar
Public Function GetSensorPressureRatingFromAnaloguePartNumber(ByVal PARTNUMBER As String, ByRef BarPressureRating As Double) As Boolean
    Dim i As Long
    Dim PressureRating As Long
    Dim PressureRatingSubString As String
    
    PressureRatingSubString = Mid$(PARTNUMBER, PRESSURE_RATING_INDEX, PRESSURE_RATING_WIDTH)
    
    If (Mid$(PressureRatingSubString, 1, 1) = "B") Then
        PressureRating = Val(Mid$(PressureRatingSubString, 2, 4))
        For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
            If AnalogueSensorCombinations.PressureRating(i).BarPressure = PressureRating Then
                BarPressureRating = AnalogueSensorCombinations.PressureRating(i).BarPressure
                GetSensorPressureRatingFromAnaloguePartNumber = True
            End If
        Next
    Else
        PressureRating = Val(PressureRatingSubString)
        For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
            If AnalogueSensorCombinations.PressureRating(i).PsiPressure = PressureRating Then
                BarPressureRating = AnalogueSensorCombinations.PressureRating(i).PsiPressure / BAR_TO_PSI
                GetSensorPressureRatingFromAnaloguePartNumber = True
            End If
        Next
    End If
    
End Function

Public Sub GetDeltaSFromFromPressureRatingIndex(ByVal PressureRatingIndex As Long, ByRef DeltaS As Double)
    
    DeltaS = AnalogueSensorCombinations.PressureRating(PressureRatingIndex).DeltaS
    
End Sub

Public Function GetDeltaSFromMCSPartNumber(ByVal PARTNUMBER As String, ByRef DeltaS As Double) As Boolean
    Dim PressureRatingIndex As Integer
    
    If (GetPressureRatingIndexFromAnaloguePartNumber(PARTNUMBER, PressureRatingIndex)) Then
        GetDeltaSFromFromPressureRatingIndex PressureRatingIndex, DeltaS
        GetDeltaSFromMCSPartNumber = True
    Else
        GetDeltaSFromMCSPartNumber = False
    End If

End Function


Private Sub GetStandardProofPressureFromPressureRatingIndex(ByVal PressureRatingIndex As Long, ByVal Side As RigSide, ByRef BarProofPressure As Double)

    BarProofPressure = AnalogueSensorCombinations.PressureRating(PressureRatingIndex).BarPressure * AnalogueSensorCombinations.PressureRating(PressureRatingIndex).StandardProofMultiplier

    ' The max available pressure on site is 2650 bar
    If BarProofPressure > GetMaxProofPressure(Side) Then
        BarProofPressure = GetMaxProofPressure(Side)
    End If

End Sub

Private Sub GetHeavyDutyProofPressureFromPressureRatingIndex(ByVal PressureRatingIndex As Long, ByVal Side As RigSide, ByRef BarProofPressure As Double)

    BarProofPressure = AnalogueSensorCombinations.PressureRating(PressureRatingIndex).BarPressure * AnalogueSensorCombinations.PressureRating(PressureRatingIndex).HeavyDutyProofMultiplier

    ' The max available pressure on site is 2650 bar
    If BarProofPressure > GetMaxProofPressure(Side) Then
        BarProofPressure = GetMaxProofPressure(Side)
    End If

End Sub


Public Function GetProofPressureFromAnaloguePartNumber(ByVal PARTNUMBER As String, ByVal Side As RigSide, ByRef BarProofPressure As Double) As Boolean
    Dim i As Long
    Dim PressureRating As Long
    Dim PressureRatingSubString As String
    
    Dim IsHeavyDuty As Boolean
    
    GetIsHeavyDutyFromMCSPartNumber PARTNUMBER, IsHeavyDuty
    
    PressureRatingSubString = Mid$(PARTNUMBER, PRESSURE_RATING_INDEX, PRESSURE_RATING_WIDTH)
    
    If (Mid$(PressureRatingSubString, 1, 1) = "B") Then
        PressureRating = Val(Mid$(PressureRatingSubString, 2, 4))
        For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
            If AnalogueSensorCombinations.PressureRating(i).BarPressure = PressureRating Then
                If (IsHeavyDuty) Then
                    GetHeavyDutyProofPressureFromPressureRatingIndex i, Side, BarProofPressure
                Else
                    GetStandardProofPressureFromPressureRatingIndex i, Side, BarProofPressure
                End If
                GetProofPressureFromAnaloguePartNumber = True
            End If
        Next
    Else
        PressureRating = Val(PressureRatingSubString)
        For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
            If AnalogueSensorCombinations.PressureRating(i).PsiPressure = PressureRating Then
                If (IsHeavyDuty) Then
                    GetHeavyDutyProofPressureFromPressureRatingIndex i, Side, BarProofPressure
                Else
                    GetStandardProofPressureFromPressureRatingIndex i, Side, BarProofPressure
                End If
                GetProofPressureFromAnaloguePartNumber = True
            End If
        Next
    End If
    
End Function


Public Sub GetAnalogueLoadFromMCSSubCode(ByVal PullUpDownCode As Long, ByRef LoadType As DigitalControlAnalogueLoadOptions, ByRef ResistorCombinationCode As Integer)
    
    If (PullUpDownCode = 0) Then
        ResistorCombinationCode = 0
        LoadType = DIG_CONTROL_ANALOGUE_LOAD_NONE
    ElseIf (PullUpDownCode >= 500) Then
        ResistorCombinationCode = PullUpDownCode - 500
        LoadType = DIG_CONTROL_ANALOGUE_LOAD_PULL_UP
    Else
        ResistorCombinationCode = PullUpDownCode
        LoadType = DIG_CONTROL_ANALOGUE_LOAD_PULL_DOWN
    End If
    
End Sub


Public Function GetAnalogueLoadFromAnaloguePartNumber(ByVal PARTNUMBER As String, ByRef LoadType As DigitalControlAnalogueLoadOptions, ByRef ResistorCombinationCode As Integer) As Boolean
    Dim PullUpDownCode As Long
    Dim PullUpDownSubString As String
    
    PullUpDownSubString = Mid$(PARTNUMBER, PULL_UP_DOWN_INDEX, PULL_UP_DOWN_WIDTH)
    PullUpDownCode = Val(PullUpDownSubString)
    
    GetAnalogueLoadFromMCSSubCode PullUpDownCode, LoadType, ResistorCombinationCode
    
    GetAnalogueLoadFromAnaloguePartNumber = True
    
End Function

Private Sub GetNonLinearityFromPressureRatingIndex(ByVal PressureRatingIndex As Long, ByRef NonLinearity As Double)
    NonLinearity = AnalogueSensorCombinations.PressureRating(PressureRatingIndex).NonLinearity
End Sub


Public Function GetNonLinearityFromAnaloguePartNumber(ByVal PARTNUMBER As String, ByRef NonLinearity As Double) As Boolean
    Dim i As Long
    Dim PressureRating As Long
    Dim PressureRatingSubString As String
    
    PressureRatingSubString = Mid$(PARTNUMBER, PRESSURE_RATING_INDEX, PRESSURE_RATING_WIDTH)
    
    If (Mid$(PressureRatingSubString, 1, 1) = "B") Then
        PressureRating = Val(Mid$(PressureRatingSubString, 2, 4))
        For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
            If AnalogueSensorCombinations.PressureRating(i).BarPressure = PressureRating Then
                GetNonLinearityFromPressureRatingIndex i, NonLinearity
                GetNonLinearityFromAnaloguePartNumber = True
            End If
        Next
    Else
        PressureRating = Val(PressureRatingSubString)
        For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
            If AnalogueSensorCombinations.PressureRating(i).PsiPressure = PressureRating Then
                GetNonLinearityFromPressureRatingIndex i, NonLinearity
                GetNonLinearityFromAnaloguePartNumber = True
            End If
        Next
    End If
    
End Function

Private Sub GetMilliVPerVFromPressureRatingIndex(ByVal PressureRatingIndex As Long, ByRef MilliVPerV As Double)
    MilliVPerV = AnalogueSensorCombinations.PressureRating(PressureRatingIndex).MilliVPerV
End Sub



Public Function GetMilliVperVFromAnaloguePartNumber(ByVal PARTNUMBER As String, ByRef MilliVPerV As Double) As Boolean
    Dim i As Long
    Dim PressureRating As Long
    Dim PressureRatingSubString As String
    
    PressureRatingSubString = Mid$(PARTNUMBER, PRESSURE_RATING_INDEX, PRESSURE_RATING_WIDTH)
    
    If (Mid$(PressureRatingSubString, 1, 1) = "B") Then
        PressureRating = Val(Mid$(PressureRatingSubString, 2, 4))
        For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
            If AnalogueSensorCombinations.PressureRating(i).BarPressure = PressureRating Then
                GetMilliVPerVFromPressureRatingIndex i, MilliVPerV
                GetMilliVperVFromAnaloguePartNumber = True
            End If
        Next
    Else
        PressureRating = Val(PressureRatingSubString)
        For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
            If AnalogueSensorCombinations.PressureRating(i).PsiPressure = PressureRating Then
                GetMilliVPerVFromPressureRatingIndex i, MilliVPerV
                GetMilliVperVFromAnaloguePartNumber = True
            End If
        Next
    End If
    
End Function



Private Sub GetMaxProofShiftsFromPressureRatingIndex(ByVal PressureRatingIndex As Long, ByRef MaxPositiveProofShiftPercent As Double, ByRef MaxNegativeProofShiftPercent As Double)
    MaxPositiveProofShiftPercent = AnalogueSensorCombinations.PressureRating(PressureRatingIndex).MaxPositiveProofShiftPercent
    MaxNegativeProofShiftPercent = AnalogueSensorCombinations.PressureRating(PressureRatingIndex).MaxNegativeProofShiftPercent
End Sub


Public Function GetMaxProofShiftPercentsFromAnaloguePartNumber(ByVal PARTNUMBER As String, ByRef MaxPositiveProofShiftPercent As Double, ByRef MaxNegativeProofShiftPercent As Double) As Boolean
    Dim i As Long
    Dim PressureRating As Long
    Dim PressureRatingSubString As String
    
    PressureRatingSubString = Mid$(PARTNUMBER, PRESSURE_RATING_INDEX, PRESSURE_RATING_WIDTH)
    
    If (Mid$(PressureRatingSubString, 1, 1) = "B") Then
        PressureRating = Val(Mid$(PressureRatingSubString, 2, 4))
        For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
            If AnalogueSensorCombinations.PressureRating(i).BarPressure = PressureRating Then
                GetMaxProofShiftsFromPressureRatingIndex i, MaxPositiveProofShiftPercent, MaxNegativeProofShiftPercent
                GetMaxProofShiftPercentsFromAnaloguePartNumber = True
            End If
        Next
    Else
        PressureRating = Val(PressureRatingSubString)
        For i = 0 To AnalogueSensorCombinations.NumPressureRatings - 1
            If AnalogueSensorCombinations.PressureRating(i).PsiPressure = PressureRating Then
                GetMaxProofShiftsFromPressureRatingIndex i, MaxPositiveProofShiftPercent, MaxNegativeProofShiftPercent
                GetMaxProofShiftPercentsFromAnaloguePartNumber = True
            End If
        Next
    End If
    
End Function



Public Function GetIsThermallyCompensatedFromMCSPartNumber(ByVal MCSPartNumber As String, ByRef IsThermallyCompensated As Boolean) As Boolean

    Dim ElementType As String
    
    GetIsThermallyCompensatedFromMCSPartNumber = True
    ElementType = Mid$(MCSPartNumber, ELEMENT_TYPE_INDEX, ELEMENT_TYPE_WIDTH)
    
    Select Case ElementType
        Case "M"
            IsThermallyCompensated = True
        Case Else
            IsThermallyCompensated = False
    End Select
        
End Function

Public Function GetIsHeavyDutyFromMCSPartNumber(ByVal MCSPartNumber As String, ByRef IsHeavyDuty As Boolean) As Boolean

    Dim ElementType As String
    
    GetIsHeavyDutyFromMCSPartNumber = True
    ElementType = Mid$(MCSPartNumber, ELEMENT_TYPE_INDEX, ELEMENT_TYPE_WIDTH)
    
    Select Case ElementType
        Case "H"
            IsHeavyDuty = True
        Case Else
            IsHeavyDuty = False
    End Select
        
End Function

Public Function GetAnalogueOutputModeFromMCSPartNumber(ByVal MCSPartNumber As String, ByVal OutputNumber As Integer, ByRef OutputMode As AnalogueOutputModeEnum) As Boolean
    Dim OutputCode As String
    GetAnalogueOutputModeFromMCSPartNumber = True
    Select Case OutputNumber
    Case 0
        OutputCode = Mid$(MCSPartNumber, OUTPUT_P1_TYPE_INDEX, OUTPUT_P1_TYPE_WIDTH)
    Case 1
        OutputCode = Mid$(MCSPartNumber, OUTPUT_P2_TYPE_INDEX, OUTPUT_P2_TYPE_WIDTH)
    Case Else
        GetAnalogueOutputModeFromMCSPartNumber = False
    End Select
    
    If GetAnalogueOutputModeFromMCSPartNumber Then
        Select Case OutputCode
        Case "A"
            OutputMode = AOM_PressureAbsoluteVoltage
        Case "B"
            OutputMode = AOM_PressureRatiometricVoltage
        Case "C"
            OutputMode = AOM_PressureCurrent
        Case "D"
            OutputMode = AOM_PressureAbsoluteVoltage
        Case "H"
            OutputMode = AOM_PressureRatiometricVoltage
        Case "S"
            OutputMode = AOM_PressureSwitch
        Case "T"
            If Mid$(MCSPartNumber, OUTPUT_P1_TYPE_INDEX, OUTPUT_P1_TYPE_WIDTH) = "A" Then
                OutputMode = AOM_TemperatureAbsoluteVoltage
            Else
                OutputMode = AOM_TemperatureRatiometricVoltage
            End If

        Case Else
            GetAnalogueOutputModeFromMCSPartNumber = False
            OutputMode = AOM_None
        End Select
        
        If RigConfig.ProofOnly Then
            OutputMode = AOM_PressureAbsoluteVoltage
        End If
        
    End If

End Function


Public Function GetOutputMultiplierFromMCSPartNumber(ByVal MCSPartNumber As String, ByVal OutputNumber As Integer, ByRef Multipler As Integer) As Boolean
    Dim OutputCharacterCode As String
    
    GetOutputMultiplierFromMCSPartNumber = True
    
    Select Case OutputNumber
    Case 0
        OutputCharacterCode = Mid$(MCSPartNumber, OUTPUT_P1_TYPE_INDEX, OUTPUT_P1_TYPE_WIDTH)
    Case 1
        OutputCharacterCode = Mid$(MCSPartNumber, OUTPUT_P2_TYPE_INDEX, OUTPUT_P2_TYPE_WIDTH)
    End Select

    Select Case OutputCharacterCode
    Case "A"    ' Absolute voltage * 10 volts
        Multipler = 10
    Case "B"     ' Ratiometric voltage * 100 volts
        Multipler = 100
    Case "C"    ' Current * 10 mA or / 100 Amps
        Multipler = 10000
    Case "D"    '  Absolute voltage * 100 volts
        Multipler = 100
    Case "H"
        Multipler = 10
    Case "T"
        If Mid$(MCSPartNumber, OUTPUT_P1_TYPE_INDEX, OUTPUT_P1_TYPE_WIDTH) = "A" Then
            Multipler = 10
        Else
            Multipler = 100
        End If
    Case Else
        GetOutputMultiplierFromMCSPartNumber = False
    End Select

End Function

' This function extracts the target offset, either in volts or amps from the part number
Public Function GetTargetOffsetFromMCSPartNumber(ByVal MCSPartNumber As String, ByVal OutputNumber As Integer, ByRef TargetOffset As Double) As Boolean
    Dim OutputInteger As Integer
    Dim OutputMultiplier As Integer
    Dim OutputSubString As String
    
    GetTargetOffsetFromMCSPartNumber = True

    Select Case OutputNumber
        Case 0
            OutputSubString = Mid$(MCSPartNumber, OUTPUT_P1_ELECTRICAL_OFFSET_INDEX, OUTPUT_P1_ELECTRICAL_OFFSET_WIDTH)
        Case 1
            OutputSubString = Mid$(MCSPartNumber, OUTPUT_P2_ELECTRICAL_OFFSET_INDEX, OUTPUT_P2_ELECTRICAL_OFFSET_WIDTH)
        Case Else
            GetTargetOffsetFromMCSPartNumber = False
        End Select
    
        If (IsStringNumericOnly(OutputSubString)) Then
            OutputInteger = Val(OutputSubString)
        Else
            GetTargetOffsetFromMCSPartNumber = False
        End If
        
        If GetOutputMultiplierFromMCSPartNumber(MCSPartNumber, OutputNumber, OutputMultiplier) Then
            TargetOffset = OutputInteger / OutputMultiplier
        Else
            GetTargetOffsetFromMCSPartNumber = False
        End If
    
End Function

Public Function GetAppliedMinPressureFromMCSPartNumber(ByVal MCSPartNumber As String, ByVal OutputNumber As Integer, ByRef Pressure As Double) As Boolean
    
    GetAppliedMinPressureFromMCSPartNumber = True

    Select Case OutputNumber
    Case 0
        GetAppliedMinPressureFromMCSPartNumber = GetPressureFromMcsSubString(Mid$(MCSPartNumber, OUTPUT_P1_PRESSURE_OFFSET_INDEX, OUTPUT_P1_PRESSURE_OFFSET_WIDTH), Pressure)
    Case 1
        GetAppliedMinPressureFromMCSPartNumber = GetPressureFromMcsSubString(Mid$(MCSPartNumber, OUTPUT_P2_PRESSURE_OFFSET_INDEX, OUTPUT_P2_PRESSURE_OFFSET_WIDTH), Pressure)
    Case Else
        GetAppliedMinPressureFromMCSPartNumber = False
    End Select

    
End Function

Public Function GetAppliedMinTemperatureFromMCSPartNumber(ByVal MCSPartNumber As String, ByRef Temperature As Double) As Boolean
    Dim TemperatureSubString As String
    
    TemperatureSubString = Mid$(MCSPartNumber, OUTPUT_P2_PRESSURE_OFFSET_INDEX, OUTPUT_P2_PRESSURE_OFFSET_WIDTH)
    
    If IsStringNumericOnly(TemperatureSubString) Then
        Temperature = Val(TemperatureSubString)
        GetAppliedMinTemperatureFromMCSPartNumber = True
    Else
        GetAppliedMinTemperatureFromMCSPartNumber = False
    End If
End Function

Public Function GetAppliedMaxTemperatureFromMCSPartNumber(ByVal MCSPartNumber As String, ByRef Temperature As Double) As Boolean
    Dim TemperatureSubString As String
    
    TemperatureSubString = Mid$(MCSPartNumber, OUTPUT_P2_PRESSURE_MAX_INDEX, OUTPUT_P2_PRESSURE_MAX_WIDTH)
    
    If IsStringNumericOnly(TemperatureSubString) Then
        Temperature = Val(TemperatureSubString)
        GetAppliedMaxTemperatureFromMCSPartNumber = True
    Else
        GetAppliedMaxTemperatureFromMCSPartNumber = False
    End If
    
End Function

Public Function GetAppliedMaxPressureFromMCSPartNumber(ByVal MCSPartNumber As String, ByVal OutputNumber As Integer, ByRef Pressure As Double) As Boolean
    
    GetAppliedMaxPressureFromMCSPartNumber = True

    Select Case OutputNumber
    Case 0
        GetAppliedMaxPressureFromMCSPartNumber = GetPressureFromMcsSubString(Mid$(MCSPartNumber, OUTPUT_P1_PRESSURE_MAX_INDEX, OUTPUT_P1_PRESSURE_MAX_WIDTH), Pressure)
    Case 1
        GetAppliedMaxPressureFromMCSPartNumber = GetPressureFromMcsSubString(Mid$(MCSPartNumber, OUTPUT_P2_PRESSURE_MAX_INDEX, OUTPUT_P2_PRESSURE_MAX_WIDTH), Pressure)
    Case Else
        GetAppliedMaxPressureFromMCSPartNumber = False
    End Select

    
End Function

Public Function GetTargetMaxOutputFromMCSPartNumber(ByVal MCSPartNumber As String, ByVal OutputNumber As Integer, ByRef TargetMaxOutput As Double) As Boolean
    Dim OutputInteger As Integer
    Dim OutputMultiplier As Integer
    Dim OutputSubString As String
    
    GetTargetMaxOutputFromMCSPartNumber = True

    Select Case OutputNumber
    Case 0
        OutputSubString = Mid$(MCSPartNumber, OUTPUT_P1_ELECTRICAL_MAX_INDEX, OUTPUT_P1_ELECTRICAL_MAX_WIDTH)
    Case 1
        OutputSubString = Mid$(MCSPartNumber, OUTPUT_P2_ELECTRICAL_MAX_INDEX, OUTPUT_P2_ELECTRICAL_MAX_WIDTH)
    Case Else
        GetTargetMaxOutputFromMCSPartNumber = False
    End Select

    If (IsStringNumericOnly(OutputSubString)) Then
        OutputInteger = Val(OutputSubString)
    Else
        GetTargetMaxOutputFromMCSPartNumber = False
    End If
    
    If GetOutputMultiplierFromMCSPartNumber(MCSPartNumber, OutputNumber, OutputMultiplier) Then
        TargetMaxOutput = OutputInteger / OutputMultiplier
    Else
        GetTargetMaxOutputFromMCSPartNumber = False
    End If
  '    TargetMaxOutput = 0.0199
End Function

Public Function ExtractDigitalPartNumberFromMCSPartNumber(ByVal MCSPartNumber As String) As String
    
    ExtractDigitalPartNumberFromMCSPartNumber = Mid$(MCSPartNumber, 19, 3) & Mid$(MCSPartNumber, 23, 3)

End Function

Public Function IsValidDigitalPartNumber(ByVal PARTNUMBER As String) As Boolean
    Dim i As Long
    Dim DigitalPartNumber As String
    
    IsValidDigitalPartNumber = False
    
    If Len(PARTNUMBER) >= 25 Then
        If (Mid$(PARTNUMBER, 18, 1) = "F") Then
            DigitalPartNumber = ExtractDigitalPartNumberFromMCSPartNumber(PARTNUMBER)
            
            For i = 0 To NumDigitalCalConfigs - 1
                If DigitalPartNumberList(i) = DigitalPartNumber Then
                    IsValidDigitalPartNumber = True
                End If
            Next
        Else
            IsValidDigitalPartNumber = False
        End If
    Else
        IsValidDigitalPartNumber = False
    End If
    

End Function

Public Function IsValidPartNumber(ByVal PARTNUMBER As String) As Boolean
    IsValidPartNumber = IsValidAnaloguePartNumber(PARTNUMBER) Or IsValidDigitalPartNumber(PARTNUMBER)
End Function
' This function adjusts the offset to compensate for negative specified input pressures
Public Sub CompensteForNegativePressure(ByRef Configuration As AnalogueCalibrationConfigurationType, ByVal OutputNumber As Integer)
Dim Span As Double
Dim pressurerange As Double

    If ((Configuration.AppliedMinInput(OutputNumber) < 0) And (IsPressureOutputType(Configuration.AnalogueOutputMode(OutputNumber)))) Then
       
    Span = Configuration.TargetMaxOutput(OutputNumber) - Configuration.TargetOffset(OutputNumber)
    pressurerange = (Configuration.AppliedMaxInput(OutputNumber) - Configuration.AppliedMinInput(OutputNumber))
    
    Configuration.TargetOffset(OutputNumber) = ((Span / pressurerange)) * Abs((Configuration.AppliedMinInput(OutputNumber))) + Configuration.TargetOffset(OutputNumber)
    Configuration.AppliedMinInput(OutputNumber) = 0
    AddToHistoryLog "New Offset Target=" & Configuration.TargetOffset(OutputNumber)
    Else
        Configuration.NegativeOffsetCalibration(OutputNumber) = False
    End If
    
End Sub

' This function adjusts the configuration to compensate for slightly limited precise pressure range
' It changes the applied max pressure to the highest precisely attainable, and recalculates what
' the output voltage should be this pressure, and recalculates TargetSpan to suit
Public Function CompensteForMaxPrecisePressure(ByRef Configuration As AnalogueCalibrationConfigurationType, ByVal Side As RigSide, ByVal OutputNumber As Integer) As Boolean


    If ((Configuration.AppliedMaxInput(OutputNumber) > GetMaxPrecisePressure(Side)) And (IsPressureOutputType(Configuration.AnalogueOutputMode(OutputNumber)))) Then

        Dim m As Double
        Dim VoltagteAtMaxPressure As Double
        Dim X1 As Double
        Dim X2 As Double
        Dim Y1 As Double
        Dim Y2 As Double
        Dim X As Double
        
        X1 = Configuration.AppliedMinInput(OutputNumber)
        X2 = Configuration.AppliedMaxInput(OutputNumber)
        
        Y1 = Configuration.TargetOffset(OutputNumber)
        Y2 = Configuration.TargetMaxOutput(OutputNumber)
        
        If (X1 = X2) Then
            CompensteForMaxPrecisePressure = False
        Else
            
            m = (Y1 - Y2) / (X1 - X2)
            
            
            VoltagteAtMaxPressure = m * (GetMaxPrecisePressure(Side) - X1) + Y1
        
            Configuration.AppliedMaxInput(OutputNumber) = GetMaxPrecisePressure(Side)
            Configuration.TargetMaxOutput(OutputNumber) = VoltagteAtMaxPressure
            Configuration.TargetSpan(OutputNumber) = Configuration.TargetMaxOutput(OutputNumber) - Configuration.TargetOffset(OutputNumber)
            CompensteForMaxPrecisePressure = True
        End If
    Else
        ' no adjustment necessary
        CompensteForMaxPrecisePressure = True
    End If

End Function

Public Function MapMAToVolts(ByVal MA As Double) As Double

        ' map the output current to the voltage seen by the ASIC comparitor
    
 If (RigConfig.AsicIssue = SEN2000) Then

   MapMAToVolts = MA * 281.25 - 1.125 ' 0 volts mapped to 4ma and 4.5 volts mapped to 20ma
   'MapMAToVolts = MA * 328.13 - 1.3125 ' 0 volts mapped to 4ma and 5.25 volts mapped to 20ma
     
     Else
    MapMAToVolts = MA * 312.5 - 1.25 '0 volts mapped to 4ma and 5 volts mapped to 20ma
    
   End If
   
End Function

Public Function SetNonLinearitySwitchPoints(ByRef Configuration As AnalogueCalibrationConfigurationType) As Boolean

    ' Note that we are only setting the switch points here, i.e. the points
    ' in the output range at which each of the different non-linearity
    ' corrections are applied.
    ' The amounts of correction to be applied in each of the output bands is set
    ' elsewhere
    '
    Dim HighPercentage As Double
    Dim LowPercentage As Double
    Dim MaxOutput As Double
    Dim MinOutput As Double

    ' Note that all linearity correction here is for output 0 only.
    ' The one set of linearity corrections applies to both offset and gains,
    ' and so cannot be programmed for a sensor where final configuration
    ' is performed by a distributor. In that case, the registers are left unprogrammed
    ' and this function is not called
    
    
    MaxOutput = Configuration.TargetMaxOutput(LINEARITY_CHECK_OUTPUT)
    MinOutput = Configuration.TargetOffset(LINEARITY_CHECK_OUTPUT)

    If (Configuration.AnalogueOutputMode(LINEARITY_CHECK_OUTPUT) = AOM_PressureCurrent) Then

        MaxOutput = MapMAToVolts(MaxOutput)
        MinOutput = MapMAToVolts(MinOutput)

    End If

    ' Non-linearity correction is performed by adding in different amounts of correction
    ' depending on which part of the output range the sensor is in (bottom , middle or top).
    ' The ASIC does this by comparing the output to an internal regulated 5 volts reference.
    ' If the max output is above 5 volts, the output must first be divided by 2
    ' to bring it into the comparitor range. Note that 5.5 is used as the limit because
    ' the lowest that the HighPercentage can be set to is 60% and it is better to see
    ' 5.5 as being 110% of 5 rather than 55% of 10

    If (MaxOutput > 5.5) Then
        
        Configuration.NL_2X = 1
  'inverted for sen2000 done in asic control
  
        MaxOutput = MaxOutput / 2
        MinOutput = MinOutput / 2

    Else

        Configuration.NL_2X = 0

'        If (Configuration.AnalogueOutputMode(LINEARITY_CHECK_OUTPUT) = AOM_PressureCurrent) Then
'
'        Configuration.NL_2X = 1
'
'        End If

    End If

    ' The start and end points of the linearity correction bands must be synchronised
    ' with actual output range.
    ' The high and low ends of the adjustment band are separately adjustable,
    ' so we need to pick the start / end point nearest to the actaul start / end point
    ' For example, a 0.5 to 4.5 volt sensor would use 10% and 90% of the regulated 5 volt
    ' reference as the correction range end points
    ' The center and mid points automatically shift within the band defined by the
    ' start and end points of overall correction

    HighPercentage = (MaxOutput / 5) * 100
    LowPercentage = (MinOutput / 5) * 100

    If (RigConfig.AsicIssue = SEN2000) Then

        If (HighPercentage >= 97.5) Then
            Configuration.NL_HP = 0   ' aboive 97.5, set for 100%
        ElseIf (HighPercentage >= 92.5) Then
            Configuration.NL_HP = 5   ' 92.5 to 97.5, set for 95%
        ElseIf (HighPercentage >= 85) Then
            Configuration.NL_HP = 1   ' 85 to 92.5%, set for 90%
        ElseIf (HighPercentage >= 75) Then
            Configuration.NL_HP = 2   ' 75 to 85%, set for 80%
        ElseIf (HighPercentage >= 67.5) Then
            Configuration.NL_HP = 3   ' 67.5 to 75%, set for 70%
        ElseIf (HighPercentage >= 62.5) Then
            Configuration.NL_HP = 4   ' 62.5 to 67.5% set for 65%
        Else
            Configuration.NL_HP = 7   ' less than 62.5%, set fo 60%
        End If

        If (LowPercentage < 2.5) Then
            Configuration.NL_LP = 0   ' 0 to 5%, set for 0%
        ElseIf (LowPercentage <= 7.5) Then
            Configuration.NL_LP = 3   ' 2.5 to 7.5%, set for 5%
        ElseIf (LowPercentage <= 12.5) Then
            Configuration.NL_LP = 1   ' 7.5 to 12.5% set for 10%
        Else
            Configuration.NL_LP = 2   ' More than 12.5 set for 20%

        End If
    Else

        ' all Sen1000 varaints
        If (HighPercentage >= 95) Then
            Configuration.NL_HP = 0   ' set for 100%
        ElseIf (HighPercentage >= 85) Then
            Configuration.NL_HP = 1   ' set for 90%
        ElseIf (HighPercentage >= 75) Then
            Configuration.NL_HP = 2   ' set for 80%
        ElseIf (HighPercentage >= 65) Then
            Configuration.NL_HP = 3   ' set for 70%
        Else
            Configuration.NL_HP = 4   ' set fo 60%
        End If
        If (LowPercentage <= 5) Then
            Configuration.NL_LP = 0   ' set for 0%
        ElseIf (LowPercentage <= 15) Then
            Configuration.NL_LP = 1   ' set for 10%
        Else
            Configuration.NL_LP = 2   ' set for 20%
        End If
    End If

    AddToHistoryLog "SetNonLinearitySwitchPoints HighPercentage=" & Format$(HighPercentage) & " LowPercentage=" & Format$(LowPercentage)
    SetNonLinearitySwitchPoints = True

End Function




Private Sub SelectNegativeVGen(ByRef Configuration As AnalogueCalibrationConfigurationType, ByVal OutputNumber As Long)
    ' enable the negative voltage generator for voltage mode where
    ' the target offset is less than 100mV
    If (Configuration.AnalogueOutputMode(OutputNumber) = AOM_PressureAbsoluteVoltage And Configuration.TargetOffset(OutputNumber) < 0.2) Then
        Configuration.EnableNegativeVGen(OutputNumber) = True
        Configuration.EnableNegativeVGen(TEMPERATURE_OUTPUT) = True
    Else
        Configuration.EnableNegativeVGen(OutputNumber) = False
    End If
End Sub
        
        
Private Sub SelectTemperatureGainRegister(ByRef Configuration As AnalogueCalibrationConfigurationType)
    Dim VCi As Double
    
    Configuration.EnableTempCh = True
    
    Configuration.DeltaVc = Configuration.DeltaS * -16
    
    VCi = AnalogueSensorCombinations.RoomTemperatureVc * Configuration.DeltaVc * Configuration.AppliedTemperatureSpan / 1000000#
    
    Configuration.RequiredTemperatureGain = Configuration.TargetSpan(TEMPERATURE_OUTPUT) / VCi
    
    ' Determine best coarse gain setting, prefering lower gains
    If (Configuration.RequiredTemperatureGain < AnalogueSensorCombinations.CoarseTemperatureGainFactor(3)) Then
        Configuration.SelectedCoarseTemperatureGainIndex = 3
    ElseIf (Configuration.RequiredTemperatureGain < AnalogueSensorCombinations.CoarseTemperatureGainFactor(2)) Then
        Configuration.SelectedCoarseTemperatureGainIndex = 2
    ElseIf (Configuration.RequiredTemperatureGain < AnalogueSensorCombinations.CoarseTemperatureGainFactor(1)) Then
        Configuration.SelectedCoarseTemperatureGainIndex = 1
    Else
        ' no point in checking that 0 is good enough, just try it and hope for the best
        Configuration.SelectedCoarseTemperatureGainIndex = 0
    End If
    
    Configuration.TemperatureGainRegister = MAX_ANALOGUE_REGISTER_VALUE * Configuration.RequiredTemperatureGain / AnalogueSensorCombinations.CoarseTemperatureGainFactor(Configuration.SelectedCoarseTemperatureGainIndex)
End Sub
        
        
Private Sub SetInitialOffsetTarget(ByRef Configuration As AnalogueCalibrationConfigurationType, ByVal OutputNumber As Long)
        
    If Not IsTemperatureOutputType(Configuration.AnalogueOutputMode(OutputNumber)) Then
        Select Case Configuration.AnalogueOutputMode(OutputNumber)
        Case AOM_PressureAbsoluteVoltage, AOM_TemperatureAbsoluteVoltage
            ' Just go straight for final target
            Configuration.InitialOffsetTarget(OutputNumber) = Configuration.TargetOffset(OutputNumber)
        Case AOM_PressureCurrent
            ' Just go straight for final target
            Configuration.InitialOffsetTarget(OutputNumber) = Configuration.TargetOffset(OutputNumber)
        Case AOM_PressureRatiometricVoltage, AOM_TemperatureRatiometricVoltage
            ' Just go straight for final target
            Configuration.InitialOffsetTarget(OutputNumber) = Configuration.TargetOffset(OutputNumber)
        End Select
    
    End If

End Sub
        
        
Private Sub CalculateRequiredAsicGain(ByRef Configuration As AnalogueCalibrationConfigurationType, ByVal OutputNumber As Long)
        
    If Configuration.AppliedMaxInput(OutputNumber) <> 0 Then
        Configuration.RequiredAsicGain(OutputNumber) = (Configuration.TargetSpan(OutputNumber) * Configuration.SensorRatedPressure * 1000) / (AnalogueSensorCombinations.SensorVdd * Configuration.SensorMilliVPerV * Configuration.AppliedMaxInput(OutputNumber))
    End If

End Sub
        
Private Sub CalculateInitialOffsetMeasTol(Configuration As AnalogueCalibrationConfigurationType, ByVal OutputNumber As Long)
    Configuration.InitialOffsetMeasTol(OutputNumber) = Configuration.TargetSpan(OutputNumber) / 20
End Sub
        
        
Private Sub SetModeBitsAndSaturationLimits(ByRef Configuration As AnalogueCalibrationConfigurationType, ByVal OutputNumber As Long)
    Configuration.EnablePermanentLock = True

    Select Case Configuration.AnalogueOutputMode(OutputNumber)
    Case AOM_PressureAbsoluteVoltage, AOM_TemperatureAbsoluteVoltage
        Configuration.SelectRatiometric(OutputNumber) = False
        Configuration.SelectIMode(OutputNumber) = False
        Configuration.HighSaturatedOutput(OutputNumber) = RigConfig.Asic_Absolute_Measure_Vdd - 4  ' bodge
        Configuration.LowSaturatedOutput(OutputNumber) = 0.5
    Case AOM_PressureRatiometricVoltage, AOM_TemperatureRatiometricVoltage
        Configuration.SelectRatiometric(OutputNumber) = True
        Configuration.SelectIMode(OutputNumber) = False
        Configuration.HighSaturatedOutput(OutputNumber) = RigConfig.Asic_Ratiometric_Measure_Vdd - 0.75
        Configuration.LowSaturatedOutput(OutputNumber) = 0.5
    Case AOM_PressureCurrent
        Configuration.SelectRatiometric(OutputNumber) = False
        Configuration.SelectIMode(OutputNumber) = True
        Configuration.HighSaturatedOutput(OutputNumber) = 0.021
        Configuration.LowSaturatedOutput(OutputNumber) = 0.0039
    Case AOM_PressureSwitch
        If Configuration.AnalogueOutputMode(0) = AOM_PressureRatiometricVoltage Then
            Configuration.SelectRatiometric(OutputNumber) = True
            Configuration.SelectIMode(OutputNumber) = False
            Configuration.HighSaturatedOutput(OutputNumber) = 0
            Configuration.LowSaturatedOutput(OutputNumber) = 0
        Else
            Configuration.SelectRatiometric(OutputNumber) = False
            Configuration.SelectIMode(OutputNumber) = False
            Configuration.HighSaturatedOutput(OutputNumber) = 0
            Configuration.LowSaturatedOutput(OutputNumber) = 0
        End If
    Case Else
        ' unused output
        Configuration.SelectRatiometric(OutputNumber) = False
        Configuration.SelectIMode(OutputNumber) = False
        Configuration.HighSaturatedOutput(OutputNumber) = 0
        Configuration.LowSaturatedOutput(OutputNumber) = 0
    End Select

    Configuration.MidRangeOutput(OutputNumber) = (Configuration.HighSaturatedOutput(OutputNumber) + Configuration.LowSaturatedOutput(OutputNumber)) / 2
End Sub
        
        

' This function takes an analogue MCS part number, and extracts all the information from it.
' Some is encoded directly, some is derived from a lookup table based on the MCS number
Public Function ExtractAnalogueCalibrationConfigurationFromMCSPartNumber(ByRef Configuration As AnalogueCalibrationConfigurationType, ByVal MCSPartNumber As String, ByVal Side As RigSide) As Boolean
    Dim i As Long
    
    ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True
    
    If RigConfig.ProofOnly Then
    
        If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
            ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetProofPressureFromAnaloguePartNumber(MCSPartNumber, Side, Configuration.ProofPressure)
        End If
    Else
            
        If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
            ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetIsThermallyCompensatedFromMCSPartNumber(MCSPartNumber, Configuration.IsThermallyCompensated)
        End If
        
        ' Note that the IsHeavyDuty flag must be set before working out proof pressure
        If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
            ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetIsHeavyDutyFromMCSPartNumber(MCSPartNumber, Configuration.IsHeavyDuty)
        End If
        
        If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
            ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetSensorPressureRatingFromAnaloguePartNumber(MCSPartNumber, Configuration.SensorRatedPressure)
        End If
        
        If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
            ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetProofPressureFromAnaloguePartNumber(MCSPartNumber, Side, Configuration.ProofPressure)
        End If
        
        If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
            ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetMilliVperVFromAnaloguePartNumber(MCSPartNumber, Configuration.SensorMilliVPerV)
        End If
        
        If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
            ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetNonLinearityFromAnaloguePartNumber(MCSPartNumber, Configuration.TypicalNonLinearity)
        End If
        
        If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
            ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetAnalogueOutputModeFromMCSPartNumber(MCSPartNumber, 0, Configuration.AnalogueOutputMode(0))
        End If
        
        If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
            ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetAnalogueLoadFromAnaloguePartNumber(MCSPartNumber, Configuration.LoadType, Configuration.ResistorCombinationCode)
        End If
        
        If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
            ' note, for second output channel, do not set error if unable to read mode
            ' since no-output is a valid configuration
            GetAnalogueOutputModeFromMCSPartNumber MCSPartNumber, 1, Configuration.AnalogueOutputMode(1)
        End If
         
        If (Configuration.AnalogueOutputMode(1) = AOM_None) Then
            Configuration.NumOutputs = 1
        Else
            Configuration.NumOutputs = 2
        End If
        
        If (Configuration.NumOutputs = 2) Then
            If (IsTemperatureOutputType(Configuration.AnalogueOutputMode(1))) Then
                ' extract temperature coefficients for temperature calibration
                ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetDeltaSFromMCSPartNumber(MCSPartNumber, Configuration.DeltaS)
                'remove next line for deltas look up table
                'Configuration.DeltaS = DeltaSEntry.GetDeltaS(Configuration.DeltaS)
            End If
        End If
        
        For i = 0 To Configuration.NumOutputs - 1
      
            If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
                ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetTargetOffsetFromMCSPartNumber(MCSPartNumber, i, Configuration.TargetOffset(i))
            End If
            
            If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
                SelectNegativeVGen Configuration, i
            End If
            
            If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
                ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetTargetMaxOutputFromMCSPartNumber(MCSPartNumber, i, Configuration.TargetMaxOutput(i))
    
            End If
            
            
            If (IsTemperatureOutputType(Configuration.AnalogueOutputMode(i))) Then
        
                If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
                    ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetAppliedMinTemperatureFromMCSPartNumber(MCSPartNumber, Configuration.AppliedTemperatureOffset)
                End If
            
                Configuration.TargetSpan(i) = Configuration.TargetMaxOutput(i) - Configuration.TargetOffset(i)
            
                If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
                    Dim MaxTemperature As Double
                    ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetAppliedMaxTemperatureFromMCSPartNumber(MCSPartNumber, MaxTemperature)
                    Configuration.AppliedTemperatureSpan = MaxTemperature - Configuration.AppliedTemperatureOffset
                End If
            
            
                If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
                    SelectTemperatureGainRegister Configuration
                End If
            
            
            Else
                If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
                    ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetAppliedMinPressureFromMCSPartNumber(MCSPartNumber, i, Configuration.AppliedMinInput(i))
                End If
            
                If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
                    Dim MaxPressure As Double
                    ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetAppliedMaxPressureFromMCSPartNumber(MCSPartNumber, i, MaxPressure)
                    Configuration.AppliedMaxInput(i) = MaxPressure
                End If
            
            End If
    
            If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
                ' make any changes necessary for negative offset pressures
                CompensteForNegativePressure Configuration, i
                Configuration.TargetSpan(i) = Configuration.TargetMaxOutput(i) - Configuration.TargetOffset(i)
                CalculateInitialOffsetMeasTol Configuration, i
            End If
        
            
            If ((AnalogueSensorCombinations.SensorVdd = 0) Or (Configuration.SensorMilliVPerV = 0) Or (AnalogueSensorCombinations.FeGain = 0)) Then
                ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = False
            End If
            
            If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
                ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = GetMaxProofShiftPercentsFromAnaloguePartNumber(MCSPartNumber, Configuration.MaxPositiveProofShiftPercent, Configuration.MaxNegativeProofShiftPercent)
            End If
            
            If (ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True) Then
                CalculateRequiredAsicGain Configuration, i
            End If
            
            If ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True Then
                SetInitialOffsetTarget Configuration, i
            End If
            
            If ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True Then
                SetModeBitsAndSaturationLimits Configuration, i
            End If
            
            
            
            ' make any changes necessary for unattainably high max pressures
            If ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = True Then
                ExtractAnalogueCalibrationConfigurationFromMCSPartNumber = CompensteForMaxPrecisePressure(Configuration, Side, i)
            End If
            
        Next
        
        If (IsTemperatureOutputType(Configuration.AnalogueOutputMode(TEMPERATURE_OUTPUT))) Then
            Configuration.EnableTempCh = True
        Else
            Configuration.EnableTempCh = False
        End If
        
        If Configuration.AnalogueOutputMode(DIGITAL_SWITCH_OUTPUT) = AOM_PressureSwitch Then
            RigSideStatus(Side).AnalogueCalibrationConfiguration.DigitalSwitch = True
        Else
            RigSideStatus(Side).AnalogueCalibrationConfiguration.DigitalSwitch = False
        End If
        
        Configuration.AppliedMidInput = (Configuration.AppliedMinInput(LINEARITY_CHECK_OUTPUT) + Configuration.AppliedMaxInput(LINEARITY_CHECK_OUTPUT)) / 2
        Configuration.TargetMidOutput = (Configuration.TargetOffset(LINEARITY_CHECK_OUTPUT) + Configuration.TargetMaxOutput(LINEARITY_CHECK_OUTPUT)) / 2
     
        If Not RigConfig.SkipLinearity Then
                
            SetNonLinearitySwitchPoints Configuration
                    
        End If
           
        
        ' Now work out what power supply the transducer needs to be calibrated
        CalculateRatiometricPowerSupplyDrop Configuration
        
    End If
End Function

Public Sub GetAnalogueOutputModeFromScreenItems(ByVal OutputNumber As Integer, ByRef OutputMode As AnalogueOutputModeEnum, ByVal Side As RigSide)
    
    
    Select Case OutputNumber
    Case 0
        Select Case MainForm.Output1TypeCombo(Side).ListIndex
        Case 0
            OutputMode = AOM_PressureAbsoluteVoltage
        Case 1
            OutputMode = AOM_PressureRatiometricVoltage
        Case 2
            OutputMode = AOM_PressureCurrent
        End Select
    Case 1
        Select Case MainForm.Output2TypeCombo(Side).ListIndex
        Case 0
            OutputMode = AOM_None
        Case 1
            If MainForm.Output1TypeCombo(Side).ListIndex = 0 Then
                OutputMode = AOM_TemperatureAbsoluteVoltage
            Else
                OutputMode = AOM_TemperatureRatiometricVoltage
            End If
        Case 2
            OutputMode = AOM_PressureSwitch
        End Select
    End Select
        
End Sub

Public Function ExtractNumberOfPartsFromScreenItems(ByRef NumParts As Long, ByVal Side As RigSide) As Boolean
    NumParts = Val(MainForm.NumPartsEntry(Side))
    ExtractNumberOfPartsFromScreenItems = True
End Function

Public Function ExtractAnalogueCalibrationConfigurationFromScreenItems(ByRef Configuration As AnalogueCalibrationConfigurationType, ByVal Side As RigSide) As Boolean
    Dim i As Long
    Dim PressureRatingIndex As Long
    
    
    If IsThermalCompensationCalAllowed(Side) Then
        If MainForm.ThermalCompCheck(Side).Value = vbChecked Then
            Configuration.IsThermallyCompensated = True
        Else
            Configuration.IsThermallyCompensated = False
        End If
    Else
        Configuration.IsThermallyCompensated = False
    End If
    
    If MainForm.HeavyDutyCheck(Side).Value = vbChecked Then
        Configuration.IsHeavyDuty = True
    Else
        Configuration.IsHeavyDuty = False
    End If
    
    
    ExtractAnalogueCalibrationConfigurationFromScreenItems = True
    
    
    PressureRatingIndex = MainForm.ElementCombo(Side).ListIndex
    
    Configuration.SensorRatedPressure = AnalogueSensorCombinations.PressureRating(PressureRatingIndex).BarPressure
    
    If Configuration.IsHeavyDuty Then
        GetHeavyDutyProofPressureFromPressureRatingIndex PressureRatingIndex, Side, Configuration.ProofPressure
    Else
        GetStandardProofPressureFromPressureRatingIndex PressureRatingIndex, Side, Configuration.ProofPressure
    End If
    GetMilliVPerVFromPressureRatingIndex PressureRatingIndex, Configuration.SensorMilliVPerV
    GetNonLinearityFromPressureRatingIndex PressureRatingIndex, Configuration.TypicalNonLinearity
    
    
    GetAnalogueOutputModeFromScreenItems 0, Configuration.AnalogueOutputMode(0), Side
    GetAnalogueOutputModeFromScreenItems 1, Configuration.AnalogueOutputMode(1), Side
    
    If (Configuration.AnalogueOutputMode(1) = AOM_None) Then
        Configuration.NumOutputs = 1
    Else
        Configuration.NumOutputs = 2
    End If
    
    Dim PullUpDownCode As Long
    
    PullUpDownCode = MainForm.GetResistorCombinationCodeFromLoadCombo(Side)
    GetAnalogueLoadFromMCSSubCode PullUpDownCode, Configuration.LoadType, Configuration.ResistorCombinationCode
    
    If (Configuration.NumOutputs = 2) Then
        If (IsTemperatureOutputType(Configuration.AnalogueOutputMode(1))) Then
            ' extract temperature coefficients for temperature calibration
            GetDeltaSFromFromPressureRatingIndex PressureRatingIndex, Configuration.DeltaS
            'remove next line for deltas look up table
            'Configuration.DeltaS = DeltaSEntry.GetDeltaS(Configuration.DeltaS)
        End If
    End If
    
    For i = 0 To Configuration.NumOutputs - 1

        Configuration.TargetOffset(i) = MainForm.GetTargetOffsetOutputFromScreenItems(Side, i)
        
        SelectNegativeVGen Configuration, i
        
        Configuration.TargetMaxOutput(i) = MainForm.GetTargetMaxOutputFromScreenItems(Side, i)
       
         
        If (IsTemperatureOutputType(Configuration.AnalogueOutputMode(i))) Then
            If (ExtractAnalogueCalibrationConfigurationFromScreenItems = True) Then
                Dim MaxTemperature As Double
                
                ' note that although this is actually a temperature, the same pressure entry boxes is used
                Configuration.AppliedTemperatureOffset = MainForm.GetAppliedMinPressureFromScreenItems(Side, 1)
        
                MaxTemperature = MainForm.GetAppliedMaxPressureFromScreenItems(Side, 1)
                Configuration.AppliedTemperatureSpan = MaxTemperature - Configuration.AppliedTemperatureOffset
            End If
        Else
        If (ExtractAnalogueCalibrationConfigurationFromScreenItems = True) Then
                
            If RigSideStatus(Side).AnalogueCalibrationConfiguration.AnalogueOutputMode(i) = AOM_PressureSwitch Then
    
                Configuration.AppliedMinInput(1) = 0
                        
            Else
                
                Configuration.AppliedMinInput(i) = MainForm.GetAppliedMinPressureFromScreenItems(Side, i)
                
            End If
            
            
                Dim MaxPressure As Double
                MaxPressure = MainForm.GetAppliedMaxPressureFromScreenItems(Side, i)
                
                Configuration.AppliedMaxInput(i) = MaxPressure
                                                 
            
                If MainForm.InputUnits1Combo(Side).ListIndex = 1 Then
                    ' text was entered as PSI, need to convert to bar
                    Configuration.AppliedMinInput(i) = Configuration.AppliedMinInput(i) / BAR_TO_PSI
                    Configuration.AppliedMaxInput(i) = Configuration.AppliedMaxInput(i) / BAR_TO_PSI
                End If
                
                
            End If
        End If
        
        If (ExtractAnalogueCalibrationConfigurationFromScreenItems = True) Then
            ' make any changes necessary for negative offset pressures
            CompensteForNegativePressure Configuration, i
            Configuration.TargetSpan(i) = Configuration.TargetMaxOutput(i) - Configuration.TargetOffset(i)
            CalculateInitialOffsetMeasTol Configuration, i
        End If
        
        If (ExtractAnalogueCalibrationConfigurationFromScreenItems = True) Then
            If (IsTemperatureOutputType(Configuration.AnalogueOutputMode(i))) Then
                SelectTemperatureGainRegister Configuration
            End If
        End If
        
        If ((AnalogueSensorCombinations.SensorVdd = 0) Or (Configuration.SensorMilliVPerV = 0) Or (AnalogueSensorCombinations.FeGain = 0)) Then
            ExtractAnalogueCalibrationConfigurationFromScreenItems = False
        End If
   
    
        If (ExtractAnalogueCalibrationConfigurationFromScreenItems = True) Then
            GetMaxProofShiftsFromPressureRatingIndex PressureRatingIndex, Configuration.MaxPositiveProofShiftPercent, Configuration.MaxNegativeProofShiftPercent
        End If
        
        
        If (ExtractAnalogueCalibrationConfigurationFromScreenItems = True) Then
            CalculateRequiredAsicGain Configuration, i
        End If
   
        If ExtractAnalogueCalibrationConfigurationFromScreenItems = True Then
            SetInitialOffsetTarget Configuration, i
            SetModeBitsAndSaturationLimits Configuration, i
        End If
        

      
        ' make any changes necessary for unattainably high max pressures
        ExtractAnalogueCalibrationConfigurationFromScreenItems = CompensteForMaxPrecisePressure(Configuration, Side, i)
      
    Next
    
    If (IsTemperatureOutputType(Configuration.AnalogueOutputMode(TEMPERATURE_OUTPUT))) Then
        Configuration.EnableTempCh = True
    Else
        Configuration.EnableTempCh = False
    End If
      
    If Configuration.AnalogueOutputMode(DIGITAL_SWITCH_OUTPUT) = AOM_PressureSwitch Then
        RigSideStatus(Side).AnalogueCalibrationConfiguration.DigitalSwitch = True
    Else
        RigSideStatus(Side).AnalogueCalibrationConfiguration.DigitalSwitch = False
    End If
    
    
    Configuration.AppliedMidInput = (Configuration.AppliedMinInput(LINEARITY_CHECK_OUTPUT) + Configuration.AppliedMaxInput(LINEARITY_CHECK_OUTPUT)) / 2
    Configuration.TargetMidOutput = (Configuration.TargetOffset(LINEARITY_CHECK_OUTPUT) + Configuration.TargetMaxOutput(LINEARITY_CHECK_OUTPUT)) / 2
    
    If Not RigConfig.SkipLinearity Then
            
        SetNonLinearitySwitchPoints Configuration
                
    End If
        
    
    ' Now work out what power supply the transducer needs to be calibrated
    CalculateRatiometricPowerSupplyDrop Configuration

End Function


Public Function IsConfigurationSensible(ByRef Configuration As AnalogueCalibrationConfigurationType, ByVal Side As RigSide) As Boolean
    IsConfigurationSensible = True
    
    Dim i As Integer
    
    If Configuration.AppliedMaxInput(0) <= Configuration.AppliedMinInput(0) Then
        IsConfigurationSensible = False
    End If
    
    If Configuration.AppliedMaxInput(0) < Configuration.SensorRatedPressure * 0.8 Then
        ' below 80% of sensor rate pressure
        IsConfigurationSensible = False
    End If
    
    If Configuration.AppliedMaxInput(0) > Configuration.SensorRatedPressure * 1.2 Then
        ' above 120% of sensor rate pressure
        IsConfigurationSensible = False
    End If
    
    
    For i = 0 To Configuration.NumOutputs - 1
        If Configuration.TargetMaxOutput(i) <= Configuration.TargetOffset(i) Then
            IsConfigurationSensible = False
        End If
    Next i
    
    
    If Configuration.SelectIMode(0) Then
        If Configuration.TargetOffset(0) > 0.005 Then
            IsConfigurationSensible = False
        End If
        
        If Configuration.TargetMaxOutput(0) > 0.022 Then
            IsConfigurationSensible = False
        End If
        
    Else
        If Configuration.TargetOffset(0) > 5 Then
            IsConfigurationSensible = False
        End If
        If Configuration.SelectRatiometric(0) Then
            If Configuration.TargetMaxOutput(0) > 5 Then
                IsConfigurationSensible = False
            End If
        Else
            If Configuration.TargetMaxOutput(0) > 12 Then
                IsConfigurationSensible = False
            End If
        End If
    End If
    
    
    
    If Configuration.NumOutputs = 2 Then
        If Configuration.AppliedTemperatureSpan <= 0 Then
            IsConfigurationSensible = False
        End If
    End If
    


End Function

