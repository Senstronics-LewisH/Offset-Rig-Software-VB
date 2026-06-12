Attribute VB_Name = "Barcode"
Option Explicit

' =========================================================================
' Barcode Parsing Module
' 
' Exposes unified parsing subroutines to extract parameters from scanned 
' sensor barcodes (First, Second, and Third).
' Updates the global/module-level variables declared in OffsetCheck.bas.
' =========================================================================

Public Sub ParseFirstBarcode(ByVal Barcode As String)
    ' Clean input and ensure prefix %0 is present
    Barcode = Trim$(Barcode)
    If Left$(Barcode, 2) <> "%0" Then
        Barcode = "%0" & Barcode
    End If

    ' Extract MCS Union Type (characters 10 to 12)
    If Len(Barcode) >= 12 Then
        MCSUnionType = Mid$(Barcode, 10, 3)
    Else
        MCSUnionType = ""
    End If

    ' Extract Board Type (characters 15 to 16)
    If Len(Barcode) >= 16 Then
        BoardType = Mid$(Barcode, 15, 2)
    Else
        BoardType = ""
    End If

    ' Extract Cable Type (characters 15 to 19 or 20)
    If Len(Barcode) >= 20 Then
        CableType1 = Mid$(Barcode, 15, 6)
        If CableType1 = "AT2XTV" Then
            CableType = "AT2YT"
        Else
            CableType = Mid$(Barcode, 15, 5)
        End If
    ElseIf Len(Barcode) >= 19 Then
        CableType = Mid$(Barcode, 15, 5)
        CableType1 = CableType
    Else
        CableType = ""
        CableType1 = ""
    End If

    ' Extract Connector Type (characters 18 to 19)
    If Len(Barcode) >= 19 Then
        ConnectorType = Mid$(Barcode, 18, 2)
    Else
        ConnectorType = ""
    End If

    ' Extract MCS Mating Connector flag (character 20)
    If Len(Barcode) >= 20 Then
        MCSMatingConnector = Mid$(Barcode, 20, 1)
    Else
        MCSMatingConnector = ""
    End If

    ' Extract MCS Product Range (characters 22 to 23)
    If Len(Barcode) >= 23 Then
        MCSProductRange = Mid$(Barcode, 22, 2)
    Else
        MCSProductRange = ""
    End If

    ' Extract Load Value (characters 25 to 27)
    If Len(Barcode) >= 27 Then
        LoadValue = Val(Mid$(Barcode, 25, 3))
    Else
        LoadValue = 0
    End If

    ' Extract End Delimiter (characters 29 to 30)
    If Len(Barcode) >= 30 Then
        MCSEndDelimiter = Mid$(Barcode, 29, 2)
    Else
        MCSEndDelimiter = ""
    End If

    ' Determine O-Ring Required
    If Len(Barcode) >= 14 And Mid$(Barcode, 14, 1) = "/" Then
        If (Mid$(Barcode, 15, 2) = "") Or (Mid$(Barcode, 15, 1) = "-") Or (Mid$(Barcode, 15, 2) = "--") Or (Len(Barcode) >= 30 And Mid$(Barcode, 29, 2) = "TT") Then
            ORingRequired = False
        Else
            ORingRequired = True
        End If
    Else
        If Len(Barcode) >= 19 Then
            If (Mid$(Barcode, 18, 2) = "") Or (Mid$(Barcode, 18, 2) = "--") Then
                ORingRequired = False
            Else
                ORingRequired = True
            End If
        Else
            ORingRequired = False
        End If
    End If

    ' Determine Restrictor Required
    If Len(Barcode) >= 13 Then
        If Mid$(Barcode, 13, 1) = "0" Then
            RestrictorReq = False
        Else
            RestrictorReq = True
        End If
    Else
        RestrictorReq = False
    End If
End Sub

Public Sub ParseSecondBarcode(ByVal Barcode As String)
    ' Clean input and ensure prefix %1 is present
    Barcode = Trim$(Barcode)
    If Left$(Barcode, 2) <> "%1" Then
        Barcode = "%1" & Barcode
    End If

    ' Extract Output Type (character 3)
    If Len(Barcode) >= 3 Then
        OutputType = Mid$(Barcode, 3, 1)
    Else
        OutputType = ""
    End If

    ' Extract Offset (characters 4 to 6)
    If Len(Barcode) >= 6 Then
        BarcodeOffset = Val(Mid$(Barcode, 4, 3))
    Else
        BarcodeOffset = 0
    End If

    ' Extract Full Scale (characters 8 to 10 or default 1025)
    If Len(Barcode) >= 10 Then
        If Mid$(Barcode, 8, 1) = "A" Then
            BarcodeFullScale = 1025
        Else
            BarcodeFullScale = Val(Mid$(Barcode, 8, 3))
        End If
    Else
        BarcodeFullScale = 0
    End If

    ' Extract Pressure Units (character 12)
    If Len(Barcode) >= 12 Then
        BarcodeUnits = Mid$(Barcode, 12, 1)
    Else
        BarcodeUnits = ""
    End If

    ' Extract zero and full scale pressure based on units
    If BarcodeUnits = "B" Then
        Units = "bar"
        If Len(Barcode) >= 14 Then
            BarcodeZeroPressure = Val(Mid$(Barcode, 13, 2))
        Else
            BarcodeZeroPressure = 0
        End If
        If Len(Barcode) >= 20 Then
            BarcodeFSPressure = Val(Mid$(Barcode, 17, 4))
        Else
            BarcodeFSPressure = 0
        End If
        ZeroPressureSign = ""
    Else
        Units = "psi"
        If Len(Barcode) >= 14 Then
            BarcodeZeroPressure = Val(Mid$(Barcode, 12, 3))
        Else
            BarcodeZeroPressure = 0
        End If
        If Len(Barcode) >= 20 Then
            BarcodeFSPressure = Val(Mid$(Barcode, 16, 5))
        Else
            BarcodeFSPressure = 0
        End If
        If Len(Barcode) >= 11 Then
            ZeroPressureSign = Mid$(Barcode, 11, 1)
        Else
            ZeroPressureSign = ""
        End If
    End If
End Sub

Public Sub ParseThirdBarcode(ByVal Barcode As String)
    Dim ModeChar As String

    ' Clean input and ensure prefix %2 is present
    Barcode = Trim$(Barcode)
    If Left$(Barcode, 2) <> "%2" Then
        Barcode = "%2" & Barcode
    End If

    ' Check if Vout2 is active (character 3 is T or S)
    If Len(Barcode) >= 3 Then
        ModeChar = Mid$(Barcode, 3, 1)
        Vout2Mode = ModeChar
        If ModeChar = "T" Or ModeChar = "S" Then
            Vout2 = True
        Else
            Vout2 = False
        End If
    Else
        Vout2 = False
        Vout2Mode = ""
        ModeChar = ""
    End If

    ' Extract format check (characters 4 to 10)
    If Len(Barcode) >= 10 Then
        ThirdBarcodeFormatCheck = Mid$(Barcode, 4, 7)
    Else
        ThirdBarcodeFormatCheck = ""
    End If

    If Vout2 Then
        If ModeChar = "T" Then
            ' Extract Temperature parameters
            If Len(Barcode) >= 6 Then
                Offset2 = Val(Mid$(Barcode, 4, 3))
            Else
                Offset2 = 0
            End If
            If Len(Barcode) >= 10 Then
                Fullscale2 = Val(Mid$(Barcode, 8, 3))
            Else
                Fullscale2 = 0
            End If
            If Len(Barcode) >= 14 Then
                VOUT2Spana = Val(Mid$(Barcode, 11, 4))
            Else
                VOUT2Spana = 0
            End If
            If Len(Barcode) >= 20 Then
                VOUT2SpanB = Val(Mid$(Barcode, 17, 4))
            Else
                VOUT2SpanB = 0
            End If
            SWPressure = 0
        ElseIf ModeChar = "S" Then
            ' Extract Switch pressure (using units parsed from second barcode)
            If BarcodeUnits = "B" Then
                If Len(Barcode) >= 20 Then
                    SWPressure = Val(Mid$(Barcode, 17, 4))
                Else
                    SWPressure = 0
                End If
            Else
                If Len(Barcode) >= 20 Then
                    SWPressure = Val(Mid$(Barcode, 16, 5)) / 14.5
                Else
                    SWPressure = 0
                End If
            End If
            Offset2 = 0
            Fullscale2 = 0
            VOUT2Spana = 0
            VOUT2SpanB = 0
        End If
    Else
        Offset2 = 0
        Fullscale2 = 0
        VOUT2Spana = 0
        VOUT2SpanB = 0
        SWPressure = 0
    End If
End Sub
