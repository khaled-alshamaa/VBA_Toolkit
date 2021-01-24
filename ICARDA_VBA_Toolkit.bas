Attribute VB_Name = "ICARDA_Toolkit"

' Name:      ICARDA-VBA-Toolkit-v2.bas
' Copyright: 2019-2021, ICARDA
' Purpose:   Set of VBA utility functions
' Author:    Khaled Al-Shamaa <k.el-shamaa@cgiar.org>
' Version:   2.0
' Revision:  25 Jan 2021 - add DD2OLC, OLC2DD, and VOLC functions
'            12 Jan 2019 - initial version
' License:   GPLv3

' Enable this flag when running in OpenOffice/Libre Office.
'Option VBASupport 1

' Encode a location coordinates (latitude and longitude in WGS84) into Open Location Code
' Ref: https://github.com/google/open-location-code/blob/master/docs/specification.md
Public Function DD2OLC(latitude As Double, longitude As Double, Optional codeLength As Integer = 10) As Variant
    Dim x, y As Integer, validChars As String
    
    codeLength = codeLength / 2
    validChars = "23456789CFGHJMPQRVWX"
    
    latitude = latitude + 90
    longitude = longitude + 180
    
    latitude = Round(latitude * 20 ^ (codeLength - 2), 0)
    longitude = Round(longitude * 20 ^ (codeLength - 2), 0)
    
    For i = 1 To codeLength
        x = longitude Mod 20
        y = latitude Mod 20
        
        longitude = Fix(longitude / 20)
        latitude = Fix(latitude / 20)
        
        olc = Mid(validChars, y + 1, 1) & Mid(validChars, x + 1, 1) & olc
        
        If i = 1 Then olc = "+" & olc
    Next i
    
    DD2OLC = olc
End Function

' Decode an Open Location Code into its location coordinates (WGS84)
' Ref: https://github.com/google/open-location-code/blob/master/docs/specification.md
Public Function OLC2DD(olc As String, Optional coordinates As Integer = 0, Optional codeLength As Integer = 10) As Variant
    Dim latitude, longitude As Double, validChars As String
    
    If VOLC(olc, codeLength) = True Then
        codeLength = codeLength / 2
        validChars = "23456789CFGHJMPQRVWX"
            
        olc = UCase(Replace(olc, "+", ""))
        
        For i = 1 To codeLength
            latitude = latitude + (InStr(validChars, Mid(olc, 2 * i - 1, 1)) - 1) * 20 ^ (2 - i)
            longitude = longitude + (InStr(validChars, Mid(olc, 2 * i, 1)) - 1) * 20 ^ (2 - i)
        Next i
        
        latitude = latitude - 90
        longitude = longitude - 180
        
        OLC2DD = IIf(coordinates = 1, latitude, IIf(coordinates = 2, longitude, latitude & ", " & longitude))
    Else
        OLC2DD = "Invalid Code!"
    End If
End Function

' Determine if an Open Location Code is valid
Public Function VOLC(olc As String, Optional codeLength As Integer = 10) As Variant
    Dim regEx As Object
    Set regEx = CreateObject("vbscript.regexp")
    
    regEx.Pattern = "[^2-9CFGHJMPQRVWX]+"
    
    If Len(olc) <> codeLength + 1 Then
        IsValid = False
    ElseIf Mid(olc, codeLength - 1, 1) <> "+" Then
        IsValid = False
    ElseIf regEx.Test(UCase(Replace(olc, "+", ""))) Then
        IsValid = False
    Else
        IsValid = True
    End If
    
    VOLC = IsValid
End Function


' Generate the Code 128 Barcode, including the checksum.
Public Function Barcode(myLabel As String) As Variant
    Dim ch As String, n As Long, sum As Long, checksum As Integer
    sum = 104
    
    For n = 1 To Len(myLabel)
        ch = Mid(myLabel, n, 1)
        sum = sum + n * (Asc(ch) - 32)
    Next n
    
    checksum = sum Mod 103

    ' Map checksum to an ASCII code. This conversion takes into account the
    ' particular mapping of the font being used
    ' this VBA function is working well for the font "Libre Barcode 128":
    ' https://fonts.google.com/specimen/Libre+Barcode+128
    If checksum = 0 Then
        checksum = 212
    ElseIf checksum <= 94 Then
        checksum = checksum + 32
    Else
        checksum = checksum + 105
    End If
    
    Barcode = "Ì" & myLabel & Chr(checksum) & "Î"
End Function

' Convert Degrees Minutes Seconds (DMS) coordinates to Decimal Degrees (DD)
Public Function DMS2DD(degStr As String) As Variant
    Dim regEx As Object
    Set regEx = CreateObject("vbscript.regexp")

    degStr = Replace(degStr, " ", "")
    degStr = Replace(degStr, "''", """")
    
    'You degree symbol by click on Alt+0176 from the numkey
    regEx.Pattern = "(([0-9\.]+)[^'""0-9\.])?(([0-9\.]+)')?(([0-9\.]+)"")?([WwSs])?"

    If regEx.Test(degStr) Then
        Set regMatchs = regEx.Execute(degStr)
        
        x = regMatchs(0).SubMatches(1)
        y = regMatchs(0).SubMatches(3)
        Z = regMatchs(0).SubMatches(5)
        
        If (Len(regMatchs(0).SubMatches(6)) = 1) Then
            D = -1
        Else
            D = 1
        End If
    Else
        MsgBox ("Oops!")
    End If
    
    DMS2DD = D * (x + (y / 60) + (Z / 3600))
End Function

' Convert Decimal Degrees (DD) coordinates to Degrees Minutes Seconds (DMS)
Public Function DD2DMS(decStr As String) As Variant
    Degrees = Int(decStr)
    Minutes = Int((decStr - Degrees) * 60)
    Seconds = Round((((decStr - Degrees) * 60) - Minutes) * 60, 4)
    
    outStr = ""
    If (Degrees > 0) Then outStr = Degrees & "°"
    
    If (Minutes >= 10) Then
        outStr = outStr & Minutes & "'"
    ElseIf (Minutes > 0) Then
        outStr = outStr & "0" & Minutes & "'"
    End If
    
    If (Seconds >= 10) Then
        outStr = outStr & Seconds & """"
    ElseIf (Seconds > 0) Then
        outStr = outStr & "0" & Seconds & """"
    End If
    
    DD2DMS = outStr
End Function
