Attribute VB_Name = "ICARDA_Toolkit"
' Name:      ICARDA-VBA-Toolkit-v1.xlsm
' Copyright: 2019, ICARDA
' Purpose:   Set of VBA utility functions
' Author:    Khaled Al-Shamaa <k.el-shamaa@cgiar.org>
' Version:   1.0
' Revision:  12 Jan 2019 - initial version
' License:   GPLv3

'Generate a Code 128 Barcode including checksum
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
Public Function DEG2DEC(degStr As String) As Variant
    Dim regEx As Object
    Set regEx = CreateObject("vbscript.regexp")

    degStr = Replace(degStr, " ", "")
    degStr = Replace(degStr, "''", """")
    
    'You degree symbol by click on Alt+0176 from the numkey
    regEx.Pattern = "(([0-9\.]+)[^'""0-9\.])?(([0-9\.]+)')?(([0-9\.]+)"")?([WwSs])?"
    If regEx.Test(degStr) Then
        X = regEx.Execute(degStr)(0).SubMatches(1)
        Y = regEx.Execute(degStr)(0).SubMatches(3)
        Z = regEx.Execute(degStr)(0).SubMatches(5)
        
        If (Len(regEx.Execute(degStr)(0).SubMatches(6)) = 1) Then
            D = -1
        Else
            D = 1
        End If
    Else
        MsgBox ("Oops!")
    End If
    
    DEG2DEC = D * (X + (Y / 60) + (Z / 3600))
End Function

' Convert Decimal Degrees (DD) coordinates to Degrees Minutes Seconds (DMS)
Public Function DEC2DEG(decStr As String) As Variant
    Degrees = Int(decStr)
    Minutes = Int((decStr - Degrees) * 60)
    Seconds = Round((((decStr - Degrees) * 60) - Minutes) * 60, 4)
    
    Output = ""
    If (Degrees > 0) Then Output = Degrees & "°"
    
    If (Minutes >= 10) Then
        Output = Output & Minutes & "'"
    ElseIf (Minutes > 0) Then
        Output = Output & "0" & Minutes & "'"
    End If
    
    If (Seconds >= 10) Then
        Output = Output & Seconds & """"
    ElseIf (Seconds > 0) Then
        Output = Output & "0" & Seconds & """"
    End If
    
    DEC2DEG = Output
End Function
