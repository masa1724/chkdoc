Attribute VB_Name = "CheckUtils"
Option Explicit

Public Function CreateRegExp(pattern As String) As Object
    Dim re As Object
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.pattern = pattern
    
    Set CreateRegExp = re
End Function

Public Function SafeGetLine(allLines As collection, lineNo As Long) As String
    If allLines.Count = 0 Then
        SafeGetLine = ""
        Exit Function
    End If
    
    If lineNo < 1 Or allLines.Count < lineNo Then
        SafeGetLine = ""
        Exit Function
    End If
    
    SafeGetLine = allLines(lineNo)
End Function

Public Function CheckRange(allLines As collection, regexp As Object, baseLineNo As Long, offsets As collection) As String
    Dim offset As Variant
        
    For Each offset In offsets
        Dim lineNo As Long
        lineNo = baseLineNo + offset
    
        If regexp.test(SafeGetLine(allLines, lineNo)) Then
            CheckRange = lineNo
            Exit Function
        End If
    Next
    
    CheckRange = -1
End Function
