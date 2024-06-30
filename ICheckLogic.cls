VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICheckLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Check(line As String, allLines As collection, lineNo As Long) As Long
End Function

Public Function SkipCheck(branch As String, filePath As String, line As String, allLines As collection, lineNo As Long) As Boolean
End Function

Public Function GetErrMsg() As String
End Function
