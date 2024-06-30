VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeCheckResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ソースコードチェック結果
Public branch As String
Public filePath As String
Public lineNo As String
Public errMsg As String
Public lineContents As String
