VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileConcat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A class implementation of the code posted by Micheal W. on 7/20/2002 12:28:11 PM
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=37109&lngWId=1

Private lngFileNumber As Long
Private strPath As String

Private Sub Class_Initialize()
    lngFileNumber = FreeFile(0)
    strPath = FormatPath(App.Path) & "temp.tmp"
    Open strPath For Append As #lngFileNumber
End Sub

Private Sub Class_Terminate()
    Close #lngFileNumber
    SetAttr strPath, vbNormal
    Kill strPath
End Sub

Friend Property Get Value() As String
    Close #lngFileNumber
    Value = Space$(FileLen(strPath))
    Open strPath For Binary Access Read As #lngFileNumber
    Get #lngFileNumber, , Value
End Property

Private Function FormatPath(strPath As String) As String
  If Right$(strPath, 1) <> "\" Then
    FormatPath = strPath & "\"
  Else
    FormatPath = strPath
  End If
End Function

Friend Sub Append(strAppend As String)
    Print #lngFileNumber, ; strAppend;
End Sub

Friend Sub Flush()
    ' Empty the old buffer
    Close #lngFileNumber
    SetAttr strPath, vbNormal
    Kill strPath
    
    ' And open a new one
    lngFileNumber = FreeFile(0)
    strPath = FormatPath(App.Path) & "temp.tmp"
    Open strPath For Append As #lngFileNumber
End Sub
