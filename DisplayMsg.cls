VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisplayMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum SoundType
    sMess = 10
    sAlert = 20
End Enum

Public Enum IconType
    iInfo = 64
    iError = 16
    iExcla = 48
    iUser = 10
End Enum


Public Function DisplayAlert(MessageText As String, Optional sTitle As String = "Atención...", Optional tIcon As IconType = 64, Optional Duration As Integer = 7, Optional Sound As SoundType = 20, Optional cFont As Long = 16777215, Optional cColor1 As Long = 0, Optional cColor2 As Long = 12582912) As String
    On Error GoTo errtrap
    Dim AlertBox As frmAlert
    Set AlertBox = New frmAlert
        AlertBox.Display MessageText, sTitle, tIcon, Duration, Sound, cFont, cColor1, cColor2
        DisplayAlert = vbNullString
    Exit Function
errtrap:
    DisplayAlert = Err.Description
End Function
