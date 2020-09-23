Attribute VB_Name = "bDisplay"
Option Explicit

Public AlertCount As Integer

Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Private m_snd() As Byte
Private Const SND_ASYNC = &H1 ' play asynchronously
Private Const SND_MEMORY = &H4 ' lpszSoundName points to a memory file

Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" _
                                           (lpData As Any, _
                                      ByVal hModule As Long, _
                                      ByVal dwFlags As Long) As Long

Public Function PlaySound(ByVal SndID As Long) As Long
       Const Flags = SND_ASYNC Or SND_MEMORY
       m_snd = LoadResData(SndID, "SOUND")
       PlaySoundData m_snd(0), 0, Flags
End Function
