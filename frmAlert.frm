VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3525
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAlert 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2880
      Top             =   0
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   2775
      TabIndex        =   0
      Top             =   0
      Width           =   2840
      Begin VB.Label lblAlert 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Left            =   100
         Picture         =   "frmAlert.frx":0000
         Top             =   100
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2460
         Picture         =   "frmAlert.frx":058A
         Top             =   100
         Width           =   240
      End
      Begin VB.Label lbTit 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   123
         Width           =   525
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   2160
         Picture         =   "frmAlert.frx":0B14
         Top             =   1080
         Width           =   480
      End
   End
   Begin VB.Image imgUser 
      Height          =   240
      Left            =   3240
      Picture         =   "frmAlert.frx":17DE
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgInfo 
      Height          =   240
      Left            =   2880
      Picture         =   "frmAlert.frx":1D68
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgError 
      Height          =   240
      Left            =   2880
      Picture         =   "frmAlert.frx":22F2
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgExcla 
      Height          =   240
      Left            =   2880
      Picture         =   "frmAlert.frx":287C
      Top             =   1120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   2880
      Picture         =   "frmAlert.frx":2E06
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   3120
      Picture         =   "frmAlert.frx":3390
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' API Declarations
Private Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)

' Constants
Const SM_CXFULLSCREEN = 16   ' Width of window client area
Const SM_CYFULLSCREEN = 17   ' Height of window client area

' Declarations
Private fx As Long
Private fY As Long
Private lngScaleX As Long
Private lngScaleY As Long

Private AlertIndex As Long

Private cGrad As New CGradient

Private Sub Form_Activate()
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Image1_Click()
    If AlertCount = AlertIndex Then AlertCount = 0
    Me.Visible = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Image1.Picture = Image2.Picture
End Sub

Private Sub lblAlert_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Show as hyperlink
    If lblAlert.FontUnderline = False Then
        lblAlert.FontUnderline = True
    End If
End Sub

Private Sub picBackground_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Show text
    If lblAlert.FontUnderline = True Then
        lblAlert.FontUnderline = False
    End If

    Image1.Picture = Image3.Picture

End Sub

Public Sub Display(MessageText As String, sTit As String, sIco As Long, Duration As Integer, Sound As Long, cF As Long, c1 As Long, c2 As Long)

    Dim wFlags As Long, x As Long

    ' Increase the alert count
    AlertCount = AlertCount + 1
    If AlertCount >= 5 Then AlertCount = 1
    AlertIndex = AlertCount

    ' Set the message
    lbTit.Caption = sTit
    lblAlert.Caption = MessageText
    
    ' set icon
    Select Case sIco
        Case Is = 16
            imgIcon.Picture = imgError.Picture
        Case Is = 48
            imgIcon.Picture = imgExcla.Picture
        Case Is = 10
            imgIcon.Picture = imgUser.Picture
        Case Else
            imgIcon.Picture = imgInfo.Picture
    End Select
    
    ' Set the duration
    tmrAlert.Interval = Duration * 1000
    
    ' Setea el color
    lbTit.ForeColor = cF
    lblAlert.ForeColor = cF
    
    ' gradiente
    With cGrad
        .Angle = 90
        .Color1 = c1
        .Color2 = c2
        .Draw picBackground
    End With
    
    ' Get the system metrics we need
    fx = GetSystemMetrics(SM_CXFULLSCREEN)
    fY = GetSystemMetrics(SM_CYFULLSCREEN)
    lngScaleX = Me.Width - Me.ScaleWidth
    lngScaleY = Me.Height - Me.ScaleHeight
    
    ' Size the form
    Me.Height = 90
    Me.Width = picBackground.Width + lngScaleX
    Me.Left = fx * Screen.TwipsPerPixelX - Me.Width - 200
    Me.Top = (fY * Screen.TwipsPerPixelY) - ((picBackground.Height + lngScaleY) * (AlertCount - 1)) + 200
    Me.Show
    
    ' Play sound
    PlaySound Sound
        
    picBackground.Refresh
    ' Open the alert box
     Call tmOpen
 
End Sub

Private Sub tmOpen()
    Dim curHeight As Long
    Dim newHeight As Long
    curHeight = Me.Height
    Do Until curHeight >= picBackground.Height + lngScaleY
       DoEvents
        newHeight = curHeight + 20 ' 15
        If newHeight > picBackground.Height + lngScaleY Then newHeight = picBackground.Height + lngScaleY
        Me.Height = Me.Height + (newHeight - curHeight)
        Me.Top = Me.Top - (newHeight - curHeight)
        curHeight = Me.Height
   Loop
        tmrAlert.Enabled = True
End Sub

Private Sub tmrAlert_Timer()
    ' Alert was displayed, now close it
    tmrAlert.Enabled = False
    Call tmClose
End Sub

Private Sub tmClose()
Dim curHeight As Long
    curHeight = Me.Height
    Do Until curHeight <= 120
       DoEvents
        Me.Height = curHeight - 10 '0
        Me.Top = Me.Top + 10 '0
        curHeight = Me.Height
    Loop
        ' Close form
        If AlertCount = AlertIndex Then AlertCount = 0
        Unload Me
End Sub

