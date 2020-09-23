VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test DisplayMess"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simple"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsMess As New DisplayMsg

Private Sub Command1_Click()
    clsMess.DisplayAlert "Este es un ejemplo simple..."
End Sub

Private Sub Command2_Click()
    clsMess.DisplayAlert "Este ejemplo es mas personalizado...", "Otro ejemplo", iError, 3, sMess, vbYellow, vbBlack, vbGreen
End Sub

Private Sub Command3_Click()
    clsMess.DisplayAlert "Ernesto F. Silva" & vbNewLine & "Basado en POPUP MsgBox..." & vbNewLine & "ernestofs7@hotmail.com", "Acerca de...", iUser
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTest = Nothing
End Sub
