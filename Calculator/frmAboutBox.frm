VERSION 5.00
Begin VB.Form frmAboutBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Tim's Scientific Calculator"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   Icon            =   "frmAboutBox.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   3765
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer ScrollingText 
      Interval        =   100
      Left            =   4320
      Top             =   2400
   End
   Begin VB.CommandButton cbClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.PictureBox pbScrollTextBG 
      BackColor       =   &H80000001&
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1755
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   240
      Width           =   3255
      Begin VB.Label lblScrollingText 
         BackColor       =   &H00000000&
         Caption         =   "Simple Scientific Calculator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2000
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "© 2005 by Tim Alvord.       All Rights Reserved"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1028
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Image imgAboutBox 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   480
      Picture         =   "frmAboutBox.frx":0442
      Top             =   2640
      Width           =   510
   End
End
Attribute VB_Name = "frmAboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call GradientForm(Me, vbBlack, vbBlue)
    lblScrollingText.Caption = "Simple Scientific Calculator v1.02" & vbCrLf & vbCrLf & _
                               "* Roman, Hex, Decimal, Octal and Binary" & vbCrLf & _
                               "* Degs, Rads and Grads." & vbCrLf & _
                               "* Standard Trig functions." & vbCrLf & _
                               "* Factorial, Sqr, Inverse, Square, Cube, X^Y."
    lblScrollingText.Top = pbScrollTextBG.Top + pbScrollTextBG.Height
End Sub

Private Sub ScrollingText_Timer()
    If lblScrollingText.Top + lblScrollingText.Height <= pbScrollTextBG.Top Then
        lblScrollingText.Top = pbScrollTextBG.Top + pbScrollTextBG.Height
    End If
    lblScrollingText.Top = lblScrollingText.Top - 15
End Sub

Private Sub cbClose_Click()
    Unload Me
End Sub
