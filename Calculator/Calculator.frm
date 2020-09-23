VERSION 5.00
Begin VB.Form Calculator 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ForeColor       =   &H00000000&
   Icon            =   "Calculator.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   48
      Left            =   1560
      Picture         =   "Calculator.frx":0442
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   64
      Top             =   7800
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   51
      Left            =   3720
      Picture         =   "Calculator.frx":0D2E
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   63
      Top             =   7800
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   45
      Left            =   3720
      Picture         =   "Calculator.frx":16B5
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   62
      Top             =   7080
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   39
      Left            =   3720
      Picture         =   "Calculator.frx":203C
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   61
      Top             =   6360
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   33
      Left            =   3720
      Picture         =   "Calculator.frx":29C3
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   60
      Top             =   5640
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   27
      Left            =   3720
      Picture         =   "Calculator.frx":334A
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   59
      Top             =   4920
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   50
      Left            =   3000
      Picture         =   "Calculator.frx":3CD1
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   58
      Top             =   7800
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   44
      Left            =   3000
      Picture         =   "Calculator.frx":4658
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   57
      Top             =   7080
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   38
      Left            =   3000
      Picture         =   "Calculator.frx":4FDF
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   56
      Top             =   6360
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   32
      Left            =   3000
      Picture         =   "Calculator.frx":5966
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   55
      Top             =   5640
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   49
      Left            =   2280
      Picture         =   "Calculator.frx":62ED
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   54
      Top             =   7800
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   43
      Left            =   2280
      Picture         =   "Calculator.frx":6C74
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   53
      Top             =   7080
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   37
      Left            =   2280
      Picture         =   "Calculator.frx":75FB
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   52
      Top             =   6360
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   31
      Left            =   2280
      Picture         =   "Calculator.frx":7F82
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   51
      Top             =   5640
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   42
      Left            =   1560
      Picture         =   "Calculator.frx":8909
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   50
      Top             =   7080
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   36
      Left            =   1560
      Picture         =   "Calculator.frx":9290
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   49
      Top             =   6360
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   30
      Left            =   1560
      Picture         =   "Calculator.frx":9C17
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   48
      Top             =   5640
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   46
      Left            =   120
      Picture         =   "Calculator.frx":A59E
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   47
      Top             =   7800
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   40
      Left            =   120
      Picture         =   "Calculator.frx":AF25
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   46
      Top             =   7080
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   34
      Left            =   120
      Picture         =   "Calculator.frx":B8AC
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   45
      Top             =   6360
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   28
      Left            =   120
      Picture         =   "Calculator.frx":C233
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   44
      Top             =   5640
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   22
      Left            =   120
      Picture         =   "Calculator.frx":CBBA
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   43
      Top             =   4920
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   16
      Left            =   120
      Picture         =   "Calculator.frx":D541
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   42
      Top             =   4200
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   21
      Left            =   3720
      Picture         =   "Calculator.frx":DEC8
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   41
      Top             =   4200
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   15
      Left            =   3720
      Picture         =   "Calculator.frx":E786
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   40
      Top             =   3480
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   47
      Left            =   840
      Picture         =   "Calculator.frx":F044
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   39
      Top             =   7800
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   41
      Left            =   840
      Picture         =   "Calculator.frx":F930
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   38
      Top             =   7080
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   35
      Left            =   840
      Picture         =   "Calculator.frx":1021C
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   37
      Top             =   6360
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   29
      Left            =   840
      Picture         =   "Calculator.frx":10B08
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   36
      Top             =   5640
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   23
      Left            =   840
      Picture         =   "Calculator.frx":113F4
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   35
      Top             =   4920
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   17
      Left            =   840
      Picture         =   "Calculator.frx":11CE0
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   34
      Top             =   4200
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   26
      Left            =   3000
      Picture         =   "Calculator.frx":125CC
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   33
      Top             =   4920
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   25
      Left            =   2280
      Picture         =   "Calculator.frx":12EB8
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   32
      Top             =   4920
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   24
      Left            =   1560
      Picture         =   "Calculator.frx":137A4
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   31
      Top             =   4920
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   20
      Left            =   3000
      Picture         =   "Calculator.frx":14090
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   30
      Top             =   4200
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   19
      Left            =   2280
      Picture         =   "Calculator.frx":1497C
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   29
      Top             =   4200
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   18
      Left            =   1560
      Picture         =   "Calculator.frx":15268
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   28
      Top             =   4200
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   14
      Left            =   3000
      Picture         =   "Calculator.frx":15B54
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   27
      Top             =   3480
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   13
      Left            =   2280
      Picture         =   "Calculator.frx":16440
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   26
      Top             =   3480
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   12
      Left            =   1560
      Picture         =   "Calculator.frx":16D2C
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   25
      Top             =   3480
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   11
      Left            =   3720
      Picture         =   "Calculator.frx":17618
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   24
      Top             =   2760
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   10
      Left            =   3000
      Picture         =   "Calculator.frx":17ED6
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   23
      Top             =   2760
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   9
      Left            =   2280
      Picture         =   "Calculator.frx":1885D
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   22
      Top             =   2760
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   8
      Left            =   1560
      Picture         =   "Calculator.frx":191E4
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   21
      Top             =   2760
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   7
      Left            =   840
      Picture         =   "Calculator.frx":19B6B
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   20
      Top             =   2760
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   6
      Left            =   120
      Picture         =   "Calculator.frx":1A4F2
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   19
      Top             =   2760
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   5
      Left            =   3720
      Picture         =   "Calculator.frx":1AE79
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   18
      Top             =   2040
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   4
      Left            =   3000
      Picture         =   "Calculator.frx":1B800
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   17
      Top             =   2040
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   3
      Left            =   2280
      Picture         =   "Calculator.frx":1C187
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   16
      Top             =   2040
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   2
      Left            =   1560
      Picture         =   "Calculator.frx":1CB0E
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   15
      Top             =   2040
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   1
      Left            =   840
      Picture         =   "Calculator.frx":1D495
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   14
      Top             =   2040
      Width           =   750
   End
   Begin VB.PictureBox pbButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   0
      Left            =   120
      Picture         =   "Calculator.frx":1DE1C
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   13
      Top             =   2040
      Width           =   750
   End
   Begin VB.Frame DegRadGradFrame 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
      Begin VB.OptionButton Grad 
         BackColor       =   &H00000000&
         Caption         =   "Grad"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         ToolTipText     =   "Gradients"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Rad 
         BackColor       =   &H00000000&
         Caption         =   "Rad"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   5
         ToolTipText     =   "Radians"
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Deg 
         BackColor       =   &H00000000&
         Caption         =   "Deg"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Degrees"
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame RadixFrame 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Radix"
      Top             =   600
      Width           =   4215
      Begin VB.OptionButton Roman 
         BackColor       =   &H00000000&
         Caption         =   "Rom"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Roman Numeral Input"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Hex 
         BackColor       =   &H00000000&
         Caption         =   "Hex"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   960
         TabIndex        =   10
         ToolTipText     =   "Hex Input"
         Top             =   180
         Width           =   615
      End
      Begin VB.OptionButton Dec 
         BackColor       =   &H00000000&
         Caption         =   "Dec"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         ToolTipText     =   "Decimal Input"
         Top             =   180
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Oct 
         BackColor       =   &H00000000&
         Caption         =   "Oct"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         ToolTipText     =   "Octal Input"
         Top             =   180
         Width           =   615
      End
      Begin VB.OptionButton Bin 
         BackColor       =   &H00000000&
         Caption         =   "Bin"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         ToolTipText     =   "Binary Input"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.TextBox OutputWindow 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0."
      ToolTipText     =   "Calculator Output Window"
      Top             =   120
      Width           =   4215
   End
   Begin VB.CheckBox InvCheckBox 
      BackColor       =   &H00000000&
      Caption         =   "Inv"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      ToolTipText     =   "Inverse - Causes some keys to perform the inverse of what's on the key"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "© 2005 by Tim Alvord"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   1365
   End
   Begin VB.Image About 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   3840
      Picture         =   "Calculator.frx":1E7A3
      ToolTipText     =   "Yankees Icon"
      Top             =   1410
      Width           =   510
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'   Scientific Calculator Program
'   Author: Timothy C. Alvord
'   E-Mail: tim8w@yahoo.com
'
'   Purpose:
'       This program is a Simple Scientific Calculator. It handles
'       Roman Numeral, Hex, Decimal, Octal and Binary numbers.
'       Degrees, Radians and Gradians. All the standard Trig functions.
'       Factorial, Sqr, Inverse, Square, Cube, X^Y.
'       Fun conversions like:
'           Ounce <-> Grams
'           Pounds <-> Kilograms
'           Gallon <-> Litre
'           Mile <-> Kilometer
'           Inch <-> Centimeter
'           Fahrenheight <-> Celsius
'***********************************************************************
Const PI = 3.14159265358979
' Function Mode
Const NONE = 0
Const MULTIPLY = 1
Const DIVIDE = 2
Const PLUS = 3
Const MINUS = 4
Const POWER = 5
' Base
Const ROMANNUM = 1
Const HEXNUM = 2
Const DECNUM = 3
Const OCTNUM = 4
Const BINNUM = 5
' Deg/Rad/Grad
Const DEGREES = 1
Const RADIANS = 2
Const GRADIANS = 3

Public xFirstNum As Double
Public xSecondNum As Double
Public xMemory As Double
Public bError As Boolean
Public bEntered As Boolean
Public bFirstNum As Boolean
Public iMathFunction As Integer
Public iCurrentRadix As Integer
Public iDegRadGrad As Integer

Private ecbButton(51) As ECommandButton
Const ButtonA = 0
Const ButtonB = 1
Const ButtonC = 2
Const ButtonD = 3
Const ButtonE = 4
Const ButtonF = 5
'
Const ButtonM = 6
Const ButtonL = 7
Const ButtonX = 8
Const ButtonV = 9
Const ButtonI = 10
Const ButtonClear = 11
'
Const ButtonSine = 12
Const ButtonCosine = 13
Const ButtonTangent = 14
Const ButtonCE = 15
'
Const ButtonPI = 16
Const ButtonOzToGram = 17
Const ButtonFactorial = 18
Const ButtonSqrRoot = 19
Const Button1OverX = 20
Const ButtonBS = 21
'
Const ButtonMC = 22
Const ButtonLbToKilo = 23
Const ButtonSquare = 24
Const ButtonCube = 25
Const ButtonPower = 26
Const ButtonDivide = 27
'
Const ButtonMR = 28
Const ButtonGallonToLitre = 29
Const Button7 = 30
Const Button8 = 31
Const Button9 = 32
Const ButtonMultiply = 33
'
Const ButtonMS = 34
Const ButtonMileToKilo = 35
Const Button4 = 36
Const Button5 = 37
Const Button6 = 38
Const ButtonMinus = 39
'
Const ButtonMPlus = 40
Const ButtonInchToCent = 41
Const Button1 = 42
Const Button2 = 43
Const Button3 = 44
Const ButtonPlus = 45
'
Const ButtonMMinus = 46
Const ButtonFToC = 47
Const ButtonPlusMinus = 48
Const Button0 = 49
Const ButtonDecimal = 50
Const ButtonEquals = 51
'
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Long, _
        lParam As Any) As Long

Private Const BM_CLICK As Long = &HF5
Private Const BM_SETSTATE As Long = &HF3

Private Const vbDarkBlue = &H900000

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
   iMathFunction = NONE
    iCurrentRadix = DECNUM
    iDegRadGrad = DEGREES
    xFirstNum = 0
    xSecondNum = 0
    bError = False
    bEntered = False
    
    SetupCommandButtons
    Call Dec_Click
End Sub

Private Sub About_Click()
    frmAboutBox.Show vbModal
End Sub
Private Sub SimulateButtonClick(iButton As Integer)
    Call pbButton_MouseDown(iButton, 1, 0, pbButton(iButton).ScaleHeight / 2, pbButton(iButton).ScaleWidth / 2)
    Call pbButton_GotFocus(iButton)
    Call pbButton_MouseUp(iButton, 1, 0, pbButton(iButton).ScaleHeight / 2, pbButton(iButton).ScaleWidth / 2)
    Call pbButton_LostFocus(iButton)
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    
    MyChar = Chr(KeyCode)
    Select Case (MyChar)
    Case "0"
        Call SimulateButtonClick(Button0)
    Case "1"
        Call SimulateButtonClick(Button1)
    Case "2"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call SimulateButtonClick(Button2)
        End If
    Case "3"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call SimulateButtonClick(Button3)
        End If
    Case "4"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call SimulateButtonClick(Button4)
        End If
    Case "5"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call SimulateButtonClick(Button5)
        End If
    Case "6"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call SimulateButtonClick(Button6)
        End If
    Case "7"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call SimulateButtonClick(Button7)
        End If
    Case "8"
        If Dec.Value = True Or Hex.Value = True Then
            Call SimulateButtonClick(Button8)
        End If
    Case "9"
        If Dec.Value = True Or Hex.Value = True Then
            Call SimulateButtonClick(Button9)
        End If
    Case "A", "a"
        If Hex.Value = True Then
            Call SimulateButtonClick(ButtonA)
        End If
    Case "B", "b"
        If Hex.Value = True Then
            Call SimulateButtonClick(ButtonB)
        End If
    Case "C", "c"
        If iCurrentRadix = ROMANNUM Or iCurrentRadix = HEXNUM Then
            Call SimulateButtonClick(ButtonC)
        End If
    Case "D", "d"
        If iCurrentRadix = ROMANNUM Or iCurrentRadix = HEXNUM Then
            Call SimulateButtonClick(ButtonD)
        End If
    Case "E", "e"
        If Hex.Value = True Then
            Call SimulateButtonClick(ButtonE)
        End If
    Case "F", "f"
        If Hex.Value = True Then
            Call SimulateButtonClick(ButtonF)
        End If
    Case "M", "m"
        If iCurrentRadix = ROMANNUM Then
            Call SimulateButtonClick(ButtonM)
        End If
    Case "L", "l"
        If iCurrentRadix = ROMANNUM Then
            Call SimulateButtonClick(ButtonL)
        End If
    Case "X", "x"
        If iCurrentRadix = ROMANNUM Then
            Call SimulateButtonClick(ButtonX)
        End If
    Case "V", "v"
        If iCurrentRadix = ROMANNUM Then
            Call SimulateButtonClick(ButtonV)
        End If
    Case "I", "i"
        If iCurrentRadix = ROMANNUM Then
            Call SimulateButtonClick(ButtonI)
        End If
    Case "."
        If Dec.Value = True Then
            Call SimulateButtonClick(ButtonDecimal)
        End If
    Case "/"
        Call SimulateButtonClick(ButtonDivide)
    Case "*"
        Call SimulateButtonClick(ButtonMultiply)
    Case "+"
        Call SimulateButtonClick(ButtonPlus)
    Case "-"
        Call SimulateButtonClick(ButtonMinus)
    Case "="
        Call SimulateButtonClick(ButtonEquals)
    Case Else
        Select Case (KeyCode)
        Case 8      '   Backspace Key
            Call SimulateButtonClick(ButtonBS)
        Case 13     '   Enter Key - Treat as Equal Key
            Call SimulateButtonClick(ButtonEquals)
        Case 27     '   Esc Key
            Call SimulateButtonClick(ButtonClear)
        End Select

    End Select
End Sub

Private Sub Roman_Click()
    If bError = False Then
        Call ConvertOutputWindowText(iCurrentRadix, ROMANNUM)
        iCurrentRadix = ROMANNUM
        
        DisableButton ecbButton(ButtonA), pbButton(ButtonA)
        DisableButton ecbButton(ButtonB), pbButton(ButtonB)
        DisableButton ecbButton(ButtonE), pbButton(ButtonE)
        DisableButton ecbButton(ButtonF), pbButton(ButtonF)
        DisableButton ecbButton(Button9), pbButton(Button9)
        DisableButton ecbButton(Button8), pbButton(Button8)
        DisableButton ecbButton(Button7), pbButton(Button7)
        DisableButton ecbButton(Button6), pbButton(Button6)
        DisableButton ecbButton(Button5), pbButton(Button5)
        DisableButton ecbButton(Button4), pbButton(Button4)
        DisableButton ecbButton(Button3), pbButton(Button3)
        DisableButton ecbButton(Button2), pbButton(Button2)
        DisableButton ecbButton(Button1), pbButton(Button1)
        DisableButton ecbButton(Button0), pbButton(Button0)
        DisableButton ecbButton(ButtonTangent), pbButton(ButtonTangent)
        DisableButton ecbButton(ButtonCosine), pbButton(ButtonCosine)
        DisableButton ecbButton(ButtonSine), pbButton(ButtonSine)
        DisableButton ecbButton(ButtonOzToGram), pbButton(ButtonOzToGram)
        DisableButton ecbButton(ButtonMileToKilo), pbButton(ButtonMileToKilo)
        DisableButton ecbButton(ButtonGallonToLitre), pbButton(ButtonGallonToLitre)
        DisableButton ecbButton(ButtonInchToCent), pbButton(ButtonInchToCent)
        DisableButton ecbButton(ButtonLbToKilo), pbButton(ButtonLbToKilo)
        DisableButton ecbButton(ButtonFToC), pbButton(ButtonFToC)
        DisableButton ecbButton(Button1OverX), pbButton(Button1OverX)
        DisableButton ecbButton(ButtonPlusMinus), pbButton(ButtonPlusMinus)
        
        EnableButton ecbButton(ButtonM), pbButton(ButtonM)
        EnableButton ecbButton(ButtonD), pbButton(ButtonD)
        EnableButton ecbButton(ButtonC), pbButton(ButtonC)
        EnableButton ecbButton(ButtonL), pbButton(ButtonL)
        EnableButton ecbButton(ButtonX), pbButton(ButtonX)
        EnableButton ecbButton(ButtonV), pbButton(ButtonV)
        EnableButton ecbButton(ButtonI), pbButton(ButtonI)
        EnableButton ecbButton(ButtonSqrRoot), pbButton(ButtonSqrRoot)
        EnableButton ecbButton(ButtonSquare), pbButton(ButtonSquare)
        EnableButton ecbButton(ButtonCube), pbButton(ButtonCube)
        EnableButton ecbButton(ButtonPower), pbButton(ButtonPower)
    End If
End Sub

Private Sub Hex_Click()
    If bError = False Then
        Call ConvertOutputWindowText(iCurrentRadix, HEXNUM)
        iCurrentRadix = HEXNUM
        
        DisableButton ecbButton(ButtonM), pbButton(ButtonM)
        DisableButton ecbButton(ButtonL), pbButton(ButtonL)
        DisableButton ecbButton(ButtonX), pbButton(ButtonX)
        DisableButton ecbButton(ButtonV), pbButton(ButtonV)
        DisableButton ecbButton(ButtonI), pbButton(ButtonI)
        DisableButton ecbButton(ButtonTangent), pbButton(ButtonTangent)
        DisableButton ecbButton(ButtonCosine), pbButton(ButtonCosine)
        DisableButton ecbButton(ButtonSine), pbButton(ButtonSine)
        DisableButton ecbButton(ButtonOzToGram), pbButton(ButtonOzToGram)
        DisableButton ecbButton(ButtonMileToKilo), pbButton(ButtonMileToKilo)
        DisableButton ecbButton(ButtonGallonToLitre), pbButton(ButtonGallonToLitre)
        DisableButton ecbButton(ButtonInchToCent), pbButton(ButtonInchToCent)
        DisableButton ecbButton(ButtonLbToKilo), pbButton(ButtonLbToKilo)
        DisableButton ecbButton(ButtonFToC), pbButton(ButtonFToC)
        
        EnableButton ecbButton(ButtonA), pbButton(ButtonA)
        EnableButton ecbButton(ButtonB), pbButton(ButtonB)
        EnableButton ecbButton(ButtonC), pbButton(ButtonC)
        EnableButton ecbButton(ButtonD), pbButton(ButtonD)
        EnableButton ecbButton(ButtonE), pbButton(ButtonE)
        EnableButton ecbButton(ButtonF), pbButton(ButtonF)
        EnableButton ecbButton(Button9), pbButton(Button9)
        EnableButton ecbButton(Button8), pbButton(Button8)
        EnableButton ecbButton(Button7), pbButton(Button7)
        EnableButton ecbButton(Button6), pbButton(Button6)
        EnableButton ecbButton(Button5), pbButton(Button5)
        EnableButton ecbButton(Button4), pbButton(Button4)
        EnableButton ecbButton(Button3), pbButton(Button3)
        EnableButton ecbButton(Button2), pbButton(Button2)
        EnableButton ecbButton(Button1), pbButton(Button1)
        EnableButton ecbButton(Button0), pbButton(Button0)
        EnableButton ecbButton(ButtonSqrRoot), pbButton(ButtonSqrRoot)
        EnableButton ecbButton(ButtonSquare), pbButton(ButtonSquare)
        EnableButton ecbButton(ButtonCube), pbButton(ButtonCube)
        EnableButton ecbButton(ButtonPower), pbButton(ButtonPower)
        EnableButton ecbButton(Button1OverX), pbButton(Button1OverX)
        EnableButton ecbButton(ButtonPlusMinus), pbButton(ButtonPlusMinus)
    End If
End Sub

Private Sub Dec_Click()
    If bError = False Then
        Call ConvertOutputWindowText(iCurrentRadix, DECNUM)
        iCurrentRadix = DECNUM
        
        DisableButton ecbButton(ButtonA), pbButton(ButtonA)
        DisableButton ecbButton(ButtonB), pbButton(ButtonB)
        DisableButton ecbButton(ButtonC), pbButton(ButtonC)
        DisableButton ecbButton(ButtonD), pbButton(ButtonD)
        DisableButton ecbButton(ButtonE), pbButton(ButtonE)
        DisableButton ecbButton(ButtonF), pbButton(ButtonF)
        DisableButton ecbButton(ButtonM), pbButton(ButtonM)
        DisableButton ecbButton(ButtonL), pbButton(ButtonL)
        DisableButton ecbButton(ButtonX), pbButton(ButtonX)
        DisableButton ecbButton(ButtonV), pbButton(ButtonV)
        DisableButton ecbButton(ButtonI), pbButton(ButtonI)
        
        EnableButton ecbButton(Button9), pbButton(Button9)
        EnableButton ecbButton(Button8), pbButton(Button8)
        EnableButton ecbButton(Button7), pbButton(Button7)
        EnableButton ecbButton(Button6), pbButton(Button6)
        EnableButton ecbButton(Button5), pbButton(Button5)
        EnableButton ecbButton(Button4), pbButton(Button4)
        EnableButton ecbButton(Button3), pbButton(Button3)
        EnableButton ecbButton(Button2), pbButton(Button2)
        EnableButton ecbButton(Button1), pbButton(Button1)
        EnableButton ecbButton(Button0), pbButton(Button0)
        EnableButton ecbButton(ButtonTangent), pbButton(ButtonTangent)
        EnableButton ecbButton(ButtonCosine), pbButton(ButtonCosine)
        EnableButton ecbButton(ButtonSine), pbButton(ButtonSine)
        EnableButton ecbButton(ButtonOzToGram), pbButton(ButtonOzToGram)
        EnableButton ecbButton(ButtonMileToKilo), pbButton(ButtonMileToKilo)
        EnableButton ecbButton(ButtonGallonToLitre), pbButton(ButtonGallonToLitre)
        EnableButton ecbButton(ButtonInchToCent), pbButton(ButtonInchToCent)
        EnableButton ecbButton(ButtonLbToKilo), pbButton(ButtonLbToKilo)
        EnableButton ecbButton(ButtonFToC), pbButton(ButtonFToC)
        EnableButton ecbButton(ButtonSqrRoot), pbButton(ButtonSqrRoot)
        EnableButton ecbButton(ButtonSquare), pbButton(ButtonSquare)
        EnableButton ecbButton(ButtonCube), pbButton(ButtonCube)
        EnableButton ecbButton(ButtonPower), pbButton(ButtonPower)
        EnableButton ecbButton(Button1OverX), pbButton(Button1OverX)
        EnableButton ecbButton(ButtonPlusMinus), pbButton(ButtonPlusMinus)
    End If
End Sub

Private Sub Oct_Click()
    If bError = False Then
        Call ConvertOutputWindowText(iCurrentRadix, OCTNUM)
        iCurrentRadix = OCTNUM
        
        DisableButton ecbButton(ButtonA), pbButton(ButtonA)
        DisableButton ecbButton(ButtonB), pbButton(ButtonB)
        DisableButton ecbButton(ButtonC), pbButton(ButtonC)
        DisableButton ecbButton(ButtonD), pbButton(ButtonD)
        DisableButton ecbButton(ButtonE), pbButton(ButtonE)
        DisableButton ecbButton(ButtonF), pbButton(ButtonF)
        DisableButton ecbButton(ButtonM), pbButton(ButtonM)
        DisableButton ecbButton(ButtonL), pbButton(ButtonL)
        DisableButton ecbButton(ButtonX), pbButton(ButtonX)
        DisableButton ecbButton(ButtonV), pbButton(ButtonV)
        DisableButton ecbButton(ButtonI), pbButton(ButtonI)
        DisableButton ecbButton(Button9), pbButton(Button9)
        DisableButton ecbButton(Button8), pbButton(Button8)
        DisableButton ecbButton(ButtonTangent), pbButton(ButtonTangent)
        DisableButton ecbButton(ButtonCosine), pbButton(ButtonCosine)
        DisableButton ecbButton(ButtonSine), pbButton(ButtonSine)
        DisableButton ecbButton(ButtonOzToGram), pbButton(ButtonOzToGram)
        DisableButton ecbButton(ButtonMileToKilo), pbButton(ButtonMileToKilo)
        DisableButton ecbButton(ButtonGallonToLitre), pbButton(ButtonGallonToLitre)
        DisableButton ecbButton(ButtonInchToCent), pbButton(ButtonInchToCent)
        DisableButton ecbButton(ButtonLbToKilo), pbButton(ButtonLbToKilo)
        DisableButton ecbButton(ButtonFToC), pbButton(ButtonFToC)
        DisableButton ecbButton(Button1OverX), pbButton(Button1OverX)
        
        EnableButton ecbButton(Button7), pbButton(Button7)
        EnableButton ecbButton(Button6), pbButton(Button6)
        EnableButton ecbButton(Button5), pbButton(Button5)
        EnableButton ecbButton(Button4), pbButton(Button4)
        EnableButton ecbButton(Button3), pbButton(Button3)
        EnableButton ecbButton(Button2), pbButton(Button2)
        EnableButton ecbButton(Button1), pbButton(Button1)
        EnableButton ecbButton(Button0), pbButton(Button0)
        EnableButton ecbButton(ButtonSqrRoot), pbButton(ButtonSqrRoot)
        EnableButton ecbButton(ButtonSquare), pbButton(ButtonSquare)
        EnableButton ecbButton(ButtonCube), pbButton(ButtonCube)
        EnableButton ecbButton(ButtonPower), pbButton(ButtonPower)
        EnableButton ecbButton(ButtonPlusMinus), pbButton(ButtonPlusMinus)
    End If
End Sub

Private Sub Bin_Click()
    If bError = False Then
        Call ConvertOutputWindowText(iCurrentRadix, BINNUM)
        iCurrentRadix = BINNUM
        
        DisableButton ecbButton(ButtonA), pbButton(ButtonA)
        DisableButton ecbButton(ButtonB), pbButton(ButtonB)
        DisableButton ecbButton(ButtonC), pbButton(ButtonC)
        DisableButton ecbButton(ButtonD), pbButton(ButtonD)
        DisableButton ecbButton(ButtonE), pbButton(ButtonE)
        DisableButton ecbButton(ButtonF), pbButton(ButtonF)
        DisableButton ecbButton(ButtonM), pbButton(ButtonM)
        DisableButton ecbButton(ButtonL), pbButton(ButtonL)
        DisableButton ecbButton(ButtonX), pbButton(ButtonX)
        DisableButton ecbButton(ButtonV), pbButton(ButtonV)
        DisableButton ecbButton(ButtonI), pbButton(ButtonI)
        DisableButton ecbButton(Button9), pbButton(Button9)
        DisableButton ecbButton(Button8), pbButton(Button8)
        DisableButton ecbButton(Button7), pbButton(Button7)
        DisableButton ecbButton(Button6), pbButton(Button6)
        DisableButton ecbButton(Button5), pbButton(Button5)
        DisableButton ecbButton(Button4), pbButton(Button4)
        DisableButton ecbButton(Button3), pbButton(Button3)
        DisableButton ecbButton(Button2), pbButton(Button2)
        DisableButton ecbButton(ButtonTangent), pbButton(ButtonTangent)
        DisableButton ecbButton(ButtonCosine), pbButton(ButtonCosine)
        DisableButton ecbButton(ButtonSine), pbButton(ButtonSine)
        DisableButton ecbButton(ButtonOzToGram), pbButton(ButtonOzToGram)
        DisableButton ecbButton(ButtonMileToKilo), pbButton(ButtonMileToKilo)
        DisableButton ecbButton(ButtonGallonToLitre), pbButton(ButtonGallonToLitre)
        DisableButton ecbButton(ButtonInchToCent), pbButton(ButtonInchToCent)
        DisableButton ecbButton(ButtonLbToKilo), pbButton(ButtonLbToKilo)
        DisableButton ecbButton(ButtonFToC), pbButton(ButtonFToC)
        DisableButton ecbButton(Button1OverX), pbButton(Button1OverX)
        
        EnableButton ecbButton(Button1), pbButton(Button1)
        EnableButton ecbButton(Button0), pbButton(Button0)
        EnableButton ecbButton(ButtonSqrRoot), pbButton(ButtonSqrRoot)
        EnableButton ecbButton(ButtonSquare), pbButton(ButtonSquare)
        EnableButton ecbButton(ButtonCube), pbButton(ButtonCube)
        EnableButton ecbButton(ButtonPower), pbButton(ButtonPower)
        EnableButton ecbButton(ButtonPlusMinus), pbButton(ButtonPlusMinus)
    End If
End Sub

Private Sub Deg_Click()
        iDegRadGrad = DEGREES
End Sub

Private Sub Rad_Click()
        iDegRadGrad = RADIANS
End Sub

Private Sub Grad_Click()
        iDegRadGrad = GRADIANS
End Sub

Private Sub DoButtonPI()
    If bError = False Then
        Select Case iCurrentRadix
            Case ROMANNUM
                OutputWindow.Text = GetDecRomanStr(PI)
            Case HEXNUM
                OutputWindow.Text = GetDecHexStr(PI)
            Case DECNUM
                OutputWindow.Text = PI
            Case OCTNUM
                OutputWindow.Text = GetDecOctStr(PI)
            Case BINNUM
                OutputWindow.Text = GetDecBinStr(PI)
        End Select
        bEntered = True
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonClear()
    bEntered = False
    bFirstNum = False
    OutputWindow.Text = "0."
    OutputWindow.ForeColor = &H0&
    iMathFunction = 0
    bError = False
    DoButtonMC
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonCE()
    If bError = False Then
        If bEntered = True Then
            If bFirstNum = False Then   '   Same as Clear
                bEntered = False
                bFirstNum = False
                OutputWindow.Text = "0."
                iMathFunction = 0
            Else    '   Allow User to Enter a new 2nd Number
                bEntered = False
                OutputWindow.Text = "0."
            End If
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonBS()
    If bError = False Then
        If bEntered = True Then
            sStr = OutputWindow.Text
            If Len(sStr) > 1 Then
                OutputWindow.Text = Left(sStr, Len(sStr) - 1)
            Else
                OutputWindow.Text = "0"
            End If
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButton0()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "0"
        Else
            OutputWindow.Text = OutputWindow.Text + "0"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButton1()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "1"
        Else
            OutputWindow.Text = OutputWindow.Text + "1"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButton2()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "2"
        Else
            OutputWindow.Text = OutputWindow.Text + "2"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButton3()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "3"
        Else
            OutputWindow.Text = OutputWindow.Text + "3"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButton4()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "4"
        Else
            OutputWindow.Text = OutputWindow.Text + "4"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButton5()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "5"
        Else
            OutputWindow.Text = OutputWindow.Text + "5"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButton6()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "6"
        Else
            OutputWindow.Text = OutputWindow.Text + "6"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButton7()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "7"
        Else
            OutputWindow.Text = OutputWindow.Text + "7"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButton8()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "8"
        Else
            OutputWindow.Text = OutputWindow.Text + "8"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButton9()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "9"
        Else
            OutputWindow.Text = OutputWindow.Text + "9"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonA()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "A"
        Else
            OutputWindow.Text = OutputWindow.Text + "A"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonB()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "B"
        Else
            OutputWindow.Text = OutputWindow.Text + "B"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonC()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "C"
        Else
            OutputWindow.Text = OutputWindow.Text + "C"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonD()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "D"
        Else
            OutputWindow.Text = OutputWindow.Text + "D"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonE()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "E"
        Else
            OutputWindow.Text = OutputWindow.Text + "E"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonF()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "F"
        Else
            OutputWindow.Text = OutputWindow.Text + "F"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonM()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "M"
        Else
            OutputWindow.Text = OutputWindow.Text + "M"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonL()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "L"
        Else
            OutputWindow.Text = OutputWindow.Text + "L"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonX()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "X"
        Else
            OutputWindow.Text = OutputWindow.Text + "X"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonV()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "V"
        Else
            OutputWindow.Text = OutputWindow.Text + "V"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonI()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "I"
        Else
            OutputWindow.Text = OutputWindow.Text + "I"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonDecimal()
    If bError = False Then
        If iCurrentRadix = DECNUM Then
            If bEntered = False Then    '   Start of New Number
                bEntered = True
                OutputWindow.Text = "."
            Else                        '   Append
                If InStr(OutputWindow.Text, ".") = 0 Then   '   Make Sure Only 1 Decimal Point
                    OutputWindow.Text = OutputWindow.Text + "."
                End If
            End If
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonPlusMinus()
    If bError = False Then
        If iCurrentRadix = DECNUM Then
            If bEntered = True Then
                iValue = OutputWindow.Text * -1
                OutputWindow.Text = iValue
            End If
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonDivide()
    If bError = False Then
        If bEntered = True Then
            If bFirstNum = True Then
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xSecondNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xSecondNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xSecondNum = OutputWindow.Text
                    Case OCTNUM
                        xSecondNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xSecondNum = GetBinDecNum(OutputWindow.Text)
                End Select
                bError = Calculate_Total(xFirstNum, xSecondNum, xTotal)
                If Not bError Then
                    xFirstNum = xTotal
                    bFirstNum = True
                    bEntered = False
                    iMathFunction = DIVIDE
                End If
            Else
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xFirstNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xFirstNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xFirstNum = OutputWindow.Text
                    Case OCTNUM
                        xFirstNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xFirstNum = GetBinDecNum(OutputWindow.Text)
                End Select
                bFirstNum = True
                bEntered = False
                iMathFunction = DIVIDE
            End If
        Else
            iMathFunction = DIVIDE
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonMultiply()
    If bError = False Then
        If bEntered = True Then
            If bFirstNum = True Then
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xSecondNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xSecondNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xSecondNum = OutputWindow.Text
                    Case OCTNUM
                        xSecondNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xSecondNum = GetBinDecNum(OutputWindow.Text)
                End Select
                bError = Calculate_Total(xFirstNum, xSecondNum, xTotal)
                If Not bError Then
                    xFirstNum = xTotal
                    bFirstNum = True
                    bEntered = False
                    iMathFunction = MULTIPLY
                End If
           Else
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xFirstNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xFirstNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xFirstNum = OutputWindow.Text
                    Case OCTNUM
                        xFirstNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xFirstNum = GetBinDecNum(OutputWindow.Text)
                End Select
                bFirstNum = True
                bEntered = False
                iMathFunction = MULTIPLY
            End If
        Else
            iMathFunction = MULTIPLY
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonPlus()
    If bError = False Then
        If bEntered = True Then
            If bFirstNum = True Then
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xSecondNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xSecondNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xSecondNum = OutputWindow.Text
                    Case OCTNUM
                        xSecondNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xSecondNum = GetBinDecNum(OutputWindow.Text)
                End Select
                bError = Calculate_Total(xFirstNum, xSecondNum, xTotal)
                If Not bError Then
                    xFirstNum = xTotal
                    bFirstNum = True
                    bEntered = False
                    iMathFunction = PLUS
                End If
            Else
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xFirstNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xFirstNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xFirstNum = OutputWindow.Text
                    Case OCTNUM
                        xFirstNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xFirstNum = GetBinDecNum(OutputWindow.Text)
                End Select
                bFirstNum = True
                bEntered = False
                iMathFunction = PLUS
            End If
        Else
            iMathFunction = PLUS
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonMinus()
    If bError = False Then
        If bEntered = True Then
            If bFirstNum = True Then
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xSecondNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xSecondNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xSecondNum = OutputWindow.Text
                    Case OCTNUM
                        xSecondNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xSecondNum = GetBinDecNum(OutputWindow.Text)
                End Select
                bError = Calculate_Total(xFirstNum, xSecondNum, xTotal)
                If Not bError Then
                    xFirstNum = xTotal
                    bFirstNum = True
                    bEntered = False
                    iMathFunction = MINUS
                End If
            Else
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xFirstNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xFirstNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xFirstNum = OutputWindow.Text
                    Case OCTNUM
                        xFirstNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xFirstNum = GetBinDecNum(OutputWindow.Text)
                End Select
                bFirstNum = True
                bEntered = False
                iMathFunction = MINUS
            End If
        Else
            iMathFunction = MINUS
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonPower()
    If bError = False Then
        If bEntered = True Then
            Select Case iCurrentRadix
                Case ROMANNUM
                    xFirstNum = GetRomanDecNum(OutputWindow.Text)
                Case HEXNUM
                    xFirstNum = Val("&H" + OutputWindow.Text)
                Case DECNUM
                    xFirstNum = OutputWindow.Text
                Case OCTNUM
                    xFirstNum = Val("&O" + OutputWindow.Text)
                Case BINNUM
                    xFirstNum = GetBinDecNum(OutputWindow.Text)
            End Select
            bFirstNum = True
            bEntered = False
            iMathFunction = POWER
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonEquals()
    If bError = False Then
        If bFirstNum = True Then
            Select Case iCurrentRadix
                Case ROMANNUM
                    xSecondNum = GetRomanDecNum(OutputWindow.Text)
                Case HEXNUM
                    xSecondNum = Val("&H" + OutputWindow.Text)
                Case DECNUM
                    xSecondNum = OutputWindow.Text
                Case OCTNUM
                    xSecondNum = Val("&O" + OutputWindow.Text)
                Case BINNUM
                    xSecondNum = GetBinDecNum(OutputWindow.Text)
            End Select
            
            Select Case iMathFunction
            Case DIVIDE
                If xSecondNum = 0 Then
                    OutputWindow.Text = "ERROR - Divide by Zero"
                    OutputWindow.ForeColor = &HFF&
                    bError = True
                Else
                    xTotal = xFirstNum / xSecondNum
                    Select Case iCurrentRadix
                        Case ROMANNUM
                            OutputWindow.Text = GetDecRomanStr(xTotal)
                        Case HEXNUM
                            OutputWindow.Text = GetDecHexStr(xTotal)
                        Case DECNUM
                            OutputWindow.Text = xTotal
                        Case OCTNUM
                            OutputWindow.Text = GetDecOctStr(xTotal)
                        Case BINNUM
                            OutputWindow.Text = GetDecBinStr(xTotal)
                    End Select
                    
                    xFirstNum = xTotal
                    bFirstNum = True
                    bEntered = False
                End If
                iMathFunction = 0
            Case MULTIPLY
                xTotal = xFirstNum * xSecondNum
                Select Case iCurrentRadix
                    Case ROMANNUM
                        OutputWindow.Text = GetDecRomanStr(xTotal)
                    Case HEXNUM
                        OutputWindow.Text = GetDecHexStr(xTotal)
                    Case DECNUM
                        OutputWindow.Text = xTotal
                    Case OCTNUM
                        OutputWindow.Text = GetDecOctStr(xTotal)
                    Case BINNUM
                        OutputWindow.Text = GetDecBinStr(xTotal)
                End Select
                
                xFirstNum = xTotal
                bFirstNum = True
                bEntered = False
                iMathFunction = 0
            Case PLUS
                xTotal = xFirstNum + xSecondNum
                Select Case iCurrentRadix
                    Case ROMANNUM
                        OutputWindow.Text = GetDecRomanStr(xTotal)
                    Case HEXNUM
                        OutputWindow.Text = GetDecHexStr(xTotal)
                    Case DECNUM
                        OutputWindow.Text = xTotal
                    Case OCTNUM
                        OutputWindow.Text = GetDecOctStr(xTotal)
                    Case BINNUM
                        OutputWindow.Text = GetDecBinStr(xTotal)
                End Select
                
                xFirstNum = xTotal
                bFirstNum = True
                bEntered = False
                iMathFunction = 0
            Case MINUS
                xTotal = xFirstNum - xSecondNum
                Select Case iCurrentRadix
                    Case ROMANNUM
                        OutputWindow.Text = GetDecRomanStr(xTotal)
                    Case HEXNUM
                        OutputWindow.Text = GetDecHexStr(xTotal)
                    Case DECNUM
                        OutputWindow.Text = xTotal
                    Case OCTNUM
                        OutputWindow.Text = GetDecOctStr(xTotal)
                    Case BINNUM
                        OutputWindow.Text = GetDecBinStr(xTotal)
                End Select
                xFirstNum = xTotal
                bFirstNum = True
                bEntered = False
                iMathFunction = 0
            Case POWER
                xTotal = 1
                If xSecondNum <> 0 Then
                    For iCounter = 1 To xSecondNum
                        xTotal = xTotal * xFirstNum
                    Next
                End If
                Select Case iCurrentRadix
                    Case ROMANNUM
                        OutputWindow.Text = GetDecRomanStr(xTotal)
                    Case HEXNUM
                        OutputWindow.Text = GetDecHexStr(xTotal)
                    Case DECNUM
                        OutputWindow.Text = xTotal
                    Case OCTNUM
                        OutputWindow.Text = GetDecOctStr(xTotal)
                    Case BINNUM
                        OutputWindow.Text = GetDecBinStr(xTotal)
                End Select
                xFirstNum = xTotal
                bFirstNum = True
                bEntered = False
                iMathFunction = 0
            End Select
        End If
    End If
    OutputWindow.SetFocus
End Sub
Private Sub DoButtonMC()
    xMemory = 0
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonMR()
    If xMemory <> 0 Then
        Select Case iCurrentRadix
            Case ROMANNUM
                OutputWindow.Text = GetDecRomanStr(xMemory)
            Case HEXNUM
                OutputWindow.Text = GetDecHexStr(xMemory)
            Case DECNUM
                OutputWindow.Text = xMemory
            Case OCTNUM
                OutputWindow.Text = GetDecOctStr(xMemory)
            Case BINNUM
                OutputWindow.Text = GetDecBinStr(xMemory)
        End Select
        bEntered = True
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonMS()
    Select Case iCurrentRadix
        Case ROMANNUM
            xMemory = GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xMemory = Val("&H" + OutputWindow.Text)
        Case DECNUM
            xMemory = OutputWindow.Text
        Case OCTNUM
            xMemory = Val("&O" + OutputWindow.Text)
        Case BINNUM
            xMemory = GetBinDecNum(OutputWindow.Text)
    End Select
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonMPlus()
    Select Case iCurrentRadix
        Case ROMANNUM
            xMemory = xMemory + GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xMemory = xMemory + Val("&H" + OutputWindow.Text)
        Case DECNUM
            xMemory = xMemory + OutputWindow.Text
        Case OCTNUM
            xMemory = xMemory + Val("&O" + OutputWindow.Text)
        Case BINNUM
            xMemory = xMemory + GetBinDecNum(OutputWindow.Text)
    End Select
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonMMinus()
    Select Case iCurrentRadix
        Case ROMANNUM
            xMemory = xMemory - GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xMemory = xMemory - Val("&H" + OutputWindow.Text)
        Case DECNUM
            xMemory = xMemory - OutputWindow.Text
        Case OCTNUM
            xMemory = xMemory - Val("&O" + OutputWindow.Text)
        Case BINNUM
            xMemory = xMemory - GetBinDecNum(OutputWindow.Text)
    End Select
    OutputWindow.SetFocus
End Sub

Private Sub InvCheckBox_Click()
    If InvCheckBox.Value = 1 Then
        ecbButton(ButtonOzToGram).Text = "Gm-Oz"
        ecbButton(ButtonMileToKilo).Text = "Km-Mi"
        ecbButton(ButtonGallonToLitre).Text = "Ltr-Gal"
        ecbButton(ButtonInchToCent).Text = "Cm-In"
        ecbButton(ButtonFToC).Text = "C - F"
        ecbButton(ButtonLbToKilo).Text = "Kg-Lb"
        ecbButton(ButtonSine).Text = "Asn"
        ecbButton(ButtonCosine).Text = "Acs"
        ecbButton(ButtonTangent).Text = "Atn"
    Else
        ecbButton(ButtonOzToGram).Text = "Oz-Gm"
        ecbButton(ButtonMileToKilo).Text = "Mi-Km"
        ecbButton(ButtonGallonToLitre).Text = "Gal-Ltr"
        ecbButton(ButtonInchToCent).Text = "In-Cm"
        ecbButton(ButtonFToC).Text = "F - C"
        ecbButton(ButtonLbToKilo).Text = "Lb-Kg"
        ecbButton(ButtonSine).Text = "Sin"
        ecbButton(ButtonCosine).Text = "Cos"
        ecbButton(ButtonTangent).Text = "Tan"
    End If
    ' Redraw the buttons
    DrawCommandButton ecbButton(ButtonOzToGram), pbButton(ButtonOzToGram)
    DrawCommandButton ecbButton(ButtonMileToKilo), pbButton(ButtonMileToKilo)
    DrawCommandButton ecbButton(ButtonGallonToLitre), pbButton(ButtonGallonToLitre)
    DrawCommandButton ecbButton(ButtonInchToCent), pbButton(ButtonInchToCent)
    DrawCommandButton ecbButton(ButtonFToC), pbButton(ButtonFToC)
    DrawCommandButton ecbButton(ButtonLbToKilo), pbButton(ButtonLbToKilo)
    DrawCommandButton ecbButton(ButtonSine), pbButton(ButtonSine)
    DrawCommandButton ecbButton(ButtonCosine), pbButton(ButtonCosine)
    DrawCommandButton ecbButton(ButtonTangent), pbButton(ButtonTangent)
    OutputWindow.SetFocus
End Sub
Private Sub DoButtonOzToGram()
    If InvCheckBox.Value = 1 Then
        xGram = OutputWindow.Text
        xOunce = xGram / 28.34952313
        OutputWindow.Text = xOunce
        InvCheckBox.Value = 0
        InvCheckBox_Click
    Else
        xOunce = OutputWindow.Text
        xGram = xOunce * 28.34952313
        OutputWindow.Text = xGram
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonMileToKilo()
    If InvCheckBox.Value = 1 Then
        xKilometer = OutputWindow.Text
        xMile = xKilometer / 1.609344
        OutputWindow.Text = xMile
        InvCheckBox.Value = 0
        InvCheckBox_Click
    Else
        xMile = OutputWindow.Text
        xKilometer = xMile * 1.609344
        OutputWindow.Text = xKilometer
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonGallonToLitre()
    If InvCheckBox.Value = 1 Then
        xLitre = OutputWindow.Text
        xGallon = xLitre / 3.785412
        OutputWindow.Text = xGallon
        InvCheckBox.Value = 0
        InvCheckBox_Click
    Else
        xGallon = OutputWindow.Text
        xLitre = xGallon * 3.785412
        OutputWindow.Text = xLitre
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonInchToCent()
    If InvCheckBox.Value = 1 Then
        xCent = OutputWindow.Text
        xInch = xCent / 2.54
        OutputWindow.Text = xInch
        InvCheckBox.Value = 0
        InvCheckBox_Click
    Else
        xInch = OutputWindow.Text
        xCent = xInch * 2.54
        OutputWindow.Text = xCent
    End If
    OutputWindow.SetFocus
End Sub
Private Sub DoButtonFToC()
    If InvCheckBox.Value = 1 Then
        xCelsius = OutputWindow.Text
        xFahrenheit = 32 + (9 * xCelsius / 5)
        OutputWindow.Text = xFahrenheit
        InvCheckBox.Value = 0
        InvCheckBox_Click
    Else
        xFahrenheit = OutputWindow.Text
        xCelsius = ((xFahrenheit - 32) * 5) / 9
        OutputWindow.Text = xCelsius
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonLbToKilo()
    If InvCheckBox.Value = 1 Then
        xKilograms = OutputWindow.Text
        xPounds = xKilograms * 2.204623
        OutputWindow.Text = xPounds
        InvCheckBox.Value = 0
        InvCheckBox_Click
    Else
        xPounds = OutputWindow.Text
        xKilograms = xPounds / 2.204623
        OutputWindow.Text = xKilograms
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonSine()
    If InvCheckBox.Value = 1 Then
        xAsin = ArcSin(OutputWindow.Text)
        If bError Then Exit Sub
        Select Case iDegRadGrad
            Case DEGREES
                OutputWindow.Text = xAsin * 180 / PI
            Case RADIANS
                OutputWindow.Text = xAsin
            Case GRADIANS
                xDeg = xAsin * 180 / PI
                OutputWindow.Text = xDeg * 10 / 9
        End Select
        InvCheckBox.Value = 0
        InvCheckBox_Click
    Else
        Select Case iDegRadGrad
            Case DEGREES
                xsin = Sin(OutputWindow.Text * PI / 180)
            Case RADIANS
                xsin = Sin(OutputWindow.Text)
            Case GRADIANS
                xDeg = (OutputWindow.Text * 9) / 10
                xsin = Sin(xDeg * PI / 180)
        End Select
        OutputWindow.Text = xsin
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonCosine()
    If InvCheckBox.Value = 1 Then
        xAcos = ArcCos(OutputWindow.Text)
        If bError Then Exit Sub
        Select Case iDegRadGrad
            Case DEGREES
                OutputWindow.Text = xAcos * 180 / PI
            Case RADIANS
                OutputWindow.Text = xAcos
            Case GRADIANS
                xDeg = xAcos * 10 / 9
                OutputWindow.Text = xDeg * 180 / PI
        End Select
        InvCheckBox.Value = 0
        InvCheckBox_Click
    Else
        Select Case iDegRadGrad
            Case DEGREES
                xcos = Cos(OutputWindow.Text * PI / 180)
            Case RADIANS
                xcos = Cos(OutputWindow.Text)
            Case GRADIANS
                xDeg = (OutputWindow.Text * 9) / 10
                xcos = Cos(xDeg * PI / 180)
        End Select
        OutputWindow.Text = xcos
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonTangent()
    If InvCheckBox.Value = 1 Then
        xAtan = Atn(OutputWindow.Text)
        Select Case iDegRadGrad
            Case DEGREES
                OutputWindow.Text = xAtan * 180 / PI
            Case RADIANS
                OutputWindow.Text = xAtan
            Case GRADIANS
                xDeg = xAtan * 180 / PI
                OutputWindow.Text = xDeg * 10 / 9
        End Select
        InvCheckBox.Value = 0
        InvCheckBox_Click
    Else
        Select Case iDegRadGrad
            Case DEGREES
                xtan = Tan(OutputWindow.Text * PI / 180)
            Case RADIANS
                xtan = Tan(OutputWindow.Text)
            Case GRADIANS
                xDeg = (OutputWindow.Text * 9) / 10
                xtan = Tan(xDeg * PI / 180)
        End Select
        OutputWindow.Text = xtan
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonSquare()
    Select Case iCurrentRadix
        Case ROMANNUM
            xNum = GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xNum = Val("&H" + OutputWindow.Text)
        Case DECNUM
            xNum = OutputWindow.Text
        Case OCTNUM
            xNum = Val("&O" + OutputWindow.Text)
        Case BINNUM
            xNum = GetBinDecNum(OutputWindow.Text)
    End Select
    xNum = xNum * xNum
    Select Case iCurrentRadix
        Case ROMANNUM
            OutputWindow.Text = GetDecRomanStr(xNum)
        Case HEXNUM
            OutputWindow.Text = GetDecHexStr(xNum)
        Case DECNUM
            OutputWindow.Text = xNum
        Case OCTNUM
            OutputWindow.Text = GetDecOctStr(xNum)
        Case BINNUM
            OutputWindow.Text = GetDecBinStr(xNum)
    End Select
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonCube()
    Select Case iCurrentRadix
        Case ROMANNUM
            xNum = GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xNum = Val("&H" + OutputWindow.Text)
        Case DECNUM
            xNum = OutputWindow.Text
        Case OCTNUM
            xNum = Val("&O" + OutputWindow.Text)
        Case BINNUM
            xNum = GetBinDecNum(OutputWindow.Text)
    End Select
    xNum = xNum * xNum * xNum
    Select Case iCurrentRadix
        Case ROMANNUM
            OutputWindow.Text = GetDecRomanStr(xNum)
        Case HEXNUM
            OutputWindow.Text = GetDecHexStr(xNum)
        Case DECNUM
            OutputWindow.Text = xNum
        Case OCTNUM
            OutputWindow.Text = GetDecOctStr(xNum)
        Case BINNUM
            OutputWindow.Text = GetDecBinStr(xNum)
    End Select
    OutputWindow.SetFocus
End Sub
Public Function ArcSin(x As Variant) As Variant
    Select Case x
        Case -1
            ArcSin = 6 * Atn(1)
        Case 0:
            ArcSin = 0
        Case 1:
            ArcSin = 2 * Atn(1)
        Case Else:
            ArcSin = Atn(x / Sqr(-x * x + 1))
    End Select
End Function
Public Function ArcCos(x As Variant) As Variant

    Select Case x
        Case -1
            ArcCos = 4 * Atn(1)
             
        Case 0:
            ArcCos = 2 * Atn(1)
             
        Case 1:
            ArcCos = 0
             
        Case Else:
            ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    End Select
End Function

Private Sub DoButton1OverX()
    xValue = OutputWindow.Text
    If xValue = 0 Then
        OutputWindow.Text = "ERROR - Divide by Zero"
        OutputWindow.ForeColor = &HFF&
        bError = True
    Else
        OutputWindow.Text = 1 / xValue
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonSqrRoot()
    Select Case iCurrentRadix
        Case ROMANNUM
            xValue = GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xValue = Val("&H" + OutputWindow.Text)
        Case DECNUM
            xValue = OutputWindow.Text
        Case OCTNUM
            xValue = Val("&O" + OutputWindow.Text)
        Case BINNUM
            xValue = GetBinDecNum(OutputWindow.Text)
    End Select
    If xValue < 0 Then
        OutputWindow.Text = "ERROR - Square Root of Negative Number"
        OutputWindow.ForeColor = &HFF&
        bError = True
    Else
        Select Case iCurrentRadix
            Case ROMANNUM
                OutputWindow.Text = GetDecRomanStr(Sqr(xValue))
            Case HEXNUM
                OutputWindow.Text = GetDecHexStr(Sqr(xValue))
            Case DECNUM
                OutputWindow.Text = Sqr(xValue)
            Case OCTNUM
                OutputWindow.Text = GetDecOctStr(Sqr(xValue))
            Case BINNUM
                OutputWindow.Text = GetDecBinStr(Sqr(xValue))
        End Select
    End If
    OutputWindow.SetFocus
End Sub

Private Sub DoButtonFactorial()
    Dim xFactorial, xNum As Double
    
    Select Case iCurrentRadix
        Case ROMANNUM
            xNum = GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xNum = Val("&H" + OutputWindow.Text)
        Case DECNUM
            xNum = OutputWindow.Text
        Case OCTNUM
            xNum = Val("&O" + OutputWindow.Text)
        Case BINNUM
            xNum = GetBinDecNum(OutputWindow.Text)
    End Select
    xFactorial = xNum
    On Error GoTo Factorial_Error
    For iCounter = (xFactorial - 1) To 1 Step -1
        xFactorial = xFactorial * iCounter
    Next
    Select Case iCurrentRadix
        Case ROMANNUM
            OutputWindow.Text = GetDecRomanStr(xFactorial)
        Case HEXNUM
            OutputWindow.Text = GetDecHexStr(xFactorial)
        Case DECNUM
            OutputWindow.Text = xFactorial
        Case OCTNUM
            OutputWindow.Text = GetDecOctStr(xFactorial)
        Case BINNUM
            OutputWindow.Text = GetDecBinStr(xFactorial)
    End Select
    OutputWindow.SetFocus
    Exit Sub
    
Factorial_Error:
    OutputWindow.Text = "ERROR - " + Err.Description
    OutputWindow.ForeColor = &HFF&
    bError = True
    Err.Clear
    OutputWindow.SetFocus
End Sub

Public Function GetDecRomanStr(ByVal xDecimal As Double) As String
    Dim iThousands, iHundreds, iTens, iOnes As Integer
    Dim sReturnStr As String
    Dim sHunds(9) As String
    Dim sTens(9) As String
    Dim sOnes(9) As String
    
    sHunds(1) = "C"
    sHunds(2) = "CC"
    sHunds(3) = "CCC"
    sHunds(4) = "CD"
    sHunds(5) = "D"
    sHunds(6) = "DC"
    sHunds(7) = "DCC"
    sHunds(8) = "DCCC"
    sHunds(9) = "CM"
    sTens(1) = "X"
    sTens(2) = "XX"
    sTens(3) = "XXX"
    sTens(4) = "XL"
    sTens(5) = "L"
    sTens(6) = "LX"
    sTens(7) = "LXX"
    sTens(8) = "LXXX"
    sTens(9) = "XC"
    sOnes(1) = "I"
    sOnes(2) = "II"
    sOnes(3) = "III"
    sOnes(4) = "IV"
    sOnes(5) = "V"
    sOnes(6) = "VI"
    sOnes(7) = "VII"
    sOnes(8) = "VIII"
    sOnes(9) = "IX"
    
    iThousands = (xDecimal - (xDecimal Mod 1000)) / 1000
    xDecimal = xDecimal Mod 1000
    iHundreds = (xDecimal - (xDecimal Mod 100)) / 100
    xDecimal = xDecimal Mod 100
    iTens = (xDecimal - (xDecimal Mod 10)) / 10
    xDecimal = xDecimal Mod 10
    iOnes = xDecimal
    
    sReturnStr = ""
    For iCount = 1 To iThousands
        sReturnStr = sReturnStr + "M"
    Next
    If iHundreds > 0 Then
        sReturnStr = sReturnStr + sHunds(iHundreds)
    End If
    If iTens > 0 Then
        sReturnStr = sReturnStr + sTens(iTens)
    End If
    If iOnes > 0 Then
        sReturnStr = sReturnStr + sOnes(iOnes)
    End If
    GetDecRomanStr = sReturnStr
End Function

Public Function GetRomanDecNum(ByVal sRomanStr As String) As Double
    Dim xDecimal As Double
    Dim sStr As String
        
    sStr = Left(sRomanStr, 1)
    While sStr = "M"
        xDecimal = xDecimal + 1000
        sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
        sStr = Left(sRomanStr, 1)
    Wend
    
    iHunds = 0
    If Left(sRomanStr, 2) = "CM" Then
        iHunds = 9
        sRomanStr = Right(sRomanStr, Len(sRomanStr) - 2)
    Else
        If Left(sRomanStr, 1) = "D" Then
            iHunds = 5
            sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
        Else
            If Left(sRomanStr, 2) = "CD" Then
                iHunds = 4
                sRomanStr = Right(sRomanStr, Len(sRomanStr) - 2)
            End If
        End If
    End If
    If iHunds = 0 Or iHunds = 5 Then
        sStr = Left(sRomanStr, 1)
        While sStr = "C"
            iHunds = iHunds + 1
            sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
            sStr = Left(sRomanStr, 1)
        Wend
    End If
    xDecimal = xDecimal + iHunds * 100
    
    iTens = 0
    If Left(sRomanStr, 2) = "XC" Then
        iTens = 9
        sRomanStr = Right(sRomanStr, Len(sRomanStr) - 2)
    Else
        If Left(sRomanStr, 1) = "L" Then
            iTens = 5
            sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
        Else
            If Left(sRomanStr, 2) = "XL" Then
                iTens = 4
                sRomanStr = Right(sRomanStr, Len(sRomanStr) - 2)
            End If
        End If
    End If
    If iTens = 0 Or iTens = 5 Then
        sStr = Left(sRomanStr, 1)
        While sStr = "X"
            iTens = iTens + 1
            sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
            sStr = Left(sRomanStr, 1)
        Wend
    End If
    xDecimal = xDecimal + iTens * 10
    
    iOnes = 0
    If Left(sRomanStr, 2) = "IX" Then
        iOnes = 9
        sRomanStr = Right(sRomanStr, Len(sRomanStr) - 2)
    Else
        If Left(sRomanStr, 1) = "V" Then
            iOnes = 5
            sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
        Else
            If Left(sRomanStr, 2) = "IV" Then
                iOnes = 4
                sRomanStr = Right(sRomanStr, Len(sRomanStr) - 2)
            End If
        End If
    End If
    If iOnes = 0 Or iOnes = 5 Then
        sStr = Left(sRomanStr, 1)
        While sStr = "I"
            iOnes = iOnes + 1
            sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
            sStr = Left(sRomanStr, 1)
        Wend
    End If
    xDecimal = xDecimal + iOnes

    GetRomanDecNum = xDecimal
End Function

Public Function GetDecHexStr(ByVal xDecimal As Double) As String
    Dim sReturnStr As String
    Dim lQuotient As Long
        
    iRemainder = xDecimal Mod 16
    lQuotient = xDecimal \ 16
    
    While lQuotient > 0
        sReturnStr = sReturnStr + GetHexDigit(iRemainder)
        xDecimal = lQuotient
        iRemainder = xDecimal Mod 16
        lQuotient = xDecimal \ 16
    Wend
    If iRemainder > 0 Then
        sReturnStr = sReturnStr + GetHexDigit(iRemainder)
    End If
    GetDecHexStr = StrReverse(sReturnStr)
    
End Function

Public Function GetHexDigit(ByVal iDigit As Integer) As String
    Dim sReturnStr As String
    
    Select Case iDigit
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
        sReturnStr = CStr(iDigit)
    Case "10"
        sReturnStr = "A"
    Case "11"
        sReturnStr = "B"
    Case "12"
        sReturnStr = "C"
    Case "13"
        sReturnStr = "D"
    Case "14"
        sReturnStr = "E"
    Case "15"
        sReturnStr = "F"
    End Select
    GetHexDigit = sReturnStr

End Function

Public Function GetDecOctStr(ByVal xDecimal As Double) As String

    Dim sReturnStr As String
    Dim lQuotient As Long
        
    iRemainder = xDecimal Mod 8
    lQuotient = xDecimal \ 8
    
    Do While lQuotient > 0
        sReturnStr = sReturnStr + CStr(iRemainder)
        xDecimal = lQuotient
        iRemainder = xDecimal Mod 8
        lQuotient = xDecimal \ 8
    Loop
    If iRemainder > 0 Then
        sReturnStr = sReturnStr + CStr(iRemainder)
    End If
    GetDecOctStr = StrReverse(sReturnStr)
    
End Function

Public Function GetDecBinStr(ByVal xDecimal As Double) As String

    Dim sReturnStr As String
    Dim lQuotient As Long
        
    iRemainder = xDecimal Mod 2
    lQuotient = xDecimal \ 2
    
    Do While lQuotient > 0
        sReturnStr = sReturnStr + CStr(iRemainder)
        xDecimal = lQuotient
        iRemainder = xDecimal Mod 2
        lQuotient = xDecimal \ 2
    Loop
    If iRemainder > 0 Then
        sReturnStr = sReturnStr + CStr(iRemainder)
    End If
    GetDecBinStr = StrReverse(sReturnStr)
    
End Function
Public Function GetBinDecNum(ByVal sBinStr As String) As Double
    Dim iLength, Counter As Integer
    Dim xReturnVal As Double
    Dim sNewString As String
        
    xReturnVal = 0
    iLength = Len(sBinStr)
    sNewString = StrReverse(sBinStr)
    For Counter = 0 To iLength - 1
        xReturnVal = xReturnVal + Left(sNewString, 1) * 2 ^ Counter
        sNewString = Right(sNewString, Len(sNewString) - 1)
    Next
    
    GetBinDecNum = xReturnVal

End Function

Public Sub ConvertOutputWindowText(ByVal iOldRadix As Integer, ByVal iNewRadix As Integer)
    Select Case (iOldRadix)
        Case ROMANNUM
            Select Case (iNewRadix)
                Case HEXNUM
                    xDecimal = GetRomanDecNum(OutputWindow.Text)
                    OutputWindow.Text = GetDecHexStr(xDecimal)
                Case DECNUM
                    xDecimal = GetRomanDecNum(OutputWindow.Text)
                    OutputWindow.Text = xDecimal
                Case OCTNUM
                    xDecimal = GetRomanDecNum(OutputWindow.Text)
                    OutputWindow.Text = GetDecOctStr(xDecimal)
                Case BINNUM
                    xDecimal = GetRomanDecNum(OutputWindow.Text)
                    OutputWindow.Text = GetDecBinStr(xDecimal)
            End Select
        Case HEXNUM
            Select Case (iNewRadix)
                Case ROMANNUM
                    xDecimal = Val("&H" + OutputWindow.Text)
                    OutputWindow.Text = GetDecRomanStr(xDecimal)
                Case DECNUM
                    OutputWindow.Text = Val("&H" + OutputWindow.Text)
                Case OCTNUM
                    OutputWindow.Text = Val("&H" + OutputWindow.Text)
                    OutputWindow.Text = GetDecOctStr(OutputWindow.Text)
                Case BINNUM
                    OutputWindow.Text = Val("&H" + OutputWindow.Text)
                    OutputWindow.Text = GetDecBinStr(OutputWindow.Text)
            End Select
        Case DECNUM
            Select Case (iNewRadix)
                Case ROMANNUM
                    OutputWindow.Text = GetDecRomanStr(OutputWindow.Text)
                Case HEXNUM
                    OutputWindow.Text = GetDecHexStr(OutputWindow.Text)
                Case OCTNUM
                    OutputWindow.Text = GetDecOctStr(OutputWindow.Text)
                Case BINNUM
                    OutputWindow.Text = GetDecBinStr(OutputWindow.Text)
            End Select
        Case OCTNUM
            Select Case (iNewRadix)
                Case ROMANNUM
                    xDecimal = Val("&O" + OutputWindow.Text)
                    OutputWindow.Text = GetDecRomanStr(xDecimal)
                Case HEXNUM
                    OutputWindow.Text = Val("&O" + OutputWindow.Text)
                    OutputWindow.Text = GetDecHexStr(OutputWindow.Text)
                Case DECNUM
                    OutputWindow.Text = Val("&O" + OutputWindow.Text)
                Case BINNUM
                    OutputWindow.Text = Val("&O" + OutputWindow.Text)
                    OutputWindow.Text = GetDecBinStr(OutputWindow.Text)
            End Select
        Case BINNUM
            Select Case (iNewRadix)
                Case ROMANNUM
                    xDecimal = GetBinDecNum(OutputWindow.Text)
                    OutputWindow.Text = GetDecRomanStr(xDecimal)
                Case HEXNUM
                    OutputWindow.Text = GetBinDecNum(OutputWindow.Text)
                    OutputWindow.Text = GetDecHexStr(OutputWindow.Text)
                Case DECNUM
                    OutputWindow.Text = GetBinDecNum(OutputWindow.Text)
                Case OCTNUM
                    OutputWindow.Text = GetBinDecNum(OutputWindow.Text)
                    OutputWindow.Text = GetDecOctStr(OutputWindow.Text)
            End Select
    End Select
End Sub

Private Function Calculate_Total(dFirstNum, dSecondNum, dTotal) As Boolean
    Select Case iMathFunction
    Case DIVIDE
        If dSecondNum = 0 Then
            OutputWindow.Text = "ERROR - Divide by Zero"
            OutputWindow.ForeColor = &HFF&
            Calculate_Total = True
        Else
            dTotal = dFirstNum / dSecondNum
            Select Case iCurrentRadix
                Case ROMANNUM
                    OutputWindow.Text = GetDecRomanStr(dTotal)
                Case HEXNUM
                    OutputWindow.Text = GetDecHexStr(dTotal)
                Case DECNUM
                    OutputWindow.Text = dTotal
                Case OCTNUM
                    OutputWindow.Text = GetDecOctStr(dTotal)
                Case BINNUM
                    OutputWindow.Text = GetDecBinStr(dTotal)
            End Select
            
            xFirstNum = dTotal
            bFirstNum = True
            bEntered = True
        End If
        iMathFunction = 0
    Case MULTIPLY
        dTotal = dFirstNum * dSecondNum
        Select Case iCurrentRadix
            Case ROMANNUM
                OutputWindow.Text = GetDecRomanStr(dTotal)
            Case HEXNUM
                OutputWindow.Text = GetDecHexStr(dTotal)
            Case DECNUM
                OutputWindow.Text = dTotal
            Case OCTNUM
                OutputWindow.Text = GetDecOctStr(dTotal)
            Case BINNUM
                OutputWindow.Text = GetDecBinStr(dTotal)
        End Select
        
        dFirstNum = dTotal
        bFirstNum = True
        bEntered = True
        iMathFunction = 0
    Case PLUS
        dTotal = dFirstNum + dSecondNum
        Select Case iCurrentRadix
            Case ROMANNUM
                OutputWindow.Text = GetDecRomanStr(dTotal)
            Case HEXNUM
                OutputWindow.Text = GetDecHexStr(dTotal)
            Case DECNUM
                OutputWindow.Text = dTotal
            Case OCTNUM
                OutputWindow.Text = GetDecOctStr(dTotal)
            Case BINNUM
                OutputWindow.Text = GetDecBinStr(dTotal)
        End Select
        
        dFirstNum = dTotal
        bFirstNum = True
        bEntered = True
        iMathFunction = 0
    Case MINUS
        dTotal = dFirstNum - dSecondNum
        Select Case iCurrentRadix
            Case ROMANNUM
                OutputWindow.Text = GetDecRomanStr(dTotal)
            Case HEXNUM
                OutputWindow.Text = GetDecHexStr(dTotal)
            Case DECNUM
                OutputWindow.Text = dTotal
            Case OCTNUM
                OutputWindow.Text = GetDecOctStr(dTotal)
            Case BINNUM
                OutputWindow.Text = GetDecBinStr(dTotal)
        End Select
        xFirstNum = dTotal
        bFirstNum = True
        bEntered = True
        iMathFunction = 0
    Case POWER
        xTotal = 1
        If dSecondNum <> 0 Then
            For iCounter = 1 To dSecondNum
                dTotal = dTotal * dFirstNum
            Next
        End If
        Select Case iCurrentRadix
            Case ROMANNUM
                OutputWindow.Text = GetDecRomanStr(dTotal)
            Case HEXNUM
                OutputWindow.Text = GetDecHexStr(dTotal)
            Case DECNUM
                OutputWindow.Text = dTotal
            Case OCTNUM
                OutputWindow.Text = GetDecOctStr(dTotal)
            Case BINNUM
                OutputWindow.Text = GetDecBinStr(dTotal)
        End Select
        xFirstNum = dTotal
        bFirstNum = True
        bEntered = True
        iMathFunction = 0
    End Select
End Function

Sub SetupCommandButtons()
    SetupStandardButton ecbButton(ButtonA), pbButton(ButtonA), "A"
    DrawCommandButton ecbButton(ButtonA), pbButton(ButtonA)

    SetupStandardButton ecbButton(ButtonB), pbButton(ButtonB), "B"
    DrawCommandButton ecbButton(ButtonB), pbButton(ButtonB)

    SetupStandardButton ecbButton(ButtonC), pbButton(ButtonC), "C"
    DrawCommandButton ecbButton(ButtonC), pbButton(ButtonC)

    SetupStandardButton ecbButton(ButtonD), pbButton(ButtonD), "D"
    DrawCommandButton ecbButton(ButtonD), pbButton(ButtonD)

    SetupStandardButton ecbButton(ButtonE), pbButton(ButtonE), "E"
    DrawCommandButton ecbButton(ButtonE), pbButton(ButtonE)

    SetupStandardButton ecbButton(ButtonF), pbButton(ButtonF), "F"
    DrawCommandButton ecbButton(ButtonF), pbButton(ButtonF)

    SetupStandardButton ecbButton(ButtonM), pbButton(ButtonM), "M"
    DrawCommandButton ecbButton(ButtonM), pbButton(ButtonM)

    SetupStandardButton ecbButton(ButtonL), pbButton(ButtonL), "L"
    DrawCommandButton ecbButton(ButtonL), pbButton(ButtonL)

    SetupStandardButton ecbButton(ButtonX), pbButton(ButtonX), "X"
    DrawCommandButton ecbButton(ButtonX), pbButton(ButtonX)

    SetupStandardButton ecbButton(ButtonV), pbButton(ButtonV), "V"
    DrawCommandButton ecbButton(ButtonV), pbButton(ButtonV)

    SetupStandardButton ecbButton(ButtonI), pbButton(ButtonI), "I"
    DrawCommandButton ecbButton(ButtonI), pbButton(ButtonI)

    SetupStandardButton ecbButton(ButtonClear), pbButton(ButtonClear), "C"
    DrawCommandButton ecbButton(ButtonClear), pbButton(ButtonClear)

    SetupStandardButton ecbButton(ButtonSine), pbButton(ButtonSine), "Sin"
    DrawCommandButton ecbButton(ButtonSine), pbButton(ButtonSine)

    SetupStandardButton ecbButton(ButtonCosine), pbButton(ButtonCosine), "Cos"
    DrawCommandButton ecbButton(ButtonCosine), pbButton(ButtonCosine)

    SetupStandardButton ecbButton(ButtonTangent), pbButton(ButtonTangent), "Tan"
    DrawCommandButton ecbButton(ButtonTangent), pbButton(ButtonTangent)

    SetupStandardButton ecbButton(ButtonCE), pbButton(ButtonCE), "CE"
    DrawCommandButton ecbButton(ButtonCE), pbButton(ButtonCE)

    SetupStandardButton ecbButton(ButtonPI), pbButton(ButtonPI), "PI"
    DrawCommandButton ecbButton(ButtonPI), pbButton(ButtonPI)

    SetupStandardButton ecbButton(ButtonOzToGram), pbButton(ButtonOzToGram), "Oz-Gm"
    DrawCommandButton ecbButton(ButtonOzToGram), pbButton(ButtonOzToGram)

    SetupStandardButton ecbButton(ButtonFactorial), pbButton(ButtonFactorial), "X!"
    DrawCommandButton ecbButton(ButtonFactorial), pbButton(ButtonFactorial)

    SetupStandardButton ecbButton(ButtonSqrRoot), pbButton(ButtonSqrRoot), "Sqr"
    DrawCommandButton ecbButton(ButtonSqrRoot), pbButton(ButtonSqrRoot)

    SetupStandardButton ecbButton(Button1OverX), pbButton(Button1OverX), "1/X"
    DrawCommandButton ecbButton(Button1OverX), pbButton(Button1OverX)

    SetupStandardButton ecbButton(ButtonBS), pbButton(ButtonBS), "BS"
    DrawCommandButton ecbButton(ButtonBS), pbButton(ButtonBS)

    SetupStandardButton ecbButton(ButtonMC), pbButton(ButtonMC), "MC"
    DrawCommandButton ecbButton(ButtonMC), pbButton(ButtonMC)

    SetupStandardButton ecbButton(ButtonLbToKilo), pbButton(ButtonLbToKilo), "Lb-Kg"
    DrawCommandButton ecbButton(ButtonLbToKilo), pbButton(ButtonLbToKilo)

    SetupStandardButton ecbButton(ButtonSquare), pbButton(ButtonSquare), "X^2"
    DrawCommandButton ecbButton(ButtonSquare), pbButton(ButtonSquare)

    SetupStandardButton ecbButton(ButtonCube), pbButton(ButtonCube), "X^3"
    DrawCommandButton ecbButton(ButtonCube), pbButton(ButtonCube)

    SetupStandardButton ecbButton(ButtonPower), pbButton(ButtonPower), "X^Y"
    DrawCommandButton ecbButton(ButtonPower), pbButton(ButtonPower)

    SetupStandardButton ecbButton(ButtonDivide), pbButton(ButtonDivide), "/"
    DrawCommandButton ecbButton(ButtonDivide), pbButton(ButtonDivide)

    SetupStandardButton ecbButton(ButtonMR), pbButton(ButtonMR), "MR"
    DrawCommandButton ecbButton(ButtonMR), pbButton(ButtonMR)

    SetupStandardButton ecbButton(ButtonGallonToLitre), pbButton(ButtonGallonToLitre), "Gal-Ltr"
    DrawCommandButton ecbButton(ButtonGallonToLitre), pbButton(ButtonGallonToLitre)

    SetupStandardButton ecbButton(Button7), pbButton(Button7), "7"
    DrawCommandButton ecbButton(Button7), pbButton(Button7)

    SetupStandardButton ecbButton(Button8), pbButton(Button8), "8"
    DrawCommandButton ecbButton(Button8), pbButton(Button8)

    SetupStandardButton ecbButton(Button9), pbButton(Button9), "9"
    DrawCommandButton ecbButton(Button9), pbButton(Button9)

    SetupStandardButton ecbButton(ButtonMultiply), pbButton(ButtonMultiply), "X"
    DrawCommandButton ecbButton(ButtonMultiply), pbButton(ButtonMultiply)

    SetupStandardButton ecbButton(ButtonMS), pbButton(ButtonMS), "MS"
    DrawCommandButton ecbButton(ButtonMS), pbButton(ButtonMS)

    SetupStandardButton ecbButton(ButtonMileToKilo), pbButton(ButtonMileToKilo), "Mi-Km"
    DrawCommandButton ecbButton(ButtonMileToKilo), pbButton(ButtonMileToKilo)

    SetupStandardButton ecbButton(Button4), pbButton(Button4), "4"
    DrawCommandButton ecbButton(Button4), pbButton(Button4)

    SetupStandardButton ecbButton(Button5), pbButton(Button5), "5"
    DrawCommandButton ecbButton(Button5), pbButton(Button5)

    SetupStandardButton ecbButton(Button6), pbButton(Button6), "6"
    DrawCommandButton ecbButton(Button6), pbButton(Button6)

    SetupStandardButton ecbButton(ButtonMinus), pbButton(ButtonMinus), "-"
    ecbButton(ButtonMinus).Font.iSize = 16
    DrawCommandButton ecbButton(ButtonMinus), pbButton(ButtonMinus)

    SetupStandardButton ecbButton(ButtonMPlus), pbButton(ButtonMPlus), "M+"
    DrawCommandButton ecbButton(ButtonMPlus), pbButton(ButtonMPlus)

    SetupStandardButton ecbButton(ButtonInchToCent), pbButton(ButtonInchToCent), "In-Cm"
    DrawCommandButton ecbButton(ButtonInchToCent), pbButton(ButtonInchToCent)

    SetupStandardButton ecbButton(Button1), pbButton(Button1), "1"
    DrawCommandButton ecbButton(Button1), pbButton(Button1)

    SetupStandardButton ecbButton(Button2), pbButton(Button2), "2"
    DrawCommandButton ecbButton(Button2), pbButton(Button2)

    SetupStandardButton ecbButton(Button3), pbButton(Button3), "3"
    DrawCommandButton ecbButton(Button3), pbButton(Button3)

    SetupStandardButton ecbButton(ButtonPlus), pbButton(ButtonPlus), "+"
    ecbButton(ButtonPlus).Font.iSize = 12
    DrawCommandButton ecbButton(ButtonPlus), pbButton(ButtonPlus)

    SetupStandardButton ecbButton(ButtonMMinus), pbButton(ButtonMMinus), "M-"
    DrawCommandButton ecbButton(ButtonMMinus), pbButton(ButtonMMinus)

    SetupStandardButton ecbButton(ButtonFToC), pbButton(ButtonFToC), "F - C"
    DrawCommandButton ecbButton(ButtonFToC), pbButton(ButtonFToC)

    SetupStandardButton ecbButton(ButtonPlusMinus), pbButton(ButtonPlusMinus), "+/-"
    DrawCommandButton ecbButton(ButtonPlusMinus), pbButton(ButtonPlusMinus)

    SetupStandardButton ecbButton(Button0), pbButton(Button0), "0"
    DrawCommandButton ecbButton(Button0), pbButton(Button0)

    SetupStandardButton ecbButton(ButtonDecimal), pbButton(ButtonDecimal), "."
    ecbButton(ButtonDecimal).Font.iSize = 20
    DrawCommandButton ecbButton(ButtonDecimal), pbButton(ButtonDecimal)

    SetupStandardButton ecbButton(ButtonEquals), pbButton(ButtonEquals), "="
    ecbButton(ButtonEquals).Font.iSize = 12
    DrawCommandButton ecbButton(ButtonEquals), pbButton(ButtonEquals)
End Sub

Private Sub SetupStandardButton(ecbInCmd As ECommandButton, pbIn As PictureBox, sText As String)
    pbIn.AutoRedraw = True                  'Allow creation of picturebox in memory
    pbIn.ScaleMode = 3                      'Set the picturebox scale to pixels
    pbIn.BorderStyle = 0                    'Remove picturebox border

    ecbInCmd.State = 1                      '0 disabled, 1 button up, 2 button down
    ecbInCmd.Bevel = 0                      'Button bevel height
    ecbInCmd.Font.sName = "MS Sans Serif"   'Button font name
    ecbInCmd.Font.iSize = 8                 'Button font size
    ecbInCmd.Font.bBold = True              'Button font bold
    ecbInCmd.Font.bItalic = False           'Button font italic
    ecbInCmd.Font.iUnderline = False        'Button font underline
    ecbInCmd.Font.lColor = QBColor(15)      'Button font color
    ecbInCmd.Text = sText                   'Button text - use the ~ to force line break - multiline must be true
    ecbInCmd.VAlign = "center"              'Button Text Vert Alignment
    ecbInCmd.HAlign = "center"              'Button Text Horz Alignment
    ecbInCmd.bMultiLine = False             'Determines if Text will span multiple lines
    ecbInCmd.bFocus = False                 'Button has focus
End Sub

Private Sub pbButton_DblClick(iIndex As Integer)
    'Double click will always set button to raised position
    ecbButton(iIndex).State = 1
    DrawCommandButton ecbButton(iIndex), pbButton(iIndex)
End Sub

Private Sub pbButton_GotFocus(iIndex As Integer)
    'Draw the button with focus outline
    ecbButton(iIndex).bFocus = True
    DrawCommandButton ecbButton(iIndex), pbButton(iIndex)
End Sub

Private Sub pbButton_LostFocus(iIndex As Integer)
    'Draw button without focus outline
    ecbButton(iIndex).bFocus = False
    DrawCommandButton ecbButton(iIndex), pbButton(iIndex)
End Sub

Private Sub pbButton_MouseMove(iIndex As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Set button state to up or down for a mousedown and then move off/on picturebox
    If Button = 1 And (x < 2 Or x > pbButton(iIndex).ScaleWidth - 3 Or y < 2 Or y > pbButton(iIndex).ScaleHeight - 3) Then
        pbButton(iIndex).Refresh
        ecbButton(iIndex).State = 1
        DrawCommandButton ecbButton(iIndex), pbButton(iIndex)
    ElseIf Button = 1 And (x >= 2 Or x <= pbButton(iIndex).ScaleWidth - 3 Or y >= 2 Or y <= pbButton(iIndex).ScaleHeight - 3) Then
        pbButton(iIndex).Refresh
        ecbButton(iIndex).State = 2
        DrawCommandButton ecbButton(iIndex), pbButton(iIndex)
    End If
End Sub

Private Sub pbButton_MouseDown(iIndex As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Draw button with focus and button down features
    pbButton(iIndex).Refresh
    ecbButton(iIndex).State = 2
    DrawCommandButton ecbButton(iIndex), pbButton(iIndex)
End Sub

Private Sub pbButton_MouseUp(iIndex As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Draw button without focus and button up features
    If ecbButton(iIndex).State = 2 Then
        pbButton(iIndex).Refresh
        ecbButton(iIndex).State = 1
        DrawCommandButton ecbButton(iIndex), pbButton(iIndex)
        Select Case iIndex
            Case ButtonA
                DoButtonA
            Case ButtonB
                DoButtonB
            Case ButtonC
                DoButtonC
            Case ButtonD
                DoButtonD
            Case ButtonE
                DoButtonE
            Case ButtonF
                DoButtonF
            '
            Case ButtonM
                DoButtonM
            Case ButtonL
                DoButtonL
            Case ButtonX
                DoButtonX
            Case ButtonV
                DoButtonV
            Case ButtonI
                DoButtonI
            Case ButtonClear
                DoButtonClear
            '
            Case ButtonSine
                DoButtonSine
            Case ButtonCosine
                DoButtonCosine
            Case ButtonTangent
                DoButtonTangent
            Case ButtonCE
                DoButtonCE
            '
            Case ButtonPI
                DoButtonPI
            Case ButtonOzToGram
                DoButtonOzToGram
            Case ButtonFactorial
                DoButtonFactorial
            Case ButtonSqrRoot
                DoButtonSqrRoot
            Case Button1OverX
                DoButton1OverX
            Case ButtonBS
                DoButtonBS
            '
            Case ButtonMC
                DoButtonMC
            Case ButtonLbToKilo
                DoButtonLbToKilo
            Case ButtonSquare
                DoButtonSquare
            Case ButtonCube
                DoButtonCube
            Case ButtonPower
                DoButtonPower
            Case ButtonDivide
                DoButtonDivide
            '
            Case ButtonMR
                DoButtonMR
            Case ButtonGallonToLitre
                DoButtonGallonToLitre
            Case Button7
                DoButton7
            Case Button8
                DoButton8
            Case Button9
                DoButton9
            Case ButtonMultiply
                DoButtonMultiply
            '
            Case ButtonMS
                DoButtonMS
            Case ButtonMileToKilo
                DoButtonMileToKilo
            Case Button4
                DoButton4
            Case Button5
                DoButton5
            Case Button6
                DoButton6
            Case ButtonMinus
                DoButtonMinus
            '
            Case ButtonMPlus
                DoButtonMPlus
            Case ButtonInchToCent
                DoButtonInchToCent
            Case Button1
                DoButton1
            Case Button2
                DoButton2
            Case Button3
                DoButton3
            Case ButtonPlus
                DoButtonPlus
            '
            Case ButtonMMinus
                DoButtonMMinus
            Case ButtonFToC
                DoButtonFToC
            Case ButtonPlusMinus
                DoButtonPlusMinus
            Case Button0
                DoButton0
            Case ButtonDecimal
                DoButtonDecimal
            Case ButtonEquals
                DoButtonEquals
        End Select
    End If
End Sub

Private Sub DisableButton(ecbIn As ECommandButton, pbIn As PictureBox)
    pbIn.Refresh
    ecbIn.State = 0
    DrawCommandButton ecbIn, pbIn
    pbIn.Enabled = False
End Sub

Private Sub EnableButton(ecbIn As ECommandButton, pbIn As PictureBox)
    pbIn.Refresh
    ecbIn.State = 1
    DrawCommandButton ecbIn, pbIn
    pbIn.Enabled = True
End Sub

