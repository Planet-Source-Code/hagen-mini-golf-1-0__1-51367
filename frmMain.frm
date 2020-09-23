VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Mini Golf 1.0"
   ClientHeight    =   8670
   ClientLeft      =   105
   ClientTop       =   0
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":0442
   MousePointer    =   99  'Custom
   ScaleHeight     =   578
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picRefresh 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   960
      MouseIcon       =   "frmMain.frx":074C
      ScaleHeight     =   327
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   415
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      DrawWidth       =   40
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   10680
      Picture         =   "frmMain.frx":0A56
      ScaleHeight     =   148
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   65
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox Screen 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   9840
      Picture         =   "frmMain.frx":0E7E
      ScaleHeight     =   1185
      ScaleWidth      =   1425
      TabIndex        =   5
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox ExitButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   9960
      Picture         =   "frmMain.frx":1AEC
      ScaleHeight     =   600
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   6840
      Width           =   1500
   End
   Begin VB.PictureBox ResetButton 
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   360
      Picture         =   "frmMain.frx":4A0E
      ScaleHeight     =   600
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   6840
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox StartButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   9960
      Picture         =   "frmMain.frx":7930
      ScaleHeight     =   600
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   6120
      Width           =   1500
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   7500
      Left            =   240
      MouseIcon       =   "frmMain.frx":A852
      MousePointer    =   99  'Custom
      ScaleHeight     =   498
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   624
      TabIndex        =   0
      Top             =   120
      Width           =   9390
   End
   Begin VB.PictureBox Shotpower 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      DrawWidth       =   40
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   10080
      ScaleHeight     =   298
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   64
      Top             =   240
      Width           =   495
   End
   Begin VB.Line Line6 
      X1              =   104
      X2              =   104
      Y1              =   568
      Y2              =   520
   End
   Begin VB.Label PlayertotalLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   68
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label partotalLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   67
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   66
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Line Line5 
      X1              =   696
      X2              =   696
      Y1              =   568
      Y2              =   520
   End
   Begin VB.Line Line4 
      X1              =   8
      X2              =   784
      Y1              =   552
      Y2              =   552
   End
   Begin VB.Line Line2 
      X1              =   8
      X2              =   784
      Y1              =   568
      Y2              =   568
   End
   Begin VB.Line Line3 
      X1              =   8
      X2              =   784
      Y1              =   520
      Y2              =   520
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   784
      Y1              =   536
      Y2              =   536
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYER"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PAR"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HOLE"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   9960
      TabIndex        =   26
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   9480
      TabIndex        =   25
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   9000
      TabIndex        =   24
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   8520
      TabIndex        =   23
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   8040
      TabIndex        =   22
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   7560
      TabIndex        =   21
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   7080
      TabIndex        =   20
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   6600
      TabIndex        =   19
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6120
      TabIndex        =   18
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5640
      TabIndex        =   17
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   5160
      TabIndex        =   16
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4680
      TabIndex        =   15
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4200
      TabIndex        =   14
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3720
      TabIndex        =   13
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   12
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   11
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label holelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   9
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   9960
      TabIndex        =   44
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   9480
      TabIndex        =   43
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   9000
      TabIndex        =   42
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   8520
      TabIndex        =   41
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   8040
      TabIndex        =   40
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   7560
      TabIndex        =   39
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   7080
      TabIndex        =   38
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   6600
      TabIndex        =   37
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   6120
      TabIndex        =   36
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   5640
      TabIndex        =   35
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5160
      TabIndex        =   34
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4680
      TabIndex        =   33
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4200
      TabIndex        =   32
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3720
      TabIndex        =   31
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3240
      TabIndex        =   30
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   29
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   28
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label parlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   27
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   62
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   9960
      TabIndex        =   61
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   9480
      TabIndex        =   60
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   9000
      TabIndex        =   59
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   8520
      TabIndex        =   58
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   8040
      TabIndex        =   57
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   7560
      TabIndex        =   56
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   7080
      TabIndex        =   55
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   6600
      TabIndex        =   54
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   6120
      TabIndex        =   53
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   5640
      TabIndex        =   52
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5160
      TabIndex        =   51
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4680
      TabIndex        =   50
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4200
      TabIndex        =   49
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3720
      TabIndex        =   48
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3240
      TabIndex        =   47
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   46
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Scribble"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   45
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label bg1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   63
      Top             =   7800
      Width           =   11655
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bpos As Single
Dim test As Boolean
Dim Sdir As Integer
Dim ang As Single
Dim shotball As Boolean
Dim oldSpeed As Single
Dim holeint As Integer


Private Sub ExitButton_Click()
Unload frmTiles
Unload Me
End Sub

Private Sub Form_Load()
'set piboxes
Me.height = 8670
Me.width = 12000

With picRefresh
.left = 0
.top = 0
.width = picMain.width
.height = picMain.height
.left = picMain.left
.top = picMain.top
Ball.left = .width / 2
Ball.top = .height / 2
Aim.X1 = .width / 2
Aim.Y1 = .height / 2
Aim.X2 = .width / 2
Aim.Y2 = .height / 2

End With
With Screen
.left = 0
.top = 0
.width = picMain.width
.height = picMain.height
.left = picMain.left
.top = picMain.top
End With
Target.width = frmTiles.targetpic(0).width
Target.height = frmTiles.targetpic(0).height
Ball.height = frmTiles.ballpic(0).height
Ball.width = frmTiles.ballpic(0).width
Target.left = Aim.X1 + (Target.width / 2)
Target.top = Aim.Y1 + (Target.height / 2)
'game varibles set
pi = 3.14159265358979
holeint = 1
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call moveball
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case 1
DoEvents
If shotball = False Then
    ballx = Ball.left + (Ball.width / 2)
    bally = Ball.top + (Ball.height / 2)
    Bpos = ballPos(ballx, bally)
    shotball = True
    sCount = sCount + 1
    oldSpeed = speed
    shootball (speed)
End If
Case 2
Me.Caption = ballPos(X, Y)
End Select

End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case 0
If shotball = False Then
With Aim
    .X2 = .X1 - (X - .X1)
    .Y2 = .Y1 - (Y - .Y1)
    Target.left = Aim.X2 - (Target.width / 2)
    Target.top = Aim.Y2 - (Target.height / 2)
Call movetarget

End With
End If
    ballx = Ball.left + (Ball.width / 2)
    bally = Ball.top + (Ball.height / 2)

    
   
Case 1
    Bpos = ballPos(ballx, bally)
   
    

Case 2
     

End Select
End Sub

Private Sub shootball(sp As Single)
Dim stp As Single

ExitButton.Visible = False

speed = sp

test = False
DestX = Aim.X2
DestY = Aim.Y2
ballx = (Ball.width / 2) + Ball.left
bally = (Ball.height / 2) + Ball.top
Dim BalPos As Integer
xDif = Abs(ballx - DestX) * 3
yDif = Abs(bally - DestY) * 3

Sdir = shotdir(DestX, DestY)



'find angle the ball is shot
ang = findangle(ballx, bally, DestX, DestY)

stp = (speed * 400) + 150
Do
speed = adjustspeed(stp)

If stp < 0 Then stp = 5

DoEvents

ballx = (Ball.width / 2) + Ball.left
bally = (Ball.height / 2) + Ball.top

ballLastX = ballx
ballLastY = bally

Select Case Sdir
Case 1

Ball.top = Ball.top - (Sin(ang) * speed)
Ball.left = Ball.left + (Cos(ang) * speed)
xf = 1
yf = -1
Case 2
Ball.top = Ball.top - (Sin(ang) * speed)
Ball.left = Ball.left - (Cos(ang) * speed)
xf = -1
yf = -1
Case 3
Ball.top = Ball.top + (Sin(ang) * speed)
Ball.left = Ball.left - (Cos(ang) * speed)
xf = -1
yf = 1
Case 4
Ball.top = Ball.top + (Sin(ang) * speed)
Ball.left = Ball.left + (Cos(ang) * speed)
xf = 1
yf = 1
Case 5
Ball.left = Ball.left + speed
xf = 1
yf = 0
Case 6
Ball.top = Ball.top - speed
xf = 0
yf = -1
Case 7
Ball.left = Ball.left - speed
xf = -1
yf = 0
Case 8
Ball.top = Ball.top + speed
xf = 0
yf = 1
End Select
Call moveball
    
    DestX = DestX + (xf * (Cos(ang) * speed))
    DestY = DestY + (yf * (Sin(ang) * speed))
    ballx = (Ball.width / 2) + Ball.left
    bally = (Ball.height / 2) + Ball.top
    BalPos = ballPos(ballx, bally)

'if ball hits a wall
If motion(BalPos) = "hole" Then
        If checHole(BalPos) = True Then
            Ball.top = tTop(BalPos) + 12.5 - (Ball.height / 2)
            Ball.left = tLeft(BalPos) + 12.5 - (Ball.width / 2)
            markScore
            GoTo z:
        End If
End If

If test = False Then
    If motion(ballPos(ballx, bally)) = "f" Then
        changedir Sdir, BalPos
        ballx = ballLastX
        bally = ballLastY
        test = True
    Else
    End If
End If

If motion(BalPos) = "t" Then test = False

stp = stp - 1

Loop Until stp = 1
Call Markshot
speed = oldSpeed
If motion(BalPos) = "f" Then ResetButton.Visible = True

shotball = False

Aim.X1 = (Ball.width / 2) + Ball.left
Aim.Y1 = (Ball.height / 2) + Ball.top
ExitButton.Visible = True
Call moveball
Exit Sub
z:
speed = oldSpeed
shotball = False
'loads next hole
If holeint <> 18 Then
holeint = holeint + 1
Call loadcourse(holeint)
picMain.Enabled = True
picMain.SetFocus
placeball
ExitButton.Visible = True
Else
Screen.Visible = True
ExitButton.Visible = True

End If
End Sub



'2 6 1   |ab
'7 * 5   |cd
'3 8 4

Private Sub changedir(olddir As Integer, bp As Integer)


Select Case olddir
Case 1
    If WallType(bp) = "v" Then
        Sdir = 2
    End If
    
    If WallType(bp) = "h" Then
        Sdir = 4
    End If
    
    If WallType(bp) = "a" Then
        Sdir = 2
    End If
    
    If WallType(bp) = "c" Then
        If ballx < (tLeft(bp) + Ball.width) And bally < tBottom(bp) Then Sdir = 2
        If ballx > (tLeft(bp)) And bally > (tBottom(bp) - (Ball.height)) Then Sdir = 4
    End If
    
    If WallType(bp) = "d" Then
        Sdir = 4
    End If
Case 2
    If WallType(bp) = "v" Then
        Sdir = 1
    End If
    
    If WallType(bp) = "h" Then
        Sdir = 3
    End If

    If WallType(bp) = "b" Then
        Sdir = 1
    End If
    
    If WallType(bp) = "c" Then
        Sdir = 3
    End If
    
    If WallType(bp) = "d" Then
        If ballx > (tRight(bp) - Ball.width) And bally < tBottom(bp) Then Sdir = 1
        If ballx < (tRight(bp)) And bally > (tBottom(bp) - (Ball.height)) Then Sdir = 3
    End If
Case 3
    If WallType(bp) = "v" Then
        Sdir = 4
    End If
    
    If WallType(bp) = "h" Then
        Sdir = 2
    End If
    
    If WallType(bp) = "a" Then
        Sdir = 2
    End If
    
    If WallType(bp) = "b" Then
        If ballx > (tRight(bp) - Ball.width) And bally > tTop(bp) Then Sdir = 4
        If ballx < (tRight(bp)) And bally < (tTop(bp) + (Ball.height)) Then Sdir = 2
    End If
    
    If WallType(bp) = "d" Then
        Sdir = 4
    End If
Case 4
    If WallType(bp) = "v" Then
        Sdir = 3
    End If
    
    If WallType(bp) = "h" Then
        Sdir = 1
    End If

    If WallType(bp) = "a" Then
        If ballx < (tLeft(bp) + Ball.width) And bally > tTop(bp) Then Sdir = 3
        If ballx > (tLeft(bp)) And bally < (tTop(bp) + (Ball.height)) Then Sdir = 1
    End If
    
    If WallType(bp) = "b" Then
        Sdir = 1
    End If
    
    If WallType(bp) = "c" Then
        Sdir = 3
    End If
Case 5
    Sdir = 7
Case 6
    Sdir = 8
Case 7
    Sdir = 5
Case 8
    Sdir = 6
End Select
'find angle the ball is shot


End Sub

Private Function placeball()
With Ball
.left = tLeft(startpos) + 12.5 - (.width / 2)
.top = tTop(startpos) + 12.5 - (.height / 2)
Aim.X1 = .left + (.width / 2)
Aim.Y1 = .top + (.height / 2)
ballx = .left + (.width / 2)
bally = .top + (.height / 2)
End With
Call moveball
End Function



Private Sub Shotpower_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y <= 300 And Y > 0 Then
Call moveball
Select Case Button
    Case 1
        If shotball = False Then
            X1 = Shotpower.width / 2
            Y1 = Shotpower.height
            Y2 = 0
            X2 = -40
            X3 = Shotpower.width + 40
            Shotpower.Line (X1, Y2)-(X1, Y), &H404040
            Shotpower.Line (X1, Y1)-(X1, Y), &H8000&
            Shotpower.Line (X2, Y - 20)-(X3, Y - 20), &H404040
            speed = (Abs(Y - Y1)) / 100
        End If
End Select
End If
End Sub

Private Sub Shotpower_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call moveball
Select Case Button
Case 1
    If shotball = False Then
        X1 = Shotpower.width / 2
        Y1 = Shotpower.height
        Y2 = 0
        X2 = -40
        X3 = Shotpower.width + 40
        Shotpower.Line (X1, Y2)-(X1, Y), &H404040
        Shotpower.Line (X1, Y1)-(X1, Y), &H8000&
        Shotpower.Line (X2, Y - 20)-(X3, Y - 20), &H404040
        speed = (Abs(Y - Y1)) / 100
    End If
End Select
End Sub

Private Sub ResetButton_Click()
placeball
sCount = 0
ResetButton.Visible = False
End Sub



Private Function checHole(bp As Integer) As Boolean
checHole = False
If motion(bp) = "hole" Then
    If speed <= 0.75 Then
        If ballx > (tLeft(bp) + 5) And ballx < (tRight(bp) - 5) And bally > (tTop(bp) + 5) And bally < (tBottom(bp) - 5) Then
            checHole = True
        End If
    End If
End If
End Function



Private Sub StartButton_Click()
Call moveball
speed = 1.5

Shotpower.Line (Shotpower.width / 2, 0)-(Shotpower.width / 2, 150), &H404040
Shotpower.Line (Shotpower.width / 2, Shotpower.height)-(Shotpower.width / 2, 150), &H8000&
Shotpower.Line (-40, 130)-(Shotpower.width + 40, 130), &H404040
'loadhole test stage
Screen.Visible = False
resetscorboard
Call loadcourse(1)
picMain.Enabled = True
picMain.SetFocus
placeball
End Sub

Private Function adjustspeed(stepcount As Single) As Single
If stepcount <= speed * 400 Then
    adjustspeed = (Abs(speed - 0.1))
Else
    adjustspeed = speed
End If

End Function

Private Sub markScore()
If hNUM = 1 Then
    PlayertotalLbl.Caption = sCount
Else
    PlayertotalLbl.Caption = Int(PlayertotalLbl.Caption) + sCount
End If

Scorelbl(hNUM).Caption = sCount
If sCount < holePar(hNUM) Then Scorelbl(hNUM).ForeColor = vbRed
If sCount > holePar(hNUM) Then Scorelbl(hNUM).ForeColor = vbBlue
If sCount = holePar(hNUM) Then Scorelbl(hNUM).ForeColor = vbBlack
End Sub

Private Sub resetscorboard()
For ResetScore = 1 To 18
parlbl(ResetScore).Caption = ""
Scorelbl(ResetScore).Caption = ""
Next
End Sub
