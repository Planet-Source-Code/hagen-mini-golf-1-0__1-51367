VERSION 5.00
Begin VB.Form frmTiles 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   526
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox numberBpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "frmTiles.frx":0000
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   65
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   66
      Left            =   1920
      Picture         =   "frmTiles.frx":07AE
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   64
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   65
      Left            =   1560
      Picture         =   "frmTiles.frx":09D1
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   63
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   64
      Left            =   1200
      Picture         =   "frmTiles.frx":0F50
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   62
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   63
      Left            =   840
      Picture         =   "frmTiles.frx":14CA
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   61
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   62
      Left            =   480
      Picture         =   "frmTiles.frx":1A44
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   60
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   61
      Left            =   120
      Picture         =   "frmTiles.frx":1FC4
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   59
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   60
      Left            =   1920
      Picture         =   "frmTiles.frx":2528
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   58
      Top             =   4560
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   59
      Left            =   1560
      Picture         =   "frmTiles.frx":2AC7
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   57
      Top             =   4560
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   58
      Left            =   1200
      Picture         =   "frmTiles.frx":30BF
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   56
      Top             =   4560
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   57
      Left            =   840
      Picture         =   "frmTiles.frx":36B1
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   55
      Top             =   4560
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   56
      Left            =   480
      Picture         =   "frmTiles.frx":3C81
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   54
      Top             =   4560
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   55
      Left            =   120
      Picture         =   "frmTiles.frx":4257
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   53
      Top             =   4560
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   54
      Left            =   1920
      Picture         =   "frmTiles.frx":47E7
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   52
      Top             =   4200
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   53
      Left            =   1560
      Picture         =   "frmTiles.frx":4D6E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   51
      Top             =   4200
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   52
      Left            =   1200
      Picture         =   "frmTiles.frx":5366
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   50
      Top             =   4200
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   51
      Left            =   840
      Picture         =   "frmTiles.frx":595A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   49
      Top             =   4200
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   50
      Left            =   480
      Picture         =   "frmTiles.frx":5F56
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   48
      Top             =   4200
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   49
      Left            =   120
      Picture         =   "frmTiles.frx":654E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   47
      Top             =   4200
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   48
      Left            =   1920
      Picture         =   "frmTiles.frx":6ADE
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   46
      Top             =   3840
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   47
      Left            =   1560
      Picture         =   "frmTiles.frx":708B
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   45
      Top             =   3840
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   46
      Left            =   1200
      Picture         =   "frmTiles.frx":7687
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   44
      Top             =   3840
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   45
      Left            =   840
      Picture         =   "frmTiles.frx":7C85
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   43
      Top             =   3840
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   44
      Left            =   480
      Picture         =   "frmTiles.frx":8259
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   42
      Top             =   3840
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   43
      Left            =   120
      Picture         =   "frmTiles.frx":8833
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   41
      Top             =   3840
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   42
      Left            =   1920
      Picture         =   "frmTiles.frx":8DBA
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   40
      Top             =   3480
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   41
      Left            =   1560
      Picture         =   "frmTiles.frx":9353
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   39
      Top             =   3480
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   40
      Left            =   1200
      Picture         =   "frmTiles.frx":9954
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   38
      Top             =   3480
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   39
      Left            =   840
      Picture         =   "frmTiles.frx":9F50
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   37
      Top             =   3480
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   38
      Left            =   480
      Picture         =   "frmTiles.frx":A518
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   36
      Top             =   3480
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   37
      Left            =   120
      Picture         =   "frmTiles.frx":AAE7
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   35
      Top             =   3480
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   36
      Left            =   1920
      Picture         =   "frmTiles.frx":B06B
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   34
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   35
      Left            =   1560
      Picture         =   "frmTiles.frx":B5D7
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   33
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   34
      Left            =   1200
      Picture         =   "frmTiles.frx":BB59
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   32
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   33
      Left            =   840
      Picture         =   "frmTiles.frx":C0D3
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   31
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   32
      Left            =   480
      Picture         =   "frmTiles.frx":C64B
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   30
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   31
      Left            =   120
      Picture         =   "frmTiles.frx":CBC1
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   29
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox targetpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   3720
      Picture         =   "frmTiles.frx":D11E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   28
      Top             =   480
      Width           =   300
   End
   Begin VB.PictureBox targetpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   3360
      Picture         =   "frmTiles.frx":D610
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   27
      Top             =   480
      Width           =   300
   End
   Begin VB.PictureBox ballpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   1
      Left            =   2520
      Picture         =   "frmTiles.frx":DB02
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   26
      Top             =   120
      Width           =   135
   End
   Begin VB.PictureBox ballpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   0
      Left            =   2280
      Picture         =   "frmTiles.frx":DC40
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   25
      Top             =   120
      Width           =   135
   End
   Begin VB.PictureBox numberpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "frmTiles.frx":DD7E
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   24
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox numberpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   4440
      Picture         =   "frmTiles.frx":E52C
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   23
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox numberpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   3960
      Picture         =   "frmTiles.frx":ECDA
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   22
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox numberpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   3480
      Picture         =   "frmTiles.frx":F488
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   21
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox numberpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   3000
      Picture         =   "frmTiles.frx":FC36
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   20
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox numberpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   2520
      Picture         =   "frmTiles.frx":103E4
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   19
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox numberpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   2040
      Picture         =   "frmTiles.frx":10B92
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   18
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox numberpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1560
      Picture         =   "frmTiles.frx":11340
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   17
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox numberpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1080
      Picture         =   "frmTiles.frx":11AEE
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   16
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox numberpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   600
      Picture         =   "frmTiles.frx":1229C
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   15
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   1320
      Picture         =   "frmTiles.frx":12A4A
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   14
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   480
      Picture         =   "frmTiles.frx":131F8
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   13
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   960
      Picture         =   "frmTiles.frx":139A6
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   1440
      Picture         =   "frmTiles.frx":14154
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   11
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   14
      Left            =   0
      Picture         =   "frmTiles.frx":14902
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   10
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   15
      Left            =   0
      Picture         =   "frmTiles.frx":150B0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   9
      Top             =   1320
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   16
      Left            =   360
      Picture         =   "frmTiles.frx":1585E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   1320
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   17
      Left            =   0
      Picture         =   "frmTiles.frx":1600C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   18
      Left            =   360
      Picture         =   "frmTiles.frx":167BA
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   19
      Left            =   1440
      Picture         =   "frmTiles.frx":16F68
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   20
      Left            =   1080
      Picture         =   "frmTiles.frx":17716
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   21
      Left            =   1440
      Picture         =   "frmTiles.frx":17EC4
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   22
      Left            =   1080
      Picture         =   "frmTiles.frx":18672
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   1320
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   24
      Left            =   720
      Picture         =   "frmTiles.frx":18E20
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox sample 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   23
      Left            =   0
      Picture         =   "frmTiles.frx":195CE
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmTiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
