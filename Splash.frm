VERSION 5.00
Begin VB.Form Splash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2400
      Top             =   2760
   End
   Begin VB.PictureBox PB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6030
      Left            =   0
      Picture         =   "Splash.frx":0000
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6030
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PauseCount As Integer
Private Sub Form_Load()
Me.ScaleMode = vbPixels
With PB
     .ScaleMode = vbPixels
     .AutoRedraw = True
     .width = 400 ' works best if height = width
     .height = 400
End With
PauseCount = 1
End Sub

Private Sub Timer1_Timer()
If PauseCount = 1 Then LaserDraw3 PB
PauseCount = PauseCount + 1
If PauseCount = 10 Then
Main.Visible = True
Splash.Visible = False
Unload Me
End If
End Sub

Private Function LaserDraw3(PB As PictureBox)
Dim PC(8) As String ' color varible'
Dim vPx, vPy, Px(8), Py(8), LPx, LPy As Integer ' vPx and vPy = constant growth with limitions of picture box width and height , Px(8) and Py(8) = cordinates for the laser to draw , LPx and LPy = the base point of the laser
Dim pbvar, pbx, pby As Single ' pbvar = growth varible , pbx and pby = picturebox width and height


'---------
'|1\2|3/4|   direction the 8 lasers shoot
'|-------|
'|8/7|6\5|
'---------

With PB
    pbx = .width    'loads picture box width and height into varible
    pby = .height
End With
LPx = pbx / 2 'loads laser base position into varible
LPy = pby / 2
pbvar = 0   'sets growth varible to starting value

For vPx = pbvar To LPx 'vPx = Growth varible to width of picture box
Me.Caption = "Laser Draw" & " = " & (100 * pbvar) / LPx & "%"
DoEvents
For vPy = pbvar To LPy 'vPy = Growth varible to height of picture box

Px(1) = vPx                     'sets xy postions to draw for lasers
Px(2) = vPy
Px(3) = Abs(vPy - pby)
Px(4) = Abs((2 * LPx) - pbvar)
Px(5) = Abs((2 * LPy) - vPx)
Px(6) = Abs(vPy - pby)
Px(7) = Abs((2 * lby) - vPy)
Px(8) = vPx

Py(1) = vPy
Py(2) = vPx
Py(3) = vPx
Py(4) = vPy
Py(5) = Abs(vPy - pby)
Py(6) = Abs((2 * LPy) - pbvar)
Py(7) = Abs((2 * LPy) - pbvar)
Py(8) = Abs(vPy - pby)

    PC(1) = PB.Point(Px(1), Py(1)) 'loads color varible at defined xy point
    PC(2) = PB.Point(Px(2), Py(2))
    PC(3) = PB.Point(Px(3), Py(3))
    PC(4) = PB.Point(Px(4), Py(4))
    PC(5) = PB.Point(Px(5), Py(5))
    PC(6) = PB.Point(Px(6), Py(6))
    PC(7) = PB.Point(Px(7), Py(7))
    PC(8) = PB.Point(Px(8), Py(8))
    
'vert                                       'the 8 lasers
    Line (LPx, LPy)-(Px(1), Py(1)), PC(1)
    Line (LPx, LPy)-(Px(4), Py(4)), PC(4)
    Line (LPx, LPy)-(Px(5), Py(5)), PC(5)
    Line (LPx, LPy)-(Px(8), Py(8)), PC(8)
'horz
    Line (LPx, LPy)-(Px(2), Py(2)), PC(2)
    Line (LPx, LPy)-(Px(3), Py(3)), PC(3)
    Line (LPx, LPy)-(Px(6), Py(6)), PC(6)
    Line (LPx, LPy)-(Px(7), Py(7)), PC(7)

Next vPy
pbvar = pbvar + 1 ' increase growth varible by one
Next vPx
Me.Caption = "Laser Draw"
End Function
