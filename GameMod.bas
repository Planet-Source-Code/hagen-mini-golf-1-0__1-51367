Attribute VB_Name = "GameMod"
Type UserType
    left As Single
    top As Single
    width As Single
    height As Single
End Type
Global Ball As UserType

Type UserType2
    left As Single
    top As Single
    width As Single
    height As Single
End Type
Global Target As UserType2

Type UserType3
    Y1 As Single
    Y2 As Single
    X1 As Single
    X2 As Single
End Type
Global Aim As UserType3

Global blocpic(499) As Integer
Global motion(499) As String
Global WallType(499) As String
Global filePath As String
Global tTop(499) As Single
Global tBottom(499) As Single
Global tRight(499) As Single
Global tLeft(499) As Single
Global holePar(18) As Integer
Global ballx As Single
Global bally As Single
Global DestX As Single
Global DestY As Single
Global ballLastX As Single
Global ballLastY As Single
Global startpos As Integer
Global hNUM As Integer
Global speed As Single
Global pi As Double
Global sCount As Integer
'for Blt Bit
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal animX As Long, ByVal animY As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020   'Copies the source over the destination
Public Const SRCINVERT = &H660046 'Copies and inverts the source over the destination
Public Const SRCAND = &H8800C6    'Adds the source to the destination



Public Sub loadcourse(holenumber As Integer)
sCount = 0

hNUM = holenumber
X = 0
Y = 0
filePath = App.Path & "\Hole" & holenumber & ".hol"
Dim Picnum As Integer
loadpic = -1

Open filePath For Input As #1
Do
Line Input #1, txtstring

If Mid(txtstring, 1, (4)) = "hole" Then GoTo X:
If Mid(txtstring, 1, (3)) = "par" Then
holePar(holenumber) = Mid(txtstring, 5, (1))
If hNUM = 1 Then
    Main.partotalLbl.Caption = holePar(hNUM)
Else
    Main.partotalLbl.Caption = Int(Main.partotalLbl.Caption) + holePar(hNUM)
End If
Main.parlbl(hNUM).Caption = holePar(holenumber)
GoTo X:
End If
loadpic = loadpic + 1

Picnum = Val(Mid(txtstring, 1, (2)))

a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.sample(Picnum).hDC, 0, 0, SRCCOPY)
'Main.picRefresh.PSet (X, Y)
'Main.picRefresh.Print loadpic
'Main.picRefresh.PSet (X, Y)
'Main.picRefresh.Line -(X, Y), RGB(0, 0, 0)

tTop(loadpic) = Y
tBottom(loadpic) = Y + 25
tLeft(loadpic) = X
tRight(loadpic) = X + 25


If Len(txtstring) < 7 Then
motion(loadpic) = Mid(txtstring, 4, (1))
WallType(loadpic) = Right(txtstring, 1)
If WallType(loadpic) = "s" Then startpos = loadpic
Else
motion(loadpic) = "hole"
End If
X = X + 25
If X >= 25 * 25 Then
X = 0
Y = Y + 25
End If
X:
Loop Until loadpic = 499
Close #1
a = BitBlt(Main.picMain.hDC, 0, 0, Main.picMain.width, Main.picMain.height, Main.picRefresh.hDC, 0, 0, SRCCOPY)
Main.picMain.Refresh
fillSign
End Sub



Function ballPos(balx As Single, baly As Single) As Single
For ballposition = 0 To 499
DoEvents
If balx >= tLeft(ballposition) And balx <= tRight(ballposition) Then

If baly >= tTop(ballposition) And baly <= tBottom(ballposition) Then

ballPos = ballposition
Exit Function
End If
End If
Next ballposition

End Function

Function findangle(bx As Single, by As Single, dx As Single, dy As Single)
xb = Abs(bx - dx)
yb = Abs(by - dy)
If xb = 0 Or yb = 0 Then Exit Function
findangle = Atn(yb / xb) ' * 180) / pi
End Function
Function shotdir(mx As Single, my As Single)
If mx < ballx And my < bally Then shotdir = 2
If mx < ballx And my > bally Then shotdir = 3
If mx > ballx And my < bally Then shotdir = 1
If mx > ballx And my > bally Then shotdir = 4

If mx > ballx And my = bally Then shotdir = 5
If Abs(mx - ballx) <= 1 And Abs(my - bally) <= 1 Then shotdir = 6
If mx < ballx And Abs(my - bally) <= 1 Then shotdir = 7
If Abs(mx - ballx) <= 1 And my > bally Then shotdir = 8
End Function
Function findangle2(angle As Single)
Dim ang As Single
Dim ang2 As Single
ang = (angle * 180) / pi
ang2 = Abs(ang - 90)
findangle2 = (ang2 * pi) / 180
End Function

Public Sub Markshot()
a = BitBlt(Main.picMain.hDC, 0, 0, Main.picMain.width, Main.picMain.height, Main.picRefresh.hDC, 0, 0, SRCCOPY)


If sCount < 10 Then
X = tLeft(104)
Y = tTop(104)
a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.sample(58).hDC, 0, 0, SRCCOPY)
a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.numberpic(sCount).hDC, 0, 0, SRCAND)
a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.numberBpic(0).hDC, 0, 0, SRCINVERT)
Else
sC1 = Int(left(sCount, 1))
sC2 = Int(Right(sCount, 1))
Xa = tLeft(104)
Ya = tTop(104)
X = tLeft(105) - 8
Y = tTop(105)
a = BitBlt(Main.picRefresh.hDC, X + 8, Y, 25, 25, frmTiles.sample(59).hDC, 0, 0, SRCCOPY)
a = BitBlt(Main.picRefresh.hDC, Xa, Ya, 25, 25, frmTiles.sample(58).hDC, 0, 0, SRCCOPY)

a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.numberpic(sC2).hDC, 0, 0, SRCAND)
a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.numberBpic(0).hDC, 0, 0, SRCINVERT)

a = BitBlt(Main.picRefresh.hDC, Xa, Ya, 25, 25, frmTiles.numberpic(sC1).hDC, 0, 0, SRCAND)
a = BitBlt(Main.picRefresh.hDC, Xa, Ya, 25, 25, frmTiles.numberBpic(0).hDC, 0, 0, SRCINVERT)

End If
a = BitBlt(Main.picMain.hDC, 0, 0, Main.picMain.width, Main.picMain.height, Main.picRefresh.hDC, 0, 0, SRCCOPY)
Main.picMain.Refresh
End Sub

Public Sub moveball()
a = BitBlt(Main.picMain.hDC, 0, 0, Main.picMain.width, Main.picMain.height, Main.picRefresh.hDC, 0, 0, SRCCOPY)

a = BitBlt(Main.picMain.hDC, Ball.left, Ball.top, 9, 9, frmTiles.ballpic(1).hDC, 0, 0, SRCAND)
a = BitBlt(Main.picMain.hDC, Ball.left, Ball.top, 9, 9, frmTiles.ballpic(0).hDC, 0, 0, SRCINVERT)


Main.picMain.Refresh

End Sub

Public Sub movetarget()
a = BitBlt(Main.picMain.hDC, 0, 0, Main.picMain.width, Main.picMain.height, Main.picRefresh.hDC, 0, 0, SRCCOPY)

a = BitBlt(Main.picMain.hDC, Target.left, Target.top, 20, 20, frmTiles.targetpic(1).hDC, 0, 0, SRCAND)
a = BitBlt(Main.picMain.hDC, Target.left, Target.top, 20, 20, frmTiles.targetpic(0).hDC, 0, 0, SRCINVERT)
a = BitBlt(Main.picMain.hDC, Ball.left, Ball.top, 9, 9, frmTiles.ballpic(1).hDC, 0, 0, SRCAND)
a = BitBlt(Main.picMain.hDC, Ball.left, Ball.top, 9, 9, frmTiles.ballpic(0).hDC, 0, 0, SRCINVERT)


Main.picMain.Refresh

End Sub

Public Sub fillSign()
a = BitBlt(Main.picMain.hDC, 0, 0, Main.picMain.width, Main.picMain.height, Main.picRefresh.hDC, 0, 0, SRCCOPY)


If hNUM < 10 Then
X = tLeft(29)
Y = tTop(29)
a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.sample(40).hDC, 0, 0, SRCCOPY)
a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.numberpic(hNUM).hDC, 0, 0, SRCAND)
a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.numberBpic(0).hDC, 0, 0, SRCINVERT)
Else
sC1 = Int(left(hNUM, 1))
sC2 = Int(Right(hNUM, 1))

X = tLeft(30) - 10
Y = tTop(30)
Xa = tLeft(29)
Ya = tTop(29)

a = BitBlt(Main.picRefresh.hDC, X + 10, Y, 25, 25, frmTiles.sample(41).hDC, 0, 0, SRCCOPY)
a = BitBlt(Main.picRefresh.hDC, Xa, Ya, 25, 25, frmTiles.sample(40).hDC, 0, 0, SRCCOPY)

a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.numberpic(sC2).hDC, 0, 0, SRCAND)
a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.numberBpic(0).hDC, 0, 0, SRCINVERT)

a = BitBlt(Main.picRefresh.hDC, Xa, Ya, 25, 25, frmTiles.numberpic(sC1).hDC, 0, 0, SRCAND)
a = BitBlt(Main.picRefresh.hDC, Xa, Ya, 25, 25, frmTiles.numberBpic(0).hDC, 0, 0, SRCINVERT)

End If

X = tLeft(54)
Y = tTop(54)
a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.sample(46).hDC, 0, 0, SRCCOPY)
a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.numberpic(holePar(hNUM)).hDC, 0, 0, SRCAND)
a = BitBlt(Main.picRefresh.hDC, X, Y, 25, 25, frmTiles.numberBpic(0).hDC, 0, 0, SRCINVERT)

a = BitBlt(Main.picMain.hDC, 0, 0, Main.picMain.width, Main.picMain.height, Main.picRefresh.hDC, 0, 0, SRCCOPY)
Main.picMain.Refresh
End Sub
