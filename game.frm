VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Snake!"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
   Icon            =   "game.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   105
   ScaleMode       =   0  'User
   ScaleWidth      =   129.947
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3120
      Top             =   1920
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   0
      ScaleHeight     =   1350
      ScaleWidth      =   2025
      TabIndex        =   0
      Top             =   0
      Width           =   2025
      Begin VB.Label lblRestart 
         BackColor       =   &H00000000&
         Caption         =   "Press Enter To Restart"
         ForeColor       =   &H0000FF00&
         Height          =   210
         Left            =   195
         TabIndex        =   3
         Top             =   570
         Visible         =   0   'False
         Width           =   1620
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3120
      Top             =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "X"
      ForeColor       =   &H0000FF00&
      Height          =   180
      Left            =   1875
      TabIndex        =   2
      Top             =   1365
      Width           =   150
   End
   Begin VB.Label scorelabel 
      BackColor       =   &H00000000&
      Caption         =   "Score:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   15
      TabIndex        =   1
      Top             =   1350
      Width           =   1740
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Dim CurDir As String
Dim Speed As Integer
Dim SnakeLength As Integer
Dim HasMoved As Boolean
Dim AppleLeft As Integer
Dim AppleTop As Integer
Dim Score As Integer
Dim RestartReady As Boolean

Dim SnakePieces(150)

Dim hMemDC As Long 'For the snake
Dim hMemBitmap As Long
Dim hPrevMemBitmap As Long

Dim hMemDC2 As Long 'For the dead snake
Dim hMemBitmap2 As Long
Dim hPrevMemBitmap2 As Long

Dim hMemDC3 As Long 'For the grid
Dim hMemBitmap3 As Long
Dim hPrevMemBitmap3 As Long

Dim BBUFFERDC As Long 'Backbuffer in RAM
Dim hMemBBuffer As Long
Dim hPrevMemBBuffer As Long

Private Sub Form_Load()
Form1.Show
If Right$(App.Path, 1) = "\" Then File = App.Path Else File = App.Path + "\"

BBUFFERDC = CreateCompatibleDC(Form1.hDC)
hMemBBuffer = CreateCompatibleBitmap(Form1.hDC, 135, 90)
hPrevMemBBuffer = SelectObject(BBUFFERDC, hMemBBuffer)

hMemDC = CreateCompatibleDC(Form1.hDC)
hMemBitmap = CreateCompatibleBitmap(Form1.hDC, 9, 9)
hPrevMemBitmap = SelectObject(hMemDC, hMemBitmap)
Picture2.Picture = LoadPicture(File + "snake.bmp")
BitBlt hMemDC, 0, 0, 9, 9, Picture2.hDC, 0, 0, SRCCOPY

hMemDC2 = CreateCompatibleDC(Form1.hDC)
hMemBitmap2 = CreateCompatibleBitmap(Form1.hDC, 9, 9)
hPrevMemBitmap2 = SelectObject(hMemDC2, hMemBitmap2)
Picture2.Picture = LoadPicture(File + "deadsnake.bmp")
BitBlt hMemDC2, 0, 0, 9, 9, Picture2.hDC, 0, 0, SRCCOPY

hMemDC3 = CreateCompatibleDC(Form1.hDC)
hMemBitmap3 = CreateCompatibleBitmap(Form1.hDC, 135, 90)
hPrevMemBitmap3 = SelectObject(hMemDC3, hMemBitmap3)
Picture2.Picture = LoadPicture(File + "grid.bmp")
BitBlt hMemDC3, 0, 0, 135, 90, Picture2.hDC, 0, 0, SRCCOPY


Randomize
Score = -10
CurDir = "down"
Speed = 9 'DONT CHANGE THIS!!!!!!!!
SnakeLength = 4
InitTop = 0
InitLeft = 0
HasMoved = False

AddApple

For i = 0 To SnakeLength
SnakePieces(i) = InitLeft & "," & InitTop
InitTop = InitTop - 9
Next

For i = 0 To SnakeLength
sp = Split(SnakePieces(i), ",")
BitBlt Picture2.hDC, Int(sp(0)), Int(sp(1)), 9, 9, hMemDC, 0, 0, SRCCOPY
Next

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

hMemBitmap = SelectObject(hMemDC, hPrevMemBitmap)
DeleteObject hMemBitmap
DeleteDC hMemDC
hMemBitmap2 = SelectObject(hMemDC2, hPrevMemBitmap2)
DeleteObject hMemBitmap2
DeleteDC hMemDC2
hMemBitmap3 = SelectObject(hMemDC3, hPrevMemBitmap3)
DeleteObject hMemBitmap3
DeleteDC hMemDC3
hMemBBuffer = SelectObject(BBUFFERDC, hPrevMemBBuffer)
DeleteObject hMemBBuffer
DeleteDC BBUFFERDC

End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Sub StartGame()
Score = -10
CurDir = "down"
SnakeLength = 4
InitTop = 0
InitLeft = 0
HasMoved = False

AddApple

For i = 0 To SnakeLength
SnakePieces(i) = InitLeft & "," & InitTop
InitTop = InitTop - 9
Next

For i = 0 To SnakeLength
sp = Split(SnakePieces(i), ",")
BitBlt Picture2.hDC, Int(sp(0)), Int(sp(1)), 9, 9, hMemDC, 0, 0, SRCCOPY
Next
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
If RestartReady And KeyCode = 13 Then
lblRestart.Visible = False
RestartReady = False
StartGame
Timer1.Enabled = True
End If

If HasMoved Then Exit Sub
HasMoved = True
Select Case KeyCode
Case vbKeyLeft
If CurDir <> "right" Then CurDir = "left"
Case vbKeyRight
If CurDir <> "left" Then CurDir = "right"
Case vbKeyDown
If CurDir <> "up" Then CurDir = "down"
Case vbKeyUp
If CurDir <> "down" Then CurDir = "up"
End Select
End Sub



Private Sub Timer1_Timer()

HasMoved = False

SPS = Split(SnakePieces(0), ",")

'--------------------Colission with side
If (SPS(0) - 9) < 0 And CurDir = "left" Then 'Out of boundary boundaries
HighLightSnake
Exit Sub
End If
If (SPS(0) + 9) > 130 And CurDir = "right" Then
HighLightSnake
Exit Sub
End If
If (SPS(1) - 9) < 0 And CurDir = "up" Then
HighLightSnake
Exit Sub
End If
If (SPS(1) + 9) > 81 And CurDir = "down" Then
HighLightSnake
Exit Sub
End If

'------------------------------detect collision with self

chkcoli = Split(SnakePieces(0), ",") 'split it here for easy access

Select Case CurDir

Case "left"
For i2 = 1 To SnakeLength
chkcoli2 = Split(SnakePieces(i2), ",")
If Int(chkcoli(0)) - 9 = Int(chkcoli2(0)) And Int(chkcoli(1)) = Int(chkcoli2(1)) Then
HighLightSnake
Exit Sub
End If
Next i2

Case "right"
For i2 = 1 To SnakeLength
chkcoli2 = Split(SnakePieces(i2), ",")
If Int(chkcoli(0)) + 9 = Int(chkcoli2(0)) And Int(chkcoli(1)) = Int(chkcoli2(1)) Then
HighLightSnake
Exit Sub
End If
Next i2

Case "up"
For i2 = 1 To SnakeLength
chkcoli2 = Split(SnakePieces(i2), ",")
If Int(chkcoli(1)) - 9 = Int(chkcoli2(1)) And Int(chkcoli(0)) = Int(chkcoli2(0)) Then
HighLightSnake
Exit Sub
End If
Next i2

Case "down"
For i2 = 1 To SnakeLength
chkcoli2 = Split(SnakePieces(i2), ",")
If Int(chkcoli(1)) + 9 = Int(chkcoli2(1)) And Int(chkcoli(0)) = Int(chkcoli2(0)) Then
HighLightSnake
Exit Sub
End If
Next i2

End Select

'-----------------------------move the snake
If CurDir = "up" Then
newLeft = Int(SPS(0))
newTop = Int(SPS(1)) - Speed
End If

If CurDir = "left" Then
newLeft = Int(SPS(0)) - Speed
newTop = Int(SPS(1))
End If

If CurDir = "right" Then
newLeft = Int(SPS(0)) + Speed
newTop = Int(SPS(1))
End If

If CurDir = "down" Then
newLeft = Int(SPS(0))
newTop = Int(SPS(1)) + Speed
End If

PCMove = SnakeLength
For i = 1 To UBound(SnakePieces) 'Slip all pieces in array back 1
If Not PCMove > 0 Then Exit For
SnakePieces(PCMove) = SnakePieces(PCMove - 1)
PCMove = PCMove - 1
Next

SnakePieces(0) = newLeft & "," & newTop 'Set the position of the head of the snake
RenderMove 'Move the snake to the arrays positions

AppleColDetect = Split(SnakePieces(0), ",")
If AppleColDetect(0) = AppleLeft And AppleColDetect(1) = AppleTop Then

ACK = Split(SnakePieces(SnakeLength), ",")
SnakeLength = SnakeLength + 1
Select Case CurDir
Case "left"
SnakePieces(SnakeLength) = ACK(0) + 9 & "," & ACK(1)
Case "right"
SnakePieces(SnakeLength) = ACK(0) - 9 & "," & ACK(1)
Case "up"
SnakePieces(SnakeLength) = ACK(0) & "," & ACK(1) + 9
Case "down"
SnakePieces(SnakeLength) = ACK(0) & "," & ACK(1) - 9
End Select

AddApple
End If

End Sub

Sub RenderMove()
BitBlt BBUFFERDC, 0, 0, 135, 90, hMemDC3, 0, 0, SRCCOPY
For i = 0 To SnakeLength
sp = Split(SnakePieces(i), ",")
BitBlt BBUFFERDC, Int(sp(0)), Int(sp(1)), 9, 9, hMemDC, 0, 0, SRCCOPY
Next
BitBlt BBUFFERDC, AppleLeft, AppleTop, 9, 9, hMemDC, 0, 0, SRCCOPY
BitBlt Picture2.hDC, 0, 0, 135, 90, BBUFFERDC, 0, 0, SRCCOPY
End Sub

Sub HighLightSnake()
Timer1.Enabled = False
BitBlt BBUFFERDC, 0, 0, 135, 90, hMemDC3, 0, 0, SRCCOPY
For i = 0 To SnakeLength
sp = Split(SnakePieces(i), ",")
BitBlt BBUFFERDC, Int(sp(0)), Int(sp(1)), 9, 9, hMemDC2, 0, 0, SRCCOPY
BitBlt Picture2.hDC, 0, 0, 135, 90, BBUFFERDC, 0, 0, SRCCOPY
Next
lblRestart.Visible = True
RestartReady = True
End Sub

Sub AddApple()
Score = Score + 10
scorelabel.Caption = "Score: " & Score
AppleLeft = Round(Rnd(1) * 14) * 9
AppleTop = Round(Rnd(1) * 9) * 9
BitBlt BBUFFERDC, AppleLeft, AppleTop, 9, 9, hMemDC, 0, 0, SRCCOPY
End Sub

