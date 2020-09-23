VERSION 5.00
Begin VB.Form FrmLevel 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   638
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ShpSpace 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   3
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   9615
      TabIndex        =   3
      Top             =   7680
      Width           =   9615
   End
   Begin VB.PictureBox ShpSpace 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   9615
      TabIndex        =   2
      Top             =   -135
      Width           =   9615
   End
   Begin VB.PictureBox ShpSpace 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Index           =   1
      Left            =   9600
      ScaleHeight     =   7695
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   -120
      Width           =   135
   End
   Begin VB.Timer TmrMove 
      Interval        =   200
      Left            =   840
      Top             =   360
   End
   Begin VB.PictureBox ShpSpace 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Index           =   0
      Left            =   -135
      ScaleHeight     =   7695
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   0
      Width           =   135
   End
   Begin PrjRpg.PictureClip CharacterClip 
      Left            =   8280
      Top             =   5520
      _ExtentX        =   1879
      _ExtentY        =   3440
      Picture         =   "FrmLevel.frx":0000
      Cols            =   3
      Rows            =   4
   End
   Begin VB.Timer TmrMovement 
      Interval        =   10
      Left            =   360
      Top             =   360
   End
   Begin PrjRpg.PictureClip Tiles 
      Left            =   9600
      Top             =   0
      _ExtentX        =   6773
      _ExtentY        =   135467
      Picture         =   "FrmLevel.frx":28E2
      Cols            =   8
      Rows            =   160
   End
   Begin VB.Shape Character 
      Height          =   450
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "FrmLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private LevelLoaded As Boolean
Private Counter As Integer

Private Sub LoadLevel(Filename As String)
On Error Resume Next

Dim Level As String
Dim i As Integer, CurrentTile As String
Dim CurrentTileInt As Integer
Dim x As Long
Dim y As Long
Dim sTemp As String

Open Filename For Input As #1
sTemp = Input(LOF(1), 1)
Close #1
sTemp = Replace(sTemp, " ", "")
Level = Replace(sTemp, vbCrLf, ",")

For i = 0 To 319
    CurrentTile = Split(Level, ",")(i)
    If CurrentTile Like "*+" Then
    CurrentTile = Mid(CurrentTile, 1, Len(CurrentTile) - 1)
    Load ShpSpace(ShpSpace.ubound + 1)
    ShpSpace(ShpSpace.ubound).Left = x
    ShpSpace(ShpSpace.ubound).Top = y
    ShpSpace(ShpSpace.ubound).ScaleMode = 3
    ShpSpace(ShpSpace.ubound).Width = 32
    ShpSpace(ShpSpace.ubound).Height = 32
    ShpSpace(ShpSpace.ubound).Visible = False
    End If
    CurrentTileInt = CurrentTile
    Me.PaintPicture Tiles.GetCell(CurrentTileInt, False), x, y, , , , , , , vbSrcCopy
    x = x + 32
    If x = 20 * 32 Then
        x = 0
        y = y + 32
    End If
Next i

SavePicture Me.Image, Filename & ".bmp"
Me.Cls
Me.Picture = LoadPicture(Filename & ".bmp")
Kill Filename & ".bmp"
End Sub

Private Sub Form_Load()
CharacterClip.GenerateMask 255, 255, 255
End Sub

Private Sub TmrMove_Timer()
Counter = Counter + 1
If Counter = 2 Then Counter = 0
End Sub

Private Sub TmrMovement_Timer()
Dim SideToDraw As Integer
Dim Walking As Boolean

If LevelLoaded = False Then LoadLevel App.Path & "\level.level"
LevelLoaded = True

Walking = False

If Not GetAsyncKeyState(vbKeyLeft) = 0 Then
    Me.Cls
    Character.Left = Character.Left - 2
    SideToDraw = 9
    Walking = True
ElseIf Not GetAsyncKeyState(vbKeyRight) = 0 Then
    Me.Cls
    Character.Left = Character.Left + 2
    SideToDraw = 3
    Walking = True
End If
            
If Not GetAsyncKeyState(vbKeyUp) = 0 Then
    Me.Cls
    Character.Top = Character.Top - 2
    SideToDraw = 0
    Walking = True
ElseIf Not GetAsyncKeyState(vbKeyDown) = 0 Then
    Me.Cls
    Character.Top = Character.Top + 2
    SideToDraw = 6
    Walking = True
End If
     
If Walking = True Then
    Me.PaintPicture CharacterClip.GetCell(SideToDraw + Counter, True), Character.Left, Character.Top, , , , , , , vbMergePaint
    Me.PaintPicture CharacterClip.GetCell(SideToDraw + Counter), Character.Left, Character.Top, , , , , , , vbSrcAnd
End If


If Collision = True Then
    If Not GetAsyncKeyState(vbKeyLeft) = 0 Then
        Character.Left = Character.Left + 2
    ElseIf Not GetAsyncKeyState(vbKeyRight) = 0 Then
        Character.Left = Character.Left - 2
    End If
            
    If Not GetAsyncKeyState(vbKeyUp) = 0 Then
        Character.Top = Character.Top + 2
    ElseIf Not GetAsyncKeyState(vbKeyDown) = 0 Then
        Character.Top = Character.Top - 2
    End If
End If

End Sub

Private Function Collision() As Boolean
Dim Number As Integer

For Number = 0 To ShpSpace.ubound
    If Character.Left < ShpSpace(Number).Left + ShpSpace(Number).Width And ShpSpace(Number).Left < Character.Left + Character.Width Then
        If Character.Top < ShpSpace(Number).Top + ShpSpace(Number).Height And ShpSpace(Number).Top < Character.Top + Character.Height Then
            Collision = True
        End If
    End If
Next
End Function




