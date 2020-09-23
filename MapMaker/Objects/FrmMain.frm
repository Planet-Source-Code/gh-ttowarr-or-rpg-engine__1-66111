VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FrmMain"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   512
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   912
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   12240
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   7320
      Width           =   1215
      Begin VB.CheckBox ChkTrans 
         Caption         =   "Object"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   10920
      TabIndex        =   2
      Top             =   7320
      Width           =   1335
   End
   Begin PrjRpg.PictureClip Tiles 
      Left            =   5280
      Top             =   0
      _ExtentX        =   6773
      _ExtentY        =   135467
      Picture         =   "FrmMain.frx":0000
      Cols            =   8
      Rows            =   160
   End
   Begin VB.PictureBox PicTiles 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   76800
      Left            =   9600
      Picture         =   "FrmMain.frx":3C0052
      ScaleHeight     =   5120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   0
      Width           =   3840
      Begin VB.Shape ShpTile 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.VScrollBar Scroll 
      Height          =   7695
      Left            =   13440
      Max             =   144
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape ShpPlace 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Level(0 To 319) As String
Private CurrentTile As Long

Private Sub CmdLoad_Click()
Com.ShowOpen
If Com.FileName = "" Then Exit Sub
LoadLevel Com.FileName
End Sub

Private Sub CmdSave_Click()
Com.ShowSave
SaveLevel Com.FileName
End Sub

Private Sub Form_Click()
On Error GoTo FndErr

Dim TileNumber As Long
Dim TileTop As Long
Dim TileLeft As Long
Dim i As Integer

i = CurrentTile

TileLeft = ShpPlace.Left / 32

TileTop = ShpPlace.Top / 32

TileNumber = TileTop * 20 + TileLeft
If ChkTrans.Value = 0 Then Level(TileNumber) = CurrentTile
If ChkTrans.Value = 1 Then Level(TileNumber) = CurrentTile & "+"

Me.PaintPicture Tiles.GetCell(i, False), TileLeft * 32, TileTop * 32, , , , , , , vbSrcCopy

Me.Caption = TileNumber & " - " & CurrentTile
Exit Sub

FndErr:
End Sub

Private Sub SaveLevel(Name As String)
Dim i As Integer
Dim File As Integer

File = FreeFile

Open Name For Output As File
For i = 0 To 319
    Print #File, Level(i)
Next i
Close File

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TileX As String, TileY As String

TileX = Split((X / 32), ",")(0) * 32
TileY = Split((Y / 32), ",")(0) * 32

ShpPlace.Left = TileX
ShpPlace.Top = TileY

If Button = 1 Then Form_Click
End Sub

Private Sub PicTiles_Click()
Dim TileNumber As Long
Dim TileTop As Long
Dim TileLeft As Long

TileLeft = ShpTile.Left / 32

TileTop = ShpTile.Top / 32

TileNumber = TileTop * 8 + TileLeft
CurrentTile = TileNumber
End Sub

Private Sub PicTiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TileX As String, TileY As String

TileX = Split((X / 32), ",")(0) * 32
TileY = Split((Y / 32), ",")(0) * 32

ShpTile.Left = TileX
ShpTile.Top = TileY
End Sub

Private Sub Scroll_Change()
Me.Caption = Scroll.Value
PicTiles.Top = 0 - Scroll.Value * 32
End Sub

Private Sub LoadLevel(Name As String)
Dim sTemp As String
Dim sLevel As String
Dim i As Integer, CurrentTile As Integer
Dim X As Long
Dim Y As Long

Open Name For Input As #1
sTemp = Input(LOF(1), 1)
Close #1
sTemp = Replace(sTemp, " ", "")
sLevel = Replace(sTemp, vbCrLf, ",")

Me.Cls

For i = 0 To 319
    CurrentTile = Split(sLevel, ",")(i)
    Level(i) = CurrentTile
    Me.PaintPicture Tiles.GetCell(CurrentTile, False), X, Y, , , , , , , vbSrcCopy
    X = X + 32
    If X = 20 * 32 Then
        X = 0
        Y = Y + 32
    End If
Next i
End Sub

