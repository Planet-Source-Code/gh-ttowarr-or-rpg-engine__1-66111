VERSION 5.00
Begin VB.UserControl PictureClip 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   132
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   ToolboxBitmap   =   "PictureClip.ctx":0000
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picSize 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   1320
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "PictureClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const m_def_Cols = 1
Const m_def_Rows = 1
Dim m_Cols As Integer
Dim m_Rows As Integer

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020

Public Property Get Cols() As Integer
Cols = m_Cols
End Property

Public Property Let Cols(ByVal New_Cols As Integer)
m_Cols = New_Cols
PropertyChanged "Cols"
End Property

Public Property Get Rows() As Integer
Rows = m_Rows
End Property

Public Property Let Rows(ByVal New_Rows As Integer)
m_Rows = New_Rows
PropertyChanged "Rows"
End Property

Private Sub UserControl_InitProperties()
m_Cols = m_def_Cols
m_Rows = m_def_Rows
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set Picture = PropBag.ReadProperty("Picture", Nothing)
m_Cols = PropBag.ReadProperty("Cols", m_def_Cols)
m_Rows = PropBag.ReadProperty("Rows", m_def_Rows)
End Sub

Private Sub UserControl_Resize()
UserControl.picSize.Picture = UserControl.Picture
UserControl.Width = picSize.Width * Screen.TwipsPerPixelX
UserControl.Height = picSize.Height * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Picture", Picture, Nothing)
Call PropBag.WriteProperty("Cols", m_Cols, m_def_Cols)
Call PropBag.WriteProperty("Rows", m_Rows, m_def_Rows)
End Sub

Public Function GetCell(Id As Integer, Optional Mask As Boolean) As IPictureDisp
Dim X, Y As Integer
Dim cellHeight, cellWidth As Single
    
cellWidth = UserControl.ScaleWidth / m_Cols
cellHeight = UserControl.ScaleHeight / m_Rows
    
Y = Int(Id / m_Cols)
X = Id - (Fix(Id / m_Cols) * m_Cols)
    
picBuffer.Width = cellWidth
picBuffer.Height = cellHeight
    
BitBlt UserControl.picBuffer.hDC, 0, 0, cellWidth, cellHeight, IIf(Mask, picMask.hDC, UserControl.hDC), cellWidth * X, cellHeight * Y, SRCCOPY
    
Set GetCell = UserControl.picBuffer.Image
End Function


Public Sub GenerateMask(R As Integer, G As Integer, B As Integer)
Dim Transp, X, Y As Integer
    
UserControl_Resize
    
picMask.BackColor = RGB(R, G, B)
picMask.Cls
picMask.Width = UserControl.ScaleWidth
picMask.Height = UserControl.ScaleHeight
    
Transp = UserControl.Point(0, 0)
    
For Y = 0 To picMask.Height
    For X = 0 To picMask.Width
        If UserControl.Point((X), (Y)) <> Transp Then picMask.PSet ((X), (Y)), vbBlack
    Next X
Next Y
End Sub

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
Set UserControl.Picture = New_Picture
UserControl.picSize.Picture = UserControl.Picture
UserControl.Width = picSize.Width * Screen.TwipsPerPixelX
UserControl.Height = picSize.Height * Screen.TwipsPerPixelY
PropertyChanged "Picture"
End Property

