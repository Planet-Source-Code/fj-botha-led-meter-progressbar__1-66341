VERSION 5.00
Begin VB.UserControl MeterControl 
   BackColor       =   &H00000000&
   ClientHeight    =   135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2610
   ScaleHeight     =   135
   ScaleWidth      =   2610
   ToolboxBitmap   =   "MeterControl.ctx":0000
   Begin VB.Shape Led 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   215
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   45
   End
   Begin VB.Label LedClick 
      BackStyle       =   0  'Transparent
      Height          =   215
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "MeterControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Clicked_Value   As Integer
Private Dark_Color      As OLE_COLOR
Private Light_Color     As OLE_COLOR
Public Event LedClick(Value As Integer)

Private Sub LedClick_Click(Index As Integer)
  Dim i As Integer
  
  For i = 0 To 50
    Led(i).FillColor = Dark_Color
  Next i
  If Index = 0 Then Exit Sub
  For i = 0 To Index
    Led(i).FillColor = Light_Color
  Next i
  
  Clicked_Value = Index
  RaiseEvent LedClick(Index * 2)
End Sub

Private Sub UserControl_Initialize()
  Dim i As Integer
  For i = 1 To 50
    Load Led(i)
    With Led(i)
      .Visible = True
      If i = 1 Then
        .Left = 0
      Else
        .Left = Led(i - 1).Left + .Width + 15
        .FillColor = Dark_Color
      End If
      .Top = 0
    End With
  Next
  
  For i = 1 To 50
    Load LedClick(i)
    With LedClick(i)
      .Visible = True
      If i = 1 Then
        .Left = 0
      Else
        .Left = LedClick(i - 1).Left + .Width + 15
      End If
      .Top = 0
    End With
  
  Next i
 
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Dark_Color = PropBag.ReadProperty("DarkColor", &H8400&)
  Light_Color = PropBag.ReadProperty("LightColor", &HFF00&)
  LedClick_Click 0
End Sub

Private Sub UserControl_Resize()
  
  UserControl.Width = 3000
  UserControl.Height = 205
End Sub

Public Property Get Value() As Integer
  Value = Clicked_Value
End Property

Public Property Let Value(Val As Integer)
  
  If Val > 100 Then Exit Property
  LedClick_Click Fix((Fix(Val / 2) / 100) * 100)
    
  Clicked_Value = Val
End Property

Public Property Get DarkColor() As OLE_COLOR
  DarkColor = Dark_Color
End Property

Public Property Let DarkColor(Color As OLE_COLOR)
  Dark_Color = Color
  setDark Color
  LedClick_Click 0
  PropertyChanged DarkColor
End Property

Public Property Get LightColor() As OLE_COLOR
  LightColor = Light_Color
End Property

Public Property Let LightColor(Color As OLE_COLOR)
  Light_Color = Color
  setLight Color
  PropertyChanged LightColor
End Property

Private Sub setLight(Color As OLE_COLOR)
  Dim i As Integer
  
  For i = 1 To Clicked_Value
    Led(i).FillColor = Light_Color
  Next i

End Sub

Private Sub setDark(Color As OLE_COLOR)
  Dim i As Integer
  
  For i = 1 To 50
    Led(i).FillColor = Dark_Color
  Next i

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "DarkColor", Dark_Color, &H8400&
  PropBag.WriteProperty "LightColor", Light_Color, &HFF00&
End Sub
