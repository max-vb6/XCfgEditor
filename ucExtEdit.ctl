VERSION 5.00
Begin VB.UserControl ucExtEdit 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtEdit 
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "ucExtEdit.ctx":0000
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "ucExtEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ExtEditMode
    ExtTextEditor = 0
    ExtHexEditor = 1
End Enum

Dim bytContains() As Byte, lMode As ExtEditMode

Private Sub UserControl_Initialize()
    Erase bytContains()
    lMode = ExtTextEditor
End Sub

Public Property Get EditMode() As ExtEditMode
    EditMode = lMode
End Property

Public Property Let EditMode(ByVal lSetMode As ExtEditMode)
    PropertyChanged "EditMode"
    lMode = lSetMode
End Property

Private Sub UserControl_Resize()
    With UserControl
        txtEdit.Move 0, 0, .ScaleWidth, .ScaleHeight
    End With
End Sub
