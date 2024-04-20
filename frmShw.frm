VERSION 5.00
Begin VB.Form frmShw 
   Caption         =   "untlt"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9405
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShw.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   9405
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtView 
      Height          =   6615
      Index           =   0
      Left            =   3600
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "frmShw.frx":000C
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.PictureBox picBar 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6915
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   3495
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.PictureBox picTools 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   3495
         TabIndex        =   1
         Top             =   0
         Width           =   3495
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   720
            TabIndex        =   3
            Text            =   "Default"
            Top             =   120
            Width           =   2055
         End
         Begin VB.Line linBrd 
            BorderColor     =   &H80000010&
            Index           =   2
            X1              =   0
            X2              =   3495
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line linBrd 
            BorderColor     =   &H80000010&
            Index           =   1
            X1              =   0
            X2              =   3500
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.ListBox lstItem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmShw.frx":001A
         Left            =   0
         List            =   "frmShw.frx":0021
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   3495
      End
   End
   Begin VB.Line linBrd 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   0
      X2              =   9360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line linBrd 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   3495
      X2              =   3495
      Y1              =   0
      Y2              =   6720
   End
End
Attribute VB_Name = "frmShw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCfgPth As String

Sub NewItem(Optional sTitle As String = "", Optional sContain As String = "")
    With lstItem
        Dim i As Long, sTlt As String
        sTlt = IIf(sTitle <> "", sTitle, "新建项" & Replace(Replace(Now, " ", ","), "/", "-"))
        For i = 0 To .ListCount - 1
            If .List(i) = sTlt Then Beep: Exit Sub
        Next i
        .AddItem sTlt, .ListCount
        .ListIndex = .ListCount - 1
    End With
    Load txtView(txtView.Count)
    Form_Resize
    With txtView(txtView.UBound)
        .Text = IIf(sTitle <> "", sContain, sTlt)
        .Visible = True
        .ZOrder 0
    End With
End Sub

Sub RemoveItem(Optional bAll As Boolean = False)
    Dim i As Long
    With lstItem
        If bAll Then
            .Clear
            For i = 1 To txtView.UBound
                Unload txtView(i)
            Next i
            NewItem
        Else
            If .ListIndex + 1 < txtView.UBound Then
                For i = .ListIndex + 1 To txtView.UBound - 1
                    txtView(i).Text = txtView(i + 1).Text
                Next i
            End If
            Unload txtView(txtView.UBound)
            i = .ListIndex
            .RemoveItem .ListIndex
            If .ListCount = 0 Then NewItem
        End If
        .ListIndex = IIf(i > .ListCount - 1, .ListCount - 1, i)
        lstItem_Click
    End With
End Sub

Sub SearchItem(sSch As String)
    If sSch = "" Or sSch = "*" Then Beep: Exit Sub
    Dim i As Long, bSrched As Boolean
    bSrched = False
    With lstItem
        For i = IIf(.ListIndex = .ListCount - 1, 0, .ListIndex + 1) To .ListCount - 1
            If InStr(LCase(.List(i)), LCase(sSch)) <> 0 Or (.List(i) Like sSch) Then
                .ListIndex = i
                lstItem_Click
                bSrched = True
                Exit For
            End If
        Next i
    End With
    If Not bSrched Then Beep
End Sub

Sub ReInit()
    sCfgPth = ""
    lstItem.Clear
    Dim i As Long
    For i = 1 To txtView.UBound
        Unload txtView(txtView.UBound)
    Next i
    lstItem.AddItem "Default"
    Load txtView(1)
    Form_Resize
    txtView(1).Visible = True
End Sub

Sub LoadXCfg(sPath As String, Optional IsReload As Boolean = False)
    On Error GoTo LoadErr
    If sPath = "" And Not (IsReload) Then Exit Sub
    If Not (IsReload) Then
        sCfgPth = sPath
    End If
    
    Dim pb As New PropertyBag
    Dim bytData() As Byte
    Open sCfgPth For Binary As #1
        If LOF(1) > 0 Then
            ReDim bytData(LOF(1) - 1)
            Get #1, 1, bytData
            pb.Contents = bytData
        End If
    Close #1
    If pb.ReadProperty("/version") = "" Then
        GoTo LoadErr
    ElseIf CLng(pb.ReadProperty("/version")) > lCfgVer Then
        MsgBox "您所打开的 XCfg 文件版本高于当前编辑器所支持的版本！" & vbCrLf & "您需要更高版本的 XCfgEditor 才能打开该文件", 48, "版本不被支持"
        Me.Caption = """" & GetFileName(sCfgPth, True) & """ 打开失败"
        ReInit
    Else
        lstItem.Clear
        Dim i As Long
        For i = 1 To txtView.UBound
            Unload txtView(txtView.UBound)
        Next i
        Dim sIndexs() As String
        sIndexs = Split(pb.ReadProperty("/index"), "/")
        For i = 0 To UBound(sIndexs)
            If sIndexs(i) = "" Then Exit For
            NewItem sIndexs(i), pb.ReadProperty(sIndexs(i))
        Next i
        Me.Caption = GetFileName(sCfgPth, True)
        lstItem.ListIndex = 0
        lstItem_Click
    End If
    
    Exit Sub
LoadErr:
    MsgBox "文件 """ & sPath & """ 不是合法的 XCfg 文件！", 48, "错误"
    Me.Caption = """" & GetFileName(sCfgPth, True) & """ 打开失败"
    ReInit
End Sub

Function SaveXCfg(Optional IsSaveAs As Boolean = False) As Long      'Return value: user clicked "cancel" button=1,else=0
    With cdlg
        .FileName = Me.Caption
        If IsSaveAs Then
            .ShowSave frmMain.hwnd, sFltr, "另存为 XCfg 文件"
        ElseIf sCfgPth = "" Then
            .ShowSave frmMain.hwnd, sFltr, "保存 XCfg 文件"
        End If
        If .FileName = "" Then SaveXCfg = 1: Exit Function
        If sCfgPth = "" And Dir(TrimFileName(.FileName)) <> "" Then
            If MsgBox("文件 """ & TrimFileName(.FileName) & """ 已存在，" & vbCrLf & "是否继续？", 32 + vbYesNo, "保存文件") = vbNo Then Exit Function
        End If
        If sCfgPth = "" Then
            If cdlg.rtnFilter = 1 Then
                sCfgPth = TrimFileName(.FileName)
            Else
                sCfgPth = .FileName
            End If
        End If
        .FileName = ""
    End With
    
    Dim pb As New PropertyBag
    Dim bytData() As Byte
    GetXCfgPack pb
    If Dir(sCfgPth) <> "" Then Kill sCfgPth
    Open sCfgPth For Binary As #1
        bytData = pb.Contents
        Put #1, 1, bytData
    Close #1
    
    Me.Caption = GetFileName(sCfgPth, True)
End Function

Function XCfgCompare(sPath As String) As Boolean
    Dim bytData() As Byte, bytPb() As Byte, pb As New PropertyBag, i As Long
    Open sPath For Binary As #1
        If LOF(1) > 0 Then
            ReDim bytData(LOF(1) - 1)
            Get #1, 1, bytData
        End If
    Close #1
    GetXCfgPack pb
    bytPb = pb.Contents
    If UBound(bytData) <> UBound(bytPb) Then
        XCfgCompare = False
        Exit Function
    Else
        For i = 0 To UBound(bytData)
            If bytData(i) <> bytPb(i) Then
                XCfgCompare = False
                Exit Function
            End If
        Next i
    End If
    XCfgCompare = True
End Function

Function GetXCfgPack(pb As PropertyBag) As Long
    On Error GoTo GetErr
    Dim sIndex As String, i As Long
    With pb
        .WriteProperty "/version", lCfgVer
        For i = 0 To lstItem.ListCount - 1
            .WriteProperty lstItem.List(i), txtView(i + 1).Text
            sIndex = sIndex & lstItem.List(i) & "/"
        Next i
        .WriteProperty "/index", sIndex
    End With
    GetXCfgPack = 0
    Exit Function
GetErr:
    GetXCfgPack = Err.Number
End Function

Private Sub Form_Load()
    Me.Show
    setBorderColor lstItem.hwnd, lstItem.BackColor
    ReInit
    lstItem.ListIndex = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    linBrd(0).Y2 = Me.ScaleHeight
    linBrd(3).X2 = Me.ScaleWidth
    Dim i As Long
    For i = 1 To txtView.UBound
        txtView(i).Move picBar.Width + 120, 120, Me.ScaleWidth - picBar.Width - 240, Me.ScaleHeight - 240
    Next i
    lstItem.Move 0, picTools.Height, picBar.ScaleWidth, Me.ScaleHeight - picTools.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    With lstItem
        If sCfgPth = "" Then
            If .ListCount = 1 And .List(0) = "Default" And txtView(1).Text = txtView(0).Text Then GoTo UnloadEnd
        Else
            If Dir(sCfgPth) <> "" Then
                If XCfgCompare(sCfgPth) Then GoTo UnloadEnd
            Else
                If MsgBox("文件 """ & sCfgPth & """ 已被移动或删除，" & vbCrLf & "现在关闭将失去已编辑的内容，是否继续？", 32 + vbYesNo, "文件未找到") = vbNo Then
                    Cancel = 1
                    GoTo UnloadEnd
                Else
                    GoTo UnloadEnd
                End If
            End If
        End If
    End With
    Dim lMsg As Long
    lMsg = MsgBox("文件 """ & Me.Caption & """ 已发生变化，" & vbCrLf & "是否保存文件？", 32 + vbYesNoCancel, "保存文件")
    If lMsg = vbYes Then
        lMsg = SaveXCfg
        If lMsg = 1 Then Cancel = 1
    ElseIf lMsg = vbCancel Then
        Cancel = 1
    End If
UnloadEnd:
End Sub

Private Sub lstItem_Click()
    On Error Resume Next
    txtName.Text = lstItem.Text
    txtView(lstItem.ListIndex + 1).ZOrder 0
End Sub

Private Sub lstItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu frmMain.mnuEdit
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("/") Or KeyAscii = Asc(" ") Then KeyAscii = 0
    If KeyAscii = 13 Then
        KeyAscii = 0
        With lstItem
            If txtName.Text = "" Then
                txtName.Text = .Text
                txtName.SelStart = Len(txtName.Text)
                Beep
                Exit Sub
            End If
            Dim i As Long
            For i = 0 To .ListCount - 1
                If .List(i) = txtName.Text Then Beep: Exit Sub
            Next i
            .List(.ListIndex) = txtName.Text
        End With
    ElseIf KeyAscii = (vbKeyA And vbKeyControl) Then
        KeyAscii = 0
        With txtName
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub

Private Sub txtName_LostFocus()
    txtName.Text = lstItem.Text
End Sub

Private Sub txtView_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = (vbKeyA And vbKeyControl) Then
        KeyAscii = 0
        With txtView(Index)
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub
