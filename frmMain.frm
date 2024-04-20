VERSION 5.00
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "XCfgEditor"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   12015
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   615
      ScaleWidth      =   12015
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.TextBox txtSrch 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   120
         Width           =   2055
      End
      Begin VB.Image imgBtn 
         Height          =   360
         Index           =   6
         Left            =   5880
         Picture         =   "frmMain.frx":4781A
         ToolTipText     =   "关于..."
         Top             =   120
         Width           =   360
      End
      Begin VB.Line linBrd 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   5640
         X2              =   5640
         Y1              =   600
         Y2              =   0
      End
      Begin VB.Image imgBtn 
         Height          =   360
         Index           =   5
         Left            =   5040
         Picture         =   "frmMain.frx":47A4B
         ToolTipText     =   "查找下一项(支持模式匹配)"
         Top             =   120
         Width           =   360
      End
      Begin VB.Line linBrd 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   2640
         X2              =   2640
         Y1              =   600
         Y2              =   0
      End
      Begin VB.Image imgBtn 
         Height          =   360
         Index           =   4
         Left            =   2160
         Picture         =   "frmMain.frx":47C9B
         ToolTipText     =   "删除项"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image imgBtn 
         Height          =   360
         Index           =   3
         Left            =   1680
         Picture         =   "frmMain.frx":47E44
         ToolTipText     =   "新建项"
         Top             =   120
         Width           =   360
      End
      Begin VB.Line linBrd 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   1560
         X2              =   1560
         Y1              =   600
         Y2              =   0
      End
      Begin VB.Image imgBtn 
         Height          =   360
         Index           =   2
         Left            =   1080
         Picture         =   "frmMain.frx":47EE3
         ToolTipText     =   "保存"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image imgBtn 
         Height          =   360
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":47F6E
         ToolTipText     =   "新建"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image imgBtn 
         Height          =   360
         Index           =   1
         Left            =   600
         Picture         =   "frmMain.frx":47FFE
         ToolTipText     =   "打开"
         Top             =   120
         Width           =   360
      End
      Begin VB.Line linBrd 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFc 
         Caption         =   "新建 XCfg 文件(&N)"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFc 
         Caption         =   "打开(&O)..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFc 
         Caption         =   "重新载入(&R)"
         Index           =   2
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFc 
         Caption         =   "保存(&S)"
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFc 
         Caption         =   "另存为(&A)..."
         Index           =   4
      End
      Begin VB.Menu mnuFc 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFc 
         Caption         =   "退出(&X)"
         Index           =   6
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEc 
         Caption         =   "新建项(&I)"
         Index           =   0
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEc 
         Caption         =   "删除项(&D)"
         Index           =   1
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEc 
         Caption         =   "删除全部(&A)"
         Index           =   2
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuEc 
         Caption         =   "重命名项(&R)"
         Index           =   3
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEc 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuEc 
         Caption         =   "导入内容(&M)"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuEc 
         Caption         =   "导出内容(&E)"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnuEc 
         Caption         =   "用查看器载入(&V)..."
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEc 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuEc 
         Caption         =   "查找下一项(&N)"
         Index           =   9
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "窗口(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWc 
         Caption         =   "关闭当前窗口(&C)"
         Index           =   0
      End
      Begin VB.Menu mnuWc 
         Caption         =   "关闭所有(&A)"
         Index           =   1
      End
      Begin VB.Menu mnuWc 
         Caption         =   "保存所有(&S)"
         Index           =   2
      End
      Begin VB.Menu mnuWc 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuWc 
         Caption         =   "查看器(&V)"
         Index           =   4
         Shortcut        =   {F4}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWc 
         Caption         =   "工具栏"
         Checked         =   -1  'True
         Index           =   5
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "关于 XCfgEditor"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFrmNum As Long

Sub NewFile(sFilePath As String)
    Dim lFrm As New frmShw
    With lFrm
        .Show
        .LoadXCfg sFilePath
    End With
End Sub

Private Sub imgBtn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0, 1
            mnuFc_Click Index
        Case 2
            mnuFc_Click 3
        Case 3, 4
            mnuEc_Click Index - 3
        Case 5
            ActiveForm.SearchItem txtSrch.Text
        Case 6
            mnuAbout_Click
    End Select
End Sub

Private Sub MDIForm_Load()
    Set cdlg = New clsCdlg
    If Command <> "" Then
        Dim sCmdFiles() As String, i As Long
        If InStr(Command, """") <> 0 Then
            sCmdFiles = Split(Command & """ ", """ ")
        Else
            sCmdFiles = Split(Command & " ", " ")
        End If
        For i = 0 To UBound(sCmdFiles)
            If sCmdFiles(i) <> "" Then
                NewFile Replace(sCmdFiles(i), """", "")
            End If
        Next i
    Else
        mnuFc_Click 0
    End If
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    linBrd(0).X2 = Me.ScaleWidth
    imgBtn(6).Move Me.ScaleWidth - imgBtn(6).Width - 240
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Set cdlg = Nothing
End Sub

Private Sub mnuAbout_Click()
    ShellAbout Me.hwnd, App.Title, App.LegalCopyright, Me.Icon
End Sub

Private Sub mnuEc_Click(Index As Integer)
    Dim i As Long
    With ActiveForm
        Select Case Index
            Case 0
                .NewItem
            Case 1
                .RemoveItem
            Case 2
                If MsgBox("在删除全部项目之前，您需要确认此操作。", 48 + vbYesNo, "删除全部项目") = vbYes Then
                    .RemoveItem True
                End If
            Case 3
                .txtName.SetFocus
                .txtName.SelStart = 0
                .txtName.SelLength = Len(.txtName.Text)
            Case 5
                '''
            Case 6
                '''
            Case 7
                'Viewer Code
            Case 9
                imgBtn_MouseUp 5, 0, 0, 0, 0
        End Select
    End With
End Sub

Private Sub mnuFc_Click(Index As Integer)
    With ActiveForm
        Select Case Index
            Case 0
                Dim nFrm As New frmShw
                lFrmNum = lFrmNum + 1
                nFrm.Caption = "未标题-" & lFrmNum
                nFrm.Show
                nFrm.SetFocus
            Case 1
                cdlg.ShowOpen Me.hwnd, sFltr, "打开 XCfg"
                If cdlg.FileName = "" Then Exit Sub
                NewFile cdlg.FileName
                cdlg.FileName = ""
            Case 2
                .LoadXCfg "", True
            Case 3
                .SaveXCfg
            Case 4
                .SaveXCfg True
            Case 6
                Unload Me
        End Select
    End With
End Sub

Private Sub mnuWc_Click(Index As Integer)
    Dim Frms As Form
    Select Case Index
        Case 0
            Unload ActiveForm
        Case 1
            For Each Frms In VB.Forms
                If Frms.Name <> "frmMain" Then Unload Frms
            Next
            lFrmNum = 0
        Case 2
            For Each Frms In VB.Forms
                If Frms.Name <> "frmMain" Then Frms.SaveXCfg
            Next
        Case 4
            '''
        Case 5
            mnuWc(5).Checked = Not mnuWc(5).Checked
            picBar.Visible = mnuWc(5).Checked
    End Select
    If VB.Forms.Count = 1 Then mnuFc_Click 0
End Sub

Private Sub picBar_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    With Data
        If .GetFormat(vbCFFiles) Then
            For i = 1 To .Files.Count
                NewFile .Files.Item(i)
            Next
        End If
    End With
End Sub

Private Sub txtSrch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        imgBtn_MouseUp 5, 0, 0, 0, 0
    End If
End Sub
