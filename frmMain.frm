VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "两面夹鸡V1.2 - by sysdzw"
   ClientHeight    =   12180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15450
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12180
   ScaleWidth      =   15450
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1440
      TabIndex        =   5
      Text            =   "855"
      Top             =   2535
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1440
      TabIndex        =   4
      Text            =   "400"
      Top             =   2175
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1440
      TabIndex        =   3
      Text            =   "100"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2280
      Top             =   4920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8760
      Top             =   960
   End
   Begin VB.Label lblTip 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "点击开始游戏"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   4440
      TabIndex        =   1
      Tag             =   "0"
      Top             =   3360
      Width           =   7935
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Height          =   1335
      Left            =   120
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "夹块宽度:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2550
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "夹鸡速度:"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2190
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "小鸡速度:"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1815
      Width           =   735
   End
   Begin VB.Shape Shape3 
      Height          =   1455
      Left            =   120
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "【游戏说明】淘气的小鸡们又在走钢丝玩了！你可点击鼠标夹住它们，夹中一只得一分。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   315
      Width           =   3375
   End
   Begin VB.Label lblFen 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "分数: 0"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   11160
      TabIndex        =   0
      Tag             =   "0"
      Top             =   240
      Width           =   2295
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   6360
      Picture         =   "frmMain.frx":08CA
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   570
      Index           =   0
      Left            =   840
      Picture         =   "frmMain.frx":11CE
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   5640
      Picture         =   "frmMain.frx":1AD2
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00F0F0F0&
      BorderStyle     =   4  'Dash-Dot
      X1              =   120
      X2              =   14760
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C000&
      BorderWidth     =   3
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   7680
      Top             =   6000
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   3
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   7680
      Top             =   240
      Width           =   855
   End
   Begin VB.Image imgGlass 
      Height          =   2430
      Left            =   0
      Picture         =   "frmMain.frx":2311
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   16500
   End
   Begin VB.Image imgSky 
      Height          =   4020
      Left            =   0
      Picture         =   "frmMain.frx":84B9B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15690
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================
'名    称：夹死小鸡鸡
'描    述：点击窗体任意地方开始夹小鸡
'使用方法：当小鸡快要通过时迅速点击窗体
'编    程：sysdzw 原创开发，如果有需要对模块扩充或更新的话请邮箱发我一份
'发布日期：2016-07-09
'博    客：http://blog.163.com/sysdzw
'          http://blog.csdn.net/sysdzw
'Email   ：sysdzw@163.com
'QQ      ：171977759
'版    本：V1.0.0   初版                                                        2016-07-09
'==============================================================================================
Option Explicit
Dim isOpen As Boolean
Dim intRnd As Integer
Const MAX_CHICKEN = 10
Dim intAddChickenSpeed As Integer
Dim lngChikenSpeed_Set As Integer
Dim lngStickSpeed_Set As Integer
Dim A() As Byte
Dim B() As Byte
Dim B2() As Byte

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Const SND_ASYNC = &H1 '异步播放，否则就独占播放
Private Const SND_NODEFAULT = &H2 '不使用缺省声音
Private Const SND_MEMORY = &H4 '指向一个内存文件
Private Const SND_FILENAME = &H20000 '指向一个实际文件
Private Const SND_LOOP = &H8 '循环播放
Private Const SND_ALIAS_START = 0 '结束播放

Private Sub Form_Load()
    Dim i!
    Randomize
    For i = 1 To MAX_CHICKEN
        Load Image3(i)
        Image3(i).Visible = True
        Image3(i).Left = Image3(0).Left - (i - 1) * 2000 - Int(Rnd * 3000)
        Image3(i).Top = Image3(0).Top
    Next
    lngChikenSpeed_Set = 100
    lngStickSpeed_Set = 400
    A = LoadResData(1, "CUSTOM")
    B = LoadResData(104, "CUSTOM")
    B2 = LoadResData(103, "CUSTOM")
    sndPlaySound B2(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY Or SND_LOOP
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call startKill
End Sub
Private Sub startKill()
    Dim i!
    If lblTip.Visible = True Then '重初始化位置
        Call initChicken
        lblFen.Tag = 0
        lblFen.Caption = "分数: 0"
        lblTip.Caption = ""
        lblTip.Visible = False
        Timer2.Enabled = True
    End If
    
    isOpen = True
    Timer1.Enabled = True
End Sub
Private Sub initChicken()
    Dim i!
    Image3(0).Move 840, 3720
    Image3(0).Tag = ""
    Image3(0).ZOrder 1
    For i = 1 To MAX_CHICKEN
        Image3(i).Left = Image3(0).Left - (i - 1) * 2000 - Int(Rnd * 3000)
        Image3(i).Top = Image3(0).Top
        Image3(i).Tag = ""
    Next
    intAddChickenSpeed = 0
End Sub

Private Sub Form_Resize()
On Error GoTo err1
    lblFen.Left = Me.ScaleWidth - lblFen.Width - 90
    Line1.X1 = 0
    Line1.Y1 = Me.ScaleHeight \ 2
    Line1.X2 = Me.ScaleWidth
    Line1.Y2 = Line1.Y1
    Image3(0).Top = Line1.Y1 - Image3(0).Height
    Shape1.Move (Me.ScaleWidth - Shape1.Width) \ 2, 90
    Shape2.Move Shape1.Left, Me.ScaleHeight - Shape2.Height - 90
    Shape2.Width = Shape1.Width
    lblTip.Move (Me.ScaleWidth - lblTip.Width) \ 2, (Me.ScaleHeight - lblTip.Height) \ 2
    
    imgGlass.Top = Me.ScaleHeight - imgGlass.Height
    imgGlass.Width = Me.ScaleWidth
    imgSky.Width = Me.ScaleWidth
err1:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (Timer2.Enabled Or lblTip.Caption = "游戏结束") Then sndPlaySound B2(0), SND_ALIAS_START
End Sub

Private Sub imgGlass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call startKill
End Sub

Private Sub imgSky_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call startKill
End Sub

Private Sub lblTip_Click()
    Call startKill
End Sub

Private Sub Text1_Change()
    If Val(Text1.Text) > 0 Then lngChikenSpeed_Set = Val(Text1.Text)
End Sub

Private Sub Text2_Change()
    If Val(Text2.Text) > 0 Then lngStickSpeed_Set = Val(Text2.Text)
End Sub

Private Sub Text3_Change()
    If Val(Text3.Text) > 0 Then
        Shape1.Width = Val(Text3.Text)
        Shape2.Width = Shape1.Width
        Shape1.ZOrder 0
        Call Form_Resize
    End If
End Sub

Private Sub Timer1_Timer()
    Dim i%, isKilled As Boolean
    If isOpen Then
        If Shape1.Top + Shape1.Height < Me.ScaleHeight \ 2 Then Shape1.Top = Shape1.Top + lngStickSpeed_Set
        If Shape2.Top > Me.ScaleHeight \ 2 Then Shape2.Top = Shape2.Top - lngStickSpeed_Set
        If Not (Shape1.Top + Shape1.Height < Me.ScaleHeight \ 2) And Not (Shape2.Top > Me.ScaleHeight \ 2) Then
            For i = 0 To MAX_CHICKEN
                If Image3(i).Tag <> "died" Then
                    If (Image3(i).Left > Shape1.Left And Image3(i).Left < Shape1.Left + Shape1.Width) Or _
                       (Image3(i).Left + Image3(i).Width > Shape1.Left And Image3(i).Left + Image3(i).Width < Shape1.Left + Shape1.Width) Then
                        lblFen.Tag = Val(lblFen.Tag) + 1
                        lblFen.Caption = "分数: " & lblFen.Tag
                        sndPlaySound B(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
    '                    Image3(i).Visible = False
                        Image3(i).Tag = "died"
                        isKilled = True
    '                    Beep
                    End If
                End If
            Next
            If Not isKilled Then sndPlaySound A(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
            isOpen = False
        End If
    ElseIf Not isOpen Then
        If Shape1.Top > lngStickSpeed_Set Then Shape1.Top = Shape1.Top - lngStickSpeed_Set
        If Shape2.Top + Shape2.Height < Me.ScaleHeight - lngStickSpeed_Set Then Shape2.Top = Shape2.Top + lngStickSpeed_Set
        If Not (Shape1.Top > 0) And Not (Shape2.Top + Shape2.Height < Me.ScaleHeight) Then Timer1.Enabled = False
    End If
End Sub

Private Sub Timer2_Timer()
    Dim i%, lngMaxHeight&
    Randomize
    For i = 0 To MAX_CHICKEN
        If Image3(i).Tag <> "died" Then
            If Image3(i).Left + Image3(i).Width \ 2 < Shape1.Left - 3000 Then
                Image3(i).Left = Image3(i).Left + 20 + Int(50 * Rnd) + intAddChickenSpeed
            Else '接近时加大速度
                Image3(i).Left = Image3(i).Left + lngChikenSpeed_Set + Int(300 * Rnd) + intAddChickenSpeed
            End If
            
            intRnd = Int(500 * Rnd)
            If intRnd Mod 2 = 0 Then
                Image3(i).Top = Me.ScaleHeight \ 2 - Image3(i).Height + intRnd
                Set Image3(i).Picture = Image2.Picture
            Else
                Image3(i).Top = Me.ScaleHeight \ 2 - Image3(i).Height - intRnd
                Set Image3(i).Picture = Image4.Picture
            End If
        ElseIf Image3(i).Tag = "died" Then
            Image3(i).Top = Image3(i).Top + 200
        End If
        '设置最低水平线
        If Image3(i).Top + Image3(i).Height > lngMaxHeight And Image3(i).Tag <> "died" Then lngMaxHeight = Image3(i).Top + Image3(i).Height
        If i = MAX_CHICKEN Then
            Line1.Y1 = lngMaxHeight
'            Line1.Y2 = lngMaxHeight
            lngMaxHeight = 0
        End If

        If Image3(i).Left + Image3(i).Width \ 2 > Shape1.Left + Shape1.Width And Image3(i).Tag <> "died" Then
            lblTip.Caption = "游戏结束"
            lblTip.Visible = True
            Timer1.Enabled = False
            intAddChickenSpeed = 2000
            Me.Enabled = False
        End If

        If lblTip.Caption = "游戏结束" And Image3(MAX_CHICKEN).Left > Me.ScaleWidth Then
            Me.Enabled = True
            Timer2.Enabled = False
            Shake
        End If
        
        If Image3(MAX_CHICKEN).Tag = "died" Then '开始下一轮
            Call initChicken
'            lblTip.Caption = "您获得了满分！"
'            lblTip.Visible = True
'            Timer1.Enabled = False
'            Timer2.Enabled = False
        End If
    Next
End Sub

Private Sub Shake(Optional ByVal shakeRepeats = 30, Optional ByVal shakePads = 200, Optional ByVal shakeInterval = 25)
    If Me.WindowState <> 0 Then Exit Sub
    Dim i%
    For i = 1 To shakeRepeats
        If i Mod 2 = 0 Then
            Me.Left = Me.Left + shakePads
            Sleep 10
            Me.Top = Me.Top + shakePads
        Else
            Me.Left = Me.Left - shakePads
            Sleep 10
            Me.Top = Me.Top - shakePads
        End If
        Sleep shakeInterval
    Next
End Sub
