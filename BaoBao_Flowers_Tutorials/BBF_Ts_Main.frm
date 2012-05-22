VERSION 5.00
Begin VB.Form BBF_Ts_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BaoBao Flowers Tutorials"
   ClientHeight    =   8250
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11880
   Icon            =   "BBF_Ts_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "BBF_Ts_Main.frx":74F2
   ScaleHeight     =   8250
   ScaleWidth      =   11880
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4860
      Left            =   6255
      Picture         =   "BBF_Ts_Main.frx":149976
      ScaleHeight     =   4860
      ScaleWidth      =   5550
      TabIndex        =   20
      Top             =   540
      Width           =   5550
   End
   Begin VB.PictureBox Step_Bar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1125
      ScaleHeight     =   210
      ScaleWidth      =   7515
      TabIndex        =   18
      Top             =   7290
      Width           =   7545
      Begin VB.PictureBox Step_CT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   15
         TabIndex        =   19
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame"
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   3870
      TabIndex        =   8
      Top             =   3465
      Width           =   7950
      Begin VB.Label Text_Show 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   0
         TabIndex        =   17
         Top             =   3195
         Width           =   90
      End
      Begin VB.Label Text_Show 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   7
         Left            =   0
         MouseIcon       =   "BBF_Ts_Main.frx":16EF08
         TabIndex        =   16
         Top             =   2880
         Width           =   90
      End
      Begin VB.Label Text_Show 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   6
         Left            =   0
         TabIndex        =   15
         Top             =   2520
         Width           =   90
      End
      Begin VB.Label Text_Show 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   0
         MouseIcon       =   "BBF_Ts_Main.frx":16F05A
         TabIndex        =   14
         Top             =   2115
         Width           =   90
      End
      Begin VB.Label Text_Show 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   0
         MouseIcon       =   "BBF_Ts_Main.frx":16F1AC
         TabIndex        =   13
         Top             =   1755
         Width           =   90
      End
      Begin VB.Label Text_Show 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   0
         MouseIcon       =   "BBF_Ts_Main.frx":16F2FE
         TabIndex        =   12
         Top             =   1350
         Width           =   90
      End
      Begin VB.Label Text_Show 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   0
         TabIndex        =   11
         Top             =   945
         Width           =   90
      End
      Begin VB.Label Text_Show 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   540
         Width           =   90
      End
      Begin VB.Label Text_Show 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   135
         Width           =   90
      End
   End
   Begin VB.ListBox List_Combo 
      Appearance      =   0  'Flat
      Height          =   2910
      Left            =   270
      TabIndex        =   2
      Top             =   3150
      Width           =   3120
   End
   Begin VB.DirListBox Dir 
      Height          =   1140
      Left            =   9360
      TabIndex        =   1
      Top             =   630
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Label Label_Info 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "当前位置：首页"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   2
      Left            =   315
      TabIndex        =   21
      Top             =   135
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   3645
      X2              =   3645
      Y1              =   675
      Y2              =   6570
   End
   Begin VB.Label Label_Command 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "下一步"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   2
      Left            =   8910
      MouseIcon       =   "BBF_Ts_Main.frx":16F450
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   7320
      Width           =   540
   End
   Begin VB.Label Label_Command 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "上一步"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   1
      Left            =   360
      MouseIcon       =   "BBF_Ts_Main.frx":16F5A2
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   7320
      Width           =   540
   End
   Begin VB.Label Label_Command 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "开始查看教程"
      ForeColor       =   &H00800000&
      Height          =   180
      Index           =   0
      Left            =   2205
      MouseIcon       =   "BBF_Ts_Main.frx":16F6F4
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   6300
      Width           =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "index"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   9855
      MouseIcon       =   "BBF_Ts_Main.frx":16F846
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   7110
      Width           =   1995
   End
   Begin VB.Label Label_Info 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2007 饱饱花房(BaoBao Flowers)"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   3015
      TabIndex        =   3
      Top             =   7920
      Width           =   3870
   End
   Begin VB.Image Image_Show 
      Height          =   2250
      Left            =   3780
      MouseIcon       =   "BBF_Ts_Main.frx":16F998
      MousePointer    =   99  'Custom
      Top             =   630
      Width           =   3465
   End
   Begin VB.Label Label_Info 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[ 返回首页 ]"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   10350
      MouseIcon       =   "BBF_Ts_Main.frx":16FAEA
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   135
      Width           =   1080
   End
   Begin VB.Image Image_Logo 
      Height          =   2250
      Left            =   90
      Picture         =   "BBF_Ts_Main.frx":16FC3C
      Top             =   675
      Width           =   3465
   End
End
Attribute VB_Name = "BBF_Ts_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FS_Run As Boolean
Dim Text As String
Dim List_Type As Integer
Dim Tutor As String
Dim Step_All As Integer
Dim Move_PC As Long
Dim PassLove As Integer


Private Sub Form_GotFocus()
List_Combo.SetFocus
End Sub

Private Sub CD_Check_f()
MsgBox "正版校验出错，您可能正在使用盗版的“饱饱纸藤花教程”。" & vbCrLf & vbCrLf & "请支持正版软件，谢谢。"
End
End Sub

Private Sub Form_Load()
Dim Str_Check As String
Dim EP_OBJ As Variant
On Error Resume Next
Dim FSO As New FileSystemObject

If FS_Run = False Then
   If App.PrevInstance = True Then End


    Open App.Path + "\SyxnX.db" For Input As #1
    While Not EOF(1)
    Line Input #1, Text
    If Len(Text) > 0 Then Str_Check = RTrim(Text)
    Wend
    Close #1
 
 Select Case Str_Check
 
 Case "hjghjjuyugygbgjgygyguytrvgvgftrhjgjhgjhghjgytytcgcfdruuicxertyvbnvxerwweqwrtyruiuyiobnbncgtfgjhtgbnvyyghghvvryutiyuibnbvhytghjh"
 If FSO.FileExists(Left(App.Path, 3) & "饱饱纸藤花教程\Flowers_bbft_a.db") = False Or FSO.FileExists(Left(App.Path, 3) & "饱饱纸藤花教程\Flowers_bbft_b.db") = False Then CD_Check_f
 Set EP_OBJ = FSO.GetFile(App.Path + "\Flowers_bbft_a.db")
 If EP_OBJ.Size < 1500000000# Then CD_Check_f
 Set EP_OBJ = FSO.GetFile(App.Path + "\Flowers_bbft_a.db")
 If EP_OBJ.Size < 1500000000# Then CD_Check_f
 Case "hjghjjuyu978766545367869hknmbgbnvcrd656gtvr5x3243675viky9otyctolou;pytvytrye5667;pp'[y7vt6r5ceduhgutvt65r653435437776b987098byu"
 
 
 Case Else
 CD_Check_f
 
 End Select
   
   
   
   
   
   
   
   
   If FSO.FileExists(FSO.GetSpecialFolder(SystemFolder) & "\VB6CHS.DLL") = False Then
        FSO.CopyFile App.Path & "\VB6CHS.DLL", FSO.GetSpecialFolder(SystemFolder) & "\VB6CHS.DLL", True
        MsgBox "自动优化已完成，请重新运行 饱饱纸藤花教程！", vbInformation
        End
   End If
End If

Set FSO = Nothing

Dim i As Integer
Dim i_find As Integer
Dir.Path = App.Path + "\纸藤花教程"

List_Combo.Clear
   
   For i = 0 To Dir.ListCount - 1
    If UCase(Right(Dir.List(i), 5)) = ".BBFT" Then
      For i_find = Len(Dir.List(i)) - 6 To 3 Step -1
         If Mid(Dir.List(i), i_find, 1) = "\" Then
               List_Combo.AddItem Mid(Dir.List(i), i_find + 1, Len(Dir.List(i)) - i_find - 5)
            Exit For
         End If
      Next
    End If
   Next
   
List_Type = 1

List_Combo.ListIndex = 0
List_Combo_Click

Image_Show.Visible = False
For i = 0 To 8
Text_Show(i).Left = 0
Text_Show(i).Caption = ""
If i > 0 Then Text_Show(i).Top = Text_Show(i - 1).Top + 350
Next

Frame.Top = 3000
Text_Show(8).ForeColor = vbRed
Text_Show(0).ForeColor = vbBlack

Text_Show(3).MousePointer = 99
Text_Show(4).MousePointer = 99
Text_Show(5).MousePointer = 99
Text_Show(7).MousePointer = 99

Text_Show(0).Caption = "饱饱纸藤花教程 版本：" & App.Major & "." & App.Minor
Text_Show(1).Caption = "教程作者：杨小妮(Syxnx)"
Text_Show(2).Caption = "作品摄影：杨小妮(Syxnx)"
Text_Show(3).Caption = "界面设计：黄思夏(Leask)"
Text_Show(4).Caption = "程序开发：黄思夏(Leask)"
Text_Show(5).Caption = "客服电邮：leaskh@gmail.com"
Text_Show(6).Caption = "客服QQ：4251-2174"
Text_Show(7).Caption = "想了解更多纸藤花品种请浏览 饱饱花房 官方相册 http://picasaweb.google.com/syxnix/"
Text_Show(8).Caption = "注意：本电子教程的教程资源及程序软件均由 饱饱花房 原创并拥有版权，严禁非法传播、分享。"
Picture1.Visible = True
Label_Info(2).Caption = "当前位置：首页"
Label_Info(2).ForeColor = vbWhite

End Sub










Private Sub Form_Unload(Cancel As Integer)
End
End Sub



Private Sub Image_Show_Click()
Label2_MouseUp 1, 0, 0, 0
End Sub

Private Sub Label_Command_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
  If Label_Info(2).ForeColor = vbRed Then
     MsgBox "尊敬的客户：" & vbCrLf & "很抱歉，您尚未购买此教程。" & vbCrLf & "如想购买更多精彩的纸藤花教程，请联系 饱饱花房。" & vbCrLf & "谢谢支持！", vbInformation
  Exit Sub
End If
    List_Combo.Clear
    Open App.Path + "\纸藤花教程\" + Tutor + ".bbft\Tutorials.bbftx" For Input As #1
    While Not EOF(1)
    Line Input #1, Text
    If UCase(Left(Text, 4)) = "STP:" Then
    List_Combo.AddItem RTrim(Right(Text, Len(Text) - 4))
    End If
    Wend
    Close #1
Step_All = List_Combo.ListCount
List_Type = 2

List_Combo.ListIndex = 0
List_Combo_Click

Case 1
List_Combo.ListIndex = List_Combo.ListIndex - 1
List_Combo_Click
Case 2
List_Combo.ListIndex = List_Combo.ListIndex + 1
List_Combo_Click
End Select

End Sub


Private Sub ShowImg(FileName As String, Size As Integer)
On Error Resume Next
Dim ww, hh As Double
Image_Show.Visible = False
Image_Show.Stretch = False
Image_Show.Picture = LoadPicture(FileName)
ww = Image_Show.Width
hh = Image_Show.Height
Select Case Size
Case 0
Image_Logo.Visible = False
Image_Show.Stretch = False
Line1.X1 = 7100
Line1.X2 = 7100
Image_Show.Move (7100 - Image_Show.Width) / 2, 520
Frame.Left = 7300
Frame.Top = 2800
Label_Command(0).Visible = False
Label_Command(1).Caption = "上一步"
Label_Command(2).Caption = "下一步"
List_Combo.Left = 12000
Case 1
Image_Logo.Visible = True
Image_Show.Move 3780, 630
Image_Show.Height = 2800
Image_Show.Width = ww * 2800 / hh
Image_Show.Stretch = True
Line1.X1 = 3646
Line1.X2 = 3646
Frame.Top = 3465
Frame.Left = 3870
Label_Command(0).Visible = True
Label_Command(1).Caption = "上一种"
Label_Command(2).Caption = "下一种"
List_Combo.Left = 270
End Select
Image_Show.Visible = True
End Sub



Private Sub Label_Info_Click(Index As Integer)
Select Case Index
Case 0
Form_Load
End Select
End Sub










Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case List_Type
Case 1
Label_Command_Click 0
Case 2
   Select Case Button
      Case 1
         If Label_Command(2).Enabled = True Then Label_Command_Click 2
      Case 2
         If Label_Command(1).Enabled = True Then Label_Command_Click 1
      Case 4
          Label_Info_Click 0
   End Select
End Select
End Sub

Private Sub List_Combo_Click()
On Error Resume Next
Dim i As Integer
Dim il As Integer
Dim br As Boolean
Dim Text_Temp As String
 Dim FSO As New FileSystemObject

Text_Show(0).Top = 0
Picture1.Visible = False
For i = 0 To 8
Text_Show(i).Left = 0
Text_Show(i).Caption = ""
If i > 0 Then Text_Show(i).Top = Text_Show(i - 1).Top + 350
Next


Text_Show(3).MousePointer = 0
Text_Show(4).MousePointer = 0
Text_Show(5).MousePointer = 0
Text_Show(7).MousePointer = 0



Frame.Height = Text_Show(8).Height + Text_Show(8).Top

Select Case List_Type
Case 1


ShowImg App.Path + "\纸藤花教程\" + List_Combo + ".bbft\cover.bbfp", 1
    Open App.Path + "\纸藤花教程\" + List_Combo + ".bbft\Tutorials.bbftx" For Input As #1
    While Not EOF(1)
    Line Input #1, Text
    Select Case UCase(Left(Text, 4))
    Case "FLR:"
    
    
         Select Case FSO.FileExists(App.Path + "\纸藤花教程\" + List_Combo + ".bbft\1.bbfp")
    
              Case True
                 Text_Show(0).Caption = "花名：" & RTrim(Right(Text, Len(Text) - 4))
                   Text_Show(0).ForeColor = vbBlack
                   Label_Command(0).Enabled = True
                 Case False
                 Text_Show(0).Caption = "花名：" & RTrim(Right(Text, Len(Text) - 4)) & "（您尚未购买此教程）"
                   Text_Show(0).ForeColor = vbRed
                    Label_Command(0).Enabled = False
                 End Select
     
     
     
     
     
     Case "SAY:"
     br = False
     For i = 2 To Len(RTrim(Right(Text, Len(Text) - 4))) - 4
      If UCase(Mid(RTrim(Right(Text, Len(Text) - 4)), i, 4)) = "<BR>" Then
        br = True
        Exit For
      End If
     Next
     Select Case br
     Case True
     Text_Show(1).Caption = "花语：" & Left(RTrim(Right(Text, Len(Text) - 4)), i - 1)
     Text_Show(2).Caption = "      " & Mid(RTrim(Right(Text, Len(Text) - 4)), i + 4, Len(RTrim(Right(Text, Len(Text) - 4))) - i - 3)
     Case False
     Text_Show(1).Caption = "花语：" & RTrim(Right(Text, Len(Text) - 4))
     Text_Show(2).Caption = ""
    End Select
        

    Case "VOL:"
    Text_Show(3).Caption = "版本：" & RTrim(Right(Text, Len(Text) - 4))
    Case "DAT:"
    Text_Show(4).Caption = "日期：" & RTrim(Right(Text, Len(Text) - 4))
    Case "ATH:"
    Text_Show(5).Caption = "作者：" & RTrim(Right(Text, Len(Text) - 4))
        Case "DSN:"
    Text_Show(6).Caption = "设计：" & RTrim(Right(Text, Len(Text) - 4))
    Case "CPR:"
    Text_Show(7).Caption = "版权：" & RTrim(Right(Text, Len(Text) - 4))
    Case "BSO:"
    Text_Show(8).Caption = "备注：本教程依赖 " & RTrim(Right(Text, Len(Text) - 4)) & " 作为基础，请先学习基础课程后再学习本教程。"
    Text_Show(8).ForeColor = vbRed
    
    
    
    
    End Select
    Wend
    Close #1
    

Select Case FSO.FileExists(App.Path + "\纸藤花教程\" + List_Combo + ".bbft\1.bbfp")
    
    Case True
          Label_Info(2).Caption = "正在浏览：" & Tutor
          Label_Info(2).ForeColor = vbWhite
    Case False
          Label_Info(2).Caption = "正在浏览：" & Tutor & "（您尚未购买此教程）"
          Label_Info(2).ForeColor = vbRed
  End Select
    
Tutor = List_Combo


Case 2
Label_Info(2).Caption = "当前位置：" & Tutor
Label_Info(2).ForeColor = vbWhite
ShowImg App.Path + "\纸藤花教程\" + Tutor + ".bbft\" & (List_Combo.ListIndex + 1) & ".bbfp", 0
il = 0
Text_Temp = List_Combo
Text_Show(0).ForeColor = vbBlack

     For i = 2 To Len(Text_Temp) - 4
      If UCase(Mid(Text_Temp, i, 4)) = "<BR>" Then
         Text_Show(il).Caption = Left(Text_Temp, i - 1)
         Text_Temp = Right(Text_Temp, Len(Text_Temp) - i - 3)
         i = 0
         il = il + 1
      End If
       If i = Len(List_Combo) - 4 Then
             Text_Show(il) = Text_Temp
     End If
     Next
Text_Show(8).ForeColor = vbBlack
End Select

Label_Command(1).Enabled = True
Label_Command(2).Enabled = True
If List_Combo.ListIndex = 0 Then
Label_Command(1).Enabled = False
Label_Command(2).Enabled = True
End If
If List_Combo.ListIndex = List_Combo.ListCount - 1 Then
Label_Command(1).Enabled = True
Label_Command(2).Enabled = False
End If





Step_CT.Width = Int(Step_Bar.Width * (List_Combo.ListIndex + 1) / List_Combo.ListCount) + 30

Label2.Caption = List_Combo.ListIndex + 1 & "/" & List_Combo.ListCount
Set FSO = Nothing

End Sub


Private Sub List_Combo_DblClick()
Label_Command_Click 0
End Sub

Private Sub List_Combo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
'76 79 86 69
Select Case KeyCode
Case 27
Label_Info_Click 0
Case 13
List_Combo_DblClick
End Select

   Select Case KeyCode

      Case 76
         PassLove = 1
      Case 79
         If PassLove = 1 Then
            PassLove = 2
         Else
            PassLove = 0
         End If
      Case 86
         If PassLove = 2 Then
            PassLove = 3
         Else
            PassLove = 0
         End If
       Case 69
         If PassLove = 3 Then
              MsgBox "小妮：" & vbCrLf & vbCrLf & "宝贝，好爱你。你的每一辈子都是我的！" & vbCrLf & vbCrLf & "——思夏"
                     If MsgBox("宝贝小妮，嫁给我好不好？", vbYesNo) = vbNo Then
                           Do Until MsgBox("不要啦，再想一想嘛？人家都爱死你了！", vbYesNo) = vbYes
                           Loop
                     End If
               MsgBox "嘻嘻，开心死了！亲一个！Zhu～Bo～！"
         End If
        PassLove = 0
         
      Case Else: PassLove = 0
 End Select
End Sub


Private Sub Step_Bar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Step_Bar_MouseMove Button, 0, X, 0
End Sub

Private Sub Step_Bar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim Move_V As Long
If Button <> 1 Then Exit Sub
Select Case X
Case Is <= 0: Move_V = 0
Case Is >= Step_Bar.Width: Move_V = Step_Bar.Width
Case Else: Move_V = X
End Select
Move_PC = Round(Move_V * List_Combo.ListCount / Step_Bar.Width)
If Move_PC < 1 Then Move_PC = 1
If Move_PC > List_Combo.ListCount Then Move_PC = List_Combo.ListCount
Label2.Caption = Move_PC & "/" & List_Combo.ListCount
Step_CT.Width = Int(Step_Bar.Width * Move_PC / List_Combo.ListCount)
End Sub


Private Sub Step_Bar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
List_Combo.ListIndex = Move_PC - 1
List_Combo_Click
End Sub

Private Sub Step_CT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Step_Bar_MouseMove Button, 0, X, 0
End Sub

Private Sub Step_CT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Step_Bar_MouseMove Button, 0, X, 0
End Sub

Private Sub Step_CT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Step_Bar_MouseUp Button, 0, 0, 0
End Sub

Private Sub Text_Show_Click(Index As Integer)
On Error Resume Next
If Label_Info(2).Caption <> "当前位置：首页" Then Exit Sub
Select Case Index
Case 3, 4
Shell "explorer http://honeonet.spaces.live.com/"
Case 5
Shell "explorer mailto:leaskh@gmail.com"
Case 7
Shell "explorer http://picasaweb.google.com/syxnix/"
End Select
End Sub
