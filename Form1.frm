VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "早打卡提醒"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   6720
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "已打卡"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   0
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lastDate As String
Private fs As New CFile
Private filePath As String

Private isWindowShow As Boolean
Private lastMsg As String

Private Sub Command1_Click()
  Dim nowDate As String
  nowDate = Format(Now, "yyyyMMdd")
  fs.OverWriteToTextFile filePath, nowDate
  lastDate = nowDate
End Sub

Private Sub Form_Load()
  Call HideWin
  filePath = App.Path & "\last.txt"
  If fs.FileExists(filePath) Then
    lastDate = fs.ReadTextFile(filePath)
  Else
    lastDate = ""
  End If
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
  Dim nowDate As String
  Dim nowTime As String
  nowDate = Format(Now, "yyyyMMdd")
  nowTime = Format(Now, "HH:mm:ss")
  ActOnTime nowDate, nowTime
End Sub

Public Sub ActOnTime(ByVal nowDate As String, ByVal nowTime As String)
  Dim msg As String
  If lastDate = "" Or nowDate <> lastDate Then
    'check if warn
    If nowTime < "08:30:00" And nowTime > "05:00:00" Then
      msg = "请检查是否已打卡，打卡后今天可以准时下班哦！"
      If Not isWindowShow Then
        Call ShowWin
      End If
    End If
    If nowTime > "08:30:00" And nowTime < "09:30:00" Then
      msg = "请检查是否已打卡！本周可能要补工时了！"
      If Not isWindowShow Then
        Call ShowWin
      End If
    End If
    If nowTime > "09:30:00" Then
      msg = "你完了！申请调休吧！"
      If Not isWindowShow Then
        Call ShowWin
      End If
    End If
    If msg <> lastMsg Then
      Label1.Caption = msg
      Label1.ForeColor = vbRed
      Label1.FontBold = True
      lastMsg = msg
    End If
  Else
    Call HideWin
  End If
End Sub

Private Sub ShowWin()
  Me.Show
  isWindowShow = True
End Sub

Private Sub HideWin()
  Me.Hide
  isWindowShow = False
End Sub
