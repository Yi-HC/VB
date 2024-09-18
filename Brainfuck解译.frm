VERSION 5.00
Begin VB.Form Braunfuck解译 
   Caption         =   "Braunfuck解译"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6525
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   195
      TabIndex        =   4
      Top             =   2340
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "解译"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4875
      TabIndex        =   1
      Top             =   1560
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   195
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   195
      Width           =   6060
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   3
      Top             =   1365
      Width           =   1770
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   195
      TabIndex        =   2
      Top             =   3510
      Width           =   6060
   End
End
Attribute VB_Name = "Braunfuck解译"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'>:指针加一
'<:指针减一
'+:指针指向的字节的值加一
'-:指针指向的字节的值减一
'[:如果指针指向的单元值为零，向后跳转到对应的“]”指令的次一指令处
']:如果指针指向的单元值不为零，向前跳转到对应的“[”指令的次一指令处
'.:输出指针指向的单元内容(ASCII码)
Dim n As Integer, m As Integer, j As Integer, w As Integer
Dim s As String, sc_ As String
Dim f As Boolean
Dim b(1000) As Integer, bj(100) As Integer
Private Sub Command1_Click()
    For j = 1 To 1000
        b(j) = 0
    Next j
    m = 0: w = 0: sc_ = ""
    f = True
    If n <= 1999 Then
        For j = 1 To n
            If Not (Mid(s, j, 1) = "<" Or Mid(s, j, 1) = ">" Or Mid(s, j, 1) = "+" Or Mid(s, j, 1) = "-" Or Mid(s, j, 1) = "[" Or Mid(s, j, 1) = "]" Or Mid(s, j, 1) = ".") Then f = False: Exit For
        Next j
    Else
        f = False
    End If
    If f = False Then Call er1 Else Call get_
End Sub
Public Sub get_()
    f = True
    For j = 1 To n
        If Mid(s, j, 1) = "<" Then Call zy
        If Mid(s, j, 1) = ">" Then Call yy
        If Mid(s, j, 1) = "+" Then Call jf
        If Mid(s, j, 1) = "-" Then Call jf_
        If Mid(s, j, 1) = "[" Then Call xh1
        If Mid(s, j, 1) = "]" Then Call xh2
        If Mid(s, j, 1) = "." Then Call sc
        If f = False Then Exit For
    Next j
    If f = True Then Label1.Caption = sc_
End Sub
Public Sub zy()
    m = m - 1
    If m < 0 Then f = False: Call er3
End Sub
Public Sub yy()
    m = m + 1
    If m > 1000 Then f = False: Call er3
End Sub
Public Sub jf()
    b(m) = b(m) + 1
    If b(m) > 30000 Then f = False: Call er4
End Sub
Public Sub jf_()
    b(m) = b(m) - 1
    If b(m) < 0 Then f = False: Call er4
End Sub
Public Sub xh1()
    w = w + 1: bj(w) = j
    If b(m) = 0 Then
        Do While Mid(s, j, 1) <> "]" And j <> n + 1
            j = j + 1
        Loop
        If j = n + 1 Then Call er2
    End If
    If w >= 100 Then f = False: Call er5
End Sub
Public Sub xh2()
    If bj(w) = 0 Then
        f = False: Call er6
    Else
        If b(m) <> 0 Then j = bj(w) - 1
    End If
    w = w - 1
End Sub
Public Sub sc()
    sc_ = sc_ + Chr(b(m))
End Sub
Public Sub er1()
    Label1.Caption = "Error1:存在错误字符或输入字符过多(超过2000个)"
End Sub
Public Sub er2()
    Label1.Caption = "Error2:发生死循环"
End Sub
Public Sub er3()
    Label1.Caption = "Error3:指针越界(小于0或大于1000)"
End Sub
Public Sub er4()
    Label1.Caption = "Error4:数值越界(小于0或大于30000)"
End Sub
Public Sub er5()
    Label1.Caption = "Error5:循环次数过多(超过100次)"
End Sub
Public Sub er6()
    Label1.Caption = "Error6:循环结构语法错误"
End Sub

Private Sub Form_Load()
List1.AddItem ">:指针加一                           <:指针减一"
List1.AddItem "+:指针指向的字节的值加一             -:指针指向的字节的值减一"
List1.AddItem "[:如果指针指向的单元值为零，向后跳转到对应的“]”指令的次一指令处"
List1.AddItem "]:如果指针指向的单元值不为零，向前跳转到对应的“[”指令的次一指令处"
List1.AddItem ".:输出指针指向的单元内容(ASCII码)"
End Sub

Private Sub Text1_Change()
    s = Text1.Text
    n = Len(s)
    If n <= 2000 Then
        Label2.Caption = "还可输入" + CStr(2000 - Val(n)) + "个字符"
    Else
        Label2.Caption = "超过" + CStr(Val(n) - 2000) + "个字符"
    End If
End Sub
