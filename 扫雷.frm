VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   4050
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer4 
      Interval        =   30
      Left            =   3630
      Top             =   2715
   End
   Begin VB.Timer Timer3 
      Interval        =   30
      Left            =   3705
      Top             =   3990
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   180
      Left            =   675
      TabIndex        =   84
      Top             =   405
      Width           =   165
   End
   Begin VB.CommandButton Command82 
      Caption         =   "Command82"
      Height          =   225
      Left            =   9345
      TabIndex        =   83
      Top             =   7830
      Width           =   225
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3630
      Top             =   3585
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3630
      Top             =   3090
   End
   Begin VB.CommandButton c1 
      BackColor       =   &H8000000E&
      Caption         =   "ÖØÐÂ¿ªÊ¼"
      BeginProperty Font 
         Name            =   "ÐÂËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   1530
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   105
      Width           =   700
   End
   Begin VB.CommandButton Command81 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   3300
      Width           =   315
   End
   Begin VB.CommandButton Command80 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2700
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   3300
      Width           =   315
   End
   Begin VB.CommandButton Command79 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   3300
      Width           =   315
   End
   Begin VB.CommandButton Command78 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   3300
      Width           =   315
   End
   Begin VB.CommandButton Command77 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   3300
      Width           =   315
   End
   Begin VB.CommandButton Command76 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   3300
      Width           =   315
   End
   Begin VB.CommandButton Command75 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   3300
      Width           =   315
   End
   Begin VB.CommandButton Command74 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   3300
      Width           =   315
   End
   Begin VB.CommandButton Command73 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   3300
      Width           =   315
   End
   Begin VB.CommandButton Command72 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   3000
      Width           =   315
   End
   Begin VB.CommandButton Command71 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2700
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   3000
      Width           =   315
   End
   Begin VB.CommandButton Command70 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   3000
      Width           =   315
   End
   Begin VB.CommandButton Command69 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   3000
      Width           =   315
   End
   Begin VB.CommandButton Command68 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   3000
      Width           =   315
   End
   Begin VB.CommandButton Command67 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   3000
      Width           =   315
   End
   Begin VB.CommandButton Command66 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   3000
      Width           =   315
   End
   Begin VB.CommandButton Command65 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   3000
      Width           =   315
   End
   Begin VB.CommandButton Command64 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   3000
      Width           =   315
   End
   Begin VB.CommandButton Command63 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   2700
      Width           =   315
   End
   Begin VB.CommandButton Command62 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2700
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   2700
      Width           =   315
   End
   Begin VB.CommandButton Command61 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   2700
      Width           =   315
   End
   Begin VB.CommandButton Command60 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   2700
      Width           =   315
   End
   Begin VB.CommandButton Command59 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   2700
      Width           =   315
   End
   Begin VB.CommandButton Command58 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   2700
      Width           =   315
   End
   Begin VB.CommandButton Command57 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   2700
      Width           =   315
   End
   Begin VB.CommandButton Command56 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   2700
      Width           =   315
   End
   Begin VB.CommandButton Command55 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2700
      Width           =   315
   End
   Begin VB.CommandButton Command54 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2400
      Width           =   315
   End
   Begin VB.CommandButton Command53 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2700
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2400
      Width           =   315
   End
   Begin VB.CommandButton Command52 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2400
      Width           =   315
   End
   Begin VB.CommandButton Command51 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2400
      Width           =   315
   End
   Begin VB.CommandButton Command50 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2400
      Width           =   315
   End
   Begin VB.CommandButton Command49 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2400
      Width           =   315
   End
   Begin VB.CommandButton Command48 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2400
      Width           =   315
   End
   Begin VB.CommandButton Command47 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2400
      Width           =   315
   End
   Begin VB.CommandButton Command46 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2400
      Width           =   315
   End
   Begin VB.CommandButton Command45 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2100
      Width           =   315
   End
   Begin VB.CommandButton Command44 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2700
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2100
      Width           =   315
   End
   Begin VB.CommandButton Command43 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2100
      Width           =   315
   End
   Begin VB.CommandButton Command42 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2100
      Width           =   315
   End
   Begin VB.CommandButton Command41 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2100
      Width           =   315
   End
   Begin VB.CommandButton Command40 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2100
      Width           =   315
   End
   Begin VB.CommandButton Command39 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2100
      Width           =   315
   End
   Begin VB.CommandButton Command38 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2100
      Width           =   315
   End
   Begin VB.CommandButton Command37 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2100
      Width           =   315
   End
   Begin VB.CommandButton Command36 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1800
      Width           =   315
   End
   Begin VB.CommandButton Command35 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2700
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1800
      Width           =   315
   End
   Begin VB.CommandButton Command34 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1800
      Width           =   315
   End
   Begin VB.CommandButton Command33 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1800
      Width           =   315
   End
   Begin VB.CommandButton Command32 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1800
      Width           =   315
   End
   Begin VB.CommandButton Command31 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1800
      Width           =   315
   End
   Begin VB.CommandButton Command30 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1800
      Width           =   315
   End
   Begin VB.CommandButton Command29 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1800
      Width           =   315
   End
   Begin VB.CommandButton Command28 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1800
      Width           =   315
   End
   Begin VB.CommandButton Command27 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1500
      Width           =   315
   End
   Begin VB.CommandButton Command26 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2700
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1500
      Width           =   315
   End
   Begin VB.CommandButton Command25 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1500
      Width           =   315
   End
   Begin VB.CommandButton Command24 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1500
      Width           =   315
   End
   Begin VB.CommandButton Command23 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1500
      Width           =   315
   End
   Begin VB.CommandButton Command22 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1500
      Width           =   315
   End
   Begin VB.CommandButton Command21 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1500
      Width           =   315
   End
   Begin VB.CommandButton Command20 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1500
      Width           =   315
   End
   Begin VB.CommandButton Command19 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1500
      Width           =   315
   End
   Begin VB.CommandButton Command18 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1200
      Width           =   315
   End
   Begin VB.CommandButton Command17 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2700
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1200
      Width           =   315
   End
   Begin VB.CommandButton Command16 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1200
      Width           =   315
   End
   Begin VB.CommandButton Command15 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1200
      Width           =   315
   End
   Begin VB.CommandButton Command14 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1200
      Width           =   315
   End
   Begin VB.CommandButton Command13 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1200
      Width           =   315
   End
   Begin VB.CommandButton Command12 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Width           =   315
   End
   Begin VB.CommandButton Command11 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1200
      Width           =   315
   End
   Begin VB.CommandButton Command10 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   315
   End
   Begin VB.CommandButton Command9 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   900
      Width           =   315
   End
   Begin VB.CommandButton Command8 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2700
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   900
      Width           =   315
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   900
      Width           =   315
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   900
      Width           =   315
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   900
      Width           =   315
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   900
      Width           =   315
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   900
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   900
      Width           =   315
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   900
      Width           =   315
   End
   Begin VB.Label Label3 
      Height          =   585
      Left            =   570
      TabIndex        =   86
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "É¨À××´Ì¬"
      Height          =   420
      Left            =   960
      TabIndex        =   85
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   2535
      TabIndex        =   82
      Top             =   210
      Width           =   1065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, n As Integer, m As Integer, jishu As Integer, t As Integer, num As Integer
Dim a(1 To 81) As String
Dim flag As Boolean, f(1 To 81) As Boolean
Private Sub kuosan()
'ºÃ¼Ò»ï£¬ÎÒÕâÊÇÐ´ÁË¸öÊ²Ã´ÍæÒâ£¬Ã¿µãÒ»´Î¾ÍÈ«²¿¼ì²âÒ»±é£¬¹Ö²»µÃ¶ÑÕ»¿Õ¼ä²»×ã''Ò²ÐíÈ±ÉÙÍË³ö£¿£¿'''²Ý£¬»áÔÚ8ºÍ9Ö®¼äËÀÑ­»·''''³¢ÊÔ²ð·ÖÒ»ÏÂ''''''²»ÊÇ£¬ÎÒÕâÀ©É¢À©É¢ÁË¸ö¼ÅÄ¯£¬½á¹ûàÏ£¬µ¥´¿ÔÚËÀÑ­»·°¡''''''''''·ÅÆúÁË
'µ±ÖÐ×ªÕ¾£¿ÔÙÃ¶¾Ù£¿£¿£¿£¿'¼¸°ÙÐÐÃ¶¾Ù£¬ÎÒÍÂÁË
If m = 1 Then
    If Command2.Enabled = True Then Call Command2_Click
    If Command10.Enabled = True Then Call Command10_Click
    If Command11.Enabled = True Then Call Command11_Click
End If
If m = 2 Then
    If Command1.Enabled = True Then Call Command1_Click
    If Command3.Enabled = True Then Call Command3_Click
    If Command10.Enabled = True Then Call Command10_Click
    If Command11.Enabled = True Then Call Command11_Click
    If Command12.Enabled = True Then Call Command12_Click
End If
If m = 3 Then
    If Command2.Enabled = True Then Call Command2_Click
    If Command4.Enabled = True Then Call Command4_Click
    If Command11.Enabled = True Then Call Command11_Click
    If Command12.Enabled = True Then Call Command12_Click
    If Command13.Enabled = True Then Call Command13_Click
End If
If m = 4 Then
    If Command3.Enabled = True Then Call Command3_Click
    If Command5.Enabled = True Then Call Command5_Click
    If Command12.Enabled = True Then Call Command12_Click
    If Command13.Enabled = True Then Call Command13_Click
    If Command14.Enabled = True Then Call Command14_Click
End If
If m = 5 Then
    If Command4.Enabled = True Then Call Command4_Click
    If Command6.Enabled = True Then Call Command6_Click
    If Command13.Enabled = True Then Call Command13_Click
    If Command14.Enabled = True Then Call Command14_Click
    If Command15.Enabled = True Then Call Command15_Click
End If
If m = 6 Then
    If Command5.Enabled = True Then Call Command5_Click
    If Command7.Enabled = True Then Call Command7_Click
    If Command14.Enabled = True Then Call Command14_Click
    If Command15.Enabled = True Then Call Command15_Click
    If Command16.Enabled = True Then Call Command16_Click
End If
If m = 7 Then
    If Command6.Enabled = True Then Call Command6_Click
    If Command8.Enabled = True Then Call Command8_Click
    If Command15.Enabled = True Then Call Command15_Click
    If Command16.Enabled = True Then Call Command16_Click
    If Command17.Enabled = True Then Call Command17_Click
End If
If m = 8 Then
    If Command7.Enabled = True Then Call Command7_Click
    If Command9.Enabled = True Then Call Command9_Click
    If Command16.Enabled = True Then Call Command16_Click
    If Command17.Enabled = True Then Call Command17_Click
    If Command18.Enabled = True Then Call Command18_Click
End If
If m = 9 Then
    If Command8.Enabled = True Then Call Command8_Click
    If Command17.Enabled = True Then Call Command17_Click
    If Command18.Enabled = True Then Call Command18_Click
End If
If m = 10 Then
    If Command1.Enabled = True Then Call Command1_Click
    If Command2.Enabled = True Then Call Command2_Click
    If Command11.Enabled = True Then Call Command11_Click
    If Command19.Enabled = True Then Call Command19_Click
    If Command20.Enabled = True Then Call Command20_Click
End If
If m = 11 Then
    If Command1.Enabled = True Then Call Command1_Click
    If Command2.Enabled = True Then Call Command2_Click
    If Command3.Enabled = True Then Call Command3_Click
    If Command10.Enabled = True Then Call Command10_Click
    If Command12.Enabled = True Then Call Command12_Click
    If Command19.Enabled = True Then Call Command19_Click
    If Command20.Enabled = True Then Call Command20_Click
    If Command21.Enabled = True Then Call Command21_Click
End If
If m = 12 Then
    If Command2.Enabled = True Then Call Command2_Click
    If Command3.Enabled = True Then Call Command3_Click
    If Command4.Enabled = True Then Call Command4_Click
    If Command11.Enabled = True Then Call Command11_Click
    If Command13.Enabled = True Then Call Command13_Click
    If Command20.Enabled = True Then Call Command20_Click
    If Command21.Enabled = True Then Call Command21_Click
    If Command22.Enabled = True Then Call Command22_Click
End If
If m = 13 Then
    If Command3.Enabled = True Then Call Command3_Click
    If Command4.Enabled = True Then Call Command4_Click
    If Command5.Enabled = True Then Call Command5_Click
    If Command12.Enabled = True Then Call Command12_Click
    If Command14.Enabled = True Then Call Command14_Click
    If Command21.Enabled = True Then Call Command21_Click
    If Command22.Enabled = True Then Call Command22_Click
    If Command23.Enabled = True Then Call Command23_Click
End If
If m = 14 Then
    If Command4.Enabled = True Then Call Command4_Click
    If Command5.Enabled = True Then Call Command5_Click
    If Command6.Enabled = True Then Call Command6_Click
    If Command13.Enabled = True Then Call Command13_Click
    If Command15.Enabled = True Then Call Command15_Click
    If Command22.Enabled = True Then Call Command22_Click
    If Command23.Enabled = True Then Call Command23_Click
    If Command24.Enabled = True Then Call Command24_Click
End If
If m = 15 Then
    If Command5.Enabled = True Then Call Command5_Click
    If Command6.Enabled = True Then Call Command6_Click
    If Command7.Enabled = True Then Call Command7_Click
    If Command14.Enabled = True Then Call Command14_Click
    If Command16.Enabled = True Then Call Command16_Click
    If Command23.Enabled = True Then Call Command23_Click
    If Command24.Enabled = True Then Call Command24_Click
    If Command25.Enabled = True Then Call Command25_Click
End If
If m = 16 Then
    If Command6.Enabled = True Then Call Command6_Click
    If Command7.Enabled = True Then Call Command7_Click
    If Command8.Enabled = True Then Call Command8_Click
    If Command15.Enabled = True Then Call Command15_Click
    If Command17.Enabled = True Then Call Command16_Click
    If Command24.Enabled = True Then Call Command24_Click
    If Command25.Enabled = True Then Call Command25_Click
    If Command26.Enabled = True Then Call Command26_Click
End If
If m = 17 Then
    If Command7.Enabled = True Then Call Command7_Click
    If Command8.Enabled = True Then Call Command8_Click
    If Command9.Enabled = True Then Call Command9_Click
    If Command16.Enabled = True Then Call Command16_Click
    If Command18.Enabled = True Then Call Command18_Click
    If Command25.Enabled = True Then Call Command25_Click
    If Command26.Enabled = True Then Call Command26_Click
    If Command27.Enabled = True Then Call Command27_Click
End If
If m = 18 Then
    If Command8.Enabled = True Then Call Command8_Click
    If Command9.Enabled = True Then Call Command9_Click
    If Command17.Enabled = True Then Call Command17_Click
    If Command26.Enabled = True Then Call Command26_Click
    If Command27.Enabled = True Then Call Command27_Click
End If
If m = 19 Then
    If Command10.Enabled = True Then Call Command10_Click
    If Command11.Enabled = True Then Call Command11_Click
    If Command20.Enabled = True Then Call Command20_Click
    If Command28.Enabled = True Then Call Command28_Click
    If Command29.Enabled = True Then Call Command29_Click
End If
If m = 20 Then
    If Command10.Enabled = True Then Call Command10_Click
    If Command11.Enabled = True Then Call Command11_Click
    If Command12.Enabled = True Then Call Command12_Click
    If Command19.Enabled = True Then Call Command19_Click
    If Command21.Enabled = True Then Call Command21_Click
    If Command28.Enabled = True Then Call Command28_Click
    If Command29.Enabled = True Then Call Command29_Click
    If Command30.Enabled = True Then Call Command30_Click
End If
If m = 21 Then
    If Command11.Enabled = True Then Call Command11_Click
    If Command12.Enabled = True Then Call Command12_Click
    If Command13.Enabled = True Then Call Command13_Click
    If Command20.Enabled = True Then Call Command20_Click
    If Command22.Enabled = True Then Call Command22_Click
    If Command29.Enabled = True Then Call Command29_Click
    If Command30.Enabled = True Then Call Command30_Click
    If Command31.Enabled = True Then Call Command31_Click
End If
If m = 22 Then
    If Command12.Enabled = True Then Call Command12_Click
    If Command13.Enabled = True Then Call Command13_Click
    If Command14.Enabled = True Then Call Command14_Click
    If Command21.Enabled = True Then Call Command21_Click
    If Command23.Enabled = True Then Call Command23_Click
    If Command30.Enabled = True Then Call Command30_Click
    If Command31.Enabled = True Then Call Command31_Click
    If Command32.Enabled = True Then Call Command32_Click
End If
If m = 23 Then
    If Command13.Enabled = True Then Call Command13_Click
    If Command14.Enabled = True Then Call Command14_Click
    If Command15.Enabled = True Then Call Command15_Click
    If Command22.Enabled = True Then Call Command22_Click
    If Command24.Enabled = True Then Call Command24_Click
    If Command31.Enabled = True Then Call Command31_Click
    If Command32.Enabled = True Then Call Command32_Click
    If Command33.Enabled = True Then Call Command33_Click
End If
If m = 24 Then
    If Command14.Enabled = True Then Call Command14_Click
    If Command15.Enabled = True Then Call Command15_Click
    If Command16.Enabled = True Then Call Command16_Click
    If Command23.Enabled = True Then Call Command23_Click
    If Command25.Enabled = True Then Call Command25_Click
    If Command32.Enabled = True Then Call Command32_Click
    If Command33.Enabled = True Then Call Command33_Click
    If Command34.Enabled = True Then Call Command34_Click
End If
If m = 25 Then
    If Command15.Enabled = True Then Call Command15_Click
    If Command16.Enabled = True Then Call Command16_Click
    If Command17.Enabled = True Then Call Command17_Click
    If Command24.Enabled = True Then Call Command24_Click
    If Command26.Enabled = True Then Call Command26_Click
    If Command33.Enabled = True Then Call Command33_Click
    If Command34.Enabled = True Then Call Command34_Click
    If Command35.Enabled = True Then Call Command35_Click
End If
If m = 26 Then
    If Command16.Enabled = True Then Call Command16_Click
    If Command17.Enabled = True Then Call Command17_Click
    If Command18.Enabled = True Then Call Command18_Click
    If Command25.Enabled = True Then Call Command25_Click
    If Command27.Enabled = True Then Call Command27_Click
    If Command34.Enabled = True Then Call Command34_Click
    If Command35.Enabled = True Then Call Command35_Click
    If Command36.Enabled = True Then Call Command36_Click
End If
If m = 27 Then
    If Command17.Enabled = True Then Call Command17_Click
    If Command18.Enabled = True Then Call Command18_Click
    If Command26.Enabled = True Then Call Command26_Click
    If Command35.Enabled = True Then Call Command35_Click
    If Command36.Enabled = True Then Call Command36_Click
End If
If m = 28 Then
    If Command19.Enabled = True Then Call Command19_Click
    If Command20.Enabled = True Then Call Command20_Click
    If Command29.Enabled = True Then Call Command29_Click
    If Command37.Enabled = True Then Call Command37_Click
    If Command38.Enabled = True Then Call Command38_Click
End If
If m = 29 Then
    If Command19.Enabled = True Then Call Command19_Click
    If Command20.Enabled = True Then Call Command20_Click
    If Command21.Enabled = True Then Call Command21_Click
    If Command28.Enabled = True Then Call Command28_Click
    If Command30.Enabled = True Then Call Command30_Click
    If Command37.Enabled = True Then Call Command37_Click
    If Command38.Enabled = True Then Call Command38_Click
    If Command39.Enabled = True Then Call Command39_Click
End If
If m = 30 Then
    If Command20.Enabled = True Then Call Command20_Click
    If Command21.Enabled = True Then Call Command21_Click
    If Command22.Enabled = True Then Call Command22_Click
    If Command29.Enabled = True Then Call Command29_Click
    If Command31.Enabled = True Then Call Command31_Click
    If Command38.Enabled = True Then Call Command38_Click
    If Command39.Enabled = True Then Call Command39_Click
    If Command40.Enabled = True Then Call Command40_Click
End If
If m = 31 Then
    If Command21.Enabled = True Then Call Command21_Click
    If Command22.Enabled = True Then Call Command22_Click
    If Command23.Enabled = True Then Call Command23_Click
    If Command30.Enabled = True Then Call Command30_Click
    If Command32.Enabled = True Then Call Command32_Click
    If Command39.Enabled = True Then Call Command39_Click
    If Command40.Enabled = True Then Call Command40_Click
    If Command41.Enabled = True Then Call Command41_Click
End If
If m = 32 Then
    If Command22.Enabled = True Then Call Command22_Click
    If Command23.Enabled = True Then Call Command23_Click
    If Command24.Enabled = True Then Call Command24_Click
    If Command31.Enabled = True Then Call Command31_Click
    If Command33.Enabled = True Then Call Command33_Click
    If Command40.Enabled = True Then Call Command40_Click
    If Command41.Enabled = True Then Call Command41_Click
    If Command42.Enabled = True Then Call Command42_Click
End If
If m = 33 Then
    If Command23.Enabled = True Then Call Command23_Click
    If Command24.Enabled = True Then Call Command24_Click
    If Command25.Enabled = True Then Call Command25_Click
    If Command32.Enabled = True Then Call Command32_Click
    If Command34.Enabled = True Then Call Command34_Click
    If Command41.Enabled = True Then Call Command41_Click
    If Command42.Enabled = True Then Call Command42_Click
    If Command43.Enabled = True Then Call Command43_Click
End If
If m = 34 Then
    If Command24.Enabled = True Then Call Command24_Click
    If Command25.Enabled = True Then Call Command25_Click
    If Command26.Enabled = True Then Call Command26_Click
    If Command33.Enabled = True Then Call Command33_Click
    If Command35.Enabled = True Then Call Command35_Click
    If Command42.Enabled = True Then Call Command42_Click
    If Command43.Enabled = True Then Call Command43_Click
    If Command44.Enabled = True Then Call Command44_Click
End If
If m = 35 Then
    If Command25.Enabled = True Then Call Command25_Click
    If Command26.Enabled = True Then Call Command26_Click
    If Command27.Enabled = True Then Call Command27_Click
    If Command34.Enabled = True Then Call Command34_Click
    If Command36.Enabled = True Then Call Command36_Click
    If Command43.Enabled = True Then Call Command43_Click
    If Command44.Enabled = True Then Call Command44_Click
    If Command45.Enabled = True Then Call Command45_Click
End If
If m = 36 Then
    If Command26.Enabled = True Then Call Command26_Click
    If Command27.Enabled = True Then Call Command27_Click
    If Command35.Enabled = True Then Call Command34_Click
    If Command44.Enabled = True Then Call Command44_Click
    If Command45.Enabled = True Then Call Command45_Click
End If
If m = 37 Then
    If Command28.Enabled = True Then Call Command28_Click
    If Command29.Enabled = True Then Call Command29_Click
    If Command38.Enabled = True Then Call Command38_Click
    If Command46.Enabled = True Then Call Command46_Click
    If Command47.Enabled = True Then Call Command47_Click
End If
If m = 38 Then
    If Command28.Enabled = True Then Call Command28_Click
    If Command29.Enabled = True Then Call Command29_Click
    If Command30.Enabled = True Then Call Command30_Click
    If Command37.Enabled = True Then Call Command37_Click
    If Command39.Enabled = True Then Call Command39_Click
    If Command46.Enabled = True Then Call Command46_Click
    If Command47.Enabled = True Then Call Command47_Click
    If Command48.Enabled = True Then Call Command48_Click
End If
If m = 39 Then
    If Command29.Enabled = True Then Call Command29_Click
    If Command30.Enabled = True Then Call Command30_Click
    If Command31.Enabled = True Then Call Command31_Click
    If Command38.Enabled = True Then Call Command38_Click
    If Command40.Enabled = True Then Call Command40_Click
    If Command47.Enabled = True Then Call Command47_Click
    If Command48.Enabled = True Then Call Command48_Click
    If Command49.Enabled = True Then Call Command49_Click
End If
If m = 40 Then
    If Command30.Enabled = True Then Call Command30_Click
    If Command31.Enabled = True Then Call Command31_Click
    If Command32.Enabled = True Then Call Command32_Click
    If Command39.Enabled = True Then Call Command39_Click
    If Command41.Enabled = True Then Call Command41_Click
    If Command48.Enabled = True Then Call Command48_Click
    If Command49.Enabled = True Then Call Command49_Click
    If Command50.Enabled = True Then Call Command50_Click
End If
If m = 41 Then
    If Command31.Enabled = True Then Call Command31_Click
    If Command32.Enabled = True Then Call Command32_Click
    If Command33.Enabled = True Then Call Command33_Click
    If Command40.Enabled = True Then Call Command40_Click
    If Command42.Enabled = True Then Call Command42_Click
    If Command49.Enabled = True Then Call Command49_Click
    If Command50.Enabled = True Then Call Command50_Click
    If Command51.Enabled = True Then Call Command51_Click
End If
If m = 42 Then
    If Command32.Enabled = True Then Call Command32_Click
    If Command33.Enabled = True Then Call Command33_Click
    If Command34.Enabled = True Then Call Command34_Click
    If Command41.Enabled = True Then Call Command41_Click
    If Command43.Enabled = True Then Call Command43_Click
    If Command50.Enabled = True Then Call Command50_Click
    If Command51.Enabled = True Then Call Command51_Click
    If Command52.Enabled = True Then Call Command52_Click
End If
If m = 43 Then
    If Command33.Enabled = True Then Call Command33_Click
    If Command34.Enabled = True Then Call Command34_Click
    If Command35.Enabled = True Then Call Command35_Click
    If Command42.Enabled = True Then Call Command42_Click
    If Command44.Enabled = True Then Call Command44_Click
    If Command51.Enabled = True Then Call Command51_Click
    If Command52.Enabled = True Then Call Command52_Click
    If Command53.Enabled = True Then Call Command53_Click
End If
If m = 44 Then
    If Command34.Enabled = True Then Call Command34_Click
    If Command35.Enabled = True Then Call Command35_Click
    If Command36.Enabled = True Then Call Command36_Click
    If Command43.Enabled = True Then Call Command43_Click
    If Command45.Enabled = True Then Call Command45_Click
    If Command52.Enabled = True Then Call Command52_Click
    If Command53.Enabled = True Then Call Command53_Click
    If Command54.Enabled = True Then Call Command54_Click
End If
If m = 45 Then
    If Command35.Enabled = True Then Call Command35_Click
    If Command36.Enabled = True Then Call Command36_Click
    If Command44.Enabled = True Then Call Command44_Click
    If Command53.Enabled = True Then Call Command53_Click
    If Command54.Enabled = True Then Call Command54_Click
End If
If m = 46 Then
    If Command37.Enabled = True Then Call Command37_Click
    If Command38.Enabled = True Then Call Command38_Click
    If Command47.Enabled = True Then Call Command47_Click
    If Command55.Enabled = True Then Call Command55_Click
    If Command56.Enabled = True Then Call Command56_Click
End If
If m = 47 Then
    If Command37.Enabled = True Then Call Command37_Click
    If Command38.Enabled = True Then Call Command38_Click
    If Command39.Enabled = True Then Call Command39_Click
    If Command46.Enabled = True Then Call Command46_Click
    If Command48.Enabled = True Then Call Command48_Click
    If Command55.Enabled = True Then Call Command55_Click
    If Command56.Enabled = True Then Call Command56_Click
    If Command57.Enabled = True Then Call Command57_Click
End If
If m = 48 Then
    If Command38.Enabled = True Then Call Command38_Click
    If Command39.Enabled = True Then Call Command39_Click
    If Command40.Enabled = True Then Call Command40_Click
    If Command47.Enabled = True Then Call Command47_Click
    If Command49.Enabled = True Then Call Command49_Click
    If Command56.Enabled = True Then Call Command56_Click
    If Command57.Enabled = True Then Call Command57_Click
    If Command58.Enabled = True Then Call Command58_Click
End If
If m = 49 Then
    If Command39.Enabled = True Then Call Command39_Click
    If Command40.Enabled = True Then Call Command40_Click
    If Command41.Enabled = True Then Call Command41_Click
    If Command48.Enabled = True Then Call Command48_Click
    If Command50.Enabled = True Then Call Command50_Click
    If Command57.Enabled = True Then Call Command57_Click
    If Command58.Enabled = True Then Call Command58_Click
    If Command59.Enabled = True Then Call Command59_Click
End If
If m = 50 Then
    If Command40.Enabled = True Then Call Command40_Click
    If Command41.Enabled = True Then Call Command41_Click
    If Command42.Enabled = True Then Call Command42_Click
    If Command49.Enabled = True Then Call Command49_Click
    If Command51.Enabled = True Then Call Command51_Click
    If Command58.Enabled = True Then Call Command58_Click
    If Command59.Enabled = True Then Call Command59_Click
    If Command60.Enabled = True Then Call Command60_Click
End If
If m = 51 Then
    If Command41.Enabled = True Then Call Command41_Click
    If Command42.Enabled = True Then Call Command42_Click
    If Command43.Enabled = True Then Call Command43_Click
    If Command50.Enabled = True Then Call Command50_Click
    If Command52.Enabled = True Then Call Command52_Click
    If Command59.Enabled = True Then Call Command59_Click
    If Command60.Enabled = True Then Call Command60_Click
    If Command61.Enabled = True Then Call Command61_Click
End If
If m = 52 Then
    If Command42.Enabled = True Then Call Command42_Click
    If Command43.Enabled = True Then Call Command43_Click
    If Command44.Enabled = True Then Call Command44_Click
    If Command51.Enabled = True Then Call Command51_Click
    If Command53.Enabled = True Then Call Command53_Click
    If Command60.Enabled = True Then Call Command60_Click
    If Command61.Enabled = True Then Call Command61_Click
    If Command62.Enabled = True Then Call Command62_Click
End If
If m = 53 Then
    If Command43.Enabled = True Then Call Command43_Click
    If Command44.Enabled = True Then Call Command44_Click
    If Command45.Enabled = True Then Call Command45_Click
    If Command52.Enabled = True Then Call Command52_Click
    If Command54.Enabled = True Then Call Command54_Click
    If Command61.Enabled = True Then Call Command61_Click
    If Command62.Enabled = True Then Call Command62_Click
    If Command63.Enabled = True Then Call Command63_Click
End If
If m = 54 Then
    If Command44.Enabled = True Then Call Command44_Click
    If Command45.Enabled = True Then Call Command45_Click
    If Command53.Enabled = True Then Call Command53_Click
    If Command62.Enabled = True Then Call Command62_Click
    If Command63.Enabled = True Then Call Command63_Click
End If
If m = 55 Then
    If Command46.Enabled = True Then Call Command46_Click
    If Command47.Enabled = True Then Call Command47_Click
    If Command56.Enabled = True Then Call Command56_Click
    If Command64.Enabled = True Then Call Command64_Click
    If Command65.Enabled = True Then Call Command65_Click
End If
If m = 56 Then
    If Command46.Enabled = True Then Call Command46_Click
    If Command47.Enabled = True Then Call Command47_Click
    If Command48.Enabled = True Then Call Command48_Click
    If Command55.Enabled = True Then Call Command55_Click
    If Command57.Enabled = True Then Call Command57_Click
    If Command64.Enabled = True Then Call Command64_Click
    If Command65.Enabled = True Then Call Command65_Click
    If Command66.Enabled = True Then Call Command66_Click
End If
If m = 57 Then
    If Command47.Enabled = True Then Call Command47_Click
    If Command48.Enabled = True Then Call Command48_Click
    If Command49.Enabled = True Then Call Command49_Click
    If Command56.Enabled = True Then Call Command56_Click
    If Command58.Enabled = True Then Call Command58_Click
    If Command65.Enabled = True Then Call Command65_Click
    If Command66.Enabled = True Then Call Command66_Click
    If Command67.Enabled = True Then Call Command67_Click
End If
If m = 58 Then
    If Command48.Enabled = True Then Call Command48_Click
    If Command49.Enabled = True Then Call Command49_Click
    If Command50.Enabled = True Then Call Command50_Click
    If Command57.Enabled = True Then Call Command57_Click
    If Command59.Enabled = True Then Call Command59_Click
    If Command66.Enabled = True Then Call Command66_Click
    If Command67.Enabled = True Then Call Command67_Click
    If Command68.Enabled = True Then Call Command68_Click
End If
If m = 59 Then
    If Command49.Enabled = True Then Call Command49_Click
    If Command50.Enabled = True Then Call Command50_Click
    If Command51.Enabled = True Then Call Command51_Click
    If Command58.Enabled = True Then Call Command58_Click
    If Command60.Enabled = True Then Call Command60_Click
    If Command67.Enabled = True Then Call Command67_Click
    If Command68.Enabled = True Then Call Command68_Click
    If Command69.Enabled = True Then Call Command69_Click
End If
If m = 60 Then
    If Command50.Enabled = True Then Call Command50_Click
    If Command51.Enabled = True Then Call Command51_Click
    If Command52.Enabled = True Then Call Command52_Click
    If Command59.Enabled = True Then Call Command59_Click
    If Command61.Enabled = True Then Call Command61_Click
    If Command68.Enabled = True Then Call Command68_Click
    If Command69.Enabled = True Then Call Command69_Click
    If Command70.Enabled = True Then Call Command70_Click
End If
If m = 61 Then
    If Command51.Enabled = True Then Call Command51_Click
    If Command52.Enabled = True Then Call Command52_Click
    If Command53.Enabled = True Then Call Command53_Click
    If Command60.Enabled = True Then Call Command60_Click
    If Command62.Enabled = True Then Call Command62_Click
    If Command69.Enabled = True Then Call Command69_Click
    If Command70.Enabled = True Then Call Command70_Click
    If Command71.Enabled = True Then Call Command71_Click
End If
If m = 62 Then
    If Command52.Enabled = True Then Call Command52_Click
    If Command53.Enabled = True Then Call Command53_Click
    If Command54.Enabled = True Then Call Command54_Click
    If Command61.Enabled = True Then Call Command61_Click
    If Command63.Enabled = True Then Call Command63_Click
    If Command70.Enabled = True Then Call Command70_Click
    If Command71.Enabled = True Then Call Command71_Click
    If Command72.Enabled = True Then Call Command72_Click
End If
If m = 63 Then
    If Command53.Enabled = True Then Call Command53_Click
    If Command54.Enabled = True Then Call Command54_Click
    If Command62.Enabled = True Then Call Command62_Click
    If Command71.Enabled = True Then Call Command71_Click
    If Command72.Enabled = True Then Call Command72_Click
End If
If m = 64 Then
    If Command55.Enabled = True Then Call Command55_Click
    If Command56.Enabled = True Then Call Command56_Click
    If Command65.Enabled = True Then Call Command65_Click
    If Command73.Enabled = True Then Call Command73_Click
    If Command74.Enabled = True Then Call Command74_Click
End If
If m = 65 Then
    If Command55.Enabled = True Then Call Command55_Click
    If Command56.Enabled = True Then Call Command56_Click
    If Command57.Enabled = True Then Call Command57_Click
    If Command64.Enabled = True Then Call Command64_Click
    If Command66.Enabled = True Then Call Command66_Click
    If Command73.Enabled = True Then Call Command73_Click
    If Command74.Enabled = True Then Call Command74_Click
    If Command75.Enabled = True Then Call Command75_Click
End If
If m = 66 Then
    If Command56.Enabled = True Then Call Command56_Click
    If Command57.Enabled = True Then Call Command57_Click
    If Command58.Enabled = True Then Call Command58_Click
    If Command65.Enabled = True Then Call Command65_Click
    If Command67.Enabled = True Then Call Command67_Click
    If Command74.Enabled = True Then Call Command74_Click
    If Command75.Enabled = True Then Call Command75_Click
    If Command76.Enabled = True Then Call Command76_Click
End If
If m = 67 Then
    If Command57.Enabled = True Then Call Command57_Click
    If Command58.Enabled = True Then Call Command58_Click
    If Command59.Enabled = True Then Call Command59_Click
    If Command66.Enabled = True Then Call Command66_Click
    If Command68.Enabled = True Then Call Command68_Click
    If Command75.Enabled = True Then Call Command75_Click
    If Command76.Enabled = True Then Call Command76_Click
    If Command77.Enabled = True Then Call Command77_Click
End If
If m = 68 Then
    If Command58.Enabled = True Then Call Command58_Click
    If Command59.Enabled = True Then Call Command59_Click
    If Command60.Enabled = True Then Call Command60_Click
    If Command67.Enabled = True Then Call Command67_Click
    If Command69.Enabled = True Then Call Command69_Click
    If Command76.Enabled = True Then Call Command76_Click
    If Command77.Enabled = True Then Call Command77_Click
    If Command78.Enabled = True Then Call Command78_Click
End If
If m = 69 Then
    If Command59.Enabled = True Then Call Command59_Click
    If Command60.Enabled = True Then Call Command60_Click
    If Command61.Enabled = True Then Call Command61_Click
    If Command68.Enabled = True Then Call Command68_Click
    If Command70.Enabled = True Then Call Command70_Click
    If Command77.Enabled = True Then Call Command77_Click
    If Command78.Enabled = True Then Call Command78_Click
    If Command79.Enabled = True Then Call Command79_Click
End If
If m = 70 Then
    If Command60.Enabled = True Then Call Command60_Click
    If Command61.Enabled = True Then Call Command61_Click
    If Command62.Enabled = True Then Call Command62_Click
    If Command69.Enabled = True Then Call Command69_Click
    If Command71.Enabled = True Then Call Command71_Click
    If Command78.Enabled = True Then Call Command78_Click
    If Command79.Enabled = True Then Call Command79_Click
    If Command80.Enabled = True Then Call Command80_Click
End If
If m = 71 Then
    If Command61.Enabled = True Then Call Command61_Click
    If Command62.Enabled = True Then Call Command62_Click
    If Command63.Enabled = True Then Call Command63_Click
    If Command70.Enabled = True Then Call Command70_Click
    If Command73.Enabled = True Then Call Command73_Click
    If Command79.Enabled = True Then Call Command79_Click
    If Command80.Enabled = True Then Call Command80_Click
    If Command81.Enabled = True Then Call Command81_Click
End If
If m = 72 Then
    If Command62.Enabled = True Then Call Command62_Click
    If Command63.Enabled = True Then Call Command63_Click
    If Command73.Enabled = True Then Call Command73_Click
    If Command80.Enabled = True Then Call Command80_Click
    If Command81.Enabled = True Then Call Command81_Click
End If
If m = 73 Then
    If Command64.Enabled = True Then Call Command64_Click
    If Command65.Enabled = True Then Call Command65_Click
    If Command74.Enabled = True Then Call Command74_Click
End If
If m = 74 Then
    If Command64.Enabled = True Then Call Command64_Click
    If Command65.Enabled = True Then Call Command65_Click
    If Command66.Enabled = True Then Call Command66_Click
    If Command73.Enabled = True Then Call Command73_Click
    If Command75.Enabled = True Then Call Command75_Click
End If
If m = 75 Then
    If Command65.Enabled = True Then Call Command65_Click
    If Command66.Enabled = True Then Call Command66_Click
    If Command67.Enabled = True Then Call Command67_Click
    If Command74.Enabled = True Then Call Command74_Click
    If Command76.Enabled = True Then Call Command76_Click
End If
If m = 76 Then
    If Command66.Enabled = True Then Call Command66_Click
    If Command67.Enabled = True Then Call Command67_Click
    If Command68.Enabled = True Then Call Command68_Click
    If Command75.Enabled = True Then Call Command75_Click
    If Command77.Enabled = True Then Call Command77_Click
End If
If m = 77 Then
    If Command67.Enabled = True Then Call Command67_Click
    If Command68.Enabled = True Then Call Command68_Click
    If Command69.Enabled = True Then Call Command69_Click
    If Command76.Enabled = True Then Call Command76_Click
    If Command78.Enabled = True Then Call Command78_Click
End If
If m = 78 Then
    If Command68.Enabled = True Then Call Command68_Click
    If Command69.Enabled = True Then Call Command69_Click
    If Command70.Enabled = True Then Call Command70_Click
    If Command77.Enabled = True Then Call Command77_Click
    If Command79.Enabled = True Then Call Command79_Click
End If
If m = 79 Then
    If Command69.Enabled = True Then Call Command69_Click
    If Command70.Enabled = True Then Call Command70_Click
    If Command71.Enabled = True Then Call Command71_Click
    If Command78.Enabled = True Then Call Command78_Click
    If Command80.Enabled = True Then Call Command80_Click
End If
If m = 80 Then
    If Command70.Enabled = True Then Call Command70_Click
    If Command71.Enabled = True Then Call Command71_Click
    If Command72.Enabled = True Then Call Command72_Click
    If Command79.Enabled = True Then Call Command79_Click
    If Command81.Enabled = True Then Call Command81_Click
End If
If m = 81 Then
    If Command71.Enabled = True Then Call Command71_Click
    If Command72.Enabled = True Then Call Command72_Click
    If Command80.Enabled = True Then Call Command80_Click
End If
End Sub
Function panding(X As Integer) As String
    If a(X) = "" Then
        If X = 1 Then jishu = jiance(X + 1) + jiance(X + 9) + jiance(X + 10)
        If X = 9 Then jishu = jiance(X - 1) + jiance(X + 9) + jiance(X + 8)
        If X = 73 Then jishu = jiance(X + 1) + jiance(X - 9) + jiance(X - 8)
        If X = 81 Then jishu = jiance(X - 1) + jiance(X - 9) + jiance(X - 10)
        If X > 1 And X < 9 Then jishu = jiance(X + 1) + jiance(X - 1) + jiance(X + 8) + jiance(X + 9) + jiance(X + 10)
        If X = 10 Or X = 19 Or X = 28 Or X = 37 Or X = 46 Or X = 55 Or X = 64 Then jishu = jiance(X + 1) + jiance(X - 9) + jiance(X - 8) + jiance(X + 9) + jiance(X + 10)
        If X = 18 Or X = 27 Or X = 36 Or X = 45 Or X = 54 Or X = 63 Or X = 72 Then jishu = jiance(X - 1) + jiance(X - 9) + jiance(X - 10) + jiance(X + 9) + jiance(X + 8)
        If X > 73 And X < 81 Then jishu = jiance(X - 1) + jiance(X + 1) + jiance(X - 8) + jiance(X - 9) + jiance(X - 10)
        If X > 10 And X < 18 Or X > 19 And X < 27 Or X > 28 And X < 36 Or X > 37 And X < 45 Or X > 46 And X < 54 Or X > 55 And X < 63 Or X > 64 And X < 72 Then jishu = jiance(X - 1) + jiance(X + 1) + jiance(X - 8) + jiance(X - 9) + jiance(X - 10) + jiance(X + 8) + jiance(X + 9) + jiance(X + 10)
    End If
    If jishu = 0 Then panding = "0" Else panding = CStr(jishu)
End Function
Function jiance(Y As Integer) As Integer
    Dim js As Integer
        If a(Y) = "*" Then js = js + 1
        jiance = js
End Function

Private Sub c1_Click()
    Call chushi1
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then Label2.Caption = "±ê¼Ç×´Ì¬": flag = True
    If Check1.Value = 0 Then Label2.Caption = "É¨À××´Ì¬": flag = False
End Sub

Private Sub Command1_Click()
    m = 1
    If flag = True Then
        Command1.Caption = "!": If Command1.Caption = "!" And f(m) = False Then Command1.Caption = "": f(m) = True: num = num - 1
    ElseIf f(m) = True Then
        If panding(m) = 0 Then Call kuosan Else Command1.Caption = panding(m)
        If a(m) = "*" Then Call jieshu
        Command1.Enabled = False
    End If
    Timer1.Enabled = True
End Sub
Private Sub Command10_Click()
    m = 10
    If flag = True Then
        Command10.Caption = "!": If Command10.Caption = "!" And f(10) = False Then Command10.Caption = "": f(10) = True: num = num - 1
    ElseIf f(10) = True Then
        If panding(m) = 0 Then Call kuosan Else Command10.Caption = panding(m)
        If a(10) = "*" Then Call jieshu
        Command10.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command11_Click()
    m = 11
    If flag = True Then
        Command11.Caption = "!": If Command11.Caption = "!" And f(11) = False Then Command11.Caption = "": f(11) = True: num = num - 1
    ElseIf f(11) = True Then
        If panding(m) = 0 Then Call kuosan Else Command11.Caption = panding(m)
        If a(11) = "*" Then Call jieshu
        Command11.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command12_Click()
    m = 12
    If flag = True Then
        Command12.Caption = "!": If Command12.Caption = "!" And f(12) = False Then Command12.Caption = "": f(12) = True: num = num - 1
    ElseIf f(12) = True Then
        If panding(m) = 0 Then Call kuosan Else Command12.Caption = panding(m)
        If a(12) = "*" Then Call jieshu
        Command12.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command13_Click()
    m = 13
    If flag = True Then
        Command13.Caption = "!": If Command13.Caption = "!" And f(13) = False Then Command13.Caption = "": f(13) = True: num = num - 1
    ElseIf f(13) = True Then
        If panding(m) = 0 Then Call kuosan Else Command13.Caption = panding(m)
        If a(13) = "*" Then Call jieshu
        Command13.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command14_Click()
    m = 14
    If flag = True Then
        Command14.Caption = "!": If Command14.Caption = "!" And f(14) = False Then Command14.Caption = "": f(14) = True: num = num - 1
    ElseIf f(14) = True Then
        If panding(m) = 0 Then Call kuosan Else Command14.Caption = panding(m)
        If a(14) = "*" Then Call jieshu
        Command14.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command15_Click()
    m = 15
    If flag = True Then
        Command15.Caption = "!": If Command15.Caption = "!" And f(15) = False Then Command15.Caption = "": f(15) = True: num = num - 1
    ElseIf f(15) = True Then
        If panding(m) = 0 Then Call kuosan Else Command15.Caption = panding(m)
        If a(15) = "*" Then Call jieshu
        Command15.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command16_Click()
    m = 16
    If flag = True Then
        Command16.Caption = "!": If Command16.Caption = "!" And f(16) = False Then Command16.Caption = "": f(16) = True: num = num - 1
    ElseIf f(16) = True Then
        If panding(m) = 0 Then Call kuosan Else Command16.Caption = panding(m)
        If a(16) = "*" Then Call jieshu
        Command16.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command17_Click()
    m = 17
    If flag = True Then
        Command17.Caption = "!": If Command17.Caption = "!" And f(17) = False Then Command17.Caption = "": f(17) = True: num = num - 1
    ElseIf f(17) = True Then
        If panding(m) = 0 Then Call kuosan Else Command17.Caption = panding(m)
        If a(17) = "*" Then Call jieshu
        Command17.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command18_Click()
    m = 18
    If flag = True Then
        Command18.Caption = "!": If Command18.Caption = "!" And f(18) = False Then Command18.Caption = "": f(18) = True: num = num - 1
    ElseIf f(18) = True Then
        If panding(m) = 0 Then Call kuosan Else Command18.Caption = panding(m)
        If a(18) = "*" Then Call jieshu
        Command18.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command19_Click()
    m = 19
    If flag = True Then
        Command19.Caption = "!": If Command19.Caption = "!" And f(19) = False Then Command19.Caption = "": f(19) = True: num = num - 1
    ElseIf f(19) = True Then
        If panding(m) = 0 Then Call kuosan Else Command19.Caption = panding(m)
        If a(19) = "*" Then Call jieshu
        Command19.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    m = 2
    If flag = True Then
        Command2.Caption = "!": If Command2.Caption = "!" And f(2) = False Then Command2.Caption = "": f(2) = True: num = num - 1
    ElseIf f(2) = True Then
        If panding(m) = 0 Then Call kuosan Else Command2.Caption = panding(m)
        If a(2) = "*" Then Call jieshu
        Command2.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command20_Click()
    m = 20
    If flag = True Then
        Command20.Caption = "!": If Command20.Caption = "!" And f(20) = False Then Command20.Caption = "": f(20) = True: num = num - 1
    ElseIf f(20) = True Then
        If panding(m) = 0 Then Call kuosan Else Command20.Caption = panding(m)
        If a(20) = "*" Then Call jieshu
        Command20.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command21_Click()
    m = 21
    If flag = True Then
        Command21.Caption = "!": If Command21.Caption = "!" And f(21) = False Then Command21.Caption = "": f(21) = True: num = num - 1
    ElseIf f(21) = True Then
        If panding(m) = 0 Then Call kuosan Else Command21.Caption = panding(m)
        If a(21) = "*" Then Call jieshu
        Command21.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command22_Click()
    m = 22
    If flag = True Then
        Command22.Caption = "!": If Command22.Caption = "!" And f(22) = False Then Command22.Caption = "": f(22) = True: num = num - 1
    ElseIf f(22) = True Then
        If panding(m) = 0 Then Call kuosan Else Command22.Caption = panding(m)
        If a(22) = "*" Then Call jieshu
        Command22.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command23_Click()
    m = 23
    If flag = True Then
        Command23.Caption = "!": If Command23.Caption = "!" And f(23) = False Then Command23.Caption = "": f(23) = True: num = num - 1
    ElseIf f(23) = True Then
        If panding(m) = 0 Then Call kuosan Else Command23.Caption = panding(m)
        If a(23) = "*" Then Call jieshu
        Command23.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command24_Click()
    m = 24
    If flag = True Then
        Command24.Caption = "!": If Command24.Caption = "!" And f(24) = False Then Command24.Caption = "": f(24) = True: num = num - 1
    ElseIf f(24) = True Then
        If panding(m) = 0 Then Call kuosan Else Command24.Caption = panding(m)
        If a(24) = "*" Then Call jieshu
        Command24.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command25_Click()
    m = 25
    If flag = True Then
        Command25.Caption = "!": If Command25.Caption = "!" And f(25) = False Then Command25.Caption = "": f(25) = True: num = num - 1
    ElseIf f(25) = True Then
        If panding(m) = 0 Then Call kuosan Else Command25.Caption = panding(m)
        If a(25) = "*" Then Call jieshu
        Command25.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command26_Click()
    m = 26
    If flag = True Then
        Command26.Caption = "!": If Command26.Caption = "!" And f(26) = False Then Command26.Caption = "": f(26) = True: num = num - 1
    ElseIf f(26) = True Then
        If panding(m) = 0 Then Call kuosan Else Command26.Caption = panding(m)
        If a(26) = "*" Then Call jieshu
        Command26.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command27_Click()
    m = 27
    If flag = True Then
        Command27.Caption = "!": If Command27.Caption = "!" And f(27) = False Then Command27.Caption = "": f(27) = True: num = num - 1
    ElseIf f(27) = True Then
        If panding(m) = 0 Then Call kuosan Else Command27.Caption = panding(m)
        If a(27) = "*" Then Call jieshu
        Command27.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command28_Click()
    m = 28
    If flag = True Then
        Command28.Caption = "!": If Command28.Caption = "!" And f(28) = False Then Command28.Caption = "": f(28) = True: num = num - 1
    ElseIf f(28) = True Then
        If panding(m) = 0 Then Call kuosan Else Command28.Caption = panding(m)
        If a(28) = "*" Then Call jieshu
        Command28.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command29_Click()
    m = 29
    If flag = True Then
        Command29.Caption = "!": If Command29.Caption = "!" And f(29) = False Then Command29.Caption = "": f(29) = True: num = num - 1
    ElseIf f(29) = True Then
        If panding(m) = 0 Then Call kuosan Else Command29.Caption = panding(m)
        If a(29) = "*" Then Call jieshu
        Command29.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
    m = 3
    If flag = True Then
        Command3.Caption = "!": If Command3.Caption = "!" And f(3) = False Then Command3.Caption = "": f(3) = True: num = num - 1
    ElseIf f(3) = True Then
        If panding(m) = 0 Then Call kuosan Else Command3.Caption = panding(m)
        If a(3) = "*" Then Call jieshu
        Command3.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command30_Click()
    m = 30
    If flag = True Then
        Command30.Caption = "!": If Command30.Caption = "!" And f(30) = False Then Command30.Caption = "": f(30) = True: num = num - 1
    ElseIf f(30) = True Then
        If panding(m) = 0 Then Call kuosan Else Command30.Caption = panding(m)
        If a(30) = "*" Then Call jieshu
        Command30.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command31_Click()
    m = 31
    If flag = True Then
        Command31.Caption = "!": If Command31.Caption = "!" And f(31) = False Then Command31.Caption = "": f(31) = True: num = num - 1
    ElseIf f(31) = True Then
        If panding(m) = 0 Then Call kuosan Else Command31.Caption = panding(m)
        If a(31) = "*" Then Call jieshu
        Command31.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command32_Click()
    m = 32
    If flag = True Then
        Command32.Caption = "!": If Command32.Caption = "!" And f(32) = False Then Command32.Caption = "": f(32) = True: num = num - 1
    ElseIf f(32) = True Then
        If panding(m) = 0 Then Call kuosan Else Command32.Caption = panding(m)
        If a(32) = "*" Then Call jieshu
        Command32.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command33_Click()
    m = 33
    If flag = True Then
        Command33.Caption = "!": If Command33.Caption = "!" And f(33) = False Then Command33.Caption = "": f(33) = True: num = num - 1
    ElseIf f(33) = True Then
        If panding(m) = 0 Then Call kuosan Else Command33.Caption = panding(m)
        If a(33) = "*" Then Call jieshu
        Command33.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command34_Click()
    m = 34
    If flag = True Then
        Command34.Caption = "!": If Command34.Caption = "!" And f(34) = False Then Command34.Caption = "": f(34) = True: num = num - 1
    ElseIf f(34) = True Then
        If panding(m) = 0 Then Call kuosan Else Command34.Caption = panding(m)
        If a(34) = "*" Then Call jieshu
        Command34.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command35_Click()
    m = 35
    If flag = True Then
        Command35.Caption = "!": If Command35.Caption = "!" And f(35) = False Then Command35.Caption = "": f(35) = True: num = num - 1
    ElseIf f(35) = True Then
        If panding(m) = 0 Then Call kuosan Else Command35.Caption = panding(m)
        If a(35) = "*" Then Call jieshu
        Command35.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command36_Click()
    m = 36
    If flag = True Then
        Command36.Caption = "!": If Command36.Caption = "!" And f(36) = False Then Command36.Caption = "": f(36) = True: num = num - 1
    ElseIf f(36) = True Then
        If panding(m) = 0 Then Call kuosan Else Command36.Caption = panding(m)
        If a(36) = "*" Then Call jieshu
        Command36.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command37_Click()
    m = 37
    If flag = True Then
        Command37.Caption = "!": If Command37.Caption = "!" And f(37) = False Then Command37.Caption = "": f(37) = True: num = num - 1
    ElseIf f(37) = True Then
        If panding(m) = 0 Then Call kuosan Else Command37.Caption = panding(m)
        If a(37) = "*" Then Call jieshu
        Command37.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command38_Click()
    m = 38
    If flag = True Then
        Command38.Caption = "!": If Command38.Caption = "!" And f(38) = False Then Command38.Caption = "": f(38) = True: num = num - 1
    ElseIf f(38) = True Then
        If panding(m) = 0 Then Call kuosan Else Command38.Caption = panding(m)
        If a(38) = "*" Then Call jieshu
        Command38.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command39_Click()
    m = 39
    If flag = True Then
        Command39.Caption = "!": If Command39.Caption = "!" And f(39) = False Then Command39.Caption = "": f(39) = True: num = num - 1
    ElseIf f(39) = True Then
        If panding(m) = 0 Then Call kuosan Else Command39.Caption = panding(m)
        If a(39) = "*" Then Call jieshu
        Command39.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command4_Click()
    m = 4
    If flag = True Then
        Command4.Caption = "!": If Command4.Caption = "!" And f(4) = False Then Command4.Caption = "": f(4) = True: num = num - 1
    ElseIf f(4) = True Then
        If panding(m) = 0 Then Call kuosan Else Command4.Caption = panding(m)
        If a(4) = "*" Then Call jieshu
        Command4.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command40_Click()
    m = 40
    If flag = True Then
        Command40.Caption = "!": If Command40.Caption = "!" And f(40) = False Then Command40.Caption = "": f(40) = True: num = num - 1
    ElseIf f(40) = True Then
        If panding(m) = 0 Then Call kuosan Else Command40.Caption = panding(m)
        If a(40) = "*" Then Call jieshu
        Command40.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command41_Click()
    m = 41
    If flag = True Then
        Command41.Caption = "!": If Command41.Caption = "!" And f(41) = False Then Command41.Caption = "": f(41) = True: num = num - 1
    ElseIf f(41) = True Then
        If panding(m) = 0 Then Call kuosan Else Command41.Caption = panding(m)
        If a(41) = "*" Then Call jieshu
        Command41.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command42_Click()
    m = 42
    If flag = True Then
        Command42.Caption = "!": If Command42.Caption = "!" And f(42) = False Then Command42.Caption = "": f(42) = True: num = num - 1
    ElseIf f(42) = True Then
        If panding(m) = 0 Then Call kuosan Else Command42.Caption = panding(m)
        If a(42) = "*" Then Call jieshu
        Command42.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command43_Click()
    m = 43
    If flag = True Then
        Command43.Caption = "!": If Command43.Caption = "!" And f(43) = False Then Command43.Caption = "": f(43) = True: num = num - 1
    ElseIf f(43) = True Then
        If panding(m) = 0 Then Call kuosan Else Command43.Caption = panding(m)
        If a(43) = "*" Then Call jieshu
        Command43.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command44_Click()
    m = 44
    If flag = True Then
        Command44.Caption = "!": If Command44.Caption = "!" And f(44) = False Then Command44.Caption = "": f(44) = True: num = num - 1
    ElseIf f(44) = True Then
        If panding(m) = 0 Then Call kuosan Else Command44.Caption = panding(m)
        If a(44) = "*" Then Call jieshu
        Command44.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command45_Click()
    m = 45
    If flag = True Then
        Command45.Caption = "!": If Command45.Caption = "!" And f(45) = False Then Command45.Caption = "": f(45) = True: num = num - 1
    ElseIf f(45) = True Then
        If panding(m) = 0 Then Call kuosan Else Command45.Caption = panding(m)
        If a(45) = "*" Then Call jieshu
        Command45.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command46_Click()
    m = 46
    If flag = True Then
        Command46.Caption = "!": If Command46.Caption = "!" And f(46) = False Then Command46.Caption = "": f(46) = True: num = num - 1
    ElseIf f(46) = True Then
        If panding(m) = 0 Then Call kuosan Else Command46.Caption = panding(m)
        If a(46) = "*" Then Call jieshu
        Command46.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command47_Click()
    m = 47
    If flag = True Then
        Command47.Caption = "!": If Command47.Caption = "!" And f(47) = False Then Command47.Caption = "": f(47) = True: num = num - 1
    ElseIf f(47) = True Then
        If panding(m) = 0 Then Call kuosan Else Command47.Caption = panding(m)
        If a(47) = "*" Then Call jieshu
        Command47.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command48_Click()
    m = 48
    If flag = True Then
        Command48.Caption = "!": If Command48.Caption = "!" And f(48) = False Then Command48.Caption = "": f(48) = True: num = num - 1
    ElseIf f(48) = True Then
        If panding(m) = 0 Then Call kuosan Else Command48.Caption = panding(m)
        If a(48) = "*" Then Call jieshu
        Command48.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command49_Click()
    m = 49
    If flag = True Then
        Command49.Caption = "!": If Command49.Caption = "!" And f(49) = False Then Command49.Caption = "": f(49) = True: num = num - 1
    ElseIf f(49) = True Then
        If panding(m) = 0 Then Call kuosan Else Command49.Caption = panding(m)
        If a(49) = "*" Then Call jieshu
        Command49.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command5_Click()
    m = 5
    If flag = True Then
        Command5.Caption = "!": If Command5.Caption = "!" And f(5) = False Then Command5.Caption = "": f(5) = True: num = num - 1
    ElseIf f(5) = True Then
        If panding(m) = 0 Then Call kuosan Else Command5.Caption = panding(m)
        If a(5) = "*" Then Call jieshu
        Command5.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command50_Click()
    m = 50
    If flag = True Then
        Command50.Caption = "!": If Command50.Caption = "!" And f(50) = False Then Command50.Caption = "": f(50) = True: num = num - 1
    ElseIf f(50) = True Then
        If panding(m) = 0 Then Call kuosan Else Command50.Caption = panding(m)
        If a(50) = "*" Then Call jieshu
        Command50.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command51_Click()
    m = 51
    If flag = True Then
        Command51.Caption = "!": If Command51.Caption = "!" And f(51) = False Then Command51.Caption = "": f(51) = True: num = num - 1
    ElseIf f(51) = True Then
        If panding(m) = 0 Then Call kuosan Else Command51.Caption = panding(m)
        If a(51) = "*" Then Call jieshu
        Command51.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command52_Click()
    m = 52
    If flag = True Then
        Command52.Caption = "!": If Command52.Caption = "!" And f(52) = False Then Command52.Caption = "": f(52) = True: num = num - 1
    ElseIf f(52) = True Then
        If panding(m) = 0 Then Call kuosan Else Command52.Caption = panding(m)
        If a(52) = "*" Then Call jieshu
        Command52.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command53_Click()
    m = 53
    If flag = True Then
        Command53.Caption = "!": If Command53.Caption = "!" And f(53) = False Then Command53.Caption = "": f(53) = True: num = num - 1
    ElseIf f(53) = True Then
        If panding(m) = 0 Then Call kuosan Else Command53.Caption = panding(m)
        If a(53) = "*" Then Call jieshu
        Command53.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command54_Click()
    m = 54
    If flag = True Then
        Command54.Caption = "!": If Command54.Caption = "!" And f(54) = False Then Command54.Caption = "": f(54) = True: num = num - 1
    ElseIf f(54) = True Then
        If panding(m) = 0 Then Call kuosan Else Command54.Caption = panding(m)
        If a(54) = "*" Then Call jieshu
        Command54.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command55_Click()
    m = 55
    If flag = True Then
        Command55.Caption = "!": If Command55.Caption = "!" And f(55) = False Then Command55.Caption = "": f(55) = True: num = num - 1
    ElseIf f(55) = True Then
        If panding(m) = 0 Then Call kuosan Else Command55.Caption = panding(m)
        If a(55) = "*" Then Call jieshu
        Command55.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command56_Click()
    m = 56
    If flag = True Then
        Command56.Caption = "!": If Command56.Caption = "!" And f(56) = False Then Command56.Caption = "": f(56) = True: num = num - 1
    ElseIf f(56) = True Then
        If panding(m) = 0 Then Call kuosan Else Command56.Caption = panding(m)
        If a(56) = "*" Then Call jieshu
        Command56.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command57_Click()
    m = 57
    If flag = True Then
        Command57.Caption = "!": If Command57.Caption = "!" And f(57) = False Then Command57.Caption = "": f(57) = True: num = num - 1
    ElseIf f(57) = True Then
        If panding(m) = 0 Then Call kuosan Else Command57.Caption = panding(m)
        If a(57) = "*" Then Call jieshu
        Command57.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command58_Click()
    m = 58
    If flag = True Then
        Command58.Caption = "!": If Command58.Caption = "!" And f(58) = False Then Command58.Caption = "": f(58) = True: num = num - 1
    ElseIf f(58) = True Then
        If panding(m) = 0 Then Call kuosan Else Command58.Caption = panding(m)
        If a(58) = "*" Then Call jieshu
        Command58.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command59_Click()
    m = 59
    If flag = True Then
        Command59.Caption = "!": If Command59.Caption = "!" And f(59) = False Then Command59.Caption = "": f(59) = True: num = num - 1
    ElseIf f(59) = True Then
        If panding(m) = 0 Then Call kuosan Else Command59.Caption = panding(m)
        If a(59) = "*" Then Call jieshu
        Command59.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command6_Click()
    m = 6
    If flag = True Then
        Command6.Caption = "!": If Command6.Caption = "!" And f(6) = False Then Command6.Caption = "": f(6) = True: num = num - 1
    ElseIf f(6) = True Then
        If panding(m) = 0 Then Call kuosan Else Command6.Caption = panding(m)
        If a(6) = "*" Then Call jieshu
        Command6.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command60_Click()
    m = 60
    If flag = True Then
        Command60.Caption = "!": If Command60.Caption = "!" And f(60) = False Then Command60.Caption = "": f(60) = True: num = num - 1
    ElseIf f(60) = True Then
        If panding(m) = 0 Then Call kuosan Else Command60.Caption = panding(m)
        If a(60) = "*" Then Call jieshu
        Command60.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command61_Click()
    m = 61
    If flag = True Then
        Command61.Caption = "!": If Command61.Caption = "!" And f(61) = False Then Command61.Caption = "": f(61) = True: num = num - 1
    ElseIf f(61) = True Then
        If panding(m) = 0 Then Call kuosan Else Command61.Caption = panding(m)
        If a(61) = "*" Then Call jieshu
        Command61.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command62_Click()
    m = 62
    If flag = True Then
        Command62.Caption = "!": If Command62.Caption = "!" And f(62) = False Then Command62.Caption = "": f(62) = True: num = num - 1
    ElseIf f(62) = True Then
        If panding(m) = 0 Then Call kuosan Else Command62.Caption = panding(m)
        If a(62) = "*" Then Call jieshu
        Command62.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command63_Click()
    m = 63
    If flag = True Then
        Command63.Caption = "!": If Command63.Caption = "!" And f(63) = False Then Command63.Caption = "": f(63) = True: num = num - 1
    ElseIf f(63) = True Then
        If panding(m) = 0 Then Call kuosan Else Command63.Caption = panding(m)
        If a(63) = "*" Then Call jieshu
        Command63.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command64_Click()
    m = 64
    If flag = True Then
        Command64.Caption = "!": If Command64.Caption = "!" And f(64) = False Then Command64.Caption = "": f(64) = True: num = num - 1
    ElseIf f(64) = True Then
        If panding(m) = 0 Then Call kuosan Else Command64.Caption = panding(m)
        If a(64) = "*" Then Call jieshu
        Command64.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command65_Click()
    m = 65
    If flag = True Then
        Command65.Caption = "!": If Command65.Caption = "!" And f(65) = False Then Command65.Caption = "": f(65) = True: num = num - 1
    ElseIf f(65) = True Then
        If panding(m) = 0 Then Call kuosan Else Command65.Caption = panding(m)
        If a(65) = "*" Then Call jieshu
        Command65.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command66_Click()
    m = 66
    If flag = True Then
        Command66.Caption = "!": If Command66.Caption = "!" And f(66) = False Then Command66.Caption = "": f(66) = True: num = num - 1
    ElseIf f(66) = True Then
        If panding(m) = 0 Then Call kuosan Else Command66.Caption = panding(m)
        If a(66) = "*" Then Call jieshu
        Command66.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command67_Click()
    m = 67
    If flag = True Then
        Command67.Caption = "!": If Command67.Caption = "!" And f(67) = False Then Command67.Caption = "": f(67) = True: num = num - 1
    ElseIf f(67) = True Then
        If panding(m) = 0 Then Call kuosan Else Command67.Caption = panding(m)
        If a(67) = "*" Then Call jieshu
        Command67.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command68_Click()
    m = 68
    If flag = True Then
        Command68.Caption = "!": If Command68.Caption = "!" And f(68) = False Then Command68.Caption = "": f(68) = True: num = num - 1
    ElseIf f(68) = True Then
        If panding(m) = 0 Then Call kuosan Else Command68.Caption = panding(m)
        If a(68) = "*" Then Call jieshu
        Command68.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command69_Click()
    m = 69
    If flag = True Then
        Command69.Caption = "!": If Command69.Caption = "!" And f(69) = False Then Command69.Caption = "": f(69) = True: num = num - 1
    ElseIf f(69) = True Then
        If panding(m) = 0 Then Call kuosan Else Command69.Caption = panding(m)
        If a(69) = "*" Then Call jieshu
        Command69.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command7_Click()
    m = 7
    If flag = True Then
        Command7.Caption = "!": If Command7.Caption = "!" And f(7) = False Then Command7.Caption = "": f(7) = True: num = num - 1
    ElseIf f(7) = True Then
        If panding(m) = 0 Then Call kuosan Else Command7.Caption = panding(m)
        If a(7) = "*" Then Call jieshu
        Command7.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command70_Click()
    m = 70
    If flag = True Then
        Command70.Caption = "!": If Command70.Caption = "!" And f(70) = False Then Command70.Caption = "": f(70) = True: num = num - 1
    ElseIf f(70) = True Then
        If panding(m) = 0 Then Call kuosan Else Command70.Caption = panding(m)
        If a(70) = "*" Then Call jieshu
        Command70.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command71_Click()
    m = 71
    If flag = True Then
        Command71.Caption = "!": If Command71.Caption = "!" And f(71) = False Then Command71.Caption = "": f(71) = True: num = num - 1
    ElseIf f(71) = True Then
        If panding(m) = 0 Then Call kuosan Else Command71.Caption = panding(m)
        If a(71) = "*" Then Call jieshu
        Command71.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command72_Click()
    m = 72
    If flag = True Then
        Command72.Caption = "!": If Command72.Caption = "!" And f(72) = False Then Command72.Caption = "": f(72) = True: num = num - 1
    ElseIf f(72) = True Then
        If panding(m) = 0 Then Call kuosan Else Command72.Caption = panding(m)
        If a(72) = "*" Then Call jieshu
        Command72.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command73_Click()
    m = 73
    If flag = True Then
        Command73.Caption = "!": If Command73.Caption = "!" And f(73) = False Then Command73.Caption = "": f(73) = True: num = num - 1
    ElseIf f(73) = True Then
        If panding(m) = 0 Then Call kuosan Else Command73.Caption = panding(m)
        If a(73) = "*" Then Call jieshu
        Command73.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command74_Click()
    m = 74
    If flag = True Then
        Command74.Caption = "!": If Command74.Caption = "!" And f(74) = False Then Command74.Caption = "": f(74) = True: num = num - 1
    ElseIf f(74) = True Then
        If panding(m) = 0 Then Call kuosan Else Command74.Caption = panding(m)
        If a(74) = "*" Then Call jieshu
        Command74.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command75_Click()
    m = 75
    If flag = True Then
        Command75.Caption = "!": If Command75.Caption = "!" And f(75) = False Then Command75.Caption = "": f(75) = True: num = num - 1
    ElseIf f(75) = True Then
        If panding(m) = 0 Then Call kuosan Else Command75.Caption = panding(m)
        If a(75) = "*" Then Call jieshu
        Command75.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command76_Click()
    m = 76
    If flag = True Then
        Command76.Caption = "!": If Command76.Caption = "!" And f(76) = False Then Command76.Caption = "": f(76) = True: num = num - 1
    ElseIf f(76) = True Then
        If panding(m) = 0 Then Call kuosan Else Command76.Caption = panding(m)
        If a(76) = "*" Then Call jieshu
        Command76.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command77_Click()
    m = 77
    If flag = True Then
        Command77.Caption = "!": If Command77.Caption = "!" And f(77) = False Then Command77.Caption = "": f(77) = True: num = num - 1
    ElseIf f(77) = True Then
        If panding(m) = 0 Then Call kuosan Else Command77.Caption = panding(m)
        If a(77) = "*" Then Call jieshu
        Command77.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command78_Click()
    m = 78
    If flag = True Then
        Command78.Caption = "!": If Command78.Caption = "!" And f(78) = False Then Command78.Caption = "": f(78) = True: num = num - 1
    ElseIf f(78) = True Then
        If panding(m) = 0 Then Call kuosan Else Command78.Caption = panding(m)
        If a(78) = "*" Then Call jieshu
        Command78.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command79_Click()
    m = 79
    If flag = True Then
        Command79.Caption = "!": If Command79.Caption = "!" And f(79) = False Then Command79.Caption = "": f(79) = True: num = num - 1
    ElseIf f(79) = True Then
        If panding(m) = 0 Then Call kuosan Else Command79.Caption = panding(m)
        If a(79) = "*" Then Call jieshu
        Command79.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command8_Click()
    m = 8
    If flag = True Then
        Command8.Caption = "!": If Command8.Caption = "!" And f(8) = False Then Command8.Caption = "": f(8) = True: num = num - 1
    ElseIf f(8) = True Then
        If panding(m) = 0 Then Call kuosan Else Command8.Caption = panding(m)
        If a(8) = "*" Then Call jieshu
        Command8.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command80_Click()
Dim m As Integer
    m = 80
    If flag = True Then
        Command80.Caption = "!": If Command80.Caption = "!" And f(80) = False Then Command80.Caption = "": f(80) = True: num = num - 1
    ElseIf f(80) = True Then
        If panding(m) = 0 Then Call kuosan Else Command80.Caption = panding(m)
        If a(80) = "*" Then Call jieshu
        Command80.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command81_Click()
Dim m As Integer
    m = 81
    If flag = True Then
        Command81.Caption = "!": If Command81.Caption = "!" And f(81) = False Then Command81.Caption = "": f(81) = True: num = num - 1
    ElseIf f(81) = True Then
        If panding(m) = 0 Then Call kuosan Else Command81.Caption = panding(m)
        If a(81) = "*" Then Call jieshu
        Command81.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Command9_Click()
Dim m As Integer
    m = 9
    If flag = True Then
        Command9.Caption = "!": If Command9.Caption = "!" And f(9) = False Then Command9.Caption = "": f(9) = True: num = num - 1
    ElseIf f(9) = True Then
        If panding(m) = 0 Then Call kuosan Else Command9.Caption = CStr(panding(m))
        If a(9) = "*" Then Call jieshu
        Command9.Enabled = False
    End If
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    Call chushi1
End Sub

Private Sub jieshu()
    If a(1) = "*" Then Command1.Caption = a(1): Command1.Enabled = True
    If a(2) = "*" Then Command2.Caption = a(2): Command2.Enabled = True
    If a(3) = "*" Then Command3.Caption = a(3): Command3.Enabled = True
    If a(4) = "*" Then Command4.Caption = a(4): Command4.Enabled = True
    If a(5) = "*" Then Command5.Caption = a(5): Command5.Enabled = True
    If a(6) = "*" Then Command6.Caption = a(6): Command6.Enabled = True
    If a(7) = "*" Then Command7.Caption = a(7): Command7.Enabled = True
    If a(8) = "*" Then Command8.Caption = a(8): Command8.Enabled = True
    If a(9) = "*" Then Command9.Caption = a(9): Command9.Enabled = True
    If a(10) = "*" Then Command10.Caption = a(10): Command10.Enabled = True
    If a(11) = "*" Then Command11.Caption = a(11): Command11.Enabled = True
    If a(12) = "*" Then Command12.Caption = a(12): Command12.Enabled = True
    If a(13) = "*" Then Command13.Caption = a(13): Command13.Enabled = True
    If a(14) = "*" Then Command14.Caption = a(14): Command14.Enabled = True
    If a(15) = "*" Then Command15.Caption = a(15): Command15.Enabled = True
    If a(16) = "*" Then Command16.Caption = a(16): Command16.Enabled = True
    If a(17) = "*" Then Command17.Caption = a(17): Command17.Enabled = True
    If a(18) = "*" Then Command18.Caption = a(18): Command18.Enabled = True
    If a(19) = "*" Then Command19.Caption = a(19): Command19.Enabled = True
    If a(20) = "*" Then Command20.Caption = a(20): Command20.Enabled = True
    If a(21) = "*" Then Command21.Caption = a(21): Command21.Enabled = True
    If a(22) = "*" Then Command22.Caption = a(22): Command22.Enabled = True
    If a(23) = "*" Then Command23.Caption = a(23): Command23.Enabled = True
    If a(24) = "*" Then Command24.Caption = a(24): Command24.Enabled = True
    If a(25) = "*" Then Command25.Caption = a(25): Command25.Enabled = True
    If a(26) = "*" Then Command26.Caption = a(26): Command26.Enabled = True
    If a(27) = "*" Then Command27.Caption = a(27): Command27.Enabled = True
    If a(28) = "*" Then Command28.Caption = a(28): Command28.Enabled = True
    If a(29) = "*" Then Command29.Caption = a(29): Command29.Enabled = True
    If a(30) = "*" Then Command30.Caption = a(30): Command30.Enabled = True
    If a(31) = "*" Then Command31.Caption = a(31): Command31.Enabled = True
    If a(32) = "*" Then Command32.Caption = a(32): Command32.Enabled = True
    If a(33) = "*" Then Command33.Caption = a(33): Command33.Enabled = True
    If a(34) = "*" Then Command34.Caption = a(34): Command34.Enabled = True
    If a(35) = "*" Then Command35.Caption = a(35): Command35.Enabled = True
    If a(36) = "*" Then Command36.Caption = a(36): Command36.Enabled = True
    If a(37) = "*" Then Command37.Caption = a(37): Command37.Enabled = True
    If a(38) = "*" Then Command38.Caption = a(38): Command38.Enabled = True
    If a(39) = "*" Then Command39.Caption = a(39): Command39.Enabled = True
    If a(40) = "*" Then Command40.Caption = a(40): Command40.Enabled = True
    If a(41) = "*" Then Command41.Caption = a(41): Command41.Enabled = True
    If a(42) = "*" Then Command42.Caption = a(42): Command42.Enabled = True
    If a(43) = "*" Then Command43.Caption = a(43): Command43.Enabled = True
    If a(44) = "*" Then Command44.Caption = a(44): Command44.Enabled = True
    If a(45) = "*" Then Command45.Caption = a(45): Command45.Enabled = True
    If a(46) = "*" Then Command46.Caption = a(46): Command46.Enabled = True
    If a(47) = "*" Then Command47.Caption = a(47): Command47.Enabled = True
    If a(48) = "*" Then Command48.Caption = a(48): Command48.Enabled = True
    If a(49) = "*" Then Command49.Caption = a(49): Command49.Enabled = True
    If a(50) = "*" Then Command50.Caption = a(50): Command50.Enabled = True
    If a(51) = "*" Then Command51.Caption = a(51): Command51.Enabled = True
    If a(52) = "*" Then Command52.Caption = a(52): Command52.Enabled = True
    If a(53) = "*" Then Command53.Caption = a(53): Command53.Enabled = True
    If a(54) = "*" Then Command54.Caption = a(54): Command54.Enabled = True
    If a(55) = "*" Then Command55.Caption = a(55): Command55.Enabled = True
    If a(56) = "*" Then Command56.Caption = a(56): Command56.Enabled = True
    If a(57) = "*" Then Command57.Caption = a(57): Command57.Enabled = True
    If a(58) = "*" Then Command58.Caption = a(58): Command58.Enabled = True
    If a(59) = "*" Then Command59.Caption = a(59): Command59.Enabled = True
    If a(60) = "*" Then Command60.Caption = a(60): Command60.Enabled = True
    If a(61) = "*" Then Command61.Caption = a(61): Command61.Enabled = True
    If a(62) = "*" Then Command62.Caption = a(62): Command62.Enabled = True
    If a(63) = "*" Then Command63.Caption = a(63): Command63.Enabled = True
    If a(64) = "*" Then Command64.Caption = a(64): Command64.Enabled = True
    If a(65) = "*" Then Command65.Caption = a(65): Command65.Enabled = True
    If a(66) = "*" Then Command66.Caption = a(66): Command66.Enabled = True
    If a(67) = "*" Then Command67.Caption = a(67): Command67.Enabled = True
    If a(68) = "*" Then Command68.Caption = a(68): Command68.Enabled = True
    If a(69) = "*" Then Command69.Caption = a(69): Command69.Enabled = True
    If a(70) = "*" Then Command70.Caption = a(70): Command70.Enabled = True
    If a(71) = "*" Then Command71.Caption = a(71): Command71.Enabled = True
    If a(72) = "*" Then Command72.Caption = a(72): Command72.Enabled = True
    If a(73) = "*" Then Command73.Caption = a(73): Command73.Enabled = True
    If a(74) = "*" Then Command74.Caption = a(74): Command74.Enabled = True
    If a(75) = "*" Then Command75.Caption = a(75): Command75.Enabled = True
    If a(76) = "*" Then Command76.Caption = a(76): Command76.Enabled = True
    If a(77) = "*" Then Command77.Caption = a(77): Command77.Enabled = True
    If a(78) = "*" Then Command78.Caption = a(78): Command78.Enabled = True
    If a(79) = "*" Then Command79.Caption = a(79): Command79.Enabled = True
    If a(80) = "*" Then Command80.Caption = a(80): Command80.Enabled = True
    If a(81) = "*" Then Command81.Caption = a(81): Command81.Enabled = True
    If Timer2.Enabled = False Then Timer2.Enabled = True
    Check1.Enabled = False
    For i = 1 To 81
        f(i) = False
    Next i
End Sub
Private Sub chushi1()
    Randomize
    For i = 1 To 81
        a(i) = ""
        f(i) = True
    Next i
    For i = 1 To 10
        n = Int(Rnd * 81 + 1)
        If a(n) = "*" Then i = i - 1
        a(n) = "*"
    Next i
    Command1.Tag = a(1)
    Command2.Tag = a(2)
    Command3.Tag = a(3)
    Command4.Tag = a(4)
    Command5.Tag = a(5)
    Command6.Tag = a(6)
    Command7.Tag = a(7)
    Command8.Tag = a(8)
    Command9.Tag = a(9)
    Command10.Tag = a(10)
    Command11.Tag = a(11)
    Command12.Tag = a(12)
    Command13.Tag = a(13)
    Command14.Tag = a(14)
    Command15.Tag = a(15)
    Command16.Tag = a(16)
    Command17.Tag = a(17)
    Command18.Tag = a(18)
    Command19.Tag = a(19)
    Command20.Tag = a(20)
    Command21.Tag = a(21)
    Command22.Tag = a(22)
    Command23.Tag = a(23)
    Command24.Tag = a(24)
    Command25.Tag = a(25)
    Command26.Tag = a(26)
    Command27.Tag = a(27)
    Command28.Tag = a(28)
    Command29.Tag = a(29)
    Command30.Tag = a(30)
    Command31.Tag = a(31)
    Command32.Tag = a(32)
    Command33.Tag = a(33)
    Command34.Tag = a(34)
    Command35.Tag = a(35)
    Command36.Tag = a(36)
    Command37.Tag = a(37)
    Command38.Tag = a(38)
    Command39.Tag = a(39)
    Command40.Tag = a(40)
    Command41.Tag = a(41)
    Command42.Tag = a(42)
    Command43.Tag = a(43)
    Command44.Tag = a(44)
    Command45.Tag = a(45)
    Command46.Tag = a(46)
    Command47.Tag = a(47)
    Command48.Tag = a(48)
    Command49.Tag = a(49)
    Command50.Tag = a(50)
    Command51.Tag = a(51)
    Command52.Tag = a(52)
    Command53.Tag = a(53)
    Command54.Tag = a(54)
    Command55.Tag = a(55)
    Command56.Tag = a(56)
    Command57.Tag = a(57)
    Command58.Tag = a(58)
    Command59.Tag = a(59)
    Command60.Tag = a(60)
    Command61.Tag = a(61)
    Command62.Tag = a(62)
    Command63.Tag = a(63)
    Command64.Tag = a(64)
    Command65.Tag = a(65)
    Command66.Tag = a(66)
    Command67.Tag = a(67)
    Command68.Tag = a(68)
    Command69.Tag = a(69)
    Command70.Tag = a(70)
    Command71.Tag = a(71)
    Command72.Tag = a(72)
    Command73.Tag = a(73)
    Command74.Tag = a(74)
    Command75.Tag = a(75)
    Command76.Tag = a(76)
    Command77.Tag = a(77)
    Command78.Tag = a(78)
    Command79.Tag = a(79)
    Command80.Tag = a(80)
    Command81.Tag = a(81)
    Command1.Caption = ""
    Command2.Caption = ""
    Command3.Caption = ""
    Command4.Caption = ""
    Command5.Caption = ""
    Command6.Caption = ""
    Command7.Caption = ""
    Command8.Caption = ""
    Command9.Caption = ""
    Command10.Caption = ""
    Command11.Caption = ""
    Command12.Caption = ""
    Command13.Caption = ""
    Command14.Caption = ""
    Command15.Caption = ""
    Command16.Caption = ""
    Command17.Caption = ""
    Command18.Caption = ""
    Command19.Caption = ""
    Command20.Caption = ""
    Command21.Caption = ""
    Command22.Caption = ""
    Command23.Caption = ""
    Command24.Caption = ""
    Command25.Caption = ""
    Command26.Caption = ""
    Command27.Caption = ""
    Command28.Caption = ""
    Command29.Caption = ""
    Command30.Caption = ""
    Command31.Caption = ""
    Command32.Caption = ""
    Command33.Caption = ""
    Command34.Caption = ""
    Command35.Caption = ""
    Command36.Caption = ""
    Command37.Caption = ""
    Command38.Caption = ""
    Command39.Caption = ""
    Command40.Caption = ""
    Command41.Caption = ""
    Command42.Caption = ""
    Command43.Caption = ""
    Command44.Caption = ""
    Command45.Caption = ""
    Command46.Caption = ""
    Command47.Caption = ""
    Command48.Caption = ""
    Command49.Caption = ""
    Command50.Caption = ""
    Command51.Caption = ""
    Command52.Caption = ""
    Command53.Caption = ""
    Command54.Caption = ""
    Command55.Caption = ""
    Command56.Caption = ""
    Command57.Caption = ""
    Command58.Caption = ""
    Command59.Caption = ""
    Command60.Caption = ""
    Command61.Caption = ""
    Command62.Caption = ""
    Command63.Caption = ""
    Command64.Caption = ""
    Command65.Caption = ""
    Command66.Caption = ""
    Command67.Caption = ""
    Command68.Caption = ""
    Command69.Caption = ""
    Command70.Caption = ""
    Command71.Caption = ""
    Command72.Caption = ""
    Command73.Caption = ""
    Command74.Caption = ""
    Command75.Caption = ""
    Command76.Caption = ""
    Command77.Caption = ""
    Command78.Caption = ""
    Command79.Caption = ""
    Command80.Caption = ""
    Command81.Caption = ""
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True
    Command9.Enabled = True
    Command10.Enabled = True
    Command11.Enabled = True
    Command12.Enabled = True
    Command13.Enabled = True
    Command14.Enabled = True
    Command15.Enabled = True
    Command16.Enabled = True
    Command17.Enabled = True
    Command18.Enabled = True
    Command19.Enabled = True
    Command20.Enabled = True
    Command21.Enabled = True
    Command22.Enabled = True
    Command23.Enabled = True
    Command24.Enabled = True
    Command25.Enabled = True
    Command26.Enabled = True
    Command27.Enabled = True
    Command28.Enabled = True
    Command29.Enabled = True
    Command30.Enabled = True
    Command31.Enabled = True
    Command32.Enabled = True
    Command33.Enabled = True
    Command34.Enabled = True
    Command35.Enabled = True
    Command36.Enabled = True
    Command37.Enabled = True
    Command38.Enabled = True
    Command39.Enabled = True
    Command40.Enabled = True
    Command41.Enabled = True
    Command42.Enabled = True
    Command43.Enabled = True
    Command44.Enabled = True
    Command45.Enabled = True
    Command46.Enabled = True
    Command47.Enabled = True
    Command48.Enabled = True
    Command49.Enabled = True
    Command50.Enabled = True
    Command51.Enabled = True
    Command52.Enabled = True
    Command53.Enabled = True
    Command54.Enabled = True
    Command55.Enabled = True
    Command56.Enabled = True
    Command57.Enabled = True
    Command58.Enabled = True
    Command59.Enabled = True
    Command60.Enabled = True
    Command61.Enabled = True
    Command62.Enabled = True
    Command63.Enabled = True
    Command64.Enabled = True
    Command65.Enabled = True
    Command66.Enabled = True
    Command67.Enabled = True
    Command68.Enabled = True
    Command69.Enabled = True
    Command70.Enabled = True
    Command71.Enabled = True
    Command72.Enabled = True
    Command73.Enabled = True
    Command74.Enabled = True
    Command75.Enabled = True
    Command76.Enabled = True
    Command77.Enabled = True
    Command78.Enabled = True
    Command79.Enabled = True
    Command80.Enabled = True
    Command81.Enabled = True
    t = 0
    Label1.Caption = "0"
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = True
    Timer4.Enabled = True
    Check1.Enabled = True
    Check1.Value = 0
    num = 0
End Sub
Private Sub bj()
    If Command1.Caption = "!" And f(1) = True Then num = num + 1: f(1) = False
    If Command2.Caption = "!" And f(2) = True Then num = num + 1: f(2) = False
    If Command3.Caption = "!" And f(3) = True Then num = num + 1: f(3) = False
    If Command4.Caption = "!" And f(4) = True Then num = num + 1: f(4) = False
    If Command5.Caption = "!" And f(5) = True Then num = num + 1: f(5) = False
    If Command6.Caption = "!" And f(6) = True Then num = num + 1: f(6) = False
    If Command7.Caption = "!" And f(7) = True Then num = num + 1: f(7) = False
    If Command8.Caption = "!" And f(8) = True Then num = num + 1: f(8) = False
    If Command9.Caption = "!" And f(9) = True Then num = num + 1: f(9) = False
    If Command10.Caption = "!" And f(10) = True Then num = num + 1: f(10) = False
    If Command11.Caption = "!" And f(11) = True Then num = num + 1: f(11) = False
    If Command12.Caption = "!" And f(12) = True Then num = num + 1: f(12) = False
    If Command13.Caption = "!" And f(13) = True Then num = num + 1: f(13) = False
    If Command14.Caption = "!" And f(14) = True Then num = num + 1: f(14) = False
    If Command15.Caption = "!" And f(15) = True Then num = num + 1: f(15) = False
    If Command16.Caption = "!" And f(16) = True Then num = num + 1: f(16) = False
    If Command17.Caption = "!" And f(17) = True Then num = num + 1: f(17) = False
    If Command18.Caption = "!" And f(18) = True Then num = num + 1: f(18) = False
    If Command19.Caption = "!" And f(19) = True Then num = num + 1: f(19) = False
    If Command20.Caption = "!" And f(20) = True Then num = num + 1: f(20) = False
    If Command21.Caption = "!" And f(21) = True Then num = num + 1: f(21) = False
    If Command22.Caption = "!" And f(22) = True Then num = num + 1: f(22) = False
    If Command23.Caption = "!" And f(23) = True Then num = num + 1: f(23) = False
    If Command24.Caption = "!" And f(24) = True Then num = num + 1: f(24) = False
    If Command25.Caption = "!" And f(25) = True Then num = num + 1: f(25) = False
    If Command26.Caption = "!" And f(26) = True Then num = num + 1: f(26) = False
    If Command27.Caption = "!" And f(27) = True Then num = num + 1: f(27) = False
    If Command28.Caption = "!" And f(28) = True Then num = num + 1: f(28) = False
    If Command29.Caption = "!" And f(29) = True Then num = num + 1: f(29) = False
    If Command30.Caption = "!" And f(30) = True Then num = num + 1: f(30) = False
    If Command31.Caption = "!" And f(31) = True Then num = num + 1: f(31) = False
    If Command32.Caption = "!" And f(32) = True Then num = num + 1: f(32) = False
    If Command33.Caption = "!" And f(33) = True Then num = num + 1: f(33) = False
    If Command34.Caption = "!" And f(34) = True Then num = num + 1: f(34) = False
    If Command35.Caption = "!" And f(35) = True Then num = num + 1: f(35) = False
    If Command36.Caption = "!" And f(36) = True Then num = num + 1: f(36) = False
    If Command37.Caption = "!" And f(37) = True Then num = num + 1: f(37) = False
    If Command38.Caption = "!" And f(38) = True Then num = num + 1: f(38) = False
    If Command39.Caption = "!" And f(39) = True Then num = num + 1: f(39) = False
    If Command40.Caption = "!" And f(40) = True Then num = num + 1: f(40) = False
    If Command41.Caption = "!" And f(41) = True Then num = num + 1: f(41) = False
    If Command42.Caption = "!" And f(42) = True Then num = num + 1: f(42) = False
    If Command43.Caption = "!" And f(43) = True Then num = num + 1: f(43) = False
    If Command44.Caption = "!" And f(44) = True Then num = num + 1: f(44) = False
    If Command45.Caption = "!" And f(45) = True Then num = num + 1: f(45) = False
    If Command46.Caption = "!" And f(46) = True Then num = num + 1: f(46) = False
    If Command47.Caption = "!" And f(47) = True Then num = num + 1: f(47) = False
    If Command48.Caption = "!" And f(48) = True Then num = num + 1: f(48) = False
    If Command49.Caption = "!" And f(49) = True Then num = num + 1: f(49) = False
    If Command50.Caption = "!" And f(50) = True Then num = num + 1: f(50) = False
    If Command51.Caption = "!" And f(51) = True Then num = num + 1: f(51) = False
    If Command52.Caption = "!" And f(52) = True Then num = num + 1: f(52) = False
    If Command53.Caption = "!" And f(53) = True Then num = num + 1: f(53) = False
    If Command54.Caption = "!" And f(54) = True Then num = num + 1: f(54) = False
    If Command55.Caption = "!" And f(55) = True Then num = num + 1: f(55) = False
    If Command56.Caption = "!" And f(56) = True Then num = num + 1: f(56) = False
    If Command57.Caption = "!" And f(57) = True Then num = num + 1: f(57) = False
    If Command58.Caption = "!" And f(58) = True Then num = num + 1: f(58) = False
    If Command59.Caption = "!" And f(59) = True Then num = num + 1: f(59) = False
    If Command60.Caption = "!" And f(60) = True Then num = num + 1: f(60) = False
    If Command61.Caption = "!" And f(61) = True Then num = num + 1: f(61) = False
    If Command62.Caption = "!" And f(62) = True Then num = num + 1: f(62) = False
    If Command63.Caption = "!" And f(63) = True Then num = num + 1: f(63) = False
    If Command64.Caption = "!" And f(64) = True Then num = num + 1: f(64) = False
    If Command65.Caption = "!" And f(65) = True Then num = num + 1: f(65) = False
    If Command66.Caption = "!" And f(66) = True Then num = num + 1: f(66) = False
    If Command67.Caption = "!" And f(67) = True Then num = num + 1: f(67) = False
    If Command68.Caption = "!" And f(68) = True Then num = num + 1: f(68) = False
    If Command69.Caption = "!" And f(69) = True Then num = num + 1: f(69) = False
    If Command70.Caption = "!" And f(70) = True Then num = num + 1: f(70) = False
    If Command71.Caption = "!" And f(71) = True Then num = num + 1: f(71) = False
    If Command72.Caption = "!" And f(72) = True Then num = num + 1: f(72) = False
    If Command73.Caption = "!" And f(73) = True Then num = num + 1: f(73) = False
    If Command74.Caption = "!" And f(74) = True Then num = num + 1: f(74) = False
    If Command75.Caption = "!" And f(75) = True Then num = num + 1: f(75) = False
    If Command76.Caption = "!" And f(76) = True Then num = num + 1: f(76) = False
    If Command77.Caption = "!" And f(77) = True Then num = num + 1: f(77) = False
    If Command78.Caption = "!" And f(78) = True Then num = num + 1: f(78) = False
    If Command79.Caption = "!" And f(79) = True Then num = num + 1: f(79) = False
    If Command80.Caption = "!" And f(80) = True Then num = num + 1: f(80) = False
    If Command81.Caption = "!" And f(81) = True Then num = num + 1: f(81) = False
End Sub
Private Sub Timer1_Timer()
    t = t + 1
    Label1.Caption = t
End Sub

Private Sub Timer2_Timer()
    If Timer1.Enabled = True Then Timer1.Enabled = False
End Sub

Private Sub Timer3_Timer()
    Call bj
    Label3.Caption = "»¹ÓÐ" + CStr(10 - num) + "¸öÀ×"
End Sub

Private Sub Timer4_Timer()
    If jg = 71 Then Call jieshu: MsgBox "ÄãÓ®ÁË£¡": Timer4.Enabled = False
End Sub
Function jg() As Integer
    If Command1.Enabled = False And a(1) <> "*" Then jg = jg + 1
    If Command2.Enabled = False And a(2) <> "*" Then jg = jg + 1
    If Command3.Enabled = False And a(3) <> "*" Then jg = jg + 1
    If Command4.Enabled = False And a(4) <> "*" Then jg = jg + 1
    If Command5.Enabled = False And a(5) <> "*" Then jg = jg + 1
    If Command6.Enabled = False And a(6) <> "*" Then jg = jg + 1
    If Command7.Enabled = False And a(7) <> "*" Then jg = jg + 1
    If Command8.Enabled = False And a(8) <> "*" Then jg = jg + 1
    If Command9.Enabled = False And a(9) <> "*" Then jg = jg + 1
    If Command10.Enabled = False And a(10) <> "*" Then jg = jg + 1
    If Command11.Enabled = False And a(11) <> "*" Then jg = jg + 1
    If Command12.Enabled = False And a(12) <> "*" Then jg = jg + 1
    If Command13.Enabled = False And a(13) <> "*" Then jg = jg + 1
    If Command14.Enabled = False And a(14) <> "*" Then jg = jg + 1
    If Command15.Enabled = False And a(15) <> "*" Then jg = jg + 1
    If Command16.Enabled = False And a(16) <> "*" Then jg = jg + 1
    If Command17.Enabled = False And a(17) <> "*" Then jg = jg + 1
    If Command18.Enabled = False And a(18) <> "*" Then jg = jg + 1
    If Command19.Enabled = False And a(19) <> "*" Then jg = jg + 1
    If Command20.Enabled = False And a(20) <> "*" Then jg = jg + 1
    If Command21.Enabled = False And a(21) <> "*" Then jg = jg + 1
    If Command22.Enabled = False And a(22) <> "*" Then jg = jg + 1
    If Command23.Enabled = False And a(23) <> "*" Then jg = jg + 1
    If Command24.Enabled = False And a(24) <> "*" Then jg = jg + 1
    If Command25.Enabled = False And a(25) <> "*" Then jg = jg + 1
    If Command26.Enabled = False And a(26) <> "*" Then jg = jg + 1
    If Command27.Enabled = False And a(27) <> "*" Then jg = jg + 1
    If Command28.Enabled = False And a(28) <> "*" Then jg = jg + 1
    If Command29.Enabled = False And a(29) <> "*" Then jg = jg + 1
    If Command30.Enabled = False And a(30) <> "*" Then jg = jg + 1
    If Command31.Enabled = False And a(31) <> "*" Then jg = jg + 1
    If Command32.Enabled = False And a(32) <> "*" Then jg = jg + 1
    If Command33.Enabled = False And a(33) <> "*" Then jg = jg + 1
    If Command34.Enabled = False And a(34) <> "*" Then jg = jg + 1
    If Command35.Enabled = False And a(35) <> "*" Then jg = jg + 1
    If Command36.Enabled = False And a(36) <> "*" Then jg = jg + 1
    If Command37.Enabled = False And a(37) <> "*" Then jg = jg + 1
    If Command38.Enabled = False And a(38) <> "*" Then jg = jg + 1
    If Command39.Enabled = False And a(39) <> "*" Then jg = jg + 1
    If Command40.Enabled = False And a(40) <> "*" Then jg = jg + 1
    If Command41.Enabled = False And a(41) <> "*" Then jg = jg + 1
    If Command42.Enabled = False And a(42) <> "*" Then jg = jg + 1
    If Command43.Enabled = False And a(43) <> "*" Then jg = jg + 1
    If Command44.Enabled = False And a(44) <> "*" Then jg = jg + 1
    If Command45.Enabled = False And a(45) <> "*" Then jg = jg + 1
    If Command46.Enabled = False And a(46) <> "*" Then jg = jg + 1
    If Command47.Enabled = False And a(47) <> "*" Then jg = jg + 1
    If Command48.Enabled = False And a(48) <> "*" Then jg = jg + 1
    If Command49.Enabled = False And a(49) <> "*" Then jg = jg + 1
    If Command50.Enabled = False And a(50) <> "*" Then jg = jg + 1
    If Command51.Enabled = False And a(51) <> "*" Then jg = jg + 1
    If Command52.Enabled = False And a(52) <> "*" Then jg = jg + 1
    If Command53.Enabled = False And a(53) <> "*" Then jg = jg + 1
    If Command54.Enabled = False And a(54) <> "*" Then jg = jg + 1
    If Command55.Enabled = False And a(55) <> "*" Then jg = jg + 1
    If Command56.Enabled = False And a(56) <> "*" Then jg = jg + 1
    If Command57.Enabled = False And a(57) <> "*" Then jg = jg + 1
    If Command58.Enabled = False And a(58) <> "*" Then jg = jg + 1
    If Command59.Enabled = False And a(59) <> "*" Then jg = jg + 1
    If Command60.Enabled = False And a(60) <> "*" Then jg = jg + 1
    If Command61.Enabled = False And a(61) <> "*" Then jg = jg + 1
    If Command62.Enabled = False And a(62) <> "*" Then jg = jg + 1
    If Command63.Enabled = False And a(63) <> "*" Then jg = jg + 1
    If Command64.Enabled = False And a(64) <> "*" Then jg = jg + 1
    If Command65.Enabled = False And a(65) <> "*" Then jg = jg + 1
    If Command66.Enabled = False And a(66) <> "*" Then jg = jg + 1
    If Command67.Enabled = False And a(67) <> "*" Then jg = jg + 1
    If Command68.Enabled = False And a(68) <> "*" Then jg = jg + 1
    If Command69.Enabled = False And a(69) <> "*" Then jg = jg + 1
    If Command70.Enabled = False And a(70) <> "*" Then jg = jg + 1
    If Command71.Enabled = False And a(71) <> "*" Then jg = jg + 1
    If Command72.Enabled = False And a(72) <> "*" Then jg = jg + 1
    If Command73.Enabled = False And a(73) <> "*" Then jg = jg + 1
    If Command74.Enabled = False And a(74) <> "*" Then jg = jg + 1
    If Command75.Enabled = False And a(75) <> "*" Then jg = jg + 1
    If Command76.Enabled = False And a(76) <> "*" Then jg = jg + 1
    If Command77.Enabled = False And a(77) <> "*" Then jg = jg + 1
    If Command78.Enabled = False And a(78) <> "*" Then jg = jg + 1
    If Command79.Enabled = False And a(79) <> "*" Then jg = jg + 1
    If Command80.Enabled = False And a(80) <> "*" Then jg = jg + 1
    If Command81.Enabled = False And a(81) <> "*" Then jg = jg + 1
End Function
