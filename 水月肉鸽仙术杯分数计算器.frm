VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ë®ÔÂÈâ¸ëÏÉÊõ±­·ÖÊý¼ÆËãÆ÷"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   14205
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CheckBox Check14 
      Caption         =   "Check14"
      Height          =   210
      Left            =   9555
      TabIndex        =   195
      Top             =   5655
      Width           =   210
   End
   Begin VB.CheckBox Check13 
      Caption         =   "Check13"
      Height          =   210
      Left            =   8775
      TabIndex        =   194
      Top             =   5655
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1320
      Left            =   11700
      TabIndex        =   190
      Top             =   975
      Width           =   2160
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3000
      Left            =   10530
      TabIndex        =   189
      Top             =   3705
      Width           =   2355
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   900
      Left            =   11700
      TabIndex        =   188
      Top             =   2535
      Width           =   1770
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1110
      Left            =   9945
      TabIndex        =   187
      Top             =   2340
      Width           =   1575
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1110
      Left            =   9945
      TabIndex        =   186
      Top             =   975
      Width           =   1575
   End
   Begin VB.CheckBox Check128 
      Caption         =   "Check128"
      Height          =   210
      Left            =   6045
      TabIndex        =   185
      Top             =   5265
      Width           =   210
   End
   Begin VB.CheckBox Check127 
      Caption         =   "Check127"
      Height          =   210
      Left            =   5655
      TabIndex        =   184
      Top             =   5265
      Width           =   210
   End
   Begin VB.CheckBox Check126 
      Caption         =   "Check126"
      Height          =   210
      Left            =   5265
      TabIndex        =   183
      Top             =   5265
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check125 
      Caption         =   "Check125"
      Height          =   210
      Left            =   6045
      TabIndex        =   182
      Top             =   4875
      Width           =   210
   End
   Begin VB.CheckBox Check124 
      Caption         =   "Check124"
      Height          =   210
      Left            =   5655
      TabIndex        =   181
      Top             =   4875
      Width           =   210
   End
   Begin VB.CheckBox Check123 
      Caption         =   "Check123"
      Height          =   210
      Left            =   5265
      TabIndex        =   180
      Top             =   4875
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check122 
      Caption         =   "Check122"
      Height          =   210
      Left            =   6045
      TabIndex        =   179
      Top             =   4485
      Width           =   210
   End
   Begin VB.CheckBox Check121 
      Caption         =   "Check121"
      Height          =   210
      Left            =   5655
      TabIndex        =   178
      Top             =   4485
      Width           =   210
   End
   Begin VB.CheckBox Check120 
      Caption         =   "Check120"
      Height          =   210
      Left            =   5265
      TabIndex        =   177
      Top             =   4485
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check119 
      Caption         =   "Check119"
      Height          =   210
      Left            =   6045
      TabIndex        =   176
      Top             =   4095
      Width           =   210
   End
   Begin VB.CheckBox Check118 
      Caption         =   "Check118"
      Height          =   210
      Left            =   5655
      TabIndex        =   175
      Top             =   4095
      Width           =   210
   End
   Begin VB.CheckBox Check117 
      Caption         =   "Check117"
      Height          =   210
      Left            =   5265
      TabIndex        =   174
      Top             =   4095
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check116 
      Caption         =   "Check116"
      Height          =   210
      Left            =   6045
      TabIndex        =   173
      Top             =   3705
      Width           =   210
   End
   Begin VB.CheckBox Check115 
      Caption         =   "Check115"
      Height          =   210
      Left            =   5655
      TabIndex        =   172
      Top             =   3705
      Width           =   210
   End
   Begin VB.CheckBox Check114 
      Caption         =   "Check114"
      Height          =   210
      Left            =   5265
      TabIndex        =   171
      Top             =   3705
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check113 
      Caption         =   "Check113"
      Height          =   210
      Left            =   6045
      TabIndex        =   170
      Top             =   3315
      Width           =   210
   End
   Begin VB.CheckBox Check112 
      Caption         =   "Check112"
      Height          =   210
      Left            =   5655
      TabIndex        =   169
      Top             =   3315
      Width           =   210
   End
   Begin VB.CheckBox Check111 
      Caption         =   "Check111"
      Height          =   210
      Left            =   5265
      TabIndex        =   168
      Top             =   3315
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check110 
      Caption         =   "Check110"
      Height          =   210
      Left            =   6045
      TabIndex        =   167
      Top             =   2925
      Width           =   210
   End
   Begin VB.CheckBox Check109 
      Caption         =   "Check109"
      Height          =   210
      Left            =   5655
      TabIndex        =   166
      Top             =   2925
      Width           =   210
   End
   Begin VB.CheckBox Check108 
      Caption         =   "Check108"
      Height          =   210
      Left            =   5265
      TabIndex        =   165
      Top             =   2925
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check107 
      Caption         =   "Check107"
      Height          =   210
      Left            =   6045
      TabIndex        =   164
      Top             =   2535
      Width           =   210
   End
   Begin VB.CheckBox Check106 
      Caption         =   "Check106"
      Height          =   210
      Left            =   5655
      TabIndex        =   163
      Top             =   2535
      Width           =   210
   End
   Begin VB.CheckBox Check105 
      Caption         =   "Check105"
      Height          =   210
      Left            =   5265
      TabIndex        =   162
      Top             =   2535
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check104 
      Caption         =   "Check104"
      Height          =   210
      Left            =   6045
      TabIndex        =   161
      Top             =   2145
      Width           =   210
   End
   Begin VB.CheckBox Check103 
      Caption         =   "Check103"
      Height          =   180
      Left            =   5655
      TabIndex        =   160
      Top             =   2145
      Width           =   210
   End
   Begin VB.CheckBox Check102 
      Caption         =   "Check102"
      Height          =   210
      Left            =   5265
      TabIndex        =   159
      Top             =   2145
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check101 
      Caption         =   "Check101"
      Height          =   210
      Left            =   6045
      TabIndex        =   158
      Top             =   1755
      Width           =   210
   End
   Begin VB.CheckBox Check100 
      Caption         =   "Check100"
      Height          =   210
      Left            =   5655
      TabIndex        =   157
      Top             =   1755
      Width           =   210
   End
   Begin VB.CheckBox Check99 
      Caption         =   "Check99"
      Height          =   210
      Left            =   5265
      TabIndex        =   156
      Top             =   1755
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check98 
      Caption         =   "Check98"
      Height          =   210
      Left            =   6045
      TabIndex        =   155
      Top             =   1365
      Width           =   210
   End
   Begin VB.CheckBox Check97 
      Caption         =   "Check97"
      Height          =   210
      Left            =   5655
      TabIndex        =   154
      Top             =   1365
      Width           =   210
   End
   Begin VB.CheckBox Check96 
      Caption         =   "Check96"
      Height          =   210
      Left            =   5265
      TabIndex        =   153
      Top             =   1365
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check95 
      Caption         =   "Check95"
      Height          =   210
      Left            =   6045
      TabIndex        =   152
      Top             =   975
      Width           =   210
   End
   Begin VB.CheckBox Check94 
      Caption         =   "Check94"
      Height          =   210
      Left            =   5655
      TabIndex        =   151
      Top             =   975
      Width           =   210
   End
   Begin VB.CheckBox Check93 
      Caption         =   "Check93"
      Height          =   210
      Left            =   5265
      TabIndex        =   150
      Top             =   975
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check92 
      Caption         =   "Check92"
      Height          =   210
      Left            =   6045
      TabIndex        =   149
      Top             =   585
      Width           =   210
   End
   Begin VB.CheckBox Check91 
      Caption         =   "Check91"
      Height          =   210
      Left            =   5655
      TabIndex        =   148
      Top             =   585
      Width           =   210
   End
   Begin VB.CheckBox Check90 
      Caption         =   "Check90"
      Height          =   210
      Left            =   5265
      TabIndex        =   147
      Top             =   585
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check89 
      Caption         =   "Check89"
      Height          =   210
      Left            =   9360
      TabIndex        =   146
      Top             =   3510
      Width           =   210
   End
   Begin VB.CheckBox Check88 
      Caption         =   "Check88"
      Height          =   210
      Left            =   9165
      TabIndex        =   145
      Top             =   3510
      Width           =   210
   End
   Begin VB.CheckBox Check87 
      Caption         =   "Check87"
      Height          =   210
      Left            =   8970
      TabIndex        =   144
      Top             =   3510
      Width           =   210
   End
   Begin VB.CheckBox Check86 
      Caption         =   "Check86"
      Height          =   210
      Left            =   8775
      TabIndex        =   143
      Top             =   3510
      Width           =   210
   End
   Begin VB.CheckBox Check85 
      Caption         =   "Check85"
      Height          =   210
      Left            =   8580
      TabIndex        =   142
      Top             =   3510
      Width           =   210
   End
   Begin VB.CheckBox Check84 
      Caption         =   "Check84"
      Height          =   210
      Left            =   8385
      TabIndex        =   141
      Top             =   3510
      Width           =   210
   End
   Begin VB.CheckBox Check83 
      Caption         =   "Check83"
      Height          =   210
      Left            =   8190
      TabIndex        =   140
      Top             =   3510
      Width           =   210
   End
   Begin VB.CheckBox Check82 
      Caption         =   "Check82"
      Height          =   210
      Left            =   7995
      TabIndex        =   139
      Top             =   3510
      Width           =   210
   End
   Begin VB.CheckBox Check81 
      Caption         =   "Check81"
      Height          =   210
      Left            =   7800
      TabIndex        =   138
      Top             =   3510
      Width           =   210
   End
   Begin VB.CheckBox Check80 
      Caption         =   "Check80"
      Height          =   210
      Left            =   7605
      TabIndex        =   137
      Top             =   3510
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check79 
      Caption         =   "Check79"
      Height          =   210
      Left            =   3120
      TabIndex        =   136
      Top             =   6630
      Width           =   210
   End
   Begin VB.CheckBox Check78 
      Caption         =   "Check78"
      Height          =   210
      Left            =   2925
      TabIndex        =   135
      Top             =   6630
      Width           =   210
   End
   Begin VB.CheckBox Check77 
      Caption         =   "Check77"
      Height          =   180
      Left            =   2730
      TabIndex        =   134
      Top             =   6630
      Width           =   210
   End
   Begin VB.CheckBox Check76 
      Caption         =   "Check76"
      Height          =   210
      Left            =   2535
      TabIndex        =   133
      Top             =   6630
      Width           =   210
   End
   Begin VB.CheckBox Check75 
      Caption         =   "Check75"
      Height          =   210
      Left            =   2340
      TabIndex        =   132
      Top             =   6630
      Width           =   210
   End
   Begin VB.CheckBox Check74 
      Caption         =   "Check74"
      Height          =   210
      Left            =   2145
      TabIndex        =   131
      Top             =   6630
      Width           =   210
   End
   Begin VB.CheckBox Check73 
      Caption         =   "Check73"
      Height          =   210
      Left            =   1950
      TabIndex        =   130
      Top             =   6630
      Width           =   210
   End
   Begin VB.CheckBox Check72 
      Caption         =   "Check72"
      Height          =   210
      Left            =   1755
      TabIndex        =   129
      Top             =   6630
      Width           =   210
   End
   Begin VB.CheckBox Check71 
      Caption         =   "Check71"
      Height          =   210
      Left            =   1560
      TabIndex        =   128
      Top             =   6630
      Width           =   210
   End
   Begin VB.CheckBox Check70 
      Caption         =   "Check70"
      Height          =   210
      Left            =   1365
      TabIndex        =   127
      Top             =   6630
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check69 
      Caption         =   "Check69"
      Height          =   210
      Left            =   3120
      TabIndex        =   126
      Top             =   6045
      Width           =   210
   End
   Begin VB.CheckBox Check68 
      Caption         =   "Check68"
      Height          =   210
      Left            =   2925
      TabIndex        =   125
      Top             =   6045
      Width           =   210
   End
   Begin VB.CheckBox Check67 
      Caption         =   "Check67"
      Height          =   210
      Left            =   2730
      TabIndex        =   124
      Top             =   6045
      Width           =   210
   End
   Begin VB.CheckBox Check66 
      Caption         =   "Check66"
      Height          =   210
      Left            =   2535
      TabIndex        =   123
      Top             =   6045
      Width           =   210
   End
   Begin VB.CheckBox Check65 
      Caption         =   "Check65"
      Height          =   210
      Left            =   2340
      TabIndex        =   122
      Top             =   6045
      Width           =   210
   End
   Begin VB.CheckBox Check64 
      Caption         =   "Check64"
      Height          =   210
      Left            =   2145
      TabIndex        =   121
      Top             =   6045
      Width           =   210
   End
   Begin VB.CheckBox Check63 
      Caption         =   "Check63"
      Height          =   210
      Left            =   1950
      TabIndex        =   120
      Top             =   6045
      Width           =   210
   End
   Begin VB.CheckBox Check62 
      Caption         =   "Check62"
      Height          =   210
      Left            =   1755
      TabIndex        =   119
      Top             =   6045
      Width           =   210
   End
   Begin VB.CheckBox Check61 
      Caption         =   "Check61"
      Height          =   210
      Left            =   1560
      TabIndex        =   118
      Top             =   6045
      Width           =   210
   End
   Begin VB.CheckBox Check60 
      Caption         =   "Check60"
      Height          =   210
      Left            =   1365
      TabIndex        =   117
      Top             =   6045
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check59 
      Caption         =   "Check59"
      Height          =   210
      Left            =   3120
      TabIndex        =   116
      Top             =   5460
      Width           =   210
   End
   Begin VB.CheckBox Check58 
      Caption         =   "Check58"
      Height          =   210
      Left            =   2925
      TabIndex        =   115
      Top             =   5460
      Width           =   210
   End
   Begin VB.CheckBox Check57 
      Caption         =   "Check57"
      Height          =   210
      Left            =   2730
      TabIndex        =   114
      Top             =   5460
      Width           =   210
   End
   Begin VB.CheckBox Check56 
      Caption         =   "Check56"
      Height          =   210
      Left            =   2535
      TabIndex        =   113
      Top             =   5460
      Width           =   210
   End
   Begin VB.CheckBox Check55 
      Caption         =   "Check55"
      Height          =   210
      Left            =   2340
      TabIndex        =   112
      Top             =   5460
      Width           =   210
   End
   Begin VB.CheckBox Check54 
      Caption         =   "Check54"
      Height          =   210
      Left            =   2145
      TabIndex        =   111
      Top             =   5460
      Width           =   210
   End
   Begin VB.CheckBox Check53 
      Caption         =   "Check53"
      Height          =   210
      Left            =   1950
      TabIndex        =   110
      Top             =   5460
      Width           =   210
   End
   Begin VB.CheckBox Check52 
      Caption         =   "Check52"
      Height          =   210
      Left            =   1755
      TabIndex        =   109
      Top             =   5460
      Width           =   210
   End
   Begin VB.CheckBox Check51 
      Caption         =   "Check51"
      Height          =   210
      Left            =   1560
      TabIndex        =   108
      Top             =   5460
      Width           =   210
   End
   Begin VB.CheckBox Check50 
      Caption         =   "Check50"
      Height          =   210
      Left            =   1365
      TabIndex        =   107
      Top             =   5460
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check49 
      Caption         =   "Check49"
      Height          =   210
      Left            =   3120
      TabIndex        =   106
      Top             =   1755
      Width           =   210
   End
   Begin VB.CheckBox Check48 
      Caption         =   "Check48"
      Height          =   210
      Left            =   2925
      TabIndex        =   105
      Top             =   1755
      Width           =   210
   End
   Begin VB.CheckBox Check47 
      Caption         =   "Check47"
      Height          =   210
      Left            =   2730
      TabIndex        =   104
      Top             =   1755
      Width           =   210
   End
   Begin VB.CheckBox Check46 
      Caption         =   "Check46"
      Height          =   210
      Left            =   2535
      TabIndex        =   103
      Top             =   1755
      Width           =   210
   End
   Begin VB.CheckBox Check45 
      Caption         =   "Check45"
      Height          =   210
      Left            =   2340
      TabIndex        =   102
      Top             =   1755
      Width           =   210
   End
   Begin VB.CheckBox Check44 
      Caption         =   "Check44"
      Height          =   210
      Left            =   2145
      TabIndex        =   101
      Top             =   1755
      Width           =   210
   End
   Begin VB.CheckBox Check43 
      Caption         =   "Check43"
      Height          =   210
      Left            =   1950
      TabIndex        =   100
      Top             =   1755
      Width           =   210
   End
   Begin VB.CheckBox Check42 
      Caption         =   "Check42"
      Height          =   210
      Left            =   1755
      TabIndex        =   99
      Top             =   1755
      Width           =   210
   End
   Begin VB.CheckBox Check41 
      Caption         =   "Check41"
      Height          =   210
      Left            =   1560
      TabIndex        =   98
      Top             =   1755
      Width           =   210
   End
   Begin VB.CheckBox Check40 
      Caption         =   "Check40"
      Height          =   210
      Left            =   1365
      TabIndex        =   97
      Top             =   1755
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check39 
      Caption         =   "Check39"
      Height          =   210
      Left            =   3120
      TabIndex        =   96
      Top             =   1170
      Width           =   210
   End
   Begin VB.CheckBox Check38 
      Caption         =   "Check38"
      Height          =   210
      Left            =   2925
      TabIndex        =   95
      Top             =   1170
      Width           =   210
   End
   Begin VB.CheckBox Check37 
      Caption         =   "Check37"
      Height          =   210
      Left            =   2730
      TabIndex        =   94
      Top             =   1170
      Width           =   210
   End
   Begin VB.CheckBox Check36 
      Caption         =   "Check36"
      Height          =   210
      Left            =   2535
      TabIndex        =   93
      Top             =   1170
      Width           =   210
   End
   Begin VB.CheckBox Check35 
      Caption         =   "Check35"
      Height          =   210
      Left            =   2340
      TabIndex        =   92
      Top             =   1170
      Width           =   210
   End
   Begin VB.CheckBox Check34 
      Caption         =   "Check34"
      Height          =   210
      Left            =   2145
      TabIndex        =   91
      Top             =   1170
      Width           =   210
   End
   Begin VB.CheckBox Check33 
      Caption         =   "Check33"
      Height          =   210
      Left            =   1950
      TabIndex        =   90
      Top             =   1170
      Width           =   210
   End
   Begin VB.CheckBox Check32 
      Caption         =   "Check32"
      Height          =   210
      Left            =   1755
      TabIndex        =   89
      Top             =   1170
      Width           =   210
   End
   Begin VB.CheckBox Check31 
      Caption         =   "Check31"
      Height          =   210
      Left            =   1560
      TabIndex        =   88
      Top             =   1170
      Width           =   210
   End
   Begin VB.CheckBox Check30 
      Caption         =   "Check30"
      Height          =   210
      Left            =   1365
      TabIndex        =   87
      Top             =   1170
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check29 
      Caption         =   "Check29"
      Height          =   210
      Left            =   3120
      TabIndex        =   86
      Top             =   585
      Width           =   210
   End
   Begin VB.CheckBox Check28 
      Caption         =   "Check28"
      Height          =   210
      Left            =   2925
      TabIndex        =   85
      Top             =   585
      Width           =   210
   End
   Begin VB.CheckBox Check27 
      Caption         =   "Check27"
      Height          =   210
      Left            =   2730
      TabIndex        =   84
      Top             =   585
      Width           =   210
   End
   Begin VB.CheckBox Check26 
      Caption         =   "Check26"
      Height          =   210
      Left            =   2535
      TabIndex        =   83
      Top             =   585
      Width           =   210
   End
   Begin VB.CheckBox Check25 
      Caption         =   "Check25"
      Height          =   210
      Left            =   2340
      TabIndex        =   82
      Top             =   585
      Width           =   210
   End
   Begin VB.CheckBox Check24 
      Caption         =   "Check24"
      Height          =   210
      Left            =   2145
      TabIndex        =   81
      Top             =   585
      Width           =   210
   End
   Begin VB.CheckBox Check23 
      Caption         =   "Check23"
      Height          =   210
      Left            =   1950
      TabIndex        =   80
      Top             =   585
      Width           =   210
   End
   Begin VB.CheckBox Check22 
      Caption         =   "Check22"
      Height          =   210
      Left            =   1755
      TabIndex        =   79
      Top             =   585
      Width           =   210
   End
   Begin VB.CheckBox Check21 
      Caption         =   "Check21"
      Height          =   210
      Left            =   1560
      TabIndex        =   78
      Top             =   585
      Width           =   210
   End
   Begin VB.CheckBox Check20 
      Caption         =   "Check20"
      Height          =   210
      Left            =   1365
      TabIndex        =   77
      Top             =   585
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   6630
      TabIndex        =   24
      Top             =   5850
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   6630
      TabIndex        =   23
      Top             =   5070
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   6630
      TabIndex        =   22
      Top             =   4290
      Width           =   3135
   End
   Begin VB.CheckBox Check12 
      Caption         =   "Check12"
      Height          =   210
      Left            =   6825
      TabIndex        =   21
      Top             =   2925
      Width           =   210
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Check11"
      Height          =   210
      Left            =   6825
      TabIndex        =   20
      Top             =   2535
      Width           =   210
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Check10"
      Height          =   210
      Left            =   6825
      TabIndex        =   19
      Top             =   2145
      Width           =   210
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Check9"
      Height          =   210
      Left            =   6825
      TabIndex        =   18
      Top             =   1755
      Width           =   210
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Check8"
      Height          =   210
      Left            =   6825
      TabIndex        =   17
      Top             =   1365
      Width           =   210
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Check7"
      Height          =   210
      Left            =   6825
      TabIndex        =   16
      Top             =   975
      Width           =   210
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Check6"
      Height          =   210
      Left            =   6825
      TabIndex        =   15
      Top             =   585
      Width           =   210
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Check5"
      Height          =   210
      Left            =   195
      TabIndex        =   14
      Top             =   4485
      Width           =   210
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check4"
      Height          =   210
      Left            =   195
      TabIndex        =   13
      Top             =   4095
      Width           =   210
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   210
      Left            =   195
      TabIndex        =   12
      Top             =   3705
      Width           =   210
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   210
      Left            =   195
      TabIndex        =   11
      Top             =   3315
      Width           =   210
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   210
      Left            =   195
      TabIndex        =   10
      Top             =   2925
      Width           =   210
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¹éÁã"
      BeginProperty Font 
         Name            =   "ºÚÌå"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6630
      TabIndex        =   9
      Top             =   6435
      Width           =   3135
   End
   Begin VB.Label Label13 
      Caption         =   "¼õ·Ö"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   9165
      TabIndex        =   193
      Top             =   5655
      Width           =   600
   End
   Begin VB.Label Label12 
      Caption         =   "¼Ó·Ö"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8385
      TabIndex        =   192
      Top             =   5655
      Width           =   600
   End
   Begin VB.Label Label11 
      BeginProperty Font 
         Name            =   "ÐÂËÎÌå"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   795
      Left            =   10920
      TabIndex        =   191
      Top             =   195
      Width           =   2355
   End
   Begin VB.Label Label6 
      Caption         =   "ÆôÊ¾£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6630
      TabIndex        =   76
      Top             =   3315
      Width           =   3135
   End
   Begin VB.Label Label47 
      Caption         =   "0 1 2 3 4 5 6 7 8 9"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   1365
      TabIndex        =   59
      Top             =   5655
      Width           =   2160
   End
   Begin VB.Label Label47 
      Caption         =   "0 1 2 3 4 5 6 7 8 9"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   1365
      TabIndex        =   75
      Top             =   6825
      Width           =   2160
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   12
      Left            =   5265
      TabIndex        =   74
      Top             =   5460
      Width           =   990
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   11
      Left            =   5265
      TabIndex        =   73
      Top             =   5070
      Width           =   990
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   10
      Left            =   5265
      TabIndex        =   72
      Top             =   4680
      Width           =   990
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   5265
      TabIndex        =   71
      Top             =   4290
      Width           =   990
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   5265
      TabIndex        =   70
      Top             =   3900
      Width           =   990
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   5265
      TabIndex        =   69
      Top             =   3510
      Width           =   990
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   5265
      TabIndex        =   68
      Top             =   3120
      Width           =   990
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   5265
      TabIndex        =   67
      Top             =   2730
      Width           =   990
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   5265
      TabIndex        =   66
      Top             =   2340
      Width           =   990
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   5265
      TabIndex        =   65
      Top             =   1950
      Width           =   990
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   5265
      TabIndex        =   64
      Top             =   1560
      Width           =   990
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   5265
      TabIndex        =   63
      Top             =   1170
      Width           =   990
   End
   Begin VB.Label Label48 
      Caption         =   "0   1   2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   5265
      TabIndex        =   62
      Top             =   780
      Width           =   990
   End
   Begin VB.Label Label47 
      Caption         =   "0 1 2 3 4 5 6 7 8 9"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   7605
      TabIndex        =   61
      Top             =   3705
      Width           =   2160
   End
   Begin VB.Label Label47 
      Caption         =   "0 1 2 3 4 5 6 7 8 9"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   1365
      TabIndex        =   60
      Top             =   6240
      Width           =   2160
   End
   Begin VB.Label Label47 
      Caption         =   "0 1 2 3 4 5 6 7 8 9"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   1365
      TabIndex        =   58
      Top             =   1950
      Width           =   2160
   End
   Begin VB.Label Label47 
      Caption         =   "0 1 2 3 4 5 6 7 8 9"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1365
      TabIndex        =   57
      Top             =   1365
      Width           =   2160
   End
   Begin VB.Label Label47 
      Caption         =   "0 1 2 3 4 5 6 7 8 9"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   1365
      TabIndex        =   56
      Top             =   780
      Width           =   2160
   End
   Begin VB.Label Label46 
      Caption         =   "¿ñÐÅÈç»ðÎÞÂ©£¨50£©"
      Height          =   210
      Left            =   7215
      TabIndex        =   55
      Top             =   2925
      Width           =   2550
   End
   Begin VB.Label Label45 
      Caption         =   "Ñ¼±¾ÔË×÷Í¨¹Ø£¨80£©"
      Height          =   210
      Left            =   7215
      TabIndex        =   54
      Top             =   2535
      Width           =   2550
   End
   Begin VB.Label Label44 
      Caption         =   "ÌøÎèÎÞÂ©È«ÊÕ£¨100£©"
      Height          =   210
      Left            =   7215
      TabIndex        =   53
      Top             =   2145
      Width           =   2355
   End
   Begin VB.Label Label43 
      Caption         =   "ÌøÎèÍ¨¹Ø£¨50£©"
      Height          =   210
      Left            =   7215
      TabIndex        =   52
      Top             =   1755
      Width           =   2550
   End
   Begin VB.Label Label42 
      Caption         =   "ÕæÏàÎÞÂ©£¨100£©"
      Height          =   210
      Left            =   7215
      TabIndex        =   51
      Top             =   1365
      Width           =   2355
   End
   Begin VB.Label Label41 
      Caption         =   "ÕæÏàÍ¨¹Ø£¨50£©"
      Height          =   210
      Left            =   7215
      TabIndex        =   50
      Top             =   975
      Width           =   2355
   End
   Begin VB.Label Label40 
      Caption         =   "¼à¹¤ÏÖ³¡»÷É±Ñ¼×Ó£¨30£©"
      Height          =   210
      Left            =   7215
      TabIndex        =   49
      Top             =   585
      Width           =   2160
   End
   Begin VB.Label Label39 
      Caption         =   "Éî¶ÈÈÏÖª£¨80£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   48
      Top             =   5265
      Width           =   1575
   End
   Begin VB.Label Label38 
      Caption         =   "Ë®»ðÏàÈÝ£¨150£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   47
      Top             =   4875
      Width           =   1380
   End
   Begin VB.Label Label37 
      Caption         =   "»úÐµÖ®ÔÖ£¨50£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   46
      Top             =   4485
      Width           =   1380
   End
   Begin VB.Label Label36 
      Caption         =   "Óà½ý·½Õó£¨90£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   45
      Top             =   4095
      Width           =   1380
   End
   Begin VB.Label Label35 
      Caption         =   "ºÃÃÎÔÚºÎ·½£¨90£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   44
      Top             =   3705
      Width           =   1575
   End
   Begin VB.Label Label34 
      Caption         =   "ÓýÉú³Ø£¨90£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   43
      Top             =   3315
      Width           =   1380
   End
   Begin VB.Label Label33 
      Caption         =   "Ê§¿Ø£¨120£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   42
      Top             =   2925
      Width           =   1380
   End
   Begin VB.Label Label32 
      Caption         =   "ÎÞÂ©äéºÛÀÖÔ°£¨50£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   41
      Top             =   2535
      Width           =   1770
   End
   Begin VB.Label Label31 
      Caption         =   "ï¥ÓëÖÈÐò£¨40£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   40
      Top             =   2145
      Width           =   1380
   End
   Begin VB.Label Label30 
      Caption         =   "á÷ÁÔ³¡£¨60£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   39
      Top             =   1755
      Width           =   1380
   End
   Begin VB.Label Label29 
      Caption         =   "ÁìµØÒâÊ¶£¨100£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   38
      Top             =   1365
      Width           =   1380
   End
   Begin VB.Label Label28 
      Caption         =   "Õ°Ç°¹Ëºó£¨40£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   37
      Top             =   975
      Width           =   1575
   End
   Begin VB.Label Label27 
      Caption         =   "³²Ñ¨£¨40£©"
      Height          =   210
      Left            =   3705
      TabIndex        =   36
      Top             =   585
      Width           =   1185
   End
   Begin VB.Label Label26 
      Caption         =   "ÐÜ£¨20£©"
      Height          =   210
      Left            =   195
      TabIndex        =   35
      Top             =   6630
      Width           =   990
   End
   Begin VB.Label Label25 
      Caption         =   "¹·Í·£¨20£©"
      Height          =   210
      Left            =   195
      TabIndex        =   34
      Top             =   6045
      Width           =   990
   End
   Begin VB.Label Label24 
      Caption         =   "Ñ¼×Ó£¨20£©"
      Height          =   210
      Left            =   195
      TabIndex        =   33
      Top             =   5460
      Width           =   990
   End
   Begin VB.Label Label23 
      Caption         =   "Ä¹±®£¨230£©"
      Height          =   210
      Left            =   585
      TabIndex        =   32
      Top             =   4485
      Width           =   1185
   End
   Begin VB.Label Label22 
      Caption         =   "º®ÔÖ£¨150£©"
      Height          =   210
      Left            =   585
      TabIndex        =   31
      Top             =   4095
      Width           =   1185
   End
   Begin VB.Label Label21 
      Caption         =   "Ðâ´¸£¨120£©"
      Height          =   210
      Left            =   585
      TabIndex        =   30
      Top             =   3705
      Width           =   1185
   End
   Begin VB.Label Label20 
      Caption         =   "Ë®ÔÂ£¨300£©"
      Height          =   210
      Left            =   585
      TabIndex        =   29
      Top             =   3315
      Width           =   1185
   End
   Begin VB.Label Label19 
      Caption         =   "ÆïÊ¿£¨450£©"
      Height          =   210
      Left            =   585
      TabIndex        =   28
      Top             =   2925
      Width           =   1185
   End
   Begin VB.Label Label18 
      Caption         =   "ËÄÐÇ£¨10£©"
      Height          =   210
      Left            =   195
      TabIndex        =   27
      Top             =   1755
      Width           =   990
   End
   Begin VB.Label Label17 
      Caption         =   "ÎåÐÇ£¨20£©"
      Height          =   210
      Left            =   195
      TabIndex        =   26
      Top             =   1170
      Width           =   990
   End
   Begin VB.Label Label16 
      Caption         =   "ÁùÐÇ£¨50£©"
      Height          =   210
      Left            =   195
      TabIndex        =   25
      Top             =   585
      Width           =   990
   End
   Begin VB.Label Label10 
      Caption         =   "×Ü·Ö£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10140
      TabIndex        =   8
      Top             =   195
      Width           =   3330
   End
   Begin VB.Label Label9 
      Caption         =   "ÐÞÕý·Ö£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6630
      TabIndex        =   7
      Top             =   5655
      Width           =   3135
   End
   Begin VB.Label Label8 
      Caption         =   "½áËã·Ö£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6630
      TabIndex        =   6
      Top             =   4875
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "²ØÆ·£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6630
      TabIndex        =   5
      Top             =   4095
      Width           =   2160
   End
   Begin VB.Label Label5 
      Caption         =   "ÊÂ¼þ£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6630
      TabIndex        =   4
      Top             =   195
      Width           =   2940
   End
   Begin VB.Label Label4 
      Caption         =   "½ô¼±£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3705
      TabIndex        =   3
      Top             =   195
      Width           =   2355
   End
   Begin VB.Label Label3 
      Caption         =   "Òþ²Ø£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   2
      Top             =   5070
      Width           =   2940
   End
   Begin VB.Label Label2 
      Caption         =   "½á¾Ö£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   1
      Top             =   2535
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "ÁÙÊ±ÕÐÄ¼£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   0
      Top             =   195
      Width           =   2160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim t11, t12, t13 As Integer 'ÁÙÊ±ÕÐÄ¼
Dim t21, t22, t23, t24, t25 As Integer '½á¾Ö
Dim t31, t32, t33 As Integer 'Òþ²Ø
Dim t4(1 To 13) As Integer '½ô¼±
Dim t5(1 To 7) As Integer  'ÊÂ¼þ
Dim t6 As Integer 'ÆôÊ¾
Dim t7, t8, t9 As Double '²ØÆ·£¬½áËã£¬ÐÞÕý
Dim f As Boolean 'ÐÞÕý·Ö¼Ó¼õ

Public Sub total()
Dim tt1, tt2, tt3, tt4, tt5, tt6, tt7, tt9 As Double
tt1 = 450 * t21 + 300 * t22 + 120 * t23 + 150 * t24 + 230 * t25
tt2 = 50 * t11 + 20 * t12 + 10 * t13
tt3 = 20 * (t31 + t32 + t33)
tt4 = t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80
tt5 = 30 * t5(1) + 50 * t5(2) + 100 * t5(3) + 50 * t5(4) + 100 * t5(5) + 80 * t5(6) + 50 * t5(7)
tt6 = 50 * t6
tt7 = 10 * t7
If f Then tt9 = t9 Else tt9 = -t9
Label11.Caption = CStr(tt1 + tt2 + tt3 + tt4 + tt5 + tt6 + tt7 + t8 + tt9) + "!"
End Sub
Public Sub list1add()
List1.Clear
If t21 = 1 Then List1.AddItem "ÆïÊ¿£¨450£©"
If t22 = 1 Then List1.AddItem "Ë®ÔÂ£¨300£©"
If t23 = 1 Then List1.AddItem "Ðâ´¸£¨120£©"
If t24 = 1 Then List1.AddItem "º®ÔÖ£¨150£©"
If t25 = 1 Then List1.AddItem "Ä¹±®£¨230£©"
End Sub
Public Sub list2add()
List2.Clear
If t11 > 0 Then List2.AddItem ("ÁùÐÇ£¨50£©*" + CStr(t11))
If t12 > 0 Then List2.AddItem ("ÎåÐÇ£¨20£©*" + CStr(t12))
If t13 > 0 Then List2.AddItem ("ËÄÐÇ£¨10£©*" + CStr(t13))
End Sub
Public Sub list3add()
List3.Clear
If t31 > 0 Then List3.AddItem ("Ñ¼×Ó£¨20£©*" + CStr(t31))
If t32 > 0 Then List3.AddItem ("¹·Í·£¨20£©*" + CStr(t32))
If t33 > 0 Then List3.AddItem ("ÐÜ£¨20£©*" + CStr(t33))
End Sub
Public Sub list4add()
List4.Clear
If t4(1) > 0 Then List4.AddItem ("³²Ñ¨£¨40£©*" + CStr(t4(1)))
If t4(2) > 0 Then List4.AddItem ("Õ°Ç°¹Ëºó£¨40£©*" + CStr(t4(2)))
If t4(3) > 0 Then List4.AddItem ("ÁìµØÒâÊ¶£¨100£©*" + CStr(t4(3)))
If t4(4) > 0 Then List4.AddItem ("á÷ÁÔ³¡£¨60£©*" + CStr(t4(4)))
If t4(5) > 0 Then List4.AddItem ("ï¥ÓëÖÈÐò£¨40£©*" + CStr(t4(5)))
If t4(6) > 0 Then List4.AddItem ("ÎÞÂ©äéºÛÀÖÔ°£¨50£©*" + CStr(t4(6)))
If t4(7) > 0 Then List4.AddItem ("Ê§¿Ø£¨120£©*" + CStr(t4(7)))
If t4(8) > 0 Then List4.AddItem ("ÓýÉú³Ø£¨90£©*" + CStr(t4(8)))
If t4(9) > 0 Then List4.AddItem ("ºÃÃÎÔÚºÎ·½£¨90£©*" + CStr(t4(9)))
If t4(10) > 0 Then List4.AddItem ("Óà½ý·½Õó£¨90£©*" + CStr(t4(10)))
If t4(11) > 0 Then List4.AddItem ("»úÐµÖ®ÔÖ£¨50£©*" + CStr(t4(11)))
If t4(12) > 0 Then List4.AddItem ("Ë®»ðÏàÈÝ£¨150£©*" + CStr(t4(12)))
If t4(13) > 0 Then List4.AddItem ("Éî¶ÈÈÏÖª£¨80£©*" + CStr(t4(13)))
End Sub
Public Sub list5add()
List5.Clear
If t5(1) > 0 Then List5.AddItem ("¼à¹¤ÏÖ³¡»÷É±Ñ¼×Ó(30)")
If t5(2) > 0 Then List5.AddItem ("ÕæÏàÍ¨¹Ø(50)")
If t5(3) > 0 Then List5.AddItem ("ÕæÏàÎÞÂ©£¨100£©")
If t5(4) > 0 Then List5.AddItem ("ÌøÎèÍ¨¹Ø£¨50£©")
If t5(5) > 0 Then List5.AddItem ("ÌøÎèÎÞÂ©È«ÊÕ£¨100£©")
If t5(6) > 0 Then List5.AddItem ("Ñ¼±¾ÔË×÷Í¨¹Ø£¨80£©")
If t5(7) > 0 Then List5.AddItem ("¿ñÐÅÈç»ðÎÞÂ©£¨50£©")
End Sub
Public Sub begin()
Dim i As Integer
Text1.Text = "": Text2.Text = "": Text3.Text = ""
Check1.Value = 0: Check2.Value = 0: Check3.Value = 0: Check4.Value = 0: Check5.Value = 0: Check6.Value = 0: Check7.Value = 0: Check8.Value = 0: Check9.Value = 0
Check10.Value = 0: Check11.Value = 0: Check12.Value = 0: Check13.Value = 1: Check14.Value = 0
Check20.Value = 1: Check21.Value = 0: Check22.Value = 0: Check23.Value = 0: Check24.Value = 0: Check25.Value = 0: Check26.Value = 0: Check27.Value = 0: Check28.Value = 0: Check29.Value = 0
Check30.Value = 1: Check31.Value = 0: Check32.Value = 0: Check33.Value = 0: Check34.Value = 0: Check35.Value = 0: Check36.Value = 0: Check37.Value = 0: Check38.Value = 0: Check39.Value = 0
Check40.Value = 1: Check41.Value = 0: Check42.Value = 0: Check43.Value = 0: Check44.Value = 0: Check45.Value = 0: Check46.Value = 0: Check47.Value = 0: Check48.Value = 0: Check49.Value = 0
Check50.Value = 1: Check51.Value = 0: Check52.Value = 0: Check53.Value = 0: Check54.Value = 0: Check55.Value = 0: Check56.Value = 0: Check57.Value = 0: Check58.Value = 0: Check59.Value = 0
Check60.Value = 1: Check61.Value = 0: Check62.Value = 0: Check63.Value = 0: Check64.Value = 0: Check65.Value = 0: Check66.Value = 0: Check67.Value = 0: Check68.Value = 0: Check69.Value = 0
Check70.Value = 1: Check71.Value = 0: Check72.Value = 0: Check73.Value = 0: Check74.Value = 0: Check75.Value = 0: Check76.Value = 0: Check77.Value = 0: Check78.Value = 0: Check79.Value = 0
Check80.Value = 1: Check81.Value = 0: Check82.Value = 0: Check83.Value = 0: Check84.Value = 0: Check85.Value = 0: Check86.Value = 0: Check87.Value = 0: Check88.Value = 0: Check89.Value = 0
Check90.Value = 1: Check91.Value = 0: Check92.Value = 0: Check93.Value = 1: Check94.Value = 0: Check95.Value = 0: Check96.Value = 1: Check97.Value = 0: Check98.Value = 0: Check99.Value = 1
Check100.Value = 0: Check101.Value = 0: Check102.Value = 1: Check103.Value = 0: Check104.Value = 0: Check105.Value = 1: Check106.Value = 0: Check107.Value = 0: Check108.Value = 1: Check109.Value = 0
Check110.Value = 0: Check111.Value = 1: Check112.Value = 0: Check113.Value = 0: Check114.Value = 1: Check115.Value = 0: Check116.Value = 0: Check117.Value = 1: Check118.Value = 0: Check119.Value = 0
Check120.Value = 1: Check121.Value = 0: Check122.Value = 0: Check123.Value = 1: Check124.Value = 0: Check125.Value = 0: Check126.Value = 1: Check127.Value = 0: Check128.Value = 0
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º"
Label2.Caption = "½á¾Ö£º"
Label3.Caption = "Òþ²Ø£º"
Label4.Caption = "½ô¼±£º"
Label5.Caption = "ÊÂ¼þ£º"
Label6.Caption = "ÆôÊ¾£º"
Label7.Caption = "²ØÆ·£º"
Label8.Caption = "½áËã·Ö£º"
Label9.Caption = "ÐÞÕý·Ö£º"
Label10.Caption = "×Ü·Ö£º"
Label11.Caption = ""
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
f = True
t11 = 0: t12 = 0: t13 = 0: t21 = 0: t22 = 0: t23 = 0: t24 = 0: t25 = 0: t31 = 0: t32 = 0: t33 = 0: t6 = 0: t7 = 0: t8 = 0: t9 = 0
For i = 1 To 13
    t4(i) = 0
Next i
For i = 1 To 7
    t5(i) = 0
Next i
End Sub



Private Sub Check1_Click()
t21 = 0
If Check1.Value = 1 Then t21 = 1
Call list1add
Label2.Caption = "½á¾Ö£º" + CStr(450 * t21 + 300 * t22 + 120 * t23 + 150 * t24 + 230 * t25)
End Sub

Private Sub Check10_Click()
t5(5) = 0
If Check10.Value = 1 Then Check9.Value = 0: t5(5) = 1
Call list5add
Label5.Caption = "ÊÂ¼þ£º" + CStr(30 * t5(1) + 50 * t5(2) + 100 * t5(3) + 50 * t5(4) + 100 * t5(5) + 80 * t5(6) + 50 * t5(7))

End Sub

Private Sub Check100_Click()
t4(4) = 0
If Check100.Value = 1 Then
Check99.Value = 0: Check101.Value = 0
t4(4) = 1
End If
If Check99.Value = 0 And Check100.Value = 0 And Check101.Value = 0 Then Check99.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check101_Click()
t4(4) = 0
If Check101.Value = 1 Then
Check100.Value = 0: Check99.Value = 0
t4(4) = 3
End If
If Check99.Value = 0 And Check100.Value = 0 And Check101.Value = 0 Then Check99.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check102_Click()
t4(5) = 0
If Check102.Value = 1 Then
Check103.Value = 0: Check104.Value = 0
t4(5) = 0
End If
If Check102.Value = 0 And Check103.Value = 0 And Check104.Value = 0 Then Check102.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check103_Click()
t4(5) = 0
If Check103.Value = 1 Then
Check102.Value = 0: Check104.Value = 0
t4(5) = 1
End If
If Check102.Value = 0 And Check103.Value = 0 And Check104.Value = 0 Then Check102.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check104_Click()
t4(5) = 0
If Check104.Value = 1 Then
Check103.Value = 0: Check102.Value = 0
t4(5) = 2
End If
If Check102.Value = 0 And Check103.Value = 0 And Check104.Value = 0 Then Check102.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check105_Click()
t4(6) = 0
If Check105.Value = 1 Then
Check106.Value = 0: Check107.Value = 0
t4(6) = 0
End If
If Check105.Value = 0 And Check106.Value = 0 And Check107.Value = 0 Then Check105.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check106_Click()
t4(6) = 0
If Check106.Value = 1 Then
Check105.Value = 0: Check107.Value = 0
t4(6) = 1
End If
If Check105.Value = 0 And Check106.Value = 0 And Check107.Value = 0 Then Check105.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check107_Click()
t4(6) = 0
If Check107.Value = 1 Then
Check106.Value = 0: Check105.Value = 0
t4(6) = 2
End If
If Check105.Value = 0 And Check106.Value = 0 And Check107.Value = 0 Then Check105.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check108_Click()
t4(7) = 0
If Check108.Value = 1 Then
Check109.Value = 0: Check110.Value = 0
t4(7) = 0
End If
If Check108.Value = 0 And Check109.Value = 0 And Check110.Value = 0 Then Check108.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check109_Click()
t4(7) = 0
If Check109.Value = 1 Then
Check110.Value = 0: Check108.Value = 0
t4(7) = 1
End If
If Check108.Value = 0 And Check109.Value = 0 And Check110.Value = 0 Then Check108.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check11_Click()
t5(6) = 0
If Check11.Value = 1 Then t5(6) = 1
Call list5add
Label5.Caption = "ÊÂ¼þ£º" + CStr(30 * t5(1) + 50 * t5(2) + 100 * t5(3) + 50 * t5(4) + 100 * t5(5) + 80 * t5(6) + 50 * t5(7))

End Sub

Private Sub Check110_Click()
t4(7) = 0
If Check110.Value = 1 Then
Check109.Value = 0: Check108.Value = 0
t4(7) = 2
End If
If Check108.Value = 0 And Check109.Value = 0 And Check110.Value = 0 Then Check108.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check111_Click()
t4(8) = 0
If Check111.Value = 1 Then
Check112.Value = 0: Check113.Value = 0
t4(8) = 0
End If
If Check111.Value = 0 And Check112.Value = 0 And Check113.Value = 0 Then Check111.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check112_Click()
t4(8) = 0
If Check112.Value = 1 Then
Check111.Value = 0: Check113.Value = 0
t4(8) = 1
End If
If Check111.Value = 0 And Check112.Value = 0 And Check113.Value = 0 Then Check111.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check113_Click()
t4(8) = 0
If Check113.Value = 1 Then
Check112.Value = 0: Check111.Value = 0
t4(8) = 2
End If
If Check111.Value = 0 And Check112.Value = 0 And Check113.Value = 0 Then Check111.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check114_Click()
t4(9) = 0
If Check114.Value = 1 Then
Check115.Value = 0: Check116.Value = 0
t4(9) = 0
End If
If Check114.Value = 0 And Check115.Value = 0 And Check116.Value = 0 Then Check114.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check115_Click()
t4(9) = 0
If Check115.Value = 1 Then
Check114.Value = 0: Check116.Value = 0
t4(9) = 1
End If
If Check114.Value = 0 And Check115.Value = 0 And Check116.Value = 0 Then Check114.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check116_Click()
t4(9) = 0
If Check116.Value = 1 Then
Check115.Value = 0: Check114.Value = 0
t4(9) = 2
End If
If Check114.Value = 0 And Check115.Value = 0 And Check116.Value = 0 Then Check114.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check117_Click()
t4(10) = 0
If Check117.Value = 1 Then
Check118.Value = 0: Check119.Value = 0
t4(10) = 0
End If
If Check117.Value = 0 And Check118.Value = 0 And Check119.Value = 0 Then Check117.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check118_Click()
t4(10) = 0
If Check118.Value = 1 Then
Check117.Value = 0: Check119.Value = 0
t4(10) = 1
End If
If Check117.Value = 0 And Check118.Value = 0 And Check119.Value = 0 Then Check117.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check119_Click()
t4(10) = 0
If Check119.Value = 1 Then
Check118.Value = 0: Check117.Value = 0
t4(10) = 2
End If
If Check117.Value = 0 And Check118.Value = 0 And Check119.Value = 0 Then Check117.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check12_Click()
t5(7) = 0
If Check12.Value = 1 Then t5(7) = 1
Call list5add
Label5.Caption = "ÊÂ¼þ£º" + CStr(30 * t5(1) + 50 * t5(2) + 100 * t5(3) + 50 * t5(4) + 100 * t5(5) + 80 * t5(6) + 50 * t5(7))

End Sub

Private Sub Check120_Click()
t4(11) = 0
If Check120.Value = 1 Then
Check121.Value = 0: Check122.Value = 0
t4(11) = 0
End If
If Check120.Value = 0 And Check121.Value = 0 And Check122.Value = 0 Then Check120.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check121_Click()
t4(11) = 0
If Check121.Value = 1 Then
Check120.Value = 0: Check122.Value = 0
t4(11) = 1
End If
If Check120.Value = 0 And Check121.Value = 0 And Check122.Value = 0 Then Check120.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check122_Click()
t4(11) = 0
If Check122.Value = 1 Then
Check121.Value = 0: Check120.Value = 0
t4(11) = 2
End If
If Check120.Value = 0 And Check121.Value = 0 And Check122.Value = 0 Then Check120.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check123_Click()
t4(12) = 0
If Check123.Value = 1 Then
Check124.Value = 0: Check125.Value = 0
t4(12) = 0
End If
If Check123.Value = 0 And Check124.Value = 0 And Check125.Value = 0 Then Check123.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check124_Click()
t4(12) = 0
If Check124.Value = 1 Then
Check123.Value = 0: Check125.Value = 0
t4(12) = 1
End If
If Check123.Value = 0 And Check124.Value = 0 And Check125.Value = 0 Then Check123.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check125_Click()
t4(12) = 0
If Check125.Value = 1 Then
Check124.Value = 0: Check123.Value = 0
t4(12) = 2
End If
If Check123.Value = 0 And Check124.Value = 0 And Check125.Value = 0 Then Check123.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check126_Click()
t4(13) = 0
If Check126.Value = 1 Then
Check127.Value = 0: Check128.Value = 0
t4(13) = 0
End If
If Check126.Value = 0 And Check127.Value = 0 And Check128.Value = 0 Then Check126.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check127_Click()
t4(13) = 0
If Check127.Value = 1 Then
Check126.Value = 0: Check128.Value = 0
t4(13) = 1
End If
If Check126.Value = 0 And Check127.Value = 0 And Check128.Value = 0 Then Check126.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check128_Click()
t4(13) = 0
If Check128.Value = 1 Then
Check127.Value = 0: Check126.Value = 0
t4(13) = 2
End If
If Check126.Value = 0 And Check127.Value = 0 And Check128.Value = 0 Then Check126.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check13_Click()
If Check13.Value = 1 Then
Check14.Value = 0: f = True
End If
If Check13.Value = 0 And Check14.Value = 0 Then Check13.Value = 1
Call Text3_Change
End Sub

Private Sub Check14_Click()
If Check14.Value = 1 Then
Check13.Value = 0: f = False
End If
If Check13.Value = 0 And Check14.Value = 0 Then Check14.Value = 1
Call Text3_Change
End Sub

Private Sub Check2_Click()
t22 = 0
If Check2.Value = 1 Then t22 = 1
Call list1add
Label2.Caption = "½á¾Ö£º" + CStr(450 * t21 + 300 * t22 + 120 * t23 + 150 * t24 + 230 * t25)
End Sub

Private Sub Check20_Click()
t11 = 0
If Check20.Value = 1 Then
Check21.Value = 0: Check22.Value = 0: Check23.Value = 0: Check24.Value = 0: Check25.Value = 0: Check26.Value = 0: Check27.Value = 0: Check28.Value = 0: Check29.Value = 0:
t11 = 0
End If
If Check20.Value = 0 And Check21.Value = 0 And Check22.Value = 0 And Check23.Value = 0 And Check24.Value = 0 And Check25.Value = 0 And Check26.Value = 0 And Check27.Value = 0 And Check28.Value = 0 And Check29.Value = 0 Then Check20.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check21_Click()
t11 = 0
If Check21.Value = 1 Then
Check20.Value = 0: Check22.Value = 0: Check23.Value = 0: Check24.Value = 0: Check25.Value = 0: Check26.Value = 0: Check27.Value = 0: Check28.Value = 0: Check29.Value = 0
t11 = 1
End If
If Check20.Value = 0 And Check21.Value = 0 And Check22.Value = 0 And Check23.Value = 0 And Check24.Value = 0 And Check25.Value = 0 And Check26.Value = 0 And Check27.Value = 0 And Check28.Value = 0 And Check29.Value = 0 Then Check20.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check22_Click()
t11 = 0
If Check22.Value = 1 Then
Check21.Value = 0: Check20.Value = 0: Check23.Value = 0: Check24.Value = 0: Check25.Value = 0: Check26.Value = 0: Check27.Value = 0: Check28.Value = 0: Check29.Value = 0
t11 = 2
End If
If Check20.Value = 0 And Check21.Value = 0 And Check22.Value = 0 And Check23.Value = 0 And Check24.Value = 0 And Check25.Value = 0 And Check26.Value = 0 And Check27.Value = 0 And Check28.Value = 0 And Check29.Value = 0 Then Check20.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check23_Click()
t11 = 0
If Check23.Value = 1 Then
Check21.Value = 0: Check22.Value = 0: Check20.Value = 0: Check24.Value = 0: Check25.Value = 0: Check26.Value = 0: Check27.Value = 0: Check28.Value = 0: Check29.Value = 0
t11 = 3
End If
If Check20.Value = 0 And Check21.Value = 0 And Check22.Value = 0 And Check23.Value = 0 And Check24.Value = 0 And Check25.Value = 0 And Check26.Value = 0 And Check27.Value = 0 And Check28.Value = 0 And Check29.Value = 0 Then Check20.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check24_Click()
t11 = 0
If Check24.Value = 1 Then
Check21.Value = 0: Check22.Value = 0: Check23.Value = 0: Check20.Value = 0: Check25.Value = 0: Check26.Value = 0: Check27.Value = 0: Check28.Value = 0: Check29.Value = 0
t11 = 4
End If
If Check20.Value = 0 And Check21.Value = 0 And Check22.Value = 0 And Check23.Value = 0 And Check24.Value = 0 And Check25.Value = 0 And Check26.Value = 0 And Check27.Value = 0 And Check28.Value = 0 And Check29.Value = 0 Then Check20.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check25_Click()
t11 = 0
If Check25.Value = 1 Then
Check21.Value = 0: Check22.Value = 0: Check23.Value = 0: Check24.Value = 0: Check20.Value = 0: Check26.Value = 0: Check27.Value = 0: Check28.Value = 0: Check29.Value = 0
t11 = 5
End If
If Check20.Value = 0 And Check21.Value = 0 And Check22.Value = 0 And Check23.Value = 0 And Check24.Value = 0 And Check25.Value = 0 And Check26.Value = 0 And Check27.Value = 0 And Check28.Value = 0 And Check29.Value = 0 Then Check20.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check26_Click()
t11 = 0
If Check26.Value = 1 Then
Check21.Value = 0: Check22.Value = 0: Check23.Value = 0: Check24.Value = 0: Check25.Value = 0: Check20.Value = 0: Check27.Value = 0: Check28.Value = 0: Check29.Value = 0
t11 = 6
End If
If Check20.Value = 0 And Check21.Value = 0 And Check22.Value = 0 And Check23.Value = 0 And Check24.Value = 0 And Check25.Value = 0 And Check26.Value = 0 And Check27.Value = 0 And Check28.Value = 0 And Check29.Value = 0 Then Check20.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check27_Click()
t11 = 0
If Check27.Value = 1 Then
Check21.Value = 0: Check22.Value = 0: Check23.Value = 0: Check24.Value = 0: Check25.Value = 0: Check26.Value = 0: Check20.Value = 0: Check28.Value = 0: Check29.Value = 0
t11 = 7
End If
If Check20.Value = 0 And Check21.Value = 0 And Check22.Value = 0 And Check23.Value = 0 And Check24.Value = 0 And Check25.Value = 0 And Check26.Value = 0 And Check27.Value = 0 And Check28.Value = 0 And Check29.Value = 0 Then Check20.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check28_Click()
t11 = 0
If Check28.Value = 1 Then
Check21.Value = 0: Check22.Value = 0: Check23.Value = 0: Check24.Value = 0: Check25.Value = 0: Check26.Value = 0: Check27.Value = 0: Check20.Value = 0: Check29.Value = 0
t11 = 8
End If
If Check20.Value = 0 And Check21.Value = 0 And Check22.Value = 0 And Check23.Value = 0 And Check24.Value = 0 And Check25.Value = 0 And Check26.Value = 0 And Check27.Value = 0 And Check28.Value = 0 And Check29.Value = 0 Then Check20.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check29_Click()
t11 = 0
If Check29.Value = 1 Then
Check21.Value = 0: Check22.Value = 0: Check23.Value = 0: Check24.Value = 0: Check25.Value = 0: Check26.Value = 0: Check27.Value = 0: Check28.Value = 0: Check20.Value = 0
t11 = 9
End If
If Check20.Value = 0 And Check21.Value = 0 And Check22.Value = 0 And Check23.Value = 0 And Check24.Value = 0 And Check25.Value = 0 And Check26.Value = 0 And Check27.Value = 0 And Check28.Value = 0 And Check29.Value = 0 Then Check20.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check3_Click()
t23 = 0
If Check3.Value = 1 Then
Check4.Value = 0: Check5.Value = 0
t23 = 1
End If
Call list1add
Label2.Caption = "½á¾Ö£º" + CStr(450 * t21 + 300 * t22 + 120 * t23 + 150 * t24 + 230 * t25)
End Sub

Private Sub Check30_Click()
t12 = 0
If Check30.Value = 1 Then
Check31.Value = 0: Check32.Value = 0: Check33.Value = 0: Check34.Value = 0: Check35.Value = 0: Check36.Value = 0: Check37.Value = 0: Check38.Value = 0: Check39.Value = 0
t12 = 0
End If
If Check30.Value = 0 And Check31.Value = 0 And Check32.Value = 0 And Check33.Value = 0 And Check34.Value = 0 And Check35.Value = 0 And Check36.Value = 0 And Check37.Value = 0 And Check38.Value = 0 And Check39.Value = 0 Then Check30.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check31_Click()
t12 = 0
If Check31.Value = 1 Then
Check30.Value = 0: Check32.Value = 0: Check33.Value = 0: Check34.Value = 0: Check35.Value = 0: Check36.Value = 0: Check37.Value = 0: Check38.Value = 0: Check39.Value = 0
t12 = 1
End If
If Check30.Value = 0 And Check31.Value = 0 And Check32.Value = 0 And Check33.Value = 0 And Check34.Value = 0 And Check35.Value = 0 And Check36.Value = 0 And Check37.Value = 0 And Check38.Value = 0 And Check39.Value = 0 Then Check30.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check32_Click()
t12 = 0
If Check32.Value = 1 Then
Check31.Value = 0: Check30.Value = 0: Check33.Value = 0: Check34.Value = 0: Check35.Value = 0: Check36.Value = 0: Check37.Value = 0: Check38.Value = 0: Check39.Value = 0
t12 = 2
End If
If Check30.Value = 0 And Check31.Value = 0 And Check32.Value = 0 And Check33.Value = 0 And Check34.Value = 0 And Check35.Value = 0 And Check36.Value = 0 And Check37.Value = 0 And Check38.Value = 0 And Check39.Value = 0 Then Check30.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check33_Click()
t12 = 0
If Check33.Value = 1 Then
Check31.Value = 0: Check32.Value = 0: Check30.Value = 0: Check34.Value = 0: Check35.Value = 0: Check36.Value = 0: Check37.Value = 0: Check38.Value = 0: Check39.Value = 0
t12 = 3
End If
If Check30.Value = 0 And Check31.Value = 0 And Check32.Value = 0 And Check33.Value = 0 And Check34.Value = 0 And Check35.Value = 0 And Check36.Value = 0 And Check37.Value = 0 And Check38.Value = 0 And Check39.Value = 0 Then Check30.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check34_Click()
t12 = 0
If Check34.Value = 1 Then
Check31.Value = 0: Check32.Value = 0: Check33.Value = 0: Check30.Value = 0: Check35.Value = 0: Check36.Value = 0: Check37.Value = 0: Check38.Value = 0: Check39.Value = 0
t12 = 4
End If
If Check30.Value = 0 And Check31.Value = 0 And Check32.Value = 0 And Check33.Value = 0 And Check34.Value = 0 And Check35.Value = 0 And Check36.Value = 0 And Check37.Value = 0 And Check38.Value = 0 And Check39.Value = 0 Then Check30.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check35_Click()
t12 = 0
If Check35.Value = 1 Then
Check31.Value = 0: Check32.Value = 0: Check33.Value = 0: Check34.Value = 0: Check30.Value = 0: Check36.Value = 0: Check37.Value = 0: Check38.Value = 0: Check39.Value = 0
t12 = 5
End If
If Check30.Value = 0 And Check31.Value = 0 And Check32.Value = 0 And Check33.Value = 0 And Check34.Value = 0 And Check35.Value = 0 And Check36.Value = 0 And Check37.Value = 0 And Check38.Value = 0 And Check39.Value = 0 Then Check30.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check36_Click()
t12 = 0
If Check36.Value = 1 Then
Check31.Value = 0: Check32.Value = 0: Check33.Value = 0: Check34.Value = 0: Check35.Value = 0: Check30.Value = 0: Check37.Value = 0: Check38.Value = 0: Check39.Value = 0
t12 = 6
End If
If Check30.Value = 0 And Check31.Value = 0 And Check32.Value = 0 And Check33.Value = 0 And Check34.Value = 0 And Check35.Value = 0 And Check36.Value = 0 And Check37.Value = 0 And Check38.Value = 0 And Check39.Value = 0 Then Check30.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check37_Click()
t12 = 0
If Check37.Value = 1 Then
Check31.Value = 0: Check32.Value = 0: Check33.Value = 0: Check34.Value = 0: Check35.Value = 0: Check36.Value = 0: Check30.Value = 0: Check38.Value = 0: Check39.Value = 0
t12 = 7
End If
If Check30.Value = 0 And Check31.Value = 0 And Check32.Value = 0 And Check33.Value = 0 And Check34.Value = 0 And Check35.Value = 0 And Check36.Value = 0 And Check37.Value = 0 And Check38.Value = 0 And Check39.Value = 0 Then Check30.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check38_Click()
t12 = 0
If Check38.Value = 1 Then
Check31.Value = 0: Check32.Value = 0: Check33.Value = 0: Check34.Value = 0: Check35.Value = 0: Check36.Value = 0: Check37.Value = 0: Check30.Value = 0: Check39.Value = 0
t12 = 8
End If
If Check30.Value = 0 And Check31.Value = 0 And Check32.Value = 0 And Check33.Value = 0 And Check34.Value = 0 And Check35.Value = 0 And Check36.Value = 0 And Check37.Value = 0 And Check38.Value = 0 And Check39.Value = 0 Then Check30.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check39_Click()
t12 = 0
If Check39.Value = 1 Then
Check31.Value = 0: Check32.Value = 0: Check33.Value = 0: Check34.Value = 0: Check35.Value = 0: Check36.Value = 0: Check37.Value = 0: Check38.Value = 0: Check30.Value = 0
t12 = 9
End If
If Check30.Value = 0 And Check31.Value = 0 And Check32.Value = 0 And Check33.Value = 0 And Check34.Value = 0 And Check35.Value = 0 And Check36.Value = 0 And Check37.Value = 0 And Check38.Value = 0 And Check39.Value = 0 Then Check30.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check4_Click()
t24 = 0
If Check4.Value = 1 Then
Check3.Value = 0: Check5.Value = 0
t24 = 1
End If
Call list1add
Label2.Caption = "½á¾Ö£º" + CStr(450 * t21 + 300 * t22 + 120 * t23 + 150 * t24 + 230 * t25)
End Sub

Private Sub Check40_Click()
t13 = 0
If Check40.Value = 1 Then
Check41.Value = 0: Check42.Value = 0: Check43.Value = 0: Check44.Value = 0: Check45.Value = 0: Check46.Value = 0: Check47.Value = 0: Check48.Value = 0: Check49.Value = 0
t13 = 0
End If
If Check40.Value = 0 And Check41.Value = 0 And Check42.Value = 0 And Check43.Value = 0 And Check44.Value = 0 And Check45.Value = 0 And Check46.Value = 0 And Check47.Value = 0 And Check48.Value = 0 And Check49.Value = 0 Then Check40.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check41_Click()
t13 = 0
If Check41.Value = 1 Then
Check40.Value = 0: Check42.Value = 0: Check43.Value = 0: Check44.Value = 0: Check45.Value = 0: Check46.Value = 0: Check47.Value = 0: Check48.Value = 0: Check49.Value = 0
t13 = 1
End If
If Check40.Value = 0 And Check41.Value = 0 And Check42.Value = 0 And Check43.Value = 0 And Check44.Value = 0 And Check45.Value = 0 And Check46.Value = 0 And Check47.Value = 0 And Check48.Value = 0 And Check49.Value = 0 Then Check40.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check42_Click()
t13 = 0
If Check42.Value = 1 Then
Check41.Value = 0: Check40.Value = 0: Check43.Value = 0: Check44.Value = 0: Check45.Value = 0: Check46.Value = 0: Check47.Value = 0: Check48.Value = 0: Check49.Value = 0
t13 = 2
End If
If Check40.Value = 0 And Check41.Value = 0 And Check42.Value = 0 And Check43.Value = 0 And Check44.Value = 0 And Check45.Value = 0 And Check46.Value = 0 And Check47.Value = 0 And Check48.Value = 0 And Check49.Value = 0 Then Check40.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check43_Click()
t13 = 0
If Check43.Value = 1 Then
Check41.Value = 0: Check42.Value = 0: Check40.Value = 0: Check44.Value = 0: Check45.Value = 0: Check46.Value = 0: Check47.Value = 0: Check48.Value = 0: Check49.Value = 0
t13 = 3
End If
If Check40.Value = 0 And Check41.Value = 0 And Check42.Value = 0 And Check43.Value = 0 And Check44.Value = 0 And Check45.Value = 0 And Check46.Value = 0 And Check47.Value = 0 And Check48.Value = 0 And Check49.Value = 0 Then Check40.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check44_Click()
t13 = 0
If Check44.Value = 1 Then
Check41.Value = 0: Check42.Value = 0: Check43.Value = 0: Check40.Value = 0: Check45.Value = 0: Check46.Value = 0: Check47.Value = 0: Check48.Value = 0: Check49.Value = 0
t13 = 4
End If
If Check40.Value = 0 And Check41.Value = 0 And Check42.Value = 0 And Check43.Value = 0 And Check44.Value = 0 And Check45.Value = 0 And Check46.Value = 0 And Check47.Value = 0 And Check48.Value = 0 And Check49.Value = 0 Then Check40.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check45_Click()
t13 = 0
If Check45.Value = 1 Then
Check41.Value = 0: Check42.Value = 0: Check43.Value = 0: Check44.Value = 0: Check40.Value = 0: Check46.Value = 0: Check47.Value = 0: Check48.Value = 0: Check49.Value = 0
t13 = 5
End If
If Check40.Value = 0 And Check41.Value = 0 And Check42.Value = 0 And Check43.Value = 0 And Check44.Value = 0 And Check45.Value = 0 And Check46.Value = 0 And Check47.Value = 0 And Check48.Value = 0 And Check49.Value = 0 Then Check40.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check46_Click()
t13 = 0
If Check46.Value = 1 Then
Check41.Value = 0: Check42.Value = 0: Check43.Value = 0: Check44.Value = 0: Check45.Value = 0: Check40.Value = 0: Check47.Value = 0: Check48.Value = 0: Check49.Value = 0
t13 = 6
End If
If Check40.Value = 0 And Check41.Value = 0 And Check42.Value = 0 And Check43.Value = 0 And Check44.Value = 0 And Check45.Value = 0 And Check46.Value = 0 And Check47.Value = 0 And Check48.Value = 0 And Check49.Value = 0 Then Check40.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check47_Click()
t13 = 0
If Check47.Value = 1 Then
Check41.Value = 0: Check42.Value = 0: Check43.Value = 0: Check44.Value = 0: Check45.Value = 0: Check46.Value = 0: Check40.Value = 0: Check48.Value = 0: Check49.Value = 0
t13 = 7
End If
If Check40.Value = 0 And Check41.Value = 0 And Check42.Value = 0 And Check43.Value = 0 And Check44.Value = 0 And Check45.Value = 0 And Check46.Value = 0 And Check47.Value = 0 And Check48.Value = 0 And Check49.Value = 0 Then Check40.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check48_Click()
t13 = 0
If Check48.Value = 1 Then
Check41.Value = 0: Check42.Value = 0: Check43.Value = 0: Check44.Value = 0: Check45.Value = 0: Check46.Value = 0: Check47.Value = 0: Check40.Value = 0: Check49.Value = 0
t13 = 8
End If
If Check40.Value = 0 And Check41.Value = 0 And Check42.Value = 0 And Check43.Value = 0 And Check44.Value = 0 And Check45.Value = 0 And Check46.Value = 0 And Check47.Value = 0 And Check48.Value = 0 And Check49.Value = 0 Then Check40.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check49_Click()
t13 = 0
If Check49.Value = 1 Then
Check41.Value = 0: Check42.Value = 0: Check43.Value = 0: Check44.Value = 0: Check45.Value = 0: Check46.Value = 0: Check47.Value = 0: Check48.Value = 0: Check40.Value = 0
t13 = 9
End If
If Check40.Value = 0 And Check41.Value = 0 And Check42.Value = 0 And Check43.Value = 0 And Check44.Value = 0 And Check45.Value = 0 And Check46.Value = 0 And Check47.Value = 0 And Check48.Value = 0 And Check49.Value = 0 Then Check40.Value = 1
Call list2add
Label1.Caption = "ÁÙÊ±ÕÐÄ¼£º" + CStr(50 * t11 + 20 * t12 + 10 * t13)
End Sub

Private Sub Check5_Click()
t25 = 0
If Check5.Value = 1 Then
Check4.Value = 0: Check3.Value = 0
t25 = 1
End If
Call list1add
Label2.Caption = "½á¾Ö£º" + CStr(450 * t21 + 300 * t22 + 120 * t23 + 150 * t24 + 230 * t25)
End Sub

Private Sub Check50_Click()
t31 = 0
If Check50.Value = 1 Then
Check51.Value = 0: Check52.Value = 0: Check53.Value = 0: Check54.Value = 0: Check55.Value = 0: Check56.Value = 0: Check57.Value = 0: Check58.Value = 0: Check59.Value = 0
t31 = 0
End If
If Check50.Value = 0 And Check51.Value = 0 And Check52.Value = 0 And Check53.Value = 0 And Check54.Value = 0 And Check55.Value = 0 And Check56.Value = 0 And Check57.Value = 0 And Check58.Value = 0 And Check59.Value = 0 Then Check50.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check51_Click()
t31 = 0
If Check51.Value = 1 Then
Check50.Value = 0: Check52.Value = 0: Check53.Value = 0: Check54.Value = 0: Check55.Value = 0: Check56.Value = 0: Check57.Value = 0: Check58.Value = 0: Check59.Value = 0
t31 = 1
End If
If Check50.Value = 0 And Check51.Value = 0 And Check52.Value = 0 And Check53.Value = 0 And Check54.Value = 0 And Check55.Value = 0 And Check56.Value = 0 And Check57.Value = 0 And Check58.Value = 0 And Check59.Value = 0 Then Check50.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check52_Click()
t31 = 0
If Check52.Value = 1 Then
Check51.Value = 0: Check50.Value = 0: Check53.Value = 0: Check54.Value = 0: Check55.Value = 0: Check56.Value = 0: Check57.Value = 0: Check58.Value = 0: Check59.Value = 0
t31 = 2
End If
If Check50.Value = 0 And Check51.Value = 0 And Check52.Value = 0 And Check53.Value = 0 And Check54.Value = 0 And Check55.Value = 0 And Check56.Value = 0 And Check57.Value = 0 And Check58.Value = 0 And Check59.Value = 0 Then Check50.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check53_Click()
t31 = 0
If Check53.Value = 1 Then
Check51.Value = 0: Check52.Value = 0: Check50.Value = 0: Check54.Value = 0: Check55.Value = 0: Check56.Value = 0: Check57.Value = 0: Check58.Value = 0: Check59.Value = 0
t31 = 3
End If
If Check50.Value = 0 And Check51.Value = 0 And Check52.Value = 0 And Check53.Value = 0 And Check54.Value = 0 And Check55.Value = 0 And Check56.Value = 0 And Check57.Value = 0 And Check58.Value = 0 And Check59.Value = 0 Then Check50.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check54_Click()
t31 = 0
If Check54.Value = 1 Then
Check51.Value = 0: Check52.Value = 0: Check53.Value = 0: Check50.Value = 0: Check55.Value = 0: Check56.Value = 0: Check57.Value = 0: Check58.Value = 0: Check59.Value = 0
t31 = 4
End If
If Check50.Value = 0 And Check51.Value = 0 And Check52.Value = 0 And Check53.Value = 0 And Check54.Value = 0 And Check55.Value = 0 And Check56.Value = 0 And Check57.Value = 0 And Check58.Value = 0 And Check59.Value = 0 Then Check50.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check55_Click()
t31 = 0
If Check55.Value = 1 Then
Check51.Value = 0: Check52.Value = 0: Check53.Value = 0: Check54.Value = 0: Check50.Value = 0: Check56.Value = 0: Check57.Value = 0: Check58.Value = 0: Check59.Value = 0
t31 = 5
End If
If Check50.Value = 0 And Check51.Value = 0 And Check52.Value = 0 And Check53.Value = 0 And Check54.Value = 0 And Check55.Value = 0 And Check56.Value = 0 And Check57.Value = 0 And Check58.Value = 0 And Check59.Value = 0 Then Check50.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check56_Click()
t31 = 0
If Check56.Value = 1 Then
Check51.Value = 0: Check52.Value = 0: Check53.Value = 0: Check54.Value = 0: Check55.Value = 0: Check50.Value = 0: Check57.Value = 0: Check58.Value = 0: Check59.Value = 0
t31 = 6
End If
If Check50.Value = 0 And Check51.Value = 0 And Check52.Value = 0 And Check53.Value = 0 And Check54.Value = 0 And Check55.Value = 0 And Check56.Value = 0 And Check57.Value = 0 And Check58.Value = 0 And Check59.Value = 0 Then Check50.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check57_Click()
t31 = 0
If Check57.Value = 1 Then
Check51.Value = 0: Check52.Value = 0: Check53.Value = 0: Check54.Value = 0: Check55.Value = 0: Check56.Value = 0: Check50.Value = 0: Check58.Value = 0: Check59.Value = 0
t31 = 7
End If
If Check50.Value = 0 And Check51.Value = 0 And Check52.Value = 0 And Check53.Value = 0 And Check54.Value = 0 And Check55.Value = 0 And Check56.Value = 0 And Check57.Value = 0 And Check58.Value = 0 And Check59.Value = 0 Then Check50.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check58_Click()
t31 = 0
If Check58.Value = 1 Then
Check51.Value = 0: Check52.Value = 0: Check53.Value = 0: Check54.Value = 0: Check55.Value = 0: Check56.Value = 0: Check57.Value = 0: Check50.Value = 0: Check59.Value = 0
t31 = 8
End If
If Check50.Value = 0 And Check51.Value = 0 And Check52.Value = 0 And Check53.Value = 0 And Check54.Value = 0 And Check55.Value = 0 And Check56.Value = 0 And Check57.Value = 0 And Check58.Value = 0 And Check59.Value = 0 Then Check50.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check59_Click()
t31 = 0
If Check59.Value = 1 Then
Check51.Value = 0: Check52.Value = 0: Check53.Value = 0: Check54.Value = 0: Check55.Value = 0: Check56.Value = 0: Check57.Value = 0: Check58.Value = 0: Check50.Value = 0
t31 = 9
End If
If Check50.Value = 0 And Check51.Value = 0 And Check52.Value = 0 And Check53.Value = 0 And Check54.Value = 0 And Check55.Value = 0 And Check56.Value = 0 And Check57.Value = 0 And Check58.Value = 0 And Check59.Value = 0 Then Check50.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check6_Click()
t5(1) = 0
If Check6.Value = 1 Then t5(1) = 1
Call list5add
Label5.Caption = "ÊÂ¼þ£º" + CStr(30 * t5(1) + 50 * t5(2) + 100 * t5(3) + 50 * t5(4) + 100 * t5(5) + 80 * t5(6) + 50 * t5(7))
End Sub

Private Sub Check60_Click()
t32 = 0
If Check60.Value = 1 Then
Check61.Value = 0: Check62.Value = 0: Check63.Value = 0: Check64.Value = 0: Check65.Value = 0: Check66.Value = 0: Check67.Value = 0: Check68.Value = 0: Check69.Value = 0
t32 = 0
End If
If Check60.Value = 0 And Check61.Value = 0 And Check62.Value = 0 And Check63.Value = 0 And Check64.Value = 0 And Check65.Value = 0 And Check66.Value = 0 And Check67.Value = 0 And Check68.Value = 0 And Check69.Value = 0 Then Check60.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check61_Click()
t32 = 0
If Check61.Value = 1 Then
Check60.Value = 0: Check62.Value = 0: Check63.Value = 0: Check64.Value = 0: Check65.Value = 0: Check66.Value = 0: Check67.Value = 0: Check68.Value = 0: Check69.Value = 0
t32 = 1
End If
If Check60.Value = 0 And Check61.Value = 0 And Check62.Value = 0 And Check63.Value = 0 And Check64.Value = 0 And Check65.Value = 0 And Check66.Value = 0 And Check67.Value = 0 And Check68.Value = 0 And Check69.Value = 0 Then Check60.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check62_Click()
t32 = 0
If Check62.Value = 1 Then
Check61.Value = 0: Check60.Value = 0: Check63.Value = 0: Check64.Value = 0: Check65.Value = 0: Check66.Value = 0: Check67.Value = 0: Check68.Value = 0: Check69.Value = 0
t32 = 2
End If
If Check60.Value = 0 And Check61.Value = 0 And Check62.Value = 0 And Check63.Value = 0 And Check64.Value = 0 And Check65.Value = 0 And Check66.Value = 0 And Check67.Value = 0 And Check68.Value = 0 And Check69.Value = 0 Then Check60.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check63_Click()
t32 = 0
If Check63.Value = 1 Then
Check61.Value = 0: Check62.Value = 0: Check60.Value = 0: Check64.Value = 0: Check65.Value = 0: Check66.Value = 0: Check67.Value = 0: Check68.Value = 0: Check69.Value = 0
t32 = 3
End If
If Check60.Value = 0 And Check61.Value = 0 And Check62.Value = 0 And Check63.Value = 0 And Check64.Value = 0 And Check65.Value = 0 And Check66.Value = 0 And Check67.Value = 0 And Check68.Value = 0 And Check69.Value = 0 Then Check60.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check64_Click()
t32 = 0
If Check64.Value = 1 Then
Check61.Value = 0: Check62.Value = 0: Check63.Value = 0: Check60.Value = 0: Check65.Value = 0: Check66.Value = 0: Check67.Value = 0: Check68.Value = 0: Check69.Value = 0
t32 = 4
End If
If Check60.Value = 0 And Check61.Value = 0 And Check62.Value = 0 And Check63.Value = 0 And Check64.Value = 0 And Check65.Value = 0 And Check66.Value = 0 And Check67.Value = 0 And Check68.Value = 0 And Check69.Value = 0 Then Check60.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check65_Click()
t32 = 0
If Check65.Value = 1 Then
Check61.Value = 0: Check62.Value = 0: Check63.Value = 0: Check64.Value = 0: Check60.Value = 0: Check66.Value = 0: Check67.Value = 0: Check68.Value = 0: Check69.Value = 0
t32 = 5
End If
If Check60.Value = 0 And Check61.Value = 0 And Check62.Value = 0 And Check63.Value = 0 And Check64.Value = 0 And Check65.Value = 0 And Check66.Value = 0 And Check67.Value = 0 And Check68.Value = 0 And Check69.Value = 0 Then Check60.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check66_Click()
t32 = 0
If Check66.Value = 1 Then
Check61.Value = 0: Check62.Value = 0: Check63.Value = 0: Check64.Value = 0: Check65.Value = 0: Check60.Value = 0: Check67.Value = 0: Check68.Value = 0: Check69.Value = 0
t32 = 6
End If
If Check60.Value = 0 And Check61.Value = 0 And Check62.Value = 0 And Check63.Value = 0 And Check64.Value = 0 And Check65.Value = 0 And Check66.Value = 0 And Check67.Value = 0 And Check68.Value = 0 And Check69.Value = 0 Then Check60.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check67_Click()
t32 = 0
If Check67.Value = 1 Then
Check61.Value = 0: Check62.Value = 0: Check63.Value = 0: Check64.Value = 0: Check65.Value = 0: Check66.Value = 0: Check60.Value = 0: Check68.Value = 0: Check69.Value = 0
t32 = 7
End If
If Check60.Value = 0 And Check61.Value = 0 And Check62.Value = 0 And Check63.Value = 0 And Check64.Value = 0 And Check65.Value = 0 And Check66.Value = 0 And Check67.Value = 0 And Check68.Value = 0 And Check69.Value = 0 Then Check60.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check68_Click()
t32 = 0
If Check68.Value = 1 Then
Check61.Value = 0: Check62.Value = 0: Check63.Value = 0: Check64.Value = 0: Check65.Value = 0: Check66.Value = 0: Check67.Value = 0: Check60.Value = 0: Check69.Value = 0
t32 = 8
End If
If Check60.Value = 0 And Check61.Value = 0 And Check62.Value = 0 And Check63.Value = 0 And Check64.Value = 0 And Check65.Value = 0 And Check66.Value = 0 And Check67.Value = 0 And Check68.Value = 0 And Check69.Value = 0 Then Check60.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check69_Click()
t32 = 0
If Check69.Value = 1 Then
Check61.Value = 0: Check62.Value = 0: Check63.Value = 0: Check64.Value = 0: Check65.Value = 0: Check66.Value = 0: Check67.Value = 0: Check68.Value = 0: Check60.Value = 0
t32 = 9
End If
If Check60.Value = 0 And Check61.Value = 0 And Check62.Value = 0 And Check63.Value = 0 And Check64.Value = 0 And Check65.Value = 0 And Check66.Value = 0 And Check67.Value = 0 And Check68.Value = 0 And Check69.Value = 0 Then Check60.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check7_Click()
t5(2) = 0
If Check7.Value = 1 Then Check8.Value = 0: t5(2) = 1
Call list5add
Label5.Caption = "ÊÂ¼þ£º" + CStr(30 * t5(1) + 50 * t5(2) + 100 * t5(3) + 50 * t5(4) + 100 * t5(5) + 80 * t5(6) + 50 * t5(7))

End Sub

Private Sub Check70_Click()
t33 = 0
If Check70.Value = 1 Then
Check71.Value = 0: Check72.Value = 0: Check73.Value = 0: Check74.Value = 0: Check75.Value = 0: Check76.Value = 0: Check77.Value = 0: Check78.Value = 0: Check79.Value = 0
t33 = 0
End If
If Check70.Value = 0 And Check71.Value = 0 And Check72.Value = 0 And Check73.Value = 0 And Check74.Value = 0 And Check75.Value = 0 And Check76.Value = 0 And Check77.Value = 0 And Check78.Value = 0 And Check79.Value = 0 Then Check70.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check71_Click()
t33 = 0
If Check71.Value = 1 Then
Check70.Value = 0: Check72.Value = 0: Check73.Value = 0: Check74.Value = 0: Check75.Value = 0: Check76.Value = 0: Check77.Value = 0: Check78.Value = 0: Check79.Value = 0
t33 = 1
End If
If Check70.Value = 0 And Check71.Value = 0 And Check72.Value = 0 And Check73.Value = 0 And Check74.Value = 0 And Check75.Value = 0 And Check76.Value = 0 And Check77.Value = 0 And Check78.Value = 0 And Check79.Value = 0 Then Check70.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check72_Click()
t33 = 0
If Check72.Value = 1 Then
Check71.Value = 0: Check70.Value = 0: Check73.Value = 0: Check74.Value = 0: Check75.Value = 0: Check76.Value = 0: Check77.Value = 0: Check78.Value = 0: Check79.Value = 0
t33 = 2
End If
If Check70.Value = 0 And Check71.Value = 0 And Check72.Value = 0 And Check73.Value = 0 And Check74.Value = 0 And Check75.Value = 0 And Check76.Value = 0 And Check77.Value = 0 And Check78.Value = 0 And Check79.Value = 0 Then Check70.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check73_Click()
t33 = 0
If Check73.Value = 1 Then
Check71.Value = 0: Check72.Value = 0: Check70.Value = 0: Check74.Value = 0: Check75.Value = 0: Check76.Value = 0: Check77.Value = 0: Check78.Value = 0: Check79.Value = 0
t33 = 3
End If
If Check70.Value = 0 And Check71.Value = 0 And Check72.Value = 0 And Check73.Value = 0 And Check74.Value = 0 And Check75.Value = 0 And Check76.Value = 0 And Check77.Value = 0 And Check78.Value = 0 And Check79.Value = 0 Then Check70.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check74_Click()
t33 = 0
If Check74.Value = 1 Then
Check71.Value = 0: Check72.Value = 0: Check73.Value = 0: Check70.Value = 0: Check75.Value = 0: Check76.Value = 0: Check77.Value = 0: Check78.Value = 0: Check79.Value = 0
t33 = 4
End If
If Check70.Value = 0 And Check71.Value = 0 And Check72.Value = 0 And Check73.Value = 0 And Check74.Value = 0 And Check75.Value = 0 And Check76.Value = 0 And Check77.Value = 0 And Check78.Value = 0 And Check79.Value = 0 Then Check70.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check75_Click()
t33 = 0
If Check75.Value = 1 Then
Check71.Value = 0: Check72.Value = 0: Check73.Value = 0: Check74.Value = 0: Check70.Value = 0: Check76.Value = 0: Check77.Value = 0: Check78.Value = 0: Check79.Value = 0
t33 = 5
End If
If Check70.Value = 0 And Check71.Value = 0 And Check72.Value = 0 And Check73.Value = 0 And Check74.Value = 0 And Check75.Value = 0 And Check76.Value = 0 And Check77.Value = 0 And Check78.Value = 0 And Check79.Value = 0 Then Check70.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check76_Click()
t33 = 0
If Check76.Value = 1 Then
Check71.Value = 0: Check72.Value = 0: Check73.Value = 0: Check74.Value = 0: Check75.Value = 0: Check70.Value = 0: Check77.Value = 0: Check78.Value = 0: Check79.Value = 0
t33 = 6
End If
If Check70.Value = 0 And Check71.Value = 0 And Check72.Value = 0 And Check73.Value = 0 And Check74.Value = 0 And Check75.Value = 0 And Check76.Value = 0 And Check77.Value = 0 And Check78.Value = 0 And Check79.Value = 0 Then Check70.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check77_Click()
t33 = 0
If Check77.Value = 1 Then
Check71.Value = 0: Check72.Value = 0: Check73.Value = 0: Check74.Value = 0: Check75.Value = 0: Check76.Value = 0: Check70.Value = 0: Check78.Value = 0: Check79.Value = 0
t33 = 7
End If
If Check70.Value = 0 And Check71.Value = 0 And Check72.Value = 0 And Check73.Value = 0 And Check74.Value = 0 And Check75.Value = 0 And Check76.Value = 0 And Check77.Value = 0 And Check78.Value = 0 And Check79.Value = 0 Then Check70.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check78_Click()
t33 = 0
If Check78.Value = 1 Then
Check71.Value = 0: Check72.Value = 0: Check73.Value = 0: Check74.Value = 0: Check75.Value = 0: Check76.Value = 0: Check77.Value = 0: Check70.Value = 0: Check79.Value = 0
t33 = 8
End If
If Check70.Value = 0 And Check71.Value = 0 And Check72.Value = 0 And Check73.Value = 0 And Check74.Value = 0 And Check75.Value = 0 And Check76.Value = 0 And Check77.Value = 0 And Check78.Value = 0 And Check79.Value = 0 Then Check70.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check79_Click()
t33 = 0
If Check79.Value = 1 Then
Check71.Value = 0: Check72.Value = 0: Check73.Value = 0: Check74.Value = 0: Check75.Value = 0: Check76.Value = 0: Check77.Value = 0: Check78.Value = 0: Check70.Value = 0
t33 = 9
End If
If Check70.Value = 0 And Check71.Value = 0 And Check72.Value = 0 And Check73.Value = 0 And Check74.Value = 0 And Check75.Value = 0 And Check76.Value = 0 And Check77.Value = 0 And Check78.Value = 0 And Check79.Value = 0 Then Check70.Value = 1
Call list3add
Label3.Caption = "Òþ²Ø£º" + CStr(20 * (t31 + t32 + t33))
End Sub

Private Sub Check8_Click()
t5(3) = 0
If Check8.Value = 1 Then Check7.Value = 0: t5(3) = 1
Call list5add
Label5.Caption = "ÊÂ¼þ£º" + CStr(30 * t5(1) + 50 * t5(2) + 100 * t5(3) + 50 * t5(4) + 100 * t5(5) + 80 * t5(6) + 50 * t5(7))

End Sub

Private Sub Check80_Click()
t6 = 0
If Check80.Value = 1 Then
Check81.Value = 0: Check82.Value = 0: Check83.Value = 0: Check84.Value = 0: Check85.Value = 0: Check86.Value = 0: Check87.Value = 0: Check88.Value = 0: Check89.Value = 0
t6 = 0
End If
If Check80.Value = 0 And Check81.Value = 0 And Check82.Value = 0 And Check83.Value = 0 And Check84.Value = 0 And Check85.Value = 0 And Check86.Value = 0 And Check87.Value = 0 And Check88.Value = 0 And Check89.Value = 0 Then Check80.Value = 1
Label6.Caption = "ÆôÊ¾£º" + CStr(t6 * 50)
End Sub

Private Sub Check81_Click()
t6 = 0
If Check81.Value = 1 Then
Check80.Value = 0: Check82.Value = 0: Check83.Value = 0: Check84.Value = 0: Check85.Value = 0: Check86.Value = 0: Check87.Value = 0: Check88.Value = 0: Check89.Value = 0
t6 = 1
End If
If Check80.Value = 0 And Check81.Value = 0 And Check82.Value = 0 And Check83.Value = 0 And Check84.Value = 0 And Check85.Value = 0 And Check86.Value = 0 And Check87.Value = 0 And Check88.Value = 0 And Check89.Value = 0 Then Check80.Value = 1
Label6.Caption = "ÆôÊ¾£º" + CStr(t6 * 50)
End Sub

Private Sub Check82_Click()
t6 = 0
If Check82.Value = 1 Then
Check81.Value = 0: Check80.Value = 0: Check83.Value = 0: Check84.Value = 0: Check85.Value = 0: Check86.Value = 0: Check87.Value = 0: Check88.Value = 0: Check89.Value = 0
t6 = 2
End If
If Check80.Value = 0 And Check81.Value = 0 And Check82.Value = 0 And Check83.Value = 0 And Check84.Value = 0 And Check85.Value = 0 And Check86.Value = 0 And Check87.Value = 0 And Check88.Value = 0 And Check89.Value = 0 Then Check80.Value = 1
Label6.Caption = "ÆôÊ¾£º" + CStr(t6 * 50)
End Sub

Private Sub Check83_Click()
t6 = 0
If Check83.Value = 1 Then
Check81.Value = 0: Check82.Value = 0: Check80.Value = 0: Check84.Value = 0: Check85.Value = 0: Check86.Value = 0: Check87.Value = 0: Check88.Value = 0: Check89.Value = 0
t6 = 3
End If
If Check80.Value = 0 And Check81.Value = 0 And Check82.Value = 0 And Check83.Value = 0 And Check84.Value = 0 And Check85.Value = 0 And Check86.Value = 0 And Check87.Value = 0 And Check88.Value = 0 And Check89.Value = 0 Then Check80.Value = 1
Label6.Caption = "ÆôÊ¾£º" + CStr(t6 * 50)
End Sub

Private Sub Check84_Click()
t6 = 0
If Check84.Value = 1 Then
Check81.Value = 0: Check82.Value = 0: Check83.Value = 0: Check80.Value = 0: Check85.Value = 0: Check86.Value = 0: Check87.Value = 0: Check88.Value = 0: Check89.Value = 0
t6 = 4
End If
If Check80.Value = 0 And Check81.Value = 0 And Check82.Value = 0 And Check83.Value = 0 And Check84.Value = 0 And Check85.Value = 0 And Check86.Value = 0 And Check87.Value = 0 And Check88.Value = 0 And Check89.Value = 0 Then Check80.Value = 1
Label6.Caption = "ÆôÊ¾£º" + CStr(t6 * 50)
End Sub

Private Sub Check85_Click()
t6 = 0
If Check85.Value = 1 Then
Check81.Value = 0: Check82.Value = 0: Check83.Value = 0: Check84.Value = 0: Check80.Value = 0: Check86.Value = 0: Check87.Value = 0: Check88.Value = 0: Check89.Value = 0
t6 = 5
End If
If Check80.Value = 0 And Check81.Value = 0 And Check82.Value = 0 And Check83.Value = 0 And Check84.Value = 0 And Check85.Value = 0 And Check86.Value = 0 And Check87.Value = 0 And Check88.Value = 0 And Check89.Value = 0 Then Check80.Value = 1
Label6.Caption = "ÆôÊ¾£º" + CStr(t6 * 50)
End Sub

Private Sub Check86_Click()
t6 = 0
If Check86.Value = 1 Then
Check81.Value = 0: Check82.Value = 0: Check83.Value = 0: Check84.Value = 0: Check85.Value = 0: Check80.Value = 0: Check87.Value = 0: Check88.Value = 0: Check89.Value = 0
t6 = 6
End If
If Check80.Value = 0 And Check81.Value = 0 And Check82.Value = 0 And Check83.Value = 0 And Check84.Value = 0 And Check85.Value = 0 And Check86.Value = 0 And Check87.Value = 0 And Check88.Value = 0 And Check89.Value = 0 Then Check80.Value = 1
Label6.Caption = "ÆôÊ¾£º" + CStr(t6 * 50)
End Sub

Private Sub Check87_Click()
t6 = 0
If Check87.Value = 1 Then
Check81.Value = 0: Check82.Value = 0: Check83.Value = 0: Check84.Value = 0: Check85.Value = 0: Check86.Value = 0: Check80.Value = 0: Check88.Value = 0: Check89.Value = 0
t6 = 7
End If
If Check80.Value = 0 And Check81.Value = 0 And Check82.Value = 0 And Check83.Value = 0 And Check84.Value = 0 And Check85.Value = 0 And Check86.Value = 0 And Check87.Value = 0 And Check88.Value = 0 And Check89.Value = 0 Then Check80.Value = 1
Label6.Caption = "ÆôÊ¾£º" + CStr(t6 * 50)
End Sub

Private Sub Check88_Click()
t6 = 0
If Check88.Value = 1 Then
Check81.Value = 0: Check82.Value = 0: Check83.Value = 0: Check84.Value = 0: Check85.Value = 0: Check86.Value = 0: Check87.Value = 0: Check80.Value = 0: Check89.Value = 0
t6 = 8
End If
If Check80.Value = 0 And Check81.Value = 0 And Check82.Value = 0 And Check83.Value = 0 And Check84.Value = 0 And Check85.Value = 0 And Check86.Value = 0 And Check87.Value = 0 And Check88.Value = 0 And Check89.Value = 0 Then Check80.Value = 1
Label6.Caption = "ÆôÊ¾£º" + CStr(t6 * 50)
End Sub

Private Sub Check89_Click()
t6 = 0
If Check89.Value = 1 Then
Check81.Value = 0: Check82.Value = 0: Check83.Value = 0: Check84.Value = 0: Check85.Value = 0: Check86.Value = 0: Check87.Value = 0: Check88.Value = 0: Check80.Value = 0
t6 = 9
End If
If Check80.Value = 0 And Check81.Value = 0 And Check82.Value = 0 And Check83.Value = 0 And Check84.Value = 0 And Check85.Value = 0 And Check86.Value = 0 And Check87.Value = 0 And Check88.Value = 0 And Check89.Value = 0 Then Check80.Value = 1
Label6.Caption = "ÆôÊ¾£º" + CStr(t6 * 50)
End Sub

Private Sub Check9_Click()
t5(4) = 0
If Check9.Value = 1 Then Check10.Value = 0: t5(4) = 1
Call list5add
Label5.Caption = "ÊÂ¼þ£º" + CStr(30 * t5(1) + 50 * t5(2) + 100 * t5(3) + 50 * t5(4) + 100 * t5(5) + 80 * t5(6) + 50 * t5(7))

End Sub

Private Sub Check90_Click()
t4(1) = 0
If Check90.Value = 1 Then
Check91.Value = 0: Check92.Value = 0
t4(1) = 0
End If
If Check90.Value = 0 And Check91.Value = 0 And Check92.Value = 0 Then Check90.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check91_Click()
t4(1) = 0
If Check91.Value = 1 Then
Check90.Value = 0: Check92.Value = 0
t4(1) = 1
End If
If Check90.Value = 0 And Check91.Value = 0 And Check92.Value = 0 Then Check90.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check92_Click()
t4(1) = 0
If Check92.Value = 1 Then
Check90.Value = 0: Check91.Value = 0
t4(1) = 2
End If
If Check90.Value = 0 And Check91.Value = 0 And Check92.Value = 0 Then Check90.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check93_Click()
t4(2) = 0
If Check93.Value = 1 Then
Check94.Value = 0: Check95.Value = 0
t4(2) = 0
End If
If Check93.Value = 0 And Check94.Value = 0 And Check95.Value = 0 Then Check93.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check94_Click()
t4(2) = 0
If Check94.Value = 1 Then
Check93.Value = 0: Check95.Value = 0
t4(2) = 1
End If
If Check93.Value = 0 And Check94.Value = 0 And Check95.Value = 0 Then Check93.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check95_Click()
t4(2) = 0
If Check95.Value = 1 Then
Check94.Value = 0: Check93.Value = 0
t4(2) = 2
End If
If Check93.Value = 0 And Check94.Value = 0 And Check95.Value = 0 Then Check93.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check96_Click()
t4(3) = 0
If Check96.Value = 1 Then
Check97.Value = 0: Check98.Value = 0
t4(3) = 0
End If
If Check96.Value = 0 And Check97.Value = 0 And Check98.Value = 0 Then Check96.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check97_Click()
t4(3) = 0
If Check97.Value = 1 Then
Check96.Value = 0: Check98.Value = 0
t4(3) = 1
End If
If Check96.Value = 0 And Check97.Value = 0 And Check98.Value = 0 Then Check96.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check98_Click()
t4(3) = 0
If Check98.Value = 1 Then
Check97.Value = 0: Check96.Value = 0
t4(3) = 2
End If
If Check96.Value = 0 And Check97.Value = 0 And Check98.Value = 0 Then Check96.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Check99_Click()
t4(4) = 0
If Check99.Value = 1 Then
Check100.Value = 0: Check101.Value = 0
t4(4) = 0
End If
If Check99.Value = 0 And Check100.Value = 0 And Check101.Value = 0 Then Check99.Value = 1
Label4.Caption = "½ô¼±£º" + CStr(t4(1) * 40 + t4(2) * 40 + t4(3) * 100 + t4(4) * 60 + t4(5) * 40 + t4(6) * 50 + t4(7) * 120 + t4(8) * 90 + t4(9) * 90 + t4(10) * 90 + t4(11) * 50 + t4(12) * 150 + t4(13) * 80)
Call list4add
End Sub

Private Sub Command1_Click()
Call begin
End Sub

Private Sub Form_Load()
Call begin
End Sub

Private Sub Label1_Change()
Call total
End Sub
Private Sub Label2_Change()
Call total
End Sub
Private Sub Label3_Change()
Call total
End Sub
Private Sub Label4_Change()
Call total
End Sub
Private Sub Label5_Change()
Call total
End Sub
Private Sub Label6_Change()
Call total
End Sub
Private Sub Label7_Change()
Call total
End Sub
Private Sub Label8_Change()
Call total
End Sub
Private Sub Label9_Change()
Call total
End Sub

Private Sub Text1_Change()
Dim i As Integer
If Len(Text1.Text) <> 0 Then
    For i = 1 To Len(Text1.Text)
        If Mid(Text1.Text, i, 1) < "0" Or Mid(Text1.Text, i, 1) > "9" Then If (i = 1) Then Text1.Text = Mid(Text1.Text, 2) Else Text1.Text = Mid(Text1.Text, 1, i - 1) + Mid(Text1.Text, i + 1)
    Next i
    If Len(Text1.Text) = 0 Then t7 = 0 Else t7 = Int(Text1.Text)
Else
    t7 = 0
End If
Label7.Caption = "²ØÆ·£º" + CStr(10 * t7)
End Sub

Private Sub Text2_Change()
Dim i As Integer
If Len(Text2.Text) <> 0 Then
    For i = 1 To Len(Text2.Text)
        If Mid(Text2.Text, i, 1) < "0" Or Mid(Text2.Text, i, 1) > "9" Then If (i = 1) Then Text2.Text = Mid(Text2.Text, 2) Else Text2.Text = Mid(Text2.Text, 1, i - 1) + Mid(Text2.Text, i + 1)
    Next i
    If Len(Text2.Text) = 0 Then t8 = 0 Else t8 = Int(Text2.Text)
Else
    t8 = 0
End If
Label8.Caption = "½áËã·Ö£º" + CStr(t8)
End Sub

Private Sub Text3_Change()
Dim i As Integer
If Len(Text3.Text) <> 0 Then
    For i = 1 To Len(Text3.Text)
        If Mid(Text3.Text, i, 1) < "0" Or Mid(Text3.Text, i, 1) > "9" Then If (i = 1) Then Text3.Text = Mid(Text3.Text, 2) Else Text3.Text = Mid(Text3.Text, 1, i - 1) + Mid(Text3.Text, i + 1)
    Next i
    If Len(Text3.Text) = 0 Then t9 = 0 Else t9 = Int(Text3.Text)
Else
    t9 = 0
End If
If f Then Label9.Caption = "ÐÞÕý·Ö£º" + CStr(t9) Else Label9.Caption = "ÐÞÕý·Ö£º-" + CStr(t9)
End Sub
