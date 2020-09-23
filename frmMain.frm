VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Map Editor"
   ClientHeight    =   4590
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgCD 
      Left            =   4080
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   164
      Left            =   5160
      TabIndex        =   172
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   163
      Left            =   4800
      TabIndex        =   171
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   162
      Left            =   4440
      TabIndex        =   170
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   161
      Left            =   4080
      TabIndex        =   169
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   160
      Left            =   3720
      TabIndex        =   168
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   159
      Left            =   3360
      TabIndex        =   167
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   158
      Left            =   3000
      TabIndex        =   166
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   157
      Left            =   2640
      TabIndex        =   165
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   156
      Left            =   2280
      TabIndex        =   164
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   155
      Left            =   1920
      TabIndex        =   163
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   154
      Left            =   1560
      TabIndex        =   162
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   153
      Left            =   1200
      TabIndex        =   161
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   152
      Left            =   840
      TabIndex        =   160
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   151
      Left            =   480
      TabIndex        =   159
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   150
      Left            =   120
      TabIndex        =   158
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   149
      Left            =   5160
      TabIndex        =   157
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   148
      Left            =   4800
      TabIndex        =   156
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   147
      Left            =   4440
      TabIndex        =   155
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   146
      Left            =   4080
      TabIndex        =   154
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   145
      Left            =   3720
      TabIndex        =   153
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   144
      Left            =   3360
      TabIndex        =   152
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   143
      Left            =   3000
      TabIndex        =   151
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   142
      Left            =   2640
      TabIndex        =   150
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   141
      Left            =   2280
      TabIndex        =   149
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   140
      Left            =   1920
      TabIndex        =   148
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   139
      Left            =   1560
      TabIndex        =   147
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   138
      Left            =   1200
      TabIndex        =   146
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   137
      Left            =   840
      TabIndex        =   145
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   136
      Left            =   480
      TabIndex        =   144
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   135
      Left            =   120
      TabIndex        =   143
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   134
      Left            =   5160
      TabIndex        =   142
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   133
      Left            =   4800
      TabIndex        =   141
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   132
      Left            =   4440
      TabIndex        =   140
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   131
      Left            =   4080
      TabIndex        =   139
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   130
      Left            =   3720
      TabIndex        =   138
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   129
      Left            =   3360
      TabIndex        =   137
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   128
      Left            =   3000
      TabIndex        =   136
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   127
      Left            =   2640
      TabIndex        =   135
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   126
      Left            =   2280
      TabIndex        =   134
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   125
      Left            =   1920
      TabIndex        =   133
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   124
      Left            =   1560
      TabIndex        =   132
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   123
      Left            =   1200
      TabIndex        =   131
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   122
      Left            =   840
      TabIndex        =   130
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   121
      Left            =   480
      TabIndex        =   129
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   120
      Left            =   120
      TabIndex        =   128
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   119
      Left            =   5160
      TabIndex        =   127
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   118
      Left            =   4800
      TabIndex        =   126
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   117
      Left            =   4440
      TabIndex        =   125
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   116
      Left            =   4080
      TabIndex        =   124
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   115
      Left            =   3720
      TabIndex        =   123
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   114
      Left            =   3360
      TabIndex        =   122
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   113
      Left            =   3000
      TabIndex        =   121
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   112
      Left            =   2640
      TabIndex        =   120
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   111
      Left            =   2280
      TabIndex        =   119
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   110
      Left            =   1920
      TabIndex        =   118
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   109
      Left            =   1560
      TabIndex        =   117
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   108
      Left            =   1200
      TabIndex        =   116
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   107
      Left            =   840
      TabIndex        =   115
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   106
      Left            =   480
      TabIndex        =   114
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   105
      Left            =   120
      TabIndex        =   113
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   104
      Left            =   5160
      TabIndex        =   112
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   103
      Left            =   4800
      TabIndex        =   111
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   102
      Left            =   4440
      TabIndex        =   110
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   101
      Left            =   4080
      TabIndex        =   109
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   100
      Left            =   3720
      TabIndex        =   108
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5040
      Picture         =   "frmMain.frx":0000
      Top             =   4120
      Width           =   480
   End
   Begin VB.Label lblRed 
      BackColor       =   &H000000C0&
      Height          =   255
      Left            =   2880
      TabIndex        =   107
      ToolTipText     =   "Insert End of Maze Tool"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Finish"
      Height          =   195
      Left            =   3240
      TabIndex        =   106
      Top             =   4220
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Start"
      Height          =   195
      Left            =   2400
      TabIndex        =   105
      Top             =   4220
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Wall"
      Height          =   195
      Left            =   1560
      TabIndex        =   104
      Top             =   4220
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Unused"
      Height          =   195
      Left            =   480
      TabIndex        =   103
      Top             =   4220
      Width           =   555
   End
   Begin VB.Label lblGreen 
      BackColor       =   &H00008000&
      Height          =   255
      Left            =   2040
      TabIndex        =   102
      ToolTipText     =   "Insert Start of Maze Tool"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblGrey 
      BackColor       =   &H00808080&
      Height          =   255
      Left            =   1200
      TabIndex        =   101
      ToolTipText     =   "Insert Wall Tool"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblBlack 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   100
      ToolTipText     =   "Unused Slot Tool"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   99
      Left            =   3360
      TabIndex        =   99
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   98
      Left            =   3000
      TabIndex        =   98
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   97
      Left            =   2640
      TabIndex        =   97
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   96
      Left            =   2280
      TabIndex        =   96
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   95
      Left            =   1920
      TabIndex        =   95
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   94
      Left            =   1560
      TabIndex        =   94
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   93
      Left            =   1200
      TabIndex        =   93
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   92
      Left            =   840
      TabIndex        =   92
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   91
      Left            =   480
      TabIndex        =   91
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   90
      Left            =   120
      TabIndex        =   90
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   89
      Left            =   5160
      TabIndex        =   89
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   88
      Left            =   4800
      TabIndex        =   88
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   87
      Left            =   4440
      TabIndex        =   87
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   86
      Left            =   4080
      TabIndex        =   86
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   85
      Left            =   3720
      TabIndex        =   85
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   84
      Left            =   3360
      TabIndex        =   84
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   83
      Left            =   3000
      TabIndex        =   83
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   82
      Left            =   2640
      TabIndex        =   82
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   81
      Left            =   2280
      TabIndex        =   81
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   80
      Left            =   1920
      TabIndex        =   80
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   79
      Left            =   1560
      TabIndex        =   79
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   78
      Left            =   1200
      TabIndex        =   78
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   77
      Left            =   840
      TabIndex        =   77
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   76
      Left            =   480
      TabIndex        =   76
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   75
      Left            =   120
      TabIndex        =   75
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   74
      Left            =   5160
      TabIndex        =   74
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   73
      Left            =   4800
      TabIndex        =   73
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   72
      Left            =   4440
      TabIndex        =   72
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   71
      Left            =   4080
      TabIndex        =   71
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   70
      Left            =   3720
      TabIndex        =   70
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   69
      Left            =   3360
      TabIndex        =   69
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   68
      Left            =   3000
      TabIndex        =   68
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   67
      Left            =   2640
      TabIndex        =   67
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   66
      Left            =   2280
      TabIndex        =   66
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   65
      Left            =   1920
      TabIndex        =   65
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   64
      Left            =   1560
      TabIndex        =   64
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   63
      Left            =   1200
      TabIndex        =   63
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   62
      Left            =   840
      TabIndex        =   62
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   61
      Left            =   480
      TabIndex        =   61
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   60
      Left            =   120
      TabIndex        =   60
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   59
      Left            =   5160
      TabIndex        =   59
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   58
      Left            =   4800
      TabIndex        =   58
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   57
      Left            =   4440
      TabIndex        =   57
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   56
      Left            =   4080
      TabIndex        =   56
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   55
      Left            =   3720
      TabIndex        =   55
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   54
      Left            =   3360
      TabIndex        =   54
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   53
      Left            =   3000
      TabIndex        =   53
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   52
      Left            =   2640
      TabIndex        =   52
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   51
      Left            =   2280
      TabIndex        =   51
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   50
      Left            =   1920
      TabIndex        =   50
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   49
      Left            =   1560
      TabIndex        =   49
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   48
      Left            =   1200
      TabIndex        =   48
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   47
      Left            =   840
      TabIndex        =   47
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   46
      Left            =   480
      TabIndex        =   46
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   45
      Left            =   120
      TabIndex        =   45
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   44
      Left            =   5160
      TabIndex        =   44
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   43
      Left            =   4800
      TabIndex        =   43
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   42
      Left            =   4440
      TabIndex        =   42
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   41
      Left            =   4080
      TabIndex        =   41
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   40
      Left            =   3720
      TabIndex        =   40
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   39
      Left            =   3360
      TabIndex        =   39
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   38
      Left            =   3000
      TabIndex        =   38
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   37
      Left            =   2640
      TabIndex        =   37
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   36
      Left            =   2280
      TabIndex        =   36
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   35
      Left            =   1920
      TabIndex        =   35
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   34
      Left            =   1560
      TabIndex        =   34
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   33
      Left            =   1200
      TabIndex        =   33
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   32
      Left            =   840
      TabIndex        =   32
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   31
      Left            =   480
      TabIndex        =   31
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   30
      Left            =   120
      TabIndex        =   30
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   29
      Left            =   5160
      TabIndex        =   29
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   28
      Left            =   4800
      TabIndex        =   28
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   27
      Left            =   4440
      TabIndex        =   27
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   26
      Left            =   4080
      TabIndex        =   26
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   25
      Left            =   3720
      TabIndex        =   25
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   24
      Left            =   3360
      TabIndex        =   24
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   23
      Left            =   3000
      TabIndex        =   23
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   22
      Left            =   2640
      TabIndex        =   22
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   21
      Left            =   2280
      TabIndex        =   21
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   20
      Left            =   1920
      TabIndex        =   20
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   19
      Left            =   1560
      TabIndex        =   19
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   18
      Left            =   1200
      TabIndex        =   18
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   17
      Left            =   840
      TabIndex        =   17
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   16
      Left            =   480
      TabIndex        =   16
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   15
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   14
      Left            =   5160
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   13
      Left            =   4800
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   12
      Left            =   4440
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   11
      Left            =   4080
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   10
      Left            =   3720
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   9
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   8
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   7
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsMapOptions 
         Caption         =   "&Map Options..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = App.Title & " [Untitled]"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Label2_Click()
    'Show which tool is selected
    lblBlack.BorderStyle = 1
    lblGrey.BorderStyle = 0
    lblGreen.BorderStyle = 0
    lblRed.BorderStyle = 0
End Sub

Private Sub Label3_Click()
    'Show which tool is selected
    lblBlack.BorderStyle = 0
    lblGrey.BorderStyle = 1
    lblGreen.BorderStyle = 0
    lblRed.BorderStyle = 0
End Sub

Private Sub Label4_Click()
    'Show which tool is selected
    lblBlack.BorderStyle = 0
    lblGrey.BorderStyle = 0
    lblGreen.BorderStyle = 1
    lblRed.BorderStyle = 0
End Sub

Private Sub Label5_Click()
    'Show which tool is selected
    lblBlack.BorderStyle = 0
    lblGrey.BorderStyle = 0
    lblGreen.BorderStyle = 0
    lblRed.BorderStyle = 1
End Sub

Private Sub lblBlack_Click()
    'Show which tool is selected
    lblBlack.BorderStyle = 1
    lblGrey.BorderStyle = 0
    lblGreen.BorderStyle = 0
    lblRed.BorderStyle = 0
End Sub

Private Sub lblGreen_Click()
    'Show which tool is selected
    lblBlack.BorderStyle = 0
    lblGrey.BorderStyle = 0
    lblGreen.BorderStyle = 1
    lblRed.BorderStyle = 0
End Sub

Private Sub lblGrey_Click()
    'Show which tool is selected
    lblBlack.BorderStyle = 0
    lblGrey.BorderStyle = 1
    lblGreen.BorderStyle = 0
    lblRed.BorderStyle = 0
End Sub

Private Sub lblMap_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Determine the text of the tooltip
    Select Case lblMap(Index).BackColor
        Case black
            lblMap(Index).ToolTipText = "Unused Slot " & lblMap(Index).Index + 1
        Case grey
            lblMap(Index).ToolTipText = "Wall"
        Case green
            lblMap(Index).ToolTipText = "Start"
        Case red
            lblMap(Index).ToolTipText = "Finish"
    End Select
End Sub

Private Sub lblMap_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        'Convienient way to erase mistakes
        lblMap(Index).BackColor = black
    Else
        Dim i As Integer
        For i = 0 To 164
            'If we already have a start point or an end
            'point then we don't need another one so
            'erase the old one and paint the new one
            If lblRed.BorderStyle = 1 Then
                If lblMap(i).BackColor = red Then
                    lblMap(i).BackColor = black
                    i = 164
                End If
            ElseIf lblGreen.BorderStyle = 1 Then
                If lblMap(i).BackColor = green Then
                    lblMap(i).BackColor = black
                    i = 164
                End If
            Else
                i = 164
            End If
        Next i
        
        'Determine which tool has been selected
        If lblBlack.BorderStyle = 1 Then
            lblMap(Index).BackColor = black
        ElseIf lblGrey.BorderStyle = 1 Then
            lblMap(Index).BackColor = grey
        ElseIf lblGreen.BorderStyle = 1 Then
            lblMap(Index).BackColor = green
        Else
            lblMap(Index).BackColor = red
        End If
    End If
End Sub

Private Sub lblRed_Click()
    'Show which tool is selected
    lblBlack.BorderStyle = 0
    lblGrey.BorderStyle = 0
    lblGreen.BorderStyle = 0
    lblRed.BorderStyle = 1
End Sub

Private Sub mnuFileExit_Click()
    'User selected exit from the file menu Abort! Abort! Abort!
    End
End Sub

Private Sub mnuFileNew_Click()
    Dim i As Integer
    'Start off with a new canvas...
    For i = 0 To 164
        lblMap(i).BackColor = black
    Next i
    '...as well as new map information
    frmMapOptions.txtAuthor = ""
    frmMapOptions.txtPassword = ""
    frmMapOptions.txtTitle = ""
    frmMapOptions.cboDifficulty.ListIndex = 0
    Unload frmMapOptions
    Me.Caption = App.Title & " [Untitled]"
End Sub

Private Sub mnuFileOpen_Click()
    With dlgCD
        .Filter = "Map Files (*.map)|*.map|All Files (*.*)|*.*"
        .ShowOpen
    End With
    
    If dlgCD.FileName <> "" Then
        Dim FileNum As Long
        FileNum = FreeFile()
        Open dlgCD.FileName For Binary As #FileNum Len = Len(dlgCD.FileName)
        Get #FileNum, , Map
        Close #FileNum
        
        'A crude way to open a file, but effective and since
        'the map file is very small then it will open very fast
        If Map.Password <> "" Then
            frmLogin.Show vbModal, frmMain
            Exit Sub
        Else
            lblMap(0).BackColor = Map.Slot0
            lblMap(1).BackColor = Map.Slot1
            lblMap(2).BackColor = Map.Slot2
            lblMap(3).BackColor = Map.Slot3
            lblMap(4).BackColor = Map.Slot4
            lblMap(5).BackColor = Map.Slot5
            lblMap(6).BackColor = Map.Slot6
            lblMap(7).BackColor = Map.Slot7
            lblMap(8).BackColor = Map.Slot8
            lblMap(9).BackColor = Map.Slot9
            lblMap(10).BackColor = Map.Slot10
            lblMap(11).BackColor = Map.Slot11
            lblMap(12).BackColor = Map.Slot12
            lblMap(13).BackColor = Map.Slot13
            lblMap(14).BackColor = Map.Slot14
            lblMap(15).BackColor = Map.Slot15
            lblMap(16).BackColor = Map.Slot16
            lblMap(17).BackColor = Map.Slot17
            lblMap(18).BackColor = Map.Slot18
            lblMap(19).BackColor = Map.Slot19
            lblMap(20).BackColor = Map.Slot20
            lblMap(21).BackColor = Map.Slot21
            lblMap(22).BackColor = Map.Slot22
            lblMap(23).BackColor = Map.Slot23
            lblMap(24).BackColor = Map.Slot24
            lblMap(25).BackColor = Map.Slot25
            lblMap(26).BackColor = Map.Slot26
            lblMap(27).BackColor = Map.Slot27
            lblMap(28).BackColor = Map.Slot28
            lblMap(29).BackColor = Map.Slot29
            lblMap(30).BackColor = Map.Slot30
            lblMap(31).BackColor = Map.Slot31
            lblMap(32).BackColor = Map.Slot32
            lblMap(33).BackColor = Map.Slot33
            lblMap(34).BackColor = Map.Slot34
            lblMap(35).BackColor = Map.Slot35
            lblMap(36).BackColor = Map.Slot36
            lblMap(37).BackColor = Map.Slot37
            lblMap(38).BackColor = Map.Slot38
            lblMap(39).BackColor = Map.Slot39
            lblMap(40).BackColor = Map.Slot40
            lblMap(41).BackColor = Map.Slot41
            lblMap(42).BackColor = Map.Slot42
            lblMap(43).BackColor = Map.Slot43
            lblMap(44).BackColor = Map.Slot44
            lblMap(45).BackColor = Map.Slot45
            lblMap(46).BackColor = Map.Slot46
            lblMap(47).BackColor = Map.Slot47
            lblMap(48).BackColor = Map.Slot48
            lblMap(49).BackColor = Map.Slot49
            lblMap(50).BackColor = Map.Slot50
            lblMap(51).BackColor = Map.Slot51
            lblMap(52).BackColor = Map.Slot52
            lblMap(53).BackColor = Map.Slot53
            lblMap(54).BackColor = Map.Slot54
            lblMap(55).BackColor = Map.Slot55
            lblMap(56).BackColor = Map.Slot56
            lblMap(57).BackColor = Map.Slot57
            lblMap(58).BackColor = Map.Slot58
            lblMap(59).BackColor = Map.Slot59
            lblMap(60).BackColor = Map.Slot60
            lblMap(61).BackColor = Map.Slot61
            lblMap(62).BackColor = Map.Slot62
            lblMap(63).BackColor = Map.Slot63
            lblMap(64).BackColor = Map.Slot64
            lblMap(65).BackColor = Map.Slot65
            lblMap(66).BackColor = Map.Slot66
            lblMap(67).BackColor = Map.Slot67
            lblMap(68).BackColor = Map.Slot68
            lblMap(69).BackColor = Map.Slot69
            lblMap(70).BackColor = Map.Slot70
            lblMap(71).BackColor = Map.Slot71
            lblMap(72).BackColor = Map.Slot72
            lblMap(73).BackColor = Map.Slot73
            lblMap(74).BackColor = Map.Slot74
            lblMap(75).BackColor = Map.Slot75
            lblMap(76).BackColor = Map.Slot76
            lblMap(77).BackColor = Map.Slot77
            lblMap(78).BackColor = Map.Slot78
            lblMap(79).BackColor = Map.Slot79
            lblMap(80).BackColor = Map.Slot80
            lblMap(81).BackColor = Map.Slot81
            lblMap(82).BackColor = Map.Slot82
            lblMap(83).BackColor = Map.Slot83
            lblMap(84).BackColor = Map.Slot84
            lblMap(85).BackColor = Map.Slot85
            lblMap(86).BackColor = Map.Slot86
            lblMap(87).BackColor = Map.Slot87
            lblMap(88).BackColor = Map.Slot88
            lblMap(89).BackColor = Map.Slot89
            lblMap(90).BackColor = Map.Slot90
            lblMap(91).BackColor = Map.Slot91
            lblMap(92).BackColor = Map.Slot92
            lblMap(93).BackColor = Map.Slot93
            lblMap(94).BackColor = Map.Slot94
            lblMap(95).BackColor = Map.Slot95
            lblMap(96).BackColor = Map.Slot96
            lblMap(97).BackColor = Map.Slot97
            lblMap(98).BackColor = Map.Slot98
            lblMap(99).BackColor = Map.Slot99
            lblMap(100).BackColor = Map.Slot100
            lblMap(101).BackColor = Map.Slot101
            lblMap(102).BackColor = Map.Slot102
            lblMap(103).BackColor = Map.Slot103
            lblMap(104).BackColor = Map.Slot104
            lblMap(105).BackColor = Map.Slot105
            lblMap(106).BackColor = Map.Slot106
            lblMap(107).BackColor = Map.Slot107
            lblMap(108).BackColor = Map.Slot108
            lblMap(109).BackColor = Map.Slot109
            lblMap(110).BackColor = Map.Slot110
            lblMap(111).BackColor = Map.Slot111
            lblMap(112).BackColor = Map.Slot112
            lblMap(113).BackColor = Map.Slot113
            lblMap(114).BackColor = Map.Slot114
            lblMap(115).BackColor = Map.Slot115
            lblMap(116).BackColor = Map.Slot116
            lblMap(117).BackColor = Map.Slot117
            lblMap(118).BackColor = Map.Slot118
            lblMap(119).BackColor = Map.Slot119
            lblMap(120).BackColor = Map.Slot120
            lblMap(121).BackColor = Map.Slot121
            lblMap(122).BackColor = Map.Slot122
            lblMap(123).BackColor = Map.Slot123
            lblMap(124).BackColor = Map.Slot124
            lblMap(125).BackColor = Map.Slot125
            lblMap(126).BackColor = Map.Slot126
            lblMap(127).BackColor = Map.Slot127
            lblMap(128).BackColor = Map.Slot128
            lblMap(129).BackColor = Map.Slot129
            lblMap(130).BackColor = Map.Slot130
            lblMap(131).BackColor = Map.Slot131
            lblMap(132).BackColor = Map.Slot132
            lblMap(133).BackColor = Map.Slot133
            lblMap(134).BackColor = Map.Slot134
            lblMap(135).BackColor = Map.Slot135
            lblMap(136).BackColor = Map.Slot136
            lblMap(137).BackColor = Map.Slot137
            lblMap(138).BackColor = Map.Slot138
            lblMap(139).BackColor = Map.Slot139
            lblMap(140).BackColor = Map.Slot140
            lblMap(141).BackColor = Map.Slot141
            lblMap(142).BackColor = Map.Slot142
            lblMap(143).BackColor = Map.Slot143
            lblMap(144).BackColor = Map.Slot144
            lblMap(145).BackColor = Map.Slot145
            lblMap(146).BackColor = Map.Slot146
            lblMap(147).BackColor = Map.Slot147
            lblMap(148).BackColor = Map.Slot148
            lblMap(149).BackColor = Map.Slot149
            lblMap(150).BackColor = Map.Slot150
            lblMap(151).BackColor = Map.Slot151
            lblMap(152).BackColor = Map.Slot152
            lblMap(153).BackColor = Map.Slot153
            lblMap(154).BackColor = Map.Slot154
            lblMap(155).BackColor = Map.Slot155
            lblMap(156).BackColor = Map.Slot156
            lblMap(157).BackColor = Map.Slot157
            lblMap(158).BackColor = Map.Slot158
            lblMap(159).BackColor = Map.Slot159
            lblMap(160).BackColor = Map.Slot160
            lblMap(161).BackColor = Map.Slot161
            lblMap(162).BackColor = Map.Slot162
            lblMap(163).BackColor = Map.Slot163
            lblMap(164).BackColor = Map.Slot164
            
            frmMapOptions.txtAuthor.Text = Map.Author
            frmMapOptions.txtPassword.Text = Map.Password
            frmMapOptions.txtTitle.Text = Map.Title
            frmMapOptions.cboDifficulty.ListIndex = Map.Difficulty - 1
            If Map.Title <> "" Then
                frmMain.Caption = App.Title & " [" & Map.Title & "]"
            Else
                frmMain.Caption = App.Title & " [" & dlgCD.FileName & "]"
            End If
        End If
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
    With dlgCD
        .Filter = "Map Files (*.map)|*.map|All Files (*.*)|*.*"
        .ShowSave
    End With
    
    'Save the file and throw the file handle away
    If dlgCD.FileName <> "" Then
        Map.Slot0 = lblMap(0).BackColor
        Map.Slot1 = lblMap(1).BackColor
        Map.Slot2 = lblMap(2).BackColor
        Map.Slot3 = lblMap(3).BackColor
        Map.Slot4 = lblMap(4).BackColor
        Map.Slot5 = lblMap(5).BackColor
        Map.Slot6 = lblMap(6).BackColor
        Map.Slot7 = lblMap(7).BackColor
        Map.Slot8 = lblMap(8).BackColor
        Map.Slot9 = lblMap(9).BackColor
        Map.Slot10 = lblMap(10).BackColor
        Map.Slot11 = lblMap(11).BackColor
        Map.Slot12 = lblMap(12).BackColor
        Map.Slot13 = lblMap(13).BackColor
        Map.Slot14 = lblMap(14).BackColor
        Map.Slot15 = lblMap(15).BackColor
        Map.Slot16 = lblMap(16).BackColor
        Map.Slot17 = lblMap(17).BackColor
        Map.Slot18 = lblMap(18).BackColor
        Map.Slot19 = lblMap(19).BackColor
        Map.Slot20 = lblMap(20).BackColor
        Map.Slot21 = lblMap(21).BackColor
        Map.Slot22 = lblMap(22).BackColor
        Map.Slot23 = lblMap(23).BackColor
        Map.Slot24 = lblMap(24).BackColor
        Map.Slot25 = lblMap(25).BackColor
        Map.Slot26 = lblMap(26).BackColor
        Map.Slot27 = lblMap(27).BackColor
        Map.Slot28 = lblMap(28).BackColor
        Map.Slot29 = lblMap(29).BackColor
        Map.Slot30 = lblMap(30).BackColor
        Map.Slot31 = lblMap(31).BackColor
        Map.Slot32 = lblMap(32).BackColor
        Map.Slot33 = lblMap(33).BackColor
        Map.Slot34 = lblMap(34).BackColor
        Map.Slot35 = lblMap(35).BackColor
        Map.Slot36 = lblMap(36).BackColor
        Map.Slot37 = lblMap(37).BackColor
        Map.Slot38 = lblMap(38).BackColor
        Map.Slot39 = lblMap(39).BackColor
        Map.Slot40 = lblMap(40).BackColor
        Map.Slot41 = lblMap(41).BackColor
        Map.Slot42 = lblMap(42).BackColor
        Map.Slot43 = lblMap(43).BackColor
        Map.Slot44 = lblMap(44).BackColor
        Map.Slot45 = lblMap(45).BackColor
        Map.Slot46 = lblMap(46).BackColor
        Map.Slot47 = lblMap(47).BackColor
        Map.Slot48 = lblMap(48).BackColor
        Map.Slot49 = lblMap(49).BackColor
        Map.Slot50 = lblMap(50).BackColor
        Map.Slot51 = lblMap(51).BackColor
        Map.Slot52 = lblMap(52).BackColor
        Map.Slot53 = lblMap(53).BackColor
        Map.Slot54 = lblMap(54).BackColor
        Map.Slot55 = lblMap(55).BackColor
        Map.Slot56 = lblMap(56).BackColor
        Map.Slot57 = lblMap(57).BackColor
        Map.Slot58 = lblMap(58).BackColor
        Map.Slot59 = lblMap(59).BackColor
        Map.Slot60 = lblMap(60).BackColor
        Map.Slot61 = lblMap(61).BackColor
        Map.Slot62 = lblMap(62).BackColor
        Map.Slot63 = lblMap(63).BackColor
        Map.Slot64 = lblMap(64).BackColor
        Map.Slot65 = lblMap(65).BackColor
        Map.Slot66 = lblMap(66).BackColor
        Map.Slot67 = lblMap(67).BackColor
        Map.Slot68 = lblMap(68).BackColor
        Map.Slot69 = lblMap(69).BackColor
        Map.Slot70 = lblMap(70).BackColor
        Map.Slot71 = lblMap(71).BackColor
        Map.Slot72 = lblMap(72).BackColor
        Map.Slot73 = lblMap(73).BackColor
        Map.Slot74 = lblMap(74).BackColor
        Map.Slot75 = lblMap(75).BackColor
        Map.Slot76 = lblMap(76).BackColor
        Map.Slot77 = lblMap(77).BackColor
        Map.Slot78 = lblMap(78).BackColor
        Map.Slot79 = lblMap(79).BackColor
        Map.Slot80 = lblMap(80).BackColor
        Map.Slot81 = lblMap(81).BackColor
        Map.Slot82 = lblMap(82).BackColor
        Map.Slot83 = lblMap(83).BackColor
        Map.Slot84 = lblMap(84).BackColor
        Map.Slot85 = lblMap(85).BackColor
        Map.Slot86 = lblMap(86).BackColor
        Map.Slot87 = lblMap(87).BackColor
        Map.Slot88 = lblMap(88).BackColor
        Map.Slot89 = lblMap(89).BackColor
        Map.Slot90 = lblMap(90).BackColor
        Map.Slot91 = lblMap(91).BackColor
        Map.Slot92 = lblMap(92).BackColor
        Map.Slot93 = lblMap(93).BackColor
        Map.Slot94 = lblMap(94).BackColor
        Map.Slot95 = lblMap(95).BackColor
        Map.Slot96 = lblMap(96).BackColor
        Map.Slot97 = lblMap(97).BackColor
        Map.Slot98 = lblMap(98).BackColor
        Map.Slot99 = lblMap(99).BackColor
        Map.Slot100 = lblMap(100).BackColor
        Map.Slot101 = lblMap(101).BackColor
        Map.Slot102 = lblMap(102).BackColor
        Map.Slot103 = lblMap(103).BackColor
        Map.Slot104 = lblMap(104).BackColor
        Map.Slot105 = lblMap(105).BackColor
        Map.Slot106 = lblMap(106).BackColor
        Map.Slot107 = lblMap(107).BackColor
        Map.Slot108 = lblMap(108).BackColor
        Map.Slot109 = lblMap(109).BackColor
        Map.Slot110 = lblMap(110).BackColor
        Map.Slot111 = lblMap(111).BackColor
        Map.Slot112 = lblMap(112).BackColor
        Map.Slot113 = lblMap(113).BackColor
        Map.Slot114 = lblMap(114).BackColor
        Map.Slot115 = lblMap(115).BackColor
        Map.Slot116 = lblMap(116).BackColor
        Map.Slot117 = lblMap(117).BackColor
        Map.Slot118 = lblMap(118).BackColor
        Map.Slot119 = lblMap(119).BackColor
        Map.Slot120 = lblMap(120).BackColor
        Map.Slot121 = lblMap(121).BackColor
        Map.Slot122 = lblMap(122).BackColor
        Map.Slot123 = lblMap(123).BackColor
        Map.Slot124 = lblMap(124).BackColor
        Map.Slot125 = lblMap(125).BackColor
        Map.Slot126 = lblMap(126).BackColor
        Map.Slot127 = lblMap(127).BackColor
        Map.Slot128 = lblMap(128).BackColor
        Map.Slot129 = lblMap(129).BackColor
        Map.Slot130 = lblMap(130).BackColor
        Map.Slot131 = lblMap(131).BackColor
        Map.Slot132 = lblMap(132).BackColor
        Map.Slot133 = lblMap(133).BackColor
        Map.Slot134 = lblMap(134).BackColor
        Map.Slot135 = lblMap(135).BackColor
        Map.Slot136 = lblMap(136).BackColor
        Map.Slot137 = lblMap(137).BackColor
        Map.Slot138 = lblMap(138).BackColor
        Map.Slot139 = lblMap(139).BackColor
        Map.Slot140 = lblMap(140).BackColor
        Map.Slot141 = lblMap(141).BackColor
        Map.Slot142 = lblMap(142).BackColor
        Map.Slot143 = lblMap(143).BackColor
        Map.Slot144 = lblMap(144).BackColor
        Map.Slot145 = lblMap(145).BackColor
        Map.Slot146 = lblMap(146).BackColor
        Map.Slot147 = lblMap(147).BackColor
        Map.Slot148 = lblMap(148).BackColor
        Map.Slot149 = lblMap(149).BackColor
        Map.Slot150 = lblMap(150).BackColor
        Map.Slot151 = lblMap(151).BackColor
        Map.Slot152 = lblMap(152).BackColor
        Map.Slot153 = lblMap(153).BackColor
        Map.Slot154 = lblMap(154).BackColor
        Map.Slot155 = lblMap(155).BackColor
        Map.Slot156 = lblMap(156).BackColor
        Map.Slot157 = lblMap(157).BackColor
        Map.Slot158 = lblMap(158).BackColor
        Map.Slot159 = lblMap(159).BackColor
        Map.Slot160 = lblMap(160).BackColor
        Map.Slot161 = lblMap(161).BackColor
        Map.Slot162 = lblMap(162).BackColor
        Map.Slot163 = lblMap(163).BackColor
        Map.Slot164 = lblMap(164).BackColor
        
        SaveMap dlgCD.FileName
        If Map.Title <> "" Then
            frmMain.Caption = App.Title & " [" & Map.Title & "]"
        Else
            frmMain.Caption = App.Title & " [" & dlgCD.FileName & "]"
        End If
    End If
End Sub

Private Sub mnuToolsMapOptions_Click()
    'Nothing special here, just showing a form
    frmMapOptions.Show vbModal, Me
End Sub
