VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Password"
   ClientHeight    =   1965
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1160.987
   ScaleMode       =   0  'User
   ScaleWidth      =   4126.667
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "The map you are attempting to open has been password protected. Please enter the password in the box below to open this map."
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   840
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'The user doesn't know the password so close
    'this cracker-jack password dialog
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'Check for the correct password
    If txtPassword = Map.Password Then
        frmMain.lblMap(0).BackColor = Map.Slot0
        frmMain.lblMap(1).BackColor = Map.Slot1
        frmMain.lblMap(2).BackColor = Map.Slot2
        frmMain.lblMap(3).BackColor = Map.Slot3
        frmMain.lblMap(4).BackColor = Map.Slot4
        frmMain.lblMap(5).BackColor = Map.Slot5
        frmMain.lblMap(6).BackColor = Map.Slot6
        frmMain.lblMap(7).BackColor = Map.Slot7
        frmMain.lblMap(8).BackColor = Map.Slot8
        frmMain.lblMap(9).BackColor = Map.Slot9
        frmMain.lblMap(10).BackColor = Map.Slot10
        frmMain.lblMap(11).BackColor = Map.Slot11
        frmMain.lblMap(12).BackColor = Map.Slot12
        frmMain.lblMap(13).BackColor = Map.Slot13
        frmMain.lblMap(14).BackColor = Map.Slot14
        frmMain.lblMap(15).BackColor = Map.Slot15
        frmMain.lblMap(16).BackColor = Map.Slot16
        frmMain.lblMap(17).BackColor = Map.Slot17
        frmMain.lblMap(18).BackColor = Map.Slot18
        frmMain.lblMap(19).BackColor = Map.Slot19
        frmMain.lblMap(20).BackColor = Map.Slot20
        frmMain.lblMap(21).BackColor = Map.Slot21
        frmMain.lblMap(22).BackColor = Map.Slot22
        frmMain.lblMap(23).BackColor = Map.Slot23
        frmMain.lblMap(24).BackColor = Map.Slot24
        frmMain.lblMap(25).BackColor = Map.Slot25
        frmMain.lblMap(26).BackColor = Map.Slot26
        frmMain.lblMap(27).BackColor = Map.Slot27
        frmMain.lblMap(28).BackColor = Map.Slot28
        frmMain.lblMap(29).BackColor = Map.Slot29
        frmMain.lblMap(30).BackColor = Map.Slot30
        frmMain.lblMap(31).BackColor = Map.Slot31
        frmMain.lblMap(32).BackColor = Map.Slot32
        frmMain.lblMap(33).BackColor = Map.Slot33
        frmMain.lblMap(34).BackColor = Map.Slot34
        frmMain.lblMap(35).BackColor = Map.Slot35
        frmMain.lblMap(36).BackColor = Map.Slot36
        frmMain.lblMap(37).BackColor = Map.Slot37
        frmMain.lblMap(38).BackColor = Map.Slot38
        frmMain.lblMap(39).BackColor = Map.Slot39
        frmMain.lblMap(40).BackColor = Map.Slot40
        frmMain.lblMap(41).BackColor = Map.Slot41
        frmMain.lblMap(42).BackColor = Map.Slot42
        frmMain.lblMap(43).BackColor = Map.Slot43
        frmMain.lblMap(44).BackColor = Map.Slot44
        frmMain.lblMap(45).BackColor = Map.Slot45
        frmMain.lblMap(46).BackColor = Map.Slot46
        frmMain.lblMap(47).BackColor = Map.Slot47
        frmMain.lblMap(48).BackColor = Map.Slot48
        frmMain.lblMap(49).BackColor = Map.Slot49
        frmMain.lblMap(50).BackColor = Map.Slot50
        frmMain.lblMap(51).BackColor = Map.Slot51
        frmMain.lblMap(52).BackColor = Map.Slot52
        frmMain.lblMap(53).BackColor = Map.Slot53
        frmMain.lblMap(54).BackColor = Map.Slot54
        frmMain.lblMap(55).BackColor = Map.Slot55
        frmMain.lblMap(56).BackColor = Map.Slot56
        frmMain.lblMap(57).BackColor = Map.Slot57
        frmMain.lblMap(58).BackColor = Map.Slot58
        frmMain.lblMap(59).BackColor = Map.Slot59
        frmMain.lblMap(60).BackColor = Map.Slot60
        frmMain.lblMap(61).BackColor = Map.Slot61
        frmMain.lblMap(62).BackColor = Map.Slot62
        frmMain.lblMap(63).BackColor = Map.Slot63
        frmMain.lblMap(64).BackColor = Map.Slot64
        frmMain.lblMap(65).BackColor = Map.Slot65
        frmMain.lblMap(66).BackColor = Map.Slot66
        frmMain.lblMap(67).BackColor = Map.Slot67
        frmMain.lblMap(68).BackColor = Map.Slot68
        frmMain.lblMap(69).BackColor = Map.Slot69
        frmMain.lblMap(70).BackColor = Map.Slot70
        frmMain.lblMap(71).BackColor = Map.Slot71
        frmMain.lblMap(72).BackColor = Map.Slot72
        frmMain.lblMap(73).BackColor = Map.Slot73
        frmMain.lblMap(74).BackColor = Map.Slot74
        frmMain.lblMap(75).BackColor = Map.Slot75
        frmMain.lblMap(76).BackColor = Map.Slot76
        frmMain.lblMap(77).BackColor = Map.Slot77
        frmMain.lblMap(78).BackColor = Map.Slot78
        frmMain.lblMap(79).BackColor = Map.Slot79
        frmMain.lblMap(80).BackColor = Map.Slot80
        frmMain.lblMap(81).BackColor = Map.Slot81
        frmMain.lblMap(82).BackColor = Map.Slot82
        frmMain.lblMap(83).BackColor = Map.Slot83
        frmMain.lblMap(84).BackColor = Map.Slot84
        frmMain.lblMap(85).BackColor = Map.Slot85
        frmMain.lblMap(86).BackColor = Map.Slot86
        frmMain.lblMap(87).BackColor = Map.Slot87
        frmMain.lblMap(88).BackColor = Map.Slot88
        frmMain.lblMap(89).BackColor = Map.Slot89
        frmMain.lblMap(90).BackColor = Map.Slot90
        frmMain.lblMap(91).BackColor = Map.Slot91
        frmMain.lblMap(92).BackColor = Map.Slot92
        frmMain.lblMap(93).BackColor = Map.Slot93
        frmMain.lblMap(94).BackColor = Map.Slot94
        frmMain.lblMap(95).BackColor = Map.Slot95
        frmMain.lblMap(96).BackColor = Map.Slot96
        frmMain.lblMap(97).BackColor = Map.Slot97
        frmMain.lblMap(98).BackColor = Map.Slot98
        frmMain.lblMap(99).BackColor = Map.Slot99
        frmMain.lblMap(100).BackColor = Map.Slot100
        frmMain.lblMap(101).BackColor = Map.Slot101
        frmMain.lblMap(102).BackColor = Map.Slot102
        frmMain.lblMap(103).BackColor = Map.Slot103
        frmMain.lblMap(104).BackColor = Map.Slot104
        frmMain.lblMap(105).BackColor = Map.Slot105
        frmMain.lblMap(106).BackColor = Map.Slot106
        frmMain.lblMap(107).BackColor = Map.Slot107
        frmMain.lblMap(108).BackColor = Map.Slot108
        frmMain.lblMap(109).BackColor = Map.Slot109
        frmMain.lblMap(110).BackColor = Map.Slot110
        frmMain.lblMap(111).BackColor = Map.Slot111
        frmMain.lblMap(112).BackColor = Map.Slot112
        frmMain.lblMap(113).BackColor = Map.Slot113
        frmMain.lblMap(114).BackColor = Map.Slot114
        frmMain.lblMap(115).BackColor = Map.Slot115
        frmMain.lblMap(116).BackColor = Map.Slot116
        frmMain.lblMap(117).BackColor = Map.Slot117
        frmMain.lblMap(118).BackColor = Map.Slot118
        frmMain.lblMap(119).BackColor = Map.Slot119
        frmMain.lblMap(120).BackColor = Map.Slot120
        frmMain.lblMap(121).BackColor = Map.Slot121
        frmMain.lblMap(122).BackColor = Map.Slot122
        frmMain.lblMap(123).BackColor = Map.Slot123
        frmMain.lblMap(124).BackColor = Map.Slot124
        frmMain.lblMap(125).BackColor = Map.Slot125
        frmMain.lblMap(126).BackColor = Map.Slot126
        frmMain.lblMap(127).BackColor = Map.Slot127
        frmMain.lblMap(128).BackColor = Map.Slot128
        frmMain.lblMap(129).BackColor = Map.Slot129
        frmMain.lblMap(130).BackColor = Map.Slot130
        frmMain.lblMap(131).BackColor = Map.Slot131
        frmMain.lblMap(132).BackColor = Map.Slot132
        frmMain.lblMap(133).BackColor = Map.Slot133
        frmMain.lblMap(134).BackColor = Map.Slot134
        frmMain.lblMap(135).BackColor = Map.Slot135
        frmMain.lblMap(136).BackColor = Map.Slot136
        frmMain.lblMap(137).BackColor = Map.Slot137
        frmMain.lblMap(138).BackColor = Map.Slot138
        frmMain.lblMap(139).BackColor = Map.Slot139
        frmMain.lblMap(140).BackColor = Map.Slot140
        frmMain.lblMap(141).BackColor = Map.Slot141
        frmMain.lblMap(142).BackColor = Map.Slot142
        frmMain.lblMap(143).BackColor = Map.Slot143
        frmMain.lblMap(144).BackColor = Map.Slot144
        frmMain.lblMap(145).BackColor = Map.Slot145
        frmMain.lblMap(146).BackColor = Map.Slot146
        frmMain.lblMap(147).BackColor = Map.Slot147
        frmMain.lblMap(148).BackColor = Map.Slot148
        frmMain.lblMap(149).BackColor = Map.Slot149
        frmMain.lblMap(150).BackColor = Map.Slot150
        frmMain.lblMap(151).BackColor = Map.Slot151
        frmMain.lblMap(152).BackColor = Map.Slot152
        frmMain.lblMap(153).BackColor = Map.Slot153
        frmMain.lblMap(154).BackColor = Map.Slot154
        frmMain.lblMap(155).BackColor = Map.Slot155
        frmMain.lblMap(156).BackColor = Map.Slot156
        frmMain.lblMap(157).BackColor = Map.Slot157
        frmMain.lblMap(158).BackColor = Map.Slot158
        frmMain.lblMap(159).BackColor = Map.Slot159
        frmMain.lblMap(160).BackColor = Map.Slot160
        frmMain.lblMap(161).BackColor = Map.Slot161
        frmMain.lblMap(162).BackColor = Map.Slot162
        frmMain.lblMap(163).BackColor = Map.Slot163
        frmMain.lblMap(164).BackColor = Map.Slot164
            
        frmMapOptions.txtAuthor.Text = Map.Author
        frmMapOptions.txtPassword.Text = Map.Password
        frmMapOptions.txtTitle.Text = Map.Title
        frmMapOptions.cboDifficulty.ListIndex = Map.Difficulty - 1
        If Map.Title <> "" Then
            frmMain.Caption = App.Title & " [" & Map.Title & "]"
        Else
            frmMain.Caption = App.Title & " [" & frmMain.dlgCD.FileName & "]"
        End If
        Unload Me
    Else
        MsgBox "Invalid Password, try again!", vbOKOnly & vbInformation, "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub
