VERSION 5.00
Begin VB.Form frmMapOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Options"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   2910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox cboDifficulty 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Difficulty:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Author:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Title:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "frmMapOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'We don't have a bin labled "Recycle" so
    'instead let's unload the form
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'The user has set a password for this map
    'and clicked the "OK" button. I was going
    'to write about my whole life here, but
    'I thought that if anyone were to use this
    'feature, they would be awful mad that it
    'didn't work, so I threw some code in here ;)
    Map.Password = txtPassword.Text
    Map.Title = txtTitle.Text
    Map.Author = txtAuthor.Text
    'Not very impressive; just another Select...End Select case
    'I could have assigned the text value to the file and read it
    'in as text, but I didn't wanna ;)
    Select Case cboDifficulty.Text
        Case "Novice"
            Map.Difficulty = 1
        Case "Easy"
            Map.Difficulty = 2
        Case "Medium"
            Map.Difficulty = 3
        Case "Hard"
            Map.Difficulty = 4
        Case "Hardest"
            Map.Difficulty = 5
    End Select
    'The user might want to know which map or file they are working on
    If Map.Title <> "" Then
        frmMain.Caption = App.Title & " [" & Map.Title & "]"
    ElseIf frmMain.dlgCD.FileName <> "" Then
        frmMain.Caption = App.Title & " [" & frmMain.dlgCD.FileName & "]"
    Else
        frmMain.Caption = App.Title & " [Untitled]"
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    'Let's set the defaults for the user
    cboDifficulty.AddItem "Novice"
    cboDifficulty.AddItem "Easy"
    cboDifficulty.AddItem "Medium"
    cboDifficulty.AddItem "Hard"
    cboDifficulty.AddItem "Hardest"
    cboDifficulty.ListIndex = 0
End Sub
