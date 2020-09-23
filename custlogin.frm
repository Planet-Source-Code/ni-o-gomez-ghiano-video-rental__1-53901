VERSION 5.00
Begin VB.Form custlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Loggin"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "custlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Loggin"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   """Your No. 1 Stop For Your Video Needs"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "My Only Bidyo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   120
      Picture         =   "custlogin.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      Caption         =   "Admin Password Log-in"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "Admin Username"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
End
Attribute VB_Name = "custlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public login As Boolean
Public emp As Integer
Public ctr As Integer
Private Sub Command1_Click()
On Error GoTo plan

    If Text1.Text = "admin" And Text2.Text = "in" Then
        'ang "admin" kay username and "in" kay password
        frmmainmdi.mnuaddvcd = True
        frmmainmdi.mnuadddvd.Enabled = True
        frmmainmdi.mnuaddcustomer.Enabled = True
        frmmainmdi.Show 'i-show ang frmmain nga form
        'exit ang custlogin na form
        
        Unload Me
    ElseIf Text1.Text = "emp" And Text2.Text = "test" Then
        'ang "emp" kay username og "test" kay password
        frmmainmdi.Toolbar1.Buttons(2).Enabled = False
        frmmainmdi.Toolbar1.Buttons(6).Enabled = False
        frmmainmdi.Toolbar1.Buttons(10).Enabled = False
        frmmainmdi.mnuaddvcd = False
        frmmainmdi.mnuadddvd.Enabled = False
        frmmainmdi.mnuaddcustomer.Enabled = False
        'i-show ang frmmain nga form
        frmmainmdi.Show
        'exit ang custlogin na form
        Unload Me
    Else
        'kung sayop ang username og password
        MsgBox "Invalid Admin Password or Username"
    End If

plan:


End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
login = True
End Sub

