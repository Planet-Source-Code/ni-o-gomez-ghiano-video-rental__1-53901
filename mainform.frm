VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmmainmdi 
   BackColor       =   &H8000000C&
   Caption         =   "M.O.B. Video"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9135
   Icon            =   "mainform.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "mainform.frx":030A
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8400
      Top             =   120
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   7920
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   17462
            MinWidth        =   17462
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "3/31/04"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "9:28 AM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainform.frx":3F0CA
            Key             =   "imgadd"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainform.frx":3F51E
            Key             =   "imgedit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainform.frx":3F972
            Key             =   "imgnext"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainform.frx":3FDC6
            Key             =   "imgnewvcd"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainform.frx":4021A
            Key             =   "imgeditvcd"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainform.frx":4066E
            Key             =   "imglogin"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainform.frx":40AC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainform.frx":40F16
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainform.frx":4136A
            Key             =   "imgcustlist"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainform.frx":417BE
            Key             =   "imgpaydue"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbnewlogin"
            Object.ToolTipText     =   "New Customer Login"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbaddc"
            Object.ToolTipText     =   "Add Customer"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbeditc"
            Object.ToolTipText     =   "Edit Customer"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbclist"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbnewdvd"
            Object.ToolTipText     =   "Add New DVD"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbneditdvd"
            Object.ToolTipText     =   "Edit DVD"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbdvdview"
            Object.ToolTipText     =   "View DVD List"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbnewvcd"
            Object.ToolTipText     =   "Add New VCD"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbeditvcd"
            Object.ToolTipText     =   "Edit VCD"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbvcdview"
            Object.ToolTipText     =   "View VCD List"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbpaydue"
            Object.ToolTipText     =   "Pay Due"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbcustlogout"
            Object.ToolTipText     =   "Customer Log-Out"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuaddcustomer 
         Caption         =   "&Add Customer"
      End
      Begin VB.Menu mnueditcustomer 
         Caption         =   "&Edit Customer"
      End
      Begin VB.Menu nothing 
         Caption         =   "-"
      End
      Begin VB.Menu mnunextcustomer 
         Caption         =   "Next &Customer"
      End
      Begin VB.Menu nothin2 
         Caption         =   "-"
      End
      Begin VB.Menu menuanemp 
         Caption         =   "Add &New Employee"
      End
      Begin VB.Menu noth 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuvideo 
      Caption         =   "&Video"
      Begin VB.Menu mnuadddvd 
         Caption         =   "&Add DVD"
      End
      Begin VB.Menu mnuedutdvd 
         Caption         =   "&Edit DVD"
      End
      Begin VB.Menu nothin3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaddvcd 
         Caption         =   "Add &VCD"
      End
      Begin VB.Menu mnueditvcd 
         Caption         =   "E&dit VCD"
      End
   End
   Begin VB.Menu mnurent 
      Caption         =   "&Rent"
      Begin VB.Menu mnurentdvd 
         Caption         =   "&DVD"
      End
      Begin VB.Menu mnurentvcd 
         Caption         =   "&VCD"
      End
      Begin VB.Menu beta 
         Caption         =   "BETA MAX"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Report Generator"
      Begin VB.Menu mnudvdreport 
         Caption         =   "DVD Report"
         Begin VB.Menu mnuedvdsales 
            Caption         =   "Sales"
         End
         Begin VB.Menu mnuallreportdvd 
            Caption         =   "All List"
         End
      End
      Begin VB.Menu reprtvcd 
         Caption         =   "VCD Report"
         Begin VB.Menu vcdlistall 
            Caption         =   "All List"
         End
      End
   End
End
Attribute VB_Name = "frmmainmdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub beta_Click()
MsgBox "HOY ULOL!!!!!! old school na kaayo ka???>.." + vbNewLine + "BETAMAX hahahaah!!!!", vbCritical, "Joke3x"
End Sub

Private Sub MDIForm_Load()
DataEnvironment1.myConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=" & App.Path & "\video.mdb"
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=" & App.Path & "\video.mdb"
    conn.Open
Toolbar1.Buttons(3).Enabled = False
Set rs = New ADODB.Recordset
rs.Open "customerinfo", conn, adOpenKeyset, adLockPessimistic, adCmdTable
If Not rs.BOF Then
rs.MoveFirst
End If
Do Until rs.EOF
rs.Fields("enter") = "no"

rs.MoveNext
Loop

Call duedate




End Sub


Private Sub mnuaddcustomer_Click()
'shows the newcustomer form
newcustomer.Show
End Sub

Private Sub mnuadddvd_Click()
adddvd.Show
End Sub

Private Sub mnuaddvcd_Click()
addvcd.Show
End Sub

Private Sub mnuallreportdvd_Click()
DVDreport.Show
End Sub

Private Sub mnuedvdsales_Click()
MsgBox "Sorry this f**king program is not f**king finish yet mothaf**ker!!!"
End Sub

Private Sub mnuexit_Click()
'exits the program
'rs.Open "customerinfo", conn
'If Not rs.BOF Then
'rs.MoveFirst
'End If''

'Do Until rs.EOF
'rs.Find "enter Like 'yes'"
'rs("enter") = " "
'rs.MoveNext
'Loop

End
End Sub

Private Sub mnunextcustomer_Click()
logincust.Show
End Sub

Private Sub mnurentdvd_Click()
listdvd.Show
End Sub

Private Sub mnurentvcd_Click()
listvcd.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal sangre As MSComctlLib.Button)
'toolbar
Select Case sangre.Key
Case "tbpaydue"
Set rs = New ADODB.Recordset
rs.Open "customerinfo", conn, adOpenKeyset, adLockPessimistic, adCmdTable
Do Until rs.EOF
If rs("enter") = "yes" Then
'rs("payment") = totcost
'rs("totcost") = totcost + Val(rs("due"))
End If
rs.MoveNext
Loop
rs.Close
If custlog = False Then
MsgBox "This button is only design for customer payments", vbExclamation, "Watchout!!!"
Else
'Call payduemethod
Call billfrm.Show
End If
Case "tbnewlogin" 'new loggin is pressed
    Call logincust.Show
Case "tbeditc"
    If custlog = False Then
    For X = 1 To frmcustlist.lv1.ListItems.Count
    If frmcustlist.lv1.ListItems(X).Selected = True Then
    indicate = frmcustlist.lv1.ListItems(X).Text
    editcust.Show
    End If
    Next
    Else
MsgBox "Please Exit Customer Before Editing", , "My Only Bidyo"
End If
Case "tbclist"
    frmcustlist.Show
Case "tbnewdvd"
    Call adddvd.Show
Case "tbaddc"
    Call newcustomer.Show
Case "tbnewvcd"
    Call addvcd.Show
Case "tbdvdview"
    Call listdvd.Show
'    Call listdvd.Show
Case "tbvcdview"
    'Call listvcd form
    Call listvcd.Show
Case "tbcustlogout"
    totcost = 0
    Set rs = New ADODB.Recordset
    rs.Open "customerinfo", conn, adOpenKeyset, adLockPessimistic, adCmdTable
    If Not rs.BOF Then
    rs.MoveFirst
    End If
    
    If custlog = True Then
    frmmainmdi.Caption = "M.O.B Video"
    MsgBox "Customer " + cfname + " " + clname + " Has Been Log-Out"
    'cfname = " "
    'clname = " "
    'If Not rs.BOF Then
    'rs.MoveFirst
    'End If
    'Do Until rs.EOF
    'rs("enter") = " "
    'rs("payment") = 0
    'rs.MoveNext
    'Loop
    
    custlog = False
    End If
    
End Select

End Sub

Private Sub payduemethod()
Set rs = New ADODB.Recordset
rs.Open "customerinfo", conn, adOpenKeyset, adLockPessimistic, adCmdTable
   
    Do Until rs.EOF
    If rs("enter") = "yes" Then
    rs("payment") = Val(totcost)
    rs("totcost") = totcost + Val(rs("due"))
    End If
    rs.MoveNext
    Loop
    reportpay.Show
   
End Sub

Private Sub vcdlistall_Click()
vcdreport.Show
End Sub
