Attribute VB_Name = "Declarations"
'Public vcdrents As Index
'Public dvdrents As Integer ' indicates how many cds rented
Public totcost As Integer 'total cost of rented cd
Public indicate As String 'used in editing a customer
Public lv As ListItem
Public custlog As Boolean 'customer loggin
Public conn As ADODB.Connection
Public rs As ADODB.Recordset
Public stockcount As Integer 'used in counting of rs
Public cfname As String, clname As String, cidno As String
Public idvd As Boolean, ivcd As Boolean
Public printpcount As Integer
'Public ihap As Integer
Public renti As String
Public duecost As Integer
Public dueid As String

Public Sub duedate()
Set rs = New ADODB.Recordset
rs.Open "dvd", conn, adOpenKeyset, adLockPessimistic, adCmdTable
If Not rs.BOF Then
rs.MoveFirst
End If
Do Until rs.EOF
If rs("return") <> "" Then
    'If DateValue(rs("return")) = DateValue(Now) Then
    If rs("return") = DateValue(Now) Then
    dueid = rs("rentedby")
    duecost = 10
    End If
End If
rs.MoveNext

Loop

End Sub
