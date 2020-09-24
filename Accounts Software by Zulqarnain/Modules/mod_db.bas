Attribute VB_Name = "mod_db"
Public con As ADODB.Connection

Public Sub DbConnection()

Set con = New ADODB.Connection
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Account.mdb;Persist Security Info=False;OLEDB:Database Password=umairzulqiok"
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Account.mdb;Jet " & "OLEDB:Database Password=zulqiscitsarani"

End Sub

