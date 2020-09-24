Attribute VB_Name = "mod_global"
Public selectRecord As ADODB.Recordset
Public insertRecord As ADODB.Recordset
Public updateRecord As ADODB.Recordset
Public deleteRecord As ADODB.Recordset

Public nd As Node
Public lv As ListItem

Public activeAccountingPeriod As String

Public transactionNo As Integer

Public accountNo As Variant

'*****For Accounting Period*****'
Public periodFrom As String
Public periodTo As String
