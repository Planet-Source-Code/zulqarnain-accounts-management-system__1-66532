VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frm_AddAccount 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Account....."
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   4095
      Begin VB.ComboBox cmb_acc_type 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txt_acc_title 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cmb_acc_head 
         Height          =   315
         ItemData        =   "frm_AddAccount.frx":0000
         Left            =   1440
         List            =   "frm_AddAccount.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Account Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1230
      End
      Begin VB.Label lbl_account_no 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Account No."
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Account Title:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Head:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1245
      End
   End
   Begin LVbuttons.LaVolpeButton cmd_save 
      Height          =   300
      Left            =   2160
      TabIndex        =   3
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frm_AddAccount.frx":0047
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmd_reset 
      Height          =   300
      Left            =   3240
      TabIndex        =   4
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "Ok"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frm_AddAccount.frx":0063
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "ACCOUNT DETAILS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frm_AddAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim headID As String    'Stores Account Head ID
Dim accountType As String

Private Sub cmb_acc_head_Click()

If cmb_acc_head.Text = "Assets" Then
    cmb_acc_type.Enabled = True
    cmb_acc_type.Clear
    cmb_acc_type.AddItem "Fixed Asset"
    cmb_acc_type.AddItem "Current Asset"
ElseIf cmb_acc_head.Text = "Liabilities" Then
    cmb_acc_type.Enabled = True
    cmb_acc_type.Clear
    cmb_acc_type.AddItem "Long Term Liability"
    cmb_acc_type.AddItem "Current Liability"
Else
    cmb_acc_type.Enabled = False
    cmb_acc_type.Clear
End If

mod_db.DbConnection
    selectRecord.Open "Select account_head_id from tbl_account_head where account_head_name = '" & cmb_acc_head.Text & "';", con
        If selectRecord.EOF = False Then
            headID = selectRecord.Fields(0)
        End If
    selectRecord.Close
con.Close

End Sub

Private Sub cmd_reset_Click()

Unload Me

End Sub

Private Sub cmd_save_Click()

If txt_acc_title.Text = "" Then
    MsgBox "Please enter Account Title", vbInformation, "Accounts Management System....."
    txt_acc_title.SetFocus
    Exit Sub
End If

If cmb_acc_head.ListIndex = -1 Then
    MsgBox "Please select Account Head", vbInformation, "Accounts Management System....."
    txt_acc_title.SetFocus
    Exit Sub
End If

If cmb_acc_head.Text = "Asset" Or cmb_acc_head.Text = "Liability" Then
    If cmb_acc_type.ListIndex = -1 Then
        MsgBox "Please select Account Type", vbInformation, "Accounts Management System....."
        cmb_acc_type.SetFocus
        Exit Sub
    End If
End If

mod_db.DbConnection
    selectRecord.Open "Select account_title from tbl_account_title where account_title = '" & txt_acc_title.Text & "';", con
        If selectRecord.EOF = False Then
            MsgBox "Please enter another Account Title because it already exist!", vbCritical, "Accounts Management System....."
            txt_acc_title.SetFocus
        Else
            If cmb_acc_type.Enabled = True Then
                If cmb_acc_type.ListIndex = -1 Then
                    MsgBox "Please select accoun type", vbExclamation, "Accounts Management Software....."
                    cmb_acc_type.SetFocus
                    Exit Sub
                Else
                    accountType = cmb_acc_type.Text
                End If
            Else
                accountType = ""
            End If
            Set insertRecord = New ADODB.Recordset
            insertRecord.Open "Insert into tbl_account_title values ('" & lbl_account_no.Caption & "', '" & headID & "', '" & txt_acc_title.Text & "', '" & accountType & "', '0', '0', '0');", con
            If cmb_acc_head.Text = "Assets" Or cmb_acc_head.Text = "Liabilities" Or cmb_acc_head.Text = "Capital" Then
                insertRecord.Open "Insert into tbl_balancesheet values ('" & activeAccountingPeriod & "', '" & headID & "', '" & lbl_account_no.Caption & "', '" & txt_acc_title.Text & "', '0', '0');", con
            End If
            If cmb_acc_head.Text = "Expense" Or cmb_acc_head.Text = "Revenue" Then
                insertRecord.Open "Insert into tbl_incomestmt values ('" & activeAccountingPeriod & "', '" & headID & "', '" & lbl_account_no.Caption & "', '" & txt_acc_title.Text & "', '0');", con
            End If
            MsgBox "Sub-Account has been successfully added!", vbInformation, "Accounts Management System....."
        End If
con.Close

GenerateAccountNo
txt_acc_title.Text = ""
cmb_acc_head.ListIndex = -1
cmb_acc_type.Enabled = False
cmb_acc_type.Clear
txt_acc_title.SetFocus

End Sub

Private Sub Form_Load()

Me.Top = 4000
Me.Left = 6000

Set selectRecord = New ADODB.Recordset

GenerateAccountNo

End Sub

Private Sub GenerateAccountNo()

mod_db.DbConnection
    selectRecord.Open "Select max(account_title_id) from tbl_account_title;", con
        If selectRecord.EOF = False Then
            lbl_account_no.Caption = selectRecord.Fields(0) + 1
        End If
    selectRecord.Close
con.Close

End Sub

Private Sub txt_acc_title_GotFocus()

txt_acc_title.BackColor = &H80000018

End Sub

Private Sub txt_acc_title_LostFocus()

txt_acc_title.BackColor = &HFFFFFF

End Sub
