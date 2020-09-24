VERSION 5.00
Begin VB.Form frm_LockApplication 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Lock Application..."
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   FillColor       =   &H00AC8576&
   ForeColor       =   &H00AC8576&
   LinkTopic       =   "Form1"
   ScaleHeight     =   630
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Type your password and press <Enter> in your keyboard."
      Top             =   160
      Width           =   3975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   195
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   2
      FillColor       =   &H00C7C1BA&
      Height          =   615
      Left            =   20
      Shape           =   4  'Rounded Rectangle
      Top             =   20
      Width           =   5895
   End
End
Attribute VB_Name = "frm_LockApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Text1.Text = ""

Me.Top = 5000
Me.Left = 5000

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Text1.Text = "admin" Then
        Unload Me
    Else
       MsgBox "The password doesn't matches", vbCritical, "Accounts Management System....."
       Text1.Text = ""
       Text1.SetFocus
    End If
End If

Call mod_Validation.ForGeneralFields(KeyAscii, Text1)

End Sub
