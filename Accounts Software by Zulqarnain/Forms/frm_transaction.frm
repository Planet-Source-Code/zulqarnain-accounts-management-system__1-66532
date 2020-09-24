VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_transaction 
   Caption         =   "Accounts Management System.....[University Version]-----[Accounting Period:1st January 2006-31st December 2006]"
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_transaction.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   10200
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   10080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":1708A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":21024
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":2AFBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":34F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":3EEF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":3F7CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":3FC1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":408F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":411D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":4122E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":4128D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":42167
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":42525
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":42935
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":42D83
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_transaction.frx":42E45
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2775
      Left            =   13080
      TabIndex        =   32
      Top             =   7560
      Width           =   2055
      Begin LVbuttons.LaVolpeButton cmd_logoff 
         Height          =   300
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "Log-&Off"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frm_transaction.frx":42FEB
         ALIGN           =   0
         IMGLST          =   "ImageList1"
         IMGICON         =   "1"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   0   'False
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmd_exit 
         Height          =   300
         Left            =   120
         TabIndex        =   34
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "E&xit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frm_transaction.frx":43007
         ALIGN           =   0
         IMGLST          =   "ImageList1"
         IMGICON         =   "4"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   0   'False
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   900
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1588
         BTYPE           =   3
         TX              =   "Add New Sub-Account"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frm_transaction.frx":43023
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "1"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   0   'False
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmd_calc 
         Height          =   300
         Left            =   120
         TabIndex        =   58
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "Calculator"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frm_transaction.frx":4303F
         ALIGN           =   0
         IMGLST          =   "ImageList1"
         IMGICON         =   "11"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   0   'False
         BSTYLE          =   0
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   120
         Picture         =   "frm_transaction.frx":4305B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   1920
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "TOOL BOX:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   180
         Left            =   480
         TabIndex        =   35
         Top             =   0
         Width           =   885
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   10200
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   10740
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16113
            Text            =   "Developed By: Zulqarnain Sarani [SCIT-012]"
            TextSave        =   "Developed By: Zulqarnain Sarani [SCIT-012]"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "9/12/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   2
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Assets"
      TabPicture(0)   =   "frm_transaction.frx":43D25
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_asset"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tv_asset"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Liabilities"
      TabPicture(1)   =   "frm_transaction.frx":43D41
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tv_liability"
      Tab(1).Control(1)=   "Label17"
      Tab(1).Control(2)=   "lbl_liability"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Capital"
      TabPicture(2)   =   "frm_transaction.frx":43D5D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl_capital"
      Tab(2).Control(1)=   "Label19"
      Tab(2).Control(2)=   "tv_capital"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Revenue"
      TabPicture(3)   =   "frm_transaction.frx":43D79
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tv_revenue"
      Tab(3).Control(1)=   "Label21"
      Tab(3).Control(2)=   "lbl_revenue"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Expense"
      TabPicture(4)   =   "frm_transaction.frx":43D95
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lbl_expense"
      Tab(4).Control(1)=   "Label23"
      Tab(4).Control(2)=   "tv_expense"
      Tab(4).ControlCount=   3
      Begin MSComctlLib.TreeView tv_asset 
         Height          =   3735
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   6588
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tv_liability 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   8
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   6588
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tv_capital 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   9
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   6588
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tv_revenue 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   10
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   6588
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tv_expense 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   11
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   6588
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Total(Rs.):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73800
         TabIndex        =   55
         Top             =   4440
         Width           =   1020
      End
      Begin VB.Label lbl_expense 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -72720
         TabIndex        =   54
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Total(Rs.):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73800
         TabIndex        =   53
         Top             =   4440
         Width           =   1020
      End
      Begin VB.Label lbl_revenue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -72720
         TabIndex        =   52
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Total(Rs.):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73800
         TabIndex        =   51
         Top             =   4440
         Width           =   1020
      End
      Begin VB.Label lbl_capital 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -72720
         TabIndex        =   50
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Total(Rs.):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73800
         TabIndex        =   49
         Top             =   4440
         Width           =   1020
      End
      Begin VB.Label lbl_liability 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -72720
         TabIndex        =   48
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label lbl_asset 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2280
         TabIndex        =   37
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total(Rs.):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   36
         Top             =   4440
         Width           =   1020
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "New Transaction:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   4800
      TabIndex        =   13
      Top             =   120
      Width           =   10335
      Begin AMS.ucProgressBar pb1 
         Height          =   255
         Left            =   2400
         TabIndex        =   57
         Top             =   6840
         Width           =   4695
         _extentx        =   8281
         _extenty        =   450
         font            =   "frm_transaction.frx":43DB1
         brushstyle      =   0
         color           =   6956042
         color2          =   12640511
         scrolling       =   1
         doublevalonmetal=   -1  'True
      End
      Begin VB.TextBox txt_cr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8160
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txt_dr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   8160
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cmb_account 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txt_exp 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1440
         Width           =   8415
      End
      Begin MSComCtl2.DTPicker dtp_trdate 
         Height          =   300
         Left            =   1680
         TabIndex        =   14
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   20250627
         CurrentDate     =   38966
      End
      Begin MSComctlLib.ListView lv_tr 
         Height          =   3855
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Account No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Account Title"
            Object.Width           =   3545
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Explanation"
            Object.Width           =   5380
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Debit (Rs.)"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Credit (Rs.)"
            Object.Width           =   2647
         EndProperty
      End
      Begin LVbuttons.LaVolpeButton cmd_add 
         Height          =   345
         Left            =   6600
         TabIndex        =   5
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         BTYPE           =   4
         TX              =   "Add to List"
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
         MICON           =   "frm_transaction.frx":43DD9
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "10"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmd_remove 
         Height          =   345
         Left            =   8160
         TabIndex        =   6
         Top             =   2400
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BTYPE           =   4
         TX              =   "Remove from List"
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
         MICON           =   "frm_transaction.frx":43DF5
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "9"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmd_save 
         Height          =   300
         Left            =   240
         TabIndex        =   30
         Top             =   6840
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
         MICON           =   "frm_transaction.frx":43E11
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
         Left            =   1320
         TabIndex        =   39
         Top             =   6840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "Reset"
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
         MICON           =   "frm_transaction.frx":43E2D
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
      Begin MSComCtl2.DTPicker dtp_trtime 
         Height          =   300
         Left            =   5640
         TabIndex        =   46
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20250626
         CurrentDate     =   38966
      End
      Begin VB.Label lbl_accno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   5640
         TabIndex        =   56
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Time:"
         Height          =   195
         Left            =   4080
         TabIndex        =   47
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Account No."
         Height          =   195
         Left            =   4560
         TabIndex        =   45
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Credit:"
         Height          =   195
         Left            =   7440
         TabIndex        =   43
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Debit:"
         Height          =   195
         Left            =   7560
         TabIndex        =   42
         Top             =   720
         Width           =   525
      End
      Begin VB.Label lbl_cr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8760
         TabIndex        =   41
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Label lbl_dr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   7200
         TabIndex        =   40
         Top             =   6840
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Date:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Transaction No:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label lbl_transaction_no 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
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
         Left            =   1680
         TabIndex        =   17
         Top             =   360
         Width           =   8415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Select Account:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Explanation:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Accounting Reports:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   21
      Top             =   5160
      Width           =   4575
      Begin MSComCtl2.DTPicker dtp_from 
         Height          =   300
         Left            =   720
         TabIndex        =   22
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   20250627
         CurrentDate     =   38969
      End
      Begin MSComCtl2.DTPicker dtp_to 
         Height          =   300
         Left            =   2760
         TabIndex        =   23
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   20250627
         CurrentDate     =   38967
      End
      Begin LVbuttons.LaVolpeButton cmd_bs 
         Height          =   345
         Left            =   2400
         TabIndex        =   26
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         BTYPE           =   1
         TX              =   "Balance Sheet"
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
         MICON           =   "frm_transaction.frx":43E49
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
      Begin LVbuttons.LaVolpeButton cmd_tb 
         Height          =   345
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
         BTYPE           =   1
         TX              =   "Trial Balance"
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
         MICON           =   "frm_transaction.frx":43E65
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
      Begin LVbuttons.LaVolpeButton cmd_is 
         Height          =   345
         Left            =   2400
         TabIndex        =   28
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         BTYPE           =   1
         TX              =   "Income Statement"
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
         MICON           =   "frm_transaction.frx":43E81
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
      Begin LVbuttons.LaVolpeButton cmd_gj 
         Height          =   345
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
         BTYPE           =   1
         TX              =   "General Journal"
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
         MICON           =   "frm_transaction.frx":43E9D
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
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Select the reporting period"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "To:"
         Height          =   195
         Left            =   2400
         TabIndex        =   25
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "From:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   510
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Transactions:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   31
      Top             =   7560
      Width           =   12855
      Begin LVbuttons.LaVolpeButton cmd_sort 
         Height          =   285
         Left            =   4200
         TabIndex        =   65
         ToolTipText     =   "Sort"
         Top             =   2355
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   4
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   12648384
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frm_transaction.frx":43EB9
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "13"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.ComboBox cmb_ad 
         Height          =   315
         ItemData        =   "frm_transaction.frx":43ED5
         Left            =   3240
         List            =   "frm_transaction.frx":43EDF
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   2360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmb_sort 
         Height          =   315
         ItemData        =   "frm_transaction.frx":43EEE
         Left            =   960
         List            =   "frm_transaction.frx":43F01
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   2360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComctlLib.ListView lv_tr_list 
         Height          =   1935
         Left            =   240
         TabIndex        =   59
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tr. No."
            Object.Width           =   1412
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Account No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Account Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Tr. Date"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Explanation"
            Object.Width           =   6792
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Debit"
            Object.Width           =   2648
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Credit"
            Object.Width           =   2648
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sort By:"
         Height          =   195
         Left            =   240
         TabIndex        =   62
         Top             =   2360
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl_tdr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   225
         Left            =   9240
         TabIndex        =   61
         Top             =   2400
         Width           =   1530
      End
      Begin VB.Label lbl_tcr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   10800
         TabIndex        =   60
         Top             =   2400
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frm_transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dr As Double    'Storing Debit
Dim cr As Double    'Storing Credit
Dim selectedItemToRemove As Integer     'For Removing an item from lv_tr
Dim transactionDate As String

Private Sub cmb_account_Click()

mod_db.DbConnection
    selectRecord.Open "Select account_title_id from tbl_account_title where account_title = '" & cmb_account.Text & "';", con
        If selectRecord.EOF = False Then
           lbl_accno = selectRecord.Fields(0)
        End If
    selectRecord.Close
con.Close

End Sub

Private Sub cmd_add_Click()

'*****Validating*****'
If cmb_account.ListIndex = -1 Then
    MsgBox "Please select the appropriate account", vbExclamation, "Software Management System....."
    cmb_account.SetFocus
    Exit Sub
End If

If txt_dr.Text = 0 And txt_cr.Text = 0 Then
    MsgBox "Please enter valid values of Debit and Credit", vbExclamation, "Accounts Management System....."
    txt_dr.SetFocus
    Exit Sub
End If

If txt_dr.Text <> 0 And txt_cr.Text <> 0 Then
    MsgBox "Please enter valid values of Debit and Credit", vbExclamation, "Accounts Management System....."
    txt_dr.SetFocus
    Exit Sub
End If

If txt_exp.Text = "" Then
    txt_exp.Text = "-"
End If
'*******************'

Dim c As String
c = lv_tr.ListItems.Count
c = c + 1

Set lv = lv_tr.ListItems.Add(, , c & ".")
lv.SubItems(1) = lbl_accno.Caption
lv.SubItems(2) = cmb_account.Text
lv.SubItems(3) = txt_exp.Text
lv.SubItems(4) = txt_dr.Text
lv.SubItems(5) = txt_cr.Text

dr = lbl_dr.Caption
lbl_dr.Caption = FormatNumber(dr + txt_dr.Text)

cr = lbl_cr.Caption
lbl_cr.Caption = FormatNumber(cr + txt_cr.Text)

For j = 1 To lv_tr.ListItems.Count
    lv_tr.ListItems(j).Text = j & "."
Next j

End Sub

Private Sub cmd_bs_Click()

rpt_balancesheet.Show vbModal

End Sub

Private Sub cmd_calc_Click()

Shell ("calc"), vbMinimizedFocus
Exit Sub

End Sub

Private Sub cmd_exit_Click()

ques = MsgBox("Do you want to exit the Application", vbQuestion + vbYesNo, "Accounts Management System.....")
If ques = vbYes Then
    End
Else
    Cancel = 0
End If

End Sub

Private Sub cmd_gj_Click()

rpt_generaljournal.Show vbModal

End Sub

Private Sub cmd_is_Click()

rpt_incomestmt.Show vbModal

End Sub

Private Sub cmd_remove_Click()

'*****Removing selected transaction from lv_tr*****'
If lv_tr.ListItems.Count >= 1 Then
    lbl_dr.Caption = FormatNumber(lbl_dr.Caption - lv_tr.SelectedItem.SubItems(4))
    lbl_cr.Caption = FormatNumber(lbl_cr.Caption - lv_tr.SelectedItem.SubItems(5))
    
    lv_tr.ListItems.Remove (selectedItemToRemove)
    
    For j = 1 To lv_tr.ListItems.Count
        lv_tr.ListItems(j).Text = I & "."
    Next j
End If
'**************************************************'

End Sub

Private Sub cmd_reset_Click()

Reset

End Sub

Private Sub cmd_save_Click()

Set updateRecord = New ADODB.Recordset

dr = lbl_dr.Caption
cr = lbl_cr.Caption

If dr <> cr Then
    MsgBox "Transactions are invalid. i.e Debit is not equal to Credit" & vbCrLf & "Please validate the transaction", vbExclamation, "Accounts Management System....."
    Exit Sub
End If

Dim drcr As Double
Dim cr1 As Double

mod_db.DbConnection
    Set insertRecord = New ADODB.Recordset
    transactionDate = dtp_trdate.Day & "-" & MonthName(dtp_trdate.Month, True) & "-" & dtp_trdate.Year
    insertRecord.Open "Insert into tbl_acc_period_transaction values ('" & activeAccountingPeriod & "', '" & lbl_transaction_no.Caption & "', '" & transactionDate & "');", con
    For j = 1 To lv_tr.ListItems.Count
        insertRecord.Open "Insert into tbl_transaction values ('" & lbl_transaction_no.Caption & "', '" & lv_tr.ListItems(j).SubItems(1) & "', '" & transactionDate & "', '" & lv_tr.ListItems(j).SubItems(3) & "', '" & lv_tr.ListItems(j).SubItems(4) & "', '" & lv_tr.ListItems(j).SubItems(5) & "');", con
        '****Updating Balance Sheet Table*****'
        selectRecord.Open "Select account_head_id from tbl_account_title where account_title_id = " & lv_tr.ListItems(j).SubItems(1) & ";", con
            If selectRecord.EOF = False Then
                hid = selectRecord.Fields(0)
            End If
        selectRecord.Close
        If hid = "1" Or hid = "2" Or hid = "3" Then
            selectRecord.Open "Select dr, cr from tbl_balancesheet where head_no = " & hid & " and account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
                If selectRecord.EOF = False Then
                    dr1 = selectRecord.Fields(0)
                    cr1 = selectRecord.Fields(1)
                End If
            selectRecord.Close
        End If
        If hid = "4" Or hid = "5" Then
            selectRecord.Open "Select drcr from tbl_incomestmt where head_no = " & hid & " and account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
                If selectRecord.EOF = False Then
                    dr1 = selectRecord.Fields(0)
                End If
            selectRecord.Close
        End If
        
        '*****Update Asset Account*****'
        If hid = "1" And lv_tr.ListItems(j).SubItems(4) <> "0.00" Then
            updateRecord.Open "Update tbl_balancesheet set dr = '" & dr1 + lv_tr.ListItems(j).SubItems(4) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
        End If
        '*****Update Asset Account*****'
        If hid = "1" And lv_tr.ListItems(j).SubItems(5) <> "0.00" Then
            updateRecord.Open "Update tbl_balancesheet set dr = '" & dr1 - lv_tr.ListItems(j).SubItems(5) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
        End If
        
        '*****Update Liability Account*****'
        If hid = "2" And lv_tr.ListItems(j).SubItems(5) <> "0.00" Then
            'updateRecord.Open "Update tbl_balancesheet set drcr = '" & drcr + lv_tr.ListItems(j).SubItems(5) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
            updateRecord.Open "Update tbl_balancesheet set cr = '" & cr1 + lv_tr.ListItems(j).SubItems(5) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
        End If
        '*****Update Liability Account*****'
        If hid = "2" And lv_tr.ListItems(j).SubItems(4) <> "0.00" Then
            'updateRecord.Open "Update tbl_balancesheet set drcr = '" & drcr - lv_tr.ListItems(j).SubItems(4) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
            updateRecord.Open "Update tbl_balancesheet set cr = '" & cr1 - lv_tr.ListItems(j).SubItems(4) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
        End If
        
        '*****Update Capital Account*****'
        If hid = "3" And lv_tr.ListItems(j).SubItems(5) <> "0.00" Then
            'updateRecord.Open "Update tbl_balancesheet set drcr = '" & drcr + lv_tr.ListItems(j).SubItems(5) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
            updateRecord.Open "Update tbl_balancesheet set cr = '" & cr1 + lv_tr.ListItems(j).SubItems(5) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
        End If
        '*****Update Capital Account*****'
        If hid = "3" And lv_tr.ListItems(j).SubItems(4) <> "0.00" Then
            'updateRecord.Open "Update tbl_balancesheet set drcr = '" & drcr - lv_tr.ListItems(j).SubItems(4) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
            updateRecord.Open "Update tbl_balancesheet set cr = '" & cr1 - lv_tr.ListItems(j).SubItems(4) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
        End If
        
        '*****Update Revenue Account*****'
        If hid = "4" And lv_tr.ListItems(j).SubItems(5) <> "0.00" Then
            updateRecord.Open "Update tbl_incomestmt set drcr = '" & dr1 + lv_tr.ListItems(j).SubItems(5) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
        End If
         '*****Update Revenue Account*****'
        If hid = "4" And lv_tr.ListItems(j).SubItems(4) <> "0.00" Then
            updateRecord.Open "Update tbl_incomestmt set drcr = '" & dr1 - lv_tr.ListItems(j).SubItems(5) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
        End If
        
        '*****Update Expense Account*****'
        If hid = "5" And lv_tr.ListItems(j).SubItems(4) <> "0.00" Then
            updateRecord.Open "Update tbl_incomestmt set drcr = '" & dr1 + lv_tr.ListItems(j).SubItems(4) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
        End If
        '*****Update Expense Account*****'
        If hid = "5" And lv_tr.ListItems(j).SubItems(5) <> "0.00" Then
            updateRecord.Open "Update tbl_incomestmt set drcr = '" & dr1 - lv_tr.ListItems(j).SubItems(4) & "' where account_no = " & lv_tr.ListItems(j).SubItems(1) & ";", con
        End If
        
    Next j
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Timer2.Enabled = True
con.Close

End Sub


Private Sub cmd_sort_Click()

'mod_db.DbConnection
    'Dim ob As String
    'Dim f As String
    'If cmb_sort.Text = "Account No." Then
        'f = "account_title_id"
    'End If
    'If cmb_sort.Text = "Credit" Then
        'f = "cr"
    'End If
    'If cmb_sort.Text = "Debit" Then
        'f = "dr"
    'End If
    'If cmb_sort.Text = "Tr. Date" Then
        'f = "transaction_date"
    'End If
    'If cmb_sort.Text = "Transaction No" Then
        'f = "transaction_id"
    'End If
    'ob = "tbl_transaction." & f & " " & cmb_ad.Text
    'selectRecord.Open "SELECT tbl_transaction.transaction_id, tbl_account_title.account_title_id, tbl_account_title.account_title, tbl_transaction.transaction_date, tbl_transaction.transaction_description, tbl_transaction.dr, tbl_transaction.cr FROM tbl_acc_period_transaction INNER JOIN (tbl_account_title INNER JOIN tbl_transaction ON tbl_account_title.account_title_id = tbl_transaction.account_title_id) ON tbl_acc_period_transaction.transaction_id = tbl_transaction.transaction_id where tbl_acc_period_transaction.acc_period_id = '" & activeAccountingPeriod & "' order by '" & ob & "';", con
        'n = 0
        'lv_tr_list.ListItems.Clear
        'While selectRecord.EOF = False
            'n = n + 1
            'Set lv = lv_tr_list.ListItems.Add(, , selectRecord.Fields(0))
            'lv_tr_list.ListItems(n).Bold = True
            'lv_tr_list.ListItems(n).SubItems(1) = selectRecord.Fields(1)
            'lv_tr_list.ListItems(n).SubItems(2) = selectRecord.Fields(2)
            'lv_tr_list.ListItems(n).SubItems(3) = selectRecord.Fields(3)
            'lv_tr_list.ListItems(n).SubItems(4) = selectRecord.Fields(4)
            'lv_tr_list.ListItems(n).SubItems(5) = selectRecord.Fields(5)
            'lv_tr_list.ListItems(n).SubItems(6) = selectRecord.Fields(6)
            'selectRecord.MoveNext
        'Wend
    'selectRecord.Close
'con.Close

'For k = 1 To lv_tr_list.ListItems.Count - 1
    'c = lv_tr_list.ListItems(k)
    'd = lv_tr_list.ListItems(k).SubItems(3)
    'For j = 1 To lv_tr_list.ListItems.Count - 1
        'If c = lv_tr_list.ListItems(k + 1) Then
            'lv_tr_list.ListItems(k + 1).Text = ""
        'End If
        'If d = lv_tr_list.ListItems(k + 1).SubItems(3) Then
            'lv_tr_list.ListItems(k + 1).SubItems(3) = ""
        'End If
    'Next j
'Next k

'mod_db.DbConnection
    'selectRecord.Open "SELECT sum(Dr), Sum(Cr) FROM tbl_acc_period_transaction INNER JOIN tbl_transaction ON tbl_acc_period_transaction.transaction_id = tbl_transaction.transaction_id where tbl_acc_period_transaction.acc_period_id = '" & activeAccountingPeriod & "';", con
        'If selectRecord.EOF = False Then
            'lbl_tdr.Caption = FormatNumber(selectRecord.Fields(0))
            'lbl_tcr.Caption = FormatNumber(selectRecord.Fields(1))
        'End If
    'selectRecord.Close
'con.Close

End Sub

Private Sub cmd_tb_Click()

rpt_trialbalance.Show vbModal

End Sub

Private Sub dtp_from_Change()

If dtp_from.Value > Date Then
    MsgBox "Please enter valid Date", vbExclamation, "Accounts Management System....."
    dtp_from.Value = Date
    dtp_from.SetFocus
End If

End Sub

Private Sub dtp_to_Change()

If dtp_to.Value > Date Then
    MsgBox "Please enter valid Date", vbExclamation, "Accounts Management System....."
    dtp_to.Value = Date
    dtp_to.SetFocus
End If

End Sub

Private Sub dtp_trdate_Change()

If dtp_trdate.Value > Date Then
    MsgBox "Please enter valid Transaction Date", vbExclamation, "Accounts Management System....."
    dtp_trdate.Value = Date
    dtp_trdate.SetFocus
End If

End Sub

Private Sub Form_Load()

Set selectRecord = New ADODB.Recordset
  
mod_db.DbConnection
    selectRecord.Open "Select acc_period_from, acc_period_to, acc_period_id from tbl_accounting_period where status = 'Active';", con
        If selectRecord.EOF = False Then
            periodFrom = selectRecord.Fields(0)
            periodTo = selectRecord.Fields(1)
            frm_transaction.Caption = "Accounts Management System.....[University Version]-----[Accounting Period: " & periodFrom & " - " & periodTo & "]"
            activeAccountingPeriod = selectRecord.Fields(2)
        End If
    selectRecord.Close
con.Close

dtp_trdate.Value = Date
dtp_from.Value = Date
dtp_to.Value = Date
cmb_ad.ListIndex = 0

GenerateTransactionNo

FillAssetTree
FillLiabilityTree
FillCapitalTree
FillRevenueTree
FillExpenseTree

LoadAccounts

LoadTransactionList
    
End Sub

Private Sub lbl_sdr_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)

ques = MsgBox("Do you want to exit the Application", vbQuestion + vbYesNo, "Accounts Management System.....")
If ques = vbYes Then
    End
Else
    Cancel = 1
End If

End Sub

Private Sub LaVolpeButton1_Click()

frm_AddAccount.Show vbModal

End Sub

Private Sub LaVolpeButton3_Click()

End Sub

Private Sub LaVolpeButton2_Click()

MsgBox "Under Construction. Sorry for the Inconvenience", vbInformation, "Accounts Management System....."

End Sub

Private Sub lv_tr_Click()

On Error Resume Next

selectedItemToRemove = lv_tr.SelectedItem.Index

End Sub

Private Sub Timer1_Timer()

StatusBar1.Panels(3).Text = Time
dtp_trtime.Value = Time

End Sub

Private Sub Timer2_Timer()

pb1.Value = pb1.Value + 1

If pb1.Value >= 100 Then
    Timer2.Enabled = False
    Screen.MousePointer = vbDefault
    MsgBox "Transaction saved successfully!!!", vbInformation, "Accounts Management Software....."
    pb1.Value = 0
    Me.Enabled = True
    Reset
    GenerateTransactionNo
End If

End Sub

Private Sub tv_asset_Click()

mod_db.DbConnection
    selectRecord.Open "Select account_title_id from tbl_account_title where account_title = '" & Mid(tv_asset.SelectedItem.Text, 6, 20) & "';", con
        If selectRecord.EOF = False Then
            accountNo = selectRecord.Fields(0)
        Else
            accountNo = 0
        End If
    selectRecord.Close
con.Close

mod_db.DbConnection
    'selectRecord.Open "Select sum(Dr) from tbl_transaction where account_title_id = " & accountNo & ";", con
    selectRecord.Open "Select sum(Dr) from tbl_balancesheet where account_no = " & accountNo & ";", con
        If selectRecord.EOF = False Then
            lbl_asset.Caption = FormatNumber(selectRecord.Fields(0))
            If lbl_asset.Caption = "" Then
                lbl_asset.Caption = "0.00"
            End If
        End If
    selectRecord.Close
con.Close

End Sub

Private Sub tv_capital_Click()

mod_db.DbConnection
    selectRecord.Open "Select account_title_id from tbl_account_title where account_title = '" & Mid(tv_capital.SelectedItem.Text, 6, 20) & "';", con
        If selectRecord.EOF = False Then
            accountNo = selectRecord.Fields(0)
        Else
            accountNo = 0
        End If
    selectRecord.Close
con.Close

mod_db.DbConnection
    selectRecord.Open "Select sum(Cr) from tbl_balancesheet where account_no = " & accountNo & ";", con
        If selectRecord.EOF = False Then
            lbl_capital.Caption = FormatNumber(selectRecord.Fields(0))
            If lbl_capital.Caption = "" Then
                lbl_capital.Caption = "0.00"
            End If
        End If
    selectRecord.Close
con.Close

End Sub

Private Sub tv_expense_Click()

mod_db.DbConnection
    selectRecord.Open "Select account_title_id from tbl_account_title where account_title = '" & Mid(tv_expense.SelectedItem.Text, 6, 20) & "';", con
        If selectRecord.EOF = False Then
            accountNo = selectRecord.Fields(0)
        Else
            accountNo = 0
        End If
    selectRecord.Close
con.Close

mod_db.DbConnection
    selectRecord.Open "Select sum(Dr) from tbl_transaction where account_title_id = " & accountNo & ";", con
        If selectRecord.EOF = False Then
            lbl_expense.Caption = FormatNumber(selectRecord.Fields(0))
            If lbl_expense.Caption = "" Then
                lbl_expense.Caption = "0.00"
            End If
        End If
    selectRecord.Close
con.Close

End Sub

Private Sub tv_liability_Click()

mod_db.DbConnection
    selectRecord.Open "Select account_title_id from tbl_account_title where account_title = '" & Mid(tv_liability.SelectedItem.Text, 6, 20) & "';", con
        If selectRecord.EOF = False Then
            accountNo = selectRecord.Fields(0)
        Else
            accountNo = 0
        End If
    selectRecord.Close
con.Close

mod_db.DbConnection
    selectRecord.Open "Select sum(Cr) from tbl_balancesheet where account_no = " & accountNo & ";", con
        If selectRecord.EOF = False Then
            lbl_liability.Caption = FormatNumber(selectRecord.Fields(0))
            If lbl_liability.Caption = "" Then
                lbl_liability.Caption = "0.00"
            End If
        End If
    selectRecord.Close
con.Close

End Sub

Private Sub tv_revenue_Click()

mod_db.DbConnection
    selectRecord.Open "Select account_title_id from tbl_account_title where account_title = '" & Mid(tv_revenue.SelectedItem.Text, 6, 20) & "';", con
        If selectRecord.EOF = False Then
            accountNo = selectRecord.Fields(0)
        Else
            accountNo = 0
        End If
    selectRecord.Close
con.Close

mod_db.DbConnection
    selectRecord.Open "Select sum(Cr) from tbl_transaction where account_title_id = " & accountNo & ";", con
        If selectRecord.EOF = False Then
            lbl_revenue.Caption = FormatNumber(selectRecord.Fields(0))
            If lbl_revenue.Caption = "" Then
                lbl_revenue.Caption = "0.00"
            End If
        End If
    selectRecord.Close
con.Close

End Sub

Private Sub txt_cr_GotFocus()

txt_cr.BackColor = &H80000018
txt_cr.SelStart = 0
txt_cr.SelLength = Len(txt_cr.Text)

End Sub

Private Sub txt_cr_LostFocus()

txt_cr.BackColor = &HFFFFFF
If txt_cr.Text = "" Then
    txt_cr.Text = "0.00"
End If
txt_cr.Text = FormatNumber(txt_cr.Text)

End Sub

Private Sub txt_dr_GotFocus()

txt_dr.BackColor = &H80000018
txt_dr.SelStart = 0
txt_dr.SelLength = Len(txt_dr.Text)

End Sub

Private Sub txt_dr_KeyPress(KeyAscii As Integer)

'Call mod_Validation.ForNumericFields(KeyAscii, txt_dr)

End Sub

Private Sub txt_dr_LostFocus()

txt_dr.BackColor = &HFFFFFF
If txt_dr.Text = "" Then
    txt_dr.Text = "0.00"
End If
txt_dr.Text = FormatNumber(txt_dr.Text)

End Sub

Private Sub txt_exp_GotFocus()

txt_exp.BackColor = &H80000018

End Sub

Private Sub txt_exp_KeyPress(KeyAscii As Integer)

Call mod_Validation.ForStringFields(KeyAscii, txt_exp)

End Sub

Private Sub txt_exp_LostFocus()

txt_exp.BackColor = &HFFFFFF

End Sub

Private Sub FillAssetTree()

tv_asset.Nodes.Clear

Set nd = tv_asset.Nodes.Add(, , "Assets", "Assets")
tv_asset.Nodes(1).Bold = True
tv_asset.Nodes(1).Expanded = True

Set nd = tv_asset.Nodes.Add("Assets", tvwChild, "Fixed Assets", "Fixed Assets")
    mod_db.DbConnection
        selectRecord.Open "Select account_title, account_title_id from tbl_account_title where account_head_id = 1 and account_type = 'Fixed Asset';", con
            While selectRecord.EOF = False
                Set nd = tv_asset.Nodes.Add("Fixed Assets", tvwChild, , selectRecord.Fields(1) & ". " & selectRecord.Fields(0))
                tv_asset.Nodes(tv_asset.Nodes.Count).ForeColor = &H8000&
                selectRecord.MoveNext
            Wend
        selectRecord.Close
    con.Close
    
Set nd = tv_asset.Nodes.Add("Assets", tvwChild, "Current Assets", "Current Assets")
    mod_db.DbConnection
        selectRecord.Open "Select account_title, account_title_id from tbl_account_title where account_head_id = 1 and account_type = 'Current Asset';", con
            While selectRecord.EOF = False
                Set nd = tv_asset.Nodes.Add("Current Assets", tvwChild, , selectRecord.Fields(1) & ". " & selectRecord.Fields(0))
                tv_asset.Nodes(tv_asset.Nodes.Count).ForeColor = &H8000&
                selectRecord.MoveNext
            Wend
        selectRecord.Close
    con.Close
    
End Sub

Private Sub FillLiabilityTree()

tv_liability.Nodes.Clear

Set nd = tv_liability.Nodes.Add(, , "Liabilities", "Liabilities")
tv_liability.Nodes(1).Bold = True
tv_liability.Nodes(1).Expanded = True

Set nd = tv_liability.Nodes.Add("Liabilities", tvwChild, "Long Term Liabilities", "Long Term Liabilities")
    mod_db.DbConnection
        selectRecord.Open "Select account_title,account_title_id  from tbl_account_title where account_head_id = 2 and account_type = 'Long Term Liability';", con
            While selectRecord.EOF = False
                Set nd = tv_liability.Nodes.Add("Long Term Liabilities", tvwChild, , selectRecord.Fields(1) & ". " & selectRecord.Fields(0))
                tv_liability.Nodes(tv_liability.Nodes.Count).ForeColor = &H8000&
                selectRecord.MoveNext
            Wend
        selectRecord.Close
    con.Close
    
Set nd = tv_liability.Nodes.Add("Liabilities", tvwChild, "Current Liabilities", "Current Liabilities")
    mod_db.DbConnection
        selectRecord.Open "Select account_title, account_title_id from tbl_account_title where account_head_id = 2 and account_type = 'Current Liability';", con
            While selectRecord.EOF = False
                Set nd = tv_liability.Nodes.Add("Current Liabilities", tvwChild, , selectRecord.Fields(1) & ". " & selectRecord.Fields(0))
                tv_liability.Nodes(tv_liability.Nodes.Count).ForeColor = &H8000&
                selectRecord.MoveNext
            Wend
        selectRecord.Close
    con.Close

End Sub

Private Sub FillCapitalTree()

tv_capital.Nodes.Clear

Set nd = tv_capital.Nodes.Add(, , "Capital", "Capital")
tv_capital.Nodes(1).Bold = True
tv_capital.Nodes(1).Expanded = True

mod_db.DbConnection
    selectRecord.Open "Select account_title, account_title_id from tbl_account_title where account_head_id = 3;", con
        While selectRecord.EOF = False
            Set nd = tv_capital.Nodes.Add("Capital", tvwChild, , selectRecord.Fields(1) & ". " & selectRecord.Fields(0))
            tv_capital.Nodes(tv_capital.Nodes.Count).ForeColor = &H8000&
            selectRecord.MoveNext
        Wend
    selectRecord.Close
con.Close

End Sub

Private Sub FillRevenueTree()

tv_revenue.Nodes.Clear

Set nd = tv_revenue.Nodes.Add(, , "Revenue", "Revenue")
tv_revenue.Nodes(1).Bold = True
tv_revenue.Nodes(1).Expanded = True

mod_db.DbConnection
    selectRecord.Open "Select account_title, account_title_id  from tbl_account_title where account_head_id = 4;", con
        While selectRecord.EOF = False
            Set nd = tv_revenue.Nodes.Add("Revenue", tvwChild, , selectRecord.Fields(1) & ". " & selectRecord.Fields(0))
            tv_revenue.Nodes(tv_revenue.Nodes.Count).ForeColor = &H8000&
            selectRecord.MoveNext
        Wend
    selectRecord.Close
con.Close

End Sub

Private Sub FillExpenseTree()

tv_expense.Nodes.Clear

Set nd = tv_expense.Nodes.Add(, , "Expense", "Expense")
tv_expense.Nodes(1).Bold = True
tv_expense.Nodes(1).Expanded = True

mod_db.DbConnection
    selectRecord.Open "Select account_title, account_title_id  from tbl_account_title where account_head_id = 5;", con
        While selectRecord.EOF = False
            Set nd = tv_expense.Nodes.Add("Expense", tvwChild, , selectRecord.Fields(1) & ". " & selectRecord.Fields(0))
            tv_expense.Nodes(tv_expense.Nodes.Count).ForeColor = &H8000&
            selectRecord.MoveNext
        Wend
    selectRecord.Close
con.Close

End Sub

Private Sub GenerateTransactionNo()

mod_db.DbConnection
    selectRecord.Open "Select count(transaction_id) from tbl_acc_period_transaction where acc_period_id = " & activeAccountingPeriod & ";", con
        If selectRecord.EOF = False Then
            If selectRecord.Fields(0) = "0" Then
                 transactionNo = "1"
                 lbl_transaction_no.Caption = transactionNo
            Else
                transactionNo = selectRecord.Fields(0) + 1
                lbl_transaction_no.Caption = transactionNo
            End If
        End If
    selectRecord.Close
con.Close

End Sub

Private Sub LoadAccounts()

mod_db.DbConnection
    cmb_account.Clear
    selectRecord.Open "Select account_title from tbl_account_title order by account_head_id;", con
        While selectRecord.EOF = False
            cmb_account.AddItem selectRecord.Fields(0)
            selectRecord.MoveNext
        Wend
    selectRecord.Close
con.Close

End Sub

Private Sub Reset()

lbl_accno.Caption = ""
dtp_trdate.Value = Date
txt_exp.Text = ""
txt_dr.Text = "0.00"
txt_cr.Text = "0.00"
lbl_dr.Caption = "0.00"
lbl_cr.Caption = "0.00"
cmb_sort.ListIndex = -1
cmb_ad.ListIndex = 0
lv_tr.ListItems.Clear
LoadAccounts
LoadTransactionList
FillAssetTree
FillLiabilityTree
FillCapitalTree
FillRevenueTree
FillExpenseTree
cmb_account.SetFocus

End Sub

Private Sub LoadTransactionList()

mod_db.DbConnection
    selectRecord.Open "SELECT tbl_transaction.transaction_id, tbl_account_title.account_title_id, tbl_account_title.account_title, tbl_transaction.transaction_date, tbl_transaction.transaction_description, tbl_transaction.dr, tbl_transaction.cr FROM tbl_acc_period_transaction INNER JOIN (tbl_account_title INNER JOIN tbl_transaction ON tbl_account_title.account_title_id = tbl_transaction.account_title_id) ON tbl_acc_period_transaction.transaction_id = tbl_transaction.transaction_id where tbl_acc_period_transaction.acc_period_id = " & activeAccountingPeriod & " order by tbl_transaction.transaction_id desc, tbl_transaction.transaction_date;", con
        n = 0
        lv_tr_list.ListItems.Clear
        While selectRecord.EOF = False
            n = n + 1
            Set lv = lv_tr_list.ListItems.Add(, , selectRecord.Fields(0))
            lv_tr_list.ListItems(n).Bold = True
            lv_tr_list.ListItems(n).SubItems(1) = selectRecord.Fields(1)
            lv_tr_list.ListItems(n).SubItems(2) = selectRecord.Fields(2)
            lv_tr_list.ListItems(n).SubItems(3) = selectRecord.Fields(3)
            lv_tr_list.ListItems(n).SubItems(4) = selectRecord.Fields(4)
            lv_tr_list.ListItems(n).SubItems(5) = selectRecord.Fields(5)
            lv_tr_list.ListItems(n).SubItems(6) = selectRecord.Fields(6)
            selectRecord.MoveNext
        Wend
    selectRecord.Close
con.Close

'*****Nested Loops for not showing repeated Transaction No. and Transaction Date in lv_tr_list*****'
For k = 1 To lv_tr_list.ListItems.Count - 1
    c = lv_tr_list.ListItems(k)
    d = lv_tr_list.ListItems(k).SubItems(3)
    For j = 1 To lv_tr_list.ListItems.Count - 1
        If c = lv_tr_list.ListItems(k + 1) Then
            lv_tr_list.ListItems(k + 1).Text = ""
        End If
        If d = lv_tr_list.ListItems(k + 1).SubItems(3) Then
            lv_tr_list.ListItems(k + 1).SubItems(3) = ""
        End If
    Next j
Next k
'***************************************************************************************************'

mod_db.DbConnection
    selectRecord.Open "SELECT sum(Dr), Sum(Cr) FROM tbl_acc_period_transaction INNER JOIN tbl_transaction ON tbl_acc_period_transaction.transaction_id = tbl_transaction.transaction_id where tbl_acc_period_transaction.acc_period_id = " & activeAccountingPeriod & ";", con
        If selectRecord.EOF = False Then
            lbl_tdr.Caption = FormatNumber(selectRecord.Fields(0))
            lbl_tcr.Caption = FormatNumber(selectRecord.Fields(1))
        End If
    selectRecord.Close
con.Close

End Sub
