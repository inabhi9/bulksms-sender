VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMSZatpat SMS Bulk SMS Sender 1.0"
   ClientHeight    =   8625
   ClientLeft      =   -90
   ClientTop       =   435
   ClientWidth     =   14715
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   14715
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrows 
      Height          =   375
      Left            =   5040
      TabIndex        =   125
      Top             =   8160
      Visible         =   0   'False
      Width           =   1095
      ExtentX         =   1931
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   4440
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "BulkSMS API Profile (*.sap)|*.sap"
   End
   Begin TabDlg.SSTab sstbMain 
      Height          =   7935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Contacts"
      TabPicture(0)   =   "frmMain.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Groups"
      TabPicture(1)   =   "frmMain.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Compose SMS"
      TabPicture(2)   =   "frmMain.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "comDlgXL"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdCancel"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdSendToOutbox"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lstDummy"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdNewListing"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Outbox"
      TabPicture(3)   =   "frmMain.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lsvOutbox"
      Tab(3).Control(1)=   "cmdForward"
      Tab(3).Control(2)=   "cmdLoadJob"
      Tab(3).Control(3)=   "cmdDelJob"
      Tab(3).Control(4)=   "cboJob"
      Tab(3).Control(5)=   "cmdRemoveSel"
      Tab(3).Control(6)=   "cmdDiscardJob"
      Tab(3).Control(7)=   "cmdSaveJob"
      Tab(3).Control(8)=   "cmdStartSender"
      Tab(3).Control(9)=   "Label5"
      Tab(3).Control(10)=   "Label3"
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "Sent"
      TabPicture(4)   =   "frmMain.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label6"
      Tab(4).Control(1)=   "lsvSent"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Tools/Settings"
      TabPicture(5)   =   "frmMain.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "WebBrowBal"
      Tab(5).Control(1)=   "Frame7"
      Tab(5).Control(2)=   "cmdSaveSetting"
      Tab(5).Control(3)=   "cmdCheckBalance"
      Tab(5).Control(4)=   "cmdChangePass"
      Tab(5).ControlCount=   5
      Begin MSComctlLib.ListView lsvOutbox 
         Height          =   6615
         Left            =   -74760
         TabIndex        =   48
         Top             =   1080
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   11668
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Number"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Message"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Char"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "GSM Sender ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "CDMA Sender ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Sent"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Status"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.CommandButton cmdChangePass 
         Caption         =   "Change Password"
         Height          =   375
         Left            =   -73320
         TabIndex        =   128
         Top             =   7200
         Width           =   1575
      End
      Begin VB.CommandButton cmdCheckBalance 
         Caption         =   "Check Balance"
         Height          =   375
         Left            =   -74760
         TabIndex        =   127
         Top             =   7200
         Width           =   1335
      End
      Begin VB.CommandButton cmdNewListing 
         Caption         =   "New Listing"
         Height          =   435
         Left            =   9840
         TabIndex        =   126
         Top             =   7320
         Width           =   1335
      End
      Begin VB.CommandButton cmdForward 
         Caption         =   "Forward to"
         Height          =   375
         Left            =   -73200
         TabIndex        =   122
         Top             =   7320
         Width           =   975
      End
      Begin VB.CommandButton cmdLoadJob 
         Caption         =   "Load"
         Height          =   375
         Left            =   -67680
         TabIndex        =   112
         Top             =   7320
         Width           =   855
      End
      Begin VB.CommandButton cmdDelJob 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -66840
         TabIndex        =   113
         Top             =   7320
         Width           =   855
      End
      Begin VB.ComboBox cboJob 
         Height          =   315
         Left            =   -69480
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   7350
         Width           =   1815
      End
      Begin VB.CommandButton cmdSaveSetting 
         Caption         =   "Save settings"
         Height          =   375
         Left            =   -62400
         TabIndex        =   94
         Top             =   7320
         Width           =   1335
      End
      Begin VB.Frame Frame7 
         Caption         =   "API Settings"
         Height          =   5655
         Left            =   -74760
         TabIndex        =   55
         Top             =   840
         Width           =   13815
         Begin VB.Frame Frame17 
            Caption         =   "CDMA Sender ID:"
            Enabled         =   0   'False
            Height          =   855
            Left            =   4800
            TabIndex        =   106
            Top             =   4560
            Width           =   4095
            Begin VB.TextBox txtCDMASenderVariable 
               Enabled         =   0   'False
               Height          =   330
               Left            =   1320
               TabIndex        =   107
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Variable name:"
               Height          =   195
               Index           =   27
               Left            =   120
               TabIndex        =   108
               Top             =   428
               Width           =   1050
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "GSM Sender ID:"
            Enabled         =   0   'False
            Height          =   855
            Left            =   240
            TabIndex        =   103
            Top             =   4560
            Width           =   4095
            Begin VB.TextBox txtGSMSenderVariable 
               Enabled         =   0   'False
               Height          =   330
               Left            =   1320
               TabIndex        =   104
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Variable name:"
               Height          =   195
               Index           =   30
               Left            =   120
               TabIndex        =   105
               Top             =   428
               Width           =   1050
            End
         End
         Begin VB.CommandButton cmdLoadProfile 
            Caption         =   "Load profile"
            Enabled         =   0   'False
            Height          =   375
            Left            =   10680
            TabIndex        =   96
            Top             =   5085
            Width           =   1335
         End
         Begin VB.CommandButton cmdSaveProfile 
            Caption         =   "Save profile"
            Height          =   375
            Left            =   12120
            TabIndex        =   95
            Top             =   5085
            Width           =   1335
         End
         Begin VB.Frame Frame13 
            Caption         =   "To"
            Enabled         =   0   'False
            Height          =   1335
            Left            =   9360
            TabIndex        =   87
            Top             =   3000
            Width           =   4095
            Begin VB.TextBox txtToSap 
               Enabled         =   0   'False
               Height          =   330
               Left            =   2280
               TabIndex        =   92
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtToVariableName 
               Enabled         =   0   'False
               Height          =   330
               Left            =   120
               TabIndex        =   88
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "e.g.: ( ; ) or ( , )"
               Height          =   195
               Index           =   21
               Left            =   2280
               TabIndex        =   93
               Top             =   960
               Width           =   1035
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Saperator:"
               Height          =   195
               Index           =   20
               Left            =   2280
               TabIndex        =   91
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "e.g.: sendto or to"
               Height          =   195
               Index           =   19
               Left            =   120
               TabIndex        =   90
               Top             =   960
               Width           =   1200
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Variable name:"
               Height          =   195
               Index           =   18
               Left            =   120
               TabIndex        =   89
               Top             =   360
               Width           =   1050
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Change password"
            Enabled         =   0   'False
            Height          =   1335
            Left            =   4800
            TabIndex        =   83
            Top             =   3000
            Width           =   4095
            Begin VB.TextBox txtPassPage 
               Enabled         =   0   'False
               Height          =   330
               Left            =   120
               TabIndex        =   101
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtChangePassVariable 
               Enabled         =   0   'False
               Height          =   330
               Left            =   1920
               TabIndex        =   84
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "e.g.: chngpwd.asp"
               Height          =   195
               Index           =   26
               Left            =   120
               TabIndex        =   102
               Top             =   960
               Width           =   1320
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Page name:"
               Height          =   195
               Index           =   25
               Left            =   120
               TabIndex        =   100
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Variable name:"
               Height          =   195
               Index           =   17
               Left            =   1920
               TabIndex        =   86
               Top             =   360
               Width           =   1050
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "e.g.: newpass or np"
               Height          =   195
               Index           =   16
               Left            =   1920
               TabIndex        =   85
               Top             =   960
               Width           =   1395
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Balance Checking"
            Enabled         =   0   'False
            Height          =   1335
            Left            =   240
            TabIndex        =   79
            Top             =   3000
            Width           =   4095
            Begin VB.TextBox txtBalancePage 
               Enabled         =   0   'False
               Height          =   330
               Left            =   120
               TabIndex        =   98
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtBalanceVariable 
               Enabled         =   0   'False
               Height          =   330
               Left            =   1800
               TabIndex        =   80
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "e.g.: balance.asp"
               Height          =   195
               Index           =   24
               Left            =   120
               TabIndex        =   99
               Top             =   960
               Width           =   1230
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Page name:"
               Height          =   195
               Index           =   23
               Left            =   120
               TabIndex        =   97
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "e.g.: bal or Balance"
               Height          =   195
               Index           =   15
               Left            =   1800
               TabIndex        =   82
               Top             =   960
               Width           =   1380
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Variable name:"
               Height          =   195
               Index           =   14
               Left            =   1800
               TabIndex        =   81
               Top             =   360
               Width           =   1050
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Password"
            Height          =   1335
            Left            =   4800
            TabIndex        =   71
            Top             =   1440
            Width           =   4095
            Begin VB.TextBox txtPassVariableName 
               Enabled         =   0   'False
               Height          =   330
               Left            =   120
               TabIndex        =   73
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox txtPassValue 
               Enabled         =   0   'False
               Height          =   330
               Left            =   2160
               TabIndex        =   72
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Variable name:"
               Height          =   195
               Index           =   13
               Left            =   120
               TabIndex        =   78
               Top             =   360
               Width           =   1050
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "e.g.: pwd or pass"
               Height          =   195
               Index           =   12
               Left            =   120
               TabIndex        =   77
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Value:"
               Height          =   195
               Index           =   9
               Left            =   2160
               TabIndex        =   76
               Top             =   360
               Width           =   450
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Your account password"
               Height          =   195
               Index           =   8
               Left            =   2160
               TabIndex        =   75
               Top             =   960
               Width           =   1680
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "="
               Height          =   195
               Index           =   7
               Left            =   1920
               TabIndex        =   74
               Top             =   675
               Width           =   90
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Message body"
            Enabled         =   0   'False
            Height          =   1335
            Left            =   9360
            TabIndex        =   67
            Top             =   1440
            Width           =   4095
            Begin VB.TextBox txtMsgVariable 
               Enabled         =   0   'False
               Height          =   330
               Left            =   120
               TabIndex        =   68
               Top             =   600
               Width           =   3735
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Variable name:"
               Height          =   195
               Index           =   11
               Left            =   120
               TabIndex        =   70
               Top             =   360
               Width           =   1050
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "e.g.: message or msg"
               Height          =   195
               Index           =   10
               Left            =   120
               TabIndex        =   69
               Top             =   960
               Width           =   1500
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "User Name"
            Height          =   1335
            Left            =   240
            TabIndex        =   59
            Top             =   1440
            Width           =   4095
            Begin VB.TextBox txtUserValue 
               Enabled         =   0   'False
               Height          =   330
               Left            =   2160
               TabIndex        =   64
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox txtUserVariableName 
               Enabled         =   0   'False
               Height          =   330
               Left            =   120
               TabIndex        =   60
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "="
               Height          =   195
               Index           =   6
               Left            =   1920
               TabIndex        =   66
               Top             =   675
               Width           =   90
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "e.g.: Your user name"
               Height          =   195
               Index           =   5
               Left            =   2160
               TabIndex        =   65
               Top             =   960
               Width           =   1470
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Value:"
               Height          =   195
               Index           =   4
               Left            =   2160
               TabIndex        =   63
               Top             =   360
               Width           =   450
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "e.g.: usrname or para1"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   62
               Top             =   960
               Width           =   1590
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Variable name:"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   61
               Top             =   360
               Width           =   1050
            End
         End
         Begin VB.TextBox txtURL 
            Enabled         =   0   'False
            Height          =   330
            Left            =   240
            TabIndex        =   56
            Top             =   600
            Width           =   13215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "e.g.: http://www.testsms.com/send.php"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   58
            Top             =   960
            Width           =   2835
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Basic URL: "
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   57
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.ListBox lstDummy 
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   7320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdRemoveSel 
         Caption         =   "Remove selected"
         Height          =   375
         Left            =   -74760
         TabIndex        =   53
         Top             =   7320
         Width           =   1455
      End
      Begin VB.CommandButton cmdDiscardJob 
         Caption         =   "Discard job"
         Height          =   375
         Left            =   -62040
         TabIndex        =   51
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaveJob 
         Caption         =   "Save job"
         Height          =   375
         Left            =   -63150
         TabIndex        =   50
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmdStartSender 
         Caption         =   "Start Sending"
         Height          =   375
         Left            =   -64680
         TabIndex        =   49
         Top             =   7320
         Width           =   1335
      End
      Begin VB.CommandButton cmdSendToOutbox 
         Caption         =   "Start sending"
         Height          =   435
         Left            =   11370
         TabIndex        =   43
         Top             =   7320
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   12720
         TabIndex        =   42
         Top             =   7320
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         Caption         =   "Message"
         Height          =   4095
         Left            =   240
         TabIndex        =   41
         Top             =   3120
         Width           =   13815
         Begin VB.CommandButton cmdAddCDMASender 
            Caption         =   "Ì"
            BeginProperty Font 
               Name            =   "Wingdings 2"
               Size            =   12
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   12840
            TabIndex        =   120
            Top             =   3480
            Width           =   375
         End
         Begin VB.CommandButton cmdRemoveCDMASender 
            Caption         =   "¬"
            BeginProperty Font 
               Name            =   "Wingdings 3"
               Size            =   14.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   13200
            TabIndex        =   121
            Top             =   3480
            Width           =   375
         End
         Begin VB.CommandButton cmdAddGSMSender 
            Caption         =   "Ì"
            BeginProperty Font 
               Name            =   "Wingdings 2"
               Size            =   12
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   12840
            TabIndex        =   118
            Top             =   2400
            Width           =   375
         End
         Begin VB.CommandButton cmdRemoveGSMSender 
            Caption         =   "¬"
            BeginProperty Font 
               Name            =   "Wingdings 3"
               Size            =   14.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   13200
            TabIndex        =   119
            Top             =   2400
            Width           =   375
         End
         Begin VB.ComboBox cboCDMASender 
            Height          =   315
            Left            =   11520
            Style           =   2  'Dropdown List
            TabIndex        =   117
            Top             =   3120
            Width           =   2055
         End
         Begin VB.ComboBox cboGSMSender 
            Height          =   315
            Left            =   11520
            Style           =   2  'Dropdown List
            TabIndex        =   115
            Top             =   2040
            Width           =   2055
         End
         Begin VB.ListBox lstVariable 
            Height          =   900
            IntegralHeight  =   0   'False
            ItemData        =   "frmMain.frx":00B4
            Left            =   11520
            List            =   "frmMain.frx":00C4
            TabIndex        =   45
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtBody 
            Height          =   2865
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   720
            Width           =   11175
         End
         Begin VB.Label lblChar 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1/160 characters available"
            Height          =   195
            Left            =   9525
            TabIndex        =   130
            Top             =   3720
            Width           =   1905
         End
         Begin VB.Label Label2 
            Caption         =   "CDMA Sender ID:"
            Height          =   255
            Index           =   3
            Left            =   11520
            TabIndex        =   116
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "GSM Sender ID:"
            Height          =   255
            Index           =   2
            Left            =   11520
            TabIndex        =   114
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Variable:"
            Height          =   255
            Index           =   1
            Left            =   11520
            TabIndex        =   47
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Message body:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   46
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "SMS Information"
         Height          =   2175
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   13815
         Begin VB.CommandButton cmdSendSMSExcel 
            Caption         =   "Send SMS from Excel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   11880
            Picture         =   "frmMain.frx":00E2
            Style           =   1  'Graphical
            TabIndex        =   131
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtSubject 
            Height          =   330
            Left            =   240
            TabIndex        =   39
            Top             =   1560
            Width           =   11415
         End
         Begin VB.TextBox txtSMSNum 
            Height          =   330
            Left            =   240
            TabIndex        =   37
            Top             =   600
            Width           =   11415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Subject:"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   40
            Top             =   1320
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Multiple numbers must be saperated by semicolon ( ; )."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   38
            Top             =   960
            Width           =   4635
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Enter number(s):"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Add/Edit Group"
         Height          =   6855
         Left            =   -65520
         TabIndex        =   27
         Top             =   840
         Width           =   4575
         Begin VB.TextBox txtGName 
            Height          =   330
            Left            =   240
            TabIndex        =   32
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox txtDescription 
            Height          =   1290
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   1560
            Width           =   3975
         End
         Begin VB.CommandButton cmdAddG 
            Caption         =   "Add"
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   6360
            Width           =   1335
         End
         Begin VB.CommandButton cmdUpdateG 
            Caption         =   "Update"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   29
            Top             =   6360
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelG 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3120
            TabIndex        =   28
            Top             =   6360
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   34
            Top             =   480
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   33
            Top             =   1320
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Groups"
         Height          =   6855
         Left            =   -74760
         TabIndex        =   21
         Top             =   840
         Width           =   9135
         Begin MSComctlLib.ListView lsvGroup 
            Height          =   4815
            Left            =   240
            TabIndex        =   25
            ToolTipText     =   "Double click to edit selected group"
            Top             =   1320
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   8493
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No."
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   7937
            EndProperty
         End
         Begin VB.TextBox txtSearchGroup 
            Height          =   345
            Left            =   240
            TabIndex        =   24
            Text            =   "Enter the keyword"
            Top             =   780
            Width           =   6375
         End
         Begin VB.CommandButton cmdImportG 
            Caption         =   "Import to Compose SMS"
            Height          =   435
            Left            =   6840
            TabIndex        =   23
            Top             =   720
            Width           =   2055
         End
         Begin VB.CommandButton cmdDeleteG 
            Caption         =   "Delete selected group(s)"
            Height          =   375
            Left            =   6960
            TabIndex        =   22
            Top             =   6360
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Search contacts:"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Contacts"
         Height          =   6855
         Left            =   -74760
         TabIndex        =   3
         Top             =   840
         Width           =   9135
         Begin VB.CheckBox chkSelectAll 
            Caption         =   "Select all"
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdSendQuick 
            Caption         =   "Send QuickSMS"
            Height          =   375
            Left            =   240
            TabIndex        =   109
            Top             =   6360
            Width           =   1455
         End
         Begin VB.CommandButton cmdDeleteC 
            Caption         =   "Delete selected contact(s)"
            Height          =   375
            Left            =   6720
            TabIndex        =   20
            Top             =   6360
            Width           =   2175
         End
         Begin VB.ComboBox cboGroups 
            Height          =   315
            ItemData        =   "frmMain.frx":220C
            Left            =   3720
            List            =   "frmMain.frx":2213
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   780
            Width           =   3015
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "Import to Compose SMS"
            Height          =   435
            Left            =   6840
            TabIndex        =   7
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtSearch 
            Height          =   345
            Left            =   240
            TabIndex        =   5
            Text            =   "Enter the keyword"
            Top             =   780
            Width           =   3255
         End
         Begin MSComctlLib.ListView lsvContact 
            Height          =   4815
            Left            =   240
            TabIndex        =   4
            ToolTipText     =   "Double click to edit selected contact"
            Top             =   1440
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   8493
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No."
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Number"
               Object.Width           =   7056
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Select group to view contacts:"
            Height          =   195
            Index           =   1
            Left            =   3720
            TabIndex        =   8
            Top             =   480
            Width           =   2160
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Search contacts:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Add/Edit Contact"
         Height          =   6855
         Left            =   -65520
         TabIndex        =   2
         Top             =   840
         Width           =   4575
         Begin VB.CommandButton cmdImportExcel 
            Caption         =   "Import from Excel"
            Height          =   495
            Left            =   2760
            TabIndex        =   132
            Top             =   3000
            Width           =   1575
         End
         Begin VB.CommandButton cmdAddC 
            Caption         =   "Add"
            Height          =   375
            Left            =   1770
            TabIndex        =   17
            Top             =   6360
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelC 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3120
            TabIndex        =   19
            Top             =   6360
            Width           =   1335
         End
         Begin VB.CommandButton cmdUpdateC 
            Caption         =   "Update"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   6360
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddGroup 
            Caption         =   "Ì"
            BeginProperty Font 
               Name            =   "Wingdings 2"
               Size            =   12
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3840
            TabIndex        =   16
            Top             =   2400
            Width           =   375
         End
         Begin VB.ComboBox cboAGroups 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2400
            Width           =   3495
         End
         Begin VB.TextBox txtNumber 
            Height          =   330
            Left            =   240
            MaxLength       =   10
            TabIndex        =   12
            Top             =   1560
            Width           =   3975
         End
         Begin VB.TextBox txtName 
            Height          =   330
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Select group:"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   14
            Top             =   2160
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Number:"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   13
            Top             =   1320
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   465
         End
      End
      Begin MSComctlLib.ListView lsvSent 
         Height          =   6495
         Left            =   -74760
         TabIndex        =   123
         Top             =   1200
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   11456
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Number"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Message"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Char"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Status"
            Object.Width           =   2999
         EndProperty
      End
      Begin SHDocVwCtl.WebBrowser WebBrowBal 
         Height          =   375
         Left            =   -69840
         TabIndex        =   129
         Top             =   7320
         Visible         =   0   'False
         Width           =   1095
         ExtentX         =   1931
         ExtentY         =   661
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin MSComDlg.CommonDialog comDlgXL 
         Left            =   7320
         Top             =   7320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "Microsoft Excel File (97/2003 Format) (*.xls)|*.xls"
      End
      Begin VB.Label Label6 
         Caption         =   "Sent messages:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   124
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Saved job: "
         Height          =   195
         Left            =   -70320
         TabIndex        =   111
         Top             =   7410
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "Message to be going to send:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   52
         Top             =   840
         Width           =   2295
      End
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8250
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15372
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "API Profile: "
            TextSave        =   "API Profile: "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2999
            MinWidth        =   2999
            Text            =   "Total SMS to send : 0"
            TextSave        =   "Total SMS to send : 0"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "2:26 AM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "2/26/2010"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3480
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lstSMSNum As ListBox
Dim j As Integer
Dim strProfile As String
Private Sub cboGroups_Click()
FillContactList (cboGroups.ItemData(cboGroups.ListIndex))
End Sub

Private Sub chkSelectAll_Click()
Dim i As Integer
For i = 1 To lsvContact.ListItems.Count
    lsvContact.ListItems(i).Checked = True
Next
End Sub

Private Sub cmdAddC_Click()
On Error GoTo EH
If Trim(txtName.Text) = "" Or Trim(txtNumber.Text) = "" Then MsgBox "Name or number cannot be blank", vbExclamation, "Add New Contact": Exit Sub
If Len(Trim(txtNumber.Text)) < 10 Then MsgBox "Mobile number should be 10 digits", vbCritical, "Add New Contact": Exit Sub
With rs
    .Open "Select * from tblContacts where number='" & Trim(txtNumber.Text) & "'", cn, adOpenStatic, adLockOptimistic
    If .RecordCount > 0 Then MsgBox "Number already exist", vbCritical, "Add New Contact": .Close: Exit Sub
    
    .AddNew
    .Fields("name") = txtName.Text
    .Fields("number") = txtNumber.Text
    .Fields("GroupID") = cboAGroups.ItemData(cboAGroups.ListIndex)
    .Update
    .Close
    txtName.Text = ""
    txtNumber.Text = ""
    cboGroups_Click
    MsgBox "New contact added successfully", vbInformation, "Add New Contact"
End With
Exit Sub
EH:
    MsgBox "Please add group first", vbExclamation
    Exit Sub
End Sub

Private Sub cmdAddCDMASender_Click()
Dim strTemp As String
Dim FileObj As New Scripting.FileSystemObject
Dim strPro As TextStream
Dim i As Integer

strTemp = InputBox("Type GSM sender id to add", "New GSM Sender ID")
If Trim(strTemp) = "" Then Exit Sub
cboCDMASender.AddItem strTemp

Set strPro = FileObj.CreateTextFile(App.Path & "\CDMASender.dat", True)
For i = 0 To cboCDMASender.ListCount - 1
    strPro.WriteLine cboCDMASender.List(i)
Next
End Sub

Private Sub cmdAddG_Click()
If Trim(txtGName.Text) = "" Then MsgBox "Blank name is not allowed", vbCritical, "Add New Group": Exit Sub

With rs
    .Open "Select * from tblGroups where name='" & Trim(txtGName.Text) & "'", cn, adOpenStatic, adLockOptimistic
    If rs.RecordCount > 0 Then MsgBox "Group already exist", vbCritical, "Add New Group": rs.Close: Exit Sub
    .AddNew
    
    .Fields("Name") = txtGName.Text
    .Fields("Description") = txtDescription.Text
    
    .Update
    .Close
    txtGName.Text = ""
    txtDescription.Text = ""
    txtGName.SetFocus
    MsgBox "New group added successfully", vbInformation, "Add New Group"
End With
FillGroupList
End Sub
Sub FillGroupList()
Dim i As Integer
lsvGroup.ListItems.Clear

cboGroups.Clear
cboGroups.AddItem "All"
cboGroups.ItemData(cboGroups.ListCount - 1) = -1

cboAGroups.Clear
With rs
    .Open "Select * from tblGroups", cn, adOpenStatic, adLockOptimistic
    For i = 0 To .RecordCount - 1
        Dim itm As ListItem
        Set itm = lsvGroup.ListItems.Add
        itm.Tag = .Fields("id")
        itm.Text = Format(i + 1, "00")
        itm.SubItems(1) = .Fields("name")
        
        cboGroups.AddItem .Fields("name")
        cboGroups.ItemData(cboGroups.ListCount - 1) = .Fields("id")
        
        cboAGroups.AddItem .Fields("name")
        cboAGroups.ItemData(cboAGroups.ListCount - 1) = .Fields("id")
        
        If Trim(.Fields("description")) = "" Then
            itm.SubItems(2) = "No description"
        Else
            itm.SubItems(2) = .Fields("description")
        End If
        .MoveNext
    Next
    .Close
End With
If cboGroups.ListCount > 0 Then cboGroups.ListIndex = 0
If cboAGroups.ListCount > 0 Then cboAGroups.ListIndex = 0
End Sub
Sub FillContactList(Optional intGroupID As Integer = -1)

Dim i As Integer
lsvContact.ListItems.Clear
With rs
If rs.State = 1 Then rs.Close

    If intGroupID = -1 Then
        .Open "Select * from tblContacts", cn, adOpenStatic, adLockOptimistic
    Else
        .Open "Select * from tblContacts where GroupID = " & intGroupID, cn, adOpenStatic, adLockOptimistic
    End If
    For i = 0 To .RecordCount - 1
        Dim itm As ListItem
        Set itm = lsvContact.ListItems.Add
        itm.Tag = .Fields("id")
        itm.Text = Format(i + 1, "00")
        itm.SubItems(1) = .Fields("Name")
        itm.SubItems(2) = .Fields("Number")
        .MoveNext
    Next
    .Close
End With
End Sub
Private Sub cmdAddGroup_Click()
sstbMain.Tab = 1
txtGName.SetFocus
End Sub

Private Sub cmdAddGSMSender_Click()
Dim strTemp As String
Dim FileObj As New Scripting.FileSystemObject
Dim strPro As TextStream
Dim i As Integer

strTemp = InputBox("Type GSM sender id to add", "New GSM Sender ID")
If Trim(strTemp) = "" Then Exit Sub
cboGSMSender.AddItem strTemp

Set strPro = FileObj.CreateTextFile(App.Path & "\GSMSender.dat", True)
For i = 0 To cboGSMSender.ListCount - 1
    strPro.WriteLine cboGSMSender.List(i)
Next
End Sub

Private Sub cmdCancel_Click()
cmdNewListing_Click
End Sub

Private Sub cmdCancelC_Click()
cmdCancelC.Enabled = False
cmdAddC.Enabled = True
cmdUpdateC.Enabled = False
lsvContact.Enabled = True
txtName.Text = ""
txtNumber.Text = ""
lsvContact.SetFocus


End Sub

Private Sub cmdCancelG_Click()
cmdAddG.Enabled = True
cmdUpdateG.Enabled = False
cmdCancelG.Enabled = False
lsvGroup.Enabled = True
lsvGroup.SetFocus
txtGName.Text = ""
txtDescription.Text = ""
End Sub

Private Sub cmdChangePass_Click()
Dim BalUrl As String
Dim pos As Integer
Dim newPass As String
newPass = InputBox("Type new password", "Change Password")
If Trim(newPass) = "" Then Exit Sub

BalUrl = Replace(LCase(txtURL.Text), "http://", "")
pos = InStr(BalUrl, "/")
BalUrl = Mid(BalUrl, 1, pos) & txtPassPage.Text & "?" & txtUserVariableName.Text & "=" & _
        txtUserValue.Text & "&" & txtPassVariableName.Text & "=" & txtPassValue.Text & "&" & _
        txtChangePassVariable.Text & "=" & newPass
        
cmdChangePass.Enabled = False
stsBar.Panels(1).Text = "Please wait, changing password..."
WebBrowBal.navigate BalUrl

End Sub

Private Sub cmdCheckBalance_Click()
Dim BalUrl As String
Dim pos As Integer

BalUrl = Replace(LCase(txtURL.Text), "http://", "")
pos = InStr(BalUrl, "/")
BalUrl = Mid(BalUrl, 1, pos) & txtBalancePage.Text & "?" & txtUserVariableName.Text & "=" & txtUserValue.Text & "&" & txtPassVariableName.Text & "=" & txtPassValue.Text
cmdCheckBalance.Enabled = False
stsBar.Panels(1).Text = "Please wait, fetching balance..."
WebBrowBal.navigate BalUrl
End Sub

Private Sub cmdDeleteC_Click()
Dim i As Integer
For i = 1 To lsvContact.ListItems.Count
    If lsvContact.ListItems(i).Checked = True Then
        cn.Execute "DELETE FROM tblContacts where id=" & lsvContact.ListItems(i).Tag
    End If
Next
Call cboGroups_Click
MsgBox "Contact(s) deleted successfully", vbInformation, "Delete Group(s)"

End Sub

Private Sub cmdDeleteG_Click()
Dim i As Integer
For i = 1 To lsvGroup.ListItems.Count
    If lsvGroup.ListItems(i).Checked = True Then
        cn.Execute "DELETE FROM tblGroups where id=" & lsvGroup.ListItems(i).Tag
    End If
Next
Call FillGroupList
MsgBox "Group deleted successfully", vbInformation, "Delete Group(s)"
End Sub

Private Sub cmdDelJob_Click()
If cboJob.ListCount = 0 Then Exit Sub
cn.Execute "DELETE FROM tblJob where jobname='" & cboJob.List(cboJob.ListIndex) & "'"
'-------------------------Saving Job title to textfile--------------------------------
cboJob.RemoveItem cboJob.ListIndex

Dim strTemp As String
Dim FileObj As New Scripting.FileSystemObject
Dim strPro As TextStream

Set strPro = FileObj.CreateTextFile(App.Path & "\Job.dat", True)
For i = 0 To cboJob.ListCount - 1
    strPro.WriteLine cboJob.List(i)
Next
If cboJob.ListCount > 0 Then cboJob.ListIndex = 0
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
MsgBox "Selected job has been deleted successfully", vbInformation, "Delete Job"
End Sub

Private Sub cmdDiscardJob_Click()
lsvOutbox.ListItems.Clear
MsgBox "Job discarded", vbInformation, "Job Discard"
End Sub

Private Sub cmdForward_Click()
On Error GoTo EH
txtBody.Text = Replace(lsvOutbox.SelectedItem.SubItems(3), Chr(160), vbCrLf)
sstbMain.Tab = 2
Exit Sub
EH:
    Exit Sub
End Sub

Private Sub cmdImport_Click()
Dim i As Integer
Dim cnt As Integer
txtSMSNum.Text = "<Collected from Contacts and Group>;"

For i = 1 To lsvContact.ListItems.Count
    If lsvContact.ListItems(i).Checked Then
        If InStr(strSMSNum, lsvContact.ListItems(i).SubItems(2) & ";") = 0 Then
            lstDummy.AddItem lsvContact.ListItems(i).SubItems(2)
            strSMSNum = strSMSNum & lsvContact.ListItems(i).SubItems(2) & ";"
            cnt = cnt + 1
        End If
    End If
Next
MsgBox "Added contact(s) : " & cnt & vbCrLf & vbCrLf & "Total contact(s) : " & lstDummy.ListCount, vbInformation, "Compose SMS"

End Sub

Private Sub cmdImportExcel_Click()
On Error GoTo ErrH
comDlgXL.CancelError = True
comDlgXL.ShowOpen
frmImportExcel.Show 1
Exit Sub
ErrH:
    Exit Sub
End Sub

Private Sub cmdImportG_Click()
Dim i, j As Integer
Dim cnt As Integer

txtSMSNum.Text = "<Collected from Contacts and Group>;"

For i = 1 To lsvGroup.ListItems.Count
    If lsvGroup.ListItems(i).Checked Then
        If rs.State = 1 Then rs.Close
        rs.Open "Select number from tblContacts Where groupid=" & lsvGroup.ListItems(i).Tag, cn, adOpenStatic, adLockOptimistic
        
        For j = 0 To rs.RecordCount - 1
            If InStr(strSMSNum, rs.Fields("number") & ";") = 0 Then
                lstDummy.AddItem rs.Fields("number")
                strSMSNum = strSMSNum & rs.Fields("number") & ";"
                cnt = cnt + 1
            End If
            rs.MoveNext
        Next
    End If
Next
MsgBox "Added contact(s) : " & cnt & vbCrLf & vbCrLf & "Total contact(s) : " & lstDummy.ListCount, vbInformation, "Compose SMS"
End Sub

Private Sub cmdLoadJob_Click()
Dim i As Integer
With rs
If .State = 1 Then .Close
.Open "Select * from tblJob where Jobname='" & cboJob.List(cboJob.ListIndex) & "'", cn, adOpenStatic, adLockOptimistic

For i = 0 To .RecordCount - 1
    Dim itm As ListItem
    Set itm = lsvOutbox.ListItems.Add
    itm.Text = Format(lsvOutbox.ListItems.Count, "00")
    itm.SubItems(1) = .Fields("Name")
    itm.SubItems(2) = .Fields("number")
    itm.SubItems(3) = .Fields("message")
    itm.SubItems(4) = Len(Replace(.Fields("message"), Chr(160), vbCrLf))
    itm.SubItems(5) = .Fields("GSMSender")
    itm.SubItems(6) = .Fields("CDMASender")
    itm.SubItems(7) = Time & ", " & Date
    itm.SubItems(8) = "Pending..."
    .MoveNext
Next
.Close
End With
End Sub

Private Sub cmdLoadProfile_Click()
On Error GoTo EH
    comDlg.DialogTitle = "Open API Profile"
    comDlg.ShowOpen
    LoadProfile (comDlg.filename)
    strProfile = comDlg.FileTitle
    stsBar.Panels(2).Text = "API Profile : " & strProfile
Exit Sub
EH:
    Exit Sub

End Sub
Public Function SendingURL(strMsg As String, strNum As String, strGSMSender As String, strCDMASender As String) As String
SendingURL = txtURL.Text & "?" & txtUserVariableName.Text & "=" & txtUserValue.Text & "&" & _
            txtPassVariableName.Text & "=" & txtPassValue.Text & "&" & txtMsgVariable.Text & "=" & URLEncode(strMsg) & _
            "&" & txtToVariableName.Text & "=" & strNum & "&" & txtGSMSenderVariable.Text & "=" & strGSMSender & _
            "&" & txtCDMASenderVariable.Text & "=" & strCDMASender
End Function

Private Sub cmdNewListing_Click()
lstDummy.Clear
strSMSNum = ""
txtSMSNum.Text = ""
txtSubject.Text = ""
txtBody.Text = ""
cmdSendToOutbox.Enabled = True
MsgBox "Old job has been discard", vbInformation
End Sub

Private Sub cmdRemoveCDMASender_Click()

Dim FileObj As New Scripting.FileSystemObject
Dim strPro As TextStream
Dim i As Integer

If cboCDMASender.ListCount = 0 Then Exit Sub
cboCDMASender.RemoveItem cboCDMASender.ListIndex
If cboCDMASender.ListCount > 0 Then cboCDMASender.ListIndex = 0

Set strPro = FileObj.CreateTextFile(App.Path & "\CDMASender.dat", True)
For i = 0 To cboCDMASender.ListCount - 1
    strPro.WriteLine cboCDMASender.List(i)
Next
End Sub

Private Sub cmdRemoveGSMSender_Click()
Dim FileObj As New Scripting.FileSystemObject
Dim strPro As TextStream
Dim i As Integer
If cboGSMSender.ListCount = 0 Then Exit Sub
cboGSMSender.RemoveItem cboGSMSender.ListIndex
If cboGSMSender.ListCount > 0 Then cboGSMSender.ListIndex = 0

Set strPro = FileObj.CreateTextFile(App.Path & "\GSMSender.dat", True)
For i = 0 To cboGSMSender.ListCount - 1
    strPro.WriteLine cboGSMSender.List(i)
Next

End Sub

Private Sub cmdRemoveSel_Click()
On Error GoTo EH
Dim i As Integer
Dim cnt As Integer
cnt = 1
i = 1
Do While i = 1
    If lsvOutbox.ListItems.Item(cnt).Checked Then
        lsvOutbox.ListItems.Remove (cnt)
        i = 1
        cnt = 1
    End If
    cnt = cnt + 1
    If cnt > lsvOutbox.ListItems.Count Then Exit Do
Loop

For i = 1 To lsvOutbox.ListItems.Count
    lsvOutbox.ListItems(i).Text = Format(i, "00")
    stsBar.Panels(3).Text = "Total SMS to send : " & lsvOutbox.ListItems.Count
Next
Exit Sub
EH:
    Exit Sub
End Sub

Private Sub cmdResetSnd_Click()
txtGSMS.Text = "10"
End Sub

Private Sub cmdSaveJob_Click()
Dim strJob As String
Dim i As Integer

strJob = InputBox("Enter a job name", "Save Job")
If Trim(strJob) = "" Then Exit Sub

If rs.State = 1 Then rs.Close
rs.Open "Select * from tblJob where jobname='" & strJob & "'", cn, adOpenStatic, adLockOptimistic
If rs.RecordCount > 0 Then MsgBox "Job name already exist", vbCritical, "Add New Job": rs.Close: Exit Sub

For i = 1 To lsvOutbox.ListItems.Count
    rs.AddNew
    rs.Fields("Jobname") = strJob
    rs.Fields("name") = lsvOutbox.ListItems(i).SubItems(1)
    rs.Fields("number") = lsvOutbox.ListItems(i).SubItems(2)
    rs.Fields("message") = lsvOutbox.ListItems(i).SubItems(3)
    rs.Fields("GSMSender") = lsvOutbox.ListItems(i).SubItems(5)
    rs.Fields("CDMASender") = lsvOutbox.ListItems(i).SubItems(6)
    rs.Update
Next

'-------------------------Saving Job title to textfile--------------------------------
cboJob.AddItem strJob
Dim strTemp As String
Dim FileObj As New Scripting.FileSystemObject
Dim strPro As TextStream

Set strPro = FileObj.CreateTextFile(App.Path & "\Job.dat", True)
For i = 0 To cboJob.ListCount - 1
    strPro.WriteLine cboJob.List(i)
Next

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
MsgBox "Job saved", vbInformation, "Save Job"
End Sub

Private Sub cmdSaveProfile_Click()
On Error GoTo EH
    comDlg.DialogTitle = "Save API Profile"
    comDlg.ShowSave
    SaveProfile (comDlg.filename)
    strProfile = comDlg.FileTitle
    stsBar.Panels(2).Text = "API Profile : " & strProfile
Exit Sub
EH:
    Exit Sub
End Sub

Private Sub cmdSaveSetting_Click()

Dim FileObj As New Scripting.FileSystemObject
Dim strPro As TextStream
Set strPro = FileObj.CreateTextFile(App.Path & "\Settings.dat")
strPro.WriteLine strProfile

End Sub

Private Sub cmdSendQuick_Click()
Dim frmQuick As New frmChat
frmQuick.Caption = "Send QuickSMS to " & UCase(lsvContact.SelectedItem.SubItems(1))
frmQuick.MobNo = lsvContact.SelectedItem.SubItems(2)
frmQuick.Show
End Sub

Private Sub cmdSendSMSExcel_Click()
On Error GoTo ErrH
comDlgXL.CancelError = True
comDlgXL.ShowOpen
frmExcel.Show 1
Exit Sub
ErrH:
    Exit Sub
End Sub

Private Sub cmdSendToOutbox_Click()
If Trim(txtBody.Text) = "" Or Trim(txtSMSNum.Text) = "" Then MsgBox "Number or message cannot be blank", vbCritical, "Send SMS": Exit Sub
Dim i As Integer
Dim pos As Integer
Dim strTemp As String
Dim strNum As String
strNum = Replace(txtSMSNum.Text, "<Collected from Contacts and Group>;", "")
pos = 1
If Right(txtSMSNum.Text, 1) <> ";" Then txtSMSNum.Text = txtSMSNum.Text & ";"
Do While pos > 0
    pos = InStr(strNum, ";")
    If pos = 0 Then Exit Do
    strTemp = Mid(strNum, 1, pos - 1)
    lstDummy.AddItem strTemp
    strNum = Replace(strNum, strTemp & ";", "")
Loop

For i = 0 To lstDummy.ListCount - 1
    Dim strMessage As String
    If Trim(txtSubject.Text) <> "" Then strMessage = "Subject: " & txtSubject.Text & Chr(160)
    strMessage = strMessage & txtBody.Text
    strMessage = Replace(strMessage, vbCrLf, Chr(160))
    With rs

    If .State = 1 Then .Close
    .Open "Select * from tblContacts where number='" & lstDummy.List(i) & "'", cn, adOpenStatic, adLockOptimistic


    If rs.RecordCount = 0 Then
        strMessage = Replace(strMessage, "<Name>", "Sir/Madam")
    Else
        strMessage = Replace(strMessage, "<Name>", .Fields("name"))
    End If

    strMessage = Replace(strMessage, "<Number>", lstDummy.List(i))
    strMessage = Replace(strMessage, "<Date>", Date)
    strMessage = Replace(strMessage, "<Time>", Time)

    Dim itm As ListItem
    Set itm = lsvOutbox.ListItems.Add
    itm.Text = Format(i + 1, "00")
    If rs.RecordCount = 0 Then itm.SubItems(1) = "No Name" Else itm.SubItems(1) = .Fields("name")
    itm.SubItems(2) = lstDummy.List(i)
    itm.SubItems(3) = strMessage
    itm.SubItems(4) = Format(Len(Replace(strMessage, Chr(160), vbCrLf)), "00")
    itm.SubItems(5) = cboGSMSender.List(cboGSMSender.ListIndex)
    itm.SubItems(6) = cboCDMASender.List(cboCDMASender.ListIndex)
    itm.SubItems(7) = Time & ", " & Date
    itm.SubItems(8) = "Pending..."

    End With
    strMessage = ""
Next
stsBar.Panels(3).Text = "Total SMS to send : " & lstDummy.ListCount
MsgBox lstDummy.ListCount & " Messages added to the outbox", vbInformation, "Send to Outbox"
txtSMSNum.Text = ""
txtSubject.Text = ""
txtBody.Text = ""
stsBar.Panels(1).Text = "Initializing..."
cmdSendToOutbox.Enabled = False
cmdNewListing_Click
cmdStartSender_Click
End Sub

Private Sub cmdStartSender_Click()
If lsvOutbox.ListItems.Count = 0 Then Exit Sub

For i = 1 To lsvOutbox.ListItems.Count
    lsvOutbox.ListItems(i).SubItems(2) = Right(lsvOutbox.ListItems(i).SubItems(2), 10)
Next


WebBrows.navigate "www.google.com"
cmdStartSender.Enabled = False
j = 1
End Sub

Private Sub cmdUpdateC_Click()
If Trim(txtName.Text) = "" Or Trim(txtNumber.Text) = "" Then MsgBox "Name or number cannot be blank", vbExclamation, "Update Contact": Exit Sub
If Len(Trim(txtNumber.Text)) < 10 Then MsgBox "Mobile number should be 10 digits", vbCritical, "Add New Contact": Exit Sub

With rs
If .State = 1 Then .Close
.Open "Select * from tblContacts Where ID=" & lsvContact.SelectedItem.Tag, cn, adOpenStatic, adLockOptimistic

.Fields("Name") = txtName.Text
.Fields("Number") = txtNumber.Text
.Fields("GroupID") = cboAGroups.ItemData(cboAGroups.ListIndex)
.Update
.Close


cmdCancelC.Enabled = False
cmdAddC.Enabled = True
cmdUpdateC.Enabled = False
lsvContact.Enabled = True
txtName.Text = ""
txtNumber.Text = ""
lsvContact.SetFocus
Call cboGroups_Click

MsgBox "Selected item updated successfully", vbInformation, "Update Item"
End With
End Sub

Private Sub cmdUpdateG_Click()
If Trim(txtGName.Text) = "" Then MsgBox "Blank name is not allowed", vbCritical, "Add New Group": Exit Sub
cn.Execute "UPDATE tblGroups SET name='" & txtGName.Text & "', description='" & txtDescription.Text & "' where id=" & lsvGroup.SelectedItem.Tag
Call FillGroupList
cmdCancelG_Click
MsgBox "Record successfully updated", vbInformation, "Update Group"
End Sub


Private Sub Form_Load()
'sstbMain.Tab = 0
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db.mdb;Persist Security Info=False"
Call FillGroupList
Call FillContactList


Dim FileObj As New Scripting.FileSystemObject
Dim strPro As TextStream
Set strPro = FileObj.OpenTextFile(App.Path & "\Settings.dat")
strProfile = strPro.ReadLine
LoadProfile App.Path & "\SMSProfile\" & strProfile
stsBar.Panels(2).Text = "API Profile : " & strProfile
strPro.Close
'-------------------------------------------------Load GSM Sender ID-------------------------------
'Set strPro = FileObj.OpenTextFile(App.Path & "\GSMSender.dat")
'For i = 0 To 50
'    Dim strTemp As String
'    If strPro.AtEndOfStream = True Then Exit For
'    strTemp = strPro.ReadLine
'    cboGSMSender.AddItem strTemp
'Next


'--------------------------------------------------------------------------------------------------

'-------------------------------------------------Load CDMA Sender ID-------------------------------
'Set strPro = FileObj.OpenTextFile(App.Path & "\CDMASender.dat")
'For i = 0 To 50
'
'    If strPro.AtEndOfStream = True Then Exit For
'    strTemp = strPro.ReadLine
'    cboCDMASender.AddItem strTemp
'Next


'--------------------------------------------------------------------------------------------------

'-------------------------------------------------Load JOB Sender ID-------------------------------
Set strPro = FileObj.OpenTextFile(App.Path & "\Job.dat")
For i = 0 To 50
    If strPro.AtEndOfStream = True Then Exit For
    strTemp = strPro.ReadLine
    cboJob.AddItem strTemp
Next
If cboJob.ListCount > 0 Then cboJob.ListIndex = 0
'--------------------------------------------------------------------------------------------------


End Sub




Private Sub lstVariable_DblClick()
txtBody.Text = txtBody.Text & "<" & lstVariable.List(lstVariable.ListIndex) & ">"
End Sub

Private Sub lsvContact_DblClick()
cmdCancelC.Enabled = True
cmdAddC.Enabled = False
cmdUpdateC.Enabled = True
lsvContact.Enabled = False
txtName.Text = lsvContact.SelectedItem.SubItems(1)
txtNumber.Text = lsvContact.SelectedItem.SubItems(2)
txtName.SetFocus

End Sub

Private Sub lsvGroup_DblClick()
cmdAddG.Enabled = False
cmdUpdateG.Enabled = True
cmdCancelG.Enabled = True
lsvGroup.Enabled = False

txtGName.Text = lsvGroup.SelectedItem.SubItems(1)
txtDescription.Text = lsvGroup.SelectedItem.SubItems(2)
End Sub

Sub SaveProfile(strFileName As String)
Dim FileObj As New Scripting.FileSystemObject
Dim strPro As TextStream
Set strPro = FileObj.CreateTextFile(strFileName)
strPro.WriteLine txtURL.Text
strPro.WriteLine txtUserVariableName.Text
strPro.WriteLine txtUserValue.Text
strPro.WriteLine txtPassVariableName.Text
strPro.WriteLine txtPassValue.Text
strPro.WriteLine txtMsgVariable.Text
strPro.WriteLine txtBalanceVariable.Text
strPro.WriteLine txtBalancePage.Text
strPro.WriteLine txtChangePassVariable.Text
strPro.WriteLine txtPassPage.Text
strPro.WriteLine txtToVariableName.Text
strPro.WriteLine txtToSap.Text
strPro.WriteLine txtGSMSenderVariable.Text
strPro.WriteLine txtCDMASenderVariable.Text
End Sub
Sub LoadProfile(strFileName As String)
Dim FileObj As New Scripting.FileSystemObject
Dim strPro As TextStream
Set strPro = FileObj.OpenTextFile(strFileName)

txtURL.Text = strPro.ReadLine
txtUserVariableName.Text = strPro.ReadLine
txtUserValue.Text = strPro.ReadLine
txtPassVariableName.Text = strPro.ReadLine
txtPassValue.Text = strPro.ReadLine
txtMsgVariable.Text = strPro.ReadLine
txtBalanceVariable.Text = strPro.ReadLine
txtBalancePage.Text = strPro.ReadLine
txtChangePassVariable.Text = strPro.ReadLine
txtPassPage.Text = strPro.ReadLine
txtToVariableName.Text = strPro.ReadLine
txtToSap.Text = strPro.ReadLine
txtGSMSenderVariable.Text = strPro.ReadLine
txtCDMASenderVariable.Text = strPro.ReadLine
End Sub

Private Sub txtBody_Change()
lblChar.Caption = CountChar(txtBody.Text)
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyTab Or KeyAscii = 13 Then Exit Sub
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
End Sub

Private Sub txtSearch_Change()
On Error Resume Next
Dim indx As Integer
indx = lsvContact.FindItem(txtSearch.Text, 1, , 1).Index
lsvContact.ListItems(indx).Selected = True
End Sub

Private Sub txtSMSNum_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyTab Or KeyAscii = 13 Or KeyAscii = Asc(";") Then Exit Sub
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0

End Sub

Private Sub txtSMSNum_LostFocus()
If Right(txtSMSNum.Text, 1) <> ";" Then txtSMSNum.Text = txtSMSNum.Text & ";"
End Sub

Private Sub WebBrowBal_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Dim htDoc As New HTMLDocument
Set htDoc = WebBrowBal.document

If InStr(LCase(htDoc.documentElement.outerText), "cannot") Then stsBar.Panels(1).Text = "Ready": Exit Sub
stsBar.Panels(1).Text = htDoc.documentElement.outerText

cmdCheckBalance.Enabled = True
cmdChangePass.Enabled = True
End Sub

Private Sub WebBrows_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Dim ahtDoc As New HTMLDocument
Set ahtDoc = WebBrows.document

If LCase(ahtDoc.documentElement.outerText) = "google" Then WebBrows.navigate SendingURL(Replace(lsvOutbox.ListItems(j).SubItems(3), Chr(160), vbCrLf), lsvOutbox.ListItems(j).SubItems(2), lsvOutbox.ListItems(j).SubItems(5), lsvOutbox.ListItems(j).SubItems(6)): j = j + 1: Exit Sub
If lsvOutbox.ListItems.Count = 0 Then Exit Sub
    If j <= lsvOutbox.ListItems.Count Then
        lsvOutbox.ListItems(j).SubItems(8) = "Sent at " & Time
        Dim itm As ListItem
        Set itm = lsvSent.ListItems.Add
        itm.Text = lsvSent.ListItems.Count
        itm.SubItems(1) = lsvOutbox.ListItems(j).SubItems(1)
        itm.SubItems(2) = lsvOutbox.ListItems(j).SubItems(2)
        itm.SubItems(3) = lsvOutbox.ListItems(j).SubItems(3)
        itm.SubItems(4) = lsvOutbox.ListItems(j).SubItems(4)
        itm.SubItems(5) = ahtDoc.documentElement.outerText
        
        WebBrows.navigate SendingURL(Replace(lsvOutbox.ListItems(j).SubItems(3), Chr(160), vbCr), lsvOutbox.ListItems(j).SubItems(2), lsvOutbox.ListItems(j).SubItems(5), lsvOutbox.ListItems(j).SubItems(6))
        stsBar.Panels(1).Text = "Seding message " & j & "/" & lsvOutbox.ListItems.Count
        
        j = j + 1
    Else
        cmdStartSender.Enabled = True
        stsBar.Panels(1).Text = "Ready"
        stsBar.Panels(3).Text = "Total SMS to send : 0"
        MsgBox "Done", vbInformation, "SMS Sent"
        lsvOutbox.ListItems.Clear
        
    End If

End Sub
