VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrows 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      ExtentX         =   2566
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
      Location        =   ""
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2445
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2445
      Width           =   1095
   End
   Begin VB.TextBox txtMsg 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      Caption         =   "1/160 characters available"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Type message: "
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MobNo As String
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSend_Click()
frmMain.stsBar.Panels(1).Text = "Sending message..."
WebBrows.navigate frmMain.SendingURL(txtMsg.Text, MobNo, frmMain.cboGSMSender, frmMain.cboCDMASender)

cmdSend.Enabled = False
End Sub

Private Sub txtMsg_Change()
lblChar.Caption = CountChar(txtMsg.Text)
End Sub

Private Sub WebBrows_NavigateComplete2(ByVal pDisp As Object, URL As Variant)


Dim htDoc As New HTMLDocument
Set htDoc = WebBrows.document
frmMain.stsBar.Panels(1).Text = "Message from server : " & htDoc.documentElement.outerText
cmdSend.Enabled = True
txtMsg.Text = ""
End Sub

