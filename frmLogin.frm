VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login - SMSZatpat.com"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox txtUsername 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   5190
      TabIndex        =   6
      Top             =   0
      Width           =   5190
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter your user name and password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   4125
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   0
         Picture         =   "frmLogin.frx":000C
         Top             =   0
         Width           =   780
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowBal 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
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
   Begin InetCtlsObjects.Inet Inet 
      Left            =   720
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "&Password:"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1740
      Width           =   1170
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "&User name:"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1260
      Width           =   1170
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdLogin_Click()
Dim BalUrl As String
Dim pos As Integer
Dim Sender As String
If Trim(txtUsername.Text) = "" Or Trim(txtPassword.Text) = "" Then
    MsgBox "User name or password cannot be blank", vbCritical
    Exit Sub
End If

BalUrl = Replace(LCase(frmMain.txtURL.Text), "http://", "")
pos = InStr(BalUrl, "/")
BalUrl = Mid(BalUrl, 1, pos) & frmMain.txtBalancePage.Text & "?" & frmMain.txtUserVariableName.Text & "=" & txtUsername.Text & "&" & frmMain.txtPassVariableName.Text & "=" & txtPassword.Text
cmdLogin.Enabled = False
Me.Caption = "Logging in..."
txtUsername.Enabled = False
txtPassword.Enabled = False


Sender = Inet.OpenURL("http://api.smszatpat.com/senderid.php?loginid=" & txtUsername.Text & "&type=gsm")
Set colSenderGSM = New Collection
Dim snd As String
While (InStr(Sender, ",") <> 0)
    pos = InStr(Sender, ",")
    snd = Mid(Sender, 1, pos - 1)
    colSenderGSM.Add snd
    Sender = Replace(Sender, snd & ",", "")
Wend
Me.Caption = "Getting sender ID..."

Sender = Inet.OpenURL("http://api.smszatpat.com/senderid.php?loginid=" & txtUsername.Text & "&type=cdma")
Set colSenderCDMA = New Collection

While (InStr(Sender, ",") <> 0)
    pos = InStr(Sender, ",")
    snd = Mid(Sender, 1, pos - 1)
    colSenderCDMA.Add snd
    Sender = Replace(Sender, snd & ",", "")
Wend

For i = 1 To colSenderGSM.Count
    frmMain.cboGSMSender.AddItem colSenderGSM.Item(i)
Next
If frmMain.cboGSMSender.ListCount > 0 Then frmMain.cboGSMSender.ListIndex = 0

For i = 1 To colSenderCDMA.Count
    frmMain.cboCDMASender.AddItem colSenderCDMA.Item(i)
Next
If frmMain.cboCDMASender.ListCount > 0 Then frmMain.cboCDMASender.ListIndex = 0

WebBrowBal.navigate BalUrl
Me.Caption = "Getting balance..."
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword)

End Sub

Private Sub txtUsername_GotFocus()
txtUsername.SelStart = 0
txtUsername.SelLength = Len(txtUsername)
End Sub

Private Sub WebBrowBal_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Dim htDoc As New HTMLDocument
Set htDoc = WebBrowBal.document

If InStr(LCase(htDoc.documentElement.outerText), "cannot") Then Exit Sub


If InStr(LCase(htDoc.documentElement.outerText), "invalid") Then
    MsgBox "Login failed!" & vbCrLf & htDoc.documentElement.outerText, vbCritical
    cmdLogin.Enabled = True
    Me.Caption = "Login failed"
    txtUsername.Enabled = True
txtPassword.Enabled = True

Else
    MsgBox "Login successfull!" & vbCrLf & htDoc.documentElement.outerText, vbInformation
    frmMain.txtUserValue.Text = txtUsername.Text
    frmMain.txtPassValue.Text = txtPassword.Text
    
    Unload Me
    
    frmMain.Show
    
End If

End Sub
