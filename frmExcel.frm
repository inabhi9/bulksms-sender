VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmExcel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send SMS from Excel sheet"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   Icon            =   "frmExcel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   9840
      TabIndex        =   24
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Send sms from excel sheet"
      Height          =   8055
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   10935
      Begin VB.TextBox txtMsgBody 
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   6000
         Width           =   7095
      End
      Begin VB.Frame Frame5 
         Caption         =   "Select mobile number field"
         Height          =   855
         Left            =   7560
         TabIndex        =   18
         Top             =   4680
         Width           =   3255
         Begin VB.ComboBox cboMobileFied 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Mobile nos in:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   420
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Select range"
         Height          =   2775
         Left            =   7560
         TabIndex        =   11
         Top             =   1800
         Width           =   3255
         Begin VB.TextBox txtRangeTo 
            Enabled         =   0   'False
            Height          =   315
            Left            =   600
            MaxLength       =   4
            TabIndex        =   16
            Top             =   2160
            Width           =   2310
         End
         Begin VB.TextBox txtRangeFrom 
            Enabled         =   0   'False
            Height          =   315
            Left            =   600
            MaxLength       =   4
            TabIndex        =   15
            Top             =   1320
            Width           =   2310
         End
         Begin VB.OptionButton optSelected 
            Caption         =   "Selected range"
            Height          =   255
            Left            =   600
            TabIndex        =   14
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton optAllNumber 
            Caption         =   "All Numbers"
            Height          =   255
            Left            =   600
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "To:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "From:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Select field"
         Height          =   2055
         Left            =   7560
         TabIndex        =   9
         Top             =   5760
         Width           =   3255
         Begin VB.ListBox lstField 
            Height          =   1425
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Choose sender ID"
         Height          =   1215
         Left            =   7560
         TabIndex        =   4
         Top             =   480
         Width           =   3255
         Begin VB.ComboBox cboSenderCDMA 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   720
            Width           =   1935
         End
         Begin VB.ComboBox cboSenderGSM 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "CDMA:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   750
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "GSM:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   270
            Width           =   735
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsFlexContact 
         Height          =   4935
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   7155
         _cx             =   12621
         _cy             =   8705
         Appearance      =   3
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   600
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Message body"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   5760
         Width           =   1035
      End
      Begin VB.Label lblChar 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1/160 characters available"
         Height          =   195
         Left            =   5445
         TabIndex        =   22
         Top             =   7680
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Excel sheet:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdSendToOutbox 
      Caption         =   "Send to outbox"
      Height          =   450
      Left            =   8160
      TabIndex        =   0
      Top             =   8400
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   8910
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19976
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdSendToOutbox_Click()
stsBar.Panels(1).Text = "Please wait, adding..."
 Dim itm As ListItem
 Dim strMessage, strNumber As String
 Dim intFrom, intTo As Integer
    
If optAllNumber.Value = True Then
    intFrom = 1: intTo = vsFlexContact.rows - 1
Else
    intFrom = CInt(txtRangeFrom.Text)
    intTo = CInt(txtRangeTo.Text)
End If
 
    
Dim i, j As Integer

For i = intFrom To intTo
    With frmMain
    strMessage = txtMsgBody.Text
    
    For j = 0 To lstField.ListCount - 1
        Dim Tempstr, TempReplace As String
        TempReplace = vsFlexContact.TextMatrix(i, j + 1)
        
        Tempstr = Replace(strMessage, "<" & lstField.List(j) & ">", TempReplace)
        strMessage = Trim(Tempstr)
    Next

        strNumber = vsFlexContact.TextMatrix(i, cboMobileFied.ListIndex + 1)
        If IsAllNumbers(strNumber) = False Then
            stsBar.Panels(1).Text = "Row " & i & " has not been added"
        Else
            If Trim(txtMsgBody.Text) = "" Then MsgBox "Message cannot be blank", vbCritical: stsBar.Panels(1).Text = "Ready": Exit Sub:
            
            strMessage = Replace(strMessage, vbCrLf, Chr(160))
            
            Set itm = .lsvOutbox.ListItems.Add
            itm.Text = Format(.lsvOutbox.ListItems.Count, "0000")
            itm.SubItems(1) = "No Name"
            itm.SubItems(2) = strNumber
            itm.SubItems(3) = strMessage
            itm.SubItems(4) = Format(Len(Replace(strMessage, Chr(160), vbCrLf)), "000")
            itm.SubItems(5) = cboSenderGSM.List(cboSenderGSM.ListIndex)
            itm.SubItems(6) = cboSenderCDMA.List(cboSenderCDMA.ListIndex)
            itm.SubItems(7) = Time & ", " & Date
            itm.SubItems(8) = "Pending..."
            stsBar.Panels(1).Text = "Please wait, adding..." & i & "/" & intTo
        End If
    End With
Next
MsgBox "Adding completed", vbInformation
Unload Me
End Sub

Private Sub LoadExcel(ByVal filename As String)
On Error GoTo EH
Dim i As Integer
vsFlexContact.LoadGrid filename, flexFileExcel

For i = 1 To vsFlexContact.rows - 1
    vsFlexContact.TextMatrix(i, 0) = i
Next
For i = 1 To vsFlexContact.cols - 1
    vsFlexContact.TextMatrix(0, i) = Chr(64 + i)
    lstField.AddItem Chr(64 + i)
    lstField.ListIndex = 0
    
    cboMobileFied.AddItem Chr(64 + i)
    cboMobileFied.ListIndex = 0
    
Next

vsFlexContact.ColWidth(0) = 600
vsFlexContact.Cell(flexcpAlignment, 1, 1, 1, 1) = 3
vsFlexContact.ColAlignment(0) = flexAlignCenterCenter
Exit Sub
EH:
    MsgBox Err.Description
    Exit Sub
    
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next

LoadExcel frmMain.comDlgXL.filename

Dim i As Integer
For i = 1 To colSenderGSM.Count
    cboSenderGSM.AddItem colSenderGSM.Item(i)
    cboSenderGSM.ListIndex = 0
Next
For i = 1 To colSenderCDMA.Count
    cboSenderCDMA.AddItem colSenderCDMA.Item(i)
    cboSenderCDMA.ListIndex = 0
Next

End Sub

Private Sub lstField_DblClick()
'txtMsgBody.Text = txtMsgBody.Text & "<" & lstField.List(lstField.ListIndex) & ">"
Dim intSelStart As Integer
Dim strSelText As String
Dim strRep As String

intSelStart = txtMsgBody.SelStart + 1

If intSelStart - 1 = Len(txtMsgBody) Then
    txtMsgBody.Text = txtMsgBody.Text & "<" & lstField.List(lstField.ListIndex) & ">"
Else
    txtMsgBody.Text = Trim(txtMsgBody.Text)
    strSelText = Mid(txtMsgBody, intSelStart)
    strRep = "<" & lstField.List(lstField.ListIndex) & ">" & strSelText
    txtMsgBody.Text = Replace(txtMsgBody, strSelText, strRep)
End If
txtMsgBody.SetFocus
txtMsgBody.SelStart = intSelStart + 2
End Sub

Private Sub optAllNumber_Click()
Label4.Enabled = False
Label5.Enabled = False
txtRangeFrom.Enabled = False
txtRangeTo.Enabled = False

End Sub

Private Sub optSelected_Click()
Label4.Enabled = True
Label5.Enabled = True
txtRangeFrom.Enabled = True
txtRangeTo.Enabled = True
End Sub

Private Sub txtMsgBody_Change()
lblChar.Caption = CountChar(txtMsgBody.Text)
End Sub

Private Sub txtRangeFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = vbKeyTab Or KeyAscii = 13 Then Exit Sub
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
End Sub

Private Sub txtRangeTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = vbKeyTab Or KeyAscii = 13 Then Exit Sub
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
End Sub
