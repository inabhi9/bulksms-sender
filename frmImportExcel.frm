VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmImportExcel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import contacts from Excel"
   ClientHeight    =   8985
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Group"
      Height          =   1215
      Left            =   5400
      TabIndex        =   17
      Top             =   5640
      Width           =   2175
      Begin VB.ComboBox cboGroups 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Select group:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdImportToContacts 
      Caption         =   "Import to Phonebook"
      Height          =   450
      Left            =   4290
      TabIndex        =   15
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   6240
      TabIndex        =   16
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select number field"
      Height          =   855
      Left            =   3960
      TabIndex        =   12
      Top             =   6960
      Width           =   3615
      Begin VB.ComboBox cboMobileFied 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mobile nos in:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   420
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Select range"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   5175
      Begin VB.OptionButton optAllNumber 
         Caption         =   "All Numbers"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optSelected 
         Caption         =   "Selected range"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtRangeFrom 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   7
         Top             =   360
         Width           =   2310
      End
      Begin VB.TextBox txtRangeTo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   6
         Top             =   720
         Width           =   2310
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "From:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "To:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   750
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Select name field"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   3735
      Begin VB.ComboBox cboNameField 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Name in:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   420
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Excel sheet preview"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VSFlex8LCtl.VSFlexGrid vsFlexContact 
         Height          =   4935
         Left            =   120
         TabIndex        =   1
         Top             =   360
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
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   8610
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13573
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmImportExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdImportToContacts_Click()
stsBar.Panels(1).Text = "Please wait, adding..."
cmdImportToContacts.Enabled = False
Dim strName, strNumber As String
Dim intFrom, intTo As Integer
    
If cboMobileFied.ListIndex = cboNameField.ListIndex Then
    MsgBox "Name and number cannot be in same field", vbCritical
    stsBar.Panels(1).Text = "Ready"
    cmdImportToContacts.Enabled = True
    Exit Sub
End If

If optAllNumber.Value = True Then
    intFrom = 1: intTo = vsFlexContact.rows - 1
Else
    intFrom = CInt(txtRangeFrom.Text)
    intTo = CInt(txtRangeTo.Text)
End If

Dim i, j As Integer

For i = intFrom To intTo
        DoEvents
        strNumber = vsFlexContact.TextMatrix(i, cboMobileFied.ListIndex + 1)
        strName = vsFlexContact.TextMatrix(i, cboNameField.ListIndex + 1)
        
        If IsAllNumbers(strNumber) = False Then
            stsBar.Panels(1).Text = "Row " & i & " has not been added"
        Else
            With rs
                If .State = 1 Then .Close
                .Open "Select * from tblContacts where number='" & Trim(strNumber) & "'", cn, adOpenStatic, adLockOptimistic
                If .RecordCount = 0 Then
                    .AddNew
                    .Fields("name") = strName
                    .Fields("number") = strNumber
                    .Fields("GroupID") = cboGroups.ItemData(cboGroups.ListIndex)
                    .Update
                    .Close
                End If
            End With

            stsBar.Panels(1).Text = "Please wait, adding..." & i & "/" & intTo
        End If
Next
MsgBox "New contact added successfully", vbInformation, "Add New Contact"
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Call FillGroupList

LoadExcel frmMain.comDlgXL.filename

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
    cboMobileFied.AddItem Chr(64 + i)
    cboMobileFied.ListIndex = 0
    
    cboNameField.AddItem Chr(64 + i)
    cboNameField.ListIndex = 0
Next

vsFlexContact.ColWidth(0) = 600
vsFlexContact.Cell(flexcpAlignment, 1, 1, 1, 1) = 3
vsFlexContact.ColAlignment(0) = flexAlignCenterCenter
Exit Sub
EH:
    MsgBox Err.Description
    Exit Sub
    
End Sub
Sub FillGroupList()
Dim i As Integer

cboGroups.Clear

With rs
    .Open "Select * from tblGroups", cn, adOpenStatic, adLockOptimistic
    For i = 0 To .RecordCount - 1
                
        cboGroups.AddItem .Fields("name")
        cboGroups.ItemData(cboGroups.ListCount - 1) = .Fields("id")
      
        .MoveNext
    Next
    .Close
End With
If cboGroups.ListCount > 0 Then cboGroups.ListIndex = 0

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
Private Sub txtRangeFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = vbKeyTab Or KeyAscii = 13 Then Exit Sub
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
End Sub

Private Sub txtRangeTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = vbKeyTab Or KeyAscii = 13 Then Exit Sub
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
End Sub

