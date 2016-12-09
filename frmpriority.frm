VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmpriority 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkloadtshowog 
      Caption         =   "Load Tshowog"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load Tshowog"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      Picture         =   "frmpriority.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
   End
   Begin VB.ListBox DZLIST 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   12135
   End
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   3735
      Left            =   3480
      TabIndex        =   0
      Top             =   2640
      Width           =   8655
      _cx             =   15266
      _cy             =   6588
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmpriority.frx":076A
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   2640
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpriority.frx":07FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpriority.frx":0B97
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpriority.frx":0F31
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpriority.frx":1C0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpriority.frx":205D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpriority.frx":2817
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   1164
      ButtonWidth     =   1217
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ADD"
            Key             =   "ADD"
            Object.ToolTipText     =   "ADDS NEW RECORD"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "New"
                  Text            =   "New"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Open"
                  Text            =   "Open"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OPEN"
            Key             =   "OPEN"
            Object.ToolTipText     =   "OPEN/EDIT EXISTING RECORD"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "SAVE"
            Key             =   "SAVE"
            Object.ToolTipText     =   "SAVES RECORD"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "DELETE"
            Key             =   "DELETE"
            Object.ToolTipText     =   "DELETE THE RECORD"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EXIT"
            Key             =   "EXIT"
            Object.ToolTipText     =   "EXIT FROM THE FORM"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   3
   End
End
Attribute VB_Name = "frmpriority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DZstr As String
Private Sub VSFlexGrid1_Click()

End Sub

Private Sub chkloadtshowog_Click()
FillGrid
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select DZONGKHAGCODE,DZONGKHAGNAME from tbldzongkhag Order by DZONGKHAGCODE", MHVDB, adOpenStatic
With rs
Do While Not .EOF
   DZLIST.AddItem Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE
   .MoveNext
Loop
End With
'FillGrid
End Sub
Private Sub FillGrid()
DZstr = ""
If chkloadtshowog.Value = 1 Then
'FRMRPTLANDDETAILS.Width = 13980

For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       DZstr = DZstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next
If Len(DZstr) > 0 Then
   DZstr = "(" + Left(DZstr, Len(DZstr) - 1) + ")"
 
Else
'FRMRPTLANDDETAILS.Width = 7560
chkloadtshowog.Value = 0
LSTPR.Clear
   MsgBox "DZONGKHAG NOT SELECTED !!!"
   Exit Sub
End If
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^Sl.Nl.|^Status|^DGT|^Priority|^"
Mygrid.ColWidth(0) = 585
Mygrid.ColWidth(1) = 585
Mygrid.ColWidth(2) = 6000
Mygrid.ColWidth(3) = 960
Mygrid.ColWidth(4) = 960
Dim rs As New ADODB.Recordset

Set rs = Nothing
'LSTPR.Clear
rs.Open "select * from tbltshewog where dzongkhagid in " & DZstr & "order by dzongkhagid,gewogid,tshewogid", MHVDB, adOpenStatic
i = 1
With rs
Do While Not .EOF
FindDZ rs!dzongkhagid
FindGE rs!dzongkhagid, rs!gewogid

Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i

Mygrid.TextMatrix(i, 2) = rs!dzongkhagid & " " & Trim(Dzname) + " |" + rs!gewogid & " " & Trim(GEname) + " |" + rs!tshewogid & " " & Trim(rs!tshewogname) 'Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE
i = i + 1
   'LSTPR.AddItem rs!dzongkhagid & " " & Trim(Dzname) + " |" + rs!gewogid & " " & Trim(GEname) + " |" + rs!tshewogid & " " & Trim(rs!tshewogname) 'Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE
   .MoveNext
Loop
End With


Else
'FRMRPTLANDDETAILS.Width = 7560
End If




End Sub

Private Sub Mygrid_Click()
Dim i As Integer
Dim CBOSTR As String
If Mygrid.col <> 0 Then
Mygrid.Editable = flexEDKbdMouse
Else

Mygrid.Editable = flexEDNone
End If





If Mygrid.col = 3 Then
CBOSTR = ""
For i = 1 To Mygrid.Rows - 1
If i <> Val(Mygrid.TextMatrix(i, 3)) Then
CBOSTR = "|" & i & CBOSTR
End If
Next
Mygrid.ComboList = CBOSTR
Else

Mygrid.ComboList = ""
End If


End Sub
