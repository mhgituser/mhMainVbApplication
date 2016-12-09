VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmchemical 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C H E M I C A L S . . ."
   ClientHeight    =   8085
   ClientLeft      =   6810
   ClientTop       =   2070
   ClientWidth     =   7560
   Icon            =   "frmchemical.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   7560
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7335
      Begin VB.TextBox txtuses 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   3000
         Width           =   4215
      End
      Begin VB.TextBox txtpreharvest 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   2520
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker txtapproveddate 
         Height          =   300
         Left            =   2880
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   75628545
         CurrentDate     =   41464
      End
      Begin VB.TextBox txttradename 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtsafereentry 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtrecomendeddose 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   3480
         Width           =   4215
      End
      Begin MSDataListLib.DataCombo cboChemical 
         Bindings        =   "frmchemical.frx":0E42
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   "ItemCode"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Uses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pre Harvest-Interval Days "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   2745
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Recomended Dose Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   2580
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Date of Approve"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Chemical Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trade Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Safe Re-entry Period days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   2760
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   3600
         Width           =   945
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   7215
      _cx             =   12726
      _cy             =   4895
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   8438015
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmchemical.frx":0E57
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
      Left            =   5280
      Top             =   0
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
            Picture         =   "frmchemical.frx":0EDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchemical.frx":1277
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchemical.frx":1611
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchemical.frx":22EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchemical.frx":273D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchemical.frx":2EF7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Active Ingredients"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   19
      Top             =   4920
      Width           =   1875
   End
End
Attribute VB_Name = "frmchemical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsF As New ADODB.Recordset

Private Sub cboChemical_LostFocus()
On Error GoTo err
   
   cboChemical.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblqmschemicalhdr where chemicalId='" & cboChemical.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
   If rs.EOF <> True Then
   fillcontroll cboChemical.BoundText
   
   Else
   MsgBox "Record Not Found."
   End If
   rs.Close
   Exit Sub
err:
   MsgBox err.Description
   'rs.Close
End Sub
Private Sub fillcontroll(id As String)
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmschemicalhdr where chemicalId = '" & id & "'", MHVDB
If rs.EOF <> True Then
   txttradename.Text = rs!tradename
   txtapproveddate.Value = Format(rs!dateapproved, "yyyy-MM-dd")
   txtrecomendeddose.Text = IIf(rs!recomendedDoseRate = 0, "", rs!recomendedDoseRate)
   txtsafereentry.Text = IIf(rs!safeReEntry = 0, "", rs!safeReEntry)
   txtpreharvest.Text = IIf(rs!preHarvest = 0, "", rs!preHarvest)
   txtuses.Text = IIf(rs!uses = 0, "", rs!uses)
   txtremarks.Text = rs!remarks
  
End If
fillgrd id
End Sub
Private Sub fillgrd(id As String)
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
rs.Open "select * from tblqmschemicaldetails where chemicalId = '" & id & "'", MHVDB
If rs.EOF <> True Then
i = 1
    Do While rs.EOF <> True
   Mygrid.TextMatrix(i, 0) = i
   FindqmsChemical rs!ingredientid
   Mygrid.TextMatrix(i, 1) = qmsChemical
   Mygrid.TextMatrix(i, 2) = rs!percentage
        
        
        
        i = i + 1
        rs.MoveNext
    Loop
End If
End Sub
Private Sub Form_Load()
On Error GoTo err
Operation = ""

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString



Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select concat(cast(chemicalId as char), '  ', tradename) as description,chemicalid  from tblqmschemicalhdr order by convert(chemicalid,unsigned integer)", db
Set cboChemical.RowSource = rsF
cboChemical.ListField = "description"
cboChemical.BoundColumn = "chemicalid"

Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Mygrid_Click()
If Mygrid.col = 1 And Len(Mygrid.TextMatrix(Mygrid.row - 1, 2)) > 0 Then
Mygrid.Editable = flexEDKbdMouse
FillGridCombo
Else
Mygrid.ComboList = ""
Mygrid.Editable = flexEDNone
End If

If Len(Mygrid.TextMatrix(Mygrid.row, 1)) > 0 And Mygrid.col = 2 Then
Mygrid.Editable = flexEDKbdMouse
End If


End Sub

Private Sub Mygrid_LeaveCell()
Dim i As Integer
If Trim(Mygrid.TextMatrix(Mygrid.row, 1)) = "" Then
Mygrid.RemoveItem (Mygrid.row)
Mygrid.Rows = Mygrid.Rows + 1
For i = 1 To Mygrid.Rows - 1
If Len(Mygrid.TextMatrix(i, 1)) = 0 Then Exit For
Mygrid.TextMatrix(i, 0) = i
Next

End If

End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
            
        cboChemical.Enabled = False
        TB.Buttons(3).Enabled = True
        Operation = "ADD"
        CLEARCONTROLL
        Dim rs As New ADODB.Recordset
        Set rs = Nothing
        rs.Open "SELECT MAX(convert(chemicalid ,unsigned integer))+1 AS MaxID from tblqmschemicalhdr", MHVDB, adOpenForwardOnly, adLockOptimistic
        If rs.EOF <> True Then
        cboChemical.Text = IIf(IsNull(rs!MaxID), 1, rs!MaxID)
        Else
        cboChemical.Text = rs!MaxID
        End If
       Case "OPEN"
        Operation = "OPEN"
        CLEARCONTROLL
        cboChemical.Enabled = True
        TB.Buttons(3).Enabled = True
             
       Case "SAVE"
        MNU_SAVE

       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub CLEARCONTROLL()
   
   txttradename.Text = ""
   txtrecomendeddose.Text = ""
   txtsafereentry.Text = ""
   txtpreharvest.Text = ""
   txtuses.Text = ""
   txtremarks.Text = ""
   txtapproveddate.Value = "01/01/1900"
  Mygrid.Clear
Mygrid.FormatString = "^Sl.No.|^Active Ingredient Name|^%|^"
Mygrid.ColWidth(0) = 960
Mygrid.ColWidth(1) = 3790
Mygrid.ColWidth(2) = 1320
Mygrid.ColWidth(3) = 1005
End Sub
Private Sub MNU_SAVE()
Dim rs As New ADODB.Recordset
On Error GoTo err
Dim i As Integer
If Len(txttradename.Text) = 0 Then
MsgBox "Input the Trade name."
Exit Sub
End If

MHVDB.BeginTrans
If Operation = "ADD" Then
LogRemarks = ""
MHVDB.Execute "INSERT INTO tblqmschemicalhdr (chemicalId,tradename,dateapproved,recomendedDoseRate," _
            & "safeReEntry,preHarvest,uses,location,remarks,status) " _
            & "VALUEs('" & cboChemical.Text & "','" & txttradename.Text & "','" & Format(txtapproveddate.Value, "yyyy-MM-dd") & "', " _
            & " '" & Val(txtrecomendeddose.Text) & "','" & Val(txtsafereentry.Text) & "','" & Val(txtpreharvest.Text) & "', " _
            & "'" & txtuses.Text & "','" & Mlocation & "','" & txtremarks.Text & "','ON')"
 
 
LogRemarks = "Inserted new record" & cboChemical.BoundText & "," & txttradename.Text & "," & Mlocation & "," & txtremarks
updatemhvlog Now, MUSER, LogRemarks, ""

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblqmschemicalhdr set tradename='" & txttradename.Text & "' " _
            & ",dateapproved='" & Format(txtapproveddate.Value, "yyyy-MM-dd") & "',remarks='" & txtremarks.Text & "' " _
            & ",recomendedDoseRate='" & Val(txtrecomendeddose.Text) & "' " _
            & ",safeReEntry='" & Val(txtsafereentry.Text) & "',preHarvest='" & Val(txtpreharvest.Text) & "',uses='" & txtuses.Text & "' " _
            & " where chemicalid='" & cboChemical.BoundText & "' and location='" & Mlocation & "'"

LogRemarks = "Updated  record" & cboChemical.BoundText & "," & txttradename.Text & "," & Mlocation
updatemhvlog Now, MUSER, LogRemarks, ""
Else
MsgBox "OPERATION NOT SELECTED."
Exit Sub
End If

MHVDB.Execute "delete from tblqmschemicaldetails where chemicalId='" & cboChemical.BoundText & "'"

For i = 1 To Mygrid.Rows - 1
If Len(Mygrid.TextMatrix(i, 1)) = 0 Then Exit For
MHVDB.Execute "insert into tblqmschemicaldetails (ingredientid,chemicalid,percentage,location) values" _
            & "('" & Mid(Mygrid.TextMatrix(i, 1), 1, 5) & "','" & cboChemical.BoundText & "' " _
            & ", '" & Val(Mygrid.TextMatrix(i, 2)) & "','" & Mlocation & "')"

Next



MHVDB.CommitTrans
TB.Buttons(3).Enabled = False
Exit Sub

err:
MsgBox err.Description
MHVDB.RollbackTrans


End Sub
Private Sub FillGridCombo()
        Dim RstTemp As New ADODB.Recordset
        Dim NmcItemData As String
        Dim StrComboList As String
        Dim mn1 As String
        StrComboList = "          |"
        'StrComboList = "a"
        
            Set RstTemp = Nothing
            RstTemp.Open ("select ingredientId,chemicalname,chemicalformula from tblqmsactiveingredients where status='ON' ORDER BY ingredientId"), MHVDB

            If Not RstTemp.BOF Then
               Do While Not RstTemp.EOF
                    NmcItemData = ""
                    StrComboList = IIf(StrComboList = "", Right("0000" & RstTemp("ingredientId").Value, 3) & " " & RstTemp("chemicalname").Value & " " & RstTemp("chemicalformula").Value, StrComboList & "|" & Right("0000" & RstTemp("ingredientId").Value, 3)) & " " & RstTemp("chemicalName").Value & " " & RstTemp("chemicalFormula").Value
                    RstTemp.MoveNext
               Loop
            End If
            RstTemp.Close
       Mygrid.ComboList = StrComboList



    End Sub

Private Sub txtpreharvest_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtrecomendeddose_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtremarks_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtremarks.SelStart + 1
    Dim sText As String
    sText = Left$(txtremarks.Text, iPos)
    If iPos = 1 Then GoTo Upit
    KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
    If iPos > 1 And _
     (InStr(iPos - 1, sText, " ") > 0 Or _
      InStr(iPos - 1, sText, "-") > 0 Or _
      InStr(iPos - 1, sText, ".") > 0 Or _
      InStr(iPos - 1, sText, "'") > 0) _
      Then GoTo Upit
    If iPos > 2 Then _
      If InStr(iPos - 2, sText, "Mc") > 0 _
        Then GoTo Upit
        
   End If
  Exit Sub
Upit:
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtsafereentry_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txttradename_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txttradename.SelStart + 1
    Dim sText As String
    sText = Left$(txttradename.Text, iPos)
    If iPos = 1 Then GoTo Upit
    KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
    If iPos > 1 And _
     (InStr(iPos - 1, sText, " ") > 0 Or _
      InStr(iPos - 1, sText, "-") > 0 Or _
      InStr(iPos - 1, sText, ".") > 0 Or _
      InStr(iPos - 1, sText, "'") > 0) _
      Then GoTo Upit
    If iPos > 2 Then _
      If InStr(iPos - 2, sText, "Mc") > 0 _
        Then GoTo Upit
        
   End If
  Exit Sub
Upit:
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub txtuses_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtuses.SelStart + 1
    Dim sText As String
    sText = Left$(txtuses.Text, iPos)
    If iPos = 1 Then GoTo Upit
    KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
    If iPos > 1 And _
     (InStr(iPos - 1, sText, " ") > 0 Or _
      InStr(iPos - 1, sText, "-") > 0 Or _
      InStr(iPos - 1, sText, ".") > 0 Or _
      InStr(iPos - 1, sText, "'") > 0) _
      Then GoTo Upit
    If iPos > 2 Then _
      If InStr(iPos - 2, sText, "Mc") > 0 _
        Then GoTo Upit
        
   End If
  Exit Sub
Upit:
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub
