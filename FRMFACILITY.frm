VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMFACILITY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F A C I L I T Y    M A I N T A I N A N C E"
   ClientHeight    =   5760
   ClientLeft      =   3840
   ClientTop       =   1860
   ClientWidth     =   12435
   Icon            =   "FRMFACILITY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   12435
   Begin VB.TextBox txtgpsLng 
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
      Left            =   10920
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   12255
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
         Left            =   1920
         TabIndex        =   15
         Top             =   1440
         Width           =   7095
      End
      Begin VB.TextBox txtarea 
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
         Left            =   1920
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtwobblers 
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
         Left            =   4920
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtgpsLat 
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
         Left            =   7800
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtdescription 
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
         Left            =   7800
         TabIndex        =   4
         Top             =   480
         Width           =   4215
      End
      Begin MSDataListLib.DataCombo cboFacilityId 
         Bindings        =   "FRMFACILITY.frx":0E42
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   480
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
      Begin MSDataListLib.DataCombo cbostatus 
         Bindings        =   "FRMFACILITY.frx":0E57
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   10800
         TabIndex        =   18
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Status"
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
         Left            =   9480
         TabIndex        =   17
         Top             =   1560
         Width           =   660
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
         TabIndex        =   10
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "GPS Lng"
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
         Left            =   9480
         TabIndex        =   9
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "GPS Lat"
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
         Left            =   6360
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No. of Wobblers"
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
         Left            =   3240
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Area (Sq.Meter)"
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description"
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
         Left            =   6360
         TabIndex        =   3
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Facility Id"
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
         TabIndex        =   1
         Top             =   600
         Width           =   1020
      End
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
            Picture         =   "FRMFACILITY.frx":0E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMFACILITY.frx":1206
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMFACILITY.frx":15A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMFACILITY.frx":227A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMFACILITY.frx":26CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMFACILITY.frx":2E86
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12435
      _ExtentX        =   21934
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
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   2775
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   12255
      _cx             =   21616
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
      BackColorAlternate=   12640511
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
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FRMFACILITY.frx":3220
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
End
Attribute VB_Name = "FRMFACILITY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsF As New ADODB.Recordset

Private Sub cboFacilityId_LostFocus()
 On Error GoTo err
   
   cbofacilityid.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblqmsfacility where facilityId='" & cbofacilityid.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
   If rs.EOF <> True Then
   fillcontroll cbofacilityid.BoundText
   
   Else
   MsgBox "Record Not Found."
   End If
   rs.Close
   Exit Sub
err:
   MsgBox err.Description
   'rs.Close
End Sub

Private Sub cboFacilityId_Validate(Cancel As Boolean)
Dim rs As New ADODB.Recordset
cbofacilityid.Text = UCase(cbofacilityid.BoundText)
rs.Open "select * from tblqmsfacility where facilityId = '" & cbofacilityid.BoundText & "'", MHVDB
If rs.EOF Then
   If Operation = "OPEN" Then
      MsgBox "This code does not exists !!! "
      Cancel = True
      Exit Sub
   End If
Else
   If Operation = "ADD" Then
      MsgBox "This code already exists !!! "
      Operation = "OPEN"
   End If
fillcontroll cbofacilityid.BoundText


End If


TB.Buttons(3).Enabled = True

TB.Buttons(4).Enabled = True
End Sub
Private Sub fillcontroll(id As String)
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmsfacility where facilityId = '" & id & "'", MHVDB
If rs.EOF <> True Then
txtdescription.Text = rs!Description
   txtarea.Text = IIf(rs!Area = 0, "", rs!Area)
   txtwobblers.Text = IIf(rs!noofwobblers = 0, "", rs!noofwobblers)
   txtgpsLat.Text = IIf(rs!gpslat = 0, "", rs!gpslat)
   txtgpsLng.Text = IIf(rs!gpslng = 0, "", rs!gpslng)
   txtremarks.Text = rs!remarks
   FindqmsStatus rs!Status
   cbostatus.Text = qmsStatus
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
rsF.Open "select statusid,status  from tblqmsstatus order by status", db
Set cbostatus.RowSource = rsF
cbostatus.ListField = "status"
cbostatus.BoundColumn = "statusid"

Set rsF = Nothing

If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select concat(facilityId , '  ', description) as description,facilityId  from tblqmsfacility order by facilityId", db
Set cbofacilityid.RowSource = rsF
cbofacilityid.ListField = "description"
cbofacilityid.BoundColumn = "facilityId"
FillGrid
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub FillGrid()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^Sl.No.|^Facility Id|^Description|^Area(Sq.m)|^No. of Wobblers|^GPS Lat|^GPS Lng|^Remarks|^Status|^"
Mygrid.ColWidth(0) = 645
Mygrid.ColWidth(1) = 975
Mygrid.ColWidth(2) = 2055
Mygrid.ColWidth(3) = 1110
Mygrid.ColWidth(4) = 1650
Mygrid.ColWidth(5) = 1230
Mygrid.ColWidth(6) = 1155
Mygrid.ColWidth(7) = 1755
Mygrid.ColWidth(8) = 1065
Mygrid.ColWidth(9) = 450

rs.Open "select * from tblqmsfacility order by facilityId", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i

Mygrid.TextMatrix(i, 1) = UCase(rs!facilityid)
Mygrid.TextMatrix(i, 2) = rs!Description
Mygrid.TextMatrix(i, 3) = rs!Area
Mygrid.TextMatrix(i, 4) = rs!noofwobblers
Mygrid.TextMatrix(i, 5) = rs!gpslat
Mygrid.TextMatrix(i, 6) = rs!gpslng
Mygrid.TextMatrix(i, 7) = rs!remarks
FindqmsStatus rs!Status
Mygrid.TextMatrix(i, 7) = qmsStatus
rs.MoveNext
i = i + 1
Loop

rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
            
        cbofacilityid.Enabled = False
        TB.Buttons(3).Enabled = True
        Operation = "ADD"
        CLEARCONTROLL
        Dim rs As New ADODB.Recordset
        Set rs = Nothing
        rs.Open "SELECT MAX(SUBSTRING(facilityid,2,2))+1 AS MaxID from tblqmsfacility where old <>'Y'", MHVDB, adOpenForwardOnly, adLockOptimistic
        If rs.EOF <> True Then
        cbofacilityid.Text = "F" & IIf(IsNull(rs!MaxID), 10, rs!MaxID)
        Else
        cbofacilityid.Text = "F10" & rs!MaxID
        End If
       Case "OPEN"
        Operation = "OPEN"
        CLEARCONTROLL
        cbofacilityid.Enabled = True
        TB.Buttons(3).Enabled = True
             
       Case "SAVE"
        MNU_SAVE
      
        FillGrid
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub MNU_SAVE()
Dim rs As New ADODB.Recordset
On Error GoTo err
If Len(cbostatus.Text) = 0 Then
MsgBox "Select Status."
Exit Sub
End If

MHVDB.BeginTrans
If Operation = "ADD" Then
LogRemarks = ""
MHVDB.Execute "INSERT INTO tblqmsfacility (facilityId,description,area,noofwobblers,gpslat,gpslng," _
            & "location,status,remarks) " _
            & "VALUEs('" & cbofacilityid.Text & "','" & txtdescription.Text & "','" & Val(txtarea.Text) & "', " _
            & " '" & Val(txtwobblers.Text) & "','" & Val(txtgpsLat.Text) & "','" & Val(txtgpsLng.Text) & "', " _
            & "'" & Mlocation & "','" & cbostatus.BoundText & "','" & txtremarks.Text & "')"
 
 
LogRemarks = "Inserted new record" & cbofacilityid.BoundText & "," & txtdescription.Text & "," & Mlocation & "," & txtremarks
updatemhvlog Now, MUSER, LogRemarks, ""

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblqmsfacility set description='" & txtdescription.Text & "' " _
            & ",area='" & Val(txtarea.Text) & "',remarks='" & txtremarks.Text & "' " _
            & ",noofwobblers='" & Val(txtwobblers.Text) & "' " _
            & ",gpslat='" & Val(txtgpsLat.Text) & "',gpslng='" & Val(txtgpsLng.Text) & "',status='" & cbostatus.BoundText & "' " _
            & " where facilityId='" & cbofacilityid.BoundText & "' and location='" & Mlocation & "'"

LogRemarks = "Updated  record" & cbofacilityid.BoundText & "," & txtdescription.Text & "," & Mlocation
updatemhvlog Now, MUSER, LogRemarks, ""
Else
MsgBox "OPERATION NOT SELECTED."
End If
  
TB.Buttons(3).Enabled = False
MHVDB.CommitTrans
Exit Sub

err:
MsgBox err.Description
MHVDB.RollbackTrans


End Sub

Private Sub CLEARCONTROLL()
    txtdescription.Text = ""
   txtarea.Text = ""
   txtwobblers.Text = ""
   txtgpsLat.Text = ""
   txtgpsLng.Text = ""
   txtremarks.Text = ""

End Sub

Private Sub txtarea_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtdescription_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtdescription.SelStart + 1
    Dim sText As String
    sText = Left$(txtdescription.Text, iPos)
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

Private Sub txtgpsLat_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtgpsLng_KeyPress(KeyAscii As Integer)
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

Private Sub txtwobblers_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
