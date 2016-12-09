VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMmediumbatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P L A N T I N G  ME D I U M  B A T C H . . ."
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13350
   Icon            =   "FRMmediumbatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   13350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txttotplants 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox txtdesc 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   7080
      Width           =   5775
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   0
      TabIndex        =   22
      Top             =   2640
      Width           =   13335
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   4095
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   12975
         _cx             =   22886
         _cy             =   7223
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
         Rows            =   1
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FRMmediumbatch.frx":0E42
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
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   13335
      Begin VB.TextBox txtcomments 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6480
         TabIndex        =   21
         Top             =   1200
         Width           =   6735
      End
      Begin VB.TextBox txtplantsno 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   11880
         TabIndex        =   17
         Top             =   720
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker txtstartdate 
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   76939265
         CurrentDate     =   41480
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "FRMmediumbatch.frx":0F9E
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker txtstarttime 
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   76939266
         CurrentDate     =   41480
      End
      Begin MSComCtl2.DTPicker txtenddate 
         Height          =   375
         Left            =   9000
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   76939265
         CurrentDate     =   41480
      End
      Begin MSComCtl2.DTPicker txtendtime 
         Height          =   375
         Left            =   11880
         TabIndex        =   11
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   76939266
         CurrentDate     =   41480
      End
      Begin MSDataListLib.DataCombo cboplantBatch 
         Bindings        =   "FRMmediumbatch.frx":0FB3
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1200
         TabIndex        =   13
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   "ItemCode"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cbomediummix 
         Bindings        =   "FRMmediumbatch.frx":0FC8
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   6480
         TabIndex        =   15
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   "ItemCode"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cbostaff 
         Bindings        =   "FRMmediumbatch.frx":0FDD
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1200
         TabIndex        =   19
         Top             =   1200
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   "ItemCode"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   20
         Top             =   1320
         Width           =   870
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Staff Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "No.of Plants"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10440
         TabIndex        =   16
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Medium Mix No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   14
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Plant Batch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Finish Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10440
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Finish Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7920
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Start Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   4
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trn. Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   585
      End
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
            Picture         =   "FRMmediumbatch.frx":0FF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMmediumbatch.frx":138C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMmediumbatch.frx":1726
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMmediumbatch.frx":2400
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMmediumbatch.frx":2852
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMmediumbatch.frx":300C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13350
      _ExtentX        =   23548
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
Attribute VB_Name = "FRMmediumbatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim noofplants As Double
Private Sub cbotrnid_LostFocus()
 On Error GoTo err
   
   cbotrnid.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblqmsmediumbatch where trnid='" & cbotrnid.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
   If rs.EOF <> True Then
   fillcontroll cbotrnid.BoundText
   
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
rs.Open "select * from tblqmsmediumbatch where trnid = '" & id & "'", MHVDB
If rs.EOF <> True Then
    txtstartdate.Value = Format(rs!startdate, "dd/MM/yyyy")
    txtenddate.Value = Format(rs!enddate, "dd/MM/yyyy")
    txtstarttime.Value = Format(rs!starttime, "HH:mm:ss")
    txtendtime.Value = Format(rs!endtime, "HH:mm:ss")
    findQmsBatchDetail rs!plantBatch
    cboplantbatch.Text = qmsBatchdetail1
    cbomediummix.Text = rs!mediummixno
    txtplantsno.Text = rs!noofplants
    FindsTAFF rs!staffid
    cbostaff.Text = rs!staffid & " " & sTAFF
    txtcomments.Text = rs!Comments
    
    
    
 
End If
End Sub

Private Sub Form_Load()
On Error GoTo err
operation = ""
Dim rsF As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
txtstartdate.Value = Format(Now, "dd/MM/yyyy")
txtstarttime.Value = Format(Now, "HH:mm:ss")
txtenddate.Value = Format(Now, "dd/MM/yyyy")
txtendtime.Value = Format(Now, "HH:mm:ss")
Set rsF = Nothing

If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select trnId as description  from tblqmsmediumbatch order by trnId", db
Set cbotrnid.RowSource = rsF
cbotrnid.ListField = "description"
cbotrnid.BoundColumn = "description"



Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "SELECT concat(cast(plantbatch as char),' ', c.description,' ', b.description) as description,plantbatch FROM  `tblqmsplantbatchdetail` a," _
 & "tblqmsplanttype b, tblqmsplantvariety c Where planttype = planttypeid" _
 & " AND plantvariety = varietyid order by plantbatch", db
Set cboplantbatch.RowSource = rsF
cboplantbatch.ListField = "description"
cboplantbatch.BoundColumn = "plantbatch"




Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
If rsF.State = adStateOpen Then Srs.Close
rsF.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff WHERE nursery='1'  order by STAFFCODE", db
Set cbostaff.RowSource = rsF
cbostaff.ListField = "STAFFNAME"
cbostaff.BoundColumn = "STAFFCODE"

Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
If rsF.State = adStateOpen Then Srs.Close
rsF.Open "select trnid  from tblqmsmediummixhdr   order by trnid", db
Set cbomediummix.RowSource = rsF
cbomediummix.ListField = "trnid"
cbomediummix.BoundColumn = "trnid"



FillGrid

Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Mygrid_Click()
txtdesc.Text = Mygrid.TextMatrix(Mygrid.row, Mygrid.col)
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
            
        cbotrnid.Enabled = False
        TB.Buttons(3).Enabled = True
        operation = "ADD"
        CLEARCONTROLL
        Dim rs As New ADODB.Recordset
        Set rs = Nothing
        rs.Open "SELECT MAX(trnid )+1 AS MaxID from tblqmsmediumbatch", MHVDB, adOpenForwardOnly, adLockOptimistic
        If rs.EOF <> True Then
        cbotrnid.Text = IIf(IsNull(rs!MaxId), 99, rs!MaxId)
        Else
        cbotrnid.Text = rs!MaxId
        End If
       Case "OPEN"
        operation = "OPEN"
        CLEARCONTROLL
        cbotrnid.Enabled = True
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
If Len(cbotrnid.Text) = 0 Then
MsgBox "Transaction No. is Required."
Exit Sub
End If

MHVDB.BeginTrans
If operation = "ADD" Then
LogRemarks = ""
MHVDB.Execute "INSERT INTO tblqmsmediumbatch (trnid,startdate,starttime,enddate,endtime,plantbatch," _
            & "mediummixno,noofplants,comments,staffid,status,location)" _
            & "VALUEs('" & cbotrnid.Text & "','" & Format(txtstartdate.Value, "yyyy-MM-dd") & "'," _
            & "'" & Format(txtstarttime.Value, "HH:mm:ss") & "','" & Format(txtenddate.Value, "yyyy-MM-dd") & "', " _
            & "'" & Format(txtendtime.Value, "HH:mm:ss") & "','" & cboplantbatch.BoundText & "'," _
            & " '" & cbomediummix.BoundText & "','" & Val(txtplantsno.Text) & "','" & txtcomments.Text & "', " _
            & "'" & cbostaff.BoundText & "','ON','" & Mlocation & "')"
 
 
LogRemarks = "Inserted new record" & cbotrnid.BoundText & "," & Mlocation & "," & txtremarks
updatemhvlog Now, MUSER, LogRemarks, ""

ElseIf operation = "OPEN" Then
MHVDB.Execute "update tblqmsmediumbatch set " _
            & "startdate='" & Format(txtstartdate.Value, "yyyy-MM-dd") & "'," _
            & "starttime='" & Format(txtstarttime.Value, "HH:mm:ss") & "'," _
            & "enddate='" & Format(txtenddate.Value, "yyyy-MM-dd") & "'," _
            & "endtime='" & Format(txtendtime.Value, "HH:mm:ss") & "'," _
            & "plantbatch='" & cboplantbatch.BoundText & "'," _
            & "mediummixno='" & cbomediummix.BoundText & "'," _
            & "noofplants='" & Val(txtplantsno.Text) & "'," _
            & "comments='" & txtcomments.Text & "'" _
            & " where trnid='" & cbotrnid.BoundText & "' and location='" & Mlocation & "'"

LogRemarks = "Updated  record" & cbotrnid.BoundText & "," & Mlocation
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
txtstartdate.Value = Format(Now, "dd/MM/yyyy")
txtstarttime.Value = Format(Now, "HH:mm:ss")
txtenddate.Value = Format(Now, "dd/MM/yyyy")
txtendtime.Value = Format(Now, "HH:mm:ss")
   cbostaff.Text = ""
   cboplantbatch.Text = ""
   cbomediummix.Text = ""
   txtplantsno.Text = ""
   txtcomments.Text = ""

End Sub

Private Sub txtstarttime_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
txtstarttime.Value = Format(txtstarttime.Value, "HH:mm:ss")

End Sub
Private Sub FillGrid()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
noofplants = 0
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^Sl.No.|^Trn. Id|^S.Date|^S.Time|^F.Date|^F.Time|^P.Batch|^Mix. No.|^Plants No.|^Comments|^Staff Id|^"
Mygrid.ColWidth(0) = 960
Mygrid.ColWidth(1) = 735
Mygrid.ColWidth(2) = 1095
Mygrid.ColWidth(3) = 1005
Mygrid.ColWidth(4) = 1095
Mygrid.ColWidth(5) = 960
Mygrid.ColWidth(6) = 1110
Mygrid.ColWidth(7) = 960
Mygrid.ColWidth(8) = 960
Mygrid.ColWidth(9) = 1245
Mygrid.ColWidth(10) = 2205
Mygrid.ColWidth(11) = 210

rs.Open "select * from tblqmsmediumbatch order by trnid", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i

Mygrid.TextMatrix(i, 1) = rs!trnid
Mygrid.TextMatrix(i, 2) = Format(rs!startdate, "dd-MM-yyyy")
Mygrid.TextMatrix(i, 3) = Format(rs!starttime, "HH:mm:ss")
Mygrid.TextMatrix(i, 4) = Format(rs!enddate, "dd-MM-yyyy")
Mygrid.TextMatrix(i, 5) = Format(rs!endtime, "HH:mm:ss")
findQmsBatchDetail rs!plantBatch
Mygrid.TextMatrix(i, 6) = qmsBatchdetail1
Mygrid.ColAlignment(6) = flexAlignLeftTop
Mygrid.TextMatrix(i, 7) = rs!mediummixno
Mygrid.ColAlignment(7) = flexAlignRightTop
Mygrid.TextMatrix(i, 8) = rs!noofplants
Mygrid.ColAlignment(8) = flexAlignRightTop
Mygrid.TextMatrix(i, 9) = rs!Comments

FindsTAFF rs!staffid
Mygrid.TextMatrix(i, 10) = rs!staffid & " " & sTAFF
Mygrid.ColAlignment(10) = flexAlignLeftTop
noofplants = noofplants + IIf(IsNull(rs!noofplants), 0, rs!noofplants)
rs.MoveNext
i = i + 1
Loop

rs.Close

txttotplants.Text = noofplants
Exit Sub
err:
MsgBox err.Description

End Sub
