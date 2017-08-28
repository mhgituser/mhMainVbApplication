VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMPLANTBATCH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P L A N T  B A T C H . . . "
   ClientHeight    =   7725
   ClientLeft      =   3735
   ClientTop       =   1650
   ClientWidth     =   16335
   Icon            =   "FRMPLANTBATCH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   16335
   Begin VB.TextBox txticedamaged 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txtoversize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txtundersize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txtweakplant 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txthealthyplant 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox txttotalplant 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdboxdetail 
      Caption         =   "Box Detail..."
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
      Left            =   11040
      Picture         =   "FRMPLANTBATCH.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Plant Batch Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   12855
      Begin VSFlex7UCtl.VSFlexGrid myGrid 
         Height          =   5055
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   12615
         _cx             =   22251
         _cy             =   8916
         _ConvInfo       =   1
         Appearance      =   1
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
         BackColor       =   12648447
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   16777215
         BackColorBkg    =   14737632
         BackColorAlternate=   16761024
         GridColor       =   14737632
         GridColorFixed  =   14737632
         TreeColor       =   16777215
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   1500
         ExtendLastCol   =   0   'False
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   12855
      Begin VB.TextBox txtnoofboxes 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   11880
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker txtentrydate 
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   102170625
         CurrentDate     =   41477
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "FRMPLANTBATCH.frx":170C
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSDataListLib.DataCombo cbostaffid 
         Bindings        =   "FRMPLANTBATCH.frx":1721
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No. of Boxes"
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
         Left            =   10440
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Staff  Id"
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
         Left            =   6000
         TabIndex        =   8
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Shipment No."
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
         TabIndex        =   3
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date"
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
         Left            =   3600
         TabIndex        =   2
         Top             =   360
         Width           =   510
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
            Picture         =   "FRMPLANTBATCH.frx":1736
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMPLANTBATCH.frx":1AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMPLANTBATCH.frx":1E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMPLANTBATCH.frx":2B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMPLANTBATCH.frx":2F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMPLANTBATCH.frx":3750
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   16335
      _ExtentX        =   28813
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total"
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
      TabIndex        =   17
      Top             =   6840
      Width           =   555
   End
End
Attribute VB_Name = "FRMPLANTBATCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim maxPlantBatch As Integer
Dim maxPlantBatchloop As Integer
Dim totplant, healthyplant, weakplant, oversize, undersize, icedamaged As Long

Private Sub cbotrnid_LostFocus()
On Error GoTo err
Dim rs As New ADODB.Recordset
Set rs = Nothing
cbotrnid.Enabled = False
TB.buttons(3).Enabled = True
rs.Open "select * from tblqmsplantbatchhdr where trnid='" & cbotrnid.BoundText & "' order by trnid", MHVDB

If rs.EOF <> True Then

txtentrydate.Value = Format(rs!entrydate, "yyyy-MM-dd")
FindsTAFF rs!staffid
cbostaffid.Text = rs!staffid & " " & sTAFF
txtnoofboxes.Text = IIf(rs!noofboxes = 0, "", rs!noofboxes)
filldetail rs!trnid

End If

TB.buttons(3).Enabled = True




Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub addcells()
Dim i As Integer
totplant = 0
healthyplant = 0
weakplant = 0
oversize = 0
undersize = 0
icedamaged = 0
For i = 1 To mygrid.rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
totplant = totplant + Val(mygrid.TextMatrix(i, 4))
healthyplant = healthyplant + Val(mygrid.TextMatrix(i, 5))
weakplant = weakplant + Val(mygrid.TextMatrix(i, 6))
undersize = undersize + Val(mygrid.TextMatrix(i, 7))
oversize = oversize + Val(mygrid.TextMatrix(i, 8))
icedamaged = icedamaged + Val(mygrid.TextMatrix(i, 9))
Next
txttotalplant.Text = Format(totplant, "#,##,##")
txthealthyplant.Text = IIf(Format(healthyplant, "#,##,##") = 0, "", Format(healthyplant, "#,##,##"))
txtweakplant.Text = IIf(Format(weakplant, "#,##,##") = 0, "", Format(weakplant, "#,##,##"))
txtundersize.Text = IIf(Format(undersize, "#,##,##") = 0, "", Format(undersize, "#,##,##"))
txtoversize.Text = IIf(Format(oversize, "#,##,##") = 0, "", Format(oversize, "#,##,##"))
txticedamaged.Text = IIf(Format(icedamaged, "#,##,##") = 0, "", Format(icedamaged, "#,##,##"))
End Sub
Private Sub filldetail(trnid As Integer)
On Error GoTo err

Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing

mygrid.Clear
'mygrid.FormatString = "^Sl.No.|^Plant Variety|^Batch No.|^Tissue Culture|^B/L Shipment Size|^Healthy Plant|^Weak Plant|^Under Size|^Over Size|^ Damaged|^TC source|"
mygrid.FormatString = pbStringFormat
'mygrid.ColWidth(0) = 570
'mygrid.ColWidth(1) = 1380
'mygrid.ColWidth(2) = 1005
'mygrid.ColWidth(3) = 1350
'mygrid.ColWidth(4) = 1710
'mygrid.ColWidth(5) = 1350
'mygrid.ColWidth(6) = 1155
'mygrid.ColWidth(7) = 1080
'mygrid.ColWidth(8) = 1005
'mygrid.ColWidth(9) = 1320
'mygrid.ColWidth(10) = 1155
'mygrid.ColWidth(11) = 120

mygrid.ColWidth(0) = 570
mygrid.ColWidth(1) = 1380
mygrid.ColWidth(2) = 1005
mygrid.ColWidth(3) = 1350
mygrid.ColWidth(4) = 1710
mygrid.ColWidth(5) = 1350
mygrid.ColWidth(6) = 1000
mygrid.ColWidth(7) = 1080
mygrid.ColWidth(8) = 900
mygrid.ColWidth(9) = 1100
mygrid.ColWidth(10) = 1000
mygrid.ColWidth(11) = 50

rs.Open "select * from tblqmsplantbatchdetail where trnid='" & trnid & "' order by trnid", MHVDB
i = 1
Do While rs.EOF <> True
mygrid.TextMatrix(i, 0) = i
FindqmsPlantVariety (Right("0000" & rs!plantvariety, 3))
mygrid.TextMatrix(i, 1) = Right("0000" & rs!plantvariety, 3) & " " & qmsPlantVariety
mygrid.TextMatrix(i, 2) = rs!plantBatch
FindqmsPlanttype rs!planttype
mygrid.TextMatrix(i, 3) = rs!planttype & " " & qmsPlantType
mygrid.TextMatrix(i, 4) = IIf(rs!shipmentsize = 0, "", rs!shipmentsize)
mygrid.TextMatrix(i, 5) = IIf(rs!healthyplant = 0, "", rs!healthyplant)
mygrid.TextMatrix(i, 6) = IIf(rs!weakplant = 0, "", rs!weakplant)
mygrid.TextMatrix(i, 7) = IIf(rs!undersize = 0, "", rs!undersize)
mygrid.TextMatrix(i, 8) = IIf(rs!oversize = 0, "", rs!oversize)
mygrid.TextMatrix(i, 9) = IIf(rs!icedamaged = 0, "", rs!icedamaged)
mygrid.TextMatrix(i, 10) = IIf(rs!tcsource = 0, "", rs!tcsource)


i = i + 1
rs.MoveNext
Loop
addcells
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub cmdboxdetail_Click()
If Len(cbotrnid.Text) = 0 Then
MsgBox "Select The Shipment No. First."
Exit Sub
End If


If TB.buttons(3).Enabled = True Then
MsgBox "Save the Plant Batch First."

Else
frmboxdetails.txtshipmentno = cbotrnid.Text
frmboxdetails.txtentrydate.Value = Format(txtentrydate.Value, "yyyy-MM-dd")
frmboxdetails.txtstaffid.Text = cbostaffid.Text
frmboxdetails.Show 1

End If
End Sub

Private Sub Form_Load()
On Error GoTo err
Operation = ""
Dim rsF As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
maxPlantBatch = 0
Set rsF = Nothing

If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select trnid  as description,trnid  from tblqmsplantbatchhdr order by trnid", db
Set cbotrnid.RowSource = rsF
cbotrnid.ListField = "description"
cbotrnid.BoundColumn = "trnid"

Set rsF = Nothing

If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff where nursery='1' order by STAFFCODE", db
Set cbostaffid.RowSource = rsF
cbostaffid.ListField = "STAFFNAME"
cbostaffid.BoundColumn = "STAFFCODE"

CLEARCONTROLL

Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub FillGridComboVariety()
        Dim RstTemp As New ADODB.Recordset
        Dim NmcItemData As String
        Dim StrComboList As String
        Dim mn1 As String
        StrComboList = "          |"
        'StrComboList = "a"
        
            Set RstTemp = Nothing
            RstTemp.Open ("select varietyid,description from tblqmsplantvariety where status='ON' ORDER BY varietyid"), MHVDB

            If Not RstTemp.BOF Then
               Do While Not RstTemp.EOF
                    NmcItemData = ""
                    StrComboList = IIf(StrComboList = "", RstTemp("varietyid").Value & " " & RstTemp("description").Value, StrComboList & "|" & RstTemp("varietyid").Value) & " " & RstTemp("description").Value
                    RstTemp.MoveNext
               Loop
            End If
            RstTemp.Close
       mygrid.ComboList = StrComboList



    End Sub
    
    Private Sub FillGridComboTC()
        Dim RstTemp As New ADODB.Recordset
        Dim NmcItemData As String
        Dim StrComboList As String
        Dim mn1 As String
        StrComboList = "          |"
        'StrComboList = "a"
        
            Set RstTemp = Nothing
            RstTemp.Open ("select planttypeid,description from tblqmsplanttype where status='ON' ORDER BY planttypeid"), MHVDB

            If Not RstTemp.BOF Then
               Do While Not RstTemp.EOF
                    NmcItemData = ""
                    StrComboList = IIf(StrComboList = "", RstTemp("planttypeid").Value & " " & RstTemp("description").Value, StrComboList & "|" & RstTemp("planttypeid").Value) & " " & RstTemp("description").Value
                    RstTemp.MoveNext
               Loop
            End If
            RstTemp.Close
       mygrid.ComboList = StrComboList



    End Sub

Private Sub mygrid_Click()
If mygrid.col = 1 And (Len(mygrid.TextMatrix(mygrid.row - 1, 4)) > 0 Or mygrid.TextMatrix(mygrid.row - 1, 10) = "SUCK") Then
mygrid.Editable = flexEDKbdMouse
FillGridComboVariety
ElseIf mygrid.col = 3 And Len(mygrid.TextMatrix(mygrid.row, 2)) > 0 Then
mygrid.Editable = flexEDKbdMouse
FillGridComboTC
ElseIf (mygrid.col = 4 Or mygrid.col = 5 Or mygrid.col = 6 Or mygrid.col = 7 Or mygrid.col = 8 Or mygrid.col = 9) And Len(mygrid.TextMatrix(mygrid.row, 3)) > 0 Then
mygrid.ComboList = ""
mygrid.Editable = flexEDKbdMouse
ElseIf (mygrid.col = 10 And Len(mygrid.TextMatrix(mygrid.row - 1, 2)) > 0 And Len(mygrid.TextMatrix(mygrid.row - 1, 4)) > 0) Then
mygrid.Editable = flexEDKbdMouse
FillGridComboTCsource
Else
mygrid.ComboList = ""
mygrid.Editable = flexEDNone
End If

addcells
End Sub

Private Sub Mygrid_LeaveCell()

Dim i As Integer
Dim j As Integer

If mygrid.col = 1 And Len(mygrid.TextMatrix(mygrid.row, 1)) > 0 And Val(mygrid.TextMatrix(mygrid.row, 2)) = 0 And Operation = "ADD" Then
mygrid.TextMatrix(mygrid.row, 2) = maxPlantBatch
maxPlantBatch = maxPlantBatch + 1
End If

If Trim(mygrid.TextMatrix(mygrid.row, 1)) = "" And Operation = "ADD" Then
mygrid.RemoveItem (mygrid.row)
j = maxPlantBatchloop
mygrid.rows = mygrid.rows + 1
For i = 1 To mygrid.rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
mygrid.TextMatrix(i, 0) = i
mygrid.TextMatrix(i, 2) = j
j = j + 1
Next
End If

If Trim(mygrid.TextMatrix(mygrid.row, 1)) = "" And Operation = "OPEN" Then
mygrid.RemoveItem (mygrid.row)
mygrid.rows = mygrid.rows + 1
For i = 1 To mygrid.rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
mygrid.TextMatrix(i, 0) = i
Next
End If

addcells
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

       Case "ADD"
        Dim rsPb As New ADODB.Recordset
        Dim rs As New ADODB.Recordset
        Set rsPb = Nothing
        rsPb.Open "Select max(plantbatch) as max from tblqmsplantbatchdetail", MHVDB
        If rsPb.EOF <> True Then
        maxPlantBatch = rsPb!max + 1
        maxPlantBatchloop = rsPb!max + 1
        End If
        Set rs = Nothing
        rs.Open "select max(trnid) as max from tblqmsplantbatchhdr", MHVDB
        If rs.EOF <> True Then
        cbotrnid.Text = rs!max + 1
        End If
        
        
        cbotrnid.Enabled = False
        TB.buttons(3).Enabled = True
        Operation = "ADD"
        boxOperation = "ADD"
        CLEARCONTROLL
        Case "OPEN"
        Operation = "OPEN"
        boxOperation = "OPEN"
        cbotrnid.Text = ""
        CLEARCONTROLL
        cbotrnid.Enabled = True
        TB.buttons(3).Enabled = False
             
       Case "SAVE"
        MNU_SAVE
        
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
End Select
End Sub
Private Sub MNU_SAVE()
Dim rs As New ADODB.Recordset
Dim i As Integer
On Error GoTo err
If Len(cbotrnid.Text) = 0 Then
MsgBox "Invalid Operation."
Exit Sub
End If

If Operation = "ADD" And Val(txtnoofboxes.Text) = 0 Then
MsgBox "Input Total No. of Boxes In the Shipment " & cbotrnid.Text
Exit Sub
End If

MHVDB.BeginTrans
If Operation = "ADD" Then
LogRemarks = ""
MHVDB.Execute "insert into tblqmsplantbatchhdr(trnid,entrydate,shipmentno,staffid,status,location,noofboxes)" _
             & "values('" & cbotrnid.Text & "','" & Format(txtentrydate.Value, "yyyy-MM-dd") & "','" & cbotrnid.Text & "', " _
             & "'" & cbostaffid.BoundText & "','ON','" & Mlocation & "','" & Val(txtnoofboxes.Text) & "')"
 
LogRemarks = "Inserted new record" & cbotrnid.BoundText & "," & Format(txtentrydate.Value, "yyyy-MM-dd") & "," & Mlocation & "," & txtremarks
updatemhvlog Now, MUSER, LogRemarks, ""

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblqmsplantbatchhdr" _
             & " set entrydate='" & Format(txtentrydate.Value, "yyyy-MM-dd") & "', " _
             & " staffid='" & cbostaffid.BoundText & "',noofboxes='" & Val(txtnoofboxes.Text) & "' where trnid='" & cbotrnid.BoundText & "' and location='" & Mlocation & "'"
 

LogRemarks = "Updated  record" & cbotrnid.BoundText & "," & Format(txtentrydate.Value, "yyyy-MM-dd") & "," & Mlocation
updatemhvlog Now, MUSER, LogRemarks, ""
Else
MsgBox "OPERATION NOT SELECTED."
MHVDB.RollbackTrans
Exit Sub
End If
MHVDB.Execute "delete from tblqmsplantbatchdetail where trnid='" & cbotrnid.BoundText & "'"
For i = 1 To mygrid.rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
MHVDB.Execute "insert into tblqmsplantbatchdetail (trnid,entrydate,plantvariety,plantbatch,planttype," _
            & "shipmentsize,healthyplant,weakplant,undersize,oversize,icedamaged,location,tcsid) values" _
            & "('" & cbotrnid.BoundText & "','" & Format(txtentrydate.Value, "yyyy-MM-dd") & "', " _
            & " '" & Mid(mygrid.TextMatrix(i, 1), 1, 3) & "','" & mygrid.TextMatrix(i, 2) & "','" & Mid(mygrid.TextMatrix(i, 3), 1, 3) & "'" _
            & ",'" & Val(mygrid.TextMatrix(i, 4)) & "','" & Val(mygrid.TextMatrix(i, 5)) & "'" _
            & ",'" & Val(mygrid.TextMatrix(i, 6)) & "','" & Val(mygrid.TextMatrix(i, 7)) & "'" _
            & ",'" & Val(mygrid.TextMatrix(i, 8)) & "','" & Val(mygrid.TextMatrix(i, 9)) & "','" & Mlocation & "','" & Mid(mygrid.TextMatrix(i, 10), 1, 4) & "')"

Next



 TB.buttons(3).Enabled = False
MHVDB.CommitTrans
Exit Sub

err:
MsgBox err.Description
MHVDB.RollbackTrans
End Sub

Private Sub CLEARCONTROLL()
mygrid.Clear
'mygrid.FormatString = "^Sl.No.|^Plant Variety|^Batch No.|^Tissue Culture|^B/L Shipment Size|^Healthy Plant|^Weak Plant|^Under Size|^Over Size|^Ice Damaged|^"
mygrid.FormatString = pbStringFormat
mygrid.ColWidth(0) = 570
mygrid.ColWidth(1) = 1380
mygrid.ColWidth(2) = 1005
mygrid.ColWidth(3) = 1350
mygrid.ColWidth(4) = 1710
mygrid.ColWidth(5) = 1350
mygrid.ColWidth(6) = 1000
mygrid.ColWidth(7) = 1080
mygrid.ColWidth(8) = 900
mygrid.ColWidth(9) = 1100
mygrid.ColWidth(10) = 1000
mygrid.ColWidth(11) = 50

cbostaffid.Text = ""
txtnoofboxes.Text = ""
txtentrydate.Value = Format(Now, "dd/MM/yyyy")
End Sub
Private Sub FillGridComboTCsource()
        Dim RstTemp As New ADODB.Recordset
        Dim NmcItemData As String
        Dim StrComboList As String
        Dim mn1 As String
        StrComboList = "          |"
            Set RstTemp = Nothing
            RstTemp.Open ("select tcsid,tcsourcename from tblqmstcsource order by tcsid"), MHVDB
            If Not RstTemp.BOF Then
               Do While Not RstTemp.EOF
                     StrComboList = IIf(StrComboList = "", RstTemp("tcsourcename").Value, StrComboList & "|" & RstTemp("tcsourcename").Value)
                     RstTemp.MoveNext
               Loop
            End If
            RstTemp.Close
       mygrid.ComboList = StrComboList
End Sub










