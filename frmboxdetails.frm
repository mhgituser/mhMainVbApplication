VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmboxdetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "B O X  D E T A I L S  [MS-GAP-11a]"
   ClientHeight    =   9015
   ClientLeft      =   2580
   ClientTop       =   1755
   ClientWidth     =   17355
   Icon            =   "frmboxdetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   17355
   Begin VB.TextBox txtavgroot 
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
      Left            =   12960
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtavgsize 
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
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtdeadplant 
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
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txticedamage 
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
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   7680
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   7680
      Width           =   735
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
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   7680
      Width           =   855
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   7680
      Width           =   975
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txttotplant 
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
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
      Left            =   6720
      Picture         =   "frmboxdetails.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
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
      Left            =   8160
      Picture         =   "frmboxdetails.frx":15EC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17175
      Begin VB.TextBox txtnoofboxes 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtstaffid 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker txtentrydate 
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   100990977
         CurrentDate     =   41478
      End
      Begin VB.TextBox txtshipmentno 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No of Boxes"
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
         Left            =   10680
         TabIndex        =   20
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label3 
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
         Left            =   5400
         TabIndex        =   5
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Entry Date"
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
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Shipment No."
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
         TabIndex        =   3
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Box Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   17175
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   6615
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   16935
         _cx             =   29871
         _cy             =   11668
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
         BackColor       =   12648447
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16761024
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
         Rows            =   100
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmboxdetails.frx":22B6
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
         Begin MSComCtl2.DTPicker dtPick 
            Height          =   315
            Left            =   1200
            TabIndex        =   11
            Top             =   840
            Visible         =   0   'False
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            Format          =   100990977
            CurrentDate     =   36473
         End
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total"
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
      Left            =   3360
      TabIndex        =   12
      Top             =   7800
      Width           =   450
   End
End
Attribute VB_Name = "frmboxdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totplant, healthyplant, weakplant, oversize, undersize, icedamaged, deadplant As Integer
Dim avgsize, avgroot As Double

Private Sub cmdsave_Click()

If Len(txtshipmentno.Text) = 0 Then
MsgBox "No Shipment No. Assigned. restart the process."
Exit Sub
End If

MHVDB.Execute "delete from tblqmsboxdetail where trnid='" & Val(txtshipmentno.Text) & "'"
For i = 1 To mygrid.Rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
MHVDB.Execute "insert into tblqmsboxdetail (trnid,entrydate,boxno,plantvariety,plantbatch,planttype," _
            & "shipmentsize,healthyplant,weakplant,undersize,oversize,icedamaged, " _
            & "deadplant,avgsize,avgroot,dateplantedbox,comments,location,status) values" _
            & "('" & Val(txtshipmentno.Text) & "','" & Format(txtentrydate.Value, "yyyy-MM-dd") & "', " _
            & " '" & mygrid.TextMatrix(i, 1) & "','" & Mid(mygrid.TextMatrix(i, 3), 1, 3) & "','" & mygrid.TextMatrix(i, 2) & "','" & Mid(mygrid.TextMatrix(i, 4), 1, 3) & "'" _
            & ",'" & Val(mygrid.TextMatrix(i, 5)) & "','" & Val(mygrid.TextMatrix(i, 6)) & "'" _
            & ",'" & Val(mygrid.TextMatrix(i, 7)) & "','" & Val(mygrid.TextMatrix(i, 8)) & "'" _
            & ",'" & Val(mygrid.TextMatrix(i, 9)) & "','" & Val(mygrid.TextMatrix(i, 10)) & "'," _
            & " '" & Val(mygrid.TextMatrix(i, 11)) & "','" & Val(mygrid.TextMatrix(i, 12)) & "'," _
            & "'" & Val(mygrid.TextMatrix(i, 13)) & "','" & Format(mygrid.TextMatrix(i, 14), "yyyy-MM-dd") & "', " _
            & "'" & mygrid.TextMatrix(i, 15) & "','" & Mlocation & "','ON')"

Next



cmdsave.Enabled = False
Operation = ""
End Sub

Private Sub Command2_Click()
Unload Me
Operation = ""
End Sub

Private Sub dtPick_Change()
mygrid.Text = dtPick.Value
End Sub

Private Sub dtPick_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
        Case vbKeyEscape
            mygrid = dtPick.Tag
            dtPick.Visible = False
        Case vbKeyReturn
            dtPick.Visible = False
    End Select
End Sub

Private Sub dtPick_LostFocus()

 dtPick.Visible = False
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmsboxdetail where trnid='" & Val(txtshipmentno.Text) & "'", MHVDB

Dim i As Integer
dtPick.Value = Format(Now, "dd/MM/yyyy")
txtshipmentno.Text = FRMPLANTBATCH.cbotrnid.Text
txtnoofboxes.Text = FRMPLANTBATCH.txtnoofboxes.Text


Operation = boxOperation
If Operation = "OPEN" Then
loaddetail Val(txtshipmentno.Text)
addcells
Else
mygrid.Rows = Val(txtnoofboxes.Text) + 1
For i = 1 To mygrid.Rows - 1
mygrid.TextMatrix(i, 0) = i
mygrid.TextMatrix(i, 1) = i

Next


End If
End Sub
Private Sub addcells()
Dim i As Integer
Dim recordcnt As Integer
totplant = 0
healthyplant = 0
weakplant = 0
oversize = 0
undersize = 0
icedamaged = 0
deadplant = 0
avgsize = 0
avgroot = 0
recordcnt = 0
For i = 1 To mygrid.Rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
totplant = totplant + Val(mygrid.TextMatrix(i, 5))
healthyplant = healthyplant + Val(mygrid.TextMatrix(i, 6))
weakplant = weakplant + Val(mygrid.TextMatrix(i, 7))
undersize = undersize + Val(mygrid.TextMatrix(i, 8))
oversize = oversize + Val(mygrid.TextMatrix(i, 9))
icedamaged = icedamaged + Val(mygrid.TextMatrix(i, 10))
deadplant = deadplant + Val(mygrid.TextMatrix(i, 11))
avgsize = avgsize + Val(mygrid.TextMatrix(i, 12))
avgroot = avgroot + Val(mygrid.TextMatrix(i, 13))
recordcnt = recordcnt + 1
Next
txttotplant.Text = Format(totplant, "#,##,##")
txthealthyplant.Text = IIf(Format(healthyplant, "#,##,##") = 0, "", Format(healthyplant, "#,##,##"))
txtweakplant.Text = IIf(Format(weakplant, "#,##,##") = 0, "", Format(weakplant, "#,##,##"))
txtundersize.Text = IIf(Format(undersize, "#,##,##") = 0, "", Format(undersize, "#,##,##"))
txtoversize.Text = IIf(Format(oversize, "#,##,##") = 0, "", Format(oversize, "#,##,##"))
txticedamage.Text = IIf(Format(icedamaged, "#,##,##") = 0, "", Format(icedamaged, "#,##,##"))
txtdeadplant.Text = IIf(Format(deadplant, "#,##,##") = 0, "", Format(icedamaged, "#,##,##"))
txtavgsize.Text = Format(avgsize / IIf(recordcnt = 0, 1, recordcnt), "##0.00")
txtavgroot.Text = Format(avgroot / IIf(recordcnt = 0, 1, recordcnt), "##0.00")
End Sub
Private Sub loaddetail(shipmentno As Integer)
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i, j As Integer
Set rs = Nothing
If Operation = "OPEN" Then
mygrid.Rows = 1
End If

mygrid.Clear
mygrid.FormatString = "^Sl.No.|^BoxNo.|^Batch No.|^Variety|^TC|^ShipmentSize|^HealthyPlant|^WeakPlant|^UnderSize|^OverSize|^IceDamaged|^DeadPlants|^Avg.Size(Cm)|^Avg.Root(Cm)|^BoxPlantedDate|^Commments|^"
mygrid.ColWidth(0) = 630
mygrid.ColWidth(1) = 675
mygrid.ColWidth(2) = 870
mygrid.ColWidth(3) = 960
mygrid.ColWidth(4) = 825
mygrid.ColWidth(5) = 1245
mygrid.ColWidth(6) = 1125
mygrid.ColWidth(7) = 1020
mygrid.ColWidth(8) = 975
mygrid.ColWidth(9) = 855
mygrid.ColWidth(10) = 1155
mygrid.ColWidth(11) = 1095
mygrid.ColWidth(12) = 1230
mygrid.ColWidth(13) = 1275
mygrid.ColWidth(14) = 1455
mygrid.ColWidth(15) = 1290
mygrid.ColWidth(16) = 135

rs.Open "select * from tblqmsboxdetail where trnid='" & shipmentno & "' order by convert(boxno,unsigned integer)", MHVDB
i = 1
If rs.EOF <> True Then
Do While rs.EOF <> True
mygrid.Rows = mygrid.Rows + 1
mygrid.TextMatrix(i, 0) = i
mygrid.TextMatrix(i, 1) = rs!boxno
mygrid.TextMatrix(i, 2) = rs!plantBatch

FindqmsPlantVariety rs!plantvariety


mygrid.TextMatrix(i, 3) = rs!plantvariety & " " & qmsPlantVariety
FindqmsPlanttype rs!planttype
mygrid.TextMatrix(i, 4) = rs!planttype & " " & qmsPlantType
mygrid.TextMatrix(i, 5) = rs!shipmentsize
mygrid.TextMatrix(i, 6) = IIf(rs!healthyplant = 0, "", rs!healthyplant)
mygrid.TextMatrix(i, 7) = IIf(rs!weakplant = 0, "", rs!weakplant)
mygrid.TextMatrix(i, 8) = IIf(rs!undersize = 0, "", rs!undersize)
mygrid.TextMatrix(i, 9) = IIf(rs!oversize = 0, "", rs!oversize)
mygrid.TextMatrix(i, 10) = IIf(rs!icedamaged = 0, "", rs!icedamaged)
mygrid.TextMatrix(i, 11) = IIf(rs!deadplant = 0, "", rs!deadplant)
mygrid.TextMatrix(i, 12) = Format(IIf(rs!avgsize = 0, "", rs!avgsize), "##0.00")
mygrid.TextMatrix(i, 13) = IIf(rs!avgroot = 0, "", rs!avgroot)
mygrid.TextMatrix(i, 14) = Format(rs!dateplantedbox, "dd/MM/yyyy")
mygrid.TextMatrix(i, 15) = rs!Comments

i = i + 1
rs.MoveNext
Loop
Else
mygrid.Rows = Val(txtnoofboxes.Text) + 1
For j = 1 To mygrid.Rows - 1
mygrid.TextMatrix(j, 0) = j
mygrid.TextMatrix(j, 1) = j

Next
End If


Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Mygrid_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
If dtPick.Visible Then Cancel = True
End Sub

Private Sub Mygrid_BeforeUserResize(ByVal row As Long, ByVal col As Long, Cancel As Boolean)
  If dtPick.Visible Then Cancel = True
End Sub

Private Sub mygrid_Click()
If mygrid.col = 14 Then

    mygrid.Editable = flexEDKbd
    mygrid.ColComboList(14) = "Dummy"
Else

    mygrid.Editable = flexEDNone

End If

If mygrid.col = 2 And Len(mygrid.TextMatrix(mygrid.row - 1, 5)) > 0 Then
    mygrid.Editable = flexEDKbdMouse
    FillGridCombobatch
ElseIf (mygrid.col = 5 Or mygrid.col = 6 Or mygrid.col = 7 Or mygrid.col = 8 Or mygrid.col = 9 Or mygrid.col = 10 Or mygrid.col = 11 Or mygrid.col = 12 Or mygrid.col = 13 Or mygrid.col = 14 Or mygrid.col = 15) And Len(Trim(mygrid.TextMatrix(mygrid.row, 4))) > 0 Then
    mygrid.ComboList = ""
    mygrid.Editable = flexEDKbdMouse

Else
    mygrid.ComboList = ""
    mygrid.Editable = flexEDNone
End If





addcells

End Sub
Private Sub FillGridCombobatch()
        Dim RstTemp As New ADODB.Recordset
        Dim NmcItemData As String
        Dim StrComboList As String
        Dim mn1 As String
        StrComboList = "          |"
        'StrComboList = "a"
        
            Set RstTemp = Nothing
            RstTemp.Open ("select distinct plantbatch from tblqmsplantbatchdetail where trnid='" & Val(txtshipmentno.Text) & "' ORDER BY plantbatch"), MHVDB

            If Not RstTemp.BOF Then
               Do While Not RstTemp.EOF
                    NmcItemData = ""
                    StrComboList = IIf(StrComboList = "", RstTemp("plantbatch").Value, StrComboList & "|" & RstTemp("plantbatch").Value)
                    RstTemp.MoveNext
               Loop
            End If
            RstTemp.Close
       mygrid.ComboList = StrComboList



    End Sub

Private Sub mygrid_DblClick()
Dim MSTR As String
Dim str As String
Dim ch As Integer
Dim row As Integer
Dim i As Integer
If mygrid.col = 1 Then
If MsgBox("Do You Want to Split The Box?", vbQuestion + vbYesNo) = vbYes Then
    MSTR = InputBox("Enter The No. of Split")
    If Not IsNumeric(MSTR) Then
    MsgBox "Invalid Input No."
    Else
    row = mygrid.row
    str = mygrid.TextMatrix(mygrid.row, 1)
    mygrid.RemoveItem mygrid.row
    
    ch = 65
     For i = 1 To MSTR
        mygrid.AddItem "", row
        
    mygrid.TextMatrix(row, 1) = str & Chr(ch)
    row = row + 1
    ch = ch + 1
    Next
    
    End If
    
   
    
 For i = 1 To mygrid.Rows - 1
mygrid.TextMatrix(i, 0) = i


Next
   
End If
End If
End Sub

Private Sub Mygrid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And Shift = 1 Then

   If mygrid.row > 0 Then
   If MsgBox("Do you want to delete this row", vbQuestion + vbYesNo) = vbYes Then
      mygrid.RemoveItem mygrid.row
      For i = 1 To mygrid.Rows - 1
        mygrid.TextMatrix(i, 0) = i
      Next
      End If
   Else
      Beep
      Beep
   End If
   
   
   
End If
End Sub

Private Sub Mygrid_LeaveCell()
If mygrid.col = 14 And dtPick.Visible = True Then
mygrid.TextMatrix(mygrid.row, mygrid.col) = dtPick.Value
End If



Dim i As Integer



If mygrid.col = 2 Then
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmsplantbatchdetail where trnid='" & Val(txtshipmentno.Text) & "' and plantbatch='" & Val(mygrid.TextMatrix(mygrid.row, 2)) & "'", MHVDB
If rs.EOF <> True Then
FindqmsPlanttype rs!planttype
FindqmsPlantVariety Right("00000" & rs!plantvariety, 3)
mygrid.TextMatrix(mygrid.row, 3) = Right("00000" & rs!plantvariety, 3) & " " & qmsPlantVariety
mygrid.TextMatrix(mygrid.row, 4) = rs!planttype & " " & qmsPlantType
Else
mygrid.TextMatrix(mygrid.row, 3) = ""
mygrid.TextMatrix(mygrid.row, 4) = ""

End If
End If



addcells

End Sub

Private Sub Mygrid_StartEdit(ByVal row As Long, ByVal col As Long, Cancel As Boolean)
 ' if this is a date column, edit it with the date picker control
    If mygrid.ColDataType(col) = flexDTDate Then
        
        ' we'll handle the editing ourselves
        Cancel = True
        
        ' position date picker control over cell
        dtPick.Move mygrid.CellLeft, mygrid.CellTop, mygrid.CellWidth, mygrid.CellHeight
        
        ' initialize value, save original in tag in case user hits escape
'        dtPick.Value = Mygrid
'        dtPick.Tag = Mygrid
        
        ' show and activate date picker control
        dtPick.Visible = True
        dtPick.SetFocus
        
        ' make it drop down the calendar
        'SendKeys "{f4}"
        
    End If

End Sub
