VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMCLIMATEMETEROLOGY01 
   Caption         =   "C L I M A T E   M E T E R O L O G Y . . . . ."
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11895
   Icon            =   "FRMCLIMATEMETEROLOGY.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Sort By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9000
      TabIndex        =   25
      Top             =   2760
      Width           =   2895
      Begin VB.CommandButton Command2 
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
         Left            =   2040
         Picture         =   "FRMCLIMATEMETEROLOGY.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   600
         TabIndex        =   27
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81395713
         CurrentDate     =   41479
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   600
         TabIndex        =   28
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81395713
         CurrentDate     =   41479
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         TabIndex        =   30
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         TabIndex        =   29
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sun Rise/Sun Set not applicable"
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
      TabIndex        =   24
      Top             =   840
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   0
      TabIndex        =   11
      Top             =   3840
      Width           =   11895
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   2775
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   11775
         _cx             =   20770
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
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FRMCLIMATEMETEROLOGY.frx":11CC
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
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   7695
      Begin VB.TextBox txtcomments 
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
         Left            =   2400
         TabIndex        =   23
         Top             =   2640
         Width           =   5175
      End
      Begin VB.TextBox txtmaxrh 
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
         Left            =   2400
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtminrh 
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
         Left            =   6240
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtmaxdegree 
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
         Left            =   2400
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtmindegree 
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
         Left            =   6240
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker txtentrydate 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81395713
         CurrentDate     =   41479
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "FRMCLIMATEMETEROLOGY.frx":1337
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
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
      Begin MSComCtl2.DTPicker txtstarttime 
         Height          =   375
         Left            =   6240
         TabIndex        =   13
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   81395714
         CurrentDate     =   41480
      End
      Begin MSComCtl2.DTPicker txtsunrisetime 
         Height          =   375
         Left            =   2400
         TabIndex        =   20
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   81395714
         CurrentDate     =   41480
      End
      Begin MSComCtl2.DTPicker txtsunsettime 
         Height          =   375
         Left            =   6240
         TabIndex        =   21
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   81395714
         CurrentDate     =   41480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Commets"
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
         TabIndex        =   22
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Sun Rise Time"
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
         TabIndex        =   19
         Top             =   2280
         Width           =   1530
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Sunset Time"
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
         Left            =   3960
         TabIndex        =   18
         Top             =   2280
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Max RH"
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
         TabIndex        =   17
         Top             =   1800
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Min RH"
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
         Left            =   3960
         TabIndex        =   16
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Max Degree Celceius"
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
         Top             =   1320
         Width           =   2235
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Record Time"
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
         Left            =   3960
         TabIndex        =   8
         Top             =   720
         Width           =   1365
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
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transaction  Id"
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
         TabIndex        =   6
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Min degree Celceius"
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
         Left            =   3960
         TabIndex        =   5
         Top             =   1320
         Width           =   2145
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
            Picture         =   "FRMCLIMATEMETEROLOGY.frx":134C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCLIMATEMETEROLOGY.frx":16E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCLIMATEMETEROLOGY.frx":1A80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCLIMATEMETEROLOGY.frx":275A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCLIMATEMETEROLOGY.frx":2BAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCLIMATEMETEROLOGY.frx":3366
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
Attribute VB_Name = "FRMCLIMATEMETEROLOGY01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbotrnid_LostFocus()
On Error GoTo err
   
   cbotrnid.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblqmsclimatemeterology01 where trnid='" & cbotrnid.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
   If rs.EOF <> True Then
    txtentrydate.Value = Format(rs!entrydate, "yyyy-MM-dd")
    txtstarttime.Value = Format(rs!recordtime, "HH:mm:ss")
    txtsunrisetime.Value = Format(rs!timesunrise, "HH:mm:ss")
    txtsunsettime.Value = Format(rs!timesunset, "HH:mm:ss")
    txtmaxdegree.Text = rs!degreemax
    txtmindegree.Text = rs!degreemin
    txtmaxrh.Text = rs!maxrh
    txtminrh.Text = rs!minrh
    txtcomments.Text = rs!Comments
    
   
   Else
   MsgBox "Record Not Found."
   End If
   rs.Close
   Exit Sub
err:
   MsgBox err.Description
End Sub

Private Sub Check1_Click()
txtsunrisetime.Value = "00:00:00"
txtsunsettime.Value = "00:00:00"
End Sub

Private Sub Command2_Click()
FillGrid txtfrmdate.Value, txttodate.Value
End Sub

Private Sub Form_Load()
On Error GoTo err
Operation = ""
Dim rsF As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString


Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select trnid  from tblqmsclimatemeterology01 order by trnid", db
Set cbotrnid.RowSource = rsF
cbotrnid.ListField = "trnid"
cbotrnid.BoundColumn = "trnid"

'FillGrid
  txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
  txttodate.Value = Format(Now, "dd/MM/yyyy")

Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
            
        cbotrnid.Enabled = False
        TB.Buttons(3).Enabled = True
        Operation = "ADD"
        CLEARCONTROLL
        Dim rs As New ADODB.Recordset
        Set rs = Nothing
        rs.Open "SELECT MAX(trnid)+1 AS MaxID from tblqmsclimatemeterology01", MHVDB, adOpenForwardOnly, adLockOptimistic
        If rs.EOF <> True Then
        cbotrnid.Text = IIf(IsNull(rs!MaxId), 1, rs!MaxId)
        Else
        cbotrnid.Text = rs!MaxId
        End If
       Case "OPEN"
        Operation = "OPEN"
        CLEARCONTROLL
        cbotrnid.Enabled = True
        TB.Buttons(3).Enabled = True
             
       Case "SAVE"
        MNU_SAVE
      
        'FillGrid
       
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
MsgBox "Transaction No. is Must."
Exit Sub
End If



MHVDB.BeginTrans
If Operation = "ADD" Then
LogRemarks = ""
MHVDB.Execute "INSERT INTO tblqmsclimatemeterology01 (trnid,entrydate,recordtime,degreemax,degreemin," _
            & "maxrh,minrh,timesunrise,timesunset,comments,status,location)" _
            & "values(" _
            & "'" & cbotrnid.BoundText & "'," _
            & "'" & Format(txtentrydate.Value, "yyyy-MM-dd") & "'," _
            & "'" & Format(txtstarttime.Value, "HH:mm:ss") & "'," _
            & "'" & Val(txtmaxdegree.Text) & "'," _
            & "'" & Val(txtmindegree.Text) & "'," _
            & "'" & Val(txtmaxrh.Text) & "'," _
            & "'" & Val(txtminrh.Text) & "'," _
            & "'" & Format(txtsunrisetime.Value, "HH:mm:ss") & "'," _
            & "'" & Format(txtsunsettime.Value, "HH:mm:ss") & "'," _
            & "'" & txtcomments.Text & "'," _
            & "'ON'," _
            & "'" & Mlocation & "'" _
            & ")"
 
 
LogRemarks = "Inserted new record" & cbotrnid.BoundText & "," & Mlocation & ","
updatemhvlog Now, MUSER, LogRemarks, ""

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblqmsclimatemeterology01 set " _
            & "entrydate='" & Format(txtentrydate.Value, "yyyy-MM-dd") & "'," _
            & "recordtime='" & Format(txtstarttime.Value, "HH:mm:ss") & "'," _
            & "degreemax='" & Val(txtmaxdegree.Text) & "'," _
            & "degreemin='" & Val(txtmindegree.Text) & "'," _
            & "maxrh='" & Val(txtmaxrh.Text) & "'," _
            & "minrh='" & Val(txtminrh.Text) & "'," _
            & "timesunrise='" & Format(txtsunrisetime.Value, "HH:mm:ss") & "'," _
            & "comments='" & txtcomments.Text & "'," _
            & "timesunset='" & Format(txtsunsettime.Value, "HH:mm:ss") & "'" _
            & " where trnid='" & cbotrnid.BoundText & "' and location='" & Mlocation & "'"
            

LogRemarks = "Updated  record" & cbotrnid.BoundText & "," & Mlocation
updatemhvlog Now, MUSER, LogRemarks, ""
Else
MsgBox "OPERATION NOT SELECTED."
End If
 TB.Buttons(3).Enabled = False
 
MHVDB.CommitTrans
FillGrid txtfrmdate.Value, txttodate.Value
Exit Sub

err:
MsgBox err.Description
MHVDB.RollbackTrans


End Sub
Private Sub CLEARCONTROLL()
  txtentrydate.Value = Format(Now, "dd/MM/yyyy")
  txtstarttime.Value = Format(Now, "HH:mm:ss")
  txtsunrisetime.Value = Format("06:30:00", "HH:mm:ss")
  txtsunsettime.Value = Format("17:30:00", "HH:mm:ss")
    txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
  txttodate.Value = Format(Now, "dd/MM/yyyy")
  txtmaxdegree.Text = ""
  txtmindegree.Text = ""
  txtmaxrh.Text = ""
  txtminrh.Text = ""
  txtcomments.Text = ""
End Sub
Private Sub FillGrid(frmdate As Date, todate As Date)
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mygrid.Clear
mygrid.Rows = 1
mygrid.FormatString = "^Sl.No.|^Trn. Id|^Date|^Record Time|^Max Degree |^Min Degree|^Max RH|^Min RH|^Sun Rise Time|^Sun Set Time|^Comments|^"
mygrid.ColWidth(0) = 645
mygrid.ColWidth(1) = 975
mygrid.ColWidth(2) = 1275
mygrid.ColWidth(3) = 1230
mygrid.ColWidth(4) = 1110
mygrid.ColWidth(5) = 1080
mygrid.ColWidth(6) = 795
mygrid.ColWidth(7) = 720
mygrid.ColWidth(8) = 1395
mygrid.ColWidth(9) = 1305
mygrid.ColWidth(10) = 960
mygrid.ColWidth(11) = 195

rs.Open "select * from tblqmsclimatemeterology01 where entrydate>='" & Format(frmdate, "yyyy-MM-dd") & "' and entrydate<='" & Format(todate, "yyyy-MM-dd") & "' order by trnid", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
mygrid.Rows = mygrid.Rows + 1
mygrid.TextMatrix(i, 0) = i

mygrid.TextMatrix(i, 1) = rs!trnid
mygrid.TextMatrix(i, 2) = Format(rs!entrydate, "dd/MM/yyyy")
mygrid.TextMatrix(i, 3) = Format(rs!recordtime, "HH:mm:ss")
mygrid.TextMatrix(i, 4) = rs!degreemax
mygrid.TextMatrix(i, 5) = rs!degreemin
mygrid.TextMatrix(i, 6) = rs!maxrh
mygrid.TextMatrix(i, 7) = rs!minrh
mygrid.TextMatrix(i, 8) = Format(rs!timesunrise, "HH:mm:ss")
mygrid.TextMatrix(i, 9) = Format(rs!timesunset, "HH:mm:ss")
mygrid.TextMatrix(i, 10) = rs!Comments
rs.MoveNext
i = i + 1
Loop

rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub
