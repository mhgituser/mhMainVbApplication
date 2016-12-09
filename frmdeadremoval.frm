VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmdeadremoval 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "D E A D   R E M O V A L"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   Icon            =   "frmdeadremoval.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   12000
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
      Height          =   2535
      Left            =   9240
      TabIndex        =   20
      Top             =   3600
      Width           =   2655
      Begin VB.CommandButton Command2 
         Caption         =   "Load Transaction"
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
         Left            =   240
         Picture         =   "frmdeadremoval.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1920
         Width           =   2055
      End
      Begin VB.OptionButton optplantbatch 
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
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optall 
         Caption         =   "All"
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
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optfacility 
         Caption         =   "Facility"
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
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   720
         TabIndex        =   25
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81920001
         CurrentDate     =   41479
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   720
         TabIndex        =   26
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81920001
         CurrentDate     =   41479
      End
      Begin VB.Label Label15 
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
         Left            =   240
         TabIndex        =   28
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label14 
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
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   420
      End
   End
   Begin VB.TextBox txtgentot 
      Alignment       =   2  'Center
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
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txtdhtot 
      Alignment       =   2  'Center
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
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dead History Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6480
      TabIndex        =   13
      Top             =   720
      Width           =   5415
      Begin VSFlex7Ctl.VSFlexGrid mygrid1 
         Height          =   2055
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   5295
         _cx             =   9340
         _cy             =   3625
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmdeadremoval.frx":11CC
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
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   9135
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   2775
         Left            =   0
         TabIndex        =   14
         Top             =   120
         Width           =   9015
         _cx             =   15901
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmdeadremoval.frx":124D
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
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6255
      Begin MSComCtl2.DTPicker txtentrydate 
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81920001
         CurrentDate     =   41479
      End
      Begin VB.TextBox txtnoOfDead 
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
         Left            =   2160
         TabIndex        =   1
         Top             =   1800
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "frmdeadremoval.frx":1322
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSDataListLib.DataCombo cbofacilityId 
         Bindings        =   "frmdeadremoval.frx":1337
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   1080
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSDataListLib.DataCombo cboplantBatch 
         Bindings        =   "frmdeadremoval.frx":134C
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Top             =   1440
         Width           =   3975
         _ExtentX        =   7011
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No. of Dead Plants"
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
         TabIndex        =   11
         Top             =   1920
         Width           =   1965
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Plant batch"
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
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Label Label3 
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
         TabIndex        =   5
         Top             =   1200
         Width           =   1020
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   240
         Width           =   1560
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
            Picture         =   "frmdeadremoval.frx":1361
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdeadremoval.frx":16FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdeadremoval.frx":1A95
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdeadremoval.frx":276F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdeadremoval.frx":2BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdeadremoval.frx":337B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   1164
      ButtonWidth     =   1376
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
            Caption         =   "CANCEL"
            Key             =   "CANCEL"
            Object.ToolTipText     =   "CANCEL THE RECORD"
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
      Left            =   6600
      TabIndex        =   19
      Top             =   6240
      Width           =   555
   End
   Begin VB.Label Label5 
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
      Left            =   9720
      TabIndex        =   16
      Top             =   3240
      Width           =   555
   End
End
Attribute VB_Name = "frmdeadremoval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dhtot As Double
Dim gentot As Double
Dim shortby As String

Private Sub cbotrnid_LostFocus()
 On Error GoTo err
   
   cbotrnid.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblqmsdeadremoval where trnid='" & cbotrnid.BoundText & "' and status='ON'", MHVDB, adOpenForwardOnly, adLockOptimistic
   If rs.EOF <> True Then
    txtentrydate.Value = Format(rs!entrydate, "yyyy-MM-dd")
    txtnoOfDead.Text = rs!noofdead
    findQmsfacility rs!facilityid
    cbofacilityid.Text = rs!facilityid & " " & qmsFacility
    findQmsBatchDetail rs!plantbatch
    cboplantBatch.Text = qmsBatchdetail
   
   Else
   MsgBox "Record Not Found. Or the record is cancelled"
   TB.Buttons(3).Enabled = False
    TB.Buttons(4).Enabled = False
   End If
   rs.Close
   Exit Sub
err:
   MsgBox err.Description
   'rs.Close
End Sub

Private Sub Command2_Click()
FillGrid txtfrmdate.Value, txttodate.Value, shortby
End Sub

Private Sub Form_Load()
On Error GoTo err
Operation = ""
Dim rsF As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString


shortby = "A"
Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select trnId as description  from tblqmsdeadremoval where status='ON' order by trnId", db
Set cbotrnid.RowSource = rsF
cbotrnid.ListField = "description"
cbotrnid.BoundColumn = "description"

Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select concat(facilityid,' ',description) as description,facilityid  from tblqmsfacility order by facilityid", db
Set cbofacilityid.RowSource = rsF
cbofacilityid.ListField = "description"
cbofacilityid.BoundColumn = "facilityid"


Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "SELECT concat(cast(plantbatch as char),' ', c.description,' ', b.description) as description,plantbatch FROM  `tblqmsplantbatchdetail` a," _
 & "tblqmsplanttype b, tblqmsplantvariety c Where planttype = planttypeid" _
 & " AND plantvariety = varietyid order by plantbatch", db
Set cboplantBatch.RowSource = rsF
cboplantBatch.ListField = "description"
cboplantBatch.BoundColumn = "plantbatch"

'FillGrid
deadhistory
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub VSFlexGrid1_Click()

End Sub

Private Sub Option1_Click()
shortby = "F"
End Sub

Private Sub Option2_Click()
shortby = "A"
End Sub

Private Sub OPTALL_Click()
shortby = "A"
End Sub

Private Sub optfacility_Click()
shortby = "F"
End Sub

Private Sub optplantbatch_Click()
shortby = "B"
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
        rs.Open "SELECT MAX(trnid)+1 AS MaxID from tblqmsdeadremoval", MHVDB, adOpenForwardOnly, adLockOptimistic
        If rs.EOF <> True Then
        cbotrnid.Text = IIf(IsNull(rs!MaxID), "1", rs!MaxID)
        Else
        cbotrnid.Text = rs!MaxID
        End If
       Case "OPEN"
        Operation = "OPEN"
        CLEARCONTROLL
        cbotrnid.Enabled = True
        TB.Buttons(3).Enabled = True
        TB.Buttons(4).Enabled = True
       Case "SAVE"
        MNU_SAVE
        'FillGrid
        deadhistory
       Case "CANCEL"
         MNU_CANCEL
         deadhistory
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub MNU_CANCEL()
If MsgBox("Do You Want to Cancel The Reccord No. " & cbotrnid.Text & "?", vbQuestion + vbYesNo) = vbYes Then
MHVDB.Execute "update tblqmsdeadremoval set status='C' where trnid='" & cbotrnid.BoundText & "'"
 TB.Buttons(3).Enabled = False
TB.Buttons(4).Enabled = False

LogRemarks = "Canceled Transaction No." & cbotrnid.BoundText & "," & "from table tblqmsdeadremoval"
updatemhvlog Now, MUSER, LogRemarks, ""


Else

End If

End Sub
Private Sub MNU_SAVE()
Dim rs As New ADODB.Recordset
On Error GoTo err
If Len(cbofacilityid.Text) = 0 Then
MsgBox "Select facility."
Exit Sub
End If

If Len(cboplantBatch.Text) = 0 Then
MsgBox "Select Plant Batch."
Exit Sub
End If

If Val(txtnoOfDead.Text) = 0 Then
MsgBox "Invalid No of Dead Input."
Exit Sub
End If
If Len(cbotrnid.Text) = 0 Then
MsgBox "Transaction No. is Must."
Exit Sub
End If

MHVDB.BeginTrans
If Operation = "ADD" Then
LogRemarks = ""
MHVDB.Execute "INSERT INTO tblqmsdeadremoval (trnid,entrydate,facilityid,plantbatch," _
            & "noofdead,status,location) " _
            & "VALUEs('" & cbotrnid.Text & "','" & Format(txtentrydate.Value, "yyyy-MM-dd") & "','" & cbofacilityid.BoundText & "', " _
            & "'" & cboplantBatch.BoundText & "','" & Val(txtnoOfDead.Text) & "','ON','" & Mlocation & "')"
 
 
LogRemarks = "Inserted new record" & cbotrnid.BoundText & "," & cboplantBatch.Text & "," & Mlocation
updatemhvlog Now, MUSER, LogRemarks, ""

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblqmsdeadremoval set entrydate='" & Format(txtentrydate.Value, "yyyy-MM-dd") & "' " _
            & ",facilityid='" & cbofacilityid.BoundText & "',plantbatch='" & cboplantBatch.BoundText & "' " _
            & ",noofdead='" & Val(txtnoOfDead.Text) & "' " _
            & " where trnid='" & cbotrnid.BoundText & "' and location='" & Mlocation & "'"

LogRemarks = "Updated  record" & cbotrnid.BoundText & "," & cboplantBatch.Text & "," & Mlocation
updatemhvlog Now, MUSER, LogRemarks, ""
Else
MsgBox "OPERATION NOT SELECTED."
End If
MHVDB.CommitTrans
TB.Buttons(3).Enabled = False
TB.Buttons(4).Enabled = False
Exit Sub

err:
MsgBox err.Description
TB.Buttons(3).Enabled = False
MHVDB.RollbackTrans


End Sub
Private Sub CLEARCONTROLL()
    cbofacilityid.Text = ""
   cboplantBatch.Text = ""
     txtnoOfDead.Text = ""
   

End Sub
Private Sub FillGrid(frmdate As Date, todate As Date, shortby As String)
On Error GoTo err
gentot = 0
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^Sl.No.|^Trn. Id|^Date|^Facility|^Plant Batch|^Dead Plants|^"
Mygrid.ColWidth(0) = 645
Mygrid.ColWidth(1) = 960
Mygrid.ColWidth(2) = 1095
Mygrid.ColWidth(3) = 2205
Mygrid.ColWidth(4) = 1905
Mygrid.ColWidth(5) = 1665
Mygrid.ColWidth(6) = 345

If shortby = "A" Then
rs.Open "select * from tblqmsdeadremoval where status='ON' and entrydate>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and entrydate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' order by  trnid", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf shortby = "F" Then
If Len(cbofacilityid.Text) = 0 Then
MsgBox "Select Facility From The Drop Down Combo."
Exit Sub
End If
rs.Open "select * from tblqmsdeadremoval where status='ON' and entrydate>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and entrydate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and facilityid='" & cbofacilityid.BoundText & "' order by  trnid", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf shortby = "B" Then
If Len(cboplantBatch.Text) = 0 Then
MsgBox "Select Batch No. From The Drop Down Combo."
Exit Sub
End If
rs.Open "select * from tblqmsdeadremoval where status='ON' and entrydate>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and entrydate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and plantbatch='" & cboplantBatch.BoundText & "' order by  trnid", MHVDB, adOpenForwardOnly, adLockOptimistic
Else
MsgBox "Invalid Selection."
Exit Sub
End If

i = 1
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i
Mygrid.TextMatrix(i, 1) = UCase(rs!trnid)
Mygrid.TextMatrix(i, 2) = Format(rs!entrydate, "dd/MM/yyyy")
findQmsfacility rs!facilityid
Mygrid.TextMatrix(i, 3) = rs!facilityid & " " & qmsFacility
findQmsBatchDetail rs!plantbatch
Mygrid.TextMatrix(i, 4) = qmsBatchdetail1
Mygrid.TextMatrix(i, 5) = rs!noofdead
gentot = gentot + rs!noofdead
rs.MoveNext
i = i + 1
Loop
txtgentot.Text = Format(gentot, "##,##,##")
rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub
Private Sub deadhistory()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mygrid1.Clear
dhtot = 0
mygrid1.Rows = 1
mygrid1.FormatString = "^Sl.No.|^Facility|^Dead Plants|^"
mygrid1.ColWidth(0) = 645
mygrid1.ColWidth(1) = 3105
mygrid1.ColWidth(2) = 1185
mygrid1.ColWidth(3) = 225



rs.Open "select facilityid,sum(noofdead) as dead from tblqmsdeadremoval where status='ON'  group by facilityid order by facilityid", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
mygrid1.Rows = mygrid1.Rows + 1
mygrid1.TextMatrix(i, 0) = i
'mygrid1.TextMatrix(i, 1) = (rs!yr)
findQmsfacility rs!facilityid
mygrid1.TextMatrix(i, 1) = rs!facilityid & " " & qmsFacility
Mygrid.ColAlignment(1) = flexAlignLeftTop
mygrid1.TextMatrix(i, 2) = rs!dead
Mygrid.ColAlignment(2) = flexAlignRightTop
dhtot = dhtot + rs!dead
rs.MoveNext
i = i + 1
Loop
txtdhtot.Text = Format(dhtot, "##,##,##")
rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub

