VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmmonitorfarmerupdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitor-Farmer Update"
   ClientHeight    =   7980
   ClientLeft      =   1785
   ClientTop       =   1095
   ClientWidth     =   20400
   Icon            =   "frmmonitorfarmerupdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   20400
   Begin VB.Frame Frame2 
      Caption         =   "Farmer Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5880
      TabIndex        =   11
      Top             =   840
      Width           =   7095
      Begin VB.TextBox txttotaldropout 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txttotalrejected 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txttotalplantedlist 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txttotalactive 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Active Farmer(Planted List)"
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
         TabIndex        =   15
         Top             =   840
         Width           =   2805
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Rejected"
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
         Left            =   4440
         TabIndex        =   14
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Active Farmer"
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
         TabIndex        =   13
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Dropout"
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
         Left            =   4440
         TabIndex        =   12
         Top             =   840
         Width           =   1185
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Monitor Without Supervisor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   14760
      TabIndex        =   9
      Top             =   1680
      Width           =   5775
      Begin VSFlex7Ctl.VSFlexGrid mwsgrid 
         Height          =   5415
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5535
         _cx             =   9763
         _cy             =   9551
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
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmmonitorfarmerupdate.frx":0E42
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
   Begin VB.Frame Frame4 
      Caption         =   "Farmer without monitor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   11760
      TabIndex        =   7
      Top             =   2160
      Width           =   5775
      Begin VSFlex7Ctl.VSFlexGrid fwmgrid 
         Height          =   5415
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5535
         _cx             =   9763
         _cy             =   9551
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
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmmonitorfarmerupdate.frx":0EB5
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
   Begin VB.Frame Frame3 
      Height          =   5775
      Left            =   5400
      TabIndex        =   5
      Top             =   2640
      Width           =   6015
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   5415
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5775
         _cx             =   10186
         _cy             =   9551
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
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmmonitorfarmerupdate.frx":0F27
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
      Caption         =   "Monitor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5775
      Begin MSDataListLib.DataCombo cbomonitor 
         Bindings        =   "frmmonitorfarmerupdate.frx":0F99
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   ""
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
      Begin MSDataListLib.DataCombo cbosupervisor 
         Bindings        =   "frmmonitorfarmerupdate.frx":0FAE
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   ""
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Supervisor"
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
         TabIndex        =   2
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Active Monitor"
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
         TabIndex        =   1
         Top             =   480
         Width           =   1245
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   9000
      Top             =   720
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
            Picture         =   "frmmonitorfarmerupdate.frx":0FC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorfarmerupdate.frx":135D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorfarmerupdate.frx":16F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorfarmerupdate.frx":23D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorfarmerupdate.frx":2823
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorfarmerupdate.frx":2FDD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   3480
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
            Picture         =   "frmmonitorfarmerupdate.frx":3377
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorfarmerupdate.frx":3711
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorfarmerupdate.frx":3AAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorfarmerupdate.frx":4785
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorfarmerupdate.frx":4BD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorfarmerupdate.frx":5391
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   1164
      ButtonWidth     =   1005
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
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
   Begin VB.Image imgBtnUp 
      Height          =   240
      Left            =   4440
      Picture         =   "frmmonitorfarmerupdate.frx":572B
      Top             =   4200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgBtnDn 
      Height          =   240
      Left            =   3840
      Picture         =   "frmmonitorfarmerupdate.frx":5AB5
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmmonitorfarmerupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBOMONITOR_LostFocus()
If Len(cbomonitor.Text) = 0 Then Exit Sub
Dim rs As New ADODB.Recordset
cbomonitor.Enabled = False

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & cbomonitor.BoundText & "'", MHVDB
If rs.EOF <> True Then
If Len(rs!msupervisor) > 0 Then
FindsTAFF rs!msupervisor
End If
cbosupervisor.Text = rs!msupervisor & " " & sTAFF
End If

fillmain
TB.buttons(3).Enabled = True
End Sub

Private Sub Form_Load()
 
    
    fillmonitor
    fillms
    fillfwm
    fillmws
    fillsummary
End Sub
Private Sub addbtn()
' initialize grid
    mygrid.Editable = flexEDKbdMouse
    mygrid.AllowUserResizing = flexResizeNone
    
    ' add some buttons to the grid
    Dim i%
    For i = 1 To mygrid.rows - 1
        mygrid.Cell(flexcpPicture, i, 2) = imgBtnUp
        mygrid.Cell(flexcpPictureAlignment, i, 2) = flexAlignRightCenter
    Next
End Sub
Private Sub fillmonitor()
On Error GoTo err
Operation = ""
Dim Srs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set Srs = Nothing

If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff where moniter='1' order by STAFFCODE", db
Set cbomonitor.RowSource = Srs
cbomonitor.ListField = "STAFFNAME"
cbomonitor.BoundColumn = "STAFFCODE"


Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub fillms()
On Error GoTo err
Operation = ""
Dim Srs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set Srs = Nothing

If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select distinct concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff where MSUPERVISOR<>''  order by STAFFCODE", db
Set cbosupervisor.RowSource = Srs
cbosupervisor.ListField = "STAFFNAME"
cbosupervisor.BoundColumn = "STAFFCODE"


Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Form_Resize()
   ' On Error Resume Next
    'Mygrid.Move Mygrid.Left, Mygrid.Top, ScaleWidth - 2 * Mygrid.Left, ScaleHeight - Mygrid.Left - Mygrid.Top
End Sub

Private Sub Mygrid_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
' only interesetd in left button
   On Error Resume Next
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    r = mygrid.MouseRow
    c = mygrid.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    ' make sure the click was on a cell with a button
    If mygrid.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
    
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = mygrid.Cell(flexcpLeft, r, c) + mygrid.Cell(flexcpWidth, r, c) - X
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    mygrid.Cell(flexcpPicture, r, c) = imgBtnDn
    'MsgBox "Thanks for clicking my custom button!"
    'MsgBox "You have clicked on row=" & r & "  col=" & c
    
    If MsgBox("Make sure that you want to delink farmer " & mygrid.TextMatrix(r, 1) & " from monitor " & cbomonitor.Text & ".", vbYesNo) = vbYes Then
    removefarmer r, c
    End If
    
    mygrid.Cell(flexcpPicture, r, c) = imgBtnUp
    
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub


Private Sub fillmain()
On Error GoTo err

Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mygrid.Clear
mygrid.rows = 1
mygrid.FormatString = "^Sl.No.|^Farmer|^|^"
mygrid.ColWidth(0) = 645
mygrid.ColWidth(1) = 4095
mygrid.ColWidth(2) = 270
mygrid.ColWidth(3) = 405

rs.Open "select * from tblfarmer where status='A' and monitor='" & cbomonitor.BoundText & "' order by idfarmer", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
mygrid.rows = mygrid.rows + 1
mygrid.TextMatrix(i, 0) = i
mygrid.TextMatrix(i, 1) = rs!idfarmer & "  " & rs!farmername
mygrid.ColAlignment(1) = flexAlignLeftTop
rs.MoveNext
i = i + 1
Loop

rs.Close
addbtn

Exit Sub
err:
MsgBox err.Description

End Sub
Private Sub fillfwm()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
fwmgrid.Clear
fwmgrid.rows = 1
fwmgrid.FormatString = "^Sl.No.|^Farmer|^|^"
fwmgrid.ColWidth(0) = 645
fwmgrid.ColWidth(1) = 4095
fwmgrid.ColWidth(2) = 270
fwmgrid.ColWidth(3) = 405
MHVDB.Execute "update tblfarmer set monitor='' where length(monitor)<>5"
rs.Open "select * from tblfarmer where status='A' and monitor='' order by idfarmer", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
fwmgrid.rows = fwmgrid.rows + 1
fwmgrid.TextMatrix(i, 0) = i
fwmgrid.TextMatrix(i, 1) = rs!idfarmer & "  " & rs!farmername
fwmgrid.ColAlignment(1) = flexAlignLeftTop
rs.MoveNext
i = i + 1
Loop

rs.Close


Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub fillmws()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mwsgrid.Clear
mwsgrid.rows = 1
mwsgrid.FormatString = "^Sl.No.|^Monitor|^|^"
mwsgrid.ColWidth(0) = 645
mwsgrid.ColWidth(1) = 4095
mwsgrid.ColWidth(2) = 270
mwsgrid.ColWidth(3) = 405
MHVDB.Execute "UPDATE  tblmhvstaff SET msupervisor = '' WHERE LENGTH( msupervisor )<>'5'"
rs.Open "select * from tblmhvstaff where moniter='1' and msupervisor='' order by staffcode", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
mwsgrid.rows = mwsgrid.rows + 1
mwsgrid.TextMatrix(i, 0) = i
mwsgrid.TextMatrix(i, 1) = rs!staffcode & "  " & rs!staffname
mwsgrid.ColAlignment(1) = flexAlignLeftTop
rs.MoveNext
i = i + 1
Loop

rs.Close


Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub mygrid_Click()
If mygrid.row = 0 Then Exit Sub
If mygrid.col = 1 And Len(mygrid.TextMatrix(mygrid.row - 1, 1)) > 0 And Len(Trim(mygrid.TextMatrix(mygrid.row, 1))) = 0 Then
mygrid.Editable = flexEDKbdMouse
FillGridCombo
Else
mygrid.ComboList = ""
mygrid.Editable = flexEDNone
End If
End Sub

Private Sub mygrid_DblClick()
'myinput = InputBox("Enter No. of Rows You Want " & 1)
'            If Not IsNumeric(myinput) Then
'            MsgBox "Invalid number,Double Click again to enable the input box."
'            Else
'
'            Mygrid.Rows = Mygrid.Rows + CInt(myinput)
'            End If
If Len(cbomonitor.Text) = 0 Then Exit Sub
If MsgBox("Do You want to add a farmer?", vbYesNo) = vbYes Then
mygrid.rows = mygrid.rows + 1
addbtn
Else

End If

End Sub

Private Sub FillGridCombo()
        Dim RstTemp As New ADODB.Recordset
        Dim NmcItemData As String
        Dim StrComboList As String
        Dim mn1 As String
        mn1 = "          |"
        StrComboList = ""
        MHVDB.Execute "UPDATE  `tblfarmer` SET monitor = '' WHERE LENGTH( monitor )<>'5'"
            Set RstTemp = Nothing
            RstTemp.Open ("select idfarmer,farmername from tblfarmer where status='A' and monitor='' ORDER BY idfarmer"), MHVDB

            If Not RstTemp.BOF Then
               Do While Not RstTemp.EOF
                    NmcItemData = ""
                    StrComboList = IIf(StrComboList = "", RstTemp("idfarmer").Value & " " & RstTemp("farmername").Value, StrComboList & "|" & RstTemp("idfarmer").Value) & " " & RstTemp("farmername").Value

                    RstTemp.MoveNext
               Loop
            End If
            RstTemp.Close
       mygrid.ComboList = StrComboList



    End Sub
    
    Private Sub removefarmer(mrow As Long, MCOL As Long)
    Dim rs As New ADODB.Recordset
    MHVDB.Execute "update tblfarmer set monitor='' where idfarmer='" & Mid(Trim(mygrid.TextMatrix(mrow, 1)), 1, 14) & "'"
    mygrid.RemoveItem (mrow)
    fillfwm
    mygrid.ComboList = ""
    mygrid.Editable = flexEDNone
    End Sub
    
Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       
       Case "OPEN"
       Operation = "OPEN"
       CLEARCONTROLL
       cbomonitor.Enabled = True
       TB.buttons(3).Enabled = False
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
cbomonitor.Text = ""
cbosupervisor.Text = ""
mygrid.Clear
mygrid.rows = 1
mygrid.FormatString = "^Sl.No.|^Farmer|^|^"
mygrid.ColWidth(0) = 645
mygrid.ColWidth(1) = 4095
mygrid.ColWidth(2) = 270
mygrid.ColWidth(3) = 405
End Sub
Private Sub MNU_SAVE()
Dim i As Integer
Dim rs As New ADODB.Recordset
If Len(cbosupervisor.Text) = 0 Then
MsgBox "Not a valid supervosor!"
Exit Sub
End If

MHVDB.BeginTrans
Set rs = Nothing
rs.Open "Select * from tblmhvstaff where staffcode='" & cbosupervisor.BoundText & "'", MHVDB
If rs.EOF <> True Then

Else
MsgBox "Not a valid supervosor!"
MHVDB.RollbackTrans
Exit Sub
End If

' update m supervisor
MHVDB.Execute "update tblmhvstaff set msupervisor='" & Mid(Trim(cbosupervisor.Text), 1, 5) & "' where staffcode='" & cbomonitor.BoundText & "'"
' update farmers with
For i = 1 To mygrid.rows - 1
If Len(Trim(mygrid.TextMatrix(i, 1))) = 0 Then Exit For
MHVDB.Execute "update tblfarmer set monitor='" & cbomonitor.BoundText & "' where idfarmer='" & Mid(Trim(mygrid.TextMatrix(i, 1)), 1, 14) & "'"
Next
MHVDB.CommitTrans

TB.buttons(3).Enabled = False
fillmain
fillfwm
fillmws

End Sub


Private Sub fillsummary()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select count(*) as cnt from tblfarmer where status='A'", MHVDB
If rs.EOF <> True Then
txttotalactive.Text = rs!cnt
End If

Set rs = Nothing
rs.Open "select count(*) as cnt from tblfarmer where status='D'", MHVDB
If rs.EOF <> True Then
txttotaldropout.Text = rs!cnt
End If

Set rs = Nothing
rs.Open "select count(*) as cnt from tblfarmer where status='R'", MHVDB
If rs.EOF <> True Then
txttotalrejected.Text = rs!cnt
End If

Set rs = Nothing
rs.Open "select count(*) as cnt from tblfarmer where idfarmer in(select farmercode from tblplanted)", MHVDB
If rs.EOF <> True Then
txttotalplantedlist.Text = rs!cnt
End If












End Sub
























