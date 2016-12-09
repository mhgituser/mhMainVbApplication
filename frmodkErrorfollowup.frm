VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmodkErrorfollowup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ODK ERROR FOLLOUP"
   ClientHeight    =   8070
   ClientLeft      =   2535
   ClientTop       =   1710
   ClientWidth     =   17685
   Icon            =   "frmodkErrorfollowup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   17685
   Begin VB.Frame Frame4 
      Caption         =   "SORT BY DATE"
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
      Left            =   8760
      TabIndex        =   10
      Top             =   720
      Width           =   8655
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   5880
         Picture         =   "frmodkErrorfollowup.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   735
      End
      Begin MSComCtl2.DTPicker TXTFROMDATE 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   110886913
         CurrentDate     =   41492
      End
      Begin VB.CheckBox CHKSELECT 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin MSComCtl2.DTPicker TXTTODATE 
         Height          =   375
         Left            =   4320
         TabIndex        =   15
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   110886913
         CurrentDate     =   41492
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FROM DATE"
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
         Left            =   600
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TO DATE"
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
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "ATTENDED ERROR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   0
      TabIndex        =   4
      Top             =   4560
      Width           =   17535
      Begin VSFlex7Ctl.VSFlexGrid mygrid1 
         Height          =   2775
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   17295
         _cx             =   30506
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
         BackColorAlternate=   12648384
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
         FormatString    =   $"frmodkErrorfollowup.frx":11CC
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8535
      Begin VB.TextBox txtparamvalue 
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
         Height          =   375
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin MSDataListLib.DataCombo cboparaid 
         Bindings        =   "frmodkErrorfollowup.frx":1303
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PARAMETER VALUE"
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
         Left            =   5520
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PARAMETER NAME"
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
         TabIndex        =   2
         Top             =   360
         Width           =   1740
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   6360
      Top             =   600
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
            Picture         =   "frmodkErrorfollowup.frx":1318
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmodkErrorfollowup.frx":16B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmodkErrorfollowup.frx":1A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmodkErrorfollowup.frx":2726
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmodkErrorfollowup.frx":2B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmodkErrorfollowup.frx":3332
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
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
            Picture         =   "frmodkErrorfollowup.frx":36CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmodkErrorfollowup.frx":3A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmodkErrorfollowup.frx":3E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmodkErrorfollowup.frx":4ADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmodkErrorfollowup.frx":4F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmodkErrorfollowup.frx":56E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   17685
      _ExtentX        =   31194
      _ExtentY        =   1164
      ButtonWidth     =   1217
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
   Begin VB.Frame Frame2 
      Caption         =   "UNATTENDED ERROR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   17535
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   2775
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Width           =   17295
         _cx             =   30506
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
         BackColorAlternate=   12648447
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmodkErrorfollowup.frx":5A80
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
         Begin VB.Frame Frame5 
            Caption         =   "FARMER DETAIL"
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
            Left            =   0
            TabIndex        =   17
            Top             =   240
            Visible         =   0   'False
            Width           =   7335
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
               Left            =   6120
               Picture         =   "frmodkErrorfollowup.frx":5C03
               Style           =   1  'Graphical
               TabIndex        =   33
               ToolTipText     =   "View Pest History"
               Top             =   1800
               Width           =   1095
            End
            Begin VB.TextBox TXTTS 
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
               Height          =   375
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   1680
               Width           =   3615
            End
            Begin VB.TextBox TXTGE 
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
               Height          =   375
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   1200
               Width           =   2175
            End
            Begin VB.TextBox TXTDZ 
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
               Height          =   375
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   1200
               Width           =   2175
            End
            Begin VB.TextBox TXTPLANTS 
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
               Height          =   375
               Left            =   6240
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox TXTAREA 
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
               Height          =   375
               Left            =   3720
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox TXTFARMERNAME 
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
               Height          =   375
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   240
               Width           =   5655
            End
            Begin VB.TextBox TXTSTATUS 
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
               Height          =   375
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   720
               Width           =   975
            End
            Begin VB.CommandButton Command4 
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
               Left            =   5400
               Picture         =   "frmodkErrorfollowup.frx":63AD
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Exit Farmer Detail"
               Top             =   1800
               Width           =   615
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "TSHOWOG"
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
               TabIndex        =   31
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "GEWOG"
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
               Left            =   4320
               TabIndex        =   30
               Top             =   1320
               Width           =   720
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "DZONGKHAG "
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
               TabIndex        =   27
               Top             =   1320
               Width           =   1245
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "NO. OF PLANTS"
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
               Left            =   4800
               TabIndex        =   26
               Top             =   840
               Width           =   1425
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "ACRE REG."
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
               TabIndex        =   23
               Top             =   840
               Width           =   1020
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "FARMER NAME"
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
               TabIndex        =   21
               Top             =   360
               Width           =   1365
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "STATUS"
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
               TabIndex        =   19
               Top             =   840
               Width           =   750
            End
         End
      End
   End
End
Attribute VB_Name = "frmodkErrorfollowup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim paraId As Integer
Private Sub cboparaid_LostFocus()
'On Error GoTo err
Dim rs As New ADODB.Recordset
Set rs = Nothing
cboparaid.Enabled = False
TB.buttons(3).Enabled = True
rs.Open "select * from tblodkfollowuplog where paraid='" & cboparaid.BoundText & "' order by paraid", ODKDB

If rs.EOF <> True Then
findParamDetails rs!paraId
If ispercentage Then
txtparamvalue.Text = paramValue & "%"
Else
txtparamvalue.Text = paramValue
End If
filldetail rs!paraId
filldetail1 rs!paraId
End If

End Sub
Private Sub filldetail(trnid As Integer)
On Error GoTo err
findParamDetails trnid
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mchk = True
mygrid.Clear
mygrid.rows = 1
mygrid.FormatString = "^SL.NO.|^ENTRY DATE|^START DATE|^" & paramName & "|^STAFF|^FARMER|^FD. CODE|^ACTION TAKEN|^RECOMMENDATION|^QC COMMENT|^uri|^email|^"
mygrid.ColWidth(0) = 750
mygrid.ColWidth(1) = 1320
mygrid.ColWidth(2) = 1665
mygrid.ColWidth(3) = 1125
'Mygrid.ColWidth(4) = 1365
mygrid.ColWidth(4) = 2460
mygrid.ColWidth(5) = 2820
mygrid.ColWidth(6) = 960
mygrid.ColWidth(7) = 1530
mygrid.ColWidth(8) = 1980
mygrid.ColWidth(9) = 2250
mygrid.ColWidth(10) = 120
mygrid.ColWidth(11) = 1     ' URI
mygrid.ColWidth(12) = 1     ' EMAIL


If CHKSELECT.Value = 0 Then
rs.Open "select * from tblodkfollowuplog where paraid='" & trnid & "' and followupstatus='ON' order by odkstartdate desc", ODKDB
Else
rs.Open "select * from tblodkfollowuplog where paraid='" & trnid & "' and followupstatus='ON' AND odkstartdate>='" & Format(txtfromdate.Value, "yyyy-MM-dd") & "' and odkstartdate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' order by odkstartdate desc", ODKDB
End If
i = 1
Do While rs.EOF <> True
mygrid.rows = mygrid.rows + 1
mygrid.TextMatrix(i, 0) = i
FindsTAFF rs!staffcode
FindFA rs!farmercode, "F"
mygrid.TextMatrix(i, 1) = Format(rs!entrydate, "dd/MM/yyyy")
mygrid.TextMatrix(i, 2) = Format(rs!odkStartDate, "dd/MM/yyyy")
If ispercentage Then
mygrid.TextMatrix(i, 3) = Format(rs!odkValue, "####0.00") & "%"
Else
mygrid.TextMatrix(i, 3) = Format(rs!odkValue, "####0.00")
End If
mygrid.ColAlignment(3) = flexAlignRightTop
'Mygrid.TextMatrix(i, 4) = paramValue
mygrid.TextMatrix(i, 4) = rs!staffcode & " " & sTAFF
mygrid.ColAlignment(4) = flexAlignLeftTop
mygrid.TextMatrix(i, 5) = rs!farmercode & " " & FAName
mygrid.ColAlignment(5) = flexAlignLeftTop
mygrid.TextMatrix(i, 6) = IIf(rs!fieldcode = 0, "", rs!fieldcode)
mygrid.TextMatrix(i, 7) = rs!actiontaken
mygrid.TextMatrix(i, 8) = rs!recommendation
mygrid.TextMatrix(i, 9) = rs!qccomment
mygrid.TextMatrix(i, 11) = rs!uri
mygrid.TextMatrix(i, 12) = rs!emailstatus
i = i + 1
rs.MoveNext
Loop
'addcells
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub filldetail1(trnid As Integer)
On Error GoTo err
mchk = True
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
findParamDetails trnid
mygrid1.Clear
mygrid1.rows = 1
mygrid1.FormatString = "^SL.NO.|^ENTRY DATE|^START DATE|^" & paramName & "|^STAFF|^FARMER|^FD. CODE|^ACTION TAKEN|^RECOMMENDATION|^"
mygrid1.ColWidth(0) = 750
mygrid1.ColWidth(1) = 1320
mygrid1.ColWidth(2) = 1665
mygrid1.ColWidth(3) = 1125
'Mygrid.ColWidth(4) = 1365
mygrid1.ColWidth(4) = 2460
mygrid1.ColWidth(5) = 2820
mygrid1.ColWidth(6) = 960
mygrid1.ColWidth(7) = 1530
mygrid1.ColWidth(8) = 4395
mygrid1.ColWidth(9) = 120




If CHKSELECT.Value = 0 Then
rs.Open "select * from tblodkfollowuplog where paraid='" & trnid & "' and followupstatus='C' order by odkstartdate", ODKDB
Else
rs.Open "select * from tblodkfollowuplog where paraid='" & trnid & "' and followupstatus='C' and odkstartdate>='" & Format(txtfromdate.Value, "yyyy-MM-dd") & "' and odkstartdate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' order by odkstartdate", ODKDB
End If
i = 1
Do While rs.EOF <> True
mygrid1.rows = mygrid1.rows + 1
mygrid1.TextMatrix(i, 0) = i
FindsTAFF rs!staffcode
FindFA rs!farmercode, "F"
mygrid1.TextMatrix(i, 1) = Format(rs!entrydate, "dd/MM/yyyy")
mygrid1.TextMatrix(i, 2) = Format(rs!odkStartDate, "dd/MM/yyyy")
If ispercentage Then
mygrid1.TextMatrix(i, 3) = Format(rs!odkValue, "####0.00") & "%"
Else
mygrid1.TextMatrix(i, 3) = Format(rs!odkValue, "####0.00")
End If
mygrid1.ColAlignment(3) = flexAlignRightTop
'Mygrid.TextMatrix(i, 4) = paramValue
mygrid1.TextMatrix(i, 4) = rs!staffcode & " " & sTAFF
mygrid1.ColAlignment(4) = flexAlignLeftTop
mygrid1.TextMatrix(i, 5) = rs!farmercode & " " & FAName
mygrid1.ColAlignment(5) = flexAlignLeftTop
mygrid1.TextMatrix(i, 6) = IIf(rs!fieldcode, "", rs!fieldcode)
mygrid1.TextMatrix(i, 7) = rs!actiontaken
mygrid1.TextMatrix(i, 8) = rs!recommendation

i = i + 1
rs.MoveNext
Loop
'addcells
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Command1_Click()
If Len(cboparaid.Text) > 0 Then
filldetail cboparaid.BoundText
filldetail1 cboparaid.BoundText
End If
End Sub

Private Sub Command2_Click()
pestFarmer = txtfarmername.Text
pestparamname = cboparaid.Text
frmpesthistory.Show 1
End Sub

Private Sub Command4_Click()
Frame5.Visible = False
End Sub

Private Sub Form_Load()

Operation = ""

Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString

Set db1 = New ADODB.Connection
db1.CursorLocation = adUseClient
db1.Open CnnString
                     





Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open " select DISTINCT concat(cast(paraid as char),'  ',paraname,'  ',fstype,'  ',cast(value as char)) as description,paraid from tblodkalarmparameter where status='ON'", db
Set cboparaid.RowSource = rs
cboparaid.ListField = "description"
cboparaid.BoundColumn = "paraid"
End Sub

Private Sub mygrid_Click()
Frame5.Visible = False





If mygrid.col = 7 Or mygrid.col = 8 Or mygrid.col = 9 Then
mygrid.Editable = flexEDKbdMouse
Else
mygrid.Editable = flexEDNone
End If
End Sub

Private Sub mygrid_DblClick()
If mygrid.col = 5 And Len(mygrid.TextMatrix(mygrid.row, 5)) > 0 Then
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblfarmer   where idfarmer='" & Mid(mygrid.TextMatrix(mygrid.row, 5), 1, 14) & "'", MHVDB
If rs.EOF <> True Then
txtfarmername.Text = ""
txtdz.Text = ""
TXTAREA.Text = ""
txtge.Text = ""
txtts.Text = ""
TXTSTATUS.Text = ""
TXTPLANTS.Text = ""
Frame5.Visible = True
txtfarmername.Text = rs!idfarmer & " " & rs!farmername
FindDZ Mid(rs!idfarmer, 1, 3)
FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
txtdz.Text = Mid(rs!idfarmer, 1, 3) & " " & Dzname
txtge.Text = Mid(rs!idfarmer, 4, 3) & " " & GEname
txtts.Text = Mid(rs!idfarmer, 7, 3) & " " & TsName
If rs!status = "A" Then
TXTSTATUS.Text = "ACTIVE"
ElseIf rs!status = "R" Then
TXTSTATUS.Text = "REJECTED"
Else
TXTSTATUS.Text = "DROPPED OUT"
End If
Set rs1 = Nothing
rs1.Open "select sum(regland) as land from tbllandreg where farmerid='" & rs!idfarmer & "'", MHVDB
If rs1.EOF <> True Then
TXTAREA.Text = Format(rs1!land, "###0.00")
End If

Set rs1 = Nothing
rs1.Open "select sum(nooftrees) as tr from tblplanted where farmercode='" & rs!idfarmer & "'", MHVDB
If rs1.EOF <> True Then
TXTPLANTS = IIf(IsNull(rs1!tr), "", rs1!tr)
End If


End If
End If
End Sub

Private Sub MYGRID1_Click()
Frame5.Visible = False
End Sub

Private Sub mygrid1_DblClick()
If mygrid1.col = 5 And Len(mygrid1.TextMatrix(mygrid1.row, 5)) > 0 Then
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblfarmer   where idfarmer='" & Mid(mygrid1.TextMatrix(mygrid1.row, 5), 1, 14) & "'", MHVDB
If rs.EOF <> True Then
txtfarmername.Text = ""
txtdz.Text = ""
TXTAREA.Text = ""
txtge.Text = ""
txtts.Text = ""
TXTSTATUS.Text = ""
TXTPLANTS.Text = ""
Frame5.Visible = True
txtfarmername.Text = rs!idfarmer & " " & rs!farmername
FindDZ Mid(rs!idfarmer, 1, 3)
FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
txtdz.Text = Mid(rs!idfarmer, 1, 3) & " " & Dzname
txtge.Text = Mid(rs!idfarmer, 4, 3) & " " & GEname
txtts.Text = Mid(rs!idfarmer, 7, 3) & " " & TsName
If rs!status = "A" Then
TXTSTATUS.Text = "ACTIVE"
ElseIf rs!status = "R" Then
TXTSTATUS.Text = "REJECTED"
Else
TXTSTATUS.Text = "DROPPED OUT"
End If
Set rs1 = Nothing
rs1.Open "select sum(regland) as land from tbllandreg where farmerid='" & rs!idfarmer & "'", MHVDB
If rs1.EOF <> True Then
TXTAREA.Text = Format(rs1!land, "###0.00")
End If

Set rs1 = Nothing
rs1.Open "select sum(nooftrees) as tr from tblplanted where farmercode='" & rs!idfarmer & "'", MHVDB
If rs1.EOF <> True Then
TXTPLANTS = rs1!tr
End If


End If
End If
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

        Case "OPEN"
        Operation = "OPEN"
        CLEARCONTROLL
        cboparaid.Enabled = True
        cboparaid.Text = ""
        TB.buttons(3).Enabled = True
             
       Case "SAVE"
        MNU_SAVE
     If Len(cboparaid.Text) > 0 Then
        filldetail cboparaid.BoundText
         filldetail1 cboparaid.BoundText
     End If
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
Dim rsp As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim i As Integer
Dim j As Integer
On Error GoTo err


MHVDB.BeginTrans
If Operation = "OPEN" Then

LogRemarks = ""
For i = 1 To mygrid.rows - 1
If mygrid.ValueMatrix(i, 7) * -1 = 1 Then
ODKDB.Execute "update tblodkfollowuplog set actiontaken='1'," _
            & "recommendation='" & mygrid.TextMatrix(i, 8) & "',followupstatus='C'" _
            & "where uri= '" & mygrid.TextMatrix(i, 11) & "' "
            
Else
If Len(mygrid.TextMatrix(i, 9)) > 0 Then

ODKDB.Execute "update tblodkfollowuplog set " _
            & "qccomment='" & mygrid.TextMatrix(i, 9) & "'" _
            & "where uri= '" & mygrid.TextMatrix(i, 11) & "' "
End If




End If


Next



'updatemhvlog Now, MUSER, LogRemarks, ""


Else
MsgBox "OPERATION NOT SELECTED."
End If
MHVDB.CommitTrans
TB.buttons(3).Enabled = False
Exit Sub

err:
MsgBox err.Description
MHVDB.RollbackTrans


End Sub
Private Sub CLEARCONTROLL()
mygrid.Clear
mygrid.rows = 1
mygrid.FormatString = "^SL.NO.|^ENTRY DATE|^START DATE|^" & paramName & "|^STAFF|^FARMER|^FD. CODE|^ACTION TAKEN|^RECOMMENDATION|^QC COMMENT|^uri|^email|^"
mygrid.ColWidth(0) = 750
mygrid.ColWidth(1) = 1320
mygrid.ColWidth(2) = 1665
mygrid.ColWidth(3) = 1125
'Mygrid.ColWidth(4) = 1365
mygrid.ColWidth(4) = 2460
mygrid.ColWidth(5) = 2820
mygrid.ColWidth(6) = 960
mygrid.ColWidth(7) = 1530
mygrid.ColWidth(8) = 1980
mygrid.ColWidth(9) = 2250
mygrid.ColWidth(10) = 120
mygrid.ColWidth(11) = 1     ' URI
mygrid.ColWidth(12) = 1     ' EMAIL

mygrid1.Clear
mygrid1.rows = 1
mygrid1.FormatString = "^SL.NO.|^ENTRY DATE|^START DATE|^ODK VALUE|^STAFF|^FARMER|^FD. CODE|^ACTION TAKEN|^RECOMMENDATION|^"
mygrid1.ColWidth(0) = 750
mygrid1.ColWidth(1) = 1320
mygrid1.ColWidth(2) = 1665
mygrid1.ColWidth(3) = 1125
mygrid1.ColWidth(4) = 2460
mygrid1.ColWidth(5) = 2820
mygrid1.ColWidth(6) = 960
mygrid1.ColWidth(7) = 1530
mygrid1.ColWidth(8) = 4395
mygrid1.ColWidth(9) = 120
End Sub
