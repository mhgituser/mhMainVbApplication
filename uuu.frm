VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCrateBatchTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEND TO FIELD"
   ClientHeight    =   8715
   ClientLeft      =   2895
   ClientTop       =   690
   ClientWidth     =   15465
   Icon            =   "uuu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   15465
   Begin VB.TextBox txtselected 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   4080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   6360
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sent To Field Transaction"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   7455
      Begin VB.TextBox txtqty 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtnoofcrates 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3360
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataListLib.DataCombo cbofacility 
         Bindings        =   "uuu.frx":0E42
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1320
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
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
      Begin MSDataListLib.DataCombo cboplantbatch 
         Bindings        =   "uuu.frx":0E57
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   480
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
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
      Begin MSFlexGridLib.MSFlexGrid ItemGrd 
         Height          =   3540
         Left            =   0
         TabIndex        =   19
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6244
         _Version        =   393216
         Rows            =   50
         Cols            =   8
         RowHeightMin    =   315
         ForeColorFixed  =   -2147483635
         ScrollTrack     =   -1  'True
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         FormatString    =   "       |^ Batch No.|^Variety |^Facility |^No. Of Crates|^ Crate #                 |Qty.      |"
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
         Caption         =   "Remarks :"
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
         Index           =   8
         Left            =   240
         TabIndex        =   20
         Top             =   4920
         Width           =   870
      End
      Begin VB.Line Line3 
         X1              =   5280
         X2              =   8880
         Y1              =   4290
         Y2              =   4290
      End
   End
   Begin VB.TextBox txtshortexcees 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   8160
      Width           =   1095
   End
   Begin VB.TextBox txtsendtofield 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   13
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox txtdsheetqty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   8160
      Width           =   855
   End
   Begin VB.TextBox txttocrate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   "Distribution Sheet Detail"
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
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Width           =   3855
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3615
         _cx             =   6376
         _cy             =   3625
         _ConvInfo       =   1
         Appearance      =   0
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
         FormatString    =   $"uuu.frx":0E6C
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
      Caption         =   "Crate Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   7680
      TabIndex        =   3
      Top             =   720
      Width           =   7695
      Begin VB.Frame frmaddcrate 
         Caption         =   "Add Crate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   5280
         TabIndex        =   32
         Top             =   5760
         Visible         =   0   'False
         Width           =   2295
         Begin VB.CommandButton Command6 
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
            Left            =   1200
            Picture         =   "uuu.frx":0EEA
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtcrateno 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1080
            TabIndex        =   33
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Crate No."
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
            TabIndex        =   35
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.CommandButton Command5 
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
         Left            =   6480
         Picture         =   "uuu.frx":1694
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   7320
         Width           =   1095
      End
      Begin VB.TextBox txtfind 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   30
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtcratecnt 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ok"
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
         Picture         =   "uuu.frx":1F5E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7320
         Width           =   975
      End
      Begin VB.ListBox DZLIST 
         Columns         =   8
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6585
         ItemData        =   "uuu.frx":22E8
         Left            =   120
         List            =   "uuu.frx":22EA
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   600
         Width           =   7455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7455
      Begin VB.TextBox txtdriver 
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
         Left            =   5040
         TabIndex        =   42
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox txtloc 
         Height          =   315
         ItemData        =   "uuu.frx":22EC
         Left            =   6000
         List            =   "uuu.frx":22FC
         TabIndex        =   38
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtyr 
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
         Height          =   360
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "uuu.frx":2314
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
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
      Begin MSComCtl2.DTPicker txtentrydate 
         Height          =   375
         Left            =   6000
         TabIndex        =   25
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   115343361
         CurrentDate     =   41479
      End
      Begin MSDataListLib.DataCombo cbostaff 
         Bindings        =   "uuu.frx":2329
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1440
         TabIndex        =   28
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
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
      Begin MSDataListLib.DataCombo cbovehicle 
         Bindings        =   "uuu.frx":233E
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1440
         TabIndex        =   41
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "Driver"
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
         Left            =   4200
         TabIndex        =   40
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle"
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
         TabIndex        =   39
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Station"
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
         Left            =   5280
         TabIndex        =   37
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Staff"
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
         TabIndex        =   27
         Top             =   720
         Width           =   480
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
         Index           =   0
         Left            =   5280
         TabIndex        =   26
         Top             =   360
         Width           =   510
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Distribution No."
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
         Width           =   1335
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
            Picture         =   "uuu.frx":2353
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "uuu.frx":26ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "uuu.frx":2A87
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "uuu.frx":3761
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "uuu.frx":3BB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "uuu.frx":436D
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
      Width           =   15465
      _ExtentX        =   27279
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Short/Excess"
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
      Left            =   4920
      TabIndex        =   15
      Top             =   8280
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Sent To Field Qty."
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
      Left            =   4680
      TabIndex        =   12
      Top             =   5880
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "D. Sheet Qty."
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
      Left            =   1680
      TabIndex        =   10
      Top             =   8280
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Crates Selected"
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
      Left            =   4680
      TabIndex        =   8
      Top             =   7800
      Width           =   1365
   End
End
Attribute VB_Name = "frmCrateBatchTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totdsheet, totsendtofield, totcrate As Double
Dim ValidRow As Boolean
Dim CurrRow As Long
Private Sub filldgrid()
Dim rs As New ADODB.Recordset

End Sub

Private Sub cbofacility_LostFocus()
Dim Issue, Recv As Double
Dim rs As New ADODB.Recordset
ItemGrd.TextMatrix(CurrRow, 3) = cbofacility.BoundText
cbofacility.Visible = False
ItemGrd.ColWidth(3) = 750
End Sub

Private Sub cbofacility_Validate(Cancel As Boolean)
ItemGrd.TextMatrix(CurrRow, 3) = cbofacility.BoundText
   cbofacility.Visible = False
End Sub

Private Sub cboplantBatch_LostFocus()
Dim Issue, Recv As Double
Dim rs As New ADODB.Recordset
 ItemGrd.TextMatrix(CurrRow, 1) = cboplantbatch.BoundText
   cboplantbatch.Visible = False

'datInvItem.Recordset.FindFirst "itemcode = '" & CboItemDesc.BoundText & "'"
  'rsdatInvItem.Find " itemcode='" & CboItemDesc.BoundText & "'", , adSearchForward, 1
  Set rs = Nothing
  rs.Open "select * from tblqmsplantbatchdetail where plantbatch='" & ItemGrd.TextMatrix(CurrRow, 1) & "'", MHVDB
With rs
If .EOF Then
   ItemGrd.TextMatrix(CurrRow, 1) = ""
   ItemGrd.TextMatrix(CurrRow, 2) = ""
   ValidRow = True
Else
findQmsBatchDetail rs!plantBatch
   'ItemGrd.TextMatrix(CurrRow, 1) = !varietyid
   ItemGrd.TextMatrix(CurrRow, 2) = qmsplantbatch3 '!Description


End If
End With
cboplantbatch.Visible = False
End Sub

Private Sub cbotrnid_LostFocus()
On Error GoTo err
cbotrnid.Enabled = False
txtyr.Text = Mid(cbotrnid.Text, Len(cbotrnid.Text) - 4, 5)
txtyr.Text = Trim(txtyr.Text)
cbotrnid.Text = cbotrnid.BoundText

TB.buttons(3).Enabled = True

Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmssendtofieldhdr where distributionno='" & cbotrnid.BoundText & "'", MHVDB
If rs.EOF <> True Then
FindsTAFF rs!staffid
cbostaff.Text = rs!staffid & "  " & sTAFF
txtentrydate.Value = Format(rs!entrydate, "dd/MM/yyyy")
txtsendtofield.Text = rs!sendtofieldqty

txtdsheetqty.Text = rs!dsheetqty
txttocrate.Text = rs!cratecount
txtloc.Text = rs!location
cbovehicle.Text = rs!vehicleno
txtdriver.Text = rs!DriverName
End If


filldsheet txtyr.Text
fillsendtofield
txtshortexcees.Text = Val(txtsendtofield.Text) - Val(txtdsheetqty.Text)
Frame2.Enabled = True
Exit Sub
err:
    cbotrnid.Text = ""
    cbotrnid.Enabled = True
    MsgBox "Invalid selection of distribution no."
End Sub
Private Sub fillsendtofield()
Dim rs As New ADODB.Recordset
Dim i As Integer
i = 1
ItemGrd.Clear
ItemGrd.FormatString = "       |^ Batch No.|^Variety |^Facility |^No. Of Crates|^ Crate #                 |Qty.      |"
Set rs = Nothing
rs.Open "select * from tblqmsplanttransaction where transactiontype='4'and  distributionno='" & cbotrnid.BoundText & "' and status='ON'", MHVDB
Do While rs.EOF <> True
findQmsBatchDetail rs!plantBatch
ItemGrd.TextMatrix(i, 1) = rs!plantBatch
ItemGrd.TextMatrix(i, 2) = qmsplantbatch3
ItemGrd.TextMatrix(i, 3) = rs!facilityid
ItemGrd.TextMatrix(i, 4) = rs!cratecount
ItemGrd.TextMatrix(i, 5) = rs!crateno
ItemGrd.TextMatrix(i, 6) = rs!credit

i = i + 1
rs.MoveNext

Loop

End Sub
Private Sub filldsheet(myear As Integer)
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim i As Integer
Dim j As Integer
j = 0
i = 1
totdsheet = 0
mygrid.Clear
mygrid.FormatString = "^SL.NO.|^PLANT VARIETY|^QTY.|"
mygrid.ColWidth(0) = 750
mygrid.ColWidth(1) = 1665
mygrid.ColWidth(2) = 840
mygrid.ColWidth(3) = 555

If Len(cbotrnid.Text) = 0 Then Exit Sub


mygrid.rows = 1
Set rs = Nothing
rs.Open "Select * from tblqmsplantvariety where status<>'C' and varietyid in (1,2,4,7,12)", MHVDB
Do While rs.EOF <> True
mygrid.rows = mygrid.rows + 1
mygrid.TextMatrix(i, 0) = i
mygrid.TextMatrix(i, 1) = "V" & i
i = i + 1
rs.MoveNext
Loop



Set rs = Nothing
'Select Case myear
'Case 2011, 2012, 2013, 2014
'rs.Open "SELECT SUM( bcrate*35 ) AS B, SUM( ecrate*35 ) AS E, SUM( bno*35 ) AS P, SUM( plno*plnofactor ) AS P1, SUM( crate *35 ) AS N FROM  `tblplantdistributiondetail` WHERE  distno='" & cbotrnid.BoundText & "' and year='" & Val(txtyr.Text) & "' and subtotindicator='' and status not in ('C','F') and distno>0 and trnid in (select trnid from tblplantdistributionheader where status='ON') GROUP BY distno", MHVDB
'Case 2015, 2016
'rs.Open "SELECT SUM( bcrate*bcratefactor ) AS B, SUM( ecrate*ecratefactor ) AS E, SUM( bno*bnofactor ) AS P, SUM( plno *plnofactor) AS P1, SUM( crate*cratefactor) AS N FROM  `tblplantdistributiondetail` WHERE  distno='" & cbotrnid.BoundText & "' and year='" & Val(txtyr.Text) & "' and subtotindicator='' and status not in ('C','F') and distno>0 and trnid in (select trnid from tblplantdistributionheader where status='ON') GROUP BY distno", MHVDB
'
'End Select
rs.Open "SELECT SUM( bcrate*bcratefactor ) AS V1, SUM( ecrate*ecratefactor ) AS V2, SUM( bno*bnofactor ) AS V3, SUM( plno *plnofactor) AS P1, SUM( crate*cratefactor) AS V4 FROM  `tblplantdistributiondetail` WHERE  distno='" & cbotrnid.BoundText & "' and year='" & Val(txtyr.Text) & "' and subtotindicator='' and status not in ('C','F') and distno>0 and trnid in (select trnid from tblplantdistributionheader where status='ON') GROUP BY distno", MHVDB

For i = 1 To mygrid.rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
For j = 0 To 4

If Trim(mygrid.TextMatrix(i, 1)) = rs.Fields(j).name Then
'Mygrid.Rows = Mygrid.Rows + 1
If rs.EOF <> True Then
mygrid.TextMatrix(i, 1) = rs.Fields(j).name
mygrid.TextMatrix(i, 2) = IIf(IsNull(rs.Fields(j).Value), "", rs.Fields(j).Value)
mygrid.ColAlignment(2) = flexAlignRightTop
totdsheet = totdsheet + rs.Fields(j).Value
Else
 TB.buttons(3).Enabled = False
End If
End If

Next
Next
txtdsheetqty.Text = totdsheet
End Sub

Private Sub Command2_Click()
Dim crtcnt As Integer
crtcnt = 0
Dim i As Integer
Dzstr = ""
For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       Dzstr = Dzstr + Trim(DZLIST.List(i)) + ","
       crtcnt = crtcnt + 1
         End If
    Next
If Len(Dzstr) > 0 Then
If Val(ItemGrd.TextMatrix(CurrRow, 4)) = crtcnt Then


   Dzstr = Left(Dzstr, Len(Dzstr) - 1)
   ItemGrd.TextMatrix(CurrRow, 5) = ""
 ItemGrd.TextMatrix(CurrRow, 5) = Dzstr
frmCrateBatchTransaction.Width = 7770
TB.buttons(3).Enabled = True
 ItemGrd.Enabled = True
 Else
 MsgBox "Crate count does not match."
 'ItemGrd.TextMatrix(CurrRow, 5) = ""
 loadcrate
 frmCrateBatchTransaction.Width = 7770
 TB.buttons(3).Enabled = True
  ItemGrd.Enabled = True
 End If
 
 
 
Else
   MsgBox "CRATE NOT SELECTED !!!"
   frmCrateBatchTransaction.Width = 7770
   TB.buttons(3).Enabled = True
     ItemGrd.Enabled = True
   Exit Sub
End If
End Sub

Private Sub cratecnt()
Dim i As Integer
Dim mcrate As String
Dim munselected As String
mcratecnt = 0
mcrate = ""
munselected = ""
For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
      mcratecnt = mcratecnt + 1
     'If InStr(1, txtselected.Text, DZLIST.List(i), vbTextCompare) = 0 Then
       mcrate = mcrate + Trim(DZLIST.List(i)) + ","
      'End If
    End If


        
    Next
    If Len(mcrate) > 0 Then
     mcrate = Left(mcrate, Len(mcrate) - 1)
    End If
    
    txtcratecnt.Text = mcratecnt
    If Val(txtcratecnt.Text) > 0 Then
        
       txtselected.Text = mcrate
        
        
    txtselected.Visible = True
    Else
    
    txtselected.Visible = False
    End If
End Sub


Private Sub Command5_Click()
frmaddcrate.Visible = True
txtcrateno.Text = ""
End Sub

Private Sub Command6_Click()
Dim rs As New ADODB.Recordset

If Len(Trim(txtcrateno.Text)) = 0 Then
frmaddcrate.Visible = False
Exit Sub
End If


Set rs = Nothing
rs.Open "select * from tblqmscrate where crateno='" & Trim(txtcrateno.Text) & "' ", MHVDB
If rs.EOF <> True Then
MsgBox "Crate Already Exists."
Else

MHVDB.Execute "insert into tblqmscrate(crateno) values('" & Trim(txtcrateno.Text) & "')"
End If
txtcrateno.Text = ""
frmaddcrate.Visible = False
Command2_Click
Frame3.Visible = False
End Sub

Private Sub DZLIST_Click()
cratecnt
End Sub

Private Sub Form_Load()
Dim cnt As Integer
Dim rs As New ADODB.Recordset
Dim i, j As Integer
Dim dd As Variant
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
                   
db.Open CnnString
frmCrateBatchTransaction.Width = 7770
Set rs = Nothing
i = 1
DZLIST.Clear
rs.Open "select * from tblqmscrate Order by  crateno", MHVDB
With rs
Do While Not .EOF

DZLIST.AddItem Trim(!crateno)

   .MoveNext
  Loop
End With




Set rs = Nothing

If rs.State = adStateOpen Then rs.Close
rs.Open "select distinct distno,concat(cast(distno as char) , '  ', cast(year as char)) dist  from tblplantdistributiondetail where  subtotindicator='' and status not in ('C','F') and distno>0 and trnid in (select trnid from tblplantdistributionheader where status='ON' and planneddist='Y')order by distno desc", db
Set cbotrnid.RowSource = rs
cbotrnid.ListField = "dist"
cbotrnid.BoundColumn = "distno"




Set rs = Nothing

If rs.State = adStateOpen Then rs.Close
rs.Open "select plantbatch  from tblqmsplantbatchdetail where plantbatch>0 order by plantbatch desc", db
Set cboplantbatch.RowSource = rs
cboplantbatch.ListField = "plantbatch"
cboplantbatch.BoundColumn = "plantbatch"

Set rs = Nothing

If rs.State = adStateOpen Then rs.Close
rs.Open "select facilityid,description  from tblqmsfacility where status='ON' order by facilityid", db
Set cbofacility.RowSource = rs
cbofacility.ListField = "description"
cbofacility.BoundColumn = "facilityid"

Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open "select concat(STAFFCODE , '  ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff WHERE nursery='1'  order by STAFFCODE", db
Set cbostaff.RowSource = rs
cbostaff.ListField = "STAFFNAME"
cbostaff.BoundColumn = "STAFFCODE"


Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open "select dcmno as dcmno,dcmno as id  from tbldistexternaldcm WHERE length(dcmno)>4  order by dcmno", db
Set cbovehicle.RowSource = rs
cbovehicle.ListField = "dcmno"
cbovehicle.BoundColumn = "id"

'i = 1


ValidRow = True
CurrRow = 1
ItemGrd.row = 1
ItemGrd.col = 1
cboplantbatch.Left = ItemGrd.Left + ItemGrd.CellLeft
cboplantbatch.Width = ItemGrd.CellWidth
cboplantbatch.Height = ItemGrd.CellHeight

ItemGrd.col = 3
cbofacility.Left = ItemGrd.Left + ItemGrd.CellLeft
cbofacility.Width = 2000 'ItemGrd.CellWidth
cbofacility.Height = ItemGrd.CellHeight

ItemGrd.col = 4
txtnoofcrates.Left = ItemGrd.Left + ItemGrd.CellLeft
txtnoofcrates.Width = ItemGrd.CellWidth
txtnoofcrates.Height = ItemGrd.CellHeight

ItemGrd.col = 6
txtqty.Left = ItemGrd.Left + ItemGrd.CellLeft
txtqty.Width = ItemGrd.CellWidth
txtqty.Height = ItemGrd.CellHeight

End Sub


Private Sub ItemGrd_Click()
Dim mrow, MCOL As Integer
txtselected.Visible = False
ItemGrd.ColWidth(3) = 750
If Not ValidRow And CurrRow <> ItemGrd.row Then
   ItemGrd.row = CurrRow
   Exit Sub
End If
mrow = ItemGrd.row
MCOL = ItemGrd.col
If mrow = 0 Then Exit Sub
If mrow > 1 And Len(ItemGrd.TextMatrix(mrow - 1, 4)) = 0 Then
   Beep
   Exit Sub
End If
ItemGrd.TextMatrix(CurrRow, 0) = CurrRow
CurrRow = mrow
ItemGrd.TextMatrix(CurrRow, 0) = Chr(174)

Select Case MCOL
       
       Case 1
       If Len(ItemGrd.TextMatrix(mrow - 1, 5)) > 0 Then
            cboplantbatch.Top = ItemGrd.Top + ItemGrd.CellTop
            cboplantbatch = ItemGrd.Text
            cboplantbatch.Visible = True
            cboplantbatch.SetFocus
            End If
       Case 3
       If Trim(Len(ItemGrd.TextMatrix(mrow, 2))) > 0 Then
            ItemGrd.ColWidth(3) = 2000
            cbofacility.Top = ItemGrd.Top + ItemGrd.CellTop
            cbofacility = ItemGrd.Text
            cbofacility.Visible = True
            cbofacility.SetFocus
            End If
            
            Dim rs As New ADODB.Recordset
            Set db = New ADODB.Connection
            db.CursorLocation = adUseClient
            Dim CONNLOCAL As New ADODB.Connection
            CONNLOCAL.Open OdkCnnString
                         
            db.Open CnnString
            
            Set rs = Nothing
            
            If rs.State = adStateOpen Then rs.Close
            rs.Open "select distinct a.facilityid,a.description  from tblqmsfacility a ,tblqmsplanttransaction b where a.facilityid=b.facilityid and plantbatch='" & Trim(ItemGrd.TextMatrix(mrow, 1)) & "' and a.status='ON' order by facilityid", db
            Set cbofacility.RowSource = rs
            cbofacility.ListField = "description"
            cbofacility.BoundColumn = "facilityid"
            
            
            
       Case 4
            If Len(ItemGrd.TextMatrix(mrow, 1)) > 0 And Len(ItemGrd.TextMatrix(mrow, 5)) = 0 Then
               txtnoofcrates.Top = ItemGrd.Top + ItemGrd.CellTop
               txtnoofcrates = ItemGrd.Text
               txtnoofcrates.Visible = True
               txtnoofcrates.SetFocus
            End If
       Case 6
'       If Len(ItemGrd.TextMatrix(mrow, 1)) > 0 And InStr(1, ItemGrd.TextMatrix(mrow, 5), ",") = 0 Then
'            txtqty.Top = ItemGrd.Top + ItemGrd.CellTop
'            txtqty = ItemGrd.Text
'            txtqty.Visible = False
'            'txtqty.SetFocus
'       End If
    Case 5
      If ItemGrd.col = 5 And Val(ItemGrd.TextMatrix(CurrRow, 4)) > 0 And Len(ItemGrd.TextMatrix(CurrRow, 5)) = 0 Then
            loadcrate
            frmCrateBatchTransaction.Width = 15525
            TB.buttons(3).Enabled = False
            ItemGrd.Enabled = False
            
            Else
            frmCrateBatchTransaction.Width = 7770
            TB.buttons(3).Enabled = True
            txtselected.Visible = True
            txtselected.Text = ItemGrd.TextMatrix(CurrRow, 5)
            End If
            
            If Len(txtselected.Text) > 15 And ItemGrd.col = 5 And Len(ItemGrd.TextMatrix(CurrRow, 5)) > 0 Then
            txtselected.Visible = True
            txtselected.Text = ItemGrd.TextMatrix(CurrRow, 5)
            Else
            txtselected.Visible = False
            End If
            
            
            
    End Select
End Sub
Private Sub loadcrate()
Dim cnt As Integer
Dim rs As New ADODB.Recordset
Dim i, j As Integer
Dim dd As Variant
Set rs = Nothing

DZLIST.Clear
rs.Open "select * from tblqmscrate where locked='0' Order by  cast(crateno as unsigned)", MHVDB
With rs
Do While Not .EOF
   DZLIST.AddItem Trim(!crateno)
   .MoveNext
Loop
End With
If Len(ItemGrd.TextMatrix(CurrRow, 5)) > 0 Then
dd = Split(ItemGrd.TextMatrix(CurrRow, 5), ",", -1, vbTextCompare)
'dd = Split("ItemGrd.TextMatrix(CurrRow, 5)", ",")
cnt = Len(ItemGrd.TextMatrix(CurrRow, 5)) - Len(Replace(ItemGrd.TextMatrix(CurrRow, 5), ",", ""))
'Len(x) - Len(Replace(x, ",", ""))
For j = 0 To cnt
For i = 0 To DZLIST.ListCount - 1

If dd(j) = Trim(DZLIST.List(i)) Then
DZLIST.Selected(i) = True
End If

'    If DZLIST.Selected(i) Then
'       DZstr = DZstr + Trim(DZLIST.List(i)) + ","
'         End If
    Next
Next


End If


End Sub

Private Sub ItemGrd_DblClick()
If Not ValidRow And CurrRow <> ItemGrd.row Then
   ItemGrd.row = CurrRow
   Exit Sub
End If
mrow = ItemGrd.row
MCOL = ItemGrd.col
CurrRow = mrow

If ItemGrd.col = 5 And Val(ItemGrd.TextMatrix(CurrRow, 4)) > 0 And Len(ItemGrd.TextMatrix(CurrRow, 5)) > 0 Then
            loadcrate
            frmCrateBatchTransaction.Width = 15525
            TB.buttons(3).Enabled = False
            ItemGrd.Enabled = False
            End If
            
   If ItemGrd.col = 4 And Val(ItemGrd.TextMatrix(CurrRow, 4)) > 0 And Len(ItemGrd.TextMatrix(CurrRow, 5)) > 0 Then
   If MsgBox("Are you sure to clear the Crate Selected?", vbYesNo) = vbYes Then
   ItemGrd.TextMatrix(CurrRow, 4) = ""
   ItemGrd.TextMatrix(CurrRow, 5) = ""
   ItemGrd.TextMatrix(CurrRow, 6) = ""
   End If
   End If
   
   
  If ItemGrd.col = 6 And Len(ItemGrd.TextMatrix(ItemGrd.row, 4)) > 0 And (ItemGrd.TextMatrix(ItemGrd.row, 2) = "P1" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "P" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "N" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "B" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "E" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "L" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "A" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "L1" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "D1" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "TDG") Then



myinput = InputBox("Enter No. of plants of variety " & ItemGrd.TextMatrix(ItemGrd.row, 2))
            If Not IsNumeric(myinput) Then
            MsgBox "Invalid number,Double Click again to enable the input box."
            Else
            ItemGrd.TextMatrix(ItemGrd.row, 6) = CInt(myinput)
            getsum
            End If

End If
          

   
   
getsum
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       Operation = "ADD"
       populatedno "ADD"
       CLEARCONTROLL
       cbotrnid.Enabled = True
       TB.buttons(3).Enabled = True
       Case "OPEN"
       Operation = "OPEN"
       populatedno "OPEN"
       CLEARCONTROLL
       cbotrnid.Enabled = True
       TB.buttons(3).Enabled = True
       
       Case "SAVE"
       MNU_SAVE
        TB.buttons(3).Enabled = False
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
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim dd As Variant
Dim mm As Variant
Dim bb As Variant
Dim cnt As Integer
Dim mMaxId As Double
Dim crateStr As String
Dim i As Integer
If Len(cbotrnid.Text) = 0 Then
MsgBox "Distribution No. cannot be empty."
Exit Sub
End If

If Len(txtloc.Text) = 0 Then
MsgBox "Select Station."
Exit Sub
End If


If Val(txtyr.Text) <= 0 Then
MsgBox "Invalid distribution no."
Exit Sub
End If

If Len(cbovehicle.Text) = 0 Then
MsgBox "Please select vehicle from the list."
Exit Sub
End If

If Val(txtsendtofield.Text) <= 0 Then
MsgBox "Invalid sent to field quantity."
Exit Sub
End If


    Set rs = Nothing
        rs.Open "SELECT MAX(trnid)+1 AS MaxID from tblqmsplanttransaction", MHVDB, adOpenForwardOnly, adLockOptimistic
        If rs.EOF <> True Then
        mMaxId = IIf(IsNull(rs!MaxId), "1", rs!MaxId)
        Else
        mMaxId = rs!MaxId
        End If

MHVDB.BeginTrans

 crateStr = ""
 For i = 1 To ItemGrd.rows - 1
 If Len(Trim(ItemGrd.TextMatrix(i, 1))) = 0 Then Exit For
 mm = Split(Trim(ItemGrd.TextMatrix(i, 5)), ",", -1, vbTextCompare)
cnt = Len(Trim(ItemGrd.TextMatrix(i, 5))) - Len(Replace(Trim(ItemGrd.TextMatrix(i, 5)), ",", ""))
For j = 0 To cnt
crateStr = Trim(ItemGrd.TextMatrix(i, 1)) & "|" & mm(j) & "," & crateStr
Next
 
 Next
 crateStr = Left(crateStr, Len(crateStr) - 1)
 
If Operation = "ADD" Then
MHVDB.Execute "insert into tblqmssendtofieldhdr(distributionno,entrydate,staffid" _
& " ,sendtofieldqty,dsheetqty,shortexcessqty,vehicleno,drivername,status,cratecount,year,location) values(" _
& "'" & cbotrnid.BoundText & "'," _
& "'" & Format(txtentrydate.Value, "yyyy-MM-dd") & "'," _
& "'" & cbostaff.BoundText & "'," _
& "'" & Val(txtsendtofield.Text) & "'," _
& "'" & Val(txtdsheetqty.Text) & "'," _
& "'" & Val(txtshortexcees.Text) & "'," _
& "'" & cbovehicle.BoundText & "'," _
& "'" & txtdriver.Text & "'," _
& "'ON'," _
& "'" & Val(txttocrate.Text) & "','" & Val(txtyr.Text) & "','" & txtloc.Text & "')"

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblqmssendtofieldhdr set " _
& "entrydate='" & Format(txtentrydate.Value, "yyyy-MM-dd") & "'," _
& "staffid='" & cbostaff.BoundText & "'," _
& "sendtofieldqty='" & Val(txtsendtofield.Text) & "'," _
& "dsheetqty='" & Val(txtdsheetqty.Text) & "'," _
& "shortexcessqty='" & Val(txtshortexcees.Text) & "'," _
& "vehicleno='" & cbovehicle.BoundText & "'," _
& "drivername='" & txtdriver.Text & "'," _
& "status='ON'," _
& "location='" & txtloc.Text & "'," _
& "cratecount='" & Val(txttocrate.Text) & "' where distributionno='" & cbotrnid.BoundText & "' and year='" & Val(txtyr.Text) & "'"
Else
MsgBox "No Operation Selected."
Exit Sub
End If





Set rs = Nothing
rs.Open "select * from tblqmsbacktonurseryhdr where distributionno='" & cbotrnid.BoundText & "' and year='" & Val(txtyr.Text) & "'", MHVDB
If rs.EOF <> True Then
Else
MHVDB.Execute "delete from tblqmssendtofielddetail where distributionno='" & cbotrnid.BoundText & "' and year='" & Val(txtyr.Text) & "'"

dd = Split(crateStr, ",", -1, vbTextCompare)
cnt = Len(crateStr) - Len(Replace(crateStr, ",", ""))
For j = 0 To cnt
bb = Split(dd(j), "|", -1, vbTextCompare)
MHVDB.Execute "insert into tblqmssendtofielddetail(distributionno,plantbatch," _
& "crateno,cratestatus,year)values(" _
& "'" & cbotrnid.BoundText & "'," _
& "'" & bb(0) & "'," _
& "'" & Val(bb(1)) & "'," _
& "'ON','" & Val(txtyr.Text) & "')"

MHVDB.Execute "update tblqmscrate set lasttrnno='" & cbotrnid.BoundText & "', lasttrnyear='" & Val(txtyr.Text) & "',lasttrntype='DIS' where crateno='" & Val(bb(1)) & "'"

Next
End If


 
MHVDB.Execute "delete from tblqmsplanttransaction where distributionno='" & cbotrnid.BoundText & "' and distyear='" & Val(txtyr.Text) & "' " _
& "and verificationtype='2' and transactiontype='4'"

For i = 1 To ItemGrd.rows - 1
If Len(Trim(ItemGrd.TextMatrix(i, 1))) = 0 Then Exit For
findQmsBatchDetail Trim(ItemGrd.TextMatrix(i, 1))
MHVDB.Execute "INSERT INTO tblqmsplanttransaction (trnid,entrydate,plantbatch,varietyid," _
            & "facilityid,verificationtype,transactiontype,staffid,debit,credit,status,location,distributionno,crateno,cratecount,distyear) " _
            & "VALUEs('" & mMaxId & "','" & Format(txtentrydate.Value, "yyyy-MM-dd") & "','" & Trim(ItemGrd.TextMatrix(i, 1)) & "', " _
            & "'" & mPlantVariety & "','" & Trim(ItemGrd.TextMatrix(i, 3)) & "'," _
            & "'2','4', " _
            & "'" & cbostaff.BoundText & "','0','" & Val(ItemGrd.TextMatrix(i, 6)) & "','ON','" & txtloc.Text & "','" & cbotrnid.BoundText & "','" & Trim(ItemGrd.TextMatrix(i, 5)) & "','" & Trim(ItemGrd.TextMatrix(i, 4)) & "','" & Val(txtyr.Text) & "')"
            

            
  mMaxId = mMaxId + 1
 Next
 
 
MHVDB.Execute "update tblplantdistributiondetail set senttofield='Y' where distno='" & cbotrnid.BoundText & "' and year='" & Val(txtyr.Text) & "' "
MHVDB.Execute "UPDATE mhv.tblregistrationsettings SET updatetable ='Yes' WHERE tblid ='4'"
MHVDB.CommitTrans


Exit Sub
err:
MHVDB.RollbackTrans
MsgBox err.Description
End Sub
Private Sub populatedno(Operation As String)

Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
                   
db.Open CnnString

Set rs = Nothing
If Operation = "ADD" Then
'If rs.State = adStateOpen Then rs.Close
'rs.Open "select distinct distno  from tblplantdistributiondetail where  subtotindicator='' and status not in ('C','F') and distno>0 and trnid in (select trnid from tblplantdistributionheader where status='ON')order by distno ", db
'Set cbotrnid.RowSource = rs
'cbotrnid.ListField = "distno"
'cbotrnid.BoundColumn = "distno"

If rs.State = adStateOpen Then rs.Close
rs.Open "select distinct distno,concat(cast(distno as char) , '  ', cast(year as char)) dist  from tblplantdistributiondetail where  subtotindicator='' and status not in ('C','F') and distno>0 and trnid in (select trnid from tblplantdistributionheader where status='ON'and planneddist='Y')order by distno desc  ", db
Set cbotrnid.RowSource = rs
cbotrnid.ListField = "dist"
cbotrnid.BoundColumn = "distno"



ElseIf Operation = "OPEN" Then
If rs.State = adStateOpen Then rs.Close
rs.Open "select distinct distno,concat(cast(distno as char) , '  ', cast(year as char)) dist  from tblplantdistributiondetail  where status<>'C' and distno  in(select distributionno from tblqmssendtofieldhdr where status='ON') order by distno desc", db

Set cbotrnid.RowSource = rs
cbotrnid.ListField = "dist"
cbotrnid.BoundColumn = "distno"
Else
MsgBox "Wrong Operation Selected."
End If
End Sub
Private Sub CLEARCONTROLL()
ItemGrd.Clear
ItemGrd.FormatString = "       |^ Batch No.|^Variety |^Facility |^No. Of Crates|^ Crate #                 |Qty.      |"
txtentrydate.Value = Format(Now, "dd/MM/yyyy")
cbotrnid.Text = ""
txtsendtofield.Text = ""
cbostaff.Text = ""
txtdsheetqty.Text = ""
txtshortexcees.Text = ""
txtselected.Text = ""
txtyr.Text = ""
End Sub

Private Sub txtfind_Change()
'onchange
End Sub
Private Sub onchange()
Dim SQLSTR As String
Dim i As Integer
'If Len(TXTSEARCHID.Text) <= 3 Then Exit Sub
Dim rs As New ADODB.Recordset
If txtfind.Text = "'" Then
MsgBox (err.Number & " : " & "Enter Valid Character for Search.")
txtfind.Text = ""
txtfind.SetFocus
Exit Sub
End If




    If txtfind.Text = "" Then
'        cleargrid
'        i = 1
  Exit Sub
  
  End If
        Set rs = Nothing
                
        
        SQLSTR = "select * from tblqmscrate where crateno like '" & txtfind.Text & "%' order by cast(crateno as unsigned)"
      
        
        
        'SQLSTR = "" '
        
        
        
        
        rs.Open SQLSTR, MHVDB
        If rs.RecordCount > 0 Then
        rs.MoveFirst
        Else
        On Error Resume Next
        End If
'         cleargrid
'         i = 1
        
        DZLIST.Clear

With rs
Do While Not .EOF
   DZLIST.AddItem Trim(!crateno)
   .MoveNext
Loop
End With
'        If ListView1.ListItems.Count <> 0 Then
'        ListView1.ListItems(1).Selected = True
'        End If
End Sub

Private Sub txtfind_DblClick()
Dim i As Integer
For i = 0 To DZLIST.ListCount - 1
If txtfind.Text = DZLIST.List(i) Then
DZLIST.Selected(i) = True
End If
Next
txtfind.Text = ""
cratecnt
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

Dim i As Integer
For i = 0 To DZLIST.ListCount - 1
If txtfind.Text = DZLIST.List(i) Then
DZLIST.Selected(i) = True

End If
Next
txtfind.Text = ""
cratecnt
End If
cratecnt
End Sub

Private Sub txtnoofcrates_Validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not IsNumeric(txtnoofcrates) Then
   Beep
   MsgBox "Enter a valid No."
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   ItemGrd.TextMatrix(CurrRow, 4) = Val(txtnoofcrates.Text)
   ItemGrd.TextMatrix(CurrRow, 6) = Val(txtnoofcrates.Text) * 35
   ValidRow = True
   
End If
End If
txtnoofcrates.Visible = False
getsum
End Sub
Private Sub getsum()
Dim i As Integer
totsendtofield = 0
totcrate = 0
For i = 1 To ItemGrd.rows - 1
If Len(ItemGrd.TextMatrix(i, 1)) = 0 Then Exit For
totsendtofield = totsendtofield + Val(ItemGrd.TextMatrix(i, 6))
totcrate = totcrate + Val(ItemGrd.TextMatrix(i, 4))
Next

txtsendtofield.Text = totsendtofield
txttocrate.Text = totcrate
txtshortexcees.Text = Val(txtdsheetqty.Text) - Val(txtsendtofield.Text)
If Val(txtshortexcees.Text) < 0 Then
'txtshortexcees.BackColor = vbRed
Else
'txtshortexcees.BackColor = vbWhite
End If


End Sub

Private Sub txtQty_validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not IsNumeric(txtqty) Then
   Beep
   MsgBox "Enter a valid No."
   ValidRow = False
   Cancel = True
   Exit Sub
Else
  
   ItemGrd.TextMatrix(CurrRow, 6) = Val(txtqty.Text)
   ValidRow = True
   
End If
End If
txtqty.Visible = False
getsum
End Sub
