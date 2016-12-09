VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmupdatedeliverystatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UPDATE DELIVERY STATUS"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11040
   Icon            =   "frmupdatedeliverystatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   10200
      TabIndex        =   36
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtreturnedqty 
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox txtassignedqty 
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox txtbacktonursary 
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "FARMER INFORMATION"
      Height          =   1215
      Left            =   6840
      TabIndex        =   12
      Top             =   2280
      Width           =   4095
      Begin VB.TextBox txttrees 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtfarmername 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtarea 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "No Of Trees"
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
         Left            =   1920
         TabIndex        =   21
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Area"
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
         TabIndex        =   20
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Farmer Description"
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
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "SCHEDULE INFORMATION"
      Height          =   1215
      Left            =   6840
      TabIndex        =   11
      Top             =   960
      Width           =   4095
      Begin VB.TextBox txtmonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtyear 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtscheduledesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Month"
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
         Left            =   2160
         TabIndex        =   15
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Year"
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
         TabIndex        =   14
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Schedule Description"
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
         TabIndex        =   13
         Top             =   240
         Width           =   1830
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6615
      Begin VB.CheckBox chkbacktonursary 
         Caption         =   "Back To Nursery?"
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
         Height          =   255
         Left            =   3960
         TabIndex        =   25
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox chkassignednext 
         Caption         =   "Assigned To Next farmer?"
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
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   1680
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "frmupdatedeliverystatus.frx":0E42
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
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
      Begin MSDataListLib.DataCombo cbodeliveryno 
         Bindings        =   "frmupdatedeliverystatus.frx":0E57
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
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
      Begin MSDataListLib.DataCombo cbofarmerid 
         Bindings        =   "frmupdatedeliverystatus.frx":0E6C
         Height          =   360
         Left            =   1320
         TabIndex        =   9
         Top             =   1200
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   635
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Farmer Id"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Delivery No."
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
         TabIndex        =   7
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Schedule No."
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
         TabIndex        =   5
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CANCELLATION TYPE"
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6615
      Begin VB.OptionButton optunplanned 
         Caption         =   "Unplanned Distribution"
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
         Left            =   3840
         TabIndex        =   34
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optdelno 
         Caption         =   "By Delivery No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optfarmerid 
         Caption         =   "By Farmer No."
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
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   1935
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
            Picture         =   "frmupdatedeliverystatus.frx":0E81
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmupdatedeliverystatus.frx":121B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmupdatedeliverystatus.frx":15B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmupdatedeliverystatus.frx":228F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmupdatedeliverystatus.frx":26E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmupdatedeliverystatus.frx":2E9B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
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
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   2775
      Left            =   120
      TabIndex        =   35
      Top             =   3960
      Width           =   9735
      _cx             =   17171
      _cy             =   4895
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   0   'False
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
      Rows            =   20
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmupdatedeliverystatus.frx":3235
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
   Begin VB.Label Label12 
      BackColor       =   &H000000FF&
      Caption         =   "You cannot proceed to save this record."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   32
      Top             =   7560
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "DIfference"
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
      Left            =   4200
      TabIndex        =   31
      Top             =   7560
      Width           =   915
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Field Drop Outs/Cencellation"
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
      TabIndex        =   29
      Top             =   7080
      Width           =   2475
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Back To Nersury"
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
      Left            =   7680
      TabIndex        =   27
      Top             =   7080
      Width           =   1440
   End
End
Attribute VB_Name = "frmupdatedeliverystatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
    

Dim assignedqty As Long
Dim ValidRow As Boolean
Dim CurrRow As Long
Dim fldtrnid, fldyear, fldmnth, fldsno, flddistno, fldtotalplant, fldcrateno, fldbcrate, fldecrate, oldSno As Long
Dim fldbno, fldplno, fldcrate, fldserialmatch As Long
Dim fldfarmercode, fldschedule, fldsubtotindicator, fldnewold, fldstatus As String
Dim fldarea, fldssp, fldmop, fldurea, flddolomite, fldtotalkg1, fldamountnu1, fldkg, fldamountnu2, fldtotalamount As Double


Private Sub FillGridCombo()
        Dim RstTemp As New ADODB.Recordset
        Dim NmcItemData As String
        Dim StrComboList As String
        Dim mn1 As String
        mn1 = "          |"
        StrComboList = ""
        
            Set RstTemp = Nothing
            RstTemp.Open ("select idfarmer,farmername from tblfarmer where status='A' ORDER BY idfarmer"), MHVDB

            If Not RstTemp.BOF Then
               Do While Not RstTemp.EOF
                    NmcItemData = ""
                    StrComboList = IIf(StrComboList = "", RstTemp("idfarmer").Value & " " & RstTemp("farmername").Value, StrComboList & "|" & RstTemp("idfarmer").Value) & " " & RstTemp("farmername").Value

                    RstTemp.MoveNext
               Loop
            End If
            RstTemp.Close
       mygrid.ComboList = mn1 + StrComboList



    End Sub

Private Sub cbodeliveryno_GotFocus()
If optfarmerid.Value = True Then
cbofarmerid.Enabled = True
Else
cbofarmerid.Enabled = False
End If
End Sub

Private Sub cbodeliveryno_LostFocus()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Dim RSTR As New ADODB.Recordset

cbodeliveryno.Enabled = False


Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
RSTR.Open "select concat(idfarmer,'  ',farmername) as farmername ,farmercode as idfarmer from tblfarmer a,tblplantdistributiondetail b where idfarmer=farmercode and trnid='" & cbotrnid.BoundText & "' and year='" & txtyear.Text & "'  and mnth='" & txtmonth.Text & "' and subtotindicator not in('T','S') and b.status<>'C' and distno='" & cbodeliveryno.BoundText & "'", db
Set cbofarmerid.RowSource = RSTR
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"

If optdelno.Value = True Then
Set RSTR = Nothing
RSTR.Open "select count(farmercode) as cnt,sum(area) as area, sum(totalplant) as totaltrees from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and status<>'C'", MHVDB
If RSTR.EOF <> True Then
txtfarmername.Text = "No. of farmers: " & RSTR!cnt
TXTAREA.Text = Format(RSTR!Area, "####0.00")
txttrees.Text = IIf(IsNull(RSTR!totaltrees), "", RSTR!totaltrees)

End If


End If

Set RSTR = Nothing
RSTR.Open "select sum(debit) debit from tblqmsplanttransaction where transactiontype='5' and distributionno='" & cbodeliveryno.BoundText & "' ", MHVDB
If RSTR.EOF <> True Then
txtbacktonursary.Text = IIf(IsNull(RSTR!debit), "", RSTR!debit)
End If

If optunplanned.Value = False Then
fillgrid
End If




chkassignednext.Enabled = True
chkbacktonursary.Enabled = True
getAssignedQty
End Sub
Private Sub fillgrid()
On Error GoTo err
Dim rs As New ADODB.Recordset


mygrid.Clear
mygrid.Rows = 20
mygrid.FormatString = "^SL.NO.|^FARMER ID|^AREA|^NO. OF TREES|B CRT.|E CRT.|N CRT.|"
mygrid.ColWidth(0) = 750
mygrid.ColWidth(1) = 2835
mygrid.ColWidth(2) = 930
mygrid.ColWidth(3) = 1455
mygrid.ColWidth(4) = 960
mygrid.ColWidth(5) = 960
mygrid.ColWidth(6) = 960
mygrid.ColWidth(7) = 300
Set rs = Nothing
rs.Open "select * from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and assignedatfield='Y' and status<>'C'", MHVDB

i = 1
If rs.EOF <> True Then
Do While rs.EOF <> True
mygrid.Rows = mygrid.Rows + 1
mygrid.TextMatrix(i, 0) = i
FindFA rs!farmercode, "F"
mygrid.TextMatrix(i, 1) = rs!farmercode & "  " & FAName
mygrid.TextMatrix(i, 2) = rs!Area
mygrid.TextMatrix(i, 3) = rs!totalplant
mygrid.TextMatrix(i, 4) = rs!bcrate
mygrid.TextMatrix(i, 5) = rs!ecrate
mygrid.TextMatrix(i, 6) = rs!crate
mygrid.TextMatrix(i, 7) = rs!isManual
rs.MoveNext
i = i + 1
Loop
chkassignednext.Value = 1

Else

End If

Set rs = Nothing
If Len(Trim(cbofarmerid.Text)) = 0 Then
rs.Open "select count(farmercode) as cnt,sum(area) area,sum(totalplant) totalplant from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and status<>'C'", MHVDB
Else
rs.Open "select count(farmercode) as cnt,sum(area) area,sum(totalplant) totalplant from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and farmercode='" & cbofarmerid.BoundText & "' and status<>'C'", MHVDB
End If
If rs.EOF <> True Then
If Len(Trim(cbofarmerid.Text)) = 0 Then
txtfarmername.Text = "No. of farmers in distribution no. :  " & cbodeliveryno.BoundText & " is " & rs!cnt
Else
FindFA cbofarmerid.BoundText, "F"
txtfarmername.Text = cbofarmerid.BoundText & "  " & FAName
End If
TXTAREA.Text = Format(rs!Area, "####0.00")
txttrees.Text = IIf(IsNull(rs!totalplant), "", rs!totalplant)

End If




rs.Close
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub cbofarmerid_LostFocus()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Dim RSTR As New ADODB.Recordset
cbofarmerid.Enabled = False

Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
RSTR.Open "select *  from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and year='" & txtyear.Text & "'  and mnth='" & txtmonth.Text & "' and subtotindicator not in('T','S') and status<>'C' and distno='" & cbodeliveryno.BoundText & "' and farmercode='" & cbofarmerid.BoundText & "'", db
If RSTR.EOF <> True Then
FindFA RSTR!farmercode, "F"
txtfarmername.Text = RSTR!farmercode & " " & FAName
txttrees.Text = RSTR!totalplant
TXTAREA.Text = Format(RSTR!Area, "####0.00")
mygrid.Clear
mygrid.Rows = 20
mygrid.FormatString = "^SL.NO.|^FARMER ID|^AREA|^NO. OF TREES|B CRT.|E CRT.|N CRT.|"
mygrid.ColWidth(0) = 750
mygrid.ColWidth(1) = 2835
mygrid.ColWidth(2) = 930
mygrid.ColWidth(3) = 1455
mygrid.ColWidth(4) = 960
mygrid.ColWidth(5) = 960
mygrid.ColWidth(6) = 960
mygrid.ColWidth(7) = 300
End If






End Sub

Private Sub cbotrnid_GotFocus()

cbodeliveryno.Enabled = True
Frame1.Enabled = False
End Sub

Private Sub cbotrnid_LostFocus()
cbotrnid.Enabled = False

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Dim RSTR As New ADODB.Recordset

Set RSTR = Nothing
RSTR.Open "select * from tblplantdistributionheader where trnid='" & cbotrnid.BoundText & "'", db
If RSTR.EOF <> True Then
txtscheduledesc.Text = RSTR!distributionname
txtyear.Text = RSTR!Year
txtmonth.Text = RSTR!mnth
End If


Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close

RSTR.Open "select distinct distno from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and year='" & txtyear.Text & "'  and mnth='" & txtmonth.Text & "' and subtotindicator not in('T','S') and status<>'C' order by distno desc", db

Set cbodeliveryno.RowSource = RSTR
cbodeliveryno.ListField = "distno"
cbodeliveryno.BoundColumn = "distno"
End Sub

Private Sub DataCombo1_Click(Area As Integer)



End Sub

Private Sub chkassignednext_Click()
If Len(cbotrnid.Text) = 0 Then Exit Sub
If chkassignednext.Value = 1 Then
mygrid.Enabled = True
chkbacktonursary.Value = 0
If optunplanned.Value = False Then
fillgrid
End If
getAssignedQty
Else
mygrid.Enabled = False

End If
End Sub

Private Sub chkbacktonursary_Click()

If chkbacktonursary.Value = 1 Then
chkassignednext.Value = 0
mygrid.Clear
mygrid.Rows = 20
mygrid.FormatString = "^SL.NO.|^FARMER ID|^AREA|^NO. OF TREES|B|E|N|"
mygrid.ColWidth(0) = 750
mygrid.ColWidth(1) = 2835
mygrid.ColWidth(2) = 930
mygrid.ColWidth(3) = 1455
mygrid.ColWidth(4) = 960
mygrid.ColWidth(5) = 960
mygrid.ColWidth(6) = 960
mygrid.ColWidth(7) = 300
mygrid.Enabled = False
getAssignedQty
Else
mygrid.Enabled = True
End If

End Sub

Private Sub Command1_Click()
dijkstra_shortest_Path
End Sub

Private Sub Form_Load()
Dim RSTR As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
RSTR.Open "select concat(cast(trnid as char) ,' ',distributionname,' ',cast(year as char),' ',cast(mnth as char)) as dname,trnid  from tblplantdistributionheader where status='ON' order by trnid desc", db
Set cbotrnid.RowSource = RSTR
cbotrnid.ListField = "dname"
cbotrnid.BoundColumn = "trnid"



'Mygrid.col = 4

End Sub

Private Sub mygrid_Click()

If chkassignednext.Value = 0 Then
mygrid.Enabled = False
Exit Sub
End If

mygrid.Editable = flexEDNone
If mygrid.col = 1 And Len(mygrid.TextMatrix(mygrid.row - 1, 3)) > 0 Then
mygrid.Editable = flexEDKbdMouse
FillGridCombo
Else
mygrid.ComboList = ""
mygrid.Editable = flexEDNone
End If
        
If (mygrid.col = 4 Or mygrid.col = 5 Or mygrid.col = 6) And (Val(mygrid.TextMatrix(mygrid.row, 2)) > 0 Or Len(mygrid.TextMatrix(mygrid.row - 1, 3))) Then
mygrid.ComboList = ""
mygrid.Editable = flexEDKbdMouse
Else
'mygrid.Editable = flexEDNone
End If
                                                       
If mygrid.col = 2 Or mygrid.col = 3 Then
mygrid.ComboList = ""
mygrid.Editable = flexEDNone
End If

If chkassignednext.Value = 1 Then
chkassignednext.Enabled = False
chkbacktonursary.Enabled = False
End If
getAssignedQty
End Sub
Private Sub getAssignedQty()
Dim diff As Long
Dim rsF As New ADODB.Recordset
assignedqty = 0
Dim i As Integer
Set rsF = Nothing
rsF.Open "select * from tbldistformula where fid='1'", MHVDB
For i = 1 To mygrid.Rows - 1
If Len(Trim(mygrid.TextMatrix(i, 1))) = 0 Then Exit For
assignedqty = assignedqty + Val(mygrid.TextMatrix(i, 3))
Next
txtassignedqty.Text = IIf(assignedqty = 0, "", assignedqty)
diff = Val(txttrees.Text) - Val(txtassignedqty.Text)

txtreturnedqty.Text = Abs(diff)
If diff >= 0 Then
txtreturnedqty.BackColor = &HC0FFFF
Label12.Visible = False
Else
'MsgBox "You Cannot save this information."
txtreturnedqty.BackColor = vbRed
Label12.Visible = True
End If



End Sub

Private Sub mygrid_DblClick()
Dim myinput As String
If mygrid.col = 3 And Len(mygrid.TextMatrix(mygrid.row, 1)) > 0 Then
If MsgBox("Do You want to enter the no. of plants manually?", vbYesNo) = vbYes Then
mygrid.TextMatrix(mygrid.row, 7) = "M"
myinput = InputBox("Enter The No. Of Plants.", "MH Distribution", mygrid.TextMatrix(mygrid.row, 3))
            If Not IsNumeric(myinput) Then
            'MsgBox "Invalid number,Double Click again to enable the input box."
            Else
            
            mygrid.TextMatrix(mygrid.row, 3) = CLng(myinput)
            End If
            Else
        mygrid.TextMatrix(mygrid.row, 4) = ""
End If

End If
End Sub

Private Sub mygrid_KeyPress(KeyAscii As Integer)
getAssignedQty
End Sub

Private Sub Mygrid_LeaveCell()
Dim rs As New ADODB.Recordset
Dim rsF As New ADODB.Recordset


If mygrid.col = 1 And Len(mygrid.TextMatrix(mygrid.row, 1)) > 0 And Trim(mygrid.TextMatrix(mygrid.row, 7)) <> "M" Then
Set rs = Nothing
rs.Open "select sum(regland) as regland from tbllandreg where farmerid='" & Mid(Trim(mygrid.TextMatrix(mygrid.row, 1)), 1, 14) & "' and plantedstatus='N'", MHVDB
If rs.EOF <> True Then
mygrid.TextMatrix(mygrid.row, 2) = Format(IIf(IsNull(rs!regland), "", rs!regland), "####0.00")
Set rsF = Nothing
rsF.Open "select * from tbldistformula where fid='1'", MHVDB
mygrid.TextMatrix(mygrid.row, 3) = Round(((Val(mygrid.TextMatrix(mygrid.row, 2)) * rsF!totalplant)), 0)
End If

End If
TB.Buttons(3).Enabled = True
End Sub

Private Sub Mygrid_Validate(Cancel As Boolean)
getAssignedQty
End Sub

Private Sub Mygrid_ValidateEdit(ByVal row As Long, ByVal col As Long, Cancel As Boolean)
getAssignedQty
End Sub

Private Sub optdelno_Click()
If optdelno.Value = True Then
chkassignednext.Enabled = False
Else
chkassignednext.Enabled = True
End If
End Sub

Private Sub optfarmerid_Click()
If optfarmerid.Value = True Then
chkassignednext.Enabled = True
Else
chkassignednext.Enabled = False
End If
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "ADD"
    
       Case "OPEN"
         
       clearcontrol
       cbotrnid.Enabled = True
       Frame1.Enabled = True
       TB.Buttons(3).Enabled = True
       Case "SAVE"
       MNU_SAVE
       
       Case "EXIT"
       Unload Me
End Select
End Sub
Private Sub MNU_SAVE()
On Error GoTo err
Dim i, j As Integer
Dim opt As Integer
Dim rs As New ADODB.Recordset
If optfarmerid.Value = True And Len(cbofarmerid.Text) = 0 Then
MsgBox "Farmer to cancel must be selected."
Exit Sub
End If
'If optfarmerid.Value = True Then
'        Set rs = Nothing
'        rs.Open "select * from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and farmercode<>'" & cbofarmerid.BoundText & "' ", MHVDB
'        If rs.EOF <> True Then
'                Do While rs.EOF <> True
'                        For i = 1 To Mygrid.Rows - 1
'                                If Trim(Mid(Mygrid.TextMatrix(i, 1), 1, 14)) = rs!farmercode Then
'                                        MsgBox "The farmer " & rs!farmercode & " already exist in the distribution no. " & rs!distno
'                                        Exit Sub
'                                End If
'                        Next
'                rs.MoveNext
'                Loop
'        End If
'End If

initvar

'opt 0 = delno, 1 = farmer

MHVDB.BeginTrans


If chkbacktonursary.Value = 1 Then

            If optdelno.Value = True And chkbacktonursary.Value = 1 Then
                MHVDB.Execute "update tblplantdistributiondetail set status='C',assignedatfield='Y',operation='Y' where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "'"
                MHVDB.Execute "update tblplantdistributiondetail set status='C',assignedatfield='Y',operation='Y' where trnid='" & cbotrnid.BoundText & "' and serialmatch='" & cbodeliveryno.BoundText & "'"
            ElseIf optfarmerid.Value = True And chkbacktonursary.Value = 1 Then
                MHVDB.Execute "update tblplantdistributiondetail set status='C',assignedatfield='Y',operation='Y' where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and farmercode='" & cbofarmerid.BoundText & "'"
            Else
                MHVDB.RollbackTrans
                MsgBox "Invalid selection of cancellation option."
            Exit Sub
            End If
ElseIf chkassignednext.Value = 1 Then
If optdelno.Value = True Then
           getval 0
            ElseIf optfarmerid.Value = True Then
           getval 1
            ElseIf optunplanned.Value = True Then
             getval 2
            Else
                MHVDB.RollbackTrans
                MsgBox "Invalid selection of cancellation option."
            Exit Sub
            End If


Else

If optfarmerid.Value = True And chkassignednext.Value = 0 Then
                MHVDB.Execute "update tblplantdistributiondetail set status='C',assignedatfield='Y',operation='Y' where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and farmercode='" & cbofarmerid.BoundText & "'"
            Else
            MsgBox "Invalid Selection of Cancellation Option."
            End If

End If
TB.Buttons(3).Enabled = False
MHVDB.CommitTrans
Exit Sub
err:
    MHVDB.RollbackTrans
    MsgBox err.Description

End Sub
Private Sub getval(opt As Integer)
Dim isManual As String
Dim isError As Boolean
Dim assignedcount As Integer
Dim rs As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim SQLSTR As String
Dim i As Integer
isError = False
assignedcount = 0
isManual = ""
Set rsF = Nothing
rsF.Open "select * from tbldistformula where fid='1'", MHVDB
' get the total no of new farmers involved
For i = 1 To mygrid.Rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
assignedcount = assignedcount + 1
Next
'-----
If assignedcount = 0 Then
isError = True
MsgBox "Must Assigned plants to some farmers."
Exit Sub
End If
SQLSTR = ""
Set rs = Nothing
If opt = 0 Then
            rs.Open "select  *,max(sno) as mno from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' order by sno", MHVDB
            If rs.EOF <> True Then
            oldSno = rs!mno + 1
            fldsno = rs!mno
            flddistno = rs!distno
            fldtrnid = rs!trnid
            fldyear = rs!Year
            fldmnth = rs!mnth
            End If
            Set rs = Nothing
            MHVDB.Execute "update tblplantdistributiondetail set status='C',assignedatfield='Y',operation='Y' where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "'"
            MHVDB.Execute "update tblplantdistributiondetail set sno=sno +  " & assignedcount & "  where sno > '" & fldsno & "' and trnid='" & cbotrnid.BoundText & "'"
            
            MHVDB.Execute "delete from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and assignedatfield='Y' and status<>'C'"
            
            Set rs = Nothing
            For i = 1 To mygrid.Rows - 1
            If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
            
            fldarea = Format(Val(mygrid.TextMatrix(i, 2)), "###0.00") '9
            If Trim(mygrid.TextMatrix(i, 7)) = "M" Then
            fldtotalplant = Val(mygrid.TextMatrix(i, 3))
            Else
            fldtotalplant = Round(((fldarea * rsF!totalplant)), 0)       '10
            End If
            isManual = Trim(mygrid.TextMatrix(i, 7))
            
            fldcrateno = Val(mygrid.TextMatrix(i, 4)) + Val(mygrid.TextMatrix(i, 5)) + Val(mygrid.TextMatrix(i, 6)) 'Round((fldtotalplant) / rsF!crateno, 0)             '11
            fldbcrate = Val(mygrid.TextMatrix(i, 4)) 'Round((fldcrateno * rsF!crateno * rsF!bcrate) / rsF!crateno, 0) '12
            fldecrate = Val(mygrid.TextMatrix(i, 5)) 'Round((fldcrateno * rsF!crateno - rsF!crateno - Val(fldbcrate * rsF!crateno)) / rsF!crateno, 0) '13
            fldbno = 0 'rsF!crateno / 35 '14
            fldplno = 0 '15
            fldcrate = Val(mygrid.TextMatrix(i, 6)) '0 '16
            
            fldserialmatch = oldSno
            fldssp = Round((fldtotalplant * rsF!ssp), 2) '17
            fldmop = Round((fldtotalplant * rsF!mop), 2) '18
            fldurea = Round((fldtotalplant * rsF!urea), 2) '19
            flddolomite = Round((fldtotalplant * rsF!dolomite), 2) '20
            fldtotalkg1 = Round(fldssp + fldmop + fldurea + flddolomite, 0) '21
            fldamountnu1 = Round(Val(fldssp * rsF!sspperkg) + Val(fldmop * rsF!mopperkg) + Val(fldurea * rsF!ureaperkg) + Val(flddolomite * rsF!dolomiteperkg), 0) '22
            fldkg = Round((fldtotalplant * rsF!kg), 0) '23
            fldamountnu2 = Round((fldkg * rsF!amountnu), 0) '24
            fldtotalamount = fldamountnu1 + fldamountnu2 '25
            fldfarmercode = Mid(mygrid.TextMatrix(i, 1), 1, 14)
            fldschedule = ""
            fldsubtotindicator = ""
            fldnewold = ""
            fldstatus = ""
            
            SQLSTR = "insert into tblplantdistributiondetail(trnid,year,mnth,sno,distno," _
                         & "farmercode,area,totalplant,crateno,bcrate,ecrate,bno,plno,crate,ssp," _
                         & "mop,urea,dolomite,totalkg1,amountnu1,kg,amountnu2,totalamount," _
                         & "schedule,serialmatch,subtotindicator,newold,assignedAtField,ismanual,operation)values(" _
                         & "'" & fldtrnid & "','" & fldyear & "','" & fldmnth & "','" & oldSno & "'," _
                         & "'" & flddistno & "','" & fldfarmercode & "','" & fldarea & "','" & fldtotalplant & "'," _
                         & "'" & fldcrateno & "','" & fldbcrate & "','" & fldecrate & "','" & fldbno & "','" & fldplno & "'," _
                         & "'" & fldcrate & "','" & fldssp & "','" & fldmop & "','" & fldurea & "','" & flddolomite & "'," _
                         & "'" & fldtotalkg1 & "','" & fldamountnu1 & "','" & fldkg & "','" & fldamountnu2 & "','" & fldtotalamount & "'," _
                         & "'" & fldschedule & "','" & flddistno & "','" & fldsubtotindicator & "','" & fldnewold & "','Y','" & isManual & "','Y')"
                         MHVDB.Execute SQLSTR
                        oldSno = oldSno + 1
            Next


ElseIf opt = 1 Then
            rs.Open "select  *,max(sno) as mno from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and farmercode='" & cbofarmerid.BoundText & "' order by sno", MHVDB
            If rs.EOF <> True Then
            oldSno = rs!mno + 1
            fldsno = rs!mno
            flddistno = rs!distno
            fldtrnid = rs!trnid
            fldyear = rs!Year
            fldmnth = rs!mnth
            End If
            Set rs = Nothing
            MHVDB.Execute "update tblplantdistributiondetail set status='C',assignedatfield='Y',operation='Y' where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and farmercode='" & cbofarmerid.BoundText & "'"
            MHVDB.Execute "update tblplantdistributiondetail set sno=sno +  " & assignedcount & "  where sno > '" & fldsno & "' and trnid='" & cbotrnid.BoundText & "'"
            Set rs = Nothing
            
            
            
            For i = 1 To mygrid.Rows - 1
            If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
            MHVDB.Execute "delete from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and assignedatfield='Y' and status<>'C' and farmercode='" & Mid(Trim(mygrid.TextMatrix(i, 1)), 1, 14) & "'"
            fldarea = Format(Val(mygrid.TextMatrix(i, 2)), "###0.00") '9
            If Trim(mygrid.TextMatrix(i, 7)) = "M" Then
            fldtotalplant = Val(mygrid.TextMatrix(i, 3))
            Else
            fldtotalplant = Round(((fldarea * rsF!totalplant)), 0)       '10
            End If
            isManual = Trim(mygrid.TextMatrix(i, 7))
            fldcrateno = Val(mygrid.TextMatrix(i, 4)) + Val(mygrid.TextMatrix(i, 5)) + Val(mygrid.TextMatrix(i, 6)) 'Round((fldtotalplant) / rsF!crateno, 0)             '11
            fldbcrate = Val(mygrid.TextMatrix(i, 4)) 'Round((fldcrateno * rsF!crateno * rsF!bcrate) / rsF!crateno, 0) '12
            fldecrate = Val(mygrid.TextMatrix(i, 5)) 'Round(fldcrateno * rsF!crateno - rsF!crateno - Val(fldbcrate * rsF!crateno), 0) '13
            fldbno = 0 'rsF!crateno / 35 '14
            fldplno = 0 '15
            fldcrate = Val(mygrid.TextMatrix(i, 6)) '0 '16
            
            fldserialmatch = oldSno
            fldssp = Round((fldtotalplant * rsF!ssp), 2) '17
            fldmop = Round((fldtotalplant * rsF!mop), 2) '18
            fldurea = Round((fldtotalplant * rsF!urea), 2) '19
            flddolomite = Round((fldtotalplant * rsF!dolomite), 2) '20
            fldtotalkg1 = Round(fldssp + fldmop + fldurea + flddolomite, 0) '21
            fldamountnu1 = Round(Val(fldssp * rsF!sspperkg) + Val(fldmop * rsF!mopperkg) + Val(fldurea * rsF!ureaperkg) + Val(flddolomite * rsF!dolomiteperkg), 0) '22
            fldkg = Round((fldtotalplant * rsF!kg), 0) '23
            fldamountnu2 = Round((fldkg * rsF!amountnu), 0) '24
            fldtotalamount = fldamountnu1 + fldamountnu2 '25
            fldfarmercode = Mid(mygrid.TextMatrix(i, 1), 1, 14)
            fldschedule = ""
            fldsubtotindicator = ""
            fldnewold = ""
            fldstatus = ""
            
            
            
            SQLSTR = "insert into tblplantdistributiondetail(trnid,year,mnth,sno,distno," _
                         & "farmercode,area,totalplant,crateno,bcrate,ecrate,bno,plno,crate,ssp," _
                         & "mop,urea,dolomite,totalkg1,amountnu1,kg,amountnu2,totalamount," _
                         & "schedule,serialmatch,subtotindicator,newold,assignedAtField,ismanual,operation)values(" _
                         & "'" & fldtrnid & "','" & fldyear & "','" & fldmnth & "','" & oldSno & "'," _
                         & "'" & flddistno & "','" & fldfarmercode & "','" & fldarea & "','" & fldtotalplant & "'," _
                         & "'" & fldcrateno & "','" & fldbcrate & "','" & fldecrate & "','" & fldbno & "','" & fldplno & "'," _
                         & "'" & fldcrate & "','" & fldssp & "','" & fldmop & "','" & fldurea & "','" & flddolomite & "'," _
                         & "'" & fldtotalkg1 & "','" & fldamountnu1 & "','" & fldkg & "','" & fldamountnu2 & "','" & fldtotalamount & "'," _
                         & "'" & fldschedule & "','" & flddistno & "','" & fldsubtotindicator & "','" & fldnewold & "','Y','" & isManual & "','Y')"
                         MHVDB.Execute SQLSTR
                        oldSno = oldSno + 1
            Next

Else
rs.Open "select  *,max(sno) as mno from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "'  order by sno", MHVDB
            If rs.EOF <> True Then
            oldSno = rs!mno + 1
            fldsno = rs!mno
            flddistno = rs!distno
            fldtrnid = rs!trnid
            fldyear = rs!Year
            fldmnth = rs!mnth
            End If
            Set rs = Nothing
           ' MHVDB.Execute "update tblplantdistributiondetail set status='C',assignedatfield='Y',operation='Y' where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and farmercode='" & cbofarmerid.BoundText & "'"
            MHVDB.Execute "update tblplantdistributiondetail set sno=sno +  " & assignedcount & "  where sno > '" & fldsno & "' and trnid='" & cbotrnid.BoundText & "'"
            Set rs = Nothing
            
            
            
            For i = 1 To mygrid.Rows - 1
            If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
           ' MHVDB.Execute "delete from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and distno='" & cbodeliveryno.BoundText & "' and assignedatfield='Y' and status<>'C' and farmercode='" & Mid(Trim(Mygrid.TextMatrix(i, 1)), 1, 14) & "'"
            fldarea = Format(Val(mygrid.TextMatrix(i, 2)), "###0.00") '9
            If Trim(mygrid.TextMatrix(i, 7)) = "M" Then
            fldtotalplant = Val(mygrid.TextMatrix(i, 3))
            Else
            fldtotalplant = Round(((fldarea * rsF!totalplant)), 0)       '10
            End If
            isManual = Trim(mygrid.TextMatrix(i, 7))
            fldcrateno = Val(mygrid.TextMatrix(i, 4)) + Val(mygrid.TextMatrix(i, 5)) + Val(mygrid.TextMatrix(i, 6)) 'Round((fldtotalplant) / rsF!crateno, 0)             '11
            fldbcrate = Val(mygrid.TextMatrix(i, 4)) 'Round((fldcrateno * rsF!crateno * rsF!bcrate) / rsF!crateno, 0) '12
            fldecrate = Val(mygrid.TextMatrix(i, 5)) 'Round(fldcrateno * rsF!crateno - rsF!crateno - Val(fldbcrate * rsF!crateno), 0) '13
            fldbno = 0 ' rsF!crateno / 35 '14
            fldplno = 0 '15
            fldcrate = Val(mygrid.TextMatrix(i, 6)) ' 0 '16
            
            fldserialmatch = oldSno
            fldssp = Round((fldtotalplant * rsF!ssp), 2) '17
            fldmop = Round((fldtotalplant * rsF!mop), 2) '18
            fldurea = Round((fldtotalplant * rsF!urea), 2) '19
            flddolomite = Round((fldtotalplant * rsF!dolomite), 2) '20
            fldtotalkg1 = Round(fldssp + fldmop + fldurea + flddolomite, 0) '21
            fldamountnu1 = Round(Val(fldssp * rsF!sspperkg) + Val(fldmop * rsF!mopperkg) + Val(fldurea * rsF!ureaperkg) + Val(flddolomite * rsF!dolomiteperkg), 0) '22
            fldkg = Round((fldtotalplant * rsF!kg), 0) '23
            fldamountnu2 = Round((fldkg * rsF!amountnu), 0) '24
            fldtotalamount = fldamountnu1 + fldamountnu2 '25
            fldfarmercode = Mid(mygrid.TextMatrix(i, 1), 1, 14)
            fldschedule = ""
            fldsubtotindicator = ""
            fldnewold = ""
            fldstatus = ""
            
            
            
            SQLSTR = "insert into tblplantdistributiondetail(trnid,year,mnth,sno,distno," _
                         & "farmercode,area,totalplant,crateno,bcrate,ecrate,bno,plno,crate,ssp," _
                         & "mop,urea,dolomite,totalkg1,amountnu1,kg,amountnu2,totalamount," _
                         & "schedule,serialmatch,subtotindicator,newold,assignedAtField,ismanual,operation)values(" _
                         & "'" & fldtrnid & "','" & fldyear & "','" & fldmnth & "','" & oldSno & "'," _
                         & "'" & flddistno & "','" & fldfarmercode & "','" & fldarea & "','" & fldtotalplant & "'," _
                         & "'" & fldcrateno & "','" & fldbcrate & "','" & fldecrate & "','" & fldbno & "','" & fldplno & "'," _
                         & "'" & fldcrate & "','" & fldssp & "','" & fldmop & "','" & fldurea & "','" & flddolomite & "'," _
                         & "'" & fldtotalkg1 & "','" & fldamountnu1 & "','" & fldkg & "','" & fldamountnu2 & "','" & fldtotalamount & "'," _
                         & "'" & fldschedule & "','" & flddistno & "','" & fldsubtotindicator & "','" & fldnewold & "','Y','" & isManual & "','Y')"
                         MHVDB.Execute SQLSTR
                        oldSno = oldSno + 1
            Next


End If

End Sub
Private Sub initvar()
oldSno = 0
fldtrnid = 0
fldyear = 0
fldmnth = 0
fldsno = 0
flddistno = 0
fldtotalplant = 0
fldcrateno = 0
fldbcrate = 0
fldecrate = 0
fldbno = 0
fldplno = 0
fldcrate = 0
fldserialmatch = 0
fldarea = 0
fldssp = 0
fldmop = 0
fldurea = 0
flddolomite = 0
fldtotalkg1 = 0
fldamountnu1 = 0
fldkg = 0
fldamountnu2 = 0
fldtotalamount = 0
fldfarmercode = ""
fldschedule = ""
fldsubtotindicator = ""
fldnewold = ""
fldstatus = ""

End Sub
Private Sub clearcontrol()
mygrid.Clear
mygrid.Rows = 20
mygrid.FormatString = "^SL.NO.|^FARMER ID|^AREA|^NO. OF TREES|B CRT.|E CRT.|N CRT.|"
mygrid.ColWidth(0) = 750
mygrid.ColWidth(1) = 2835
mygrid.ColWidth(2) = 930
mygrid.ColWidth(3) = 1455
mygrid.ColWidth(4) = 960
mygrid.ColWidth(5) = 960
mygrid.ColWidth(6) = 960
mygrid.ColWidth(7) = 300
cbotrnid.Text = ""
cbodeliveryno.Text = ""
cbofarmerid.Text = ""
txtassignedqty.Text = ""
txtreturnedqty.Text = ""
chkassignednext.Enabled = False
chkassignednext.Value = 0
chkbacktonursary.Enabled = False
chkbacktonursary.Value = 0
End Sub
 
    
    
     
