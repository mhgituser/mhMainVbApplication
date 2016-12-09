VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMABSENTEE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABSENTEE INFORMATION"
   ClientHeight    =   7995
   ClientLeft      =   4515
   ClientTop       =   1530
   ClientWidth     =   12150
   Icon            =   "FRMABSENTEE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   12150
   Begin VB.Frame Frame1 
      Caption         =   "ADMINISTRATIVE LOCATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   12135
      Begin MSDataListLib.DataCombo cboDzongkhag 
         Bindings        =   "FRMABSENTEE.frx":0E42
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   2040
         TabIndex        =   29
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
      Begin MSDataListLib.DataCombo cbogewog 
         Bindings        =   "FRMABSENTEE.frx":0E57
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   8400
         TabIndex        =   30
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DZONGKHAG"
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
         TabIndex        =   32
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label2 
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
         Left            =   6480
         TabIndex        =   31
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ABSENTEE INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   12135
      Begin VB.CheckBox CHKINF 
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
         Left            =   8160
         TabIndex        =   39
         Top             =   2640
         Width           =   255
      End
      Begin VB.CommandButton CMDINF 
         Enabled         =   0   'False
         Height          =   495
         Left            =   8400
         Picture         =   "FRMABSENTEE.frx":0E6C
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   6000
         Picture         =   "FRMABSENTEE.frx":1616
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "REFRESH"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cbosex 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FRMABSENTEE.frx":1DC0
         Left            =   8040
         List            =   "FRMABSENTEE.frx":1DCA
         TabIndex        =   34
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtabsenteename 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   18
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtcid 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtaddress 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   1560
         Width           =   9975
      End
      Begin VB.TextBox txtphone1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txtphone2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   2520
         Width           =   3975
      End
      Begin VB.TextBox txthouseno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   13
         Top             =   1080
         Width           =   3975
      End
      Begin MSDataListLib.DataCombo cboabsenteeid 
         Bindings        =   "FRMABSENTEE.frx":1DDC
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSDataListLib.DataCombo CBOCARETAKER 
         Bindings        =   "FRMABSENTEE.frx":1DF1
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   8040
         TabIndex        =   36
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
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
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "INFLUENTIAL PERSON"
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
         Left            =   6120
         TabIndex        =   40
         Top             =   2640
         Width           =   2040
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CARETAKER"
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
         Left            =   6480
         TabIndex        =   35
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ADSENTEE ID"
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
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ABSENTEE NAME"
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
         TabIndex        =   26
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CID"
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
         TabIndex        =   25
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "PRESENT ADDRESS"
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
         TabIndex        =   24
         Top             =   1680
         Width           =   1830
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "PHONE-1"
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
         TabIndex        =   23
         Top             =   2160
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "PHONE-2"
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
         TabIndex        =   22
         Top             =   2640
         Width           =   825
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "SEX"
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
         Left            =   6480
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "HOUSE NO."
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
         Left            =   6480
         TabIndex        =   20
         Top             =   1200
         Width           =   1050
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "LAND INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   6240
      TabIndex        =   7
      Top             =   5160
      Width           =   6015
      Begin VB.TextBox txttotalarea 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtregland 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL FALLOW LAND ACRE"
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
         TabIndex        =   11
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL REGISTERED LAND ACRE"
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
         TabIndex        =   10
         Top             =   840
         Width           =   2985
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "REMARKS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   6720
      Width           =   12135
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   12015
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "CONTRACT INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   6015
      Begin VB.CheckBox CHKISCONTRACTSIGNED 
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
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin MSComCtl2.DTPicker TXTCONTRACTDATE 
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   79429633
         CurrentDate     =   41208
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "CONTRACT DATE"
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
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "IS CONTRACT SIGNED"
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
         Width           =   2025
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
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMABSENTEE.frx":1E06
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMABSENTEE.frx":21A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMABSENTEE.frx":253A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMABSENTEE.frx":3214
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMABSENTEE.frx":3666
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMABSENTEE.frx":3E20
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMABSENTEE.frx":41BA
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
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   1164
      ButtonWidth     =   1217
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
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
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "PRINT"
            Key             =   "PRINT"
            Object.ToolTipText     =   "PRINTS THE DISPLAYED INFORMATION OF ABSENTEE."
            ImageIndex      =   7
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   3
   End
End
Attribute VB_Name = "FRMABSENTEE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsDz As New ADODB.Recordset
Dim rsGe As New ADODB.Recordset
Dim rsfr As New ADODB.Recordset
Dim rsAb As New ADODB.Recordset

Dim CrName As String
Private Sub cboabsenteeid_LostFocus()
On Error GoTo ERR
cboabsenteeid.Enabled = False
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblabsentee where absenteeid='" & cboabsenteeid.BoundText & "'", MHVDB
If rs.EOF <> True Then
FindDZ Mid(rs!ABSENTEEID, 1, 3)
FindGE Mid(rs!ABSENTEEID, 1, 3), Mid(rs!ABSENTEEID, 4, 3)
FindCaretaker rs!caretakerid
cboDzongkhag.Text = Mid(rs!ABSENTEEID, 1, 3) & " " & Dzname
cbogewog.Text = Mid(rs!ABSENTEEID, 4, 3) & " " & GEname
CBOCARETAKER.Text = rs!caretakerid & " " & CrName
txtabsenteename.Text = rs!ABSENTEENAME
txtcid.Text = rs!cidno
If rs!sex = 0 Then
cbosex.Text = "Male"
ElseIf rs!sex = 1 Then
cbosex.Text = "Female"
End If
txthouseno.Text = IIf(IsNull(rs!HOUSENO), "", rs!HOUSENO)
txtaddress.Text = IIf(IsNull(rs!Address), "", rs!Address)
txtphone1.Text = IIf(IsNull(rs!phone1), "", rs!phone1)
txtphone2.Text = IIf(IsNull(rs!phone2), "", rs!phone2)
CHKISCONTRACTSIGNED.Value = IIf(IsNull(rs!ISCONTRACTSIGNED), "", rs!ISCONTRACTSIGNED)
TXTCONTRACTDATE.Value = IIf(IsNull(rs!CONTRACTDATE), "", rs!CONTRACTDATE)
txttotalarea.Text = Format(IIf(IsNull(rs!TOTALAREA), 0, rs!TOTALAREA), "#####0.00")
txtregland.Text = Format(IIf(IsNull(rs!REGAREA), 0, rs!REGAREA), "####0.00")
txtremarks.Text = IIf(IsNull(rs!remarks), "", rs!remarks)

Set rs = Nothing
 
 rs.Open "select sum(regland) as regland from tbllandregabsentee where farmerid='" & cboabsenteeid.BoundText & "'", MHVDB
 If rs.EOF <> True Then
 txtregland.Text = Format(IIf(IsNull(rs!REGLAND), 0, rs!REGLAND) + Val(txtregland.Text), "####0.00")
 End If
 
Else

MsgBox "Record Not Found."
End If
Exit Sub
ERR:
MsgBox ERR.Description
End Sub

Private Sub cboDzongkhag_GotFocus()
cbogewog.Enabled = True
End Sub

Private Sub cboDzongkhag_LostFocus()
cboDzongkhag.Enabled = False
End Sub

Private Sub cbogewog_LostFocus()
On Error GoTo ERR
Dim Idloc As String
Dim id As String
id = 0
Idloc = ""
Idloc = cboDzongkhag.BoundText & cbogewog.BoundText
If Operation = "ADD" Then
Dim rs As New ADODB.Recordset
If Len(cbogewog.Text) = 0 Then
         MsgBox "Please,Select Gewog."
         cbogewog.SetFocus
        Exit Sub
        End If
        cbogewog.BackColor = vbWhite
        cbogewog.Enabled = False
Set rs = Nothing
rs.Open "select max(substring(absenteeid,8,4)+1) as MaxId from tblabsentee WHERE SUBSTRING(absenteeid,1,6)='" + Idloc + "'", MHVDB, adOpenForwardOnly, adLockOptimistic

If rs.EOF <> True Then
id = IIf(IsNull(rs!MaxID), 1, rs!MaxID)
If Len(id) = 1 Then
id = "000" & id
ElseIf Len(id) = 2 Then
id = "00" & id
ElseIf Len(id) = 3 Then
id = "0" & id

Else

End If
cboabsenteeid.Text = Idloc & "A" & id
Else
cboabsenteeid.Text = Idloc + "A" + "0001"
End If
        
        
ElseIf Operation = "OPEN" Then

Else

End If

Exit Sub
ERR:
MsgBox ERR.Description
End Sub

Private Sub CHKINF_Click()
If CHKINF.Value = 1 Then

CMDINF.Enabled = True





Else

CMDINF.Enabled = False


End If
End Sub

Private Sub CMDINF_Click()
If MsgBox("Do You Want To Save The Influential Person Record", vbYesNo) = vbYes Then


If Len(cboabsenteeid.BoundText) = 0 Then
MsgBox "Please Check The Entries In The Farmer Registration."
Exit Sub
Else

MNU_SAVE
mbypass = True
Mcaretaker = cboabsenteeid.BoundText
FATYPEINF = "A"
frminf.Show 1

End If




Else
cmdnext.Enabled = False
End If
mbypass = False
End Sub

Private Sub Command3_Click()
On Error GoTo ERR
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsDz = Nothing



Set rsAb = Nothing

If rsAb.State = adStateOpen Then rsAb.Close
rsAb.Open "select concat(absenteeid , ' ', absenteename) as absenteename,absenteeid  from tblabsentee order by absenteeid", db
Set cboabsenteeid.RowSource = rsAb
cboabsenteeid.ListField = "absenteename"
cboabsenteeid.BoundColumn = "absenteeid"
Exit Sub
ERR:
MsgBox ERR.Description
End Sub

Private Sub Form_Load()
On Error GoTo ERR
Operation = ""


Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsDz = Nothing

If rsDz.State = adStateOpen Then rsDz.Close
rsDz.Open "select concat(dzongkhagcode , ' ', dzongkhagname) as dzongkhagname,dzongkhagcode  from tbldzongkhag order by dzongkhagcode", db
Set cboDzongkhag.RowSource = rsDz
cboDzongkhag.ListField = "dzongkhagname"
cboDzongkhag.BoundColumn = "dzongkhagcode"

If rsGe.State = adStateOpen Then rsGe.Close
rsGe.Open "select concat(gewogid , ' ', gewogname) as gewogname,gewogid  from tblgewog order by dzongkhagid,gewogid", db
Set cbogewog.RowSource = rsGe
cbogewog.ListField = "gewogname"
cbogewog.BoundColumn = "gewogid"


Set rsfr = Nothing

If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer WHERE ISCARETAKER='1' order by idfarmer", db
Set CBOCARETAKER.RowSource = rsfr
CBOCARETAKER.ListField = "farmername"
CBOCARETAKER.BoundColumn = "idfarmer"

Set rsAb = Nothing

If rsAb.State = adStateOpen Then rsAb.Close
rsAb.Open "select concat(absenteeid , ' ', absenteename) as absenteename,absenteeid  from tblabsentee order by absenteeid", db
Set cboabsenteeid.RowSource = rsAb
cboabsenteeid.ListField = "absenteename"
cboabsenteeid.BoundColumn = "absenteeid"


If mbypass = True Then
cboDzongkhag.Enabled = False
        TB.Buttons(3).Enabled = True
       Operation = "ADD"
       CLEARCONTROLL
       cboDzongkhag.Enabled = True
     cboabsenteeid.Enabled = False
       
      
CBOCARETAKER.Text = Mcaretaker
CBOCARETAKER.Enabled = False
Else

CBOCARETAKER.Enabled = True
End If


If ISCARETAKER = True Then
CHKINF.Enabled = False
Else
CHKINF.Enabled = True
End If







Exit Sub
ERR:
MsgBox ERR.Description
End Sub

Private Sub Tb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
Case "ADD"
       cboDzongkhag.Enabled = False
        TB.Buttons(3).Enabled = True
       Operation = "ADD"
       CLEARCONTROLL
       cboDzongkhag.Enabled = True
     cboabsenteeid.Enabled = False
       
       Case "OPEN"
       Operation = "OPEN"
       CLEARCONTROLL
       cboabsenteeid.Enabled = True
       cboDzongkhag.Enabled = False
       cbogewog.Enabled = False
      TB.Buttons(3).Enabled = True
       
       Case "SAVE"
       If Len(CBOCARETAKER.Text) = 0 Then
        MsgBox "PLEASE SELECT CARETAKER FOR THIS ABSENTEE."
        CBOCARETAKER.SetFocus
        Exit Sub
        End If
       MNU_SAVE
        TB.Buttons(3).Enabled = False
        FillGrid
       TB.Buttons(6).Enabled = True
       Case "DELETE"
         Case "PRINT"
         PRINTINFO
          TB.Buttons(6).Enabled = False
       Case "EXIT"
       Unload Me
       
       
End Select
End Sub
Private Sub PRINTINFO()

On Error Resume Next
Dim excel_app As Object
Dim excel_sheet As Object
Dim row As Long
Dim statement As String
Dim i, j, K As Integer
Screen.MousePointer = vbHourglass

FileCopy App.Path + "\absenteeinfo.XLS", App.Path + "\" + cboabsenteeid.Text + Format(Now, "ddMMyyyy") + ".XLS"
Set excel_app = CreateObject("Excel.Application")
excel_app.Workbooks.Open FileName:=App.Path + "\" + cboabsenteeid.Text + Format(Now, "ddMMyyyy") + ".XLS"
If Val(excel_app.Application.Version) >= 8 Then
   Set excel_sheet = excel_app.ActiveSheet
Else
   Set excel_sheet = excel_app
End If
excel_app.Visible = True
excel_sheet.Cells(5, 2) = cboDzongkhag.Text
excel_sheet.Cells(5, 6) = cbogewog.Text
excel_sheet.Cells(7, 2) = cboabsenteeid.Text
excel_sheet.Cells(7, 6) = CBOCARETAKER.Text
excel_sheet.Cells(9, 2) = "'" & txtcid.Text
excel_sheet.Cells(9, 6) = cbosex.Text '
excel_sheet.Cells(10, 2) = txthouseno.Text
excel_sheet.Cells(11, 2) = txtaddress.Text
excel_sheet.Cells(12, 2) = txtphone1.Text
excel_sheet.Cells(12, 6) = txtphone2.Text
If CHKISCONTRACTSIGNED.Value = 1 Then
excel_sheet.Cells(14, 2) = "YES"
excel_sheet.Cells(14, 6) = "'" & Format(TXTCONTRACTDATE.Value, "dd/MM/yyyy") & "DD/MM/YYYY"
Else
excel_sheet.Cells(14, 2) = "NO"
excel_sheet.Cells(14, 6) = ""

End If
excel_sheet.Cells(15, 2) = "'" & Format(txttotalarea.Text, "#####0.00")
excel_sheet.Cells(15, 6) = "'" & Format(txtregland.Text, "#####0.00")
excel_sheet.Cells(18, 1) = txtremarks.Text
With excel_app.ActiveSheet.Pictures.Insert(App.Path + "\image\" + cboabsenteeid.BoundText & ".jpg")
    With .ShapeRange
        .LockAspectRatio = msoTrue
        .Width = 60
        .Height = 60
    End With
    .Left = excel_app.ActiveSheet.Cells(1, 8).Left
    .Top = excel_app.ActiveSheet.Cells(1, 8).Top
    .Placement = 1
    .PrintObject = True
End With



Screen.MousePointer = vbDefault
End Sub
Private Sub CLEARCONTROLL()
txtabsenteename.Text = ""
If mbypass = False Then
CBOCARETAKER.Text = ""
End If
txtcid.Text = ""
cbosex.Text = ""
txtphone1.Text = ""
txtphone2.Text = ""
txtaddress.Text = ""

txthouseno.Text = ""
txttotalarea.Text = ""
txtregland.Text = ""
TXTCONTRACTDATE.Value = "01/01/1900"
End Sub
Private Sub MNU_SAVE()
On Error GoTo ERR
If Len(cboabsenteeid.Text) = 0 Then
MsgBox "Please Select The Appropriate Information From The Drop Down Controll.", , "MHV ERROR BOX"
Exit Sub
End If
MHVDB.BeginTrans
If Operation = "ADD" Then
MHVDB.Execute "insert into tblabsentee(absenteeid,caretakerid,absenteename,cidno,SEX,houseno,address,phone1," _
& "phone2,iscontractsigned,contractdate,totalarea,regarea,remarks)" _
& "values('" & cboabsenteeid.Text & "','" & CBOCARETAKER.BoundText & "','" & txtabsenteename.Text & "','" & txtcid.Text & "','" & cbosex.ListIndex & "'" _
& " ,'" & txthouseno.Text & "','" & txtaddress.Text & "','" & txtphone1.Text & "','" & txtphone2.Text & "','" & CHKISCONTRACTSIGNED.Value & "','" & Format(TXTCONTRACTDATE.Value, "yyyy-MM-dd") & "'" _
& ",'" & txttotalarea.Text & "','" & txtregland.Text & "','" & txtremarks.Text & "')"



ElseIf Operation = "OPEN" Then

MHVDB.Execute "update tblabsentee set caretakerid='" & CBOCARETAKER.BoundText & "',absenteename='" & txtabsenteename.Text & "',cidno='" & txtcid.Text & "',SEX='" & cbosex.ListIndex & "',houseno='" & txthouseno.Text & "',address='" & txtaddress.Text & "',phone1='" & txtphone1.Text & "'," _
& "phone2='" & txtphone2.Text & "',iscontractsigned='" & CHKISCONTRACTSIGNED.Value & "',contractdate='" & Format(TXTCONTRACTDATE.Value, "yyyy-MM-dd") & "',totalarea='" & txttotalarea.Text & "',regarea='" & txtregland.Text & "',remarks='" & txtremarks.Text & "'  where absenteeid='" & cboabsenteeid.BoundText & "'"

' Val(txtregland.Text)


Else
MsgBox "OPERATION NOT SELECTED."
End If
MHVDB.CommitTrans
mbypass = False
Mcaretaker = ""
Exit Sub
ERR:
MsgBox ERR.Description
MHVDB.RollbackTrans
End Sub

Private Sub FillGrid()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text3_Change()

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text4_Change()

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Text6_Change()

End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtabsenteename_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtabsenteename.SelStart + 1
    Dim sText As String
    sText = Left$(txtabsenteename.Text, iPos)
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

Private Sub txtaddress_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtaddress.SelStart + 1
    Dim sText As String
    sText = Left$(txtaddress.Text, iPos)
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

Private Sub txtcid_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txthouseno_Change()
txthouseno.Text = StrConv(txthouseno.Text, vbUpperCase)
txthouseno.SelStart = Len(txthouseno)
End Sub

Private Sub txtphone1_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtphone2_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtregland_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txttotalarea_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txttshowogname_Change()

End Sub

Private Sub txttshowogname_KeyPress(KeyAscii As Integer)

End Sub
Private Sub FindDZ(dd As String)
On Error GoTo ERR
Dim rs As New ADODB.Recordset
Dzname = ""
Set rs = Nothing
rs.Open "select * from tbldzongkhag where dzongkhagcode='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
Dzname = rs!DZONGKHAGNAME
Else
MsgBox "Record Not Found."
End If
Exit Sub
ERR:
MsgBox ERR.Description
End Sub
Private Sub FindGE(dd As String, GG As String)
On Error GoTo ERR
Dim rs As New ADODB.Recordset
GEname = ""
Set rs = Nothing
rs.Open "select * from tblgewog where dzongkhagID='" & dd & "' AND GEWOGID='" & GG & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
GEname = rs!gewogname
Else
MsgBox "Record Not Found."
End If
Exit Sub
ERR:
MsgBox ERR.Description
End Sub
Private Sub FindCaretaker(dd As String)
On Error GoTo ERR
Dim rs As New ADODB.Recordset
CrName = ""
Set rs = Nothing
rs.Open "select * from tblfarmer where idfarmer='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
CrName = rs!FARMERNAME
Else
MsgBox "Record Not Found."
End If
Exit Sub
ERR:
MsgBox ERR.Description
End Sub
