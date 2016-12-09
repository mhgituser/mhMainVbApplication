VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMINSECTICIDEFUNJICIDE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "I N S E C T I C I D E & F U N J I C I D E"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15330
   Icon            =   "FRMINSECTICIDEFUNJICIDE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   15330
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
      Left            =   12360
      TabIndex        =   30
      Top             =   2880
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
         Picture         =   "FRMINSECTICIDEFUNJICIDE.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   600
         TabIndex        =   32
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   102039553
         CurrentDate     =   41479
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   600
         TabIndex        =   33
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   102039553
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   15255
      Begin VB.TextBox txtreason 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6840
         TabIndex        =   28
         Top             =   1680
         Width           =   8295
      End
      Begin VB.TextBox txtnoofminutes 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   13920
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtareasprayed 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6840
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtchemicalqty 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   9960
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtvolume 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   13920
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker txtstartdate 
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   102039553
         CurrentDate     =   41480
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "FRMINSECTICIDEFUNJICIDE.frx":11CC
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1560
         TabIndex        =   8
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
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   102039554
         CurrentDate     =   41480
      End
      Begin MSComCtl2.DTPicker txtendtime 
         Height          =   375
         Left            =   9960
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   102039554
         CurrentDate     =   41480
      End
      Begin MSDataListLib.DataCombo cbofacilityid 
         Bindings        =   "FRMINSECTICIDEFUNJICIDE.frx":11E1
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1560
         TabIndex        =   11
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
      Begin MSDataListLib.DataCombo cbofertilizer 
         Bindings        =   "FRMINSECTICIDEFUNJICIDE.frx":11F6
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1560
         TabIndex        =   12
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
      Begin MSDataListLib.DataCombo cbostaff 
         Bindings        =   "FRMINSECTICIDEFUNJICIDE.frx":120B
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1560
         TabIndex        =   13
         Top             =   1680
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
      Begin MSDataListLib.DataCombo cboappmethod 
         Bindings        =   "FRMINSECTICIDEFUNJICIDE.frx":1220
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   6840
         TabIndex        =   27
         Top             =   1200
         Width           =   4455
         _ExtentX        =   7858
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "App. Method"
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
         TabIndex        =   26
         ToolTipText     =   "Method of Application"
         Top             =   1320
         Width           =   1095
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
         TabIndex        =   25
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fertilizer Mix"
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
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   615
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
         Left            =   8760
         TabIndex        =   22
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
         Left            =   5400
         TabIndex        =   21
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
         Left            =   3000
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No.of Minutes"
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
         Left            =   12480
         TabIndex        =   18
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Sprayed Area"
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
         TabIndex        =   17
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Chemical Qty"
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
         Left            =   8760
         TabIndex        =   16
         ToolTipText     =   "kg or litres of Chemical"
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Total Volume"
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
         Left            =   12480
         TabIndex        =   15
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Reason"
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
         TabIndex        =   14
         ToolTipText     =   "Reason for applying"
         Top             =   1800
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   15255
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   4095
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   15135
         _cx             =   26696
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
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FRMINSECTICIDEFUNJICIDE.frx":1235
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
            Picture         =   "FRMINSECTICIDEFUNJICIDE.frx":13D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMINSECTICIDEFUNJICIDE.frx":176E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMINSECTICIDEFUNJICIDE.frx":1B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMINSECTICIDEFUNJICIDE.frx":27E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMINSECTICIDEFUNJICIDE.frx":2C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMINSECTICIDEFUNJICIDE.frx":33EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   15330
      _ExtentX        =   27040
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
Attribute VB_Name = "FRMINSECTICIDEFUNJICIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbotrnid_LostFocus()
On Error GoTo err
   
   cbotrnid.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblqmsspray where trnid='" & cbotrnid.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
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
rs.Open "select * from tblqmsspray where trnid = '" & id & "'", MHVDB
If rs.EOF <> True Then
txtstartdate.Value = Format(rs!entrydate, "dd/MM/yyyy")
txtstarttime.Value = Format(rs!starttime, "HH:mm:ss")
txtendtime.Value = Format(rs!endtime, "HH:mm:ss")
txtnoofminutes.Text = IIf(DateTime.DateDiff("n", Format(rs!starttime, "HH:mm:ss"), Format(rs!endtime, "HH:mm:ss")) < 0, "", DateTime.DateDiff("n", Format(rs!starttime, "HH:mm:ss"), Format(rs!endtime, "HH:mm:ss")))
findQmsfacility rs!facilityid
cbofacilityid.Text = rs!facilityid & " " & qmsFacility
txtvolume.Text = rs!totalvol
txtareasprayed.Text = rs!Area
cbofertilizer.Text = rs!chemicalid
txtchemicalqty.Text = rs!chemicalqty
findQmsApplicationMethod rs!applicationmethod
cboappmethod.Text = rs!applicationmethod & " " & qmsApplicationMethod
FindsTAFF rs!staffid
cbostaff.Text = rs!staffid & "  " & sTAFF
txtreason.Text = rs!reason
End If

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
txtstartdate.Value = Format(Now, "dd/MM/yyyy")
txtstarttime.Value = Format(Now, "HH:mm:ss")
txtendtime.Value = Format(Now, "HH:mm:ss")
Set rsF = Nothing

If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select trnId as description  from tblqmsspray order by trnId", db
Set cbotrnid.RowSource = rsF
cbotrnid.ListField = "description"
cbotrnid.BoundColumn = "description"


Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
If rsF.State = adStateOpen Then Srs.Close
rsF.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff WHERE nursery='1'  order by STAFFCODE", db
Set cbostaff.RowSource = rsF
cbostaff.ListField = "STAFFNAME"
cbostaff.BoundColumn = "STAFFCODE"

Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select concat(facilityId , '  ', description) as description,facilityId  from tblqmsfacility order by facilityId", db
Set cbofacilityid.RowSource = rsF
cbofacilityid.ListField = "description"
cbofacilityid.BoundColumn = "facilityId"

Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select concat(cast(fertilizermixno as char),'   ',cast(mixeddate as char)) as description,fertilizermixno  from tblqmsfertilizermixhdr order by fertilizermixno", db
Set cbofertilizer.RowSource = rsF
cbofertilizer.ListField = "description"
cbofertilizer.BoundColumn = "fertilizermixno"

Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select concat(cast(methodid as char),' ',description) as description,methodid  from tblqmsapplicationmethod order by methodid", db
Set cboappmethod.RowSource = rsF
cboappmethod.ListField = "description"
cboappmethod.BoundColumn = "methodid"

txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")
'FillGrid

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
        rs.Open "SELECT MAX(trnid)+1 AS MaxID from tblqmsspray", MHVDB, adOpenForwardOnly, adLockOptimistic
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
       
       ' FillGrid
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub CLEARCONTROLL()
txtstartdate.Value = Format(Now, "dd/MM/yyyy")
txtstarttime.Value = Format(Now, "HH:mm:ss")
txtendtime.Value = Format(Now, "HH:mm:ss")
txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")
txtnoofminutes.Text = ""
cbofacilityid.Text = ""
cbostaff.Text = ""
cboappmethod.Text = ""
cbofertilizer.Text = ""
txtareasprayed.Text = ""
txtvolume.Text = ""
txtchemicalqty.Text = ""
txtreason.Text = ""
End Sub
Private Sub FillGrid(frmdate As Date, todate As Date)
'On Error GoTo err
On Error Resume Next
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^Sl.No.|^Trn. Id|^S.Date|S.Time|^F.Time|^Minutes|^Facility|^Area|^Chemical Qty.|^Total Vol.|^Fertilizer Mix|^App. Method|^Staff Id|^"
Mygrid.ColWidth(0) = 570
Mygrid.ColWidth(1) = 735
Mygrid.ColWidth(2) = 1200
Mygrid.ColWidth(3) = 855
Mygrid.ColWidth(4) = 855
Mygrid.ColWidth(5) = 765
Mygrid.ColWidth(6) = 2025
Mygrid.ColWidth(7) = 585
Mygrid.ColWidth(8) = 1260
Mygrid.ColWidth(9) = 855
Mygrid.ColWidth(10) = 1170
Mygrid.ColWidth(11) = 1605
Mygrid.ColWidth(12) = 2340
Mygrid.ColWidth(13) = 210


rs.Open "select * from tblqmsspray where entrydate>='" & Format(frmdate, "yyyy-MM-dd") & "' and entrydate<='" & Format(todate, "yyyy-MM-dd") & "' order by trnid desc", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i

Mygrid.TextMatrix(i, 1) = rs!trnid
Mygrid.TextMatrix(i, 2) = Format(rs!entrydate, "dd/MM/yyyy")
Mygrid.TextMatrix(i, 3) = Format(rs!starttime, "HH:mm:ss")
Mygrid.TextMatrix(i, 4) = Format(rs!endtime, "HH:mm:ss")
Mygrid.TextMatrix(i, 5) = IIf(DateTime.DateDiff("n", Format(rs!starttime, "HH:mm:ss"), Format(rs!endtime, "HH:mm:ss")) < 0, "", DateTime.DateDiff("n", Format(rs!starttime, "HH:mm:ss"), Format(rs!endtime, "HH:mm:ss")))
findQmsfacility rs!facilityid
Mygrid.TextMatrix(i, 6) = rs!facilityid & " " & qmsFacility

Mygrid.TextMatrix(i, 7) = rs!Area
Mygrid.TextMatrix(i, 8) = rs!chemicalqty
Mygrid.TextMatrix(i, 9) = rs!totalvol

Mygrid.TextMatrix(i, 10) = rs!chemicalid
findQmsApplicationMethod rs!applicationmethod
Mygrid.TextMatrix(i, 11) = qmsApplicationMethod

FindsTAFF rs!staffid
Mygrid.TextMatrix(i, 12) = rs!staffid & " " & sTAFF

rs.MoveNext
i = i + 1
Loop

rs.Close
'Exit Sub
'err:
'MsgBox err.Description

End Sub

Private Sub txtendtime_Change()
txtnoofminutes.Text = IIf(DateTime.DateDiff("n", txtstarttime.Value, txtendtime.Value) < 0, "", DateTime.DateDiff("n", txtstarttime.Value, txtendtime.Value))
End Sub

Private Sub txtstarttime_Change()
txtnoofminutes.Text = IIf(DateTime.DateDiff("n", txtstarttime.Value, txtendtime.Value) < 0, "", DateTime.DateDiff("n", txtstarttime.Value, txtendtime.Value))
End Sub
Private Sub MNU_SAVE()
Dim rs As New ADODB.Recordset
On Error GoTo err
If Len(cbotrnid.Text) = 0 Then
MsgBox "Transaction No. is Must."
Exit Sub
End If

If Val(txtnoofminutes.Text) = 0 Then
MsgBox "Invalid Minutes."
Exit Sub
End If

MHVDB.BeginTrans
If Operation = "ADD" Then
LogRemarks = ""
MHVDB.Execute "INSERT INTO tblqmsspray (trnid,entrydate,starttime,endtime,facilityid," _
            & "area,chemicalid,chemicalqty,totalvol,applicationmethod,reason,staffid,status,location)" _
            & "values(" _
            & "'" & cbotrnid.BoundText & "'," _
            & "'" & Format(txtstartdate.Value, "yyyy-MM-dd") & "'," _
            & "'" & Format(txtstarttime.Value, "HH:mm:ss") & "'," _
            & "'" & Format(txtendtime.Value, "HH:mm:ss") & "'," _
            & "'" & cbofacilityid.BoundText & "'," _
            & "'" & Val(txtareasprayed.Text) & "'," _
            & "'" & cbofertilizer.BoundText & "'," _
            & "'" & Val(txtchemicalqty.Text) & "'," _
            & "'" & Val(txtvolume.Text) & "'," _
            & "'" & cboappmethod.BoundText & "'," _
            & "'" & txtreason.Text & "'," _
            & "'" & cbostaff.BoundText & "'," _
            & "'ON'," _
            & "'" & Mlocation & "'" _
            & ")"
 
 
LogRemarks = "Inserted new record" & cbotrnid.BoundText & "," & Mlocation & ","
updatemhvlog Now, MUSER, LogRemarks, ""

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblqmsspray set " _
            & "entrydate='" & Format(txtstartdate.Value, "yyyy-MM-dd") & "'," _
            & "starttime='" & Format(txtstarttime.Value, "HH:mm:ss") & "'," _
            & "endtime='" & Format(txtendtime.Value, "HH:mm:ss") & "'," _
            & "facilityid='" & cbofacilityid.BoundText & "'," _
            & "area='" & Val(txtareasprayed.Text) & "'," _
            & "chemicalid='" & cbofertilizer.BoundText & "'," _
            & "chemicalqty='" & Val(txtchemicalqty.Text) & "'," _
            & "totalvol='" & Val(txtvolume.Text) & "'," _
            & "applicationmethod='" & cboappmethod.BoundText & "'," _
            & "reason='" & txtreason.Text & "'," _
            & "staffid='" & cbostaff.BoundText & "'" _
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

