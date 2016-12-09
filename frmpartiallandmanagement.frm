VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmpartiallandmanagement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Partial Plantaion Management"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10455
   Icon            =   "frmpartiallandmanagement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtplantedindetail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      TabIndex        =   26
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Land Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   4680
      TabIndex        =   5
      Top             =   1440
      Width           =   5775
      Begin VB.TextBox txtdate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   3840
         TabIndex        =   25
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox txtcurrentplan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   3840
         TabIndex        =   23
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtbalanceacre 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   3840
         TabIndex        =   22
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox txtacreplanted 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   3840
         TabIndex        =   21
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox txttotalacre 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   3840
         TabIndex        =   20
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtts 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         TabIndex        =   19
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         TabIndex        =   18
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtdz 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         TabIndex        =   17
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtfarmername 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         TabIndex        =   16
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtfarmercode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         TabIndex        =   15
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Date of Registration"
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
         TabIndex        =   24
         Top             =   4920
         Width           =   1725
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Current Planting Plan"
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
         Left            =   1800
         TabIndex        =   14
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Balance Acre"
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
         TabIndex        =   13
         Top             =   3720
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Acre Planted"
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
         TabIndex        =   12
         Top             =   3240
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Acre"
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
         Left            =   2880
         TabIndex        =   11
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tshowog"
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
         Left            =   1320
         TabIndex        =   10
         Top             =   2280
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Gewog"
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
         Left            =   1440
         TabIndex        =   9
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dongkhag"
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
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Farmer Name"
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
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Farmer Code"
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
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Farmer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
      Begin VSFlex7Ctl.VSFlexGrid mygrid 
         Height          =   4935
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   4095
         _cx             =   7223
         _cy             =   8705
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
         FormatString    =   $"frmpartiallandmanagement.frx":0E42
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
      Begin VB.Image imgBtnDn 
         Height          =   240
         Left            =   5040
         Picture         =   "frmpartiallandmanagement.frx":0EC1
         Top             =   960
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBtnUp 
         Height          =   240
         Left            =   5040
         Picture         =   "frmpartiallandmanagement.frx":124B
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin MSDataListLib.DataCombo cbotrnid 
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   ""
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
   Begin MSComctlLib.ImageList IMG 
      Left            =   9720
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
            Picture         =   "frmpartiallandmanagement.frx":15D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpartiallandmanagement.frx":196F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpartiallandmanagement.frx":1D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpartiallandmanagement.frx":29E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpartiallandmanagement.frx":2E35
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpartiallandmanagement.frx":35EF
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
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1164
      ButtonWidth     =   820
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Open"
            Key             =   "OPEN"
            Object.ToolTipText     =   "OPEN/EDIT EXISTING RECORD"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save"
            Key             =   "SAVE"
            Object.ToolTipText     =   "SAVES RECORD"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Delete"
            Key             =   "DELETE"
            Object.ToolTipText     =   "DELETE THE RECORD"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
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
      Caption         =   "Farmer Code"
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
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   1080
   End
End
Attribute VB_Name = "frmpartiallandmanagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chqty As Double
Private Sub cbotrnid_LostFocus()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
If Len(cbotrnid.Text) = 0 Then Exit Sub
 TB.Buttons(3).Enabled = True
 cbotrnid.Enabled = False
Set rs = Nothing
rs.Open "select * from tbllandreg where trnid='" & cbotrnid.BoundText & "'", MHVDB
If rs.EOF <> True Then
FindFA rs!farmerid, "F"
FindDZ Mid(rs!farmerid, 1, 3)
FindGE Mid(rs!farmerid, 1, 3), Mid(rs!farmerid, 4, 3)
FindTs Mid(rs!farmerid, 1, 3), Mid(rs!farmerid, 4, 3), Mid(rs!farmerid, 7, 3)

 
txtfarmercode.Text = rs!farmerid
txtfarmername.Text = FAName
txtdz.Text = Mid(rs!farmerid, 1, 3) & "  " & Dzname
txtge.Text = Mid(rs!farmerid, 4, 3) & "  " & GEname
txtts.Text = Mid(rs!farmerid, 7, 3) & "  " & TsName

txttotalacre.Text = Format(rs!regland, "##0.00")
Set rs1 = Nothing

txtacreplanted.Text = ""
txtbalanceacre.Text = ""
txtcurrentplan.Text = ""
txtdate.Text = Format(rs!regdate, "dd/MM/yyyy")

End If

Set rs = Nothing
rs.Open "select sum(acre) as pt from tbllandregdetail where farmercode='" & Mid(Trim(cbotrnid.Text), 1, 14) & "' and plantedstatus='P'", MHVDB
If rs.EOF <> True Then
txtacreplanted.Text = Format(IIf(IsNull(rs!pt), 0, rs!pt), "##0.00")
Else
txtacreplanted.Text = "0.00"
End If
txtbalanceacre.Text = Format(Val(txttotalacre.Text) - Val(txtacreplanted.Text), "##0.00")

FillGrid

getnew

End Sub
Private Sub FillGrid()
Dim rs As New ADODB.Recordset
Dim plantedacre As Double
Dim isalreadynew As Boolean
Dim i As Integer
Set rs = Nothing
isalreadynew = False
i = 1
plantedacre = 0


mygrid.Clear
mygrid.Rows = 1
mygrid.FormatString = "^Sl.No.|^Acre|^Planted Status|^"
mygrid.ColWidth(0) = 705
mygrid.ColWidth(1) = 990
mygrid.ColWidth(2) = 1860
mygrid.ColWidth(3) = 405


rs.Open "select * from tbllandregdetail where farmercode='" & Mid(Trim(cbotrnid.Text), 1, 14) & "'", MHVDB
Do While rs.EOF <> True
mygrid.Rows = mygrid.Rows + 1
mygrid.TextMatrix(i, 0) = rs!SLNO
mygrid.TextMatrix(i, 1) = rs!acre
Select Case rs!plantedstatus
Case "N"
mygrid.TextMatrix(i, 2) = "New"
isalreadynew = True
Case "P"
mygrid.TextMatrix(i, 2) = "Planted"
plantedacre = plantedacre + rs!acre
End Select
i = i + 1
rs.MoveNext
Loop
txtplantedindetail.Text = plantedacre

If isalreadynew = False Then
mygrid.Rows = mygrid.Rows + 1
mygrid.TextMatrix(i, 0) = i
mygrid.TextMatrix(i, 1) = Format(Val(txttotalacre.Text) - Val(txtplantedindetail.Text), "##0.00")
mygrid.TextMatrix(i, 2) = "New"
End If


End Sub

Private Sub mygrid_DblClick()
  If mygrid.col = 1 And mygrid.TextMatrix(mygrid.row, 2) = "New" Then
  


myinput = InputBox("Enter Acre for the current plan", "Current Acre", mygrid.TextMatrix(mygrid.row, 1))

If myinput > Val(txtbalanceacre.Text) Then
MsgBox "Please check balance acre!"
Exit Sub
End If

            If Not IsNumeric(myinput) Then
           ' MsgBox "Invalid number,Double Click again to enable the input box."
            Else
            mygrid.TextMatrix(mygrid.row, 1) = CDbl(myinput)
            getnew
            End If
            End If




  If mygrid.col = 3 And mygrid.TextMatrix(mygrid.row, 2) = "New" Then
  'mygrid.Editable = flexEDKbdMouse
  'mygrid.ComboList = "Planted"
  Else
mygrid.Editable = flexEDNone
  mygrid.ComboList = ""
  End If
getnew
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       Operation = "ADD"
       CLEARCONTROLL
       populatefarmer
       cbotrnid.Enabled = True
       TB.Buttons(3).Enabled = False
       Case "OPEN"
       Operation = "OPEN"
       CLEARCONTROLL
    TB.Buttons(3).Enabled = False
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
Private Sub populatefarmer()
Dim RSTR As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
'RSTR.Open "select concat(farmerid,'  ',farmername,'  ',cast(trnid as char)) as farmername,trnid  from tbllandreg as a,tblfarmer as b where a.farmerid=b.idfarmer and plantedstatus='P' and a.status not in ('R','D') order by trnid", db
RSTR.Open "select concat(farmerid,'  ',farmername,'  ',cast(trnid as char)) as farmername,trnid  from tbllandreg as a,tblfarmer as b where a.farmerid=b.idfarmer and plantedstatus='P' and a.status not in ('R','D') and farmerid not in(select farmercode from tbllandregdetail where plantedstatus='N') order by trnid", db
Set cbotrnid.RowSource = RSTR
cbotrnid.ListField = "farmername"
cbotrnid.BoundColumn = "trnid"
End Sub
Private Sub MNU_SAVE()
Dim i As Integer
If Len(txtfarmercode.Text) = 0 Then Exit Sub
MHVDB.Execute "delete from tbllandregdetail where farmercode='" & Trim(txtfarmercode.Text) & "' and headerid='" & cbotrnid.BoundText & "'"

For i = 1 To mygrid.Rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
MHVDB.Execute "insert into tbllandregdetail(farmercode,slno,acre,plantedstatus,headerid) values(" _
& " '" & Trim(txtfarmercode.Text) & "','" & mygrid.TextMatrix(i, 0) & "','" & mygrid.TextMatrix(i, 1) & "','" & mygrid.TextMatrix(i, 2) & "','" & cbotrnid.BoundText & "')"

Next
TB.Buttons(3).Enabled = False
End Sub
Private Sub CLEARCONTROLL()
txtplantedindetail.Text = ""

cbotrnid.Text = ""
txtfarmercode.Text = ""
txtfarmername.Text = ""
txtdz.Text = ""
txtge.Text = ""
txtts.Text = ""

txttotalacre.Text = ""
txtacreplanted.Text = ""
txtbalanceacre.Text = ""
txtcurrentplan.Text = ""
txtdate.Text = ""



mygrid.Clear
mygrid.Rows = 1
mygrid.FormatString = "^Sl.No.|^Acre|^Planted Status|^"
mygrid.ColWidth(0) = 705
mygrid.ColWidth(1) = 990
mygrid.ColWidth(2) = 1860
mygrid.ColWidth(3) = 405

End Sub





Private Sub getnew()
Dim tt As Double
Dim i As Integer
tt = 0
For i = 1 To mygrid.Rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
If Trim(mygrid.TextMatrix(i, 2)) = "New" Then
tt = tt + Val(mygrid.TextMatrix(i, 1))
End If

Next

txtcurrentplan.Text = Format(tt, "##0.00")

End Sub
