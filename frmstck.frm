VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmsTOCKeNTRY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK ENTRY"
   ClientHeight    =   6930
   ClientLeft      =   4680
   ClientTop       =   1275
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmstck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   11415
   Visible         =   0   'False
   Begin MSComCtl2.DTPicker custdate 
      Height          =   375
      Left            =   4320
      TabIndex        =   26
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   80412673
      CurrentDate     =   37359
   End
   Begin VB.TextBox txtpart 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7440
      MaxLength       =   25
      TabIndex        =   25
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox txtCustom 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   22
      Top             =   2400
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker LBLnOW 
      Height          =   360
      Left            =   3720
      TabIndex        =   19
      Top             =   675
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      _Version        =   393216
      Format          =   80412673
      CurrentDate     =   36797
   End
   Begin VB.TextBox txtChallan 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   3
      ToolTipText     =   "Enter Challan No"
      Top             =   1440
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker txtChDate 
      Height          =   315
      Index           =   0
      Left            =   7440
      TabIndex        =   2
      ToolTipText     =   "Enter Challan Date"
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   80412673
      CurrentDate     =   36383
      MinDate         =   36161
   End
   Begin VB.TextBox txtChallan 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   1
      ToolTipText     =   "Enter Challan No"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtRemark 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7440
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   11055
      Begin MSDataListLib.DataCombo CboItemDesc 
         Bindings        =   "frmstck.frx":076A
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1440
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
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
      End
      Begin VB.TextBox txtAmt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7800
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox cboItemCode 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid ItemGrd 
         Height          =   3660
         Left            =   0
         TabIndex        =   0
         Top             =   120
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6456
         _Version        =   393216
         Rows            =   201
         Cols            =   8
         RowHeightMin    =   315
         ForeColorFixed  =   -2147483635
         ScrollTrack     =   -1  'True
         Enabled         =   0   'False
         HighLight       =   0
         FormatString    =   $"frmstck.frx":077F
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
         TabIndex        =   10
         Top             =   4920
         Width           =   870
      End
      Begin VB.Line Line2 
         X1              =   5280
         X2              =   8880
         Y1              =   4290
         Y2              =   4290
      End
   End
   Begin MSDataListLib.DataCombo CboBillNo 
      Bindings        =   "frmstck.frx":0848
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1920
      TabIndex        =   27
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo CboPO 
      Bindings        =   "frmstck.frx":085D
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   7440
      TabIndex        =   28
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DBcboParty 
      Bindings        =   "frmstck.frx":0872
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1920
      TabIndex        =   30
      Top             =   1920
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstck.frx":0887
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstck.frx":0C21
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstck.frx":0FBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstck.frx":1C95
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstck.frx":20E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstck.frx":28A1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
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
   Begin VB.Label Label7 
      Caption         =   "Particulars"
      Height          =   255
      Left            =   6240
      TabIndex        =   24
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Date"
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Custom Entry No."
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblyr 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1755
      TabIndex        =   18
      Top             =   720
      Width           =   180
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   9720
      TabIndex        =   17
      Top             =   6600
      Width           =   240
   End
   Begin VB.Label Label4 
      Caption         =   "Purchase Order no"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   16
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Gate Pass No."
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
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Challan Date"
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
      Index           =   2
      Left            =   6120
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Challan No."
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
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblParty 
      AutoSize        =   -1  'True
      Caption         =   "Supplier"
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
      Top             =   1920
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Entry No."
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
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmsTOCKeNTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BillDetailRec As New ADODB.Recordset
Dim Bill As New ADODB.Recordset
Dim PORec As New ADODB.Recordset
Dim rsbrbill As New ADODB.Recordset
Dim rsInvItem As New ADODB.Recordset
Dim dataPO As New ADODB.Recordset
Dim DatBrBill As New ADODB.Recordset
Dim CurrRow, Jkey, ErrCTR As Long
Dim ValidRow As Boolean

Dim ltot As Double
Const fmString = "        |^  Code        |^                                            Item Name                                  |^   Unit   |^       Qty      |^    Pur. Rate        |^      Amount         |       "
Private Sub PrintBill()
Dim Supp
Dim Party As ADODB.Recordset
Dim i, Lin As Integer
Dim pfile As String
pfile = "SE" + Trim(CboBillNo) + ".txt"
On Error GoTo jerr:
Do While True
   Select Case MsgBox("Printer Ready ? ", vbYesNoCancel)
          Case vbYes
               Open "lpt1:" For Output As #1
               Exit Do
          Case vbNo
               If MsgBox("Connect/Swich on the printer and retry.", vbRetryCancel) = vbCancel Then
                  MsgBox "Bill stored to " + pfile
                  Open pfile For Output As #1
                  Exit Do
               End If
          Case vbCancel
               MsgBox "Bill stored to " + pfile
               Open pfile For Output As #1
               Exit Do
   End Select
Loop
'Open "lpt1:" For Output As #1
Set Party = MHVDB.Execute("SELECT * FROM supplier WHERE suplCODE='" & DBcboParty.BoundText & "'", dbOpenDynaset, DBReadOnly)
If Party.EOF Then
   MsgBox "Not a valid Supplier !!!"
Else
   Supp = Party!Name + " ( " + DBcboParty.BoundText + " )"
End If
Lin = 61
For i = 1 To ItemGrd.Rows - 1
    If Len(Trim(ItemGrd.TextMatrix(i, 1))) = 0 Then Exit For
    If Lin > 60 Then
       Print #1, Chr(18);
       Print #1, Chr(14) + "MOUNTAIN HAZELNUTS VENTURE PRIVATE LIMITED." + Chr(20)
       Print #1, "LINGMETHANG:BHUTAN"
       Print #1, "Stock Entry Slip"
       Print #1, String(79, "_")
       Print #1, PadWithChar("Entry No.", 12, " ", 0) + PadWithChar(lblYR + Trim(CboBillNo), 16, " ", 0) + " " + PadWithChar(" Date ", 5, " ", 0) + Format(lblnow, "dd/mm/yyyy") + "  " + PadWithChar("Purchase Order No.", 19, " ", 0) + PadWithChar(": " + CboPO, 16, " ", 0)
       Print #1, PadWithChar("Challan No.", 12, " ", 0) + PadWithChar(Trim(txtChallan(0)), 16, " ", 0) + " " + PadWithChar(" Date ", 5, " ", 0) + Format(txtChDate(0), "dd/mm/yyyy") + "  " + PadWithChar("Gate Pass No.", 19, " ", 0) + PadWithChar(": " + txtChallan(1), 16, " ", 0)
       Print #1, PadWithChar("Invoice No.", 12, " ", 0) + PadWithChar("", 16, " ", 0) + " " + PadWithChar(" Date ", 5, " ", 0)
       Print #1, "Supplier :  " + Supp
       Print #1, "Remarks  :  " + txtRemark
       Print #1, String(79, "_")
       Print #1, "Srl.|Code  |Description                    |    Qty.      |  Rate  |     Amount  "
       Print #1, String(79, "_")
       Lin = 12
    End If
    Print #1, PadWithChar(Str(i), 4, " ", 1) + " " + PadWithChar(ItemGrd.TextMatrix(i, 1), 7, " ", 0);
    Print #1, PadWithChar(ItemGrd.TextMatrix(i, 2), 31, " ", 0) + PadWithChar(ItemGrd.TextMatrix(i, 4), 7, " ", 1) + " " + PadWithChar(ItemGrd.TextMatrix(i, 3), 7, " ", 0);
    Print #1, PadWithChar(Round(ItemGrd.TextMatrix(i, 5), 2), 9, " ", 1) + " " + PadWithChar(Round(ItemGrd.TextMatrix(i, 6), 2), 11, " ", 1)
    If Len(ItemGrd.TextMatrix(i, 2)) > 31 Then
       Print #1, "     " + Mid(ItemGrd.TextMatrix(i, 2), 32)
       Lin = Lin + 1
    End If
    Lin = Lin + 1
Next
Print #1, String(79, "_")
Print #1, PadWithChar("Total", 57, " ", 1) + PadWithChar(lblTot, 22, " ", 1)
Print #1, String(79, "_")
Print #1,
Print #1,
Print #1, "Prepared By    Checked By    Sr. Manager(S&P)    Sr. Manager(F&A)   Chief Executive"
Print #1, String(79, "_")
Print #1,
Print #1,
Print #1,
Close #1 '*/
Exit Sub
jerr:
MsgBox ERR.Description
ERR.Clear
End Sub



Private Sub CboBillNo_GotFocus()
'Frame2.Enabled = False
End Sub

Private Sub CboBillNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub CboBillNo_LostFocus()
Dim i As Integer
Dim Issue, Recv As Double
Dim imast As ADODB.Recordset
If Operation = "ADD" Then Exit Sub
ltot = 0
Set Bill = MHVDB.Execute("select * from tranhdr where procyear='" & SysYear & "' and ((billno))=('" & CboBillNo & "') AND billtype = 'EN' and status<>'C'")
If Bill.EOF Then
   MsgBox CboBillNo + " Does not exists "
   'CboBillNo.SetFocus
   Exit Sub
Else
   If Bill!ProcYear <> Val(SysYear) Then
      MsgBox " Please login For  " + Str(Bill!ProcYear)
      CboBillNo.SetFocus
      Exit Sub
   End If
   If Bill!Status = "F" Then
      MsgBox "This Bill already finalised. U cant modify it !!!"
      Exit Sub
   End If
   With Bill
   lblnow = Format(!billdate, "dd/mm/yyyy") ' DatKotHead.Recordset!Time
   txtChallan(0) = !challanno
   txtChallan(1) = !gatepassno
   txtChDate(0) = !challandate
   CboPO = IIf(IsNull(!purorderid), "", !purorderid)
   txtRemark = IIf(IsNull(!remarks), "", !remarks)
   txtCustom = IIf(IsNull(!customno), "", !customno)
   txtpart = IIf(IsNull(!PARTICULARS), "", !PARTICULARS)
   custdate = IIf(IsNull(!customDate), Format(Date, "dd/mm/yyyy"), !customDate)
   DBcboParty = IIf(IsNull(!suplcode), "", !suplcode)
   End With
   Set BillDetailRec = MHVDB.Execute("select d.itemcode,d.qty,d.rate,B.unit,itemname,avgstockrate from tranfile as d,invitems as b where d.itemcode=b.itemcode and billno=('" & CboBillNo & "') and d.billtype='EN' and d.procyear='" & SysYear & "'", dbOpenForwardOnly)
   ItemGrd.Clear
   ItemGrd.FormatString = fmString
   With BillDetailRec
   i = 1
   ltot = 0
   Do While Not .EOF
      ItemGrd.TextMatrix(i, 0) = i
      ItemGrd.TextMatrix(i, 1) = !itemcode
      ItemGrd.TextMatrix(i, 2) = !ITEMNAME
      ItemGrd.TextMatrix(i, 3) = !unit
      ItemGrd.TextMatrix(i, 4) = !qty
      ItemGrd.TextMatrix(i, 5) = !Rate
      ItemGrd.TextMatrix(i, 6) = Round(!qty * !Rate, 2)
      ltot = ltot + !qty * !Rate
      .MoveNext
      i = i + 1
   Loop
   End With
'   If MsgBox("Cant be modified.You can Cancel or Print It ! Do you want Print ?", vbYesNo) = vbYes Then
'      PrintBill
'   End If
'   Operation = ""
  ' txtRemark.SetFocus
End If
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
ItemGrd.Enabled = True
Frame2.Enabled = True

TB.Buttons(3).Enabled = True

TB.Buttons(4).Enabled = True
End Sub

Private Sub cboItemCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub cboItemCode_LostFocus()
Dim prevamt, CurrAmt, Jstock As Double
cboItemCode.Text = UCase(cboItemCode.Text)
If ItemGrd.TextMatrix(CurrRow, 1) = cboItemCode Then
   cboItemCode.Visible = False
   Exit Sub
End If
prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
If Len(Trim(CboPO)) > 0 Then
   Set PORec = MHVDB.Execute("select d.itemcode from purorderdtl as d,purorderhdr as a where str(a.purordno)+'/'+str(a.procyear)='" & CboPO & "' and a.status='ON' and  a.procyear=d.procyear and a.purordno=d.purordno and d.itemcode=itemcode = '" & CboItemDesc.BoundText & "'", DBReadOnly)
   If PORec.EOF Then
      If MsgBox("This Item is not in the Purchase Order " + CboPO + " ! Do u want to Enter this ?", vbYesNo) = vbNo Then
         ItemGrd.TextMatrix(CurrRow, 1) = ""
         ItemGrd.TextMatrix(CurrRow, 2) = ""
         ItemGrd.TextMatrix(CurrRow, 3) = ""
         ItemGrd.TextMatrix(CurrRow, 4) = ""
         ItemGrd.TextMatrix(CurrRow, 5) = ""
         ItemGrd.TextMatrix(CurrRow, 6) = 0
         txtQty = ""
         txtRate = ""
         ValidRow = True
         Exit Sub
      End If
   End If
   Set PORec = Nothing
End If

'datInvItem.Recordset.FindFirst "trim(itemcode) = trim('" & cboItemCode & "')"
 rsInvItem.Find "ITEMCODE='" & cboItemCode & "'", , adSearchForward, 1
With rsInvItem
If .EOF Then
   ItemGrd.TextMatrix(CurrRow, 1) = ""
   ItemGrd.TextMatrix(CurrRow, 2) = ""
   ItemGrd.TextMatrix(CurrRow, 3) = ""
   ItemGrd.TextMatrix(CurrRow, 4) = ""
   ItemGrd.TextMatrix(CurrRow, 5) = ""
   ItemGrd.TextMatrix(CurrRow, 6) = 0
   txtQty = ""
   txtRate = ""
   ValidRow = True
Else
   ItemGrd.TextMatrix(CurrRow, 1) = !itemcode
   ItemGrd.TextMatrix(CurrRow, 2) = !ITEMNAME
   ItemGrd.TextMatrix(CurrRow, 3) = IIf(IsNull(!unit), "Nos", !unit)
   txtQty = ItemGrd.TextMatrix(CurrRow, 4)
   If Not Val(txtQty) > 0 Then
      ValidRow = False
      ItemGrd.row = CurrRow
      txtQty.Top = ItemGrd.Top + ItemGrd.CellTop
      txtQty.Visible = True
      txtQty.SetFocus
   End If

End If
End With
CurrAmt = Val(ItemGrd.TextMatrix(i, 6))
ltot = Round(ltot + CurrAmt - prevamt, 2)
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
cboItemCode.Visible = False
End Sub

Private Sub CboItemDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub cboItemDesc_LostFocus()
Dim MSTR As String
Dim prevamt, CurrAmt, Jstock As Double
Dim imast As ADODB.Recordset
If ItemGrd.TextMatrix(CurrRow, 1) = CboItemDesc.BoundText Then
   CboItemDesc.Visible = False
   Exit Sub
End If
prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
If Len(Trim(CboPO)) > 0 Then
MSTR = "select d.itemcode from purorderdtl as d,purorderhdr as a where LTRIM(str(a.purordno))+'/'+ LTRIM(str(a.procyear))='" & CboPO & "' and a.status='ON' and  a.procyear=d.procyear and a.purordno=d.purordno and d.itemcode=itemcode = '" & CboItemDesc.BoundText & "'"
   Set PORec = MHVDB.Execute("select d.itemcode from purorderdtl as d,purorderhdr as a where LTRIM(str(a.purordno))+'/'+ LTRIM(str(a.procyear))='" & CboPO & "' and a.status='ON' and  a.procyear=d.procyear and a.purordno=d.purordno and d.itemcode= '" & CboItemDesc.BoundText & "'", DBReadOnly)
   If PORec.EOF Then
      If MsgBox("This Item is not in the Purchase Order " + CboPO + " ! Do u want to Enter this ?", vbYesNo) = vbNo Then
         ItemGrd.TextMatrix(CurrRow, 1) = ""
         ItemGrd.TextMatrix(CurrRow, 2) = ""
         ItemGrd.TextMatrix(CurrRow, 3) = ""
         ItemGrd.TextMatrix(CurrRow, 4) = ""
         ItemGrd.TextMatrix(CurrRow, 5) = ""
         ItemGrd.TextMatrix(CurrRow, 6) = 0
         txtQty = ""
         txtRate = ""
         ValidRow = True
         Exit Sub
      End If
   End If
   Set PORec = Nothing
End If
'datInvItem.Recordset.FindFirst "itemcode = '" & CboItemDesc.BoundText & "'"
 rsInvItem.Find "ITEMCODE='" & CboItemDesc.BoundText & "'", , adSearchForward, 1
With rsInvItem
If .EOF Then
   ItemGrd.TextMatrix(CurrRow, 1) = ""
   ItemGrd.TextMatrix(CurrRow, 2) = ""
   ItemGrd.TextMatrix(CurrRow, 3) = ""
   ItemGrd.TextMatrix(CurrRow, 4) = ""
   ItemGrd.TextMatrix(CurrRow, 5) = ""
   ItemGrd.TextMatrix(CurrRow, 6) = 0
   txtQty = ""
   txtRate = ""
   ValidRow = True
Else
   ItemGrd.TextMatrix(CurrRow, 1) = !itemcode
   ItemGrd.TextMatrix(CurrRow, 2) = !ITEMNAME
   ItemGrd.TextMatrix(CurrRow, 3) = IIf(IsNull(!unit), "Nos", !unit)
   txtQty = ItemGrd.TextMatrix(CurrRow, 4)
   If Not Val(txtQty) > 0 Then
      ValidRow = False
      ItemGrd.row = CurrRow
      txtQty.Top = ItemGrd.Top + ItemGrd.CellTop
      txtQty.Visible = True
      txtQty.SetFocus
   End If
End If
End With
CurrAmt = Val(ItemGrd.TextMatrix(i, 6))
ltot = Round(ltot + CurrAmt - prevamt, 2)
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
CboItemDesc.Visible = False
End Sub




Private Sub CboPO_Validate(Cancel As Boolean)
If Len(Trim(CboPO)) > 0 Then
   ltot = 0
   Set PORec = MHVDB.Execute("select a.suplcode,d.itemcode,d.qty,d.rate,B.unit,itemname from purorderdtl as d,purorderhdr as a,invitems as b where concat(cast(A.purOrdNo as char),'/',cast(A.procyear as char))='" & CboPO & "' and a.status='ON' and  a.procyear=d.procyear and a.purordno=d.purordno and d.itemcode=b.itemcode ", dbOpenForwardOnly)
   If PORec.EOF Then
      MsgBox "Wrong Purchase Order No!!!"
      Cancel = True
      Exit Sub
   End If
   DBcboParty = PORec!suplcode
   ItemGrd.Clear
   ItemGrd.FormatString = fmString
   i = 1
   With PORec
   Do While Not .EOF
      ItemGrd.TextMatrix(i, 0) = i
      ItemGrd.TextMatrix(i, 1) = !itemcode
      ItemGrd.TextMatrix(i, 2) = !ITEMNAME
      ItemGrd.TextMatrix(i, 3) = !unit
      ItemGrd.TextMatrix(i, 4) = !qty
      ItemGrd.TextMatrix(i, 5) = Format(!Rate, "####0.00")
      ItemGrd.TextMatrix(i, 6) = !qty * !Rate
      ltot = ltot + !qty * !Rate
      i = i + 1
      .MoveNext
   Loop
   End With
   lblTot.Caption = Format(ltot, "###,##,##,##0.00")
   Set PORec = Nothing
End If
End Sub

Private Sub DBcboParty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub DBcboParty_Validate(Cancel As Boolean)
Dim Party As New ADODB.Recordset
If Len(Trim(DBcboParty.BoundText)) = 0 Then
   MsgBox "Supplier should not be blank "
   Exit Sub
End If
Set Party = MHVDB.Execute("SELECT * FROM supplier WHERE suplCODE='" & DBcboParty.BoundText & "'")
If Party.EOF Then
   MsgBox "Not a valid Supplier !!!"
   Cancel = True
End If
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
' ITEM
If rsInvItem.State = adStateOpen Then rsInvItem.Close
rsInvItem.Open "select *  from invitems order by itemname", db, adOpenForwardOnly, adLockReadOnly
Set CboItemDesc.RowSource = rsInvItem
CboItemDesc.ListField = "ITEMNAME"
CboItemDesc.BoundColumn = "ItemCode"


' BILLNO
If DatBrBill.State = adStateOpen Then DatBrBill.Close
 DatBrBill.Open "select * from tranhdr where status <> 'C' and billtype='EN' and procyear='" & SysYear & "' order by billno desc", db
Set CboBillNo.RowSource = DatBrBill
CboBillNo.ListField = "billno"
CboBillNo.BoundColumn = "billno"

'supplier
 
If rsbrbill.State = adStateOpen Then rsbrbill.Close
 rsbrbill.Open ("SELECT * FROM supplieR"), db

Set DBcboParty.RowSource = rsbrbill
DBcboParty.ListField = "Name"
DBcboParty.BoundColumn = "SuplCode"
'PO BILL NO
If dataPO.State = adStateOpen Then dataPO.Close
 dataPO.Open "SELECT purOrdNo,CONCAT(cast(purOrdNo as char), '/' , cast(procyear as char)) as po  FROM purorderhdr where status='ON' order by purordno desc  ", db
Set CboPO.RowSource = dataPO
CboPO.ListField = "PO"
CboPO.BoundColumn = "purOrdNo"

ValidRow = True
lblYR = "EN\" & SysYear & "\"
CurrRow = 1
ItemGrd.row = 1
ItemGrd.Col = 1
cboItemCode.Left = ItemGrd.Left + ItemGrd.CellLeft
cboItemCode.Width = ItemGrd.CellWidth
cboItemCode.Height = ItemGrd.CellHeight
ItemGrd.Col = 2
CboItemDesc.Left = ItemGrd.Left + ItemGrd.CellLeft
CboItemDesc.Width = ItemGrd.CellWidth
CboItemDesc.Height = ItemGrd.CellHeight
ItemGrd.Col = 4
txtQty.Left = ItemGrd.Left + ItemGrd.CellLeft
txtQty.Width = ItemGrd.CellWidth
txtQty.Height = ItemGrd.CellHeight
ItemGrd.Col = 5
txtRate.Left = ItemGrd.Left + ItemGrd.CellLeft
txtRate.Width = ItemGrd.CellWidth
txtRate.Height = ItemGrd.CellHeight
ItemGrd.Col = 6
txtAmt.Left = ItemGrd.Left + ItemGrd.CellLeft
txtAmt.Width = ItemGrd.CellWidth
txtAmt.Height = ItemGrd.CellHeight
End Sub

Private Sub ItemGrd_Click()
Dim jrow, jCol As Integer
If Not ValidRow And CurrRow <> ItemGrd.row Then
   ItemGrd.row = CurrRow
   Exit Sub
End If
jrow = ItemGrd.row
jCol = ItemGrd.Col
If jrow = 0 Then Exit Sub
If jrow > 1 And Len(ItemGrd.TextMatrix(jrow - 1, 1)) = 0 Then
   Beep
   Exit Sub
End If
If CurrRow > ItemGrd.Rows - 2 Then
   ItemGrd.Rows = CurrRow + 3
End If
ItemGrd.TextMatrix(CurrRow, 0) = CurrRow
CurrRow = jrow
ItemGrd.TextMatrix(CurrRow, 0) = Chr(174)
Select Case jCol
       Case 1
            cboItemCode.Top = ItemGrd.Top + ItemGrd.CellTop
            cboItemCode = ItemGrd.Text
            cboItemCode.Visible = True
            cboItemCode.SetFocus
       Case 2
            CboItemDesc.Top = ItemGrd.Top + ItemGrd.CellTop
            CboItemDesc = ItemGrd.Text
            CboItemDesc.BoundText = ItemGrd.TextMatrix(CurrRow, 1)
            CboItemDesc.Visible = True
            CboItemDesc.SetFocus
       Case 4
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtQty.Top = ItemGrd.Top + ItemGrd.CellTop
               txtQty = ItemGrd.Text
               txtQty.Visible = True
               txtQty.SetFocus
            End If
       Case 5
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtRate.Top = ItemGrd.Top + ItemGrd.CellTop
               txtRate = ItemGrd.Text
               txtRate.Visible = True
               txtRate.SetFocus
            End If
        Case 6
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtAmt.Top = ItemGrd.Top + ItemGrd.CellTop
               txtAmt = ItemGrd.Text
               txtAmt.Visible = True
               txtAmt.SetFocus
            End If
    End Select
End Sub

Private Sub ItemGrd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And Shift = 2 Then
If MsgBox("Are u sure to Delete this row ?", vbYesNo) = vbNo Then Exit Sub
   If CurrRow > 0 And Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
      ltot = ltot - Val(ItemGrd.TextMatrix(CurrRow, 6))
      lblTot.Caption = Format(ltot, "###,##,##,##0.00")
      ItemGrd.RemoveItem CurrRow
      ItemGrd.AddItem ""
   Else
      Beep
      Beep
   End If
End If
End Sub

Private Sub ItemGrd_Scroll()
'SendKeys "{TAB}", True
End Sub

Private Sub mnuadd_Click()




Dim lastbill As ADODB.Recordset
lblnow = Format(Now, "dd/mm/yyyy")
ValidRow = True
Operation = "ADD"
CurrRow = 1
txtRemark = ""
txtNoTbl = ""
txtpax = ""
txtBillTo = ""
ltot = 0
ErrCTR = 0
txtChDate(0) = Format(Now, "dd/mm/yyyy")
'txtChDate(1) = Format(Now, "dd/mm/yyyy")
cboItemCode.Visible = False
CboItemDesc.Visible = False
txtQty.Visible = False
Set lastbill = MHVDB.Execute("select max(billno) as lno from tranhdr where procyear='" & SysYear & "' and billtype='EN' ")
CboBillNo = IIf(IsNull(lastbill!lno), 1, lastbill!lno + 1)
Set lastbill = Nothing
CboBillNo.Enabled = False
Frame2.Enabled = True
ItemGrd.Enabled = True
ItemGrd.Clear
ItemGrd.FormatString = fmString
TB.Buttons(4).Enabled = False
TB.Buttons(3).Enabled = True
End Sub
Private Sub mnuCancel_Click()




Dim UpdtStr
Dim jrec As ADODB.Recordset
If MsgBox("Cancel it !!!Are u Sure ?", vbYesNo) = vbNo Then Exit Sub
On Error GoTo ERR

UpdtStr = "UPDATE  tranhdr SET STATUS = 'C',REMARKs = '" & txtRemark & "' WHERE  procyear='" & SysYear & "' and billno = ('" & CboBillNo & "') AND billtype = 'EN'"
MHVDB.Execute UpdtStr, dbSeeChanges + dbFailOnError
Set jrec = MHVDB.Execute("select * from tranfile where  procyear='" & SysYear & "' and billno =('" & CboBillNo & "') AND billtype = 'EN'", dbOpenDynaset)
With jrec
Do While Not .EOF
   MHVDB.Execute "update ITEMSTOCK set totpur=totpur-('" & !qty & "') where procyear='" & SysYear & "' and ITEMCODE = '" & !itemcode & "'", dbFailOnError
   .MoveNext
Loop
End With
Frame2.Enabled = False


Operation = ""
CboBillNo.Enabled = False

TB.Buttons(4).Enabled = False
Exit Sub
ERR:
MsgBox "error :" + IIf(IsNull(ERR.Description), " ", ERR.Description)
ERR.Clear

End Sub

Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnuOpen_Click()
Operation = "OPEN"
ItemGrd.Enabled = False
Frame2.Enabled = True
CboBillNo.Enabled = True

TB.Buttons(4).Enabled = False
TB.Buttons(3).Enabled = True
ErrCTR = 0
CboBillNo.Refresh
End Sub
Private Sub mnuSave_Click()
If Len(ItemGrd.TextMatrix(1, 1)) = 0 And Len(ItemGrd.TextMatrix(1, 2)) = 0 Then

MsgBox "Cannot Saved Blank Item Details, Pls Try Again."
Exit Sub
End If






Dim i, j, K As Integer
Dim printNow As Boolean
Dim jrec As ADODB.Recordset
Dim InsStr, JStat, pcODE As String
'If txtQty.Visible Then txtQty_validate
If Not (Operation = "OPEN" Or Operation = "ADD") Then
   Beep
   Exit Sub
End If
If Not ValidRow Then Exit Sub
printNow = True
0:
On Error GoTo ERR

If Operation = "ADD" Then
   
   InsStr = "insert into tranHdR ( procyear,billtype,billno,BILLDATE,suplcode,Status,challanno,challandate,gatepassno,purorderid,remarks,customno,customdate,particulars) values ( '" & SysYear & "','EN','" & CboBillNo & "'," _
                  & " '" & Format(lblnow, "yyyyMMdd") & "','" & DBcboParty.BoundText & "','OK','" & txtChallan(0) & "','" & Format(txtChDate(0), "yyyyMMdd") & "','" & txtChallan(1) & "','" & CboPO & "','" & txtRemark & "','" & txtCustom & "','" & Format(custdate, "yyyyMMdd") & "','" & txtpart & "')"
   MHVDB.Execute InsStr, dbSeeChanges + dbFailOnError
   For i = 1 To 994
       If Len(Trim(ItemGrd.TextMatrix(i, 1))) > 0 Then
          InsStr = "insert into tranfile (procyear,BILLTYPE, billno,itemcode,qty,rate,lrate,amt) values ('" & SysYear & "', 'EN','" & CboBillNo & "'," _
                  & " '" & ItemGrd.TextMatrix(i, 1) & "',('" & ItemGrd.TextMatrix(i, 4) & "'),('" & ItemGrd.TextMatrix(i, 5) & "'),('" & ItemGrd.TextMatrix(i, 5) & "'),('" & ItemGrd.TextMatrix(i, 6) & "'))"
                  
          MHVDB.Execute InsStr, dbSeeChanges + dbFailOnError
          MHVDB.Execute "UPDATE INVITEMS SET AVGSTOCKRATE=('" & ItemGrd.TextMatrix(i, 5) & "') where itemcode='" & ItemGrd.TextMatrix(i, 1) & "'", dbFailOnError
          MHVDB.Execute "update ITEMSTOCK set totpur=totpur+('" & ItemGrd.TextMatrix(i, 4) & "') where procyear='" & SysYear & "' and ITEMCODE = '" & ItemGrd.TextMatrix(i, 1) & "'", dbFailOnError
       Else
          Exit For
       End If
   Next
Else
   InsStr = "update tranHdR set suplcode='" & DBcboParty.BoundText & "',challanno='" & txtChallan(0) & "',challandate='" & Format(txtChDate(0), "yyyyMMdd") & "',customno='" & txtCustom & "',customdate='" & Format(custdate, "yyyyMMdd") & "',particulars='" & txtpart & "', " _
          & " BILLDATE='" & Format(lblnow, "yyyyMMdd") & "' ,gatepassno='" & txtChallan(1) & "',purorderid='" & CboPO & "',remarks='" & txtRemark & "' where  procyear='" & SysYear & "' and billno =( '" & CboBillNo & "') and billtype='EN'"
   MHVDB.Execute InsStr, dbSeeChanges + dbFailOnError
   Set jrec = MHVDB.Execute("select * from tranfile where procyear='" & SysYear & "' and billno =('" & CboBillNo & "') AND BILLTYPE='EN'", dbOpenDynaset)
   With jrec
   Do While Not .EOF
      MHVDB.Execute "update ITEMSTOCK set totpur=totpur-('" & !qty & "') where procyear='" & SysYear & "' and ITEMCODE = '" & !itemcode & "'", dbFailOnError
      .MoveNext
   Loop
   End With
   MHVDB.Execute "delete from tranfile where procyear='" & SysYear & "' and billno =('" & CboBillNo & "') AND BILLTYPE='EN'", dbFailOnError
   For i = 1 To 994
       If Len(Trim(ItemGrd.TextMatrix(i, 1))) > 0 Then
          InsStr = "insert into tranfile (procyear,BILLTYPE, billno,itemcode,qty,rate,lrate,amt) values ('" & SysYear & "', 'EN','" & CboBillNo & "'," _
                  & " '" & ItemGrd.TextMatrix(i, 1) & "',('" & ItemGrd.TextMatrix(i, 4) & "'),('" & ItemGrd.TextMatrix(i, 5) & "'),('" & ItemGrd.TextMatrix(i, 5) & "'),('" & ItemGrd.TextMatrix(i, 6) & "'))"

          MHVDB.Execute InsStr, dbSeeChanges + dbFailOnError
          MHVDB.Execute "UPDATE INVITEMS SET AVGSTOCKRATE=('" & ItemGrd.TextMatrix(i, 5) & "') where itemcode='" & ItemGrd.TextMatrix(i, 1) & "'", dbFailOnError
          MHVDB.Execute "update ITEMSTOCK set totpur=totpur+('" & ItemGrd.TextMatrix(i, 4) & "') where procyear='" & SysYear & "' and ITEMCODE = '" & ItemGrd.TextMatrix(i, 1) & "'", dbFailOnError
       Else
          Exit For
       End If
   Next
End If
printNow = IIf(MsgBox("Print Now ?", vbYesNo) = vbYes, True, False)
If printNow Then PrintBill

'DatBrBill.Refresh
Operation = ""
CboBillNo.Enabled = False
Frame2.Enabled = False

TB.Buttons(4).Enabled = False
TB.Buttons(3).Enabled = False
ErrCTR = 0
Exit Sub
ERR:
ErrCTR = ErrCTR + 1
If ErrCTR > 5 Then
   If DBEngine.Errors.Count > 0 Then
   For Each errLoop In DBEngine.Errors
       MsgBox "Error number: " & errLoop.Number & vbCr & _
       errLoop.Description
   Next errLoop
'Exit Sub
   End If
End If
ERR.Clear

If ErrCTR < 6 Then
   For i = 1 To 1000
       For j = 1 To 9999
       Next
   Next
   GoTo 0
End If
End Sub
Private Sub Tb_ButtonClick(ByVal Button As msComctlLib.Button)
Select Case Button.Key
       Case "ADD"
      
           mnuadd_Click
       Case "OPEN"
           mnuOpen_Click
       Case "SAVE"
      
           mnuSave_Click
         
       Case "DELETE"
          ' mnuCancel_Click
       Case "EXIT"
           Unload Me
End Select
End Sub


Private Sub txtAmt_Validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not (IsNumeric(txtAmt)) Then
   Beep
   MsgBox "Enter a valid Amount !!!"
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   ValidRow = True
End If
End If
prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
ItemGrd.TextMatrix(CurrRow, 6) = Val(txtAmt)
CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 6))
If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
   ItemGrd.TextMatrix(CurrRow, 5) = CurrAmt / Val(ItemGrd.TextMatrix(CurrRow, 4))
Else
   Beep
   MsgBox "Enter a valid Quantity !!!"
   ValidRow = False
   Cancel = True
   Exit Sub
End If
ltot = Round(ltot + CurrAmt - prevamt, 2)
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
txtAmt.Visible = False
End Sub

Private Sub txtChallan_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub txtChDate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
Dim prevamt, CurrAmt As Double
If KeyAscii = 13 Then
   
   If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
   If Not (IsNumeric(txtQty)) Then
      Beep
      MsgBox "Enter a valid Quantity"
      ValidRow = False
      Exit Sub
   Else
      ItemGrd.TextMatrix(CurrRow, 4) = Val(txtQty)
      prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
      ItemGrd.TextMatrix(CurrRow, 6) = Val(txtQty) * Val(ItemGrd.TextMatrix(CurrRow, 5))
      CurrAmt = ItemGrd.TextMatrix(CurrRow, 6)
      ltot = Round(ltot + CurrAmt - prevamt, 2)
      lblTot.Caption = Format(ltot, "###,##,##,##0.00")
      ValidRow = True
   End If
   End If
   txtQty.Visible = False
   ItemGrd.row = CurrRow
   ItemGrd.Col = 5
   txtRate.Top = ItemGrd.Top + ItemGrd.CellTop
   txtRate = ItemGrd.Text
   txtRate.Visible = True
   txtRate.SetFocus
End If
End Sub

Private Sub txtQty_validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not (IsNumeric(txtQty)) Then
   Beep
   MsgBox "Enter a valid Quantity"
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   ItemGrd.TextMatrix(CurrRow, 4) = txtQty
   ValidRow = True
End If
End If
prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
ItemGrd.TextMatrix(CurrRow, 6) = Val(txtQty) * Val((ItemGrd.TextMatrix(CurrRow, 5)))
CurrAmt = ItemGrd.TextMatrix(CurrRow, 6)
ltot = Round(ltot + CurrAmt - prevamt, 2)
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
txtQty.Visible = False
End Sub


Private Sub txtRate_KeyPress(KeyAscii As Integer)
Dim prevamt, CurrAmt As Double
If KeyAscii = 13 Then
   If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
   If Not (IsNumeric(txtRate) Or Val(txtRate) > 0) Then
      Beep
      MsgBox "Enter a valid Rate !!!"
      ValidRow = False
      Exit Sub
   Else
      ItemGrd.TextMatrix(CurrRow, 5) = txtRate
      prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
      ItemGrd.TextMatrix(CurrRow, 6) = Round(Val(txtRate) * (ItemGrd.TextMatrix(CurrRow, 4)), 2)
      CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 6))
      ltot = Round(ltot + CurrAmt - prevamt, 2)
      lblTot.Caption = Format(ltot, "###,##,##,##0.00")
      ValidRow = True
   End If
   End If
   txtRate.Visible = False
   ItemGrd.TextMatrix(CurrRow, 0) = CurrRow
   CurrRow = CurrRow + 1
   If CurrRow > ItemGrd.Rows - 2 Then
      ItemGrd.Rows = CurrRow + 3
   End If
   ItemGrd.row = CurrRow
   ItemGrd.TextMatrix(CurrRow, 0) = Chr(174)
   ItemGrd.Col = 1
   cboItemCode.Top = ItemGrd.Top + ItemGrd.CellTop
   cboItemCode = ItemGrd.Text
   cboItemCode.Visible = True
   cboItemCode.SetFocus
End If
End Sub

Private Sub txtrate_Validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not (IsNumeric(txtRate) Or Val(txtRate) > 0) Then
   Beep
   MsgBox "Enter a valid Rate !!!"
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   ItemGrd.TextMatrix(CurrRow, 5) = txtRate
   ValidRow = True
End If
End If
prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
ItemGrd.TextMatrix(CurrRow, 6) = Round(Val(txtRate) * (ItemGrd.TextMatrix(CurrRow, 4)), 2)
CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 6))
ltot = Round(ltot + CurrAmt - prevamt, 2)
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
txtRate.Visible = False
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   ItemGrd.Enabled = True
   CurrRow = 1
   ItemGrd.row = CurrRow
   ItemGrd.TextMatrix(CurrRow, 0) = Chr(174)
   ItemGrd.Col = 1
   cboItemCode.Top = ItemGrd.Top + ItemGrd.CellTop
   cboItemCode = ItemGrd.Text
   cboItemCode.Visible = True
   cboItemCode.SetFocus
End If
End Sub
