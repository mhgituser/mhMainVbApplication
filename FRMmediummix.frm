VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMmediummix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M E D I U M  M I X"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7740
   Icon            =   "FRMmediummix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   7575
      Begin VB.TextBox txtmedium4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   21
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtmedium5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3960
         TabIndex        =   19
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtleafmold 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6600
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtblacksawdust 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3960
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtbrownsawdust 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin MSComCtl2.DTPicker txtenddate 
         Height          =   375
         Left            =   5640
         TabIndex        =   3
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81264641
         CurrentDate     =   41479
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "FRMmediummix.frx":0E42
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1560
         TabIndex        =   4
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
      Begin MSDataListLib.DataCombo cbostaffid 
         Bindings        =   "FRMmediummix.frx":0E57
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   5400
         TabIndex        =   5
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
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
      Begin MSComCtl2.DTPicker txtstartdate 
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81264641
         CurrentDate     =   41479
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Medium 4 (%)"
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
         Top             =   1800
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Medium 5 (%)"
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
         TabIndex        =   18
         Top             =   1800
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Leaf Mold/Soil (%)"
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
         TabIndex        =   16
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Black Saw Dust"
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
         TabIndex        =   14
         Top             =   1320
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Brown Saw Dust"
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
         TabIndex        =   12
         Top             =   1320
         Width           =   1410
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Finish Date"
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
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Medium Mix No."
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
         Top             =   360
         Width           =   1365
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
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Staff"
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
         TabIndex        =   6
         Top             =   1800
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Chemical"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   7335
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7215
         _cx             =   12726
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
         Rows            =   50
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FRMmediummix.frx":0E6C
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
            Picture         =   "FRMmediummix.frx":0EF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMmediummix.frx":128F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMmediummix.frx":1629
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMmediummix.frx":2303
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMmediummix.frx":2755
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMmediummix.frx":2F0F
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
      Width           =   7740
      _ExtentX        =   13653
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
Attribute VB_Name = "FRMmediummix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbotrnid_LostFocus()
On Error GoTo err
   
   cbotrnid.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblqmsmediummixhdr where trnid='" & cbotrnid.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
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
rs.Open "select * from tblqmsmediummixhdr where trnid = '" & id & "'", MHVDB
If rs.EOF <> True Then
   txtstartdate.Value = Format(rs!startdate, "dd/MM/yyyy")
   txtenddate.Value = Format(rs!enddate, "dd/MM/yyyy")
   txtbrownsawdust.Text = IIf(rs!brownsawdust = 0, "", rs!brownsawdust)
   txtblacksawdust.Text = IIf(rs!blacksawdust = 0, "", rs!blacksawdust)
   txtleafmold.Text = IIf(rs!leafmold = 0, "", rs!leafmold)
   txtmedium4.Text = IIf(rs!medium4 = 0, "", rs!medium4)
   txtmedium5.Text = IIf(rs!medium5 = 0, "", rs!medium5)
   FindsTAFF rs!staffid
   cbostaffid.Text = rs!staffid & " " & sTAFF
 
End If
fillgrd id
End Sub
Private Sub fillgrd(id As String)
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
rs.Open "select * from tblqmsmediummixdetail where trnid = '" & id & "'", MHVDB
If rs.EOF <> True Then
i = 1
    Do While rs.EOF <> True
   Mygrid.TextMatrix(i, 0) = i
   FindqmsChemicalTradeName rs!chemicalid
   Mygrid.TextMatrix(i, 1) = qmsChemicalTradeName
   Mygrid.TextMatrix(i, 2) = rs!qty
      
        i = i + 1
        rs.MoveNext
    Loop
End If
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
rsF.Open "select trnid as description  from tblqmsmediummixhdr order by trnid", db
Set cbotrnid.RowSource = rsF
cbotrnid.ListField = "description"
cbotrnid.BoundColumn = "description"

Set rsF = Nothing

If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff where nursery='1' order by STAFFCODE", db
Set cbostaffid.RowSource = rsF
cbostaffid.ListField = "STAFFNAME"
cbostaffid.BoundColumn = "STAFFCODE"



Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Mygrid_Click()
If Mygrid.col = 1 And Len(Mygrid.TextMatrix(Mygrid.row - 1, 2)) > 0 Then
Mygrid.Editable = flexEDKbdMouse
FillGridCombo
Else
Mygrid.ComboList = ""
Mygrid.Editable = flexEDNone
End If

If Len(Mygrid.TextMatrix(Mygrid.row, 1)) > 0 And Mygrid.col = 2 Then
Mygrid.Editable = flexEDKbdMouse
End If
End Sub
Private Sub FillGridCombo()
        Dim RstTemp As New ADODB.Recordset
        Dim NmcItemData As String
        Dim StrComboList As String
        Dim mn1 As String
        StrComboList = "          |"
        
        
            Set RstTemp = Nothing
            RstTemp.Open ("select chemicalid,tradename from tblqmschemicalhdr where status='ON' ORDER BY chemicalid"), MHVDB

            If Not RstTemp.BOF Then
               Do While Not RstTemp.EOF
                    NmcItemData = ""
                    StrComboList = IIf(StrComboList = "", Right("0000" & RstTemp("chemicalid").Value, 3) & " " & RstTemp("tradename").Value, StrComboList & "|" & Right("0000" & RstTemp("chemicalid").Value, 3)) & " " & RstTemp("tradename").Value
                    RstTemp.MoveNext
               Loop
            End If
            RstTemp.Close
       Mygrid.ComboList = StrComboList



    End Sub


Private Sub Mygrid_LeaveCell()
Dim i As Integer
If Trim(Mygrid.TextMatrix(Mygrid.row, 1)) = "" Then
Mygrid.RemoveItem (Mygrid.row)
Mygrid.Rows = Mygrid.Rows + 1
For i = 1 To Mygrid.Rows - 1
If Len(Mygrid.TextMatrix(i, 1)) = 0 Then Exit For
Mygrid.TextMatrix(i, 0) = i
Next

End If
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
        rs.Open "SELECT MAX(trnid )+1 AS MaxID from tblqmsmediummixhdr", MHVDB, adOpenForwardOnly, adLockOptimistic
        If rs.EOF <> True Then
        cbotrnid.Text = IIf(IsNull(rs!MaxID), 1, rs!MaxID)
        Else
        cbotrnid.Text = rs!MaxID
        End If
       Case "OPEN"
        Operation = "OPEN"
        CLEARCONTROLL
        cbotrnid.Enabled = True
        TB.Buttons(3).Enabled = True
             
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
   
   cbostaffid.Text = ""
   txtbrownsawdust.Text = ""
   txtblacksawdust.Text = ""
   txtleafmold.Text = ""
   txtmedium4.Text = ""
   txtmedium5.Text = ""
   cbostaffid.Text = ""
   
Mygrid.Clear
Mygrid.FormatString = "^Sl.No.|^Chemical Name|^Chemical g/m3|^"
Mygrid.ColWidth(0) = 960
Mygrid.ColWidth(1) = 3795
Mygrid.ColWidth(2) = 1395
Mygrid.ColWidth(3) = 1005
End Sub
Private Sub MNU_SAVE()
Dim rs As New ADODB.Recordset
On Error GoTo err
Dim i As Integer
If Len(cbotrnid.Text) = 0 Then
MsgBox "Input the Fertilizer Mix No."
Exit Sub
End If

MHVDB.BeginTrans
If Operation = "ADD" Then
LogRemarks = ""
MHVDB.Execute "INSERT INTO tblqmsmediummixhdr (trnid,startdate,enddate,brownsawdust,blacksawdust," _
            & "leafmold,medium4,medium5,staffid,status,location) " _
            & "VALUEs('" & cbotrnid.Text & "','" & Format(txtstartdate.Value, "yyyy-MM-dd") & "'," _
            & "'" & Format(txtenddate.Value, "yyyy-MM-dd") & "','" & Val(txtbrownsawdust.Text) & "', " _
            & "'" & Val(txtblacksawdust.Text) & "','" & Val(txtleafmold.Text) & "'," _
            & "'" & Val(txtmedium4.Text) & "','" & Val(txtmedium5.Text) & "'," _
            & " '" & cbostaffid.BoundText & "','ON','" & Mlocation & "')"
 
 
LogRemarks = "Inserted new record" & cbotrnid.BoundText & "," & Format(txtstartdate.Value, "yyyy-MM-dd") & "," & Mlocation & "," & txtremarks
updatemhvlog Now, MUSER, LogRemarks, ""

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblqmsmediummixhdr set " _
            & "startdate='" & Format(txtstartdate.Value, "yyyy-MM-dd") & "',enddate='" & Format(txtenddate.Value, "yyyy-MM-dd") & "', " _
            & "staffid='" & cbostaffid.BoundText & "',brownsawdust='" & Val(txtbrownsawdust.Text) & "', " _
            & "blacksawdust='" & Val(txtblacksawdust.Text) & "',leafmold='" & Val(txtleafmold.Text) & "'," _
            & "medium4='" & Val(txtmedium4.Text) & "',medium5='" & Val(txtmedium5.Text) & "'" _
            & " where trnid='" & cbotrnid.BoundText & "' and location='" & Mlocation & "'"

LogRemarks = "Updated  record" & cbotrnid.BoundText & "," & Format(txtstartdate.Value, "yyyy-MM-dd") & "," & Mlocation
updatemhvlog Now, MUSER, LogRemarks, ""
Else
MsgBox "OPERATION NOT SELECTED."
MHVDB.RollbackTrans
Exit Sub
End If

MHVDB.Execute "delete from tblqmsmediummixdetail where trnid='" & cbotrnid.BoundText & "'"

For i = 1 To Mygrid.Rows - 1
If Len(Mygrid.TextMatrix(i, 1)) = 0 Then Exit For
MHVDB.Execute "insert into tblqmsmediummixdetail (trnid,chemicalid,qty,location) values" _
            & "('" & cbotrnid.Text & "','" & Mid(Mygrid.TextMatrix(i, 1), 1, 3) & "', " _
            & " '" & Val(Mygrid.TextMatrix(i, 2)) & "','" & Mlocation & "')"

Next



MHVDB.CommitTrans
TB.Buttons(3).Enabled = False
Exit Sub

err:
MsgBox err.Description
MHVDB.RollbackTrans


End Sub



