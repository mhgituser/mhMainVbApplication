VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMSTATUSUPDATE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FARMER STATUS UPDATE"
   ClientHeight    =   3630
   ClientLeft      =   7665
   ClientTop       =   1335
   ClientWidth     =   6315
   Icon            =   "FRMSTATUSUPDATE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6315
   Begin VB.Frame Frame1 
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
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   6135
      Begin MSDataListLib.DataCombo cbostatus 
         Bindings        =   "FRMSTATUSUPDATE.frx":0E42
         DataField       =   "desc"
         Height          =   360
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   ""
         BoundColumn     =   "statusid"
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
      Begin MSDataListLib.DataCombo cbomonitor 
         Bindings        =   "FRMSTATUSUPDATE.frx":0E57
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Top             =   840
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CONFIRMED BY"
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
         TabIndex        =   9
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "STATUS UPDATE TO"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
      Begin MSDataListLib.DataCombo cbofarmerid 
         Bindings        =   "FRMSTATUSUPDATE.frx":0E6C
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   2040
         TabIndex        =   1
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FARMER ID"
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
         Width           =   1035
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   6360
      Top             =   840
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
            Picture         =   "FRMSTATUSUPDATE.frx":0E81
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMSTATUSUPDATE.frx":121B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMSTATUSUPDATE.frx":15B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMSTATUSUPDATE.frx":228F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMSTATUSUPDATE.frx":26E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMSTATUSUPDATE.frx":2E9B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   1164
      ButtonWidth     =   1217
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OPEN"
            Key             =   "OPEN"
            Object.ToolTipText     =   "OPEN/EDIT EXISTING RECORD"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "SAVE"
            Key             =   "SAVE"
            Object.ToolTipText     =   "SAVES RECORD"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "DELETE"
            Key             =   "DELETE"
            Object.ToolTipText     =   "DELETE THE RECORD"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EXIT"
            Key             =   "EXIT"
            Object.ToolTipText     =   "EXIT FROM THE FORM"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   3
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   75
   End
End
Attribute VB_Name = "FRMSTATUSUPDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsfr As New ADODB.Recordset
Dim rsst As New ADODB.Recordset

Private Sub cbofarmerid_LostFocus()
cbofarmerid.Enabled = False



showstatus



End Sub
Private Sub showstatus()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblfarmer where idfarmer='" & cbofarmerid.BoundText & "'", MHVDB
If rs.EOF <> True Then
If Len(rs!monitor) = 5 Then
FindsTAFF rs!monitor
cbomonitor.Text = rs!monitor & "  " & sTAFF
End If




If rs!status = "A" Then
Label2.Caption = "ACTIVE"
Label2.ForeColor = vbBlue
ElseIf rs!status = "D" Then
Label2.Caption = "DROPPED OUT"
Label2.ForeColor = vbRed
ElseIf rs!status = "R" Then
Label2.Caption = "REJECTED"
Label2.ForeColor = vbRed
Else
Label2.Caption = "ERROR"
TB.buttons(3).Enabled = False
Label2.ForeColor = vbRed
End If

End If

End Sub


Private Sub Form_Load()
On Error GoTo err
Operation = ""


Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

Set rsfr = Nothing

If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"



If rsst.State = adStateOpen Then rsst.Close
rsst.Open "SELECT  `statusid` ,  `desc` FROM  `tblstatus`", db
Set cbostatus.RowSource = rsst
cbostatus.ListField = "desc"
cbostatus.BoundColumn = "statusid"



Set rsst = Nothing

If rsst.State = adStateOpen Then rsst.Close
rsst.Open "SELECT concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE from mhv.tblmhvstaff where dept='106' and status not in('D','R','C','T') ORDER BY STAFFCODE", db
Set cbomonitor.RowSource = rsst
cbomonitor.ListField = "STAFFNAME"
cbomonitor.BoundColumn = "STAFFCODE"



Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "ADD"
       
        
Case "OPEN"
         Operation = "OPEN"
         
         cbofarmerid.Enabled = True
        cbofarmerid.Text = ""
         TB.buttons(2).Enabled = True
         cbomonitor.Text = ""
       
       Case "SAVE"
      
       MNU_SAVE
        
       Case "DELETE"
         Case "PRINT"
         
       Case "EXIT"
       Unload Me
       
       
End Select
End Sub
Private Sub MNU_SAVE()
LogRemarks = ""
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim rmrk As String

rmrk = ""


If Len(cbomonitor.Text) = 0 Then
MsgBox "Select the confirming monitor.", , "MHV ERROR BOX"
Exit Sub
End If



If Len(cbofarmerid.Text) = 0 Then
MsgBox "Please Select The Appropriate Information From The Drop Down Controll.", , "MHV ERROR BOX"
Exit Sub
End If
If Len(cbostatus.Text) = 0 Then
MsgBox "Please Select The Appropriate Status From The Drop Down Controll.", , "MHV ERROR BOX"
Exit Sub
End If


MHVDB.BeginTrans
If Operation = "ADD" Then



ElseIf Operation = "OPEN" Then


Set rs = Nothing
rs.Open "select * from tblplanted where farmercode='" & cbofarmerid.BoundText & "'", MHVDB
If rs.EOF <> True Then

 MsgBox ("This farmer exists in planted list so you cannot proceed to save. " & ProperCase(cbostatus.Text))
Exit Sub

End If




rmrk = cbostatus.Text & " confirmed by monitor " & cbomonitor.Text
LogRemarks = "Stasus of farmer " & cbofarmerid.Text & " is modified to " & cbostatus.Text
MHVDB.Execute "update tblfarmer set status='" & cbostatus.BoundText & "',monitor='',remarks='" & rmrk & "' where idfarmer='" & cbofarmerid.BoundText & "'"
MHVDB.Execute "update tbllandreg set status='" & cbostatus.BoundText & "' where farmerid='" & cbofarmerid.BoundText & "'"


Set rs = Nothing
rs.Open "select * from tblplanted where farmercode='" & cbofarmerid.BoundText & "'", MHVDB
If rs.EOF <> True Then

If MsgBox("This farmer exists in planted list,Do you want to continue to " & ProperCase(cbostatus.Text), vbYesNo) = vbYes Then

MHVDB.Execute "update tblplanted set status='C' where farmercode='" & cbofarmerid.BoundText & "'"

MHVDB.Execute "update tblfarmer set status='" & cbostatus.BoundText & "',monitor='' where idfarmer='" & cbofarmerid.BoundText & "'"
MHVDB.Execute "update tbllandreg set status='" & cbostatus.BoundText & "' where farmerid='" & cbofarmerid.BoundText & "'"
Else
MHVDB.Execute "update tblfarmer set status='A' where idfarmer='" & cbofarmerid.BoundText & "'"
MHVDB.Execute "update tbllandreg set status='A' where farmerid='" & cbofarmerid.BoundText & "'"


End If


End If




updatemhvlog Now, MUSER, LogRemarks, ""
Else
MsgBox "OPERATION NOT SELECTED."
End If
MHVDB.CommitTrans
showstatus
TB.buttons(2).Enabled = False
Exit Sub
err:
MsgBox err.Description
MHVDB.RollbackTrans
End Sub
