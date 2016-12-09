VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptpersonregsummary 
   Caption         =   "REPORT"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   Icon            =   "frmrptpersonregsummary.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35719
   _ExtentY        =   19315
   SectionData     =   "frmrptpersonregsummary.dsx":0E42
End
Attribute VB_Name = "rptpersonregsummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SLNO As Integer

Private Sub ActiveReport_ReportStart()
SLNO = 1
End Sub

Private Sub Detail_Format()
TXTSLNO.Text = SLNO
SLNO = SLNO + 1
FindFA txtFARMERID.Text, "F"
TXTFARMERNAME.Text = FAName
End Sub

Private Sub GroupHeader1_Format()
FindsTAFF txtMONITOR.Text
TXTSTAFFNAME.Text = sTAFF
End Sub
