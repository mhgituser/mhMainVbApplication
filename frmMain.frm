VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000D&
   Caption         =   "MHV"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   11280
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7605
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "USER"
            TextSave        =   "USER"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "SERVER:"
            TextSave        =   "SERVER:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "LOCATION"
            TextSave        =   "LOCATION"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin VB.Menu MNUREG 
      Caption         =   "REGISTRATION"
      HelpContextID   =   1000
      Index           =   100
      WindowList      =   -1  'True
      Begin VB.Menu MNUREGMASTER 
         Caption         =   "MASTER                 "
         HelpContextID   =   1000
         Index           =   101
         Begin VB.Menu MNUDZ 
            Caption         =   "DZONGKHAG MASTER"
            HelpContextID   =   1000
            Index           =   102
         End
         Begin VB.Menu MNUGEWOG 
            Caption         =   "GEWOG MASTER"
            HelpContextID   =   1000
            Index           =   103
         End
         Begin VB.Menu MNUTSHOWOG 
            Caption         =   "TSHOWOG MASTER"
            HelpContextID   =   1000
            Index           =   104
         End
         Begin VB.Menu mnumdep 
            Caption         =   "DEPARTMENT"
            HelpContextID   =   1000
            Index           =   105
         End
         Begin VB.Menu MNUSTAFF 
            Caption         =   "STAFF MASTER"
            HelpContextID   =   1000
            Index           =   105
         End
         Begin VB.Menu MNUFIELD 
            Caption         =   "FIELD MASTER"
            HelpContextID   =   1000
            Index           =   106
         End
         Begin VB.Menu MNUROLE 
            Caption         =   "ROLE"
            HelpContextID   =   1000
            Index           =   123
         End
      End
      Begin VB.Menu MNUREGTRN 
         Caption         =   "TRANSACTION"
         HelpContextID   =   1001
         Index           =   107
         Begin VB.Menu MNUFARMER 
            Caption         =   "FARMER REGISTRATION"
            HelpContextID   =   1001
            Index           =   108
         End
         Begin VB.Menu MNUAB 
            Caption         =   "ABSENTEE INFORMATION"
            HelpContextID   =   1001
            Index           =   109
         End
         Begin VB.Menu MNUNEWREG 
            Caption         =   "NEW LAND REGISTRATION(FARMER)"
            HelpContextID   =   1001
            Index           =   110
         End
         Begin VB.Menu MNUNEWAB 
            Caption         =   "NEW LAND REGISTRATION(ABSENTEE)"
            HelpContextID   =   1001
            Index           =   111
         End
         Begin VB.Menu mnulm 
            Caption         =   "LAND MANAGEMENT"
            HelpContextID   =   1001
            Index           =   112
         End
         Begin VB.Menu mnuppp 
            Caption         =   "PARTIAL PLANTAION PLAN"
            HelpContextID   =   1001
            Index           =   113
         End
         Begin VB.Menu MNUQ 
            Caption         =   "QUESTIONARIES"
            HelpContextID   =   1001
            Index           =   112
            Begin VB.Menu MNUABQ 
               Caption         =   "ABSENTEE"
               HelpContextID   =   1001
               Index           =   113
            End
            Begin VB.Menu MNUFRQ 
               Caption         =   "FARMER"
               HelpContextID   =   1001
               Index           =   114
            End
         End
         Begin VB.Menu MNUCONTACT 
            Caption         =   "CONTACT"
            HelpContextID   =   1000
            Index           =   124
         End
         Begin VB.Menu MNUPA 
            Caption         =   "CHALLAN ENTRY"
            HelpContextID   =   1001
            Index           =   125
         End
         Begin VB.Menu MNUUPSTATUS 
            Caption         =   "UPDATE FARMER STATUS"
            HelpContextID   =   1001
            Index           =   126
         End
         Begin VB.Menu MNUDISTPRIO 
            Caption         =   "DISTRIBUTION PRIORITY "
            HelpContextID   =   1001
            Index           =   128
         End
         Begin VB.Menu MNUDESTSCH 
            Caption         =   "DISTRIBUTION SCHEDULE"
            HelpContextID   =   1001
            Index           =   127
         End
         Begin VB.Menu mnudistsumm 
            Caption         =   "DISTRIBUTION SUMMARY"
            HelpContextID   =   1001
            Index           =   128
         End
         Begin VB.Menu MNUUDS 
            Caption         =   "UPDATE DISTRIBUTION SCHEDULE"
            HelpContextID   =   1001
            Index           =   128
         End
         Begin VB.Menu mnudfc 
            Caption         =   "DUPLICATE FARMER CHECK"
            HelpContextID   =   1001
            Index           =   129
         End
         Begin VB.Menu mnumfu 
            Caption         =   "MONITOR-FARMER UPDATE"
            HelpContextID   =   1001
            Index           =   130
         End
      End
      Begin VB.Menu MNUREGRPT 
         Caption         =   "REPORTS"
         HelpContextID   =   1002
         Index           =   115
         Begin VB.Menu MNUFL 
            Caption         =   "FARMER LISTING"
            HelpContextID   =   1002
            Index           =   116
         End
         Begin VB.Menu MNULRI 
            Caption         =   "MONITOR WISE FARMER LISTING"
            HelpContextID   =   1002
            Index           =   117
         End
         Begin VB.Menu MNULD 
            Caption         =   "LAND DETAILS"
            HelpContextID   =   1002
            Index           =   118
            Begin VB.Menu MNULDD 
               Caption         =   "LAND DETAILS (DZONGKHAG)"
               HelpContextID   =   1002
               Index           =   119
            End
            Begin VB.Menu MNULDS 
               Caption         =   "LAND DETAILS (SELECTION MODE)"
               HelpContextID   =   1002
               Index           =   120
            End
         End
         Begin VB.Menu MNUINF 
            Caption         =   "INFLUENTIAL"
            HelpContextID   =   1002
            Index           =   121
         End
         Begin VB.Menu MNUPR 
            Caption         =   "PERSON REGISTERING"
            HelpContextID   =   1002
            Index           =   122
         End
         Begin VB.Menu MNURPTCONTACT 
            Caption         =   "CONTACT DETAILS"
            HelpContextID   =   1000
            Index           =   127
         End
         Begin VB.Menu MNUODKREG 
            Caption         =   "ODK"
            HelpContextID   =   1002
            Index           =   123
         End
         Begin VB.Menu MNUBAR 
            Caption         =   "BAR CODE"
            HelpContextID   =   1002
            Index           =   124
            Begin VB.Menu MNUFBARCODE 
               Caption         =   "FARMER BAR CODE"
               HelpContextID   =   1002
               Index           =   126
            End
            Begin VB.Menu MNUSBARCODE 
               Caption         =   "STAFF BARCODE"
               HelpContextID   =   1002
               Index           =   127
            End
         End
         Begin VB.Menu mnud 
            Caption         =   "DISTRIBUTION"
            HelpContextID   =   1002
            Index           =   126
            Begin VB.Menu MNUDISTLIST 
               Caption         =   "DISTRIBUTION LIST"
               HelpContextID   =   1002
               Index           =   125
            End
            Begin VB.Menu MNUMDR 
               Caption         =   "MORE DISTRIBUTION REPORT"
               HelpContextID   =   1002
               Index           =   127
            End
         End
         Begin VB.Menu mnuemaillog 
            Caption         =   "EMAIL LOG"
            HelpContextID   =   1002
            Index           =   126
         End
      End
   End
   Begin VB.Menu mnuodk 
      Caption         =   "ODK"
      HelpContextID   =   2000
      Index           =   200
      Begin VB.Menu MNUODKMASTER 
         Caption         =   "MASTER"
         HelpContextID   =   2000
         Index           =   201
         Begin VB.Menu MNUODKUPLOAD 
            Caption         =   "UPLOAD DATA FROM ODK SERVER"
            HelpContextID   =   2000
            Index           =   202
         End
      End
      Begin VB.Menu MNUODKTRANSACTION 
         Caption         =   "TRANSACTION"
         HelpContextID   =   2001
         Index           =   203
         Begin VB.Menu MNUODKMODI 
            Caption         =   "ODK RECORD MODIFICATION"
            HelpContextID   =   2001
            Index           =   204
         End
         Begin VB.Menu MNUODKFLAG 
            Caption         =   "AlLARM PARAMETER SETTING"
            HelpContextID   =   2001
            Index           =   206
         End
         Begin VB.Menu MNUERRFOLLOW 
            Caption         =   "ODK ERROR FOLLOW UP"
            HelpContextID   =   2001
            Index           =   207
         End
      End
      Begin VB.Menu MNUODKREPORT 
         Caption         =   "REPORT"
         HelpContextID   =   2002
         Index           =   205
         Begin VB.Menu MNUDAILYACT 
            Caption         =   "DAILY ACTIVITY"
            HelpContextID   =   2002
            Index           =   206
         End
         Begin VB.Menu MNUALLFIELDS 
            Caption         =   "FIELD"
            HelpContextID   =   2002
            Index           =   207
            Begin VB.Menu MNUFDETAIL 
               Caption         =   "DETAIL"
               HelpContextID   =   2002
               Index           =   208
            End
            Begin VB.Menu MNUFSUMMARY 
               Caption         =   "SUMMARY"
               HelpContextID   =   2002
               Index           =   209
            End
            Begin VB.Menu MNUFVISIT 
               Caption         =   "VISIT"
               HelpContextID   =   2002
               Index           =   210
            End
         End
         Begin VB.Menu mnust 
            Caption         =   "STORAGE"
            Index           =   211
            Begin VB.Menu MNUSTD 
               Caption         =   "DETAIL"
               Index           =   212
            End
            Begin VB.Menu MNUSTSUM 
               Caption         =   "SUMMARY"
               Index           =   213
            End
            Begin VB.Menu MNUSTV 
               Caption         =   "VISIT"
               Index           =   214
            End
         End
         Begin VB.Menu MNUERR 
            Caption         =   "ERROR CHECK LIST"
            HelpContextID   =   2002
            Index           =   211
         End
         Begin VB.Menu mnulog 
            Caption         =   "LOG"
            HelpContextID   =   2002
            Index           =   208
         End
         Begin VB.Menu MNUGEV 
            Caption         =   "GOOGLE EARTH VIEW"
            HelpContextID   =   2002
            Index           =   209
         End
         Begin VB.Menu mnuodkdb 
            Caption         =   "ODK DASHBOARD"
            HelpContextID   =   2002
            Index           =   210
         End
      End
   End
   Begin VB.Menu MNUQMS 
      Caption         =   "QMS"
      HelpContextID   =   3000
      Index           =   300
      Begin VB.Menu MNUQMSMASTER 
         Caption         =   "MASTER"
         HelpContextID   =   3000
         Index           =   301
         Begin VB.Menu MNUFAC 
            Caption         =   "FACILITY [MS-GAP-14]"
            HelpContextID   =   3003
            Index           =   303
         End
         Begin VB.Menu MNUAI 
            Caption         =   "ACTIVE INGREDIENTS"
            HelpContextID   =   3003
            Index           =   304
         End
         Begin VB.Menu MNUCH 
            Caption         =   "CHEMICALS [MS-GAP-07]"
            HelpContextID   =   3003
            Index           =   305
         End
         Begin VB.Menu MNUPV 
            Caption         =   "PLANT VARIETY [MS-GAP-16]"
            HelpContextID   =   3003
            Index           =   306
         End
         Begin VB.Menu MNUVT 
            Caption         =   "VERIFICATION TYPE [MS-GAP-17 VT]"
            HelpContextID   =   3003
            Index           =   307
         End
         Begin VB.Menu MNUTT 
            Caption         =   "TRANSITION TYPE [MS-GAP-17 TP]"
            HelpContextID   =   3003
            Index           =   308
         End
         Begin VB.Menu mnupt 
            Caption         =   "PLANT TYPE"
            HelpContextID   =   3003
            Index           =   310
         End
         Begin VB.Menu MNUAM 
            Caption         =   "APPLICATION METHOD [MS-GAP-2a]"
            HelpContextID   =   3003
            Index           =   311
         End
      End
      Begin VB.Menu MNUQMSTRANSACTION 
         Caption         =   "TRANSACTION"
         HelpContextID   =   3001
         Index           =   302
         Begin VB.Menu mnugap 
            Caption         =   "GAP"
            HelpContextID   =   3003
            Index           =   318
            Begin VB.Menu MNUPB 
               Caption         =   "PLANT BATCH [MS-GAP-11]"
               HelpContextID   =   3003
               Index           =   309
            End
            Begin VB.Menu MNUFERMIX 
               Caption         =   "FERTILIZER MIX [MS-GAP-06]"
               HelpContextID   =   3003
               Index           =   310
            End
            Begin VB.Menu MNUMED 
               Caption         =   "MEDIUM MIX [MS-GAP-05]"
               HelpContextID   =   3003
               Index           =   315
            End
            Begin VB.Menu MNUDEADREM 
               Caption         =   "DEAD REMOVAL [MS-GAP-04 DeadRemoval]"
               HelpContextID   =   3003
               Index           =   312
            End
            Begin VB.Menu MNUSPRAY 
               Caption         =   "SPRAY"
               HelpContextID   =   3003
               Index           =   311
               Begin VB.Menu MNUIRRI 
                  Caption         =   "IRRIGATION & FERTIGATION  [MS-GAP-1]"
                  HelpContextID   =   3003
                  Index           =   313
               End
               Begin VB.Menu MNUFERT 
                  Caption         =   "INSECTICIDE & FUNJICIDE [MS-GAP-02]"
                  HelpContextID   =   3003
                  Index           =   314
               End
            End
            Begin VB.Menu MNUPMB 
               Caption         =   "PLANTING MEDIUM BATCH [MS-GAP-03]"
               HelpContextID   =   3003
               Index           =   316
            End
            Begin VB.Menu mnuptrn 
               Caption         =   "PLANT TRANSACTION [MS-GAP-04 Nursery]"
               HelpContextID   =   3003
               Index           =   317
            End
         End
         Begin VB.Menu MNUMON 
            Caption         =   "MONITOR"
            HelpContextID   =   3003
            Index           =   319
            Begin VB.Menu MNUMETER 
               Caption         =   "CLIMATE METEROLOGY [MS-MONITOR-01]"
               HelpContextID   =   3003
               Index           =   320
            End
            Begin VB.Menu MNUMET1 
               Caption         =   "CLIMATE METEROLOGY [MS-MONITOR-01a]"
               HelpContextID   =   3003
               Index           =   322
            End
            Begin VB.Menu MNUTEMP 
               Caption         =   "TEMPRATURE RECORDING [MS-MONITOR-03]"
               HelpContextID   =   3003
               Index           =   321
            End
         End
         Begin VB.Menu MNUQMSDIDT 
            Caption         =   "DISTRIBUTION"
            HelpContextID   =   3001
            Index           =   322
            Begin VB.Menu mnubac 
               Caption         =   "SENT TO FIELD"
               HelpContextID   =   3001
               Index           =   323
            End
            Begin VB.Menu MNUBTN 
               Caption         =   "BACK TO NURSERY"
               HelpContextID   =   3001
               Index           =   324
            End
            Begin VB.Menu mnustts 
               Caption         =   "SENT TO TEMPORARY STORAGE"
               HelpContextID   =   3001
               Index           =   325
            End
            Begin VB.Menu MNUCC 
               Caption         =   "CHEMICAL CHECKLIST"
               HelpContextID   =   3001
               Index           =   325
            End
         End
         Begin VB.Menu mnudnngt 
            Caption         =   "DOWNLOAD NGT DATA"
            HelpContextID   =   3001
            Index           =   325
         End
      End
      Begin VB.Menu MNUQMSREPORT 
         Caption         =   "REPORTS"
         HelpContextID   =   3002
         Index           =   303
         Begin VB.Menu mnuque 
            Caption         =   "QUERIES"
            HelpContextID   =   3003
            Index           =   304
         End
         Begin VB.Menu MNUSTF 
            Caption         =   "SENT TO FIELD HISTORY"
            HelpContextID   =   3003
            Index           =   305
         End
         Begin VB.Menu MNUSHIP 
            Caption         =   "SHIPMENT (BY BOX)"
            HelpContextID   =   3003
            Index           =   307
         End
         Begin VB.Menu MNUNTC 
            Caption         =   "SURVIVAL %"
            HelpContextID   =   3001
            Index           =   307
         End
         Begin VB.Menu mnunurdboard 
            Caption         =   "DASH BOARD"
            HelpContextID   =   3001
            Index           =   308
         End
         Begin VB.Menu mnuhlmt 
            Caption         =   "HARD LMT"
            HelpContextID   =   3001
            Index           =   309
         End
      End
   End
   Begin VB.Menu MNUSTORES 
      Caption         =   "STORES   "
      HelpContextID   =   4000
      Index           =   400
      Begin VB.Menu MNUSTOREMASTER 
         Caption         =   "MASTER"
         HelpContextID   =   4000
         Index           =   401
         Begin VB.Menu MNUCATEGORY 
            Caption         =   "ITEM CATEGORY"
            HelpContextID   =   4000
            Index           =   406
         End
         Begin VB.Menu MNUSTOREITEMGROUP 
            Caption         =   "ITEM GROUP"
            HelpContextID   =   4000
            Index           =   405
         End
         Begin VB.Menu MNUSUPPLIER 
            Caption         =   "SUPPLIER"
            HelpContextID   =   4000
            Index           =   408
         End
         Begin VB.Menu MNUSTOREITEMMASTER 
            Caption         =   "ITEM MASTER"
            HelpContextID   =   4000
            Index           =   404
         End
         Begin VB.Menu MNUDEPT 
            Caption         =   "DEPARTMENT"
            HelpContextID   =   4000
            Index           =   407
         End
      End
      Begin VB.Menu MNUSTORETRANSACTION 
         Caption         =   "TRANSACTION"
         HelpContextID   =   4001
         Index           =   402
         Begin VB.Menu MNUISSUE 
            Caption         =   "ISSUE"
            HelpContextID   =   4001
            Index           =   409
         End
         Begin VB.Menu MNUENTRY 
            Caption         =   "STOCK ENTRY"
            HelpContextID   =   4001
            Index           =   410
         End
         Begin VB.Menu MNULC 
            Caption         =   "LANDING COST"
            HelpContextID   =   4001
            Index           =   413
         End
         Begin VB.Menu MNUADJ 
            Caption         =   "STOCK ADJUSTMENT"
            HelpContextID   =   4001
            Index           =   411
         End
         Begin VB.Menu MNUPO 
            Caption         =   "PURCHASE ORDER"
            HelpContextID   =   4001
            Index           =   412
         End
      End
      Begin VB.Menu MNUSTORESREPORTS 
         Caption         =   "REPORTS"
         HelpContextID   =   4002
         Index           =   403
      End
   End
   Begin VB.Menu MNUATTEN 
      Caption         =   "ATTENDANCE"
      HelpContextID   =   7000
      Index           =   700
      Begin VB.Menu MNUATTMASTER 
         Caption         =   "MASTER"
         HelpContextID   =   7000
         Index           =   701
      End
      Begin VB.Menu MNUATTTRANSACTION 
         Caption         =   "TRANSACTION"
         HelpContextID   =   7001
         Index           =   702
      End
      Begin VB.Menu MNUATTREPORTS 
         Caption         =   "REPORTS"
         HelpContextID   =   7002
         Index           =   703
      End
   End
   Begin VB.Menu MNUMAINTAINANCE 
      Caption         =   "MAINTAINANCE"
      HelpContextID   =   5000
      Index           =   500
      Begin VB.Menu MNUU 
         Caption         =   "USER MAINTAINANCE"
         HelpContextID   =   1000
         Index           =   501
      End
      Begin VB.Menu MNUUPDATEMODULE 
         Caption         =   "UPDATE MODULE"
         HelpContextID   =   1000
         Index           =   504
      End
      Begin VB.Menu mnum 
         Caption         =   "MODULE MAINTAINANCE"
         HelpContextID   =   1000
         Index           =   502
      End
      Begin VB.Menu MNUBACKUP 
         Caption         =   "BACK UP DATABASE"
         HelpContextID   =   1000
         Index           =   503
      End
   End
   Begin VB.Menu MNUMISC 
      Caption         =   "DASHBOARD"
      HelpContextID   =   6000
      Index           =   600
      Begin VB.Menu mnuoperation 
         Caption         =   "OPERATION"
         HelpContextID   =   1001
         Index           =   601
         Begin VB.Menu mnuopnur 
            Caption         =   "NURSERY"
            HelpContextID   =   1001
            Index           =   602
         End
         Begin VB.Menu mnuopfin 
            Caption         =   "FINANCE"
            HelpContextID   =   1001
            Index           =   603
         End
         Begin VB.Menu mnuopsto 
            Caption         =   "ODK"
            HelpContextID   =   1001
            Index           =   604
         End
      End
      Begin VB.Menu mnudbmen 
         Caption         =   "MAINTAINANCE"
         HelpContextID   =   1001
         Index           =   606
         Begin VB.Menu mnudbmenfile 
            Caption         =   "D.BOARD MASTER FILE MAINTAINANCE"
            HelpContextID   =   1001
            Index           =   607
         End
      End
   End
   Begin VB.Menu mnubackmeup 
      Caption         =   "BACKUP"
      HelpContextID   =   9990
      Index           =   999
   End
   Begin VB.Menu MNUEXIT 
      Caption         =   "EXIT  "
      HelpContextID   =   9990
      Index           =   999
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MDIForm_Load()
sb.Panels(2) = UCase(MUSER)
sb.Panels(4) = UCase(Mserver)
sb.Panels(6) = UCase(Mlocation)
mbypass = False
Mcaretaker = ""
If UCase(MUSER) <> "ADMIN" Then
LoadModule UserId
End If



   
End Sub
Private Sub module(module As Stream, user As String)
    
End Sub
Private Sub MNUAB_Click(Index As Integer)
FRMABSENTEE.Show 1
End Sub
Private Sub MNUCOLLECT_Click()
FRMODKDATACOLLECT.Show 1
End Sub
Private Sub MNUABQ_Click(Index As Integer)
'FRMABSENTEEQUESTION.Show 1
'frmprac.Show 1
frmmedia.Show 1
End Sub
Private Sub MNUADJ_Click(Index As Integer)
frmsTOCKAdj.Show 1
End Sub
Private Sub MNUAI_Click(Index As Integer)
frmactiveingredients.Show 1
End Sub
Private Sub MNUALLFIELDS_Click(Index As Integer)
'FRMALLFIELDS.Show 1
End Sub
Private Sub MNUAM_Click(Index As Integer)
FRMAPPLICATIONMETHOD.Show 1
End Sub

Private Sub MNUBAC_Click(Index As Integer)
frmCrateBatchTransaction.Show 1
End Sub

Private Sub mnubackmeup_Click(Index As Integer)
MHVDB.Execute "update tblmhvsys set mbackup='Yes' where status='ON'"
MsgBox "Backed up successfully."
End Sub

Private Sub MNUBACKUP_Click(Index As Integer)
MHVDB.Execute "update tblmhvsys set mbackup='Yes' where status='ON'"
End Sub

Private Sub MNUBTN_Click(Index As Integer)
frmbacktonursery.Show 1
End Sub

Private Sub MNUCATEGORY_Click(Index As Integer)
frmCat.Show 1
End Sub

Private Sub MNUCC_Click(Index As Integer)
frmchemicalchecklistprn.Show 1
End Sub

Private Sub MNUCH_Click(Index As Integer)
frmchemical.Show 1
End Sub
Private Sub MNUCONTACT_Click(Index As Integer)
FRMCONTACT.Show 1
End Sub
Private Sub MNUDAILYACT_Click(Index As Integer)
FRMDAILYACT.Show 1
End Sub

Private Sub mnudbmenfile_Click(Index As Integer)
frmdashbordtrn.Show 1
End Sub

Private Sub MNUDEADREM_Click(Index As Integer)
frmdeadremoval.Show 1
End Sub

Private Sub MNUDEPT_Click(Index As Integer)
frmDep.Show 1
End Sub

Private Sub MNUDESTSCH_Click(Index As Integer)
frmdistributionschedule.Show 1
End Sub

Private Sub mnudfc_Click(Index As Integer)
frmduplicatecheck.Show 1
End Sub

Private Sub MNUDISTLIST_Click(Index As Integer)
FRMDISTLISTPRINT.Show 1
End Sub

Private Sub MNUDISTPRIO_Click(Index As Integer)
frmpriority.Show 1
End Sub

Private Sub mnudistsumm_Click(Index As Integer)
frmdistributionsummary.Show 1
End Sub

Private Sub mnudnngt_Click(Index As Integer)
frmdatadownload.Show 1
End Sub

Private Sub MNUDZ_Click(Index As Integer)
frmDzongkhag.Show 1
End Sub

Private Sub mnuemaillog_Click(Index As Integer)
frmemaillog.Show 1
End Sub

Private Sub MNUENTRY_Click(Index As Integer)
frmsTOCKeNTRY.Show 1
End Sub

Private Sub MNUERR_Click(Index As Integer)
FRMFIELDERROR.Show 1
End Sub

Private Sub MNUERRFOLLOW_Click(Index As Integer)
frmodkErrorfollowup.Show 1
End Sub

Private Sub mnuexit_Click(Index As Integer)
Unload Me
End Sub




Private Sub MNUFAC_Click(Index As Integer)
FRMFACILITY.Show 1
End Sub

Private Sub MNUFARMER_Click(Index As Integer)
frmfarmerreg.Show 1
End Sub

Private Sub MNUFBARCODE_Click(Index As Integer)
frmbarcode.Show 1
End Sub

Private Sub MNUFDETAIL_Click(Index As Integer)
frmfield.Show 1
End Sub

Private Sub MNUFERMIX_Click(Index As Integer)
frmfertilizermix.Show 1
End Sub

Private Sub MNUFERT_Click(Index As Integer)
FRMINSECTICIDEFUNJICIDE.Show 1
End Sub

Private Sub MNUFL_Click(Index As Integer)
FRMFARMERLISTING.Show 1
End Sub

Private Sub MNUFRQ_Click(Index As Integer)
FRMFARMERQUESTION.Show 1
End Sub

Private Sub MNUFSUMMARY_Click(Index As Integer)
FRMFIELDSUMMMARY.Show 1
End Sub

Private Sub MNUFVISIT_Click(Index As Integer)
frmfieldvisit.Show 1
End Sub

Private Sub MNUGEV_Click(Index As Integer)
frmgoogleview.Show 1
End Sub

Private Sub MNUGEWOG_Click(Index As Integer)
frmgewog.Show 1
End Sub



Private Sub MNUMODI_Click()
frmodkcollect.Show 1
End Sub

Private Sub mnuhlmt_Click(Index As Integer)
frmhardreport.Show 1
End Sub

Private Sub MNUINF_Click(Index As Integer)
FRMRPTINF.Show 1
End Sub

Private Sub MNUIRRI_Click(Index As Integer)
frmirrigationfertigation.Show 1
End Sub

Private Sub MNUISSUE_Click(Index As Integer)
frmsTOCKIssue.Show 1
End Sub

Private Sub MNULC_Click(Index As Integer)
frmsTOCKCost.Show 1
End Sub

Private Sub MNULDD_Click(Index As Integer)
FRMRPTLANDDETAILS.Show 1
End Sub

Private Sub MNULDS_Click(Index As Integer)
RptOption = "LS"
FRMLANDDETAILSEL.Show 1
End Sub

Private Sub mnulm_Click(Index As Integer)
frmmanageland.Show 1
End Sub

Private Sub mnulog_Click(Index As Integer)
frmodklog.Show 1
End Sub

Private Sub MNULRI_Click(Index As Integer)
FRMFARMERLISTINGMONITORWISE.Show 1
End Sub

Private Sub MNUM_Click(Index As Integer)
If UCase(MUSER) = UCase("ADMIN") Then
FRMMODULEMAINTAINENCE.Show 1
Else

End If
End Sub

Private Sub mnumdep_Click(Index As Integer)
frmDep.Show 1
End Sub

Private Sub MNUMDR_Click(Index As Integer)
frmdistributionreport.Show 1
End Sub

Private Sub MNUMED_Click(Index As Integer)
FRMmediummix.Show 1
End Sub

Private Sub MNUMET1_Click(Index As Integer)
FRMCLIMATEMETEROLOGY01.Show 1
End Sub

Private Sub MNUMETER_Click(Index As Integer)
frmclimatemeterology.Show 1
End Sub

Private Sub mnumfu_Click(Index As Integer)
'frmmonitorfarmerupdate.Show 1
'frmmonitorassignment.Show 1
frmextensionupdate.Show 1
End Sub

Private Sub MNUNEWREG_Click(Index As Integer)
FRMNEWLANDREG.Show 1
End Sub


Private Sub MNUNTC_Click(Index As Integer)
frmnethoopservival.Show 1

End Sub

Private Sub mnunurdboard_Click(Index As Integer)
frmdashboardnursery.Show 1
End Sub

Private Sub mnuodkdb_Click(Index As Integer)
frmodkdashboard.Show 1
End Sub

Private Sub MNUODKFLAG_Click(Index As Integer)
FRMALARMPARAMETER.Show 1
End Sub

Private Sub MNUODKMODI_Click(Index As Integer)
frmodkcollect.Show 1
End Sub

Private Sub MNUODKREG_Click(Index As Integer)
frmodkreg.Show 1
End Sub

Private Sub MNUODKUPLOAD_Click(Index As Integer)
FRMODKDATACOLLECT.Show 1
End Sub

Private Sub mnuopfil_Click(Index As Integer)
frmodkdashboard.Show
End Sub

Private Sub mnuopfin_Click(Index As Integer)
frmdashboardnursery.Show 1
End Sub

Private Sub mnuopnur_Click(Index As Integer)
'frmdashboardnursery.Show 1
frmnurserydashboard.Show 1
End Sub

Private Sub mnuopsto_Click(Index As Integer)
frmodkdashboard.Show 1
End Sub

Private Sub MNUPA_Click(Index As Integer)
frmplantedlist.Show 1
End Sub

Private Sub MNUPB_Click(Index As Integer)
FRMPLANTBATCH.Show 1
End Sub

Private Sub MNUPMB_Click(Index As Integer)
FRMmediumbatch.Show 1
End Sub

Private Sub MNUPO_Click(Index As Integer)
frmPOeNTRY.Show 1
End Sub

Private Sub mnuppp_Click(Index As Integer)
frmpartiallandmanagement.Show 1
End Sub

Private Sub MNUPR_Click(Index As Integer)
frmrptpersonregistering.Show 1
End Sub

Private Sub MNUPT_Click(Index As Integer)
FRMPLANTTYPE.Show 1
End Sub

Private Sub mnuptrn_Click(Index As Integer)
frmplanttransaction.Show 1
End Sub

Private Sub MNUPV_Click(Index As Integer)
frmplantvariety.Show 1
End Sub

Private Sub mnuque_Click(Index As Integer)
frmqueries.Show 1
End Sub

Private Sub MNUROLE_Click(Index As Integer)
FRMROLE.Show 1
End Sub

Private Sub MNURPTCONTACT_Click(Index As Integer)
RptOption = "CD"
FRMLANDDETAILSEL.Show 1
End Sub

Private Sub MNUSBARCODE_Click(Index As Integer)
frmbarcodemhvstaff.Show 1
End Sub

Private Sub MNUSHIP_Click(Index As Integer)
FRMSHIPMENTBYBOX.Show 1
End Sub

Private Sub MNUSTAFF_Click(Index As Integer)
frmmhvstaff.Show 1
End Sub

Private Sub MNUSTD_Click(Index As Integer)
frmstorage.Show 1
End Sub

Private Sub MNUSTF_Click(Index As Integer)
FRMPLANTHISTORY.Show 1
End Sub

Private Sub MNUSTOREITEMGROUP_Click(Index As Integer)
frmGrp.Show 1
End Sub

Private Sub MNUSTOREITEMMASTER_Click(Index As Integer)
frmitemMAST.Show 1
End Sub

Private Sub MNUSTSUM_Click(Index As Integer)
FRMSTORAGESUMMARY.Show 1
End Sub

Private Sub mnustts_Click(Index As Integer)
frmtempstorage.Show
End Sub

Private Sub MNUSTV_Click(Index As Integer)
FRMSTORAGEVISIT.Show 1
End Sub

Private Sub MNUSUPPLIER_Click(Index As Integer)
frmSupplMaster.Show 1
End Sub

Private Sub MNUTEMP_Click(Index As Integer)
FRMTEMPRATURERECORDING.Show 1
End Sub

Private Sub MNUTSHOWOG_Click(Index As Integer)
frmtshowog.Show 1
End Sub

Private Sub MNUTT_Click(Index As Integer)
FRMTRANSITIONTYPE.Show 1
End Sub

Private Sub MNUU_Click(Index As Integer)
frmuser.Show 1
End Sub

Private Sub MNUUDS_Click(Index As Integer)
frmupdatedeliverystatus.Show 1
End Sub

Private Sub MNUUPDATEMODULE_Click(Index As Integer)
If UCase(MUSER) = "ADMIN" Then
LoadModule "ADMIN"
End If
End Sub
Private Sub LoadModule(user As String)
On Error Resume Next
Dim tEMPmODULE As String
Dim rs As New ADODB.Recordset
Dim rschk As New ADODB.Recordset
Dim tt As String
Dim i As Integer
Dim VKEY As String
Dim Ctrl As Control
Set rs = Nothing
i = 1

   For Each Ctrl In Me.Controls
         If TypeOf Ctrl Is Menu Then
               If Ctrl.name <> "MNUEXIT" And Ctrl.name <> "mnubackmeup" Then
               Ctrl.Enabled = False
               Else
               Ctrl.Enabled = True
               End If
               
         End If
       
    Next

If user <> 100 Then
rs.Open "select * from tblmodule where userid='" & user & "' and mainmodule='" & MAINMODULEID & "' ", MHVDB

Else
rs.Open "select * from tblmodule where userid='" & user & "' ", MHVDB

End If
If rs.EOF <> True Then
 'With ctlPopMenu
'.SubClassMenu Me
'.ImageList = ilsIcons
Do While rs.EOF <> True
 For Each Ctrl In Me.Controls
 
              
        If TypeOf Ctrl Is Menu Then
        
        
        If user = 100 Then
        Set rschk = Nothing
               rschk.Open "select * from tblmodule where userid='100'  and moduleid='" & Ctrl.name & "'", MHVDB
               If rschk.EOF <> True Then
               
               Else
               MHVDB.Execute "insert into tblmodule values('" & Ctrl.name & "','" & Ctrl.Caption & "','100','1','" & Mid(Ctrl.HelpContextID, 4, 1) & "','" & i & "','" & Ctrl.Index & "','" & Mid(Ctrl.HelpContextID, 1, 3) & "')"
               End If
               
               End If
               
               
               
               
          If UCase(rs!moduleid) = UCase(Ctrl.name) Then
                If rs!userrights = 1 Then
                    Ctrl.Enabled = True
                    VKEY = Ctrl.name & "(" & rs!VKEY & ")"
                    'If Ctrl.Index - Mid(Ctrl.HelpContextID, 1, 3) <> 0 Then
                   ' .ItemIcon(VKEY) = rs!ICONID
                   ' End If
                    
                     VKEY = ""
                   
                Else
                   
                End If
              Else
              
              End If
                
              
            End If
            
              
 
      
    Next
rs.MoveNext
Loop
 'End With
Else
  If user = 100 Then

For Each Ctrl In Me.Controls
         If TypeOf Ctrl Is Menu Then
             
               Set rschk = Nothing
               rschk.Open "select * from tblmodule where userid='100' and mainmodule='" & Mid(Ctrl.HelpContextID, 1, 3) & "' and moduleid='" & Ctrl.name & "'", MHVDB
               If rschk.EOF <> True Then
               
               Else
               
               
               MHVDB.Execute "insert into tblmodule values('" & Ctrl.name & "','" & Ctrl.Caption & "','100','1','" & Mid(Ctrl.HelpContextID, 4, 1) & "','" & i & "','" & Ctrl.Index & "','" & Mid(Ctrl.HelpContextID, 1, 3) & "')"
               End If
              
         End If
       i = i + 1
    Next
End If

End If
End Sub


Private Sub MNUUPSTATUS_Click(Index As Integer)
FRMSTATUSUPDATE.Show 1
End Sub

Private Sub MNUVT_Click(Index As Integer)
FRMVERIFICATIONTYPE.Show 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key
Case "B"
updatebackup
     
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub updatebackup()
MHVDB.Execute "update tblmhvsys set mbackup='Yes' where status='ON'"
End Sub
