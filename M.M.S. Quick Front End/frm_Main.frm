VERSION 5.00
Begin VB.Form frm_Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "M.M.S. Quick Front End"
   ClientHeight    =   2655
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fme_CommandsConsolle 
      Caption         =   "3. Consolle:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   825
      Left            =   45
      TabIndex        =   9
      Top             =   1785
      Width           =   6630
      Begin VB.CommandButton cmd_ExecuteProcess 
         Caption         =   "Make Reports"
         Height          =   420
         Index           =   2
         Left            =   3795
         TabIndex        =   10
         Top             =   255
         Width           =   1290
      End
      Begin VB.CommandButton cmd_ExecuteProcess 
         Caption         =   "Make Supports"
         Height          =   420
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   255
         Width           =   1290
      End
      Begin VB.CommandButton cmd_ExecuteProcess 
         Caption         =   "Generate Docs"
         Height          =   420
         Index           =   1
         Left            =   5190
         TabIndex        =   4
         Top             =   255
         Width           =   1290
      End
      Begin VB.Image img_BButton 
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Index           =   1
         Left            =   3765
         Top             =   225
         Width           =   1350
      End
      Begin VB.Image img_BButton 
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Index           =   3
         Left            =   105
         Top             =   225
         Width           =   1350
      End
      Begin VB.Image img_BButton 
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Index           =   2
         Left            =   5160
         Top             =   225
         Width           =   1350
      End
   End
   Begin VB.Frame fme_AvailableImports 
      Caption         =   "2. Available Imports:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   825
      Left            =   45
      TabIndex        =   5
      Top             =   930
      Width           =   6630
      Begin VB.CommandButton cmd_PrjAddImport 
         Caption         =   "Add Import"
         Height          =   405
         Left            =   5190
         TabIndex        =   2
         Top             =   255
         Width           =   1290
      End
      Begin VB.ComboBox cmb_Workings 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   285
         Width           =   3825
      End
      Begin VB.Image img_BButton 
         BorderStyle     =   1  'Fixed Single
         Height          =   465
         Index           =   0
         Left            =   5160
         Top             =   225
         Width           =   1350
      End
      Begin VB.Label lbl_Descr 
         Alignment       =   1  'Right Justify
         Caption         =   "Select Working:"
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   8
         Top             =   330
         Width           =   1125
      End
   End
   Begin VB.Frame fme_AvailableProjects 
      Caption         =   "1. Available Projects:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   825
      Left            =   45
      TabIndex        =   6
      Top             =   75
      Width           =   6630
      Begin VB.ComboBox cmb_Projects 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   285
         Width           =   5220
      End
      Begin VB.Label lbl_Descr 
         Alignment       =   1  'Right Justify
         Caption         =   "Select Project:"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   330
         Width           =   1050
      End
   End
   Begin VB.Image img_TopLine 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   -15
      Top             =   0
      Width           =   6750
   End
   Begin VB.Menu mnu_Prj 
      Caption         =   "Project"
      Begin VB.Menu mnu_PrjFilter 
         Caption         =   "Projects Filter"
         Shortcut        =   ^F
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_BaseWorkDir 
         Caption         =   "Base Work Dir."
         Shortcut        =   ^B
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Space00 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_SetDNS 
         Caption         =   "Set DNS Connection"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_SetTNS 
         Caption         =   "Set TNS Connection"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Space01 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Prints 
         Caption         =   "Prints"
         Visible         =   0   'False
         Begin VB.Menu mnu_BillingPrvw 
            Caption         =   "Billing - Preview"
         End
         Begin VB.Menu mnu_BillingPrnt 
            Caption         =   "Billing - Print"
         End
      End
      Begin VB.Menu mnu_Space02 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_SetDLLIDPExtraParams 
         Caption         =   "Set IDP Extra Parameters"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnu_SetDLLODPExtraParams 
         Caption         =   "Set ODP Extra Parameters"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_Space03 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Quit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedPrj As Long

Private Sub cmb_Projects_Click()

    SelectedPrj = cmb_GetTagValue(frm_Main.cmb_Projects, True)

    If myMMS.ProjectOpen(SelectedPrj) Then
        frm_Main.Caption = "MMS_QFE - " & myMMS.ProjectName
        
        GetPrjWorkings
    Else
        SelectedPrj = -1
    End If

End Sub

Private Sub cmb_Workings_Click()
    
    myMMS.SetWorking = cmb_GetTagValue(frm_Main.cmb_Workings)

End Sub

Private Sub cmd_ExecuteProcess_Click(Index As Integer)

    If AppSettings.PrjFilter = "" Or AppSettings.PrjFolder = "" Then
        MsgBox "Impossibile proseguire.", vbCritical, "Attenzione:"
    Else
        If MsgBox(LBL_MSGBOX02, vbQuestion + vbYesNo, LBL_MSGTTL02) = vbYes Then
            frm_Main.Enabled = False
            
            Select Case Index
                Case 0
                    If myMMS.CustomerOrganize() Then MsgBox LBL_MSGBOX01, vbInformation, "Customer File Organizer:"
            
                Case 1
                    Select Case AppSettings.PrjRenderMode
                    Case 0
                        If myMMS.MakeDocsMode00(-1) Then
                            MsgBox LBL_MSGBOX01, vbInformation, "Generate Docs:"
                        Else
                            MsgBox LBL_MSGBOX03, vbExclamation, "Generate Docs:"
                        End If
                    
                    Case 1
                        If myMMS.MakeDocsMode01 Then
                            MsgBox LBL_MSGBOX01, vbInformation, "Generate Docs:"
                        Else
                            MsgBox LBL_MSGBOX03, vbExclamation, "Generate Docs:"
                        End If
                    
                    Case 2
                        If myMMS.MakeDocsMode02 Then
                            MsgBox LBL_MSGBOX01, vbInformation, "Generate Docs:"
                        Else
                            MsgBox LBL_MSGBOX03, vbExclamation, "Generate Docs:"
                        End If
                    
                    End Select
                        
                Case 2
                    If myMMS.MakeReports(DB_SubProject_SELECT(SelectedPrj)) Then
                        MsgBox LBL_MSGBOX01, vbInformation, "Generate Reports:"
                    Else
                        MsgBox LBL_MSGBOX03, vbExclamation, "Generate Reports:"
                    End If

            End Select
            
            frm_Main.Enabled = True
        End If
    End If

End Sub

Private Sub cmd_PrjAddImport_Click()
    
    If AppSettings.PrjFilter = "" Then
        MsgBox "Impossibile proseguire.", vbCritical, "Attenzione:"
    Else
        Dim OpenDlg     As New cls_CommonDialog
        Dim Tmp_Path    As String
        
        Tmp_Path = OpenDlg.Get_FileOpenName(frm_Main.hwnd, "All Files (*.*)" & Chr$(0) & "*.*", "", "Load File:", False)
                    
        Set OpenDlg = Nothing
    
        If Tmp_Path <> "" Then
            If myMMS.ImportData(Tmp_Path) Then
                If (AppSettings.PrjRenderMode = False) Then myMMS.Serialize

                GetPrjWorkings
                
                MsgBox LBL_MSGBOX01, vbInformation, LBL_MSGTTL01
            End If
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (Shift = 7) Then
        Select Case KeyCode
            Case vbKeyA
                GUI_MenuAdminEnabler True
            
            Case vbKeyZ
                GUI_MenuAdminEnabler False

        End Select
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Load()

    If App.PrevInstance Then End
    
    Dim MMSConfig As cls_MMSConfig

    Set MMSConfig = New cls_MMSConfig
    
    AppPath = Fix_Paths(App.Path)
    
    MMSConfig.setAppPath = AppPath
    
    If (MMSConfig.MMSConfigOpen) Then
        AppSettings.PrjDNS = MMSConfig.getPrjDNS                ' GetSetting("MMS_QFE", "Settings", "PrjDNS", "")
        AppSettings.PrjFolder = MMSConfig.getPrjFolder          ' GetSetting("MMS_QFE", "Settings", "PrjFolder", "")
        AppSettings.PrjFilter = MMSConfig.getPrjFilter          ' GetSetting("MMS_QFE", "Settings", "PrjFilter", "")
        AppSettings.PrjLocale = MMSConfig.getPrjLocale          ' GetSetting("MMS_QFE", "Settings", "PrjLocale", "")
        AppSettings.PrjRenderMode = MMSConfig.getPrjRenderMode
        AppSettings.PrjTNS = MMSConfig.getPrjTNS                ' GetSetting("MMS_QFE", "Settings", "PrjTNS", "")

        If DB_ConnectInit = False Then End
        
        If (AppSettings.PrjDNS = "") Then mnu_SetDNS_Click
        If (AppSettings.PrjTNS = "") Then mnu_SetTNS_Click
        
        Set myMMS = New cls_MMS
                
        With myMMS
            .AutoMergePacks = True
            .BaseWorkDir = AppSettings.PrjFolder
            .DSN = AppSettings.PrjDNS
            .TNS = AppSettings.PrjTNS
        
            If .Init = False Then
                Set myMMS = Nothing
        
                End
            End If
        End With
        
        Set MMSConfig = Nothing
        
        ' Locale Init
        '
        LBL_IBXPARAMSTITLE = "Parametri:"
        LBL_MSGBOX01 = "Operazione eseguita con successo"
        LBL_MSGBOX02 = "Sicuri di voler proseguire?"
        LBL_MSGBOX03 = "Impossibile eseguire l'operazione"
        LBL_MSGTTL01 = "Import Data:"
        LBL_MSGTTL02 = "Execute Command:"
        
        If (AppSettings.PrjLocale <> "") Then GUI_GetLocalization
        DB_Projects_SELECT
    Else
        MsgBox MMSConfig.getErrMsg, vbCritical, "Warning:"
        
        Set MMSConfig = Nothing
        
        End
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set myMMS = Nothing

End Sub

Private Sub mnu_BaseWorkDir_Click()
    
    Dim OpenDlg     As New cls_CommonDialog
    Dim Tmp_Path    As String
    
    Tmp_Path = OpenDlg.BrowseForFolder(frm_Main.hwnd, "Select Base Work Dir.:")
                
    Set OpenDlg = Nothing

    If Tmp_Path <> "" Then
        AppSettings.PrjFolder = Fix_Paths(Tmp_Path)
        
        SaveSetting "MMS_QFE", "Settings", "PrjFolder", AppSettings.PrjFolder
        
        myMMS.BaseWorkDir = AppSettings.PrjFolder
    End If

End Sub

Private Sub mnu_BillingPrnt_Click()
    
    Set mySDF = New cls_SDFP

    mySDF.SetMode MPrint
           
    If mySDF.PrinterSetUp(frm_Main.hwnd) Then DB_PRNTPRVWReportGen 1
     
    Set mySDF = Nothing

End Sub

Private Sub mnu_BillingPrvw_Click()
     
    DB_PRNTPRVWReportGen 0

End Sub

Private Sub mnu_PrjFilter_Click()

    Dim PrjFilter As String

    PrjFilter = InputBox("Insert Projects Availability:", "Projects Filter:", AppSettings.PrjFilter)

    If PrjFilter <> "" Then
        SaveSetting "MMS_QFE", "Settings", "PrjFilter", PrjFilter
        
        AppSettings.PrjFilter = PrjFilter
        
        DB_Projects_SELECT
    End If

End Sub

Private Sub mnu_Quit_Click()

    Unload Me

End Sub

Private Sub mnu_SetDLLIDPExtraParams_Click()

    Dim DLLExtraParams As String

    DLLExtraParams = InputBox("IDP DLL Extra " & IIf(LBL_IBXPARAMSTITLE = "", "Parameters:", LBL_IBXPARAMSTITLE), IIf(LBL_IBXPARAMSTITLE = "", "Parameters:", LBL_IBXPARAMSTITLE), myMMS.GetIDPPlugInParams)
    
    If ((DLLExtraParams <> "") And (DLLExtraParams <> myMMS.GetIDPPlugInParams)) Then
        If (DLLExtraParams = "NULL") Then DLLExtraParams = ""
        
        myMMS.SetIDPPlugInParams = DLLExtraParams
    End If
    
End Sub

Private Sub mnu_SetDLLODPExtraParams_Click()
    
    Dim DLLExtraParams As String

    DLLExtraParams = InputBox("ODP DLL Extra Parameters:", "Parameters:", myMMS.GetODPPlugInParams)
    
    If ((DLLExtraParams <> "") And (DLLExtraParams <> myMMS.GetODPPlugInParams)) Then
        If (DLLExtraParams = "NULL") Then DLLExtraParams = ""
        
        myMMS.SetODPPlugInParams = DLLExtraParams
    End If

End Sub

Private Sub mnu_SetDNS_Click()

    Dim PrjDNS As String

    PrjDNS = InputBox("Insert DNS Connection:", "DNS:", AppSettings.PrjDNS)

    If PrjDNS <> "" Then
        SaveSetting "MMS_QFE", "Settings", "PrjDNS", PrjDNS
        
        AppSettings.PrjDNS = PrjDNS
    End If

End Sub

Private Sub mnu_SetTNS_Click()

    Dim PrjTNS As String

    PrjTNS = InputBox("Insert TNS Connection:", "TNS:", AppSettings.PrjTNS)

    If PrjTNS <> "" Then
        SaveSetting "MMS_QFE", "Settings", "PrjTNS", PrjTNS
        
        AppSettings.PrjTNS = PrjTNS
    End If

End Sub

