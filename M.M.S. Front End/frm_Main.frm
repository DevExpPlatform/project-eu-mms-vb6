VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Main 
   Caption         =   "Massive Mail System"
   ClientHeight    =   7890
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic_HCBar 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   0
      Left            =   15
      ScaleHeight     =   150
      ScaleWidth      =   1500
      TabIndex        =   60
      Top             =   450
      Width           =   1500
      Begin VB.Image img_HCBar 
         Height          =   150
         Index           =   0
         Left            =   0
         Picture         =   "frm_Main.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   150
      End
   End
   Begin TabDlg.SSTab sst_MMSTabs 
      Height          =   7845
      Left            =   0
      TabIndex        =   59
      Top             =   45
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   13838
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   688
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      TabCaption(0)   =   "Projects"
      TabPicture(0)   =   "frm_Main.frx":005F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fme_Projects"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fme_Projects 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   7275
         Left            =   75
         TabIndex        =   61
         Top             =   480
         Width           =   11715
         Begin VB.Frame Frame2 
            Caption         =   "Frame2"
            Height          =   1875
            Left            =   90
            TabIndex        =   84
            Top             =   5310
            Width           =   4425
            Begin VB.CommandButton cmd_MMSPackagesMerger 
               Caption         =   "Make Packs"
               Height          =   375
               Left            =   2895
               TabIndex        =   58
               Top             =   1410
               Width           =   1395
            End
            Begin VB.CommandButton cmd_MMSCustomerFileOrganizer 
               Caption         =   "Cust. Files Org."
               Height          =   375
               Left            =   2895
               TabIndex        =   55
               Top             =   1050
               Width           =   1395
            End
            Begin VB.CommandButton cmd_MMSGenerate 
               Caption         =   "Gen. Single"
               Height          =   375
               Index           =   2
               Left            =   2895
               TabIndex        =   92
               Top             =   690
               Width           =   1395
            End
            Begin VB.ComboBox cmb_Workings 
               Height          =   315
               Left            =   195
               Style           =   2  'Dropdown List
               TabIndex        =   91
               Top             =   270
               Width           =   4065
            End
            Begin VB.CommandButton cmd_MMSGenerate 
               Caption         =   "Gen. + Packs"
               Height          =   375
               Index           =   1
               Left            =   1515
               TabIndex        =   57
               Top             =   1410
               Width           =   1395
            End
            Begin VB.CommandButton cmd_MMSEtichette 
               Caption         =   "Etichette"
               Height          =   375
               Left            =   1515
               TabIndex        =   54
               Top             =   1050
               Width           =   1395
            End
            Begin VB.CommandButton cmd_MMSGenerate 
               Caption         =   "Gen. - Packs"
               Height          =   375
               Index           =   0
               Left            =   135
               TabIndex        =   56
               Top             =   1410
               Width           =   1395
            End
            Begin VB.CommandButton cmd_MMSSerialize 
               Caption         =   "Serializza"
               Height          =   375
               Left            =   135
               TabIndex        =   53
               Top             =   1050
               Width           =   1395
            End
            Begin VB.CommandButton cmd_MMSImportData 
               Caption         =   "Import Data"
               Height          =   375
               Left            =   1515
               TabIndex        =   52
               Top             =   690
               Width           =   1395
            End
            Begin VB.CommandButton cmd_MMSProjectLoad 
               Caption         =   "Open Project"
               Height          =   375
               Left            =   135
               TabIndex        =   51
               Top             =   690
               Width           =   1395
            End
         End
         Begin VB.PictureBox pic_PrjSettings 
            Height          =   7020
            Left            =   4590
            ScaleHeight     =   6960
            ScaleWidth      =   6975
            TabIndex        =   63
            Top             =   165
            Width           =   7035
            Begin VB.PictureBox pic_HCBar 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'None
               Height          =   135
               Index           =   1
               Left            =   15
               ScaleHeight     =   135
               ScaleWidth      =   1500
               TabIndex        =   83
               Top             =   405
               Width           =   1500
               Begin VB.Image img_HCBar 
                  Height          =   135
                  Index           =   1
                  Left            =   0
                  Picture         =   "frm_Main.frx":03AD
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   150
               End
            End
            Begin TabDlg.SSTab sst_GeneralSettings 
               Height          =   6960
               Left            =   0
               TabIndex        =   1
               Top             =   0
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   12277
               _Version        =   393216
               Style           =   1
               TabHeight       =   688
               WordWrap        =   0   'False
               ShowFocusRect   =   0   'False
               ForeColor       =   -2147483635
               TabCaption(0)   =   "Project"
               TabPicture(0)   =   "frm_Main.frx":0413
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "fme_PrjSettings"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "SubProject"
               TabPicture(1)   =   "frm_Main.frx":0647
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "fme_SubPrjSettings"
               Tab(1).ControlCount=   1
               TabCaption(2)   =   "Templates"
               TabPicture(2)   =   "frm_Main.frx":0888
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "fme_TemplatesConsolle"
               Tab(2).ControlCount=   1
               Begin VB.Frame fme_TemplatesConsolle 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   6390
                  Left            =   -74925
                  TabIndex        =   85
                  Top             =   480
                  Width           =   6810
                  Begin VB.ComboBox cmb_TemplatesRefField 
                     Height          =   315
                     Left            =   960
                     Style           =   2  'Dropdown List
                     TabIndex        =   89
                     Top             =   3015
                     Width           =   2190
                  End
                  Begin VB.TextBox txt_TemplatesRefDescr 
                     Height          =   285
                     Left            =   960
                     TabIndex        =   42
                     Top             =   2685
                     Width           =   5310
                  End
                  Begin VB.TextBox txt_TemplatesRefFValue 
                     Height          =   315
                     Left            =   4065
                     MaxLength       =   32
                     TabIndex        =   43
                     Top             =   3015
                     Width           =   2205
                  End
                  Begin VB.CommandButton cmd_TemplatesRefDEL 
                     Caption         =   "r"
                     BeginProperty Font 
                        Name            =   "Marlett"
                        Size            =   8.25
                        Charset         =   2
                        Weight          =   500
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   6360
                     TabIndex        =   45
                     Top             =   195
                     Width           =   330
                  End
                  Begin VB.CommandButton cmd_TemplatesRefAM 
                     Caption         =   "A"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   6360
                     TabIndex        =   44
                     Top             =   2970
                     Width           =   330
                  End
                  Begin VB.CommandButton cmd_TemplatesDetAM 
                     Caption         =   "A"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   6360
                     TabIndex        =   47
                     Top             =   5925
                     Width           =   330
                  End
                  Begin VB.CommandButton cmd_PrjTmplFieldUD 
                     Caption         =   "6"
                     BeginProperty Font 
                        Name            =   "Marlett"
                        Size            =   12
                        Charset         =   2
                        Weight          =   500
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   1
                     Left            =   6360
                     TabIndex        =   49
                     Top             =   3945
                     Width           =   330
                  End
                  Begin VB.CommandButton cmd_PrjTmplFieldUD 
                     Caption         =   "5"
                     BeginProperty Font 
                        Name            =   "Marlett"
                        Size            =   12
                        Charset         =   2
                        Weight          =   500
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   0
                     Left            =   6360
                     TabIndex        =   48
                     Top             =   3630
                     Width           =   330
                  End
                  Begin VB.CommandButton cmd_PrjTmplFieldDEL 
                     Caption         =   "r"
                     BeginProperty Font 
                        Name            =   "Marlett"
                        Size            =   8.25
                        Charset         =   2
                        Weight          =   500
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   6360
                     TabIndex        =   50
                     Top             =   4380
                     Width           =   330
                  End
                  Begin MSComctlLib.ListView lvw_TemplatesDetails 
                     Height          =   2715
                     Left            =   75
                     TabIndex        =   46
                     Top             =   3585
                     Width           =   6210
                     _ExtentX        =   10954
                     _ExtentY        =   4789
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     Appearance      =   1
                     NumItems        =   0
                  End
                  Begin MSComctlLib.ListView lvw_TemplatesRef 
                     Height          =   2505
                     Left            =   75
                     TabIndex        =   41
                     Top             =   150
                     Width           =   6210
                     _ExtentX        =   10954
                     _ExtentY        =   4419
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     Appearance      =   1
                     NumItems        =   0
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Filter Field:"
                     Height          =   195
                     Index           =   18
                     Left            =   165
                     TabIndex        =   90
                     Top             =   3075
                     Width           =   795
                  End
                  Begin VB.Label lbl_PrjDescr 
                     AutoSize        =   -1  'True
                     Caption         =   " Customer Org. Fields: "
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
                     Height          =   165
                     Index           =   20
                     Left            =   90
                     TabIndex        =   88
                     Top             =   3405
                     Width           =   1710
                  End
                  Begin VB.Line ln_TemplatesHStripe 
                     BorderColor     =   &H80000010&
                     Index           =   0
                     X1              =   0
                     X2              =   3975
                     Y1              =   3480
                     Y2              =   3480
                  End
                  Begin VB.Line ln_TemplatesHStripe 
                     BorderColor     =   &H8000000E&
                     Index           =   1
                     X1              =   15
                     X2              =   3990
                     Y1              =   3495
                     Y2              =   3495
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Field Value:"
                     Height          =   195
                     Index           =   19
                     Left            =   3240
                     TabIndex        =   87
                     Top             =   3075
                     Width           =   825
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Description:"
                     Height          =   195
                     Index           =   17
                     Left            =   105
                     TabIndex        =   86
                     Top             =   2730
                     Width           =   855
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   390
                     Index           =   12
                     Left            =   6330
                     Top             =   165
                     Width           =   390
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   390
                     Index           =   8
                     Left            =   6330
                     Top             =   2940
                     Width           =   390
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   390
                     Index           =   14
                     Left            =   6330
                     Top             =   5895
                     Width           =   390
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   705
                     Index           =   16
                     Left            =   6330
                     Top             =   3600
                     Width           =   390
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   390
                     Index           =   15
                     Left            =   6330
                     Top             =   4350
                     Width           =   390
                  End
               End
               Begin VB.Frame fme_SubPrjSettings 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   6390
                  Left            =   -74925
                  TabIndex        =   78
                  Top             =   480
                  Width           =   6810
                  Begin VB.TextBox txt_SubPrjBaseFName 
                     Height          =   285
                     Left            =   1350
                     MaxLength       =   48
                     TabIndex        =   27
                     Top             =   825
                     Width           =   5370
                  End
                  Begin VB.TextBox txt_SubPrjWorkDir 
                     Height          =   285
                     Left            =   1350
                     MaxLength       =   48
                     TabIndex        =   26
                     Top             =   495
                     Width           =   5370
                  End
                  Begin VB.CommandButton cmd_SubPrjCFOFieldUD 
                     Caption         =   "6"
                     BeginProperty Font 
                        Name            =   "Marlett"
                        Size            =   12
                        Charset         =   2
                        Weight          =   500
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   1
                     Left            =   6360
                     TabIndex        =   35
                     Top             =   1695
                     Width           =   330
                  End
                  Begin VB.ComboBox cmb_SubPrjCFOAliasType 
                     Height          =   315
                     Left            =   5625
                     Style           =   2  'Dropdown List
                     TabIndex        =   32
                     Top             =   4965
                     Width           =   645
                  End
                  Begin VB.TextBox txt_SubPrjCFOFieldAlias 
                     Height          =   315
                     Left            =   3405
                     MaxLength       =   32
                     TabIndex        =   31
                     Top             =   4965
                     Width           =   2175
                  End
                  Begin VB.ComboBox cmb_SubPrjCFOField 
                     Height          =   315
                     Left            =   90
                     Style           =   2  'Dropdown List
                     TabIndex        =   29
                     Top             =   4965
                     Width           =   2190
                  End
                  Begin VB.CommandButton cmd_SubPrjCFOFieldUD 
                     Caption         =   "5"
                     BeginProperty Font 
                        Name            =   "Marlett"
                        Size            =   12
                        Charset         =   2
                        Weight          =   500
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   0
                     Left            =   6360
                     TabIndex        =   34
                     Top             =   1380
                     Width           =   330
                  End
                  Begin VB.CommandButton cmd_SubPrjCFOFieldAM 
                     Caption         =   "A"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Left            =   6360
                     TabIndex        =   33
                     Top             =   4995
                     Width           =   330
                  End
                  Begin VB.CommandButton cmd_SubPrjCFOFieldDEL 
                     Caption         =   "r"
                     BeginProperty Font 
                        Name            =   "Marlett"
                        Size            =   8.25
                        Charset         =   2
                        Weight          =   500
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   6360
                     TabIndex        =   36
                     Top             =   2130
                     Width           =   330
                  End
                  Begin VB.ComboBox cmb_SubPrjCFOSortMode 
                     Height          =   315
                     Left            =   2325
                     Style           =   2  'Dropdown List
                     TabIndex        =   30
                     Top             =   4965
                     Width           =   1035
                  End
                  Begin VB.TextBox txt_SubPrjDescr 
                     Height          =   285
                     Left            =   1350
                     MaxLength       =   48
                     TabIndex        =   25
                     Top             =   165
                     Width           =   5370
                  End
                  Begin VB.CheckBox chk_SubPrjGenOMR 
                     Alignment       =   1  'Right Justify
                     Caption         =   "O.M.R."
                     Height          =   225
                     Left            =   75
                     TabIndex        =   37
                     Top             =   5445
                     Width           =   825
                  End
                  Begin VB.CheckBox chk_SubPrjMakePackages 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Make Packages"
                     Height          =   225
                     Left            =   1005
                     TabIndex        =   38
                     Top             =   5445
                     Width           =   1395
                  End
                  Begin VB.CommandButton cmd_SubPrjSettingsAM 
                     Caption         =   "Modifica"
                     Height          =   390
                     Left            =   5700
                     TabIndex        =   39
                     Top             =   5880
                     Width           =   990
                  End
                  Begin VB.CommandButton cmd_SubPrjSettingsNew 
                     Caption         =   "Nuovo"
                     Height          =   390
                     Left            =   120
                     TabIndex        =   40
                     Top             =   5880
                     Width           =   990
                  End
                  Begin MSComctlLib.ListView lvw_SubPrjCFOrganizer 
                     Height          =   3600
                     Left            =   75
                     TabIndex        =   28
                     Top             =   1335
                     Width           =   6210
                     _ExtentX        =   10954
                     _ExtentY        =   6350
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     Appearance      =   1
                     NumItems        =   0
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Base File Name:"
                     Height          =   195
                     Index           =   5
                     Left            =   210
                     TabIndex        =   82
                     Top             =   870
                     Width           =   1140
                  End
                  Begin VB.Line ln_SubPrjHStripe 
                     BorderColor     =   &H80000010&
                     Index           =   4
                     X1              =   0
                     X2              =   3975
                     Y1              =   5760
                     Y2              =   5760
                  End
                  Begin VB.Line ln_SubPrjHStripe 
                     BorderColor     =   &H8000000E&
                     Index           =   5
                     X1              =   15
                     X2              =   3990
                     Y1              =   5775
                     Y2              =   5775
                  End
                  Begin VB.Label lbl_PrjDescr 
                     AutoSize        =   -1  'True
                     Caption         =   " Customer Org. Fields: "
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
                     Height          =   165
                     Index           =   15
                     Left            =   90
                     TabIndex        =   81
                     Top             =   1140
                     Width           =   1710
                  End
                  Begin VB.Line ln_SubPrjHStripe 
                     BorderColor     =   &H8000000E&
                     Index           =   3
                     X1              =   15
                     X2              =   3990
                     Y1              =   5340
                     Y2              =   5340
                  End
                  Begin VB.Line ln_SubPrjHStripe 
                     BorderColor     =   &H80000010&
                     Index           =   2
                     X1              =   0
                     X2              =   3975
                     Y1              =   5325
                     Y2              =   5325
                  End
                  Begin VB.Line ln_SubPrjHStripe 
                     BorderColor     =   &H8000000E&
                     Index           =   1
                     X1              =   15
                     X2              =   3990
                     Y1              =   1230
                     Y2              =   1230
                  End
                  Begin VB.Line ln_SubPrjHStripe 
                     BorderColor     =   &H80000010&
                     Index           =   0
                     X1              =   0
                     X2              =   3975
                     Y1              =   1215
                     Y2              =   1215
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Sub Project Dir.:"
                     Height          =   195
                     Index           =   16
                     Left            =   165
                     TabIndex        =   80
                     Top             =   540
                     Width           =   1185
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   705
                     Index           =   5
                     Left            =   6330
                     Top             =   1350
                     Width           =   390
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   360
                     Index           =   4
                     Left            =   6330
                     Top             =   2100
                     Width           =   390
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   315
                     Index           =   3
                     Left            =   6330
                     Top             =   4965
                     Width           =   390
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Description:"
                     Height          =   195
                     Index           =   11
                     Left            =   495
                     TabIndex        =   79
                     Top             =   210
                     Width           =   855
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   450
                     Index           =   10
                     Left            =   5670
                     Top             =   5850
                     Width           =   1050
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   450
                     Index           =   9
                     Left            =   90
                     Top             =   5850
                     Width           =   1050
                  End
               End
               Begin VB.Frame fme_PrjSettings 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   6390
                  Left            =   75
                  TabIndex        =   64
                  Top             =   480
                  Width           =   6810
                  Begin VB.ComboBox cmb_PrjSortField 
                     Height          =   315
                     Left            =   90
                     Style           =   2  'Dropdown List
                     TabIndex        =   17
                     Top             =   5385
                     Width           =   5115
                  End
                  Begin VB.CommandButton cmd_PrjSortFieldAM 
                     Caption         =   "A"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Left            =   6360
                     TabIndex        =   19
                     Top             =   5415
                     Width           =   330
                  End
                  Begin VB.CommandButton cmd_PrjSortFieldDEL 
                     Caption         =   "r"
                     BeginProperty Font 
                        Name            =   "Marlett"
                        Size            =   8.25
                        Charset         =   2
                        Weight          =   500
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Left            =   6360
                     TabIndex        =   22
                     Top             =   5025
                     Width           =   330
                  End
                  Begin VB.ComboBox cmb_PrjSortFieldMode 
                     Height          =   315
                     Left            =   5250
                     Style           =   2  'Dropdown List
                     TabIndex        =   18
                     Top             =   5385
                     Width           =   1020
                  End
                  Begin VB.CommandButton cmd_PrjSettingsAM 
                     Caption         =   "Modifica"
                     Height          =   390
                     Left            =   5700
                     TabIndex        =   23
                     Top             =   5880
                     Width           =   990
                  End
                  Begin VB.CommandButton cmd_PrjSettingsNew 
                     Caption         =   "Nuovo"
                     Height          =   390
                     Left            =   120
                     TabIndex        =   24
                     Top             =   5880
                     Width           =   990
                  End
                  Begin VB.CommandButton cmd_PrjSortFieldUD 
                     Caption         =   "6"
                     BeginProperty Font 
                        Name            =   "Marlett"
                        Size            =   12
                        Charset         =   2
                        Weight          =   500
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   1
                     Left            =   6360
                     TabIndex        =   21
                     Top             =   4605
                     Width           =   330
                  End
                  Begin VB.TextBox txt_PrjDescr 
                     Height          =   285
                     Left            =   1350
                     MaxLength       =   64
                     TabIndex        =   2
                     Top             =   165
                     Width           =   5370
                  End
                  Begin VB.TextBox txt_PrjRefTable 
                     Height          =   285
                     Left            =   1350
                     MaxLength       =   32
                     TabIndex        =   6
                     Top             =   1560
                     Width           =   5370
                  End
                  Begin VB.TextBox txt_PrjJobId 
                     Height          =   285
                     Left            =   1350
                     MaxLength       =   32
                     TabIndex        =   4
                     Top             =   825
                     Width           =   5370
                  End
                  Begin VB.TextBox txt_PrjWrkDir 
                     Height          =   285
                     Left            =   1350
                     MaxLength       =   32
                     TabIndex        =   3
                     Top             =   495
                     Width           =   5370
                  End
                  Begin VB.CommandButton cmd_PrjSortFieldUD 
                     Caption         =   "5"
                     BeginProperty Font 
                        Name            =   "Marlett"
                        Size            =   12
                        Charset         =   2
                        Weight          =   500
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   0
                     Left            =   6360
                     TabIndex        =   20
                     Top             =   4290
                     Width           =   330
                  End
                  Begin VB.ComboBox cmb_PrjSerializeMode 
                     Height          =   315
                     Left            =   1350
                     Style           =   2  'Dropdown List
                     TabIndex        =   9
                     Top             =   2610
                     Width           =   2205
                  End
                  Begin VB.ComboBox cmb_PrjDataCutter 
                     Height          =   315
                     Left            =   1350
                     Style           =   2  'Dropdown List
                     TabIndex        =   7
                     Top             =   1890
                     Width           =   5370
                  End
                  Begin VB.ComboBox cmb_PrjPstlField 
                     Height          =   315
                     Left            =   4335
                     Style           =   2  'Dropdown List
                     TabIndex        =   10
                     Top             =   2610
                     Width           =   2385
                  End
                  Begin VB.CheckBox chk_PrjPstlExtraSort 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Apply Extra Sort"
                     Height          =   225
                     Left            =   60
                     TabIndex        =   11
                     Top             =   2970
                     Width           =   1485
                  End
                  Begin VB.ComboBox cmb_PrjBarCodeType 
                     Height          =   315
                     Left            =   1350
                     Style           =   2  'Dropdown List
                     TabIndex        =   12
                     Top             =   3300
                     Width           =   2205
                  End
                  Begin VB.CheckBox chk_PrjShowBarCodeTxt 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Show Text"
                     Height          =   225
                     Left            =   4005
                     TabIndex        =   13
                     Top             =   3360
                     Width           =   1065
                  End
                  Begin VB.CheckBox chk_PrjWeightTolerance 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Apply Tolerance"
                     Height          =   225
                     Left            =   3615
                     TabIndex        =   15
                     Top             =   3795
                     Width           =   1455
                  End
                  Begin VB.ComboBox cmb_PrjProduct 
                     Height          =   315
                     Left            =   1350
                     Style           =   2  'Dropdown List
                     TabIndex        =   14
                     Top             =   3735
                     Width           =   2205
                  End
                  Begin VB.ComboBox cmb_PrjNormalizer 
                     Height          =   315
                     Left            =   1350
                     Style           =   2  'Dropdown List
                     TabIndex        =   8
                     Top             =   2250
                     Width           =   5370
                  End
                  Begin VB.TextBox txt_PrjWeight 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   1350
                     MaxLength       =   32
                     TabIndex        =   5
                     Top             =   1155
                     Width           =   720
                  End
                  Begin MSComctlLib.ListView lvw_PrjSortFields 
                     Height          =   1110
                     Left            =   75
                     TabIndex        =   16
                     Top             =   4245
                     Width           =   6210
                     _ExtentX        =   10954
                     _ExtentY        =   1958
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     Appearance      =   1
                     NumItems        =   0
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   345
                     Index           =   1
                     Left            =   6330
                     Top             =   4995
                     Width           =   390
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   315
                     Index           =   2
                     Left            =   6330
                     Top             =   5385
                     Width           =   390
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   450
                     Index           =   6
                     Left            =   5670
                     Top             =   5850
                     Width           =   1050
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   450
                     Index           =   7
                     Left            =   90
                     Top             =   5850
                     Width           =   1050
                  End
                  Begin VB.Line ln_PrjHStripe 
                     BorderColor     =   &H8000000E&
                     Index           =   9
                     X1              =   15
                     X2              =   3990
                     Y1              =   5775
                     Y2              =   5775
                  End
                  Begin VB.Line ln_PrjHStripe 
                     BorderColor     =   &H80000010&
                     Index           =   8
                     X1              =   0
                     X2              =   3975
                     Y1              =   5760
                     Y2              =   5760
                  End
                  Begin VB.Label lbl_PrjDescr 
                     AutoSize        =   -1  'True
                     Caption         =   " Sort Fields: "
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
                     Height          =   165
                     Index           =   12
                     Left            =   105
                     TabIndex        =   68
                     Top             =   4065
                     Width           =   930
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Description:"
                     Height          =   195
                     Index           =   4
                     Left            =   495
                     TabIndex        =   77
                     Top             =   210
                     Width           =   855
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Job Id.:"
                     Height          =   195
                     Index           =   1
                     Left            =   780
                     TabIndex        =   76
                     Top             =   870
                     Width           =   570
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Ref. Table Name:"
                     Height          =   195
                     Index           =   6
                     Left            =   90
                     TabIndex        =   75
                     Top             =   1605
                     Width           =   1260
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "WorkDir:"
                     Height          =   195
                     Index           =   0
                     Left            =   720
                     TabIndex        =   74
                     Top             =   540
                     Width           =   630
                  End
                  Begin VB.Image img_BButton 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   705
                     Index           =   0
                     Left            =   6330
                     Top             =   4260
                     Width           =   390
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Serialize Mode:"
                     Height          =   195
                     Index           =   2
                     Left            =   270
                     TabIndex        =   73
                     Top             =   2670
                     Width           =   1080
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Data Cutter:"
                     Height          =   195
                     Index           =   3
                     Left            =   435
                     TabIndex        =   72
                     Top             =   1950
                     Width           =   915
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Pstl Field:"
                     Height          =   195
                     Index           =   7
                     Left            =   3645
                     TabIndex        =   71
                     Top             =   2670
                     Width           =   690
                  End
                  Begin VB.Line ln_PrjHStripe 
                     BorderColor     =   &H80000010&
                     Index           =   0
                     X1              =   0
                     X2              =   3975
                     Y1              =   1485
                     Y2              =   1485
                  End
                  Begin VB.Line ln_PrjHStripe 
                     BorderColor     =   &H8000000E&
                     Index           =   1
                     X1              =   15
                     X2              =   3990
                     Y1              =   1500
                     Y2              =   1500
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "BarCode Type:"
                     Height          =   195
                     Index           =   8
                     Left            =   270
                     TabIndex        =   70
                     Top             =   3360
                     Width           =   1080
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Product:"
                     Height          =   195
                     Index           =   9
                     Left            =   735
                     TabIndex        =   69
                     Top             =   3795
                     Width           =   615
                  End
                  Begin VB.Line ln_PrjHStripe 
                     BorderColor     =   &H8000000E&
                     Index           =   2
                     X1              =   15
                     X2              =   3990
                     Y1              =   3240
                     Y2              =   3240
                  End
                  Begin VB.Line ln_PrjHStripe 
                     BorderColor     =   &H80000010&
                     Index           =   3
                     X1              =   0
                     X2              =   3975
                     Y1              =   3225
                     Y2              =   3225
                  End
                  Begin VB.Line ln_PrjHStripe 
                     BorderColor     =   &H8000000E&
                     Index           =   5
                     X1              =   15
                     X2              =   3990
                     Y1              =   3675
                     Y2              =   3675
                  End
                  Begin VB.Line ln_PrjHStripe 
                     BorderColor     =   &H80000010&
                     Index           =   4
                     X1              =   0
                     X2              =   3975
                     Y1              =   3660
                     Y2              =   3660
                  End
                  Begin VB.Line ln_PrjHStripe 
                     BorderColor     =   &H8000000E&
                     Index           =   7
                     X1              =   15
                     X2              =   3990
                     Y1              =   4155
                     Y2              =   4155
                  End
                  Begin VB.Line ln_PrjHStripe 
                     BorderColor     =   &H80000010&
                     Index           =   6
                     X1              =   0
                     X2              =   3975
                     Y1              =   4140
                     Y2              =   4140
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Normalizer:"
                     Height          =   195
                     Index           =   10
                     Left            =   540
                     TabIndex        =   67
                     Top             =   2310
                     Width           =   810
                  End
                  Begin VB.Label lbl_PrjDescr 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "SubPrjs Weight:"
                     Height          =   195
                     Index           =   13
                     Left            =   195
                     TabIndex        =   66
                     Top             =   1200
                     Width           =   1155
                  End
                  Begin VB.Label lbl_PrjDescr 
                     AutoSize        =   -1  'True
                     Caption         =   "gr."
                     Height          =   195
                     Index           =   14
                     Left            =   2100
                     TabIndex        =   65
                     Top             =   1200
                     Width           =   210
                  End
               End
            End
         End
         Begin VB.Frame fme_ProjectsList 
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   5145
            Left            =   30
            TabIndex        =   62
            Top             =   165
            Width           =   4560
            Begin MSComctlLib.ImageList iml_PrjTree 
               Left            =   90
               Top             =   30
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   24
               ImageHeight     =   24
               MaskColor       =   16711935
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   2
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frm_Main.frx":0BD2
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frm_Main.frx":0E04
                     Key             =   ""
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.TreeView tvw_Projects 
               Height          =   5145
               Left            =   60
               TabIndex        =   0
               Top             =   0
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   9075
               _Version        =   393217
               Indentation     =   529
               Style           =   7
               Appearance      =   1
            End
         End
      End
   End
   Begin VB.Line ln_HStripe 
      BorderColor     =   &H8000000E&
      Index           =   0
      X1              =   0
      X2              =   3975
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line ln_HStripe 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   3975
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnu_Project 
      Caption         =   "Project"
      Begin VB.Menu mnu_PrjNew 
         Caption         =   "Nuovo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_PrjDel 
         Caption         =   "Elimina"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myMMS           As cls_MMS

Public InsertModePrj    As Boolean
Public InsertModeSubPrj As Boolean
Public SelectedNode     As String
Public SelectedPrj      As String

Private Sub cmb_PrjBarCodeType_Click()
    
    Dim BValue As Boolean
    
    With frm_Main
        BValue = (.cmb_PrjBarCodeType.ListIndex > 0)
        
        .chk_PrjShowBarCodeTxt.Enabled = BValue
        
        If BValue = False Then .chk_PrjShowBarCodeTxt.Value = 0
    End With

End Sub

Private Sub cmb_PrjDataCutter_Click()

    Dim BValue As Boolean

    With frm_Main
        .cmb_PrjPstlField.Clear
        .cmb_PrjSortField.Clear
        .cmb_SubPrjCFOField.Clear
        .cmb_TemplatesRefField.Clear
        
        .lvw_PrjSortFields.ListItems.Clear
        lvw_Autosize .lvw_PrjSortFields, lvwCONTROL
        
        .lvw_SubPrjCFOrganizer.ListItems.Clear
        lvw_Autosize .lvw_SubPrjCFOrganizer, lvwCONTROL
        
        If .cmb_PrjDataCutter.ListIndex = 0 Then
            .cmb_PrjNormalizer.ListIndex = 0
            .cmb_PrjSerializeMode.ListIndex = 1
        Else
            DB_PrjDataCutterFields_SELECTCombo cmb_GetTagValue(.cmb_PrjDataCutter, True)
        End If
            
        BValue = (.cmb_PrjDataCutter.ListIndex > 0)
        
        .cmb_PrjNormalizer.Enabled = BValue
        .cmb_PrjSerializeMode.Enabled = BValue
    End With

End Sub

Private Sub cmb_PrjSerializeMode_Click()

    Dim BValue As Boolean
    
    With frm_Main
        BValue = (.cmb_PrjSerializeMode.ListIndex = 0)
    
        .cmb_PrjPstlField.Enabled = BValue
        .chk_PrjPstlExtraSort.Enabled = BValue
    End With
    
End Sub

Private Sub cmb_TemplatesRefField_Click()

    With frm_Main
        .txt_TemplatesRefFValue.Enabled = (.cmb_TemplatesRefField.ListIndex > 0)
            
        If .txt_TemplatesRefFValue.Enabled = False Then .txt_TemplatesRefFValue.Text = ""
    End With

End Sub

Private Sub cmb_Workings_Click()

    myMMS.SetWorking = Replace$(cmb_GetTagValue(frm_Main.cmb_Workings), "'", "")
    
End Sub

'Private Sub cmb_SubPrjFilterField_Click()
'
'    With frm_Main
'        .txt_SubPrjFilterFieldValue.Enabled = (.cmb_SubPrjFilterField.ListIndex > 0)
'        .txt_SubPrjFilterFieldValue.Text = ""
'    End With
'
'End Sub

Private Sub cmd_MMSCustomerFileOrganizer_Click()
    
    If MsgBox("Sicuri di voler proseguire?", vbQuestion + vbYesNo, "Customer Files Organizer:") = vbYes Then
        If myMMS.CustomerOrganize Then MsgBox "Operazione eseguita correttamente.", vbInformation, "Customer File Organizer:"
    End If

End Sub

Private Sub cmd_MMSEtichette_Click()
    
    If SelectedNode <> "" Then
        If MsgBox("Sicuri di voler proseguire?", vbQuestion + vbYesNo, "Packages Labels Manager:") = vbYes Then
            If myMMS.MakeReports(SelectedNode) Then
                MsgBox "Operazione eseguita correttamente.", vbInformation, "Packages Labels Manager:"
            End If
        End If
    End If

End Sub

Private Sub cmd_MMSGenerate_Click(Index As Integer)
        
    Dim CanDo       As Boolean
    Dim EndPack     As String
    Dim SplitData() As String
    Dim SQLWhere    As String
    Dim StartPack   As String
    
    StartPack = ""
    EndPack = ""
    
    If MsgBox("Sicuri di voler proseguire?", vbQuestion + vbYesNo, "Generate Docs:") = vbYes Then
        Select Case Index
            Case 0, 1
                myMMS.AutoMergePacks = (Index = 1)

                If (AppSettings.PrjRenderMode = 0) Then
                    SplitData = Split(InputBox("PackStart, PackEnd", "Select Packages:"), ", ")
            
                    If chk_Array(SplitData) Then
                        StartPack = SplitData(0)
                    
                        If (UBound(SplitData) > 0) Then EndPack = SplitData(1)
                    End If
                End If
                
                CanDo = True
            
            Case 2
                SQLWhere = InputBox("Generate Single:", "SQL Clause:")
                ' SQLWhere = "NMR_ANNO = '2010' AND NMR_FATTURA = '215721'"
            
                CanDo = (Trim$(SQLWhere) <> "")
            
        End Select
        
        If CanDo Then
            Select Case AppSettings.PrjRenderMode
            Case 0
                If myMMS.MakeDocsMode00(-1, StartPack, EndPack, SQLWhere) Then MsgBox "Operazione eseguita correttamente.", vbInformation, "Generate Docs:"
            
            Case 1
                If myMMS.MakeDocsMode01 Then MsgBox "Operazione eseguita correttamente.", vbInformation, "Generate Docs:"
            
            Case 2
                If myMMS.MakeDocsMode02 Then MsgBox "Operazione eseguita correttamente.", vbInformation, "Generate Docs:"
            
            End Select
        End If
    End If

End Sub

Private Sub cmd_MMSImportData_Click()
    
    Dim OpenDlg     As New cls_CommonDialog
    Dim Tmp_Path    As String
    
    Tmp_Path = OpenDlg.Get_FileOpenName(0, "All Files (*.*)" & Chr$(0) & "*.*", "", "Load File:", False)
                
    Set OpenDlg = Nothing

    If Tmp_Path <> "" Then
        If myMMS.ImportData(Tmp_Path) Then
            If (AppSettings.PrjRenderMode = False) Then myMMS.Serialize
            
            GetPrjWorkings
            
            MsgBox "Operazione eseguita correttamente.", vbInformation, "Import Data:"
        End If
    End If
    
End Sub

Private Sub cmd_MMSPackagesMerger_Click()
    
    Dim EndPack     As String
    Dim SplitData() As String
    Dim StartPack   As String
    
    If MsgBox("Sicuri di voler proseguire?", vbQuestion + vbYesNo, "Make Packages:") = vbYes Then
        SplitData = Split(InputBox("PackStart, PackEnd", "Select Packages:"), ", ")

        If chk_Array(SplitData) Then
            StartPack = SplitData(0)
            
            If (UBound(SplitData) > 0) Then EndPack = SplitData(1)
        End If
        
        If myMMS.MakePackages(AppSettings.PrjRenderMode, StartPack, EndPack, "0") Then MsgBox "Operazione eseguita correttamente.", vbExclamation, "Merging Packages:"
    End If

End Sub

Private Sub cmd_MMSProjectLoad_Click()
    
    If myMMS.ProjectOpen(SelectedPrj) Then
        frm_Main.Caption = "Massive Mail System - " & myMMS.ProjectName
        frm_Main.cmb_Workings.Clear
        
        GetPrjWorkings
        
        'myMMS.AddExtField = "txt_DataSpedizione|" & Format$(Now, "dd/MM/yyyy")
    End If

End Sub

Private Sub cmd_MMSSerialize_Click()
    
    If MsgBox("Sicuri di voler proseguire?", vbQuestion + vbYesNo, "Esegui Serializzazione:") = vbYes Then
        If myMMS.Serialize Then MsgBox "Operazione eseguita correttamente.", vbInformation, "Serializzazione:"
    End If

End Sub

Private Sub cmd_PrjSettingsAM_Click()

    If MsgBox("Sicuro di voler proseguire?", vbQuestion + vbYesNo, IIf(InsertModePrj, "Insert", "Update") & " Progetto:") = vbNo Then Exit Sub
    
    Dim RValue As Boolean
    
    RValue = DB_PrjSettings_AM
    
    If RValue Then
        'If InsertModePrj Then GUI_SubPrjInsertCmds False
    
        MsgBox "Operazione eseguita correttamente.", vbInformation, IIf(InsertModePrj, "Insert", "Update") & " Progetto:"
    End If

End Sub

Private Sub cmd_PrjSettingsNew_Click()
    
    GUI_PrjInsertCmds (Not InsertModePrj)
    
End Sub

Private Sub cmd_PrjSortFieldAM_Click()
    
    If frm_Main.cmb_PrjSortField.ListCount > 0 Then
        ' If MsgBox("Sicuro di voler proseguire?", vbQuestion + vbYesNo, IIf(frm_Main.cmd_SubPrjSortFieldAM.Caption = "A", "Aggiungi", "Modifica") & " Sort Field:") = vbNo Then Exit Sub
        
        Dim AddMode As Boolean
        Dim I       As Byte
    
        With frm_Main
            AddMode = (.cmd_PrjSortFieldAM.Caption = "A")
            
            For I = 1 To .lvw_PrjSortFields.ListItems.Count
                If IIf(AddMode, True, (I <> .lvw_PrjSortFields.SelectedItem.Index)) And (.cmb_PrjSortField.Text = .lvw_PrjSortFields.ListItems(I).Text) Then
                    MsgBox "Non è possibile inserire o modificare un campo già presente in altre posizioni.", vbExclamation, "Attenzione:"
                
                    Exit Sub
                End If
            Next I
        
            If AddMode Then
                Dim myItem  As ListItem
                
                Set myItem = .lvw_PrjSortFields.ListItems.Add(, , .cmb_PrjSortField.Text)
                
                myItem.SubItems(1) = .cmb_PrjSortFieldMode.Text
                
                myItem.Selected = True
                myItem.EnsureVisible
                
                Set myItem = Nothing
            Else
                .lvw_PrjSortFields.SelectedItem.Text = .cmb_PrjSortField.Text
                .lvw_PrjSortFields.SelectedItem.SubItems(1) = .cmb_PrjSortFieldMode.Text
            End If
        
            lvw_Autosize .lvw_PrjSortFields, lvwITEMS
        End With
    
        GUI_PrjSortCmdsEnabler True
    End If

End Sub

Private Sub cmd_PrjSortFieldDEL_Click()
    
    With frm_Main.lvw_PrjSortFields
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
             
            If .ListItems.Count > 0 Then .SelectedItem.Selected = True
        End If
    End With

End Sub

Private Sub cmd_PrjSortFieldUD_Click(Index As Integer)
    
    lvw_SwapRows frm_Main.lvw_PrjSortFields, Index

End Sub

Private Sub cmd_SubPrjCFOFieldAM_Click()
    
    If frm_Main.cmb_SubPrjCFOField.ListCount > 0 Then
        ' If MsgBox("Sicuro di voler proseguire?", vbQuestion + vbYesNo, IIf(frm_Main.cmd_SubPrjCFOFieldAM.Caption = "A", "Aggiungi", "Modifica") & " Sort Field:") = vbNo Then Exit Sub
        
        Dim AddMode As Boolean
        Dim I       As Byte
    
        With frm_Main
            AddMode = (.cmd_SubPrjCFOFieldAM.Caption = "A")
                    
            'If .txt_SubPrjCFOFieldAlias.Text = "" Then
            '    MsgBox "Alias mancante.", vbExclamation, "Attenzione:"
          
            '    Exit Sub
            'End If
            
            For I = 1 To .lvw_SubPrjCFOrganizer.ListItems.Count
                If IIf(AddMode, True, (I <> .lvw_SubPrjCFOrganizer.SelectedItem.Index)) And (.cmb_SubPrjCFOField.Text = .lvw_SubPrjCFOrganizer.ListItems(I).Text) Then
                    MsgBox "Non è possibile inserire o modificare un campo già presente in altre posizioni.", vbExclamation, "Attenzione:"
                
                    Exit Sub
                End If
            Next I
            
            If AddMode Then
                Dim tmpLastIndex As Byte
                
                If .lvw_SubPrjCFOrganizer.ListItems.Count > 0 Then
                    For I = 1 To .lvw_SubPrjCFOrganizer.ListItems.Count
                        If .lvw_SubPrjCFOrganizer.ListItems(I).SubItems(3) = .cmb_SubPrjCFOAliasType.Text Then
                            tmpLastIndex = .lvw_SubPrjCFOrganizer.ListItems(I).Index + 1
                        End If
                    Next I
                End If
                
                If tmpLastIndex = 0 Then tmpLastIndex = 1
                
                ' Add Item
                '
                Dim myItem  As ListItem
                
                Set myItem = .lvw_SubPrjCFOrganizer.ListItems.Add(tmpLastIndex, , .cmb_SubPrjCFOField.Text)
                
                myItem.SubItems(1) = .cmb_SubPrjCFOSortMode.Text
                myItem.SubItems(2) = IIf(.txt_SubPrjCFOFieldAlias.Text = "", " ", .txt_SubPrjCFOFieldAlias.Text)
                myItem.SubItems(3) = .cmb_SubPrjCFOAliasType.Text
                
                myItem.Selected = True
                myItem.EnsureVisible
                
                Set myItem = Nothing
            Else
                .lvw_SubPrjCFOrganizer.SelectedItem.Text = .cmb_SubPrjCFOField.Text
                .lvw_SubPrjCFOrganizer.SelectedItem.SubItems(1) = .cmb_PrjSortFieldMode.Text
                .lvw_SubPrjCFOrganizer.SelectedItem.SubItems(2) = .txt_SubPrjCFOFieldAlias.Text
                .lvw_SubPrjCFOrganizer.SelectedItem.SubItems(3) = .cmb_SubPrjCFOAliasType.Text
            End If
        
            lvw_Autosize .lvw_SubPrjCFOrganizer, lvwITEMS
        End With
    
        GUI_SubPrjCFOCmdsEnabler True
    End If

End Sub

Private Sub cmd_SubPrjCFOFieldDEL_Click()
    
    With frm_Main.lvw_SubPrjCFOrganizer
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
             
            If .ListItems.Count > 0 Then .SelectedItem.Selected = True
        End If
    End With

End Sub

Private Sub cmd_SubPrjCFOFieldUD_Click(Index As Integer)
    
    lvw_SwapRows frm_Main.lvw_SubPrjCFOrganizer, Index, 3

End Sub

Private Sub cmd_SubPrjSettingsAM_Click()
   
    If MsgBox("Sicuro di voler proseguire?", vbQuestion + vbYesNo, IIf(InsertModeSubPrj, "Insert", "Update") & " SubProgetto:") = vbNo Then Exit Sub
    
    Dim RValue As Boolean
    
    RValue = DB_SubPrjSettings_AM
    
    If RValue Then
        'If InsertModeSubPrj Then GUI_SubPrjInsertCmds False
        
        MsgBox "Operazione eseguita correttamente.", vbInformation, IIf(InsertModePrj, "Insert", "Update") & " SubProgetto:"
    End If

End Sub

Private Sub cmd_SubPrjSettingsNew_Click()

    GUI_SubPrjInsertCmds (Not InsertModeSubPrj)

End Sub

Private Sub cmd_TemplatesRefAM_Click()

    With frm_Main
        If Trim$(.txt_TemplatesRefDescr.Text) = "" Then
            MsgBox "Descrizione mancante.", vbExclamation, "Attenzione:"

            Exit Sub
        End If
        
        Dim AddMode     As Boolean
        Dim SQLString   As String
            
        AddMode = (.cmd_TemplatesRefAM.Caption = "A")

        If AddMode Then
            SQLString = "INSERT INTO ref_Templates (id_SubProject, descr_Template, str_QField, str_QValue) VALUES(" & _
                        SelectedNode & ", " & _
                        Conv_String2SQLString(.txt_TemplatesRefDescr.Text) & ", " & _
                        IIf(.cmb_TemplatesRefField.ListIndex = 0, "NULL", Conv_String2SQLString(.cmb_TemplatesRefField.Text)) & ", " & _
                        Conv_String2SQLString(.txt_TemplatesRefFValue.Text) & ")"
        Else
            SQLString = "UPDATE ref_Templates SET " & _
                        "descr_Template = " & Conv_String2SQLString(.txt_TemplatesRefDescr.Text) & ", " & _
                        "str_QField = " & IIf(.cmb_TemplatesRefField.ListIndex = 0, "NULL", Conv_String2SQLString(.cmb_TemplatesRefField.Text)) & ", " & _
                        "str_QValue = " & Conv_String2SQLString(.txt_TemplatesRefFValue.Text) & _
                        " WHERE id_Template = " & .lvw_TemplatesRef.SelectedItem.Tag
        End If
        
        If DB_ExecuteQuery(SQLString, False, False, True, IIf(AddMode, "Aggiungi", "Modifica") & " riferimento Template:") Then
            If AddMode Then
                DB_TemplatesReferences_SELECListView
            Else
                .lvw_TemplatesRef.SelectedItem.Text = .txt_TemplatesRefDescr.Text
                .lvw_TemplatesRef.SelectedItem.SubItems(1) = IIf(.cmb_TemplatesRefField.ListIndex = 0, " ", .cmb_TemplatesRefField.Text)
                .lvw_TemplatesRef.SelectedItem.SubItems(2) = IIf(Trim$(.txt_TemplatesRefFValue.Text) = "", " ", .txt_TemplatesRefFValue.Text)
            End If
            
            lvw_Autosize .lvw_TemplatesRef, lvwITEMS
        
            GUI_TemplRefCmdsEnabler True
        End If
    End With

End Sub

Private Sub cmd_TemplatesRefDEL_Click()

    With frm_Main.lvw_TemplatesRef
        If .ListItems.Count > 0 Then
            If DB_ExecuteQuery("DELETE FROM ref_Templates WHERE id_Template = " & frm_Main.lvw_TemplatesRef.SelectedItem.Tag, True, False, True, "Cancella riferimento Template:") Then
                .ListItems.Remove .SelectedItem.Index
             
                If .ListItems.Count > 0 Then .SelectedItem.Selected = True
            End If
        End If
    End With

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Load()

    If App.PrevInstance Then End
    If DB_ConnectInit = False Then End

    Dim MMSConfig As cls_MMSConfig

    Set MMSConfig = New cls_MMSConfig
    Set myMMS = New cls_MMS
            
    MMSConfig.setAppPath = Fix_Paths(App.Path)
    
    If (MMSConfig.MMSConfigOpen) Then
        AppSettings.PrjRenderMode = MMSConfig.getPrjRenderMode
        
        With myMMS
            .DSN = MMSConfig.getPrjDNS
            .TNS = MMSConfig.getPrjTNS
            .BaseWorkDir = MMSConfig.getPrjFolder
            
            If .Init = False Then
                Set myMMS = Nothing
        
                End
            End If
        End With
    Else
        MsgBox MMSConfig.getErrMsg, vbCritical, "Warning:"
        
        Set MMSConfig = Nothing
    
        End
    End If
        
    ' Start Form Init
    '
    Dim I As Byte
        
    With frm_Main
        .Width = 1024 * 15
        .Height = 768 * 15
        .Left = (Screen.Width \ 2) - (.Width \ 2)
        .Top = (Screen.Height \ 2) - (.Height \ 2) - 210
    
        ' Projects Tab
        '
        .fme_ProjectsList.BackColor = &H8000000F
        
        .pic_HCBar(1).Width = .sst_GeneralSettings.Width - 45
        .img_HCBar(1).Width = .pic_HCBar(1).Width
        
        For I = 0 To 9
            .ln_PrjHStripe(I).X2 = .fme_PrjSettings.Width - 30
            .ln_PrjHStripe(I).Y2 = .ln_PrjHStripe(I).Y1
        Next I
        
        For I = 0 To 5
            .ln_SubPrjHStripe(I).X2 = .fme_SubPrjSettings.Width - 30
            .ln_SubPrjHStripe(I).Y2 = .ln_SubPrjHStripe(I).Y1
        Next I
                        
        For I = 0 To 1
            .ln_TemplatesHStripe(I).X2 = .fme_TemplatesConsolle.Width - 30
            .ln_TemplatesHStripe(I).Y2 = .ln_TemplatesHStripe(I).Y1
        Next I
        
        ' Init Components
        '
        cmb_PrjNormalizer_INIT myMMS.GetNormalizersPlugIns
        
        With .cmb_PrjSerializeMode
            .AddItem "Postalizzazione"
            .AddItem "Standard"

            .Tag = "1|0"

            .ListIndex = 1
        End With
            
        With .cmb_PrjBarCodeType
            .AddItem "Nessuno"
            .AddItem "EAN39"
            .AddItem "Interleaved 2 of 5"

            .Tag = "NULL|EAN39|ITF"
        End With
        
        '.sst_GeneralSettings.Tab = 0
    End With

    tvw_Projects_INIT
    lvw_PrjSortFields_INIT
    lvw_SubPrjCFO_INIT
    lvw_TemplatesReferences_INIT
    lvw_TemplatesDetails_INIT
        
    DB_PrjDataCutter_SELECTCombo
    DB_PrjProducts_SELECTCombo
    DB_Projects_SELECTTreeView
        

End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    With frm_Main
        .ln_HStripe(0).X2 = .Width
        .ln_HStripe(1).X2 = .ln_HStripe(0).X2
        
        .sst_MMSTabs.Move 0, 45, .Width - 120, .Height - 740
        
        .pic_HCBar(0).Width = .sst_MMSTabs.Width - 45
        .img_HCBar(0).Width = .pic_HCBar(0).Width
    
        Select Case .sst_MMSTabs.Tab
            Case 0
                .fme_Projects.Move 75, 480, .sst_MMSTabs.Width - 165, .sst_MMSTabs.Height - 570
                
                .pic_PrjSettings.Move (.fme_Projects.Width - .pic_PrjSettings.Width - 90), 165, 7035, (.fme_Projects.Height - 255)
                .sst_GeneralSettings.Height = .pic_PrjSettings.Height - 60
                
                ' Project Tab
                '
                .fme_PrjSettings.Height = .sst_GeneralSettings.Height - 570
                
                .img_BButton(6).Top = .fme_PrjSettings.Height - .img_BButton(6).Height - 90
                .img_BButton(7).Top = .img_BButton(6).Top
                .cmd_PrjSettingsNew.Top = .img_BButton(6).Top + 30
                .cmd_PrjSettingsAM.Top = .cmd_PrjSettingsNew.Top
                
                .ln_PrjHStripe(8).Y1 = .cmd_PrjSettingsNew.Top - 120
                .ln_PrjHStripe(8).Y2 = .ln_PrjHStripe(8).Y1
                .ln_PrjHStripe(9).Y1 = .ln_PrjHStripe(8).Y1 + 15
                .ln_PrjHStripe(9).Y2 = .ln_PrjHStripe(9).Y1
                
                .cmb_PrjSortField.Top = .ln_PrjHStripe(9).Y1 - .cmb_PrjSortField.Height - 75
                .cmb_PrjSortFieldMode.Top = .cmb_PrjSortField.Top
                .img_BButton(2).Top = .cmb_PrjSortField.Top
                .cmd_PrjSortFieldAM.Top = .cmb_PrjSortField.Top + 30
                .lvw_PrjSortFields.Height = .cmb_PrjSortField.Top - .lvw_PrjSortFields.Top - 30
                
                ' SubProject Tab
                '
                .fme_SubPrjSettings.Height = .fme_PrjSettings.Height
                
                .img_BButton(9).Top = .img_BButton(6).Top
                .img_BButton(10).Top = .img_BButton(9).Top
                .cmd_SubPrjSettingsNew.Top = .img_BButton(9).Top + 30
                .cmd_SubPrjSettingsAM.Top = .cmd_SubPrjSettingsNew.Top
                
                .ln_SubPrjHStripe(4).Y1 = .cmd_SubPrjSettingsNew.Top - 120
                .ln_SubPrjHStripe(4).Y2 = .ln_SubPrjHStripe(4).Y1
                .ln_SubPrjHStripe(5).Y1 = .ln_SubPrjHStripe(4).Y1 + 15
                .ln_SubPrjHStripe(5).Y2 = .ln_SubPrjHStripe(5).Y1
                
                .chk_SubPrjGenOMR.Top = .ln_SubPrjHStripe(5).Y1 - .chk_SubPrjGenOMR.Height - 105
                .chk_SubPrjMakePackages.Top = .chk_SubPrjGenOMR.Top
                
                .ln_SubPrjHStripe(2).Y1 = .chk_SubPrjGenOMR.Top - 120
                .ln_SubPrjHStripe(2).Y2 = .ln_SubPrjHStripe(2).Y1
                .ln_SubPrjHStripe(3).Y1 = .ln_SubPrjHStripe(2).Y1 + 15
                .ln_SubPrjHStripe(3).Y2 = .ln_SubPrjHStripe(3).Y1
                
                .cmb_SubPrjCFOField.Top = .ln_SubPrjHStripe(2).Y1 - .cmb_SubPrjCFOField.Height - 45
                .cmb_SubPrjCFOSortMode.Top = .cmb_SubPrjCFOField.Top
                .txt_SubPrjCFOFieldAlias.Top = .cmb_SubPrjCFOField.Top
                .cmb_SubPrjCFOAliasType.Top = .cmb_SubPrjCFOField.Top
                .img_BButton(3).Top = .cmb_SubPrjCFOField.Top
                .cmd_SubPrjCFOFieldAM.Top = .cmb_SubPrjCFOField.Top + 30
                .lvw_SubPrjCFOrganizer.Height = .cmb_SubPrjCFOField.Top - .lvw_SubPrjCFOrganizer.Top - 30
                
                ' Templates Tab
                '
                .fme_TemplatesConsolle.Height = .fme_PrjSettings.Height
                
                .lvw_TemplatesDetails.Height = .fme_TemplatesConsolle.Height - .lvw_TemplatesDetails.Top - 75
                .img_BButton(14).Top = .fme_TemplatesConsolle.Height - .img_BButton(14).Height - 90
                .cmd_TemplatesDetAM.Top = .img_BButton(14).Top + 30
                
                '
                .Frame2.Top = .sst_GeneralSettings.Height - .Frame2.Height + 225
                
                .fme_ProjectsList.Width = .pic_PrjSettings.Left - .fme_ProjectsList.Left
                .fme_ProjectsList.Height = .Frame2.Top - .fme_PrjSettings.Top + 315
                .tvw_Projects.Width = .fme_ProjectsList.Width - 120
                .tvw_Projects.Height = .fme_ProjectsList.Height
        
        End Select
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    DB_ConnectRelease

    Set myMMS = Nothing

End Sub

Private Sub GetPrjWorkings()
        
    Dim PrjWorkings() As String

    PrjWorkings = myMMS.GetWorkings
    
    If chk_Array(PrjWorkings) Then
        If PrjWorkings(0) <> "Error" Then
            Dim I       As Integer
            Dim tmpStr  As String
            
            With frm_Main.cmb_Workings
                .Clear
                .Tag = ""
                
                For I = 0 To UBound(PrjWorkings)
                    tmpStr = Mid$(PrjWorkings(I), 7, 2) & "/" & Mid$(PrjWorkings(I), 5, 2) & "/" & Mid$(PrjWorkings(I), 1, 4) & " - " & Mid$(PrjWorkings(I), 9, 2) & "." & Mid$(PrjWorkings(I), 11, 2) & "." & Mid$(PrjWorkings(I), 13, 2)
                    
                    .AddItem (I + 1) & ". " & tmpStr
                    .Tag = .Tag & PrjWorkings(I) & "|"
                Next I
            
                .ListIndex = .ListCount - 1
            End With
        End If
    End If
    
    Erase PrjWorkings

End Sub

Private Sub lvw_PrjSortFields_DblClick()

    GUI_PrjSortCmdsEnabler False

End Sub

Private Sub lvw_SubPrjCFOrganizer_DblClick()
    
    GUI_SubPrjCFOCmdsEnabler False

End Sub

Private Sub lvw_TemplatesRef_DblClick()
    
    GUI_TemplRefCmdsEnabler False
    
End Sub

Private Sub lvw_TemplatesRef_ItemClick(ByVal Item As MSComctlLib.ListItem)

    DB_TemplatesDetails_SELECListView

End Sub

Private Sub sst_GeneralSettings_Click(PreviousTab As Integer)
    
    With frm_Main
        If .Visible Then
            Select Case .sst_GeneralSettings.Tab
                Case 0
                    .txt_PrjDescr.SetFocus
                    
                Case 1
                    .txt_SubPrjDescr.SetFocus
                    
                Case 2
                    .lvw_TemplatesRef.SetFocus
                    
            End Select
        End If
    End With

End Sub

Private Sub tvw_Projects_NodeClick(ByVal Node As MSComctlLib.Node)

    Static tmp_SelPrj       As String
    Static tmp_SelSubPrj    As String
    
    If Left$(Node.Key, 1) = "S" Then
        SelectedNode = Right$(Node.Key, Len(Node.Key) - 1)
        SelectedPrj = Right$(Node.Parent.Key, Len(Node.Parent.Key) - 1)

        If tmp_SelPrj <> SelectedPrj Then
            DB_PrjINFO_SELECT
            
            tmp_SelPrj = SelectedPrj
        End If
        
        If tmp_SelSubPrj <> SelectedNode Then
            DB_SubPrjINFO_SELECT
            
            tmp_SelSubPrj = SelectedNode
        End If
    Else
        SelectedNode = ""
        SelectedPrj = Right$(Node.Key, Len(Node.Key) - 1)
        
        If tmp_SelPrj <> SelectedPrj Then
            DB_PrjINFO_SELECT
            
            tmp_SelPrj = SelectedPrj
        End If
    End If
        
    GUI_SubPrjInsertCmds (SelectedNode = "")

End Sub

Private Sub txt_PrjWeight_KeyPress(KeyAscii As Integer)
    
    NumOnlyFilter KeyAscii, frm_Main.txt_PrjWeight.Text

End Sub
