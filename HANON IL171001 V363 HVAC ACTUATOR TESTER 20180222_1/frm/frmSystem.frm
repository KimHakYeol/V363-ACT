VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{F36F6844-D389-11D1-8968-006097AA579E}#1.0#0"; "HiResTimer.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmSystem 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SYSTEM"
   ClientHeight    =   15000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   19110
   BeginProperty Font 
      Name            =   "³ª´®°íµñ"
      Size            =   12
      Charset         =   129
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15000
   ScaleWidth      =   19110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   13455
      Left            =   180
      TabIndex        =   1
      Top             =   1320
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   23733
      _Version        =   393216
      Tab             =   1
      TabHeight       =   1323
      TabCaption(0)   =   "GENERAL"
      TabPicture(0)   =   "frmSystem.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "fraPassword"
      Tab(0).Control(2)=   "fraLang"
      Tab(0).Control(3)=   "fraHidden"
      Tab(0).Control(4)=   "fraGeneral(0)"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "OFFSET"
      TabPicture(1)   =   "frmSystem.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraSystem(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraCorrelation"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "MASTER SAMPLE"
      TabPicture(2)   =   "frmSystem.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCal(3)"
      Tab(2).Control(1)=   "fraCal(4)"
      Tab(2).Control(2)=   "fraCal(0)"
      Tab(2).Control(3)=   "fraCal(1)"
      Tab(2).ControlCount=   4
      Begin VB.Frame fraCal 
         Caption         =   "LEAK"
         Height          =   2175
         Index           =   1
         Left            =   -74460
         TabIndex        =   105
         Top             =   8880
         Visible         =   0   'False
         Width           =   4275
         Begin VB.TextBox txtLeakGroup 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   2040
            TabIndex        =   108
            Text            =   "15"
            Top             =   1140
            Width           =   1095
         End
         Begin VB.CheckBox chkLEAKTEST 
            Caption         =   "LEAK CAL"
            Height          =   495
            Left            =   1320
            TabIndex        =   106
            Top             =   540
            Width           =   1515
         End
         Begin VB.Label lblTemp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "GROUP"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   960
            TabIndex        =   107
            Top             =   1200
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "SERIAL PORT"
         Height          =   2175
         Left            =   -72180
         TabIndex        =   91
         Top             =   6840
         Width           =   6315
         Begin VB.TextBox txtSteppingPort 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   11.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   2220
            MaxLength       =   20
            TabIndex        =   200
            Text            =   "0"
            Top             =   1380
            Width           =   735
         End
         Begin VB.TextBox txtSteppingPort 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   11.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   2220
            MaxLength       =   20
            TabIndex        =   198
            Text            =   "0"
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox txtNvhPort 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   11.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   5160
            MaxLength       =   20
            TabIndex        =   113
            Text            =   "2"
            Top             =   1380
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtPowerPort 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   11.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   2220
            MaxLength       =   20
            TabIndex        =   92
            Text            =   "0"
            Top             =   420
            Width           =   735
         End
         Begin VB.Label lblSteppingPort 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "STEPPING PORT 2"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   11.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   201
            Top             =   1440
            Width           =   1875
         End
         Begin VB.Label lblSteppingPort 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "STEPPING PORT 1"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   11.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   199
            Top             =   960
            Width           =   1875
         End
         Begin VB.Label lblNvhPort 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NVH PORT"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   11.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3870
            TabIndex        =   114
            Top             =   1440
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lblPowerPort 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "POWER PORT"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   11.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   93
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame fraPassword 
         Caption         =   "PASSWORD CHANGE"
         ForeColor       =   &H00000000&
         Height          =   3795
         Left            =   -67140
         TabIndex        =   72
         Top             =   1980
         Width           =   8535
         Begin VB.TextBox txtPassword 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   468
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   3120
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   75
            Top             =   1020
            Width           =   2325
         End
         Begin VB.TextBox txtPassword 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   468
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   3120
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   74
            Top             =   1980
            Width           =   2325
         End
         Begin VB.TextBox txtPassword 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   468
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   3120
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   73
            Top             =   2580
            Width           =   2325
         End
         Begin BHButton.BHImageButton btnPassword 
            Height          =   1995
            Left            =   5640
            TabIndex        =   76
            Top             =   1020
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   3519
            Caption         =   "CONFIRM"
            CaptionChecked  =   "BHImageButton3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   24
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonAttrib    =   2
            BackColor       =   15790320
            ImgOutLineSize  =   3
         End
         Begin VB.Label lblTemp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NEW PASSWORD"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   1035
            TabIndex        =   79
            Top             =   2040
            Width           =   1965
         End
         Begin VB.Label lblTemp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CURRENT PASSWORD"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   495
            TabIndex        =   78
            Top             =   1080
            Width           =   2505
         End
         Begin VB.Label lblTemp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CONFIRM PASSWORD"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   480
            TabIndex        =   77
            Top             =   2640
            Width           =   2520
         End
      End
      Begin VB.Frame fraLang 
         Caption         =   "LANGUAGE"
         Height          =   2775
         Left            =   -72180
         TabIndex        =   68
         Top             =   9720
         Width           =   4875
         Begin VB.OptionButton optLang 
            Caption         =   "ENGLISH"
            Height          =   495
            Index           =   1
            Left            =   1620
            TabIndex        =   70
            Top             =   1620
            Width           =   1815
         End
         Begin VB.OptionButton optLang 
            Caption         =   "KOREAN"
            Height          =   495
            Index           =   0
            Left            =   1620
            TabIndex        =   69
            Top             =   780
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.Label lblTemp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "±¹³» Àü¿ëÀº »ç¿ë ¾ÈÇÔ"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   2520
            TabIndex        =   71
            Top             =   2340
            Visible         =   0   'False
            Width           =   2205
         End
      End
      Begin VB.Frame fraHidden 
         Caption         =   "Hidden"
         Height          =   6195
         Left            =   -65280
         TabIndex        =   62
         Top             =   6180
         Visible         =   0   'False
         Width           =   6675
         Begin VB.CommandButton btnSerialCheck 
            Caption         =   "ALL SERIAL PORT CHECK"
            Height          =   915
            Left            =   3060
            TabIndex        =   90
            Top             =   540
            Width           =   1875
         End
         Begin VB.Frame Frame 
            Caption         =   "FTP"
            Height          =   2115
            Left            =   360
            TabIndex        =   81
            Top             =   3300
            Width           =   3015
            Begin VB.TextBox txtFtpInfo 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   960
               TabIndex        =   85
               Top             =   300
               Width           =   1815
            End
            Begin VB.TextBox txtFtpInfo 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   960
               TabIndex        =   84
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox txtFtpInfo 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   960
               TabIndex        =   83
               Top             =   1140
               Width           =   1815
            End
            Begin VB.TextBox txtFtpInfo 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   960
               TabIndex        =   82
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "IP"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   360
               TabIndex        =   89
               Top             =   360
               Width           =   165
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "PORT"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   360
               TabIndex        =   88
               Top             =   780
               Width           =   495
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "ID"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   360
               TabIndex        =   87
               Top             =   1200
               Width           =   180
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "PW"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   360
               TabIndex        =   86
               Top             =   1620
               Width           =   315
            End
         End
         Begin VB.CommandButton btnPrintName 
            Caption         =   "PRINTER NAME"
            Height          =   615
            Left            =   360
            TabIndex        =   80
            Top             =   1740
            Width           =   2415
         End
         Begin VB.TextBox txtDBCol 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   1560
            TabIndex        =   65
            Text            =   "30"
            Top             =   540
            Width           =   1215
         End
         Begin VB.TextBox txtDBRow 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   1560
            TabIndex        =   64
            Text            =   "1"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtPrintName 
            Height          =   405
            Left            =   360
            TabIndex        =   63
            Text            =   "Datamax-O'Neil E-4304B Mark III"
            Top             =   2460
            Width           =   4575
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "DB COL"
            Height          =   285
            Left            =   480
            TabIndex        =   67
            Top             =   600
            Width           =   870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "DB ROW"
            Height          =   285
            Left            =   480
            TabIndex        =   66
            Top             =   1140
            Width           =   975
         End
      End
      Begin VB.Frame fraGeneral 
         Caption         =   "GENERAL"
         Height          =   4515
         Index           =   0
         Left            =   -72180
         TabIndex        =   56
         Top             =   1980
         Width           =   4875
         Begin VB.CheckBox chkSideDoor 
            Caption         =   "SIDE DOOR CHECK"
            Height          =   375
            Left            =   960
            TabIndex        =   134
            Top             =   3720
            Visible         =   0   'False
            Width           =   3195
         End
         Begin VB.TextBox txtNgTableList 
            Height          =   435
            Left            =   3360
            TabIndex        =   95
            Top             =   1500
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkNgTable 
            Caption         =   "USE NG TABLE"
            Height          =   495
            Left            =   960
            TabIndex        =   94
            Top             =   1500
            Visible         =   0   'False
            Width           =   2052
         End
         Begin VB.CheckBox chkPLCUse 
            Caption         =   "USE PLC COMM"
            Height          =   375
            Left            =   960
            TabIndex        =   60
            Top             =   2100
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CheckBox chkScreenKeyboardUse 
            Caption         =   "USE DISPLAY KEYBOARD"
            Height          =   375
            Left            =   960
            TabIndex        =   59
            Top             =   2640
            Visible         =   0   'False
            Width           =   3195
         End
         Begin VB.CheckBox chkCaptureSend 
            Caption         =   "CAPTURE SEND (VISION)"
            Height          =   375
            Left            =   960
            TabIndex        =   58
            Top             =   3180
            Width           =   3195
         End
         Begin VB.TextBox txtRetest 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            IMEMode         =   3  'DISABLE
            Left            =   2340
            MaxLength       =   20
            TabIndex        =   57
            Text            =   "0"
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblTemp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "RETEST"
            Height          =   285
            Index           =   8
            Left            =   1335
            TabIndex        =   61
            Top             =   660
            Width           =   825
         End
      End
      Begin VB.Frame fraCal 
         Caption         =   "SETUP"
         Height          =   9615
         Index           =   0
         Left            =   -69660
         TabIndex        =   43
         Top             =   1560
         Width           =   12735
         Begin VB.TextBox txtMSAfterDelay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Left            =   10020
            TabIndex        =   48
            Text            =   "3.00"
            Top             =   420
            Width           =   1395
         End
         Begin VB.TextBox txtCalDelay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   9780
            TabIndex        =   47
            Text            =   "3.00"
            Top             =   1980
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCalUse 
            Caption         =   "CH 00"
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   46
            Top             =   2040
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txtCalMin 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   5700
            TabIndex        =   45
            Text            =   "0.00"
            Top             =   1980
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtCalMax 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   7740
            TabIndex        =   44
            Text            =   "999.00"
            Top             =   1980
            Visible         =   0   'False
            Width           =   1815
         End
         Begin Threed.SSPanel pnlCalName 
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   49
            Top             =   1980
            Visible         =   0   'False
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   767
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   1
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   4
            X1              =   9660
            X2              =   9660
            Y1              =   1380
            Y2              =   9000
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   3
            X1              =   7620
            X2              =   7620
            Y1              =   1380
            Y2              =   9000
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   2
            X1              =   5580
            X2              =   5580
            Y1              =   1380
            Y2              =   9000
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   1
            X1              =   2340
            X2              =   2340
            Y1              =   1380
            Y2              =   9000
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   0
            X1              =   960
            X2              =   11820
            Y1              =   1860
            Y2              =   1860
         End
         Begin VB.Label lblTitleTemp 
            AutoSize        =   -1  'True
            Caption         =   "Sec"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   11520
            TabIndex        =   55
            Top             =   480
            Width           =   390
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "MIN"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   12
            Left            =   5700
            TabIndex        =   54
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "MAX"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   11
            Left            =   7740
            TabIndex        =   53
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TEST DELAY"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   8520
            TabIndex        =   52
            Top             =   480
            Width           =   1380
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "TIME (Sec)"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   9
            Left            =   9780
            TabIndex        =   51
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "NAME"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   8
            Left            =   2640
            TabIndex        =   50
            Top             =   1440
            Width           =   2820
         End
      End
      Begin VB.Frame fraCal 
         Caption         =   "TYPE"
         Height          =   4875
         Index           =   4
         Left            =   -74460
         TabIndex        =   14
         Top             =   3900
         Width           =   4275
         Begin VB.TextBox txtMS4Time 
            Alignment       =   2  'Center
            Height          =   435
            Index           =   3
            Left            =   3420
            TabIndex        =   110
            Text            =   "00"
            Top             =   4080
            Width           =   675
         End
         Begin VB.TextBox txtMS4Time 
            Alignment       =   2  'Center
            Height          =   435
            Index           =   2
            Left            =   3420
            TabIndex        =   109
            Text            =   "00"
            Top             =   3480
            Width           =   675
         End
         Begin VB.TextBox txtMS4Time 
            Alignment       =   2  'Center
            Height          =   435
            Index           =   1
            Left            =   2640
            TabIndex        =   21
            Text            =   "00"
            Top             =   4080
            Width           =   675
         End
         Begin VB.TextBox txtMS4Time 
            Alignment       =   2  'Center
            Height          =   435
            Index           =   0
            Left            =   2640
            TabIndex        =   20
            Text            =   "00"
            Top             =   3480
            Width           =   675
         End
         Begin VB.OptionButton optMSUse 
            Caption         =   "USER SET TIME"
            Height          =   435
            Index           =   3
            Left            =   300
            TabIndex        =   19
            Top             =   2880
            Width           =   2235
         End
         Begin VB.TextBox txtMS2Time 
            Alignment       =   2  'Center
            Height          =   435
            Left            =   2700
            TabIndex        =   18
            Text            =   "0"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton optMSUse 
            Caption         =   "DISABLE MASTER SAMPLE"
            Height          =   435
            Index           =   0
            Left            =   300
            TabIndex        =   17
            Top             =   540
            Value           =   -1  'True
            Width           =   3375
         End
         Begin VB.OptionButton optMSUse 
            Caption         =   "EVERY HOURS"
            Height          =   435
            Index           =   1
            Left            =   300
            TabIndex        =   16
            Top             =   1320
            Width           =   1995
         End
         Begin VB.OptionButton optMSUse 
            Caption         =   "EACH MODEL CHANGE"
            Height          =   435
            Index           =   2
            Left            =   300
            TabIndex        =   15
            Top             =   2100
            Width           =   3015
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "MM"
            Height          =   315
            Index           =   3
            Left            =   3480
            TabIndex        =   112
            Top             =   3120
            Width           =   555
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "HH"
            Height          =   315
            Index           =   2
            Left            =   2700
            TabIndex        =   111
            Top             =   3120
            Width           =   555
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "SHIFT 2 (1200~2359)"
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   23
            Top             =   4140
            Width           =   2475
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "SHIFT 1 (0000~1159)"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   22
            Top             =   3540
            Width           =   2475
         End
      End
      Begin VB.Frame fraCal 
         Caption         =   "GENERAL"
         Height          =   2235
         Index           =   3
         Left            =   -74460
         TabIndex        =   8
         Top             =   1560
         Width           =   4275
         Begin VB.CheckBox chkMSTEST 
            Caption         =   "NG TEST"
            Height          =   495
            Index           =   1
            Left            =   2340
            TabIndex        =   11
            Top             =   1320
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.CheckBox chkMSTEST 
            Caption         =   "OK TEST"
            Height          =   495
            Index           =   0
            Left            =   480
            TabIndex        =   10
            Top             =   1320
            Width           =   1515
         End
         Begin VB.TextBox txtMSVolt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   1980
            TabIndex        =   9
            Text            =   "12.0"
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblTemp 
            Caption         =   "Volt"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   3180
            TabIndex        =   13
            Top             =   705
            Width           =   555
         End
         Begin VB.Label lblTemp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TEST VOLT"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   660
            TabIndex        =   12
            Top             =   660
            Width           =   1215
         End
      End
      Begin VB.Frame fraCorrelation 
         Caption         =   " OFFSET "
         Height          =   12015
         Left            =   240
         TabIndex        =   24
         Top             =   900
         Visible         =   0   'False
         Width           =   18255
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   14
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   253
            Top             =   8520
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   13
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   252
            Top             =   8040
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   12
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   251
            Top             =   7560
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   11
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   250
            Top             =   7080
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   10
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   249
            Top             =   6600
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   9
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   248
            Top             =   6120
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   8
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   247
            Top             =   5640
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   7
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   246
            Top             =   5160
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   0
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   244
            Top             =   1800
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   1
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   243
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   2
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   242
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   3
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   241
            Top             =   3240
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   4
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   240
            Top             =   3720
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   5
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   239
            Top             =   4200
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkPercent 
            Caption         =   "USE"
            Height          =   435
            Index           =   6
            Left            =   15600
            Style           =   1  'Graphical
            TabIndex        =   238
            Top             =   4680
            Width           =   1335
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   14
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   236
            Top             =   8520
            Width           =   1935
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   14
            Left            =   14220
            TabIndex        =   235
            Text            =   "0"
            Top             =   8520
            Width           =   1335
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   14
            Left            =   12840
            TabIndex        =   234
            Text            =   "#0.00"
            Top             =   8520
            Width           =   1335
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   14
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   233
            Top             =   8520
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   14
            Left            =   2280
            TabIndex        =   232
            Text            =   "SENSOR 6"
            Top             =   8520
            Width           =   2835
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   14
            Left            =   7980
            TabIndex        =   231
            Text            =   "1.00"
            Top             =   8520
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   14
            Left            =   6600
            TabIndex        =   230
            Text            =   "0.00"
            Top             =   8520
            Width           =   1335
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   14
            Left            =   5160
            TabIndex        =   229
            Text            =   "1"
            Top             =   8520
            Width           =   1395
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   13
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   227
            Top             =   8040
            Width           =   1935
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   13
            Left            =   14220
            TabIndex        =   226
            Text            =   "0"
            Top             =   8040
            Width           =   1335
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   13
            Left            =   12840
            TabIndex        =   225
            Text            =   "#0.00"
            Top             =   8040
            Width           =   1335
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   13
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   224
            Top             =   8040
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   13
            Left            =   2280
            TabIndex        =   223
            Text            =   "SENSOR 5"
            Top             =   8040
            Width           =   2835
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   13
            Left            =   7980
            TabIndex        =   222
            Text            =   "1.00"
            Top             =   8040
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   13
            Left            =   6600
            TabIndex        =   221
            Text            =   "0.00"
            Top             =   8040
            Width           =   1335
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   13
            Left            =   5160
            TabIndex        =   220
            Text            =   "1"
            Top             =   8040
            Width           =   1395
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   12
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   218
            Top             =   7560
            Width           =   1935
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   12
            Left            =   14220
            TabIndex        =   217
            Text            =   "0"
            Top             =   7560
            Width           =   1335
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   12
            Left            =   12840
            TabIndex        =   216
            Text            =   "#0.00"
            Top             =   7560
            Width           =   1335
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   12
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   215
            Top             =   7560
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   12
            Left            =   2280
            TabIndex        =   214
            Text            =   "SENSOR 4"
            Top             =   7560
            Width           =   2835
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   12
            Left            =   7980
            TabIndex        =   213
            Text            =   "1.00"
            Top             =   7560
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   12
            Left            =   6600
            TabIndex        =   212
            Text            =   "0.00"
            Top             =   7560
            Width           =   1335
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   12
            Left            =   5160
            TabIndex        =   211
            Text            =   "1"
            Top             =   7560
            Width           =   1395
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   11
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   209
            Top             =   7080
            Width           =   1935
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   11
            Left            =   14220
            TabIndex        =   208
            Text            =   "0"
            Top             =   7080
            Width           =   1335
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   11
            Left            =   12840
            TabIndex        =   207
            Text            =   "#0.00"
            Top             =   7080
            Width           =   1335
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   11
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   206
            Top             =   7080
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   11
            Left            =   2280
            TabIndex        =   205
            Text            =   "SENSOR 3"
            Top             =   7080
            Width           =   2835
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   11
            Left            =   7980
            TabIndex        =   204
            Text            =   "1.00"
            Top             =   7080
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   11
            Left            =   6600
            TabIndex        =   203
            Text            =   "0.00"
            Top             =   7080
            Width           =   1335
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   11
            Left            =   5160
            TabIndex        =   202
            Text            =   "1"
            Top             =   7080
            Width           =   1395
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   10
            Left            =   5160
            TabIndex        =   196
            Text            =   "1"
            Top             =   6600
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   10
            Left            =   6600
            TabIndex        =   195
            Text            =   "0.00"
            Top             =   6600
            Width           =   1335
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   10
            Left            =   7980
            TabIndex        =   194
            Text            =   "1.00"
            Top             =   6600
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   10
            Left            =   2280
            TabIndex        =   193
            Text            =   "SENSOR 2"
            Top             =   6600
            Width           =   2835
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   10
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   192
            Top             =   6600
            Width           =   1395
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   10
            Left            =   12840
            TabIndex        =   191
            Text            =   "#0.00"
            Top             =   6600
            Width           =   1335
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   10
            Left            =   14220
            TabIndex        =   190
            Text            =   "0"
            Top             =   6600
            Width           =   1335
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   10
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   189
            Top             =   6600
            Width           =   1935
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   9
            Left            =   5160
            TabIndex        =   187
            Text            =   "1"
            Top             =   6120
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   9
            Left            =   6600
            TabIndex        =   186
            Text            =   "0.00"
            Top             =   6120
            Width           =   1335
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   9
            Left            =   7980
            TabIndex        =   185
            Text            =   "1.00"
            Top             =   6120
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   9
            Left            =   2280
            TabIndex        =   184
            Text            =   "SENSOR 1"
            Top             =   6120
            Width           =   2835
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   9
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   183
            Top             =   6120
            Width           =   1395
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   9
            Left            =   12840
            TabIndex        =   182
            Text            =   "#0.00"
            Top             =   6120
            Width           =   1335
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   9
            Left            =   14220
            TabIndex        =   181
            Text            =   "0"
            Top             =   6120
            Width           =   1335
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   9
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   180
            Top             =   6120
            Width           =   1935
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   8
            Left            =   5160
            TabIndex        =   178
            Text            =   "1"
            Top             =   5640
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   8
            Left            =   6600
            TabIndex        =   177
            Text            =   "0.00"
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   8
            Left            =   7980
            TabIndex        =   176
            Text            =   "1.00"
            Top             =   5640
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   8
            Left            =   2280
            TabIndex        =   175
            Text            =   "ACT04 FEEDBACK"
            Top             =   5640
            Width           =   2835
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   8
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   174
            Top             =   5640
            Width           =   1395
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   8
            Left            =   12840
            TabIndex        =   173
            Text            =   "#0.00"
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   8
            Left            =   14220
            TabIndex        =   172
            Text            =   "0"
            Top             =   5640
            Width           =   1335
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   8
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   171
            Top             =   5640
            Width           =   1935
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   7
            Left            =   5160
            TabIndex        =   169
            Text            =   "1000"
            Top             =   5160
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   7
            Left            =   6600
            TabIndex        =   168
            Text            =   "0.00"
            Top             =   5160
            Width           =   1335
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   7
            Left            =   7980
            TabIndex        =   167
            Text            =   "1.00"
            Top             =   5160
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   7
            Left            =   2280
            TabIndex        =   166
            Text            =   "ACT04 CURRENT"
            Top             =   5160
            Width           =   2835
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   7
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   165
            Top             =   5160
            Width           =   1395
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   7
            Left            =   12840
            TabIndex        =   164
            Text            =   "#0"
            Top             =   5160
            Width           =   1335
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   7
            Left            =   14220
            TabIndex        =   163
            Text            =   "0"
            Top             =   5160
            Width           =   1335
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   7
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   162
            Top             =   5160
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   6
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   160
            Top             =   4680
            Width           =   1935
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   6
            Left            =   14220
            TabIndex        =   159
            Text            =   "0"
            Top             =   4680
            Width           =   1335
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   6
            Left            =   12840
            TabIndex        =   158
            Text            =   "#0.00"
            Top             =   4680
            Width           =   1335
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   6
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   157
            Top             =   4680
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   6
            Left            =   2280
            TabIndex        =   156
            Text            =   "ACT03 FEEDBACK"
            Top             =   4680
            Width           =   2835
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   6
            Left            =   7980
            TabIndex        =   155
            Text            =   "1.00"
            Top             =   4680
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   6
            Left            =   6600
            TabIndex        =   154
            Text            =   "0.00"
            Top             =   4680
            Width           =   1335
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   6
            Left            =   5160
            TabIndex        =   153
            Text            =   "1"
            Top             =   4680
            Width           =   1395
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   5
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   151
            Top             =   4200
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   5
            Left            =   14220
            TabIndex        =   150
            Text            =   "0"
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   12840
            TabIndex        =   149
            Text            =   "#0"
            Top             =   4200
            Width           =   1335
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   5
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   148
            Top             =   4200
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   2280
            TabIndex        =   147
            Text            =   "ACT03 CURRENT"
            Top             =   4200
            Width           =   2835
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   7980
            TabIndex        =   146
            Text            =   "1.00"
            Top             =   4200
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   6600
            TabIndex        =   145
            Text            =   "0.00"
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   5160
            TabIndex        =   144
            Text            =   "1000"
            Top             =   4200
            Width           =   1395
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   4
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   142
            Top             =   3720
            Width           =   1935
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   4
            Left            =   14220
            TabIndex        =   141
            Text            =   "0"
            Top             =   3720
            Width           =   1335
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   12840
            TabIndex        =   140
            Text            =   "#0.00"
            Top             =   3720
            Width           =   1335
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   4
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   139
            Top             =   3720
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   2280
            TabIndex        =   138
            Text            =   "ACT02 FEEDBACK"
            Top             =   3720
            Width           =   2835
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   7980
            TabIndex        =   137
            Text            =   "1.00"
            Top             =   3720
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   6600
            TabIndex        =   136
            Text            =   "0.00"
            Top             =   3720
            Width           =   1335
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   5160
            TabIndex        =   135
            Text            =   "1"
            Top             =   3720
            Width           =   1395
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   5160
            TabIndex        =   132
            Text            =   "1000"
            Top             =   3240
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   6600
            TabIndex        =   131
            Text            =   "0.00"
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   7980
            TabIndex        =   130
            Text            =   "1.00"
            Top             =   3240
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2280
            TabIndex        =   129
            Text            =   "ACT02 CURRENT"
            Top             =   3240
            Width           =   2835
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   3
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   3240
            Width           =   1395
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   12840
            TabIndex        =   127
            Text            =   "#0"
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   3
            Left            =   14220
            TabIndex        =   126
            Text            =   "0"
            Top             =   3240
            Width           =   1335
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   3
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   125
            Top             =   3240
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   5160
            TabIndex        =   123
            Text            =   "10"
            Top             =   1800
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   6600
            TabIndex        =   122
            Text            =   "0.00"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   7980
            TabIndex        =   121
            Text            =   "1.00"
            Top             =   1800
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2280
            TabIndex        =   120
            Text            =   "SUPPLY VOLT"
            Top             =   1800
            Width           =   2835
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   0
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   119
            Top             =   1800
            Width           =   1395
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   12840
            TabIndex        =   118
            Text            =   "#0.0"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   0
            Left            =   14220
            TabIndex        =   117
            Text            =   "0"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   0
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   116
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   2
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "USE"
            Height          =   435
            Index           =   1
            Left            =   10860
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   2280
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   1
            Left            =   14220
            TabIndex        =   102
            Text            =   "0"
            Top             =   2280
            Width           =   1335
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   12840
            TabIndex        =   101
            Text            =   "#0"
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   1
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   2280
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2280
            TabIndex        =   99
            Text            =   "ACT01 CURRENT"
            Top             =   2280
            Width           =   2835
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   7980
            TabIndex        =   98
            Text            =   "1.00"
            Top             =   2280
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   6600
            TabIndex        =   97
            Text            =   "0.00"
            Top             =   2280
            Width           =   1335
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   5160
            TabIndex        =   96
            Text            =   "1000"
            Top             =   2280
            Width           =   1395
         End
         Begin VB.TextBox txtFiltering 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   16020
            TabIndex        =   32
            Text            =   "3"
            Top             =   300
            Width           =   1335
         End
         Begin VB.TextBox txtDisp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   2
            Left            =   14220
            TabIndex        =   31
            Text            =   "0"
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   12840
            TabIndex        =   30
            Text            =   "#0.00"
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CheckBox chkMinus 
            Caption         =   "USE"
            Height          =   435
            Index           =   2
            Left            =   9420
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   2760
            Width           =   1395
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2280
            TabIndex        =   28
            Text            =   "ACT01 FEEDBACK"
            Top             =   2760
            Width           =   2835
         End
         Begin VB.TextBox txtMulti 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   7980
            TabIndex        =   27
            Text            =   "1.00"
            Top             =   2760
            Width           =   1395
         End
         Begin VB.TextBox txtAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   6600
            TabIndex        =   26
            Text            =   "0.00"
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox txtNaive 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   5160
            TabIndex        =   25
            Text            =   "1"
            Top             =   2760
            Width           =   1395
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "PERCENT"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   15
            Left            =   15600
            TabIndex        =   245
            Top             =   1320
            Width           =   1320
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 14"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   14
            Left            =   1200
            TabIndex        =   237
            Top             =   8580
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 13"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   13
            Left            =   1200
            TabIndex        =   228
            Top             =   8100
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 12"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   12
            Left            =   1200
            TabIndex        =   219
            Top             =   7620
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 11"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   11
            Left            =   1200
            TabIndex        =   210
            Top             =   7140
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 10"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   10
            Left            =   1200
            TabIndex        =   197
            Top             =   6660
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 09"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   9
            Left            =   1200
            TabIndex        =   188
            Top             =   6180
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 08"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   8
            Left            =   1200
            TabIndex        =   179
            Top             =   5700
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 07"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   7
            Left            =   1200
            TabIndex        =   170
            Top             =   5220
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 06"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   6
            Left            =   1200
            TabIndex        =   161
            Top             =   4740
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 05"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   1200
            TabIndex        =   152
            Top             =   4260
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 04"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   1200
            TabIndex        =   143
            Top             =   3780
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 03"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   3
            Left            =   1200
            TabIndex        =   133
            Top             =   3300
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 00"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   124
            Top             =   1860
            Width           =   975
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CH 01"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   1470
            TabIndex        =   104
            Top             =   2340
            Width           =   705
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "FILTERING :"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   7
            Left            =   14340
            TabIndex        =   42
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "ZERO BALANCE"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   6
            Left            =   10860
            TabIndex        =   41
            Top             =   1320
            Width           =   1920
         End
         Begin VB.Label lblLabelTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "CH 02"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   2
            Left            =   1200
            TabIndex        =   40
            Top             =   2820
            Width           =   975
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "DISPLAY"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   14220
            TabIndex        =   39
            Top             =   1320
            Width           =   1320
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "NAMING"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   4
            Left            =   2340
            TabIndex        =   38
            Top             =   1320
            Width           =   2700
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "FORMAT"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   12840
            TabIndex        =   37
            Top             =   1320
            Width           =   1320
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "MINUS"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   9420
            TabIndex        =   36
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "OFFSET"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   6660
            TabIndex        =   35
            Top             =   1320
            Width           =   1200
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "SLOPE"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   7980
            TabIndex        =   34
            Top             =   1320
            Width           =   1380
         End
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "STANDARD"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   13
            Left            =   5160
            TabIndex        =   33
            Top             =   1320
            Width           =   1365
         End
      End
      Begin VB.Frame fraSystem 
         Caption         =   "CORELLATION PASSWORD"
         ForeColor       =   &H00000000&
         Height          =   2475
         Index           =   2
         Left            =   6938
         TabIndex        =   5
         Top             =   6263
         Width           =   5235
         Begin VB.TextBox txtCorellation 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   468
            IMEMode         =   3  'DISABLE
            Left            =   2220
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   1080
            Width           =   2325
         End
         Begin VB.Label lblTemp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PASSWORD :"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   555
            TabIndex        =   7
            Top             =   1140
            Width           =   1485
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "DISPLAY = ((((Hardware Val) * STANDARD) + OFFSET) * SLOPE)"
         Height          =   315
         Left            =   9180
         TabIndex        =   2
         Top             =   13020
         Width           =   9435
      End
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   36
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10620
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   4035
   End
   Begin VB.CommandButton btnReturn 
      Caption         =   "RETURN"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   36
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   14820
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   4095
   End
   Begin HIRESTIMERLib.HiResTimer tmrSystemLoop 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   64
      Enabled         =   0   'False
      Interval        =   1
   End
   Begin VB.ListBox lstSystemMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   10275
   End
End
Attribute VB_Name = "frmSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim bLoading As Boolean

Private Sub Form_Activate()
    nNowForm = FM_SYSTEM
    
    Call LoadLangFile(FM_SYSTEM)
    
    If bLoading = False Then
        bLoading = True
        
        Call OnStart
    End If
End Sub

Private Sub Form_Load()
    bLoading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call OnEnd
End Sub

Private Sub btnReturn_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Button
        Case vbLeftButton: Unload Me
        Case vbRightButton: frmSystem.fraHidden.Visible = IIf(frmSystem.fraHidden.Visible = True, False, True)
    End Select
End Sub

Private Sub btnReturn_Click()
    Unload Me
End Sub

Private Sub btnPrintName_Click()
    If Printer.DeviceName <> "" Then
        frmSystem.txtPrintName.Text = Printer.DeviceName
    Else
        Call MsgBox("NO PRINTER")
    End If
End Sub

Private Sub btnSerialCheck_Click()
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_SerialPort", , 48)
    
    For Each objItem In colItems
        MsgBox objItem.DeviceID
        MsgBox objItem.Description
    Next
    
    Set colItems = Nothing
    Set objWMIService = Nothing
End Sub

Private Sub btnSave_Click()
    Call SystemDisp2Mem
    Call SaveSystemFile
    Call SaveLangFile
    Call OnLog("System File Save")
End Sub

Private Sub btnPassword_Click()
    If (UCase(Trim$(txtPassword(0).Text)) = UCase(Trim$(SysVar.lpPassword))) Or (UCase(Trim$(txtPassword(0).Text)) = MASTER_PASSWORD) Then
        If (Trim$(txtPassword(1).Text) = Trim$(txtPassword(2).Text)) Then
            SysVar.lpPassword = Trim$(txtPassword(1).Text)
            
            Call OnLog("Change Password...")
            Call SystemDisp2Mem
            Call SaveSystemFile
            Call OnLog("System File Save")
        End If
    End If
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
    txtCorellation.SetFocus
    
    If DEBUGMODE Then
        fraCorrelation.Visible = True
        fraCorrelation.ZOrder 0
    End If
End Sub

Private Sub txtCorellation_Change()
    Dim lpPassword As String
    
    lpPassword = UCase(Trim$(txtCorellation.Text))
    
    If lpPassword = MASTER_PASSWORD Then
        fraCorrelation.Visible = True
        fraCorrelation.ZOrder 0
    Else
        fraCorrelation.Visible = False
    End If
End Sub

Private Sub tmrSystemLoop_Timer()
    Call SystemLoop
End Sub

Private Sub OnStart()
    Dim bRes As Boolean
    
    SSTab.Tab = 0
    lpPath = App.Path
    
    Call SystemChCtlArray
    
    bRes = LoadModelName
    bRes = LoadSystemFile
    
    If bRes Then
        Call SystemMem2Disp
    Else
        If SETUPSELECT Then
            lpNowModel = Format(nNowModelNo, "0000") & "_" & SelectCar(0).ModelName & "_" & SelectCar(0).ModelNameSub(0)
        Else
            lpNowModel = "0000_DEFAULT"
        End If
        SysVar.lpModel = lpNowModel
    End If
    
    frmSystem.fraHidden.Visible = DEBUGMODE
    frmSystem.chkPLCUse.Visible = PLCUSE
    frmSystem.lblNvhPort.Visible = NVHUSE
    frmSystem.txtNvhPort.Visible = NVHUSE
    frmSystem.chkNgTable.Visible = IIf(TABLETYPE, DEBUGMODE, False)
    frmSystem.txtNgTableList.Visible = IIf(TABLETYPE, DEBUGMODE, False)
    
    If POWERTYPE <> 0 Then
        frmSystem.lblLabelTemp(0).Visible = False
        frmSystem.txtName(0).Visible = False
        frmSystem.txtNaive(0).Visible = False
        frmSystem.txtAdd(0).Visible = False
        frmSystem.txtMulti(0).Visible = False
        frmSystem.chkMinus(0).Visible = False
        frmSystem.chkZero(0).Visible = False
        frmSystem.txtUnit(0).Visible = False
        frmSystem.txtDisp(0).Visible = False
    End If
    
    tmrSystemLoop.Enabled = True
    
    Call OnLog("System File Load")
    Call OnLog("System Start")
End Sub

Private Sub OnEnd()
    tmrSystemLoop.Enabled = False
    Call OnLog("System End")
End Sub

Private Sub SystemLoop()
    Dim i As Integer
    
    nDispCounter = nDispCounter + 1
    
    If nDispCounter > DISP_TIME Then
        nDispCounter = 0
        bDisp = True
    End If
    
    If bDisp Then
        bDisp = False
        
        For i = 0 To MAX_AD_CHANNEL
            txtDisp(i).Text = Format(ADRead(i), SysVar.lpUnit(i))
        Next
    End If  ' bDisp
End Sub

Private Sub SystemChCtlArray()
    Dim i As Integer
    Dim Index As Integer
    
    frmSystem.pnlCalName(0).Outline = False
    
    ' ¹è¿­ 0, 1 ÄÁÆ®·Ñ Á¦¿Ü
    For i = 1 To MAX_AD_CHANNEL
        Index = frmSystem.chkCalUse.Count
        
        Call Load(frmSystem.chkCalUse(Index))
        Call Load(frmSystem.pnlCalName(Index))
        Call Load(frmSystem.txtCalMin(Index))
        Call Load(frmSystem.txtCalMax(Index))
        Call Load(frmSystem.txtCalDelay(Index))
        
        frmSystem.chkCalUse(Index).Top = frmSystem.chkCalUse(Index - 1).Top + 480
        frmSystem.pnlCalName(Index).Top = frmSystem.pnlCalName(Index - 1).Top + 480
        frmSystem.txtCalMin(Index).Top = frmSystem.txtCalMin(Index - 1).Top + 480
        frmSystem.txtCalMax(Index).Top = frmSystem.txtCalMax(Index - 1).Top + 480
        frmSystem.txtCalDelay(Index).Top = frmSystem.txtCalDelay(Index - 1).Top + 480
        
        frmSystem.chkCalUse(Index).Caption = "CH " & Format(Index, "00")
        
        frmSystem.chkCalUse(Index).Visible = True
        frmSystem.pnlCalName(Index).Visible = True
        frmSystem.txtCalMin(Index).Visible = True
        frmSystem.txtCalMax(Index).Visible = True
        frmSystem.txtCalDelay(Index).Visible = True
    Next
End Sub
