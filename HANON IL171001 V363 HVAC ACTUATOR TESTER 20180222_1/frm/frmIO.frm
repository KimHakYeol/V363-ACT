VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{F36F6844-D389-11D1-8968-006097AA579E}#1.0#0"; "HiResTimer.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmIO 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "I/O"
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15000
   ScaleWidth      =   19110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   9375
      Index           =   0
      Left            =   18660
      TabIndex        =   2
      Top             =   14100
      Visible         =   0   'False
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   16536
      _Version        =   393216
      Tab             =   2
      TabHeight       =   970
      TabCaption(0)   =   "HIDDEN"
      TabPicture(0)   =   "frmIO.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "pnlAD(0)"
      Tab(0).Control(1)=   "lblAdChName(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmIO.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmIO.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   0
         Left            =   -74400
         TabIndex        =   22
         Top             =   1426
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0.00"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME 0"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   0
         Left            =   -73860
         TabIndex        =   23
         Top             =   2266
         Width           =   945
      End
   End
   Begin Threed.SSPanel SSPanel 
      Height          =   1635
      Index           =   3
      Left            =   120
      TabIndex        =   25
      Top             =   180
      Width           =   10635
      _Version        =   65536
      _ExtentX        =   18759
      _ExtentY        =   2884
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
      Begin VB.TextBox txtTestVolt 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6300
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   840
         Width           =   1575
      End
      Begin BHButton.BHImageButton btnVoltSupply 
         Height          =   1275
         Left            =   8640
         TabIndex        =   27
         Top             =   180
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   2249
         Caption         =   "VOLT OUT"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         ForeColor       =   0
         BackColor       =   15790320
         ImgOutLineSize  =   3
      End
      Begin Threed.SSPanel pnlSetVolt 
         Height          =   615
         Left            =   3420
         TabIndex        =   28
         Top             =   180
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSetCurr 
         Height          =   615
         Left            =   3420
         TabIndex        =   29
         Top             =   840
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin VB.Label lblTemp 
         Caption         =   "SUPPLY VOLT :"
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   0
         Left            =   5760
         TabIndex        =   33
         Top             =   480
         Width           =   2235
      End
      Begin VB.Label lblTemp 
         Caption         =   "VOLT"
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   1
         Left            =   7920
         TabIndex        =   32
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SUPPLY VOLT"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   31
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SUPPLY CURR"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   30
         Top             =   1020
         Width           =   1575
      End
   End
   Begin BHButton.BHImageButton btnReturn 
      Height          =   1635
      Left            =   15000
      TabIndex        =   0
      Top             =   180
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2884
      Caption         =   "RETURN"
      CaptionChecked  =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ"
         Size            =   27.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentPicture=   "frmIO.frx":0054
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnSystemReset 
      Height          =   1095
      Left            =   10800
      TabIndex        =   1
      Top             =   720
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   1931
      Caption         =   "SYSTEM RESET"
      CaptionChecked  =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ"
         Size            =   26.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentPicture=   "frmIO.frx":E32A
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnClearAD 
      Height          =   495
      Left            =   10800
      TabIndex        =   3
      Top             =   180
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Caption         =   "INTIAL AD"
      CaptionChecked  =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentPicture=   "frmIO.frx":245BC
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnZeroAD 
      Height          =   495
      Left            =   12900
      TabIndex        =   4
      Top             =   180
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Caption         =   "ZERO AD"
      CaptionChecked  =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentPicture=   "frmIO.frx":3A84E
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin HIRESTIMERLib.HiResTimer tmrLoop 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   64
      Enabled         =   0   'False
      Interval        =   1
   End
   Begin Threed.SSPanel SSPanel 
      Height          =   12915
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   18855
      _Version        =   65536
      _ExtentX        =   33258
      _ExtentY        =   22781
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
      Begin VB.TextBox txtAct03Volt 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   18
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   15360
         TabIndex        =   88
         Text            =   "0.0"
         Top             =   2940
         Width           =   1455
      End
      Begin VB.TextBox txtAct02Volt 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   18
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   15360
         TabIndex        =   77
         Text            =   "0.0"
         Top             =   1620
         Width           =   1455
      End
      Begin VB.TextBox txtAct01Volt 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   18
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   15360
         TabIndex        =   68
         Text            =   "0.0"
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox txtAct04Volt 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   18
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   15360
         TabIndex        =   34
         Text            =   "0.0"
         Top             =   4260
         Width           =   1455
      End
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   2
         Left            =   3420
         TabIndex        =   6
         Top             =   960
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   1
         Left            =   3420
         TabIndex        =   7
         Top             =   300
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   2475
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   10260
         Width           =   9135
         _Version        =   65536
         _ExtentX        =   16113
         _ExtentY        =   4366
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
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   2
            Left            =   3120
            TabIndex        =   13
            Top             =   1680
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y02] OK LAMP"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   3
            Left            =   6060
            TabIndex        =   14
            Top             =   120
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y03] NG LAMP"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   4
            Left            =   6060
            TabIndex        =   15
            Top             =   900
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y04] RUN LAMP"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
         Begin BHButton.BHImageButton btnDI 
            Height          =   675
            Index           =   0
            Left            =   180
            TabIndex        =   16
            Top             =   120
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[X00] START S/W"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Gap             =   1
            ButtonAttrib    =   2
            ImgOutLineSize  =   1
            ImgUpOutLineSize=   1
            ImgDownOutLineSize=   1
            ImgDisableOutLineSize=   1
         End
         Begin BHButton.BHImageButton btnDI 
            Height          =   675
            Index           =   1
            Left            =   180
            TabIndex        =   17
            Top             =   900
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[X01] STOP S/W"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Gap             =   1
            ButtonAttrib    =   2
            ImgOutLineSize  =   1
            ImgUpOutLineSize=   1
            ImgDownOutLineSize=   1
            ImgDisableOutLineSize=   1
         End
         Begin BHButton.BHImageButton btnDI 
            Height          =   675
            Index           =   2
            Left            =   180
            TabIndex        =   18
            Top             =   1680
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[X02] AUTO/MANU"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Gap             =   1
            ButtonAttrib    =   2
            ImgOutLineSize  =   1
            ImgUpOutLineSize=   1
            ImgDownOutLineSize=   1
            ImgDisableOutLineSize=   1
         End
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   0
            Left            =   3120
            TabIndex        =   19
            Top             =   120
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y00] START LAMP"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   1
            Left            =   3120
            TabIndex        =   20
            Top             =   900
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y01] STOP LAMP"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   5
            Left            =   6060
            TabIndex        =   21
            Top             =   1680
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y05] BUZZER"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   2475
         Index           =   2
         Left            =   9540
         TabIndex        =   24
         Top             =   10260
         Width           =   9135
         _Version        =   65536
         _ExtentX        =   16113
         _ExtentY        =   4366
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
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   6
            Left            =   180
            TabIndex        =   47
            Top             =   120
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y06] MASTER OK"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   16
            Left            =   180
            TabIndex        =   66
            Top             =   900
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y16] STEP #1 RESET"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   17
            Left            =   180
            TabIndex        =   67
            Top             =   1680
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y17] STEP #2 RESET"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   20
            Left            =   3120
            TabIndex        =   105
            Top             =   900
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y20] ACT01 CHANGE"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   21
            Left            =   6060
            TabIndex        =   106
            Top             =   900
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y21] ACT02 CHANGE"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   22
            Left            =   3120
            TabIndex        =   107
            Top             =   1680
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y22] ACT03 CHANGE"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
         Begin BHButton.BHImageButton btnDO 
            Height          =   675
            Index           =   23
            Left            =   6060
            TabIndex        =   108
            Top             =   1680
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1191
            Caption         =   "[Y23] ACT04 CHANGE"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12
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
      End
      Begin BHButton.BHImageButton btnAct04Pos 
         Height          =   1275
         Index           =   0
         Left            =   8640
         TabIndex        =   35
         Top             =   4260
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         Caption         =   "P1"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnDO 
         Height          =   1275
         Index           =   10
         Left            =   5580
         TabIndex        =   36
         Top             =   300
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2249
         Caption         =   "[Y10] ACT01 POWER"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   3
         Left            =   3420
         TabIndex        =   37
         Top             =   1620
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   4
         Left            =   3420
         TabIndex        =   40
         Top             =   2280
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   5
         Left            =   3420
         TabIndex        =   43
         Top             =   2940
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin BHButton.BHImageButton btnDO 
         Height          =   3915
         Index           =   14
         Left            =   5580
         TabIndex        =   46
         Top             =   5580
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   6906
         Caption         =   "[Y14] SENSOR POWER"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   6
         Left            =   3420
         TabIndex        =   48
         Top             =   3600
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   7
         Left            =   3420
         TabIndex        =   51
         Top             =   4260
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   8
         Left            =   3420
         TabIndex        =   54
         Top             =   4920
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   9
         Left            =   3420
         TabIndex        =   57
         Top             =   5580
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   10
         Left            =   3420
         TabIndex        =   60
         Top             =   6240
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin BHButton.BHImageButton btnDO 
         Height          =   1275
         Index           =   11
         Left            =   5580
         TabIndex        =   63
         Top             =   1620
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2249
         Caption         =   "[Y11] ACT02 POWER"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnDO 
         Height          =   1275
         Index           =   12
         Left            =   5580
         TabIndex        =   64
         Top             =   2940
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2249
         Caption         =   "[Y12] ACT03 POWER"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnDO 
         Height          =   1275
         Index           =   13
         Left            =   5580
         TabIndex        =   65
         Top             =   4260
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2249
         Caption         =   "[Y13] ACT04 POWER"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct01Pos 
         Height          =   1275
         Index           =   0
         Left            =   8640
         TabIndex        =   69
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   2249
         Caption         =   "P1"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         ForeColor       =   0
         BackColor       =   15790320
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnAct01Pos 
         Height          =   1275
         Index           =   1
         Left            =   9600
         TabIndex        =   70
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   2249
         Caption         =   "P2"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         ForeColor       =   0
         BackColor       =   15790320
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnAct01Pos 
         Height          =   1275
         Index           =   2
         Left            =   10560
         TabIndex        =   71
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   2249
         Caption         =   "P3"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         ForeColor       =   0
         BackColor       =   15790320
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnAct01Pos 
         Height          =   1275
         Index           =   3
         Left            =   11520
         TabIndex        =   72
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   2249
         Caption         =   "P4"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         ForeColor       =   0
         BackColor       =   15790320
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnAct01Pos 
         Height          =   1275
         Index           =   4
         Left            =   12480
         TabIndex        =   73
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   2249
         Caption         =   "P5"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         ForeColor       =   0
         BackColor       =   15790320
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnAct01Pos 
         Height          =   1275
         Index           =   8
         Left            =   16860
         TabIndex        =   74
         Top             =   300
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2249
         Caption         =   "SET"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct01Pos 
         Height          =   1275
         Index           =   5
         Left            =   13440
         TabIndex        =   75
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   2249
         Caption         =   "P6"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         ForeColor       =   0
         BackColor       =   15790320
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnAct01Pos 
         Height          =   1275
         Index           =   6
         Left            =   14400
         TabIndex        =   76
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   2249
         Caption         =   "P7"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         ForeColor       =   0
         BackColor       =   15790320
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnAct02Pos 
         Height          =   1275
         Index           =   0
         Left            =   8640
         TabIndex        =   78
         Top             =   1620
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         Caption         =   "P1"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct02Pos 
         Height          =   1275
         Index           =   1
         Left            =   10320
         TabIndex        =   79
         Top             =   1620
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         Caption         =   "P2"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct02Pos 
         Height          =   1275
         Index           =   4
         Left            =   16860
         TabIndex        =   80
         Top             =   1620
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2249
         Caption         =   "SET"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct02Pos 
         Height          =   1275
         Index           =   2
         Left            =   12000
         TabIndex        =   81
         Top             =   1620
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         Caption         =   "P3"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct02Pos 
         Height          =   1275
         Index           =   3
         Left            =   13680
         TabIndex        =   82
         Top             =   1620
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         Caption         =   "P4"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct03Pos 
         Height          =   1275
         Index           =   0
         Left            =   8640
         TabIndex        =   83
         Top             =   2940
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         Caption         =   "P1"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct03Pos 
         Height          =   1275
         Index           =   1
         Left            =   10320
         TabIndex        =   84
         Top             =   2940
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         Caption         =   "P2"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct03Pos 
         Height          =   1275
         Index           =   2
         Left            =   12000
         TabIndex        =   85
         Top             =   2940
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         Caption         =   "P3"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct03Pos 
         Height          =   1275
         Index           =   3
         Left            =   13680
         TabIndex        =   86
         Top             =   2940
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         Caption         =   "P4"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct03Pos 
         Height          =   1275
         Index           =   4
         Left            =   16860
         TabIndex        =   87
         Top             =   2940
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2249
         Caption         =   "SET"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct04Pos 
         Height          =   1275
         Index           =   1
         Left            =   10320
         TabIndex        =   89
         Top             =   4260
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         Caption         =   "P2"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct04Pos 
         Height          =   1275
         Index           =   2
         Left            =   12000
         TabIndex        =   90
         Top             =   4260
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         Caption         =   "P3"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct04Pos 
         Height          =   1275
         Index           =   3
         Left            =   13680
         TabIndex        =   91
         Top             =   4260
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2249
         Caption         =   "P4"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin BHButton.BHImageButton btnAct04Pos 
         Height          =   1275
         Index           =   4
         Left            =   16860
         TabIndex        =   92
         Top             =   4260
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2249
         Caption         =   "SET"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
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
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   11
         Left            =   3420
         TabIndex        =   93
         Top             =   6900
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   12
         Left            =   3420
         TabIndex        =   96
         Top             =   7560
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   13
         Left            =   3420
         TabIndex        =   99
         Top             =   8220
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAD 
         Height          =   615
         Index           =   14
         Left            =   3420
         TabIndex        =   102
         Top             =   8880
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "0"
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 14"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   15
         Left            =   300
         TabIndex        =   104
         Top             =   9060
         Width           =   705
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   14
         Left            =   1200
         TabIndex        =   103
         Top             =   9060
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 13"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   14
         Left            =   300
         TabIndex        =   101
         Top             =   8400
         Width           =   705
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   13
         Left            =   1200
         TabIndex        =   100
         Top             =   8400
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 12"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   13
         Left            =   300
         TabIndex        =   98
         Top             =   7740
         Width           =   705
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   12
         Left            =   1200
         TabIndex        =   97
         Top             =   7740
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 11"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   12
         Left            =   300
         TabIndex        =   95
         Top             =   7080
         Width           =   705
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   11
         Left            =   1200
         TabIndex        =   94
         Top             =   7080
         Width           =   735
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   10
         Left            =   1200
         TabIndex        =   62
         Top             =   6420
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 10"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   9
         Left            =   300
         TabIndex        =   61
         Top             =   6420
         Width           =   705
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   9
         Left            =   1200
         TabIndex        =   59
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 09"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   8
         Left            =   300
         TabIndex        =   58
         Top             =   5760
         Width           =   705
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   56
         Top             =   5100
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 08"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   7
         Left            =   300
         TabIndex        =   55
         Top             =   5100
         Width           =   705
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   53
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 07"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   6
         Left            =   300
         TabIndex        =   52
         Top             =   4440
         Width           =   705
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   50
         Top             =   3780
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 06"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   5
         Left            =   300
         TabIndex        =   49
         Top             =   3780
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 05"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   4
         Left            =   300
         TabIndex        =   45
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   44
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 04"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   3
         Left            =   300
         TabIndex        =   42
         Top             =   2460
         Width           =   705
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   41
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 03"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   39
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   38
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   11
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label lblAdChName 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 01"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   10
         Left            =   300
         TabIndex        =   9
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CH 02"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   11
         Left            =   300
         TabIndex        =   8
         Top             =   1140
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bLoading As Boolean

Private Sub btnAct01Pos_Click(Index As Integer)
    Dim i As Integer
    Dim lpRotation As String
    Dim lpPos(1) As String
    
    Select Case Index
        Case frmIO.btnAct01Pos.UBound:
            For i = 0 To frmIO.btnAct01Pos.UBound - 1
                frmIO.btnAct01Pos(i).ForeColor = vbBlack
            Next
            
            frmIO.btnAct01Pos(Index).ForeColor = vbRed
            
            If SetupVar.nAct01TestType = 1 Then
                txtAct01Volt.Text = Format(Val(txtAct01Volt.Text), SysVar.lpUnit(ActNo(0).AD_VOLT))
            End If
        
        Case Else:
            For i = 0 To frmIO.btnAct01Pos.UBound - 1
                frmIO.btnAct01Pos(i).ForeColor = IIf(Index = i, vbRed, vbBlack)
            Next
            
            If SetupVar.nAct01TestType = 1 Then
                txtAct01Volt.Text = Format(SetupVar.dAct01SetVolt(Index), SysVar.lpUnit(ActNo(0).AD_VOLT))
            End If
        
    End Select
    
    Select Case SetupVar.nAct01TestType
        Case 0:
            Select Case Index
                ' ACT01
                Case 0:
                    lpRotation = "00": lpPos(0) = STEP_POS1: lpPos(1) = STEP_POS2
                Case 1, 2, 3, 4, 5:
                    lpRotation = "01"
                    lpPos(0) = Dec2Hex(Val2Byte(CM_LO, SetupVar.dAct01SetVolt(Index)))
                    lpPos(1) = Dec2Hex(Val2Byte(CM_HI, SetupVar.dAct01SetVolt(Index)))
                Case 6:
                    lpRotation = "01": lpPos(0) = STEP_POS1: lpPos(1) = STEP_POS2
            End Select
            
            Call StepSend(STEP_START & STEP_STOP & "01" & "00" & "00" & "00" & STEP_END, 0)
            Call StepSend(STEP_START & STEP_MOVE & "01" & lpRotation & lpPos(1) & lpPos(0) & STEP_END, 0)
        Case 1:
            Call OutDa(ActNo(0).DA_NO, Val(frmIO.txtAct01Volt.Text), SysVar.bPercent(ActNo(0).AD_VOLT))
    End Select
End Sub

Private Sub btnAct02Pos_Click(Index As Integer)
    Dim i As Integer
    Dim lpRotation As String
    Dim lpPos(1) As String
    
    Select Case Index
        Case frmIO.btnAct02Pos.UBound:
            For i = 0 To frmIO.btnAct02Pos.UBound - 1
                frmIO.btnAct02Pos(i).ForeColor = vbBlack
            Next
            
            frmIO.btnAct02Pos(Index).ForeColor = vbRed
            
            frmIO.txtAct02Volt.Text = Format(Val(frmIO.txtAct02Volt.Text), SysVar.lpUnit(ActNo(1).AD_VOLT))
            
        Case Else:
            For i = 0 To frmIO.btnAct02Pos.UBound - 1
                frmIO.btnAct02Pos(i).ForeColor = IIf(Index = i, vbRed, vbBlack)
            Next
            
            If SetupVar.nAct02TestType = 1 Then
                frmIO.txtAct02Volt.Text = Format(SetupVar.dAct02SetVolt(Index), SysVar.lpUnit(ActNo(1).AD_VOLT))
            End If
    
    End Select
    
    Select Case SetupVar.nAct02TestType
        Case 0:
            Select Case Index
                ' ACT02
                Case 0:
                    lpRotation = "00": lpPos(0) = STEP_POS1: lpPos(1) = STEP_POS2
                Case 1, 2:
                    lpRotation = "01"
                    lpPos(0) = Dec2Hex(Val2Byte(CM_LO, SetupVar.dAct02SetVolt(Index)))
                    lpPos(1) = Dec2Hex(Val2Byte(CM_HI, SetupVar.dAct02SetVolt(Index)))
                Case 3:
                    lpRotation = "01": lpPos(0) = STEP_POS1: lpPos(1) = STEP_POS2
            End Select
            
            Call StepSend(STEP_START & STEP_STOP & "02" & "00" & "00" & "00" & STEP_END, 0)
            Call StepSend(STEP_START & STEP_MOVE & "02" & lpRotation & lpPos(1) & lpPos(0) & STEP_END, 0)
        Case 1:
            Call OutDa(ActNo(1).DA_NO, Val(frmIO.txtAct02Volt.Text), SysVar.bPercent(ActNo(1).AD_VOLT))
    End Select
End Sub

Private Sub btnAct03Pos_Click(Index As Integer)
    Dim i As Integer
    Dim lpRotation As String
    Dim lpPos(1) As String
    
    Select Case Index
        Case frmIO.btnAct03Pos.UBound:
            For i = 0 To frmIO.btnAct03Pos.UBound - 1
                frmIO.btnAct03Pos(i).ForeColor = vbBlack
            Next
            
            frmIO.btnAct03Pos(Index).ForeColor = vbRed
            
            frmIO.txtAct03Volt.Text = Format(Val(frmIO.txtAct03Volt.Text), SysVar.lpUnit(ActNo(2).AD_VOLT))
            
        Case Else:
            For i = 0 To frmIO.btnAct03Pos.UBound - 1
                frmIO.btnAct03Pos(i).ForeColor = IIf(Index = i, vbRed, vbBlack)
            Next
            
            If SetupVar.nAct03TestType = 1 Then
                frmIO.txtAct03Volt.Text = Format(SetupVar.dAct03SetVolt(Index), SysVar.lpUnit(ActNo(2).AD_VOLT))
            End If
    
    End Select
    
    Select Case SetupVar.nAct03TestType
        Case 0:
            Select Case Index
                ' ACT03
                Case 0:
                    lpRotation = "00": lpPos(0) = STEP_POS1: lpPos(1) = STEP_POS2
                Case 1, 2:
                    lpRotation = "01"
                    lpPos(0) = Dec2Hex(Val2Byte(CM_LO, SetupVar.dAct03SetVolt(Index)))
                    lpPos(1) = Dec2Hex(Val2Byte(CM_HI, SetupVar.dAct03SetVolt(Index)))
                Case 3:
                    lpRotation = "01": lpPos(0) = STEP_POS1: lpPos(1) = STEP_POS2
            End Select
            
            Call StepSend(STEP_START & STEP_STOP & "01" & "00" & "00" & "00" & STEP_END, 1)
            Call StepSend(STEP_START & STEP_MOVE & "01" & lpRotation & lpPos(1) & lpPos(0) & STEP_END, 1)
        Case 1:
            Call OutDa(ActNo(2).DA_NO, Val(frmIO.txtAct03Volt.Text), SysVar.bPercent(ActNo(2).AD_VOLT))
    End Select
End Sub

Private Sub btnAct04Pos_Click(Index As Integer)
    Dim i As Integer
    Dim lpRotation As String
    Dim lpPos(1) As String
    
    Select Case Index
        Case frmIO.btnAct04Pos.UBound:
            For i = 0 To frmIO.btnAct04Pos.UBound - 1
                frmIO.btnAct04Pos(i).ForeColor = vbBlack
            Next
            
            frmIO.btnAct04Pos(Index).ForeColor = vbRed
            
            frmIO.txtAct04Volt.Text = Format(Val(frmIO.txtAct04Volt.Text), SysVar.lpUnit(ActNo(3).AD_VOLT))
            
        Case Else:
            For i = 0 To frmIO.btnAct04Pos.UBound - 1
                frmIO.btnAct04Pos(i).ForeColor = IIf(Index = i, vbRed, vbBlack)
            Next
            
            If SetupVar.nAct04TestType = 1 Then
                frmIO.txtAct04Volt.Text = Format(SetupVar.dAct04SetVolt(Index), SysVar.lpUnit(ActNo(3).AD_VOLT))
            End If
    
    End Select
    
    Select Case SetupVar.nAct04TestType
        Case 0:
            Select Case Index
                ' ACT04
                Case 0:
                    lpRotation = "00": lpPos(0) = STEP_POS1: lpPos(1) = STEP_POS2
                Case 1, 2:
                    lpRotation = "01"
                    lpPos(0) = Dec2Hex(Val2Byte(CM_LO, SetupVar.dAct04SetVolt(Index)))
                    lpPos(1) = Dec2Hex(Val2Byte(CM_HI, SetupVar.dAct04SetVolt(Index)))
                Case 3:
                    lpRotation = "01": lpPos(0) = STEP_POS1: lpPos(1) = STEP_POS2
            End Select
            
            Call StepSend(STEP_START & STEP_STOP & "02" & "00" & "00" & "00" & STEP_END, 1)
            Call StepSend(STEP_START & STEP_MOVE & "02" & lpRotation & lpPos(1) & lpPos(0) & STEP_END, 1)
        Case 1:
            Call OutDa(ActNo(3).DA_NO, Val(frmIO.txtAct04Volt.Text), SysVar.bPercent(ActNo(3).AD_VOLT))
    End Select
End Sub

Private Sub btnReturn_Click()
    Call OnEnd
End Sub

Private Sub btnSystemReset_Click()
    frmMain.tmrPower.Enabled = False
    
    Call Sleep(250)
    Call AD_Close
    Call StepClose
    
    If POWERTYPE > 0 Then Call PowerClose
    
    bSplash = True
    Unload Me
End Sub

Private Sub btnVoltSupply_Click()
    Call SetVolt(Format(Val(frmIO.txtTestVolt.Text), "#0.0"))
End Sub

Private Sub btnDO_Click(Index As Integer)
    Dim bMove As Boolean
    
    bMove = True
    
    Select Case Index
        Case O_LIN_POWER:
            If SetupVar.nBlowerType = 2 Then
                Call LinBlrWrite(SetupVar.nLinSpeed(0))
            End If
    
    End Select
    
    If TABLETYPE Then
        bMove = MarkingInterlock(Index)
    End If
    
    ' ÁøÂ¥ ¿òÁ÷ÀÌ´Â ºÎºÐ
    If bMove Then
        Call DO_CtrlinIO(Index)
    End If
End Sub

Private Sub btnClearAD_Click()
    Call ClearAD
End Sub

Private Sub btnZeroAD_Click()
    Call ZeroAD
End Sub

Private Sub Form_Activate()
    nNowForm = FM_IO
    
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

Private Sub tmrLoop_Timer()
    Call IOLoop
End Sub

Private Sub OnStart()
    Dim i As Integer
    Dim bDoRes(MAX_DIO_CHANNEL) As Boolean
    Dim bDiRes(MAX_DIO_CHANNEL) As Boolean
    Dim btnCtl As BHImageButton
    
    For Each btnCtl In frmIO.btnDO
        bDoRes(btnCtl.Index) = True
    Next
    
    For i = 0 To MAX_DIO_CHANNEL - 1
        If bDoRes(i) Then
            Call DOS_Color(i)
        End If
    Next
    
    For Each btnCtl In frmIO.btnDI
        bDiRes(btnCtl.Index) = True
    Next
    
    For i = 0 To MAX_DIO_CHANNEL - 1
        If bDiRes(i) Then
            Select Case i
                Case I_START_SW, I_STOP_SW, I_AUTO_SW:
                
                Case Else:
                    frmIO.btnDI(i).Caption = "[X" & Format(i, "00") & "]"
            End Select
        End If
    Next
    
    For i = 0 To MAX_AD_CHANNEL
        frmIO.lblAdChName(i).Caption = SysVar.lpName(i)
    Next
    
    txtTestVolt.Text = Format(SetupVar.dTestVolt, SysVar.lpUnit(AD_SUPPLY_VOLT))
    
    Call SetVolt(SetupVar.dTestVolt)
    
    Call ZeroAD
    
    tmrLoop.Enabled = True
End Sub

Private Sub OnEnd()
    Call DO_Clear
    tmrLoop.Enabled = False
    Call Delay(100)
    Unload Me
End Sub

Private Sub DIS_Color(ByVal nCh As Integer)
    btnDI(nCh).BackColor = IIf(DIS(nCh), vbGreen, CO_NONE)
End Sub

Private Sub DOS_Color(ByVal nCh As Integer)
    btnDO(nCh).ForeColor = IIf(DOS(nCh), vbRed, vbBlack)
End Sub

Private Sub DO_CtrlinIO(ByVal nCh As Integer)
    Dim bOutData    As Boolean
    
    bOutData = IIf(DOS(nCh), False, True)

    Call DO_Control(nCh, bOutData)
    Call DOS_Color(nCh)
End Sub

Private Sub IOLoop()
    Dim i As Integer
    Dim bDiRes(MAX_DIO_CHANNEL) As Boolean
    Dim btnCtl As BHImageButton
    
    For Each btnCtl In btnDI
        bDiRes(btnCtl.Index) = True
    Next
    
    nDispCounter = nDispCounter + 1
    
    If nDispCounter > DISP_TIME Then
        nDispCounter = 0
        bDisp = True
    End If
    
    If bDisp Then
        bDisp = False
        
        For i = 0 To MAX_DIO_CHANNEL - 1
            If bDiRes(i) Then
                Call DIS_Color(i)
            End If
        Next
        
        If POWERGET Then
            If POWERTYPE = 0 Then
                dSupplyVolt = Format(ADRead(AD_SUPPLY_VOLT), "0.0")
                pnlSetVolt.Caption = Format(dSupplyVolt, "0.0")
            Else
                pnlSetVolt.Caption = Format(dSupplyVolt, "0.0")
                pnlSetCurr.Caption = Format(dSupplyCurr, "0.0")
            End If
        End If
        
        For i = 1 To MAX_AD_CHANNEL
            pnlAD(i).Caption = Format(ADRead(i), SysVar.lpUnit(i))
        Next
    End If
End Sub

