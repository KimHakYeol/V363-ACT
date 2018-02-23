VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmSetupSelect 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   14925
   ClientLeft      =   120
   ClientTop       =   375
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   14925
   ScaleWidth      =   19080
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlFrame 
      Height          =   12735
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   2100
      Width           =   18915
      _Version        =   65536
      _ExtentX        =   33364
      _ExtentY        =   22463
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Begin Threed.SSPanel pnlCarName 
         Height          =   1155
         Index           =   0
         Left            =   1500
         TabIndex        =   5
         Top             =   1320
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarNo 
         Height          =   1155
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   1320
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   0
         Left            =   3900
         TabIndex        =   7
         Top             =   1320
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   1335
         Index           =   1
         Left            =   1500
         TabIndex        =   13
         Top             =   0
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   2355
         _StockProps     =   15
         Caption         =   "MODEL"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   1335
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2355
         _StockProps     =   15
         Caption         =   "NO."
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   615
         Index           =   2
         Left            =   3900
         TabIndex        =   15
         Top             =   0
         Width           =   15015
         _Version        =   65536
         _ExtentX        =   26485
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "SEQUENCE"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.26
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   735
         Index           =   3
         Left            =   3900
         TabIndex        =   17
         Top             =   600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   735
         Index           =   4
         Left            =   5400
         TabIndex        =   18
         Top             =   600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   735
         Index           =   5
         Left            =   6900
         TabIndex        =   19
         Top             =   600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   735
         Index           =   6
         Left            =   8400
         TabIndex        =   20
         Top             =   600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   735
         Index           =   7
         Left            =   9900
         TabIndex        =   21
         Top             =   600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   735
         Index           =   8
         Left            =   11400
         TabIndex        =   22
         Top             =   600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   735
         Index           =   9
         Left            =   12900
         TabIndex        =   23
         Top             =   600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   735
         Index           =   10
         Left            =   14400
         TabIndex        =   24
         Top             =   600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   735
         Index           =   11
         Left            =   15900
         TabIndex        =   25
         Top             =   600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   735
         Index           =   12
         Left            =   17400
         TabIndex        =   26
         Top             =   600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarNo 
         Height          =   1155
         Index           =   1
         Left            =   0
         TabIndex        =   27
         Top             =   2460
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarNo 
         Height          =   1155
         Index           =   2
         Left            =   0
         TabIndex        =   28
         Top             =   3600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarNo 
         Height          =   1155
         Index           =   3
         Left            =   0
         TabIndex        =   29
         Top             =   4740
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarNo 
         Height          =   1155
         Index           =   4
         Left            =   0
         TabIndex        =   30
         Top             =   5880
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarNo 
         Height          =   1155
         Index           =   5
         Left            =   0
         TabIndex        =   31
         Top             =   7020
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarNo 
         Height          =   1155
         Index           =   6
         Left            =   0
         TabIndex        =   32
         Top             =   8160
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarNo 
         Height          =   1155
         Index           =   7
         Left            =   0
         TabIndex        =   33
         Top             =   9300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarNo 
         Height          =   1155
         Index           =   8
         Left            =   0
         TabIndex        =   34
         Top             =   10440
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarNo 
         Height          =   1155
         Index           =   9
         Left            =   0
         TabIndex        =   35
         Top             =   11580
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarName 
         Height          =   1155
         Index           =   1
         Left            =   1500
         TabIndex        =   36
         Top             =   2460
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarName 
         Height          =   1155
         Index           =   2
         Left            =   1500
         TabIndex        =   37
         Top             =   3600
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarName 
         Height          =   1155
         Index           =   3
         Left            =   1500
         TabIndex        =   38
         Top             =   4740
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarName 
         Height          =   1155
         Index           =   4
         Left            =   1500
         TabIndex        =   39
         Top             =   5880
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarName 
         Height          =   1155
         Index           =   5
         Left            =   1500
         TabIndex        =   40
         Top             =   7020
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarName 
         Height          =   1155
         Index           =   6
         Left            =   1500
         TabIndex        =   41
         Top             =   8160
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarName 
         Height          =   1155
         Index           =   7
         Left            =   1500
         TabIndex        =   42
         Top             =   9300
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarName 
         Height          =   1155
         Index           =   8
         Left            =   1500
         TabIndex        =   43
         Top             =   10440
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlCarName 
         Height          =   1155
         Index           =   9
         Left            =   1500
         TabIndex        =   44
         Top             =   11580
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   1
         Left            =   5400
         TabIndex        =   45
         Top             =   1320
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   2
         Left            =   6900
         TabIndex        =   46
         Top             =   1320
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   3
         Left            =   8400
         TabIndex        =   47
         Top             =   1320
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   4
         Left            =   9900
         TabIndex        =   48
         Top             =   1320
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   5
         Left            =   11400
         TabIndex        =   49
         Top             =   1320
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   6
         Left            =   12900
         TabIndex        =   50
         Top             =   1320
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   7
         Left            =   14400
         TabIndex        =   51
         Top             =   1320
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   8
         Left            =   15900
         TabIndex        =   52
         Top             =   1320
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   9
         Left            =   17400
         TabIndex        =   53
         Top             =   1320
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   10
         Left            =   3900
         TabIndex        =   54
         Top             =   2460
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   11
         Left            =   5400
         TabIndex        =   55
         Top             =   2460
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   12
         Left            =   6900
         TabIndex        =   56
         Top             =   2460
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   13
         Left            =   8400
         TabIndex        =   57
         Top             =   2460
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   14
         Left            =   9900
         TabIndex        =   58
         Top             =   2460
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   15
         Left            =   11400
         TabIndex        =   59
         Top             =   2460
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   16
         Left            =   12900
         TabIndex        =   60
         Top             =   2460
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   17
         Left            =   14400
         TabIndex        =   61
         Top             =   2460
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   18
         Left            =   15900
         TabIndex        =   62
         Top             =   2460
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   19
         Left            =   17400
         TabIndex        =   63
         Top             =   2460
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   20
         Left            =   3900
         TabIndex        =   64
         Top             =   3600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   21
         Left            =   5400
         TabIndex        =   65
         Top             =   3600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   22
         Left            =   6900
         TabIndex        =   66
         Top             =   3600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   23
         Left            =   8400
         TabIndex        =   67
         Top             =   3600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   24
         Left            =   9900
         TabIndex        =   68
         Top             =   3600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   25
         Left            =   11400
         TabIndex        =   69
         Top             =   3600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   26
         Left            =   12900
         TabIndex        =   70
         Top             =   3600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   27
         Left            =   14400
         TabIndex        =   71
         Top             =   3600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   28
         Left            =   15900
         TabIndex        =   72
         Top             =   3600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   29
         Left            =   17400
         TabIndex        =   73
         Top             =   3600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   30
         Left            =   3900
         TabIndex        =   74
         Top             =   4740
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   31
         Left            =   5400
         TabIndex        =   75
         Top             =   4740
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   32
         Left            =   6900
         TabIndex        =   76
         Top             =   4740
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   33
         Left            =   8400
         TabIndex        =   77
         Top             =   4740
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   34
         Left            =   9900
         TabIndex        =   78
         Top             =   4740
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   35
         Left            =   11400
         TabIndex        =   79
         Top             =   4740
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   36
         Left            =   12900
         TabIndex        =   80
         Top             =   4740
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   37
         Left            =   14400
         TabIndex        =   81
         Top             =   4740
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   38
         Left            =   15900
         TabIndex        =   82
         Top             =   4740
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   39
         Left            =   17400
         TabIndex        =   83
         Top             =   4740
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   40
         Left            =   3900
         TabIndex        =   84
         Top             =   5880
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   41
         Left            =   5400
         TabIndex        =   85
         Top             =   5880
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   42
         Left            =   6900
         TabIndex        =   86
         Top             =   5880
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   43
         Left            =   8400
         TabIndex        =   87
         Top             =   5880
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   44
         Left            =   9900
         TabIndex        =   88
         Top             =   5880
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   45
         Left            =   11400
         TabIndex        =   89
         Top             =   5880
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   46
         Left            =   12900
         TabIndex        =   90
         Top             =   5880
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   47
         Left            =   14400
         TabIndex        =   91
         Top             =   5880
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   48
         Left            =   15900
         TabIndex        =   92
         Top             =   5880
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   49
         Left            =   17400
         TabIndex        =   93
         Top             =   5880
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   50
         Left            =   3900
         TabIndex        =   94
         Top             =   7020
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   51
         Left            =   5400
         TabIndex        =   95
         Top             =   7020
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   52
         Left            =   6900
         TabIndex        =   96
         Top             =   7020
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   53
         Left            =   8400
         TabIndex        =   97
         Top             =   7020
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   54
         Left            =   9900
         TabIndex        =   98
         Top             =   7020
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   55
         Left            =   11400
         TabIndex        =   99
         Top             =   7020
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   56
         Left            =   12900
         TabIndex        =   100
         Top             =   7020
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   57
         Left            =   14400
         TabIndex        =   101
         Top             =   7020
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   58
         Left            =   15900
         TabIndex        =   102
         Top             =   7020
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   59
         Left            =   17400
         TabIndex        =   103
         Top             =   7020
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   60
         Left            =   3900
         TabIndex        =   104
         Top             =   8160
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   61
         Left            =   5400
         TabIndex        =   105
         Top             =   8160
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   62
         Left            =   6900
         TabIndex        =   106
         Top             =   8160
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   63
         Left            =   8400
         TabIndex        =   107
         Top             =   8160
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   64
         Left            =   9900
         TabIndex        =   108
         Top             =   8160
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   65
         Left            =   11400
         TabIndex        =   109
         Top             =   8160
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   66
         Left            =   12900
         TabIndex        =   110
         Top             =   8160
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   67
         Left            =   14400
         TabIndex        =   111
         Top             =   8160
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   68
         Left            =   15900
         TabIndex        =   112
         Top             =   8160
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   69
         Left            =   17400
         TabIndex        =   113
         Top             =   8160
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   70
         Left            =   3900
         TabIndex        =   114
         Top             =   9300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   71
         Left            =   5400
         TabIndex        =   115
         Top             =   9300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   72
         Left            =   6900
         TabIndex        =   116
         Top             =   9300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   73
         Left            =   8400
         TabIndex        =   117
         Top             =   9300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   74
         Left            =   9900
         TabIndex        =   118
         Top             =   9300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   75
         Left            =   11400
         TabIndex        =   119
         Top             =   9300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   76
         Left            =   12900
         TabIndex        =   120
         Top             =   9300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   77
         Left            =   14400
         TabIndex        =   121
         Top             =   9300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   78
         Left            =   15900
         TabIndex        =   122
         Top             =   9300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   79
         Left            =   17400
         TabIndex        =   123
         Top             =   9300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   80
         Left            =   3900
         TabIndex        =   124
         Top             =   10440
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   81
         Left            =   5400
         TabIndex        =   125
         Top             =   10440
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   82
         Left            =   6900
         TabIndex        =   126
         Top             =   10440
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   83
         Left            =   8400
         TabIndex        =   127
         Top             =   10440
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   84
         Left            =   9900
         TabIndex        =   128
         Top             =   10440
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   85
         Left            =   11400
         TabIndex        =   129
         Top             =   10440
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   86
         Left            =   12900
         TabIndex        =   130
         Top             =   10440
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   87
         Left            =   14400
         TabIndex        =   131
         Top             =   10440
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   88
         Left            =   15900
         TabIndex        =   132
         Top             =   10440
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   89
         Left            =   17400
         TabIndex        =   133
         Top             =   10440
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   90
         Left            =   3900
         TabIndex        =   134
         Top             =   11580
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   91
         Left            =   5400
         TabIndex        =   135
         Top             =   11580
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   92
         Left            =   6900
         TabIndex        =   136
         Top             =   11580
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   93
         Left            =   8400
         TabIndex        =   137
         Top             =   11580
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   94
         Left            =   9900
         TabIndex        =   138
         Top             =   11580
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   95
         Left            =   11400
         TabIndex        =   139
         Top             =   11580
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   96
         Left            =   12900
         TabIndex        =   140
         Top             =   11580
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   97
         Left            =   14400
         TabIndex        =   141
         Top             =   11580
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   98
         Left            =   15900
         TabIndex        =   142
         Top             =   11580
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlRankName 
         Height          =   1155
         Index           =   99
         Left            =   17400
         TabIndex        =   143
         Top             =   11580
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   2037
         _StockProps     =   15
         Caption         =   "1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlFrame 
      Height          =   1875
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   120
      Width           =   18915
      _Version        =   65536
      _ExtentX        =   33364
      _ExtentY        =   3307
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Begin VB.CommandButton btnAllModelSave 
         Caption         =   "All Model Save"
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   1500
         Width           =   1515
      End
      Begin VB.CommandButton btnReturn 
         Caption         =   "RETURN"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   27.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   14100
         TabIndex        =   12
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton btnSetup 
         Caption         =   "GROUP SETUP"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   27.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   9480
         TabIndex        =   11
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton btnNameSave 
         Caption         =   "MODEL SAVE"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   27.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4860
         TabIndex        =   10
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton btnModelLoading 
         Caption         =   "LOADING PLC"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   27.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   4335
      End
   End
   Begin Threed.SSPanel pnlLoadingPlc 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _Version        =   65536
      _ExtentX        =   15901
      _ExtentY        =   5530
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         Caption         =   "LOADING PLC PROGRASS"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   36
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Top             =   660
         Width           =   4860
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProgressPersent 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   36
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6180
         TabIndex        =   2
         Top             =   1140
         Width           =   1545
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   36
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   1
         Left            =   7800
         TabIndex        =   1
         Top             =   1140
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmSetupSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bLoading As Boolean

Private Sub btnAllModelSave_Click()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To CARSAVECOUNT - 1
        For j = 0 To SUBSAVECOUNT - 1
            If SelectCar(i).ModelName <> "" Then
                Call SaveSetupFile(Format(i * SUBSAVECOUNT + j, "0000") & "_" & SelectCar(i).ModelName & "_" & SelectCar(i).ModelNameSub(j))
'                Call OnLog(i * SUBSAVECOUNT + j)
                Call Delay(100)
            End If
        Next
    Next
    
    Call MsgBox("COMPLETE...")
End Sub

Private Sub btnModelLoading_Click()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim nCnt(9) As Integer
    Dim lRes(9) As Long
    Dim lArrRes1(9) As Long
    Dim lArrRes2(9) As Long
    Dim lpAddrModel(9) As String
    Dim lAddrModel(9) As Long
    Dim lpAddrNo(9) As String
    Dim lptxtName(9) As String
    Dim nReloadCount As Integer
    
    nReloadCount = 10
    
    DoEvents
    
    frmSetupSelect.pnlLoadingPlc.Visible = True
    frmSetupSelect.pnlLoadingPlc.ZOrder 0
    frmSetupSelect.pnlLoadingPlc.Top = frmSetupSelect.Height / 2 - frmSetupSelect.pnlLoadingPlc.Height / 2
    frmSetupSelect.pnlLoadingPlc.Left = frmSetupSelect.Width / 2 - frmSetupSelect.pnlLoadingPlc.Width / 2
    
    Call SelectButtonEnable(False)
    
    If PLCUSE = False Then
        For i = 0 To CARSAVECOUNT
            frmSetupSelect.lblProgressPersent.Caption = Format((i + 1) * (100 / (CARSAVECOUNT + 1)), "#0")
            
            Call Delay(200)
        Next
        
        If Val(frmSetupSelect.lblProgressPersent.Caption) >= 100 Then
            Call Delay(500)
            
            frmSetupSelect.lblProgressPersent.Caption = ""
            frmSetupSelect.pnlLoadingPlc.Visible = False
        End If
        
        GoTo LOADEND
    End If
    
    For i = 0 To 2
        lpAddrModel(i) = Left(PlcVar.lpAddrLoadModel(i), 1)
        lAddrModel(i) = Val(Mid(PlcVar.lpAddrLoadModel(i), 2))
    Next
    
    If lAddrModel(0) = 0 Or lAddrModel(1) = 0 Or lAddrModel(2) = 0 Then
        GoTo LOADEND
    End If
    
    For i = 0 To CARSAVECOUNT - 1
        nCnt(0) = 0

MODELRELOAD:
        
        lpAddrNo(0) = lpAddrModel(0) & lAddrModel(0) + (i * PlcVar.nNextLoadModel(0))
        lRes(0) = frmMain.ActPlc.ReadDeviceBlock(lpAddrNo(0), PlcVar.nSizeLoadModel(0), lArrRes1(0))
        
        If lRes(0) = 0 Then
            lptxtName(0) = ""
            
            For j = 0 To PlcVar.nSizeLoadModel(0) - 1
                lptxtName(0) = lptxtName(0) & AsciiChange(lArrRes1(j))
            Next
        Else
            nCnt(0) = nCnt(0) + 1
            
            If nCnt(0) > nReloadCount Then
                GoTo LOADERROR
            Else
                GoTo MODELRELOAD
            End If
        End If
        
        If frmSetupSelect.pnlCarName(i).Caption <> Trim$(lptxtName(0)) Then
            frmSetupSelect.pnlCarName(i).ForeColor = vbRed
        Else
            frmSetupSelect.pnlCarName(i).ForeColor = vbBlack
        End If
        
        SelectCar(i).ModelName = Trim(lptxtName(0))
        frmSetupSelect.pnlCarName(i).Caption = Trim(lptxtName(0))
        
        For k = 0 To SUBSAVECOUNT - 1
            nCnt(1) = 0

SUBMODELRELOAD:
            
            lpAddrNo(1) = lpAddrModel(1) & lAddrModel(1) + (k * PlcVar.nNextLoadModel(1))
            lRes(1) = frmMain.ActPlc.ReadDeviceBlock(lpAddrNo(1), PlcVar.nSizeLoadModel(1), lArrRes2(0))
            
            If lRes(1) = 0 Then
                lptxtName(1) = ""
                
                For j = 0 To PlcVar.nSizeLoadModel(1) - 1
                    lptxtName(1) = lptxtName(1) & AsciiChange(lArrRes2(j))
                Next
            Else
                nCnt(1) = nCnt(1) + 1
                
                If nCnt(1) > nReloadCount Then
                    GoTo LOADERROR
                Else
                    GoTo SUBMODELRELOAD
                End If
            End If
            
            If frmSetupSelect.pnlRankName(k).Caption <> Trim$(lptxtName(1)) Then
                frmSetupSelect.pnlRankName(k).ForeColor = vbRed
            Else
                frmSetupSelect.pnlRankName(k).ForeColor = vbBlack
            End If
            
            SelectCar(i).ModelNameSub(k) = Trim$(lptxtName(1))
            frmSetupSelect.pnlRankName(k).Caption = Trim$(lptxtName(1))
        Next k
    Next i

LOADEND:
    
    frmSetupSelect.pnlLoadingPlc.Visible = False
    
    Call Delay(100)
    Call btnNameSave_Click
    Call SelectButtonEnable(True)
    
    Exit Sub

LOADERROR:

    frmSetupSelect.pnlLoadingPlc.Visible = False
    
    Call MsgBox("MODEL LOAD READING ERROR... PLAESE TRY AGAIN...")
    Call PLC_Close
    Call Delay(100)
    Call PLC_Open
    Call SelectButtonEnable(True)
End Sub

Private Sub btnNameSave_Click()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To CARSAVECOUNT - 1
        SelectCar(i).ModelName = frmSetupSelect.pnlCarName(i).Caption
        
        For j = 0 To SUBSAVECOUNT - 1
            SelectCar(i).ModelNameSub(j) = frmSetupSelect.pnlRankName((i * SUBSAVECOUNT) + j).Caption
        Next
    Next
    
    Call SaveModelName
    Call MsgBox("Save complete...")
End Sub

Private Sub btnReturn_Click()
    Unload Me
End Sub

Private Sub btnSetup_Click()
    Dim i As Integer
    
    For i = 0 To TOTALSAVECOUNT - 1
        If frmSetupSelect.pnlRankName(i).BackColor = vbGreen Then
            nNowModelNo = i
            
            Exit For
        End If
    Next
    
    If i > TOTALSAVECOUNT - 1 Then
        Call MsgBox("Model select please...")
        
        Exit Sub
    End If
    
    If frmSetupSelect.pnlRankName(i).Caption = "" Then
        Call MsgBox("Not select model...")
        
        Exit Sub
    End If
    
    lpNowModel = Format(nNowModelNo, "0000") & "_" & frmSetupSelect.pnlCarName(Int(nNowModelNo / 3)).Caption & "_" & frmSetupSelect.pnlRankName(nNowModelNo).Caption
    
    frmSetupSelect.Hide
    frmSetup.Show 1
End Sub

Private Sub pnlCarName_Click(Index As Integer)
    Dim lpRes As String
    
    If TABLETYPE And PLCUSE = False Then
        lpRes = InputBox("¹Ù²î¾î¾ßÇÏ´Â ¸ðµ¨¸íÀ» ÀÔ·ÂÇÏ¼¼¿ä.", frmSetupSelect.pnlCarName(Index).Caption, frmSetupSelect.pnlCarName(Index).Caption)
        
        frmSetupSelect.pnlCarName(Index).Caption = lpRes
    End If
End Sub

Private Sub pnlRankName_Click(Index As Integer)
    Dim lpRes As String
    Dim lpOldName As String
    
    lpOldName = frmSetupSelect.pnlRankName(Index).Caption
    
    If TABLETYPE And PLCUSE = False And frmSetupSelect.pnlRankName(Index).BackColor = vbGreen Then
        lpRes = InputBox("¹Ù²î¾î¾ßÇÏ´Â ¼­¿­¸íÀ» ÀÔ·ÂÇÏ¼¼¿ä.", frmSetupSelect.pnlRankName(Index).Caption, frmSetupSelect.pnlRankName(Index).Caption)
        
        Select Case lpRes
            Case "": frmSetupSelect.pnlRankName(Index).Caption = lpOldName
            Case " ": frmSetupSelect.pnlRankName(Index).Caption = ""
            Case Else: frmSetupSelect.pnlRankName(Index).Caption = lpRes
        End Select
    End If
    
    Call pnlNameClear
    
    frmSetupSelect.pnlRankName(Index).BackColor = vbGreen
End Sub

Private Sub Form_Activate()
    nNowForm = FM_SETUP
    
    Call LoadLangFile(FM_SETUPSELECT)
    
    If bLoading = False Then
        Call OnStart
        
        bLoading = True
    End If
End Sub

Private Sub Form_Load()
    bLoading = False
End Sub

Private Sub OnStart()
    Dim i As Integer
    Dim j As Integer
    Dim bRes As Boolean
    
    frmSetupSelect.pnlLoadingPlc.Visible = False
    frmSetupSelect.btnModelLoading.Visible = PLCUSE
    
    bRes = LoadModelName
    
    If bRes = False Then Exit Sub
    
    For i = 0 To SUBSAVECOUNT - 1
        frmSetupSelect.pnlTitle(i + 3).Caption = i + 1
    Next
    
    For i = 0 To CARSAVECOUNT - 1
        frmSetupSelect.pnlCarNo(i).Caption = i + 1
        frmSetupSelect.pnlCarName(i).Caption = Trim$(SelectCar(i).ModelName)
        
        For j = 0 To SUBSAVECOUNT - 1
            frmSetupSelect.pnlRankName((i * SUBSAVECOUNT) + j).Caption = Trim$(SelectCar(i).ModelNameSub(j))
        Next
    Next
End Sub

Private Sub pnlNameClear()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To CARSAVECOUNT - 1
        For j = 0 To SUBSAVECOUNT - 1
            If frmSetupSelect.pnlRankName((i * SUBSAVECOUNT) + j).BackColor = vbGreen Then frmSetupSelect.pnlRankName((i * SUBSAVECOUNT) + j).BackColor = vbWhite
        Next
    Next
End Sub

Private Function SelectButtonEnable(ByVal bRes As Boolean)
    Dim pnlNameCtl As SSPanel
    
    frmSetupSelect.btnModelLoading.Enabled = bRes
    frmSetupSelect.btnNameSave.Enabled = bRes
    frmSetupSelect.btnReturn.Enabled = bRes
    frmSetupSelect.btnSetup.Enabled = bRes
    
    For Each pnlNameCtl In frmSetupSelect.pnlRankName
        frmSetupSelect.pnlRankName(pnlNameCtl.Index).Enabled = bRes
    Next
End Function

