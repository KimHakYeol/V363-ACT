VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{F36F6844-D389-11D1-8968-006097AA579E}#1.0#0"; "HiResTimer.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmRun 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RUN"
   ClientHeight    =   15000
   ClientLeft      =   45
   ClientTop       =   390
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRun.frx":0000
   ScaleHeight     =   15000
   ScaleWidth      =   19110
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   9255
      Left            =   18360
      TabIndex        =   25
      Top             =   13920
      Visible         =   0   'False
      Width           =   18105
      _ExtentX        =   31935
      _ExtentY        =   16325
      _Version        =   393216
      Tabs            =   9
      Tab             =   8
      TabsPerRow      =   9
      TabHeight       =   1411
      TabCaption(0)   =   "PLC"
      TabPicture(0)   =   "frmRun.frx":0C42
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraNGCount"
      Tab(0).Control(1)=   "pnlModelFullName"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "BLOWER"
      TabPicture(1)   =   "frmRun.frx":0C5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pnlFrame(5)"
      Tab(1).Control(1)=   "pnlNvhRpm"
      Tab(1).Control(2)=   "pnlFrame(6)"
      Tab(1).Control(3)=   "pnlFrame(9)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "ION"
      TabPicture(2)   =   "frmRun.frx":0C7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pnlFrame(10)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "LEAK"
      TabPicture(3)   =   "frmRun.frx":0C96
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pnlFrame(7)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "VISION"
      TabPicture(4)   =   "frmRun.frx":0CB2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "pnlFrame(8)"
      Tab(4).Control(1)=   "pnlVisionOpen(0)"
      Tab(4).Control(2)=   "pnlVisionClose(2)"
      Tab(4).Control(3)=   "pnlVisionOpen(2)"
      Tab(4).Control(4)=   "pnlVisionDoorOpen(2)"
      Tab(4).Control(5)=   "pnlVisionDoorClose(2)"
      Tab(4).Control(6)=   "pnlVisionDoorOpenNum(2)"
      Tab(4).Control(7)=   "pnlVisionDoorCloseNum(2)"
      Tab(4).Control(8)=   "pnlVisionClose(1)"
      Tab(4).Control(9)=   "pnlVisionOpen(1)"
      Tab(4).Control(10)=   "pnlVisionDoorOpen(1)"
      Tab(4).Control(11)=   "pnlVisionDoorClose(1)"
      Tab(4).Control(12)=   "pnlVisionDoorOpenNum(1)"
      Tab(4).Control(13)=   "pnlVisionDoorCloseNum(1)"
      Tab(4).Control(14)=   "pnlVisionClose(0)"
      Tab(4).Control(15)=   "pnlVisionDoorOpen(0)"
      Tab(4).Control(16)=   "pnlVisionDoorClose(0)"
      Tab(4).Control(17)=   "pnlVisionDoorOpenNum(0)"
      Tab(4).Control(18)=   "pnlVisionDoorCloseNum(0)"
      Tab(4).Control(19)=   "pnlVisionClose(3)"
      Tab(4).Control(20)=   "pnlVisionDoorClose(3)"
      Tab(4).Control(21)=   "pnlVisionDoorCloseNum(3)"
      Tab(4).Control(22)=   "pnlVisionOpen(3)"
      Tab(4).Control(23)=   "pnlVisionDoorOpen(3)"
      Tab(4).Control(24)=   "pnlVisionDoorOpenNum(3)"
      Tab(4).Control(25)=   "btnManualVisionClose(1)"
      Tab(4).Control(26)=   "btnManualVisionOpen(1)"
      Tab(4).Control(27)=   "btnManualVisionClose(2)"
      Tab(4).Control(28)=   "btnManualVisionOpen(2)"
      Tab(4).Control(29)=   "btnManualVisionClose(0)"
      Tab(4).Control(30)=   "btnManualVisionOpen(0)"
      Tab(4).ControlCount=   31
      TabCaption(5)   =   "LIN"
      TabPicture(5)   =   "frmRun.frx":0CCE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "pnlLinFrame(2)"
      Tab(5).Control(1)=   "pnlLinFrame(0)"
      Tab(5).Control(2)=   "Frame"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "ACT"
      TabPicture(6)   =   "frmRun.frx":0CEA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "I/O"
      TabPicture(7)   =   "frmRun.frx":0D06
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "btnBarcodePrint"
      Tab(7).Control(1)=   "pnlFrame(15)"
      Tab(7).Control(2)=   "btnPartLR"
      Tab(7).ControlCount=   3
      TabCaption(8)   =   "SENSOR"
      TabPicture(8)   =   "frmRun.frx":0D22
      Tab(8).ControlEnabled=   -1  'True
      Tab(8).ControlCount=   0
      Begin VB.CommandButton btnManualVisionOpen 
         BackColor       =   &H000080FF&
         Caption         =   "CAM1 OPEN"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   -71700
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton btnManualVisionClose 
         BackColor       =   &H000080FF&
         Caption         =   "CAM1 CLOSE"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   -70500
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton btnManualVisionOpen 
         BackColor       =   &H000080FF&
         Caption         =   "CAM2 OPEN"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   -66900
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton btnManualVisionClose 
         BackColor       =   &H000080FF&
         Caption         =   "CAM2 CLOSE"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   -65700
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton btnManualVisionOpen 
         BackColor       =   &H000080FF&
         Caption         =   "CAM2 OPEN"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   -69300
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton btnManualVisionClose 
         BackColor       =   &H000080FF&
         Caption         =   "CAM2 CLOSE"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   -68100
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   3660
         Width           =   1215
      End
      Begin VB.Frame Frame 
         Height          =   2295
         Left            =   -63360
         TabIndex        =   118
         Top             =   5460
         Width           =   3315
         Begin Threed.SSPanel pnlPTCVolt 
            Height          =   975
            Left            =   180
            TabIndex        =   119
            Top             =   1140
            Width           =   2955
            _Version        =   65536
            _ExtentX        =   5212
            _ExtentY        =   1720
            _StockProps     =   15
            Caption         =   "0.0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   27.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlPTCName 
            Height          =   735
            Left            =   180
            TabIndex        =   120
            Top             =   360
            Width           =   2955
            _Version        =   65536
            _ExtentX        =   5212
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "PTC"
            BackColor       =   8454143
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel pnlVisionDoorOpenNum 
         Height          =   675
         Index           =   3
         Left            =   -65760
         TabIndex        =   29
         Top             =   2040
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "DOOR 1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorOpen 
         Height          =   675
         Index           =   3
         Left            =   -65760
         TabIndex        =   30
         Top             =   2700
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionOpen 
         Height          =   615
         Index           =   3
         Left            =   -65760
         TabIndex        =   31
         Top             =   1440
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "OPEN"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorCloseNum 
         Height          =   675
         Index           =   3
         Left            =   -64860
         TabIndex        =   32
         Top             =   2040
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "DOOR 1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorClose 
         Height          =   675
         Index           =   3
         Left            =   -64860
         TabIndex        =   33
         Top             =   2700
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionClose 
         Height          =   615
         Index           =   3
         Left            =   -64860
         TabIndex        =   34
         Top             =   1440
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "CLOSE"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlFrame 
         Height          =   2175
         Index           =   7
         Left            =   -73860
         TabIndex        =   35
         Top             =   1380
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   3836
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
         Begin Threed.SSPanel pnlManual 
            Height          =   2055
            Index           =   8
            Left            =   60
            TabIndex        =   36
            Top             =   60
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   3625
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Begin VB.CommandButton btnManualLeak 
               BackColor       =   &H000080FF&
               Caption         =   "HTR LEAK"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   15.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1035
               Index           =   0
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   0
               Width           =   1935
            End
            Begin VB.CommandButton btnManualLeak 
               BackColor       =   &H000080FF&
               Caption         =   "EVA LEAK"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   15.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1035
               Index           =   1
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   1020
               Width           =   1935
            End
         End
         Begin Threed.SSPanel pnlLeakName 
            Height          =   1035
            Index           =   1
            Left            =   60
            TabIndex        =   39
            Top             =   1080
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   1826
            _StockProps     =   15
            Caption         =   "EVA LEAK"
            BackColor       =   8454143
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLeakData 
            Height          =   495
            Index           =   0
            Left            =   2100
            TabIndex        =   40
            Top             =   60
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "0.00"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLeakData 
            Height          =   495
            Index           =   1
            Left            =   2100
            TabIndex        =   41
            Top             =   1080
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "0.00"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLeakResult 
            Height          =   555
            Index           =   0
            Left            =   2100
            TabIndex        =   42
            Top             =   540
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLeakResult 
            Height          =   555
            Index           =   1
            Left            =   2100
            TabIndex        =   43
            Top             =   1560
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLeakName 
            Height          =   1035
            Index           =   0
            Left            =   60
            TabIndex        =   44
            Top             =   60
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   1826
            _StockProps     =   15
            Caption         =   "HTR LEAK"
            BackColor       =   8454143
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin BHButton.BHImageButton btnPartLR 
         Height          =   495
         Left            =   -69060
         TabIndex        =   45
         Top             =   2700
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   873
         Caption         =   "PART 2"
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
      Begin Threed.SSPanel pnlLinFrame 
         Height          =   1875
         Index           =   0
         Left            =   -73680
         TabIndex        =   109
         Top             =   5520
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
         _ExtentY        =   3307
         _StockProps     =   15
         BackColor       =   15790320
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
         Begin Threed.SSPanel pnlLinItem 
            Height          =   1755
            Index           =   0
            Left            =   60
            TabIndex        =   110
            Top             =   60
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   3096
            _StockProps     =   15
            Caption         =   "AUTO ADDRESS"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinItem 
            Height          =   735
            Index           =   1
            Left            =   2220
            TabIndex        =   111
            Top             =   60
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "VERIFICATION"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinAddr 
            Height          =   735
            Index           =   0
            Left            =   4860
            TabIndex        =   112
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinAddr 
            Height          =   735
            Index           =   1
            Left            =   6120
            TabIndex        =   113
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinAddr 
            Height          =   735
            Index           =   2
            Left            =   7380
            TabIndex        =   114
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinAddr 
            Height          =   735
            Index           =   3
            Left            =   8640
            TabIndex        =   115
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinItem 
            Height          =   975
            Index           =   2
            Left            =   2220
            TabIndex        =   116
            Top             =   840
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1720
            _StockProps     =   15
            Caption         =   "RESULT"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinAddrResult 
            Height          =   975
            Left            =   4860
            TabIndex        =   117
            Top             =   840
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   1720
            _StockProps     =   15
            Caption         =   "VERIFY ADDRESS"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel pnlLinFrame 
         Height          =   8895
         Index           =   2
         Left            =   -74580
         TabIndex        =   46
         Top             =   1380
         Width           =   18855
         _Version        =   65536
         _ExtentX        =   33258
         _ExtentY        =   15690
         _StockProps     =   15
         BackColor       =   15790320
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
         Begin Threed.SSPanel pnlManual 
            Height          =   1095
            Index           =   16
            Left            =   4860
            TabIndex        =   47
            Top             =   60
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1931
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Begin BHButton.BHImageButton btnDI 
               Height          =   615
               Index           =   37
               Left            =   60
               TabIndex        =   48
               Top             =   60
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   1085
               Caption         =   "LIMIT S/W"
               CaptionChecked  =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "³ª´®°íµñ"
                  Size            =   15.75
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
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Caption         =   "MOVING STEP"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   15.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   49
               Top             =   720
               Width           =   2355
            End
         End
         Begin Threed.SSPanel pnlManual 
            Height          =   1875
            Index           =   14
            Left            =   60
            TabIndex        =   52
            Top             =   6960
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   3307
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Begin VB.CommandButton btnLinActManual 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               Caption         =   "ACT 04 CCW"
               Height          =   915
               Index           =   6
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   0
               Width           =   1995
            End
            Begin VB.CommandButton btnLinActManual 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               Caption         =   "ACT 04 CW"
               Height          =   915
               Index           =   7
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   960
               Width           =   1995
            End
         End
         Begin Threed.SSPanel pnlManual 
            Height          =   1875
            Index           =   13
            Left            =   60
            TabIndex        =   55
            Top             =   5040
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   3307
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Begin VB.CommandButton btnLinActManual 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               Caption         =   "ACT 03 CCW"
               Height          =   915
               Index           =   4
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   57
               Top             =   0
               Width           =   1995
            End
            Begin VB.CommandButton btnLinActManual 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               Caption         =   "ACT 03 CW"
               Height          =   915
               Index           =   5
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   960
               Width           =   1995
            End
         End
         Begin Threed.SSPanel pnlManual 
            Height          =   1875
            Index           =   11
            Left            =   60
            TabIndex        =   58
            Top             =   3120
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   3307
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Begin VB.CommandButton btnLinActManual 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               Caption         =   "ACT 02 CCW"
               Height          =   915
               Index           =   2
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   60
               Top             =   0
               Width           =   1995
            End
            Begin VB.CommandButton btnLinActManual 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               Caption         =   "ACT 02 CW"
               Height          =   915
               Index           =   3
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   59
               Top             =   960
               Width           =   1995
            End
         End
         Begin Threed.SSPanel pnlLinActCurr 
            Height          =   1875
            Index           =   0
            Left            =   10140
            TabIndex        =   66
            Top             =   1200
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActCurr 
            Height          =   1875
            Index           =   3
            Left            =   10140
            TabIndex        =   67
            Top             =   6960
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActCurr 
            Height          =   1875
            Index           =   2
            Left            =   10140
            TabIndex        =   68
            Top             =   5040
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActCurr 
            Height          =   1875
            Index           =   1
            Left            =   10140
            TabIndex        =   69
            Top             =   3120
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinShipp 
            Height          =   1095
            Index           =   0
            Left            =   2220
            TabIndex        =   70
            Top             =   1980
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActItem 
            Height          =   1875
            Index           =   0
            Left            =   60
            TabIndex        =   71
            Top             =   1200
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "ACT 01"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActClose 
            Height          =   735
            Index           =   0
            Left            =   3540
            TabIndex        =   72
            Top             =   1200
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "CLOSE"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActOpen 
            Height          =   735
            Index           =   0
            Left            =   2220
            TabIndex        =   73
            Top             =   1200
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "OPEN"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActMove 
            Height          =   1875
            Index           =   0
            Left            =   4860
            TabIndex        =   74
            Top             =   1200
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActFinal 
            Height          =   1875
            Index           =   0
            Left            =   7500
            TabIndex        =   75
            Top             =   1200
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActResult 
            Height          =   1875
            Index           =   0
            Left            =   15420
            TabIndex        =   76
            Top             =   1200
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActTime 
            Height          =   1875
            Index           =   0
            Left            =   12780
            TabIndex        =   77
            Top             =   1200
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinShipp 
            Height          =   1095
            Index           =   1
            Left            =   2220
            TabIndex        =   78
            Top             =   3900
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActItem 
            Height          =   1875
            Index           =   1
            Left            =   60
            TabIndex        =   79
            Top             =   3120
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "ACT 02"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActClose 
            Height          =   735
            Index           =   1
            Left            =   3540
            TabIndex        =   80
            Top             =   3120
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "CLOSE"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActMove 
            Height          =   1875
            Index           =   1
            Left            =   4860
            TabIndex        =   81
            Top             =   3120
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActFinal 
            Height          =   1875
            Index           =   1
            Left            =   7500
            TabIndex        =   82
            Top             =   3120
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActResult 
            Height          =   1875
            Index           =   1
            Left            =   15420
            TabIndex        =   83
            Top             =   3120
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActTime 
            Height          =   1875
            Index           =   1
            Left            =   12780
            TabIndex        =   84
            Top             =   3120
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinShipp 
            Height          =   1095
            Index           =   2
            Left            =   2220
            TabIndex        =   85
            Top             =   5820
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActItem 
            Height          =   1875
            Index           =   2
            Left            =   60
            TabIndex        =   86
            Top             =   5040
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "ACT 03"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActClose 
            Height          =   735
            Index           =   2
            Left            =   3540
            TabIndex        =   87
            Top             =   5040
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "CLOSE"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActOpen 
            Height          =   735
            Index           =   2
            Left            =   2220
            TabIndex        =   88
            Top             =   5040
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "OPEN"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActMove 
            Height          =   1875
            Index           =   2
            Left            =   4860
            TabIndex        =   89
            Top             =   5040
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActFinal 
            Height          =   1875
            Index           =   2
            Left            =   7500
            TabIndex        =   90
            Top             =   5040
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActResult 
            Height          =   1875
            Index           =   2
            Left            =   15420
            TabIndex        =   91
            Top             =   5040
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActTime 
            Height          =   1875
            Index           =   2
            Left            =   12780
            TabIndex        =   92
            Top             =   5040
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinShipp 
            Height          =   1095
            Index           =   3
            Left            =   2220
            TabIndex        =   93
            Top             =   7740
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActItem 
            Height          =   1875
            Index           =   3
            Left            =   60
            TabIndex        =   94
            Top             =   6960
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "ACT 04"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActClose 
            Height          =   735
            Index           =   3
            Left            =   3540
            TabIndex        =   95
            Top             =   6960
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "CLOSE"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActOpen 
            Height          =   735
            Index           =   3
            Left            =   2220
            TabIndex        =   96
            Top             =   6960
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "OPEN"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActMove 
            Height          =   1875
            Index           =   3
            Left            =   4860
            TabIndex        =   97
            Top             =   6960
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActFinal 
            Height          =   1875
            Index           =   3
            Left            =   7500
            TabIndex        =   98
            Top             =   6960
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActResult 
            Height          =   1875
            Index           =   3
            Left            =   15420
            TabIndex        =   99
            Top             =   6960
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActTime 
            Height          =   1875
            Index           =   3
            Left            =   12780
            TabIndex        =   100
            Top             =   6960
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   3307
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActTitle 
            Height          =   1095
            Index           =   1
            Left            =   2220
            TabIndex        =   101
            Top             =   60
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "STALL"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActTitle 
            Height          =   1095
            Index           =   2
            Left            =   4860
            TabIndex        =   102
            Top             =   60
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "MOVING ANGLE"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActTitle 
            Height          =   1095
            Index           =   4
            Left            =   10140
            TabIndex        =   103
            Top             =   60
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "CURRENT"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActTitle 
            Height          =   1095
            Index           =   3
            Left            =   7500
            TabIndex        =   104
            Top             =   60
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "FINAL ANGLE"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActTitle 
            Height          =   1095
            Index           =   5
            Left            =   12780
            TabIndex        =   105
            Top             =   60
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "TIME"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActTitle 
            Height          =   1095
            Index           =   6
            Left            =   15420
            TabIndex        =   106
            Top             =   60
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "FINAL"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActTitle 
            Height          =   1095
            Index           =   0
            Left            =   60
            TabIndex        =   107
            Top             =   60
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "NAME"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlLinActOpen 
            Height          =   735
            Index           =   1
            Left            =   2220
            TabIndex        =   108
            Top             =   3120
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "OPEN"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   15.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlManual 
            Height          =   1095
            Index           =   10
            Left            =   60
            TabIndex        =   64
            Top             =   60
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   1931
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Begin VB.CheckBox chkLinActPower 
               BackColor       =   &H000080FF&
               Caption         =   "POWER"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   21.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   65
               Top             =   0
               Width           =   1995
            End
         End
         Begin Threed.SSPanel pnlManual 
            Height          =   1095
            Index           =   12
            Left            =   2220
            TabIndex        =   50
            Top             =   60
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1931
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Begin VB.CommandButton btnLinInit 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "LIN INIT"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   18
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   120
               Width           =   2355
            End
         End
         Begin Threed.SSPanel pnlManual 
            Height          =   1875
            Index           =   15
            Left            =   60
            TabIndex        =   61
            Top             =   1200
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   3307
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Begin VB.CommandButton btnLinActManual 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               Caption         =   "ACT 01 CW"
               Height          =   915
               Index           =   1
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   63
               Top             =   960
               Width           =   1995
            End
            Begin VB.CommandButton btnLinActManual 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               Caption         =   "ACT 01 CCW"
               Height          =   915
               Index           =   0
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   62
               Top             =   0
               Width           =   1995
            End
         End
      End
      Begin Threed.SSPanel pnlFrame 
         Height          =   5835
         Index           =   9
         Left            =   -74760
         TabIndex        =   121
         Top             =   1080
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   10292
         _StockProps     =   15
         BackColor       =   15790320
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
         Begin Threed.SSPanel pnlVibName 
            Height          =   1575
            Left            =   60
            TabIndex        =   122
            Top             =   60
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   2778
            _StockProps     =   15
            Caption         =   "VIBRATION"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlVibCurr 
            Height          =   1995
            Left            =   60
            TabIndex        =   123
            Top             =   1680
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   3519
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   27.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlVibResult 
            Height          =   2055
            Left            =   60
            TabIndex        =   124
            Top             =   3720
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   3625
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   27.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel pnlFrame 
         Height          =   4935
         Index           =   6
         Left            =   -70980
         TabIndex        =   125
         Top             =   1020
         Width           =   18855
         _Version        =   65536
         _ExtentX        =   33258
         _ExtentY        =   8705
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
         Begin VB.PictureBox picVib 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   4575
            Left            =   60
            ScaleHeight     =   4545
            ScaleWidth      =   18705
            TabIndex        =   126
            Top             =   300
            Width           =   18735
            Begin VB.Label lblVibSubject 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "VIBRATION"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   48
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0C0C0&
               Height          =   1110
               Left            =   7290
               TabIndex        =   127
               Top             =   1620
               Width           =   4995
            End
         End
         Begin VB.Label lblGraphTemp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "G"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   129
            Top             =   60
            Width           =   105
         End
         Begin VB.Label lblAxis 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   18360
            TabIndex        =   128
            Top             =   60
            Width           =   390
         End
      End
      Begin Threed.SSPanel pnlModelFullName 
         Height          =   735
         Left            =   -74400
         TabIndex        =   136
         Top             =   3120
         Width           =   6975
         _Version        =   65536
         _ExtentX        =   12303
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "ABCDEFG"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   24
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel fraNGCount 
         Height          =   1095
         Left            =   -73980
         TabIndex        =   137
         Top             =   5400
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   1931
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
         Begin VB.CommandButton btnNGCountReset 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "¿¬¼Ó ºÒ·® ¼ö·® ¸®¼Â"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   11.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            MaskColor       =   &H00C0C0FF&
            Style           =   1  'Graphical
            TabIndex        =   138
            Top             =   120
            Width           =   1215
         End
         Begin Threed.SSPanel pnlNGCount 
            Height          =   435
            Left            =   1500
            TabIndex        =   139
            Top             =   540
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   767
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   12640511
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
         End
         Begin Threed.SSPanel pnlNGLabel 
            Height          =   315
            Index           =   1
            Left            =   3120
            TabIndex        =   140
            Top             =   120
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "ºÒ·® ¼ö·® È®ÀÎ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel pnlNGLabel 
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   141
            Top             =   120
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "¿¬¼Ó ºÒ·® ¼ö·®"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel pnlNGCountValue 
            Height          =   435
            Left            =   3120
            TabIndex        =   142
            Top             =   540
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   767
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   8454143
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
         End
      End
      Begin Threed.SSPanel pnlVisionDoorCloseNum 
         Height          =   675
         Index           =   0
         Left            =   -71760
         TabIndex        =   143
         Top             =   2040
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "DOOR 1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorOpenNum 
         Height          =   675
         Index           =   0
         Left            =   -72960
         TabIndex        =   144
         Top             =   2040
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "DOOR 1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorClose 
         Height          =   675
         Index           =   0
         Left            =   -71760
         TabIndex        =   145
         Top             =   2700
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorOpen 
         Height          =   675
         Index           =   0
         Left            =   -72960
         TabIndex        =   146
         Top             =   2700
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionClose 
         Height          =   615
         Index           =   0
         Left            =   -71760
         TabIndex        =   147
         Top             =   1440
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "CLOSE"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorCloseNum 
         Height          =   675
         Index           =   1
         Left            =   -69360
         TabIndex        =   148
         Top             =   2040
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "DOOR 1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorOpenNum 
         Height          =   675
         Index           =   1
         Left            =   -70560
         TabIndex        =   149
         Top             =   2040
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "DOOR 1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorClose 
         Height          =   675
         Index           =   1
         Left            =   -69360
         TabIndex        =   150
         Top             =   2700
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorOpen 
         Height          =   675
         Index           =   1
         Left            =   -70560
         TabIndex        =   151
         Top             =   2700
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionOpen 
         Height          =   615
         Index           =   1
         Left            =   -70560
         TabIndex        =   152
         Top             =   1440
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "OPEN"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionClose 
         Height          =   615
         Index           =   1
         Left            =   -69360
         TabIndex        =   153
         Top             =   1440
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "CLOSE"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorCloseNum 
         Height          =   675
         Index           =   2
         Left            =   -66960
         TabIndex        =   154
         Top             =   2040
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "DOOR 1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorOpenNum 
         Height          =   675
         Index           =   2
         Left            =   -68160
         TabIndex        =   155
         Top             =   2040
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "DOOR 1"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorClose 
         Height          =   675
         Index           =   2
         Left            =   -66960
         TabIndex        =   156
         Top             =   2700
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionDoorOpen 
         Height          =   675
         Index           =   2
         Left            =   -68160
         TabIndex        =   157
         Top             =   2700
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionOpen 
         Height          =   615
         Index           =   2
         Left            =   -68160
         TabIndex        =   158
         Top             =   1440
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "OPEN"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionClose 
         Height          =   615
         Index           =   2
         Left            =   -66960
         TabIndex        =   159
         Top             =   1440
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "CLOSE"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlVisionOpen 
         Height          =   615
         Index           =   0
         Left            =   -72960
         TabIndex        =   160
         Top             =   1440
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "OPEN"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlFrame 
         Height          =   1575
         Index           =   8
         Left            =   -74100
         TabIndex        =   332
         Top             =   4860
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
         _ExtentY        =   2778
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
         Begin Threed.SSPanel pnlManual 
            Height          =   1095
            Index           =   9
            Left            =   60
            TabIndex        =   333
            Top             =   420
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   1931
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
            Begin VB.CommandButton btnManualVisionRun 
               BackColor       =   &H000080FF&
               Caption         =   "VISION"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   24
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   334
               Top             =   0
               Width           =   2895
            End
         End
         Begin Threed.SSPanel pnlVisionResult 
            Height          =   1095
            Left            =   3120
            TabIndex        =   335
            Top             =   420
            Width           =   4695
            _Version        =   65536
            _ExtentX        =   8281
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   21.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlVisionName 
            Height          =   1095
            Left            =   60
            TabIndex        =   336
            Top             =   420
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   1931
            _StockProps     =   15
            Caption         =   "VISION"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   24
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlFrameTitle 
            Height          =   315
            Index           =   9
            Left            =   120
            TabIndex        =   337
            Top             =   60
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "VISION"
            ForeColor       =   0
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
            Alignment       =   1
         End
      End
      Begin Threed.SSPanel pnlFrame 
         Height          =   1935
         Index           =   15
         Left            =   -73320
         TabIndex        =   338
         Top             =   4320
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
         _ExtentY        =   3413
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
         Begin Threed.SSPanel pnlFrameTitle 
            Height          =   315
            Index           =   10
            Left            =   120
            TabIndex        =   339
            Top             =   60
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "OTHER"
            ForeColor       =   0
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
            Alignment       =   1
         End
         Begin BHButton.BHImageButton btnPartProduct 
            Height          =   735
            Left            =   60
            TabIndex        =   340
            Top             =   1140
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   1296
            Caption         =   "PART 1"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TransparentPicture=   "frmRun.frx":0D3E
            Gap             =   1
            ButtonAttrib    =   2
            ImgOutLineSize  =   1
            ImgUpOutLineSize=   1
            ImgDownOutLineSize=   1
            ImgDisableOutLineSize=   1
         End
         Begin Threed.SSPanel pnlScannerResult 
            Height          =   675
            Left            =   60
            TabIndex        =   341
            Top             =   420
            Width           =   7755
            _Version        =   65536
            _ExtentX        =   13679
            _ExtentY        =   1191
            _StockProps     =   15
            Caption         =   "BARCODE SCAN"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ ExtraBold"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlDoorResult 
            Height          =   735
            Left            =   3960
            TabIndex        =   342
            Top             =   1140
            Width           =   3855
            _Version        =   65536
            _ExtentX        =   6800
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "CHECK DOOR"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
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
         Height          =   2535
         Index           =   10
         Left            =   -74220
         TabIndex        =   343
         Top             =   2340
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
         _ExtentY        =   4471
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
         Begin Threed.SSPanel pnlManual 
            Height          =   2055
            Index           =   5
            Left            =   60
            TabIndex        =   344
            Top             =   420
            Width           =   5295
            _Version        =   65536
            _ExtentX        =   9340
            _ExtentY        =   3625
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   12.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Begin VB.CheckBox chkIonManual 
               BackColor       =   &H000080FF&
               Caption         =   "POWER"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   20.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2055
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   347
               Top             =   0
               Width           =   2895
            End
            Begin VB.OptionButton optIonManual 
               BackColor       =   &H00FF80FF&
               Caption         =   "DIAG OFF"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   20.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Index           =   0
               Left            =   3060
               Style           =   1  'Graphical
               TabIndex        =   346
               Top             =   0
               Width           =   2295
            End
            Begin VB.OptionButton optIonManual 
               BackColor       =   &H00FF80FF&
               Caption         =   "DIAG ON"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   20.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1035
               Index           =   1
               Left            =   3060
               Style           =   1  'Graphical
               TabIndex        =   345
               Top             =   1020
               Width           =   2295
            End
         End
         Begin Threed.SSPanel pnlIonName 
            Height          =   1035
            Index           =   2
            Left            =   3120
            TabIndex        =   348
            Top             =   1440
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
            _ExtentY        =   1826
            _StockProps     =   15
            Caption         =   "DIAG ON"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlIonData 
            Height          =   1035
            Index           =   1
            Left            =   5400
            TabIndex        =   349
            Top             =   1440
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   1826
            _StockProps     =   15
            Caption         =   "0.0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   24
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlIonData 
            Height          =   975
            Index           =   0
            Left            =   5400
            TabIndex        =   350
            Top             =   420
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   1720
            _StockProps     =   15
            Caption         =   "0.0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   24
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlIonName 
            Height          =   975
            Index           =   1
            Left            =   3120
            TabIndex        =   351
            Top             =   420
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
            _ExtentY        =   1720
            _StockProps     =   15
            Caption         =   "DIAG OFF"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlFrameTitle 
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   352
            Top             =   60
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "IONIZER"
            ForeColor       =   0
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnlIonName 
            Height          =   2055
            Index           =   0
            Left            =   60
            TabIndex        =   353
            Top             =   420
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   3625
            _StockProps     =   15
            Caption         =   "IONIZER"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   24
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel pnlNvhRpm 
         Height          =   555
         Left            =   -70920
         TabIndex        =   411
         Top             =   6360
         Width           =   3915
         _Version        =   65536
         _ExtentX        =   6906
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin BHButton.BHImageButton btnBarcodePrint 
         Height          =   615
         Left            =   -72660
         TabIndex        =   412
         Top             =   2700
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "BARCODE PRINT"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
      Begin Threed.SSPanel pnlFrame 
         Height          =   10755
         Index           =   5
         Left            =   -72960
         TabIndex        =   413
         Top             =   4020
         Width           =   6015
         _Version        =   65536
         _ExtentX        =   10610
         _ExtentY        =   18971
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
         Begin Threed.SSPanel pnlManual 
            Height          =   9075
            Index           =   7
            Left            =   60
            TabIndex        =   414
            Top             =   420
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   16007
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Begin VB.OptionButton optBlowerManual 
               BackColor       =   &H00FF80FF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Index           =   0
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   424
               Top             =   1020
               Width           =   1515
            End
            Begin VB.CheckBox chkBlowerPower 
               BackColor       =   &H000080FF&
               Caption         =   "POWER"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   423
               Top             =   0
               Width           =   1515
            End
            Begin VB.OptionButton optBlowerManual 
               BackColor       =   &H00FF80FF&
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Index           =   1
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   422
               Top             =   1920
               Width           =   1515
            End
            Begin VB.OptionButton optBlowerManual 
               BackColor       =   &H00FF80FF&
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Index           =   2
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   421
               Top             =   2820
               Width           =   1515
            End
            Begin VB.OptionButton optBlowerManual 
               BackColor       =   &H00FF80FF&
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Index           =   3
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   420
               Top             =   3720
               Width           =   1515
            End
            Begin VB.OptionButton optBlowerManual 
               BackColor       =   &H00FF80FF&
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Index           =   4
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   419
               Top             =   4620
               Width           =   1515
            End
            Begin VB.OptionButton optBlowerManual 
               BackColor       =   &H00FF80FF&
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Index           =   5
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   418
               Top             =   5520
               Width           =   1515
            End
            Begin VB.OptionButton optBlowerManual 
               BackColor       =   &H00FF80FF&
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Index           =   6
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   417
               Top             =   6420
               Width           =   1515
            End
            Begin VB.OptionButton optBlowerManual 
               BackColor       =   &H00FF80FF&
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Index           =   7
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   416
               Top             =   7320
               Width           =   1515
            End
            Begin VB.OptionButton optBlowerManual 
               BackColor       =   &H00FF80FF&
               Caption         =   "8"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Index           =   8
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   415
               Top             =   8220
               Width           =   1515
            End
         End
         Begin Threed.SSPanel pnlBlowerItem 
            Height          =   975
            Index           =   1
            Left            =   1740
            TabIndex        =   425
            Top             =   420
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1720
            _StockProps     =   15
            Caption         =   "CURR (A)"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerItem 
            Height          =   975
            Index           =   2
            Left            =   3180
            TabIndex        =   426
            Top             =   420
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   1720
            _StockProps     =   15
            Caption         =   "TIME (Sec)"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerItem 
            Height          =   975
            Index           =   3
            Left            =   4380
            TabIndex        =   427
            Top             =   420
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   1720
            _StockProps     =   15
            Caption         =   "FINAL"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerTime 
            Height          =   855
            Index           =   0
            Left            =   3180
            TabIndex        =   428
            Top             =   1440
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlRpmName 
            Height          =   1155
            Left            =   60
            TabIndex        =   429
            Top             =   9540
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   2037
            _StockProps     =   15
            Caption         =   "RPM"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerName 
            Height          =   855
            Index           =   0
            Left            =   60
            TabIndex        =   430
            Top             =   1440
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerResult 
            Height          =   855
            Index           =   0
            Left            =   4380
            TabIndex        =   431
            Top             =   1440
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlRpmCurr 
            Height          =   1155
            Left            =   1740
            TabIndex        =   432
            Top             =   9540
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   2037
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlRpmResult 
            Height          =   1155
            Left            =   4380
            TabIndex        =   433
            Top             =   9540
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   2037
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerCurr 
            Height          =   855
            Index           =   0
            Left            =   1740
            TabIndex        =   434
            Top             =   1440
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerItem 
            Height          =   975
            Index           =   0
            Left            =   60
            TabIndex        =   435
            Top             =   420
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   1720
            _StockProps     =   15
            Caption         =   "BLOWER"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlFrameTitle 
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   436
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "BLOWER"
            ForeColor       =   0
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnlBlowerTime 
            Height          =   855
            Index           =   1
            Left            =   3180
            TabIndex        =   437
            Top             =   2340
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerName 
            Height          =   855
            Index           =   1
            Left            =   60
            TabIndex        =   438
            Top             =   2340
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "1"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerResult 
            Height          =   855
            Index           =   1
            Left            =   4380
            TabIndex        =   439
            Top             =   2340
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerCurr 
            Height          =   855
            Index           =   1
            Left            =   1740
            TabIndex        =   440
            Top             =   2340
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerTime 
            Height          =   855
            Index           =   2
            Left            =   3180
            TabIndex        =   441
            Top             =   3240
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerName 
            Height          =   855
            Index           =   2
            Left            =   60
            TabIndex        =   442
            Top             =   3240
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "2"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerResult 
            Height          =   855
            Index           =   2
            Left            =   4380
            TabIndex        =   443
            Top             =   3240
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerCurr 
            Height          =   855
            Index           =   2
            Left            =   1740
            TabIndex        =   444
            Top             =   3240
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerTime 
            Height          =   855
            Index           =   3
            Left            =   3180
            TabIndex        =   445
            Top             =   4140
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerName 
            Height          =   855
            Index           =   3
            Left            =   60
            TabIndex        =   446
            Top             =   4140
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "3"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerResult 
            Height          =   855
            Index           =   3
            Left            =   4380
            TabIndex        =   447
            Top             =   4140
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerCurr 
            Height          =   855
            Index           =   3
            Left            =   1740
            TabIndex        =   448
            Top             =   4140
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerTime 
            Height          =   855
            Index           =   4
            Left            =   3180
            TabIndex        =   449
            Top             =   5040
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerName 
            Height          =   855
            Index           =   4
            Left            =   60
            TabIndex        =   450
            Top             =   5040
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "4"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerResult 
            Height          =   855
            Index           =   4
            Left            =   4380
            TabIndex        =   451
            Top             =   5040
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerCurr 
            Height          =   855
            Index           =   4
            Left            =   1740
            TabIndex        =   452
            Top             =   5040
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerTime 
            Height          =   855
            Index           =   5
            Left            =   3180
            TabIndex        =   453
            Top             =   5940
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerName 
            Height          =   855
            Index           =   5
            Left            =   60
            TabIndex        =   454
            Top             =   5940
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "5"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerResult 
            Height          =   855
            Index           =   5
            Left            =   4380
            TabIndex        =   455
            Top             =   5940
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerCurr 
            Height          =   855
            Index           =   5
            Left            =   1740
            TabIndex        =   456
            Top             =   5940
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerTime 
            Height          =   855
            Index           =   6
            Left            =   3180
            TabIndex        =   457
            Top             =   6840
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerName 
            Height          =   855
            Index           =   6
            Left            =   60
            TabIndex        =   458
            Top             =   6840
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "6"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerResult 
            Height          =   855
            Index           =   6
            Left            =   4380
            TabIndex        =   459
            Top             =   6840
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerCurr 
            Height          =   855
            Index           =   6
            Left            =   1740
            TabIndex        =   460
            Top             =   6840
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerTime 
            Height          =   855
            Index           =   7
            Left            =   3180
            TabIndex        =   461
            Top             =   7740
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerName 
            Height          =   855
            Index           =   7
            Left            =   60
            TabIndex        =   462
            Top             =   7740
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "7"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerResult 
            Height          =   855
            Index           =   7
            Left            =   4380
            TabIndex        =   463
            Top             =   7740
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerCurr 
            Height          =   855
            Index           =   7
            Left            =   1740
            TabIndex        =   464
            Top             =   7740
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerTime 
            Height          =   855
            Index           =   8
            Left            =   3180
            TabIndex        =   465
            Top             =   8640
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerName 
            Height          =   855
            Index           =   8
            Left            =   60
            TabIndex        =   466
            Top             =   8640
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "8"
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerResult 
            Height          =   855
            Index           =   8
            Left            =   4380
            TabIndex        =   467
            Top             =   8640
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "OK"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlBlowerCurr 
            Height          =   855
            Index           =   8
            Left            =   1740
            TabIndex        =   468
            Top             =   8640
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1508
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
   End
   Begin Threed.SSPanel pnlFrame 
      Height          =   5115
      Index           =   1
      Left            =   6420
      TabIndex        =   256
      Top             =   9720
      Width           =   6315
      _Version        =   65536
      _ExtentX        =   11139
      _ExtentY        =   9022
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
      Begin Threed.SSPanel pnlAct03StallDelta 
         Height          =   915
         Left            =   1200
         TabIndex        =   479
         Top             =   4140
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlManual 
         Height          =   795
         Index           =   1
         Left            =   60
         TabIndex        =   257
         Top             =   420
         Width           =   5175
         _Version        =   65536
         _ExtentX        =   9128
         _ExtentY        =   1402
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   8.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Begin VB.OptionButton optAct03Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P4"
            Height          =   795
            Index           =   3
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   289
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optAct03Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P3"
            Height          =   795
            Index           =   2
            Left            =   3180
            Style           =   1  'Graphical
            TabIndex        =   288
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optAct03Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P2"
            Height          =   795
            Index           =   1
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   287
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optAct03Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P1"
            Height          =   795
            Index           =   0
            Left            =   1140
            Style           =   1  'Graphical
            TabIndex        =   286
            Top             =   0
            Width           =   975
         End
         Begin VB.CheckBox chkActPower 
            BackColor       =   &H000080FF&
            Caption         =   "POWER"
            Height          =   795
            Index           =   2
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   258
            Top             =   0
            Width           =   975
         End
      End
      Begin Threed.SSPanel pnlFrameTitle 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   260
         Top             =   60
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "ACT 03"
         ForeColor       =   0
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
         Alignment       =   1
      End
      Begin Threed.SSPanel pnlAct03Item 
         Height          =   795
         Index           =   0
         Left            =   60
         TabIndex        =   261
         Top             =   420
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "TEMP2"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Item 
         Height          =   915
         Index           =   1
         Left            =   60
         TabIndex        =   262
         Top             =   1260
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "CURR (mA)"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Item 
         Height          =   915
         Index           =   2
         Left            =   60
         TabIndex        =   263
         Top             =   2220
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "F/B (V)"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Item 
         Height          =   915
         Index           =   3
         Left            =   60
         TabIndex        =   264
         Top             =   3180
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "TIME (Sec)"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Item 
         Height          =   915
         Index           =   4
         Left            =   60
         TabIndex        =   265
         Top             =   4140
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "FINAL"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Name 
         Height          =   795
         Index           =   0
         Left            =   1200
         TabIndex        =   266
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "P1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Name 
         Height          =   795
         Index           =   1
         Left            =   2220
         TabIndex        =   267
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "P2"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Name 
         Height          =   795
         Index           =   2
         Left            =   3240
         TabIndex        =   268
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "P3"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Name 
         Height          =   795
         Index           =   3
         Left            =   4260
         TabIndex        =   269
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "P4"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Curr 
         Height          =   915
         Index           =   0
         Left            =   1200
         TabIndex        =   270
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Volt 
         Height          =   915
         Index           =   0
         Left            =   1200
         TabIndex        =   271
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Time 
         Height          =   915
         Index           =   0
         Left            =   1200
         TabIndex        =   272
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Result 
         Height          =   915
         Index           =   0
         Left            =   1200
         TabIndex        =   273
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Curr 
         Height          =   915
         Index           =   1
         Left            =   2220
         TabIndex        =   274
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Volt 
         Height          =   915
         Index           =   1
         Left            =   2220
         TabIndex        =   275
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Time 
         Height          =   915
         Index           =   1
         Left            =   2220
         TabIndex        =   276
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Result 
         Height          =   915
         Index           =   1
         Left            =   2220
         TabIndex        =   277
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Curr 
         Height          =   915
         Index           =   2
         Left            =   3240
         TabIndex        =   278
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Volt 
         Height          =   915
         Index           =   2
         Left            =   3240
         TabIndex        =   279
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Time 
         Height          =   915
         Index           =   2
         Left            =   3240
         TabIndex        =   280
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Result 
         Height          =   915
         Index           =   2
         Left            =   3240
         TabIndex        =   281
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Curr 
         Height          =   915
         Index           =   3
         Left            =   4260
         TabIndex        =   282
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Volt 
         Height          =   915
         Index           =   3
         Left            =   4260
         TabIndex        =   283
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Time 
         Height          =   915
         Index           =   3
         Left            =   4260
         TabIndex        =   284
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Result 
         Height          =   915
         Index           =   3
         Left            =   4260
         TabIndex        =   285
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Name 
         Height          =   795
         Index           =   4
         Left            =   5280
         TabIndex        =   370
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "FINAL"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Curr 
         Height          =   915
         Index           =   4
         Left            =   5280
         TabIndex        =   371
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Volt 
         Height          =   915
         Index           =   4
         Left            =   5280
         TabIndex        =   372
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Time 
         Height          =   915
         Index           =   4
         Left            =   5280
         TabIndex        =   373
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct03Result 
         Height          =   915
         Index           =   4
         Left            =   5280
         TabIndex        =   374
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
   End
   Begin HIRESTIMERLib.HiResTimer tmrSteppingSend 
      Left            =   540
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   64
      Interval        =   160
   End
   Begin Threed.SSPanel pnlFrame 
      Height          =   5115
      Index           =   2
      Left            =   12720
      TabIndex        =   290
      Top             =   9720
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   9022
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
      Begin Threed.SSPanel pnlAct04StallDelta 
         Height          =   915
         Left            =   1140
         TabIndex        =   480
         Top             =   4140
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlManual 
         Height          =   795
         Index           =   4
         Left            =   60
         TabIndex        =   292
         Top             =   420
         Width           =   5115
         _Version        =   65536
         _ExtentX        =   9022
         _ExtentY        =   1402
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   8.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Begin VB.OptionButton optAct04Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P4"
            Height          =   795
            Index           =   3
            Left            =   4140
            Style           =   1  'Graphical
            TabIndex        =   322
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optAct04Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P3"
            Height          =   795
            Index           =   2
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   321
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optAct04Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P2"
            Height          =   795
            Index           =   1
            Left            =   2100
            Style           =   1  'Graphical
            TabIndex        =   320
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optAct04Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P1"
            Height          =   795
            Index           =   0
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   319
            Top             =   0
            Width           =   975
         End
         Begin VB.CheckBox chkActPower 
            BackColor       =   &H000080FF&
            Caption         =   "POWER"
            Height          =   795
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   293
            Top             =   0
            Width           =   915
         End
      End
      Begin Threed.SSPanel pnlFrameTitle 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   291
         Top             =   60
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "ACT 04"
         ForeColor       =   0
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
         Alignment       =   1
      End
      Begin Threed.SSPanel pnlAct04Item 
         Height          =   795
         Index           =   0
         Left            =   60
         TabIndex        =   294
         Top             =   420
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "INTAKE"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Item 
         Height          =   915
         Index           =   1
         Left            =   60
         TabIndex        =   295
         Top             =   1260
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "CURR (mA)"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Item 
         Height          =   915
         Index           =   2
         Left            =   60
         TabIndex        =   296
         Top             =   2220
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "F/B (V)"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Item 
         Height          =   915
         Index           =   3
         Left            =   60
         TabIndex        =   297
         Top             =   3180
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "TIME (Sec)"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Item 
         Height          =   915
         Index           =   4
         Left            =   60
         TabIndex        =   298
         Top             =   4140
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "FINAL"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Curr 
         Height          =   915
         Index           =   0
         Left            =   1140
         TabIndex        =   299
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Volt 
         Height          =   915
         Index           =   0
         Left            =   1140
         TabIndex        =   300
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Time 
         Height          =   915
         Index           =   0
         Left            =   1140
         TabIndex        =   301
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Result 
         Height          =   915
         Index           =   0
         Left            =   1140
         TabIndex        =   302
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Name 
         Height          =   795
         Index           =   0
         Left            =   1140
         TabIndex        =   303
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "P1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Curr 
         Height          =   915
         Index           =   1
         Left            =   2160
         TabIndex        =   304
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Volt 
         Height          =   915
         Index           =   1
         Left            =   2160
         TabIndex        =   305
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Time 
         Height          =   915
         Index           =   1
         Left            =   2160
         TabIndex        =   306
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Result 
         Height          =   915
         Index           =   1
         Left            =   2160
         TabIndex        =   307
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Name 
         Height          =   795
         Index           =   1
         Left            =   2160
         TabIndex        =   308
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "P2"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Curr 
         Height          =   915
         Index           =   2
         Left            =   3180
         TabIndex        =   309
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Volt 
         Height          =   915
         Index           =   2
         Left            =   3180
         TabIndex        =   310
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Time 
         Height          =   915
         Index           =   2
         Left            =   3180
         TabIndex        =   311
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Result 
         Height          =   915
         Index           =   2
         Left            =   3180
         TabIndex        =   312
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Name 
         Height          =   795
         Index           =   2
         Left            =   3180
         TabIndex        =   313
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "P3"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Curr 
         Height          =   915
         Index           =   3
         Left            =   4200
         TabIndex        =   314
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Volt 
         Height          =   915
         Index           =   3
         Left            =   4200
         TabIndex        =   315
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Time 
         Height          =   915
         Index           =   3
         Left            =   4200
         TabIndex        =   316
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Result 
         Height          =   915
         Index           =   3
         Left            =   4200
         TabIndex        =   317
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Name 
         Height          =   795
         Index           =   3
         Left            =   4200
         TabIndex        =   318
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "P4"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Curr 
         Height          =   915
         Index           =   4
         Left            =   5220
         TabIndex        =   365
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Volt 
         Height          =   915
         Index           =   4
         Left            =   5220
         TabIndex        =   366
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Time 
         Height          =   915
         Index           =   4
         Left            =   5220
         TabIndex        =   367
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Result 
         Height          =   915
         Index           =   4
         Left            =   5220
         TabIndex        =   368
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct04Name 
         Height          =   795
         Index           =   4
         Left            =   5220
         TabIndex        =   369
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "FINAL"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlFrame 
      Height          =   5115
      Index           =   4
      Left            =   120
      TabIndex        =   222
      Top             =   9720
      Width           =   6315
      _Version        =   65536
      _ExtentX        =   11139
      _ExtentY        =   9022
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
      Begin Threed.SSPanel pnlAct02StallDelta 
         Height          =   915
         Left            =   1200
         TabIndex        =   478
         Top             =   4140
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlManual 
         Height          =   795
         Index           =   6
         Left            =   60
         TabIndex        =   223
         Top             =   420
         Width           =   5175
         _Version        =   65536
         _ExtentX        =   9128
         _ExtentY        =   1402
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   8.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Begin VB.OptionButton optAct02Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P4"
            Height          =   795
            Index           =   3
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   255
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optAct02Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P3"
            Height          =   795
            Index           =   2
            Left            =   3180
            Style           =   1  'Graphical
            TabIndex        =   254
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optAct02Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P2"
            Height          =   795
            Index           =   1
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   253
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optAct02Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P1"
            Height          =   795
            Index           =   0
            Left            =   1140
            Style           =   1  'Graphical
            TabIndex        =   252
            Top             =   0
            Width           =   975
         End
         Begin VB.CheckBox chkActPower 
            BackColor       =   &H000080FF&
            Caption         =   "POWER"
            Height          =   795
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   224
            Top             =   0
            Width           =   975
         End
      End
      Begin Threed.SSPanel pnlAct02Item 
         Height          =   915
         Index           =   1
         Left            =   60
         TabIndex        =   225
         Top             =   1260
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "CURR (mA)"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Item 
         Height          =   915
         Index           =   2
         Left            =   60
         TabIndex        =   226
         Top             =   2220
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "F/B (V)"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Item 
         Height          =   915
         Index           =   3
         Left            =   60
         TabIndex        =   227
         Top             =   3180
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "TIME (Sec)"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Item 
         Height          =   915
         Index           =   4
         Left            =   60
         TabIndex        =   228
         Top             =   4140
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "FINAL"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Curr 
         Height          =   915
         Index           =   0
         Left            =   1200
         TabIndex        =   229
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Name 
         Height          =   795
         Index           =   0
         Left            =   1200
         TabIndex        =   230
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "P1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Volt 
         Height          =   915
         Index           =   0
         Left            =   1200
         TabIndex        =   231
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Time 
         Height          =   915
         Index           =   0
         Left            =   1200
         TabIndex        =   232
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Result 
         Height          =   915
         Index           =   0
         Left            =   1200
         TabIndex        =   233
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Curr 
         Height          =   915
         Index           =   1
         Left            =   2220
         TabIndex        =   235
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Name 
         Height          =   795
         Index           =   1
         Left            =   2220
         TabIndex        =   236
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "P2"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Volt 
         Height          =   915
         Index           =   1
         Left            =   2220
         TabIndex        =   237
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Time 
         Height          =   915
         Index           =   1
         Left            =   2220
         TabIndex        =   238
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Result 
         Height          =   915
         Index           =   1
         Left            =   2220
         TabIndex        =   239
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Curr 
         Height          =   915
         Index           =   2
         Left            =   3240
         TabIndex        =   240
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Name 
         Height          =   795
         Index           =   2
         Left            =   3240
         TabIndex        =   241
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "P3"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Volt 
         Height          =   915
         Index           =   2
         Left            =   3240
         TabIndex        =   242
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Time 
         Height          =   915
         Index           =   2
         Left            =   3240
         TabIndex        =   243
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Result 
         Height          =   915
         Index           =   2
         Left            =   3240
         TabIndex        =   244
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Curr 
         Height          =   915
         Index           =   3
         Left            =   4260
         TabIndex        =   245
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Name 
         Height          =   795
         Index           =   3
         Left            =   4260
         TabIndex        =   246
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "P4"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Volt 
         Height          =   915
         Index           =   3
         Left            =   4260
         TabIndex        =   247
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Time 
         Height          =   915
         Index           =   3
         Left            =   4260
         TabIndex        =   248
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Result 
         Height          =   915
         Index           =   3
         Left            =   4260
         TabIndex        =   249
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlFrameTitle 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   259
         Top             =   60
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "ACT 02"
         ForeColor       =   0
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
         Alignment       =   1
      End
      Begin Threed.SSPanel pnlAct02Item 
         Height          =   795
         Index           =   0
         Left            =   60
         TabIndex        =   234
         Top             =   420
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "TEMP1"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Curr 
         Height          =   915
         Index           =   4
         Left            =   5280
         TabIndex        =   360
         Top             =   1260
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Name 
         Height          =   795
         Index           =   4
         Left            =   5280
         TabIndex        =   361
         Top             =   420
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "FINAL"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Volt 
         Height          =   915
         Index           =   4
         Left            =   5280
         TabIndex        =   362
         Top             =   2220
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Time 
         Height          =   915
         Index           =   4
         Left            =   5280
         TabIndex        =   363
         Top             =   3180
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct02Result 
         Height          =   915
         Index           =   4
         Left            =   5280
         TabIndex        =   364
         Top             =   4140
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1614
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlFrame 
      Height          =   5655
      Index           =   0
      Left            =   120
      TabIndex        =   173
      Top             =   4080
      Width           =   13575
      _Version        =   65536
      _ExtentX        =   23945
      _ExtentY        =   9975
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
      Begin Threed.SSPanel pnlAct01StallDelta 
         Height          =   1035
         Left            =   2040
         TabIndex        =   475
         Top             =   4560
         Width           =   4275
         _Version        =   65536
         _ExtentX        =   7541
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Item 
         Height          =   1035
         Index           =   5
         Left            =   6360
         TabIndex        =   476
         Top             =   4560
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "TOTAL TIME"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   15.76
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01TotalTime 
         Height          =   1035
         Left            =   7800
         TabIndex        =   477
         Top             =   4560
         Width           =   5715
         _Version        =   65536
         _ExtentX        =   10081
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlManual 
         Height          =   855
         Index           =   2
         Left            =   60
         TabIndex        =   174
         Top             =   420
         Width           =   12015
         _Version        =   65536
         _ExtentX        =   21193
         _ExtentY        =   1508
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Begin VB.OptionButton optAct01Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P7"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   6
            Left            =   10620
            Style           =   1  'Graphical
            TabIndex        =   251
            Top             =   0
            Width           =   1395
         End
         Begin VB.OptionButton optAct01Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P6"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   5
            Left            =   9180
            Style           =   1  'Graphical
            TabIndex        =   250
            Top             =   0
            Width           =   1395
         End
         Begin VB.OptionButton optAct01Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P5"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   4
            Left            =   7740
            Style           =   1  'Graphical
            TabIndex        =   180
            Top             =   0
            Width           =   1395
         End
         Begin VB.OptionButton optAct01Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P4"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   3
            Left            =   6300
            Style           =   1  'Graphical
            TabIndex        =   179
            Top             =   0
            Width           =   1395
         End
         Begin VB.OptionButton optAct01Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P3"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   2
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   178
            Top             =   0
            Width           =   1395
         End
         Begin VB.OptionButton optAct01Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P2"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   1
            Left            =   3420
            Style           =   1  'Graphical
            TabIndex        =   177
            Top             =   0
            Width           =   1395
         End
         Begin VB.OptionButton optAct01Manual 
            BackColor       =   &H00FF80FF&
            Caption         =   "P1"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   1980
            Style           =   1  'Graphical
            TabIndex        =   176
            Top             =   0
            Width           =   1395
         End
         Begin VB.CheckBox chkActPower 
            BackColor       =   &H000080FF&
            Caption         =   "POWER"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   14.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   175
            Top             =   0
            Width           =   1815
         End
      End
      Begin Threed.SSPanel pnlAct01Item 
         Height          =   1035
         Index           =   1
         Left            =   60
         TabIndex        =   181
         Top             =   1320
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "CURR (mA)"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Item 
         Height          =   1035
         Index           =   2
         Left            =   60
         TabIndex        =   182
         Top             =   2400
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "F/B (V)"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Item 
         Height          =   1035
         Index           =   3
         Left            =   60
         TabIndex        =   183
         Top             =   3480
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "TIME (Sec)"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Item 
         Height          =   1035
         Index           =   4
         Left            =   60
         TabIndex        =   184
         Top             =   4560
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "FINAL"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Curr 
         Height          =   1035
         Index           =   0
         Left            =   2040
         TabIndex        =   190
         Top             =   1320
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Curr 
         Height          =   1035
         Index           =   1
         Left            =   3480
         TabIndex        =   191
         Top             =   1320
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Curr 
         Height          =   1035
         Index           =   2
         Left            =   4920
         TabIndex        =   192
         Top             =   1320
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Curr 
         Height          =   1035
         Index           =   3
         Left            =   6360
         TabIndex        =   193
         Top             =   1320
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Curr 
         Height          =   1035
         Index           =   4
         Left            =   7800
         TabIndex        =   194
         Top             =   1320
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Volt 
         Height          =   1035
         Index           =   0
         Left            =   2040
         TabIndex        =   195
         Top             =   2400
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Volt 
         Height          =   1035
         Index           =   1
         Left            =   3480
         TabIndex        =   196
         Top             =   2400
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Volt 
         Height          =   1035
         Index           =   2
         Left            =   4920
         TabIndex        =   197
         Top             =   2400
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Volt 
         Height          =   1035
         Index           =   3
         Left            =   6360
         TabIndex        =   198
         Top             =   2400
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Volt 
         Height          =   1035
         Index           =   4
         Left            =   7800
         TabIndex        =   199
         Top             =   2400
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Time 
         Height          =   1035
         Index           =   0
         Left            =   2040
         TabIndex        =   200
         Top             =   3480
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Time 
         Height          =   1035
         Index           =   1
         Left            =   3480
         TabIndex        =   201
         Top             =   3480
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Time 
         Height          =   1035
         Index           =   2
         Left            =   4920
         TabIndex        =   202
         Top             =   3480
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Time 
         Height          =   1035
         Index           =   3
         Left            =   6360
         TabIndex        =   203
         Top             =   3480
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Time 
         Height          =   1035
         Index           =   4
         Left            =   7800
         TabIndex        =   204
         Top             =   3480
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Result 
         Height          =   1035
         Index           =   0
         Left            =   2040
         TabIndex        =   205
         Top             =   4560
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Result 
         Height          =   1035
         Index           =   1
         Left            =   3480
         TabIndex        =   206
         Top             =   4560
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Result 
         Height          =   1035
         Index           =   2
         Left            =   4920
         TabIndex        =   207
         Top             =   4560
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Result 
         Height          =   1035
         Index           =   3
         Left            =   6360
         TabIndex        =   208
         Top             =   4560
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Result 
         Height          =   1035
         Index           =   4
         Left            =   7800
         TabIndex        =   209
         Top             =   4560
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlFrameTitle 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   210
         Top             =   60
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "ACT 01"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12.01
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Alignment       =   1
      End
      Begin Threed.SSPanel pnlAct01Curr 
         Height          =   1035
         Index           =   5
         Left            =   9240
         TabIndex        =   214
         Top             =   1320
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Volt 
         Height          =   1035
         Index           =   5
         Left            =   9240
         TabIndex        =   215
         Top             =   2400
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Time 
         Height          =   1035
         Index           =   5
         Left            =   9240
         TabIndex        =   216
         Top             =   3480
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Result 
         Height          =   1035
         Index           =   5
         Left            =   9240
         TabIndex        =   217
         Top             =   4560
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Curr 
         Height          =   1035
         Index           =   6
         Left            =   10680
         TabIndex        =   218
         Top             =   1320
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Volt 
         Height          =   1035
         Index           =   6
         Left            =   10680
         TabIndex        =   219
         Top             =   2400
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Time 
         Height          =   1035
         Index           =   6
         Left            =   10680
         TabIndex        =   220
         Top             =   3480
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Result 
         Height          =   1035
         Index           =   6
         Left            =   10680
         TabIndex        =   221
         Top             =   4560
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Item 
         Height          =   855
         Index           =   0
         Left            =   60
         TabIndex        =   211
         Top             =   420
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "MODE"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Name 
         Height          =   855
         Index           =   6
         Left            =   10680
         TabIndex        =   213
         Top             =   420
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "P7"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Name 
         Height          =   855
         Index           =   5
         Left            =   9240
         TabIndex        =   212
         Top             =   420
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "P6"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Name 
         Height          =   855
         Index           =   4
         Left            =   7800
         TabIndex        =   189
         Top             =   420
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "P5"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Name 
         Height          =   855
         Index           =   3
         Left            =   6360
         TabIndex        =   188
         Top             =   420
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "P4"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Name 
         Height          =   855
         Index           =   2
         Left            =   4920
         TabIndex        =   187
         Top             =   420
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "P3"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Name 
         Height          =   855
         Index           =   1
         Left            =   3480
         TabIndex        =   186
         Top             =   420
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "P2"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Name 
         Height          =   855
         Index           =   0
         Left            =   2040
         TabIndex        =   185
         Top             =   420
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "P1"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Curr 
         Height          =   1035
         Index           =   7
         Left            =   12120
         TabIndex        =   355
         Top             =   1320
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Volt 
         Height          =   1035
         Index           =   7
         Left            =   12120
         TabIndex        =   356
         Top             =   2400
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Time 
         Height          =   1035
         Index           =   7
         Left            =   12120
         TabIndex        =   357
         Top             =   3480
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Result 
         Height          =   1035
         Index           =   7
         Left            =   12120
         TabIndex        =   358
         Top             =   4560
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1826
         _StockProps     =   15
         Caption         =   "OK"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   20.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlAct01Name 
         Height          =   855
         Index           =   7
         Left            =   12120
         TabIndex        =   359
         Top             =   420
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "FINAL"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlFrame 
      Height          =   5655
      Index           =   3
      Left            =   13680
      TabIndex        =   161
      Top             =   4080
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   9975
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
      Begin Threed.SSPanel pnlManual 
         Height          =   5175
         Index           =   3
         Left            =   60
         TabIndex        =   171
         Top             =   420
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   9128
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Begin VB.CheckBox chkSensorManual 
            BackColor       =   &H000080FF&
            Caption         =   "SENSOR 6"
            Height          =   795
            Index           =   5
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   474
            Top             =   4380
            Width           =   1755
         End
         Begin VB.CheckBox chkSensorManual 
            BackColor       =   &H000080FF&
            Caption         =   "SENSOR 5"
            Height          =   795
            Index           =   4
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   473
            Top             =   3540
            Width           =   1755
         End
         Begin VB.CheckBox chkSensorManual 
            BackColor       =   &H000080FF&
            Caption         =   "SENSOR 4"
            Height          =   795
            Index           =   3
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   331
            Top             =   2700
            Width           =   1755
         End
         Begin VB.CheckBox chkSensorManual 
            BackColor       =   &H000080FF&
            Caption         =   "SENSOR 3"
            Height          =   855
            Index           =   2
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   330
            Top             =   1800
            Width           =   1755
         End
         Begin VB.CheckBox chkSensorManual 
            BackColor       =   &H000080FF&
            Caption         =   "SENSOR 2"
            Height          =   855
            Index           =   1
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   329
            Top             =   900
            Width           =   1755
         End
         Begin VB.CheckBox chkSensorManual 
            BackColor       =   &H000080FF&
            Caption         =   "SENSOR 1"
            Height          =   855
            Index           =   0
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   172
            Top             =   0
            Width           =   1755
         End
      End
      Begin Threed.SSPanel pnlFrameTitle 
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   162
         Top             =   60
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "SENSOR"
         ForeColor       =   0
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
         Alignment       =   1
      End
      Begin Threed.SSPanel pnlSensorVolt 
         Height          =   855
         Index           =   0
         Left            =   1920
         TabIndex        =   170
         Top             =   420
         Width           =   3315
         _Version        =   65536
         _ExtentX        =   5847
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSensorName 
         Height          =   855
         Index           =   1
         Left            =   60
         TabIndex        =   323
         Top             =   1320
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "SENSOR 2"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSensorVolt 
         Height          =   855
         Index           =   1
         Left            =   1920
         TabIndex        =   324
         Top             =   1320
         Width           =   3315
         _Version        =   65536
         _ExtentX        =   5847
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSensorName 
         Height          =   855
         Index           =   2
         Left            =   60
         TabIndex        =   325
         Top             =   2220
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "SENSOR 3"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSensorVolt 
         Height          =   855
         Index           =   2
         Left            =   1920
         TabIndex        =   326
         Top             =   2220
         Width           =   3315
         _Version        =   65536
         _ExtentX        =   5847
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSensorName 
         Height          =   795
         Index           =   3
         Left            =   60
         TabIndex        =   327
         Top             =   3120
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "SENSOR 4"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSensorVolt 
         Height          =   795
         Index           =   3
         Left            =   1920
         TabIndex        =   328
         Top             =   3120
         Width           =   3315
         _Version        =   65536
         _ExtentX        =   5847
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSensorName 
         Height          =   855
         Index           =   0
         Left            =   60
         TabIndex        =   169
         Top             =   420
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "SENSOR 1"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSensorName 
         Height          =   795
         Index           =   4
         Left            =   60
         TabIndex        =   469
         Top             =   3960
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "SENSOR 5"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSensorVolt 
         Height          =   795
         Index           =   4
         Left            =   1920
         TabIndex        =   470
         Top             =   3960
         Width           =   3315
         _Version        =   65536
         _ExtentX        =   5847
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSensorName 
         Height          =   795
         Index           =   5
         Left            =   60
         TabIndex        =   471
         Top             =   4800
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "SENSOR 6"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSensorVolt 
         Height          =   795
         Index           =   5
         Left            =   1920
         TabIndex        =   472
         Top             =   4800
         Width           =   3315
         _Version        =   65536
         _ExtentX        =   5847
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "0.0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlInformation 
      Height          =   2775
      Left            =   14220
      TabIndex        =   26
      Top             =   1140
      Width           =   4755
      _Version        =   65536
      _ExtentX        =   8387
      _ExtentY        =   4895
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
      Alignment       =   0
      Begin VB.ListBox lstMsg 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         ItemData        =   "frmRun.frx":136A0
         Left            =   0
         List            =   "frmRun.frx":136A2
         TabIndex        =   27
         Top             =   360
         Width           =   4755
      End
      Begin Threed.SSPanel pnlInfoTitle 
         Height          =   255
         Left            =   60
         TabIndex        =   28
         Top             =   60
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "INFORMATION"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin HIRESTIMERLib.HiResTimer tmrLoop 
      Left            =   18660
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   64
      Enabled         =   0   'False
      Interval        =   1
   End
   Begin VB.CommandButton btnReturn 
      Caption         =   "RETURN"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   27.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   14220
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   120
      Width           =   4755
   End
   Begin VB.Frame fraHidden 
      Caption         =   "HIDDEN"
      Height          =   2115
      Left            =   18120
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   3915
      Begin VB.PictureBox picCapture 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   1020
         ScaleHeight     =   1005
         ScaleWidth      =   1125
         TabIndex        =   6
         Top             =   600
         Width           =   1155
      End
      Begin VB.Image imgCapture 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1035
         Left            =   2160
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1155
      End
   End
   Begin VB.Timer AutoLinRead 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picLog 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   10755
      Left            =   120
      ScaleHeight     =   10725
      ScaleWidth      =   18825
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   18855
      Begin VB.Label lblLog 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "... !! @@"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   72
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Index           =   2
         Left            =   6615
         TabIndex        =   23
         Top             =   4500
         Width           =   5895
      End
      Begin VB.Label lblLog 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "... !! @@"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   72
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Index           =   4
         Left            =   6615
         TabIndex        =   22
         Top             =   8220
         Width           =   5895
      End
      Begin VB.Label lblLog 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "... !! @@"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   72
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Index           =   3
         Left            =   6615
         TabIndex        =   21
         Top             =   6420
         Width           =   5895
      End
      Begin VB.Label lblLog 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "... !! @@"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   72
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Index           =   1
         Left            =   6630
         TabIndex        =   20
         Top             =   2640
         Width           =   5895
      End
      Begin VB.Label lblLog 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "... !! @@"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   72
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Index           =   0
         Left            =   6615
         TabIndex        =   19
         Top             =   780
         Width           =   5895
      End
   End
   Begin Threed.SSPanel pnlManual 
      Height          =   3795
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9075
      _Version        =   65536
      _ExtentX        =   16007
      _ExtentY        =   6694
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
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2340
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "[Y06] MASTER OK"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
         Height          =   615
         Index           =   2
         Left            =   2520
         TabIndex        =   9
         Top             =   1620
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "[Y02] OK LAMP"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
         Height          =   615
         Index           =   3
         Left            =   4920
         TabIndex        =   10
         Top             =   180
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "[Y03] NG LAMP"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
         Height          =   615
         Index           =   4
         Left            =   4920
         TabIndex        =   11
         Top             =   900
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "[Y04] RUN LAMP"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "[X00] START S/W"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   900
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "[X01] STOP S/W"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1620
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "[X02] AUTO/MANU"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
         Height          =   615
         Index           =   0
         Left            =   2520
         TabIndex        =   15
         Top             =   180
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "[Y00] START LAMP"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
         Height          =   615
         Index           =   1
         Left            =   2520
         TabIndex        =   16
         Top             =   900
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "[Y01] STOP LAMP"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
         Height          =   615
         Index           =   5
         Left            =   4920
         TabIndex        =   17
         Top             =   1620
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "[Y05] BUZZER"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
      Begin BHButton.BHImageButton btnSteppingInfo 
         Height          =   615
         Left            =   120
         TabIndex        =   354
         Top             =   3060
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "Stepping Set"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
         Height          =   615
         Index           =   16
         Left            =   4920
         TabIndex        =   375
         Top             =   2340
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "[Y16] STEP #1 RESET"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
         Height          =   615
         Index           =   17
         Left            =   4920
         TabIndex        =   376
         Top             =   3060
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "[Y17] STEP #2 RESET"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
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
   Begin Threed.SSPanel pnlMS 
      Height          =   4995
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7995
      _Version        =   65536
      _ExtentX        =   14102
      _ExtentY        =   8811
      _StockProps     =   15
      Caption         =   "123456789"
      BackColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Outline         =   -1  'True
      FloodColor      =   16777215
      Alignment       =   6
      Begin VB.CommandButton btnMS 
         BackColor       =   &H00808080&
         Caption         =   "YES"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ ExtraBold"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   3840
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   2
         Top             =   3120
         Width           =   3855
      End
      Begin VB.CommandButton btnMS 
         BackColor       =   &H00808080&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ ExtraBold"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   3840
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   1
         Top             =   3960
         Width           =   3855
      End
      Begin Threed.SSPanel MSOKFail 
         Height          =   735
         Left            =   480
         TabIndex        =   3
         Top             =   3120
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "OK FAIL"
         ForeColor       =   0
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   23.99
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel MSNGFail 
         Height          =   735
         Left            =   480
         TabIndex        =   4
         Top             =   3960
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "NG FAIL"
         ForeColor       =   0
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   23.99
            Charset         =   129
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
   Begin Threed.SSPanel pnlTop 
      Height          =   3795
      Left            =   120
      TabIndex        =   130
      Top             =   120
      Width           =   13995
      _Version        =   65536
      _ExtentX        =   24686
      _ExtentY        =   6694
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
      Begin VB.CommandButton btnStatistical 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "STATISTICAL"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3660
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   390
         Top             =   3060
         Width           =   3555
      End
      Begin VB.ComboBox cboCarType 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         ItemData        =   "frmRun.frx":136A4
         Left            =   2820
         List            =   "frmRun.frx":136A6
         Sorted          =   -1  'True
         TabIndex        =   389
         Top             =   120
         Width           =   4395
      End
      Begin VB.CommandButton btnClearCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "COUNTER RESET"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   388
         Top             =   3060
         Width           =   3555
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   3555
         Index           =   1
         Left            =   7320
         TabIndex        =   380
         Top             =   120
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   6271
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   72.01
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   2
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Begin Threed.SSPanel pnlPlcStatus 
            Height          =   735
            Index           =   0
            Left            =   60
            TabIndex        =   381
            Top             =   60
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   1296
            _StockProps     =   15
            Caption         =   "COMM"
            ForeColor       =   16777215
            BackColor       =   12632064
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlPlcStatus 
            Height          =   495
            Index           =   1
            Left            =   60
            TabIndex        =   382
            Top             =   840
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "READY"
            ForeColor       =   16777215
            BackColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlPlcStatus 
            Height          =   495
            Index           =   2
            Left            =   60
            TabIndex        =   383
            Top             =   1380
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "RUNNING"
            ForeColor       =   16777215
            BackColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlPlcStatus 
            Height          =   495
            Index           =   3
            Left            =   60
            TabIndex        =   384
            Top             =   1920
            Visible         =   0   'False
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "VIBRATION"
            ForeColor       =   16777215
            BackColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnlPlcStatus 
            Height          =   495
            Index           =   5
            Left            =   60
            TabIndex        =   385
            Top             =   3000
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "OK   "
            ForeColor       =   16777215
            BackColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
            Begin VB.Label lblFinalStatus 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TO PLC"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   225
               Left            =   180
               TabIndex        =   386
               Top             =   120
               Width           =   660
            End
         End
         Begin Threed.SSPanel pnlPlcStatus 
            Height          =   495
            Index           =   4
            Left            =   60
            TabIndex        =   387
            Top             =   2460
            Visible         =   0   'False
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "NOTHING"
            ForeColor       =   16777215
            BackColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
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
      Begin VB.Frame fraDebug 
         Caption         =   "DEBUG"
         Height          =   1515
         Left            =   9180
         TabIndex        =   131
         Top             =   2160
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton Command 
            Caption         =   "CAP"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   3
            Left            =   3600
            TabIndex        =   410
            Top             =   900
            Width           =   915
         End
         Begin VB.CommandButton Command 
            Caption         =   "STOP"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   2
            Left            =   2880
            TabIndex        =   379
            Top             =   900
            Width           =   675
         End
         Begin VB.CommandButton Command 
            Caption         =   "MOVE"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   2160
            TabIndex        =   378
            Top             =   900
            Width           =   675
         End
         Begin VB.CommandButton Command 
            Caption         =   "READ"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   0
            Left            =   1440
            TabIndex        =   377
            Top             =   900
            Width           =   675
         End
         Begin VB.TextBox txtDebugModelNo 
            Alignment       =   2  'Center
            Height          =   435
            Left            =   180
            TabIndex        =   134
            ToolTipText     =   "Insert Model No and Debug Start"
            Top             =   900
            Width           =   1215
         End
         Begin VB.CheckBox chkStart 
            Caption         =   "START"
            Height          =   495
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   133
            Top             =   300
            Width           =   2115
         End
         Begin VB.CheckBox chkStop 
            Caption         =   "STOP"
            Height          =   495
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   132
            Top             =   300
            Width           =   2115
         End
      End
      Begin Threed.SSPanel pnlMessage 
         Height          =   3555
         Left            =   9180
         TabIndex        =   135
         Top             =   120
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   6271
         _StockProps     =   15
         Caption         =   "WAIT"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   72.01
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   2
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel lblCounter 
         Height          =   555
         Index           =   0
         Left            =   2400
         TabIndex        =   391
         Top             =   1800
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   18
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel lblCounter 
         Height          =   555
         Index           =   1
         Left            =   2400
         TabIndex        =   392
         Top             =   2400
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   18
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel lblCounter 
         Height          =   555
         Index           =   2
         Left            =   5940
         TabIndex        =   393
         Top             =   2400
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   18
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel lblCounter 
         Height          =   555
         Index           =   3
         Left            =   5940
         TabIndex        =   394
         Top             =   1800
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   18
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSerial 
         Height          =   435
         Left            =   120
         TabIndex        =   395
         Top             =   1320
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "000000000000000"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlSerialName 
         Height          =   315
         Left            =   180
         TabIndex        =   396
         Top             =   900
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "SERIAL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
            Size            =   14.26
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Alignment       =   1
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   915
         Index           =   0
         Left            =   2820
         TabIndex        =   401
         Top             =   840
         Width           =   4395
         _Version        =   65536
         _ExtentX        =   7752
         _ExtentY        =   1614
         _StockProps     =   15
         BackColor       =   16777215
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
         Begin Threed.SSPanel pnlCarType 
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   402
            Top             =   120
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0"
            ForeColor       =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ ExtraBold"
               Size            =   12.01
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel pnlCarName 
            Height          =   315
            Index           =   0
            Left            =   300
            TabIndex        =   403
            Top             =   120
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CarType :"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ ExtraBold"
               Size            =   11.26
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Alignment       =   1
         End
         Begin Threed.SSPanel pnlCarType 
            Height          =   315
            Index           =   1
            Left            =   3660
            TabIndex        =   404
            Top             =   120
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ ExtraBold"
               Size            =   12.01
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel pnlCarName 
            Height          =   315
            Index           =   1
            Left            =   2460
            TabIndex        =   405
            Top             =   120
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CarRank :"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ ExtraBold"
               Size            =   11.26
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Alignment       =   1
         End
         Begin Threed.SSPanel pnlCarType 
            Height          =   315
            Index           =   2
            Left            =   1500
            TabIndex        =   406
            Top             =   480
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ ExtraBold"
               Size            =   12.01
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel pnlCarName 
            Height          =   315
            Index           =   2
            Left            =   300
            TabIndex        =   407
            Top             =   480
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Group :"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ ExtraBold"
               Size            =   11.26
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Alignment       =   1
         End
         Begin Threed.SSPanel pnlCarType 
            Height          =   315
            Index           =   3
            Left            =   3660
            TabIndex        =   408
            Top             =   480
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ ExtraBold"
               Size            =   12.01
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel pnlCarName 
            Height          =   315
            Index           =   3
            Left            =   2460
            TabIndex        =   409
            Top             =   480
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Pallet :"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "³ª´®°íµñ ExtraBold"
               Size            =   11.26
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Alignment       =   1
         End
      End
      Begin VB.Image Image 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   690
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2670
      End
      Begin VB.Label lblCounterName 
         Caption         =   "RATIO"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ ExtraBold"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3840
         TabIndex        =   400
         Top             =   1860
         Width           =   2055
      End
      Begin VB.Label lblCounterName 
         Caption         =   "NG"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ ExtraBold"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3840
         TabIndex        =   399
         Top             =   2460
         Width           =   2055
      End
      Begin VB.Label lblCounterName 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ ExtraBold"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   300
         TabIndex        =   398
         Top             =   2460
         Width           =   2055
      End
      Begin VB.Label lblCounterName 
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ ExtraBold"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   300
         TabIndex        =   397
         Top             =   1860
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bLoading As Boolean

Private Sub AutoLinRead_Timer()
    Const LOGUSE As Boolean = True
    Const DETECTCOUNT As Byte = 3
    
    Dim ReceiveBuf(250) As NCTYPE_CAN_STRUCT
    Dim ActualDataSize As Long
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim lRes As Long
    Dim lpRes As String
    
    If bAutoLinRead = False Or LINUSE = False Then Exit Sub
    
    DoEvents
    
    Status = ncReadMultiple(LinTxRx, 250 * LenB(ReceiveBuf(0)), ReceiveBuf, ActualDataSize)
    If (CheckStatus(Status, "ncRead ") = True) Then GoTo ERROR
    
    'if frames were received, display them
    If (ActualDataSize >= 1) Then
        For i = 0 To (ActualDataSize / 24) - 1
            lpRes = ""
            
            Select Case ReceiveBuf(i).FrameType
                Case NC_FRMTYPE_LIN_FULL:
                    If LOGUSE And ReceiveBuf(i).ArbitrationId <> &H22 Then
                        lpRes = lpRes + CStr(Format(Hex(ReceiveBuf(i).ArbitrationId), "00")) + " "
                        lpRes = lpRes + CStr(Format(ReceiveBuf(i).DataLength, "00")) + " "
                        
                        For j = 0 To (ReceiveBuf(i).DataLength) - 1
                            lpRes = lpRes + CStr(Format(Hex(ReceiveBuf(i).Data(j)), "00")) & " "
                        Next j
                        
                        Call OnLog("[" & Format(nLinReadSeq(nLinActNo), "00") & "]" & " " & lpRes)
                    End If
                    
                    Select Case ReceiveBuf(i).Data(6)
                        Case BYTE1_ACT01: nLinActNo = 0
                        Case BYTE1_ACT02: nLinActNo = 1
                        Case BYTE1_ACT03: nLinActNo = 2
                        Case BYTE1_ACT04: nLinActNo = 3
                        Case Else: nLinActNo = 9
                    End Select
                    
                    If nLinActNo = 9 Then
                        Select Case ReceiveBuf(i).ArbitrationId
                            Case BYTE1_PTC: nLinActNo = 4
                            Case BYTE1_BLOWER: nLinActNo = 5
                        End Select
                    End If
                    
                    If nLinActNo = 9 Then Exit Sub
                    
                    Select Case nLinReadSeq(nLinActNo)
                        Case 1, 2:
                            If RunVar.bLinDataResult(nLinReadSeq(nLinActNo), nLinActNo) = False Then
                                frmRun.pnlLinActMove(nLinActNo).Caption = Format((CLng(ReceiveBuf(i).Data(4)) * 256) + CLng(ReceiveBuf(i).Data(3)), "0")
                                
                                If RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = False Then
                                    If nLinActNo = 0 And SetupVar.nLinAct01RefPos = 0 Then
                                        RunVar.bLinRefPos(nLinReadSeq(nLinActNo)) = bLinAct01Detect
                                        
                                        If nLinReadSeq(nLinActNo) = 2 Then
                                            If (RunVar.lLinDataMove(1, nLinActNo) + 1000) < ((CLng(ReceiveBuf(i).Data(4)) * 256) + CLng(ReceiveBuf(i).Data(3))) Then
                                                If RunVar.bLinRefPos(nLinReadSeq(nLinActNo)) Then
                                                    RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = True
                                                End If
                                            End If
                                        ElseIf nLinReadSeq(nLinActNo) = 1 Then
                                            If RunVar.bLinRefPos(nLinReadSeq(nLinActNo)) Then
                                                RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = True
                                            End If
                                        End If
                                    Else
                                        RunVar.bLinRefPos(nLinReadSeq(nLinActNo)) = (ReceiveBuf(i).Data(1) And &H10) = &H10
                                        
                                        If RunVar.bLinRefPos(nLinReadSeq(nLinActNo)) Then
                                            RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = True
                                        End If
                                    End If
                                End If
                                
                                If RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) Then
                                    If SetupVar.nLinActFirstMove(nLinActNo) = 0 Then
                                        If nLinReadSeq(nLinActNo) = 1 Then
                                            frmRun.pnlLinActClose(nLinActNo).BackColor = vbGreen
                                        Else
                                            frmRun.pnlLinActOpen(nLinActNo).BackColor = vbGreen
                                        End If
                                    Else
                                        If nLinReadSeq(nLinActNo) = 1 Then
                                            frmRun.pnlLinActOpen(nLinActNo).BackColor = vbGreen
                                        Else
                                            frmRun.pnlLinActClose(nLinActNo).BackColor = vbGreen
                                        End If
                                    End If
                                    
                                    RunVar.lLinDataMove(nLinReadSeq(nLinActNo), nLinActNo) = (CLng(ReceiveBuf(i).Data(4)) * 256) + CLng(ReceiveBuf(i).Data(3))
                                    
                                    If SetupVar.bStallUse = False Then
                                        If nLinReadSeq(nLinActNo) = 1 Then
                                            RunVar.lLinDataMove(nLinReadSeq(nLinActNo) + 1, nLinActNo) = RunVar.lLinDataMove(nLinReadSeq(nLinActNo), nLinActNo) + SetupVar.lLinActAngle(nLinActNo)
                                        End If
                                    End If
                                    
                                    Call LinInit(ReceiveBuf(i).Data(6))
                                    
                                    RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = False
                                    RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 1) = False
                                    RunVar.bLinDataResult(nLinReadSeq(nLinActNo), nLinActNo) = True
                                End If
                            End If
                        
                        Case 3:
                            If RunVar.bLinDataResult(nLinReadSeq(nLinActNo), nLinActNo) = False Then
                                frmRun.pnlLinActMove(nLinActNo).Caption = Format((CLng(ReceiveBuf(i).Data(4)) * 256) + CLng(ReceiveBuf(i).Data(3)), "0")
                                
                                If RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = False Then
                                    If RunVar.lLinDataCP(nLinActNo, nLinTimeNo) = RunVar.lLinDataMove(nLinReadSeq(nLinActNo), nLinActNo) Then
                                        nLinCount(nLinReadSeq(nLinActNo), nLinActNo) = nLinCount(nLinReadSeq(nLinActNo), nLinActNo) + 1
                                    Else
                                        nLinCount(nLinReadSeq(nLinActNo), nLinActNo) = 0
                                    End If
                                    
                                    If nLinCount(nLinReadSeq(nLinActNo), nLinActNo) > DETECTCOUNT Then
                                        Call SetTime(TM_LIN)
                                        
                                        RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = True
                                    End If
                                End If
                                
                                If RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) Then
                                    Call OnLog("[DEBUG] CP TIME : " & RunVar.dLinCPTime(nLinActNo, nLinTimeNo) & "   ACT NO : " & nLinActNo & "   TIME NO : " & nLinTimeNo)
                                    
                                    If ElapseTime(TM_LIN) > RunVar.dLinCPTime(nLinActNo, nLinTimeNo) Then
                                        RunVar.lLinDataMove(nLinReadSeq(nLinActNo), nLinActNo) = (CLng(ReceiveBuf(i).Data(4)) * 256) + CLng(ReceiveBuf(i).Data(3))
                                        
                                        nLinCount(nLinReadSeq(nLinActNo), nLinActNo) = 0
                                        
                                        RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = False
                                        RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 1) = False
                                        RunVar.bLinCheckPoint(nLinActNo, nLinTimeNo) = False
                                        RunVar.bLinDataResult(nLinReadSeq(nLinActNo), nLinActNo) = True
                                        
                                        For k = 0 To 4
                                            If RunVar.bLinCheckPoint(nLinActNo, k) Then
                                                RunVar.bLinDataResult(nLinReadSeq(nLinActNo), nLinActNo) = False
                                                
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        
                        Case 4:
                            If RunVar.bLinDataResult(nLinReadSeq(nLinActNo), nLinActNo) = False Then
                                frmRun.pnlLinActFinal(nLinActNo).Caption = Format((CLng(ReceiveBuf(i).Data(4)) * 256) + CLng(ReceiveBuf(i).Data(3)), "0")
                                
                                If RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = False Then
                                    If RunVar.lLinDataFinal(nLinActNo) = Val(frmRun.pnlLinActFinal(nLinActNo).Caption) Then
                                        nLinCount(nLinReadSeq(nLinActNo), nLinActNo) = nLinCount(nLinReadSeq(nLinActNo), nLinActNo) + 1
                                    Else
                                        nLinCount(nLinReadSeq(nLinActNo), nLinActNo) = 0
                                    End If
                                    
                                    If nLinCount(nLinReadSeq(nLinActNo), nLinActNo) > DETECTCOUNT Then
                                        RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = True
                                    End If
                                End If
                                
                                If RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) Then
                                    RunVar.lLinDataMove(nLinReadSeq(nLinActNo), nLinActNo) = (CLng(ReceiveBuf(i).Data(4)) * 256) + CLng(ReceiveBuf(i).Data(3))
                                    
                                    nLinCount(nLinReadSeq(nLinActNo), nLinActNo) = 0
                                    
                                    Call LinInit(ReceiveBuf(i).Data(6))
                                    
                                    RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = False
                                    RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 1) = False
                                    RunVar.bLinDataResult(nLinReadSeq(nLinActNo), nLinActNo) = True
                                End If
                            End If
                        
                        Case 5:
                            If RunVar.bLinDataResult(nLinReadSeq(nLinActNo), nLinActNo) = False Then
                                If RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = False Then
                                    lRes = &H18
                                    
                                    If CLng(ReceiveBuf(i).Data(6)) = lRes Then
                                        RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = True
                                    End If
                                End If
                                
                                If RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) Then
                                    RunVar.lLinDataMove(nLinReadSeq(nLinActNo), nLinActNo) = CLng(ReceiveBuf(i).Data(6))
                                    
                                    nLinCount(nLinReadSeq(nLinActNo), nLinActNo) = 0
                                    
                                    Select Case nLinReadSeq(nLinActNo)
                                        Case 5: Call LinInit(BYTE1_PTC)
                                        Case 6: Call LinInit(BYTE1_BLOWER)
                                    End Select
                                    
                                    RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 0) = False
                                    RunVar.bLinReadRes(nLinReadSeq(nLinActNo), 1) = False
                                    RunVar.bLinDataResult(nLinReadSeq(nLinActNo), nLinActNo) = True
                                End If
                            End If
                        
                        Case 10, 11, 12, 13:
                            If CLng(ReceiveBuf(i).Data(6)) = Choose(nLinReadSeq(nLinActNo) - 9, BYTE1_ACT01, BYTE1_ACT02, BYTE1_ACT03, BYTE1_ACT04) Then
                                frmRun.pnlLinAddr(nLinReadSeq(nLinActNo) - 10).Caption = "OK"
                                frmRun.pnlLinAddr(nLinReadSeq(nLinActNo) - 10).BackColor = vbGreen
                            Else
                                frmRun.pnlLinAddr(nLinReadSeq(nLinActNo) - 10).Caption = "NG"
                                frmRun.pnlLinAddr(nLinReadSeq(nLinActNo) - 10).BackColor = vbRed
                            End If
                        
                        Case 18, 19:
                            frmRun.pnlLinActMove(nLinActNo).Caption = Format((CLng(ReceiveBuf(i).Data(4)) * 256) + CLng(ReceiveBuf(i).Data(3)), "0")
                    
                    End Select
                    
                    DoEvents
                
                Case NC_FRMTYPE_LIN_WAKEUP_RECEIVED:
                    If LOGUSE Then
                        lpRes = lpRes + CStr(Format(Hex(ReceiveBuf(i).ArbitrationId), "00")) + " "
                        lpRes = lpRes + "LIN Wakeup" + " "
                        lpRes = lpRes + "0" + " "
                        lpRes = lpRes + "No Data"
                        
                        Call OnLog("[" & nLinReadSeq(nLinActNo) & " " & lpRes)
                    End If
                
                Case NC_FRMTYPE_LIN_BUS_INACTIVE:
                    If LOGUSE Then
                        lpRes = lpRes + CStr(Format(Hex(ReceiveBuf(i).ArbitrationId), "00")) + " "
                        lpRes = lpRes + "LIN Bus Inactive" + " "
                        lpRes = lpRes + "0" + " "
                        lpRes = lpRes + "No Data"
                        
                        Call OnLog("[" & nLinReadSeq(nLinActNo) & "]" & " " & lpRes)
                    End If
                
                Case NC_FRMTYPE_LIN_BUS_ERR:
                    If LOGUSE Then
                        lpRes = lpRes + CStr(Format(Hex(ReceiveBuf(i).ArbitrationId), "00")) + " "
                        lpRes = lpRes + CStr(Format(ReceiveBuf(i).DataLength, "00")) + " "
                        
                        For j = 0 To (ReceiveBuf(i).DataLength) - 1
                            lpRes = lpRes + CStr(Format(Hex(ReceiveBuf(i).Data(j)), "00")) & " "
                        Next j
                        
                        lpRes = lpRes & "ERROR"
                        
                        Call OnLog("[" & nLinReadSeq(nLinActNo) & "]" & " " & lpRes)
                    End If
                    
                    bLinErrorResult(nLinActNo) = True
                    
                    GoTo ERROR
                    
                    DoEvents
            
            End Select
            
            DoEvents
        Next i
    End If
    
    lpRes = ""
    
    Exit Sub

ERROR:
    
    Select Case nLinActNo
        Case 0: RunVar.bTestEnd(TP_LINACT01) = True
        Case 1: RunVar.bTestEnd(TP_LINACT02) = True
        Case 2: RunVar.bTestEnd(TP_LINACT03) = True
        Case 3: RunVar.bTestEnd(TP_LINACT04) = True
        Case 4: RunVar.bTestEnd(TP_LINPTC) = True
    End Select
End Sub

Private Sub Command_Click(Index As Integer)
    Select Case Index
        Case 0:
            nSteppingMode = STEP_MODE_READ
        Case 1:
            nSteppingMode = STEP_MODE_MOVE
            bSteppingWrite(0) = True
            
            lpWriteStepRotation(0) = STEP_ROTA_P2
            lpWriteStepData(0) = "EE"
            lpWriteStepData(1) = "EE"
        Case 3:
            Call SaveCapture
    End Select
End Sub

Private Sub tmrSteppingSend_Timer()
    Dim i As Integer
    Dim nData(9) As Byte
    Dim lpData(1) As String
    Dim lpStr As String
    
    If LOCALTEST Or STEPUSE = False Or nSteppingMode = 0 Then
        Exit Sub
    End If
    
    DoEvents
    
    Select Case nSteppingMode
        Case STEP_MODE_READ:
            lpData(0) = STEP_START & STEP_READ & "03" & STEP_NULL & STEP_NULL & STEP_NULL & STEP_NULL & STEP_NULL & STEP_NULL & STEP_END
            lpData(1) = STEP_START & STEP_READ & "03" & STEP_NULL & STEP_NULL & STEP_NULL & STEP_NULL & STEP_NULL & STEP_NULL & STEP_END
        
        Case STEP_MODE_MOVE:

            If bSteppingWrite(0) Or bSteppingWrite(1) Then
                If bSteppingWrite(0) Then nSteppingMoveBit(0) = 1
                If bSteppingWrite(1) Then nSteppingMoveBit(0) = 2
                If bSteppingWrite(0) And bSteppingWrite(1) Then nSteppingMoveBit(0) = 3
                
                lpData(0) = STEP_START & STEP_MOVE & Dec2Hex(Hex(nSteppingMoveBit(0))) & lpWriteStepRotation(0) & lpWriteStepData(1) & lpWriteStepData(0) & lpWriteStepRotation(1) & lpWriteStepData(3) & lpWriteStepData(2) & STEP_END
                
                If bSteppingWrite(0) Then nSteppingMoveBit(0) = nSteppingMoveBit(0) - 1
                If bSteppingWrite(1) Then nSteppingMoveBit(0) = nSteppingMoveBit(0) - 2
                If bSteppingWrite(0) And bSteppingWrite(1) Then nSteppingMoveBit(0) = nSteppingMoveBit(0) - 3
                
                If bSteppingWrite(0) Then
                    bSteppingWrite(0) = False
                    lpWriteStepRotation(0) = STEP_NULL
                    lpWriteStepData(0) = STEP_NULL
                    lpWriteStepData(1) = STEP_NULL
                End If
                
                If bSteppingWrite(1) Then
                    bSteppingWrite(1) = False
                    lpWriteStepRotation(1) = STEP_NULL
                    lpWriteStepData(2) = STEP_NULL
                    lpWriteStepData(3) = STEP_NULL
                End If
            End If
            
            If bSteppingWrite(2) Or bSteppingWrite(3) Then
                If bSteppingWrite(2) Then nSteppingMoveBit(1) = 1
                If bSteppingWrite(3) Then nSteppingMoveBit(1) = 2
                If bSteppingWrite(2) And bSteppingWrite(3) Then nSteppingMoveBit(1) = 3
                
                lpData(1) = STEP_START & STEP_MOVE & Dec2Hex(Hex(nSteppingMoveBit(1))) & lpWriteStepRotation(2) & lpWriteStepData(5) & lpWriteStepData(4) & lpWriteStepRotation(3) & lpWriteStepData(7) & lpWriteStepData(6) & STEP_END
                
                If bSteppingWrite(2) Then nSteppingMoveBit(1) = nSteppingMoveBit(1) - 1
                If bSteppingWrite(3) Then nSteppingMoveBit(1) = nSteppingMoveBit(1) - 2
                If bSteppingWrite(2) And bSteppingWrite(3) Then nSteppingMoveBit(1) = nSteppingMoveBit(1) - 3
                
                If bSteppingWrite(2) Then
                    bSteppingWrite(2) = False
                    lpWriteStepRotation(2) = STEP_NULL
                    lpWriteStepData(4) = STEP_NULL
                    lpWriteStepData(5) = STEP_NULL
                End If
                
                If bSteppingWrite(3) Then
                    bSteppingWrite(3) = False
                    lpWriteStepRotation(3) = STEP_NULL
                    lpWriteStepData(6) = STEP_NULL
                    lpWriteStepData(7) = STEP_NULL
                End If
            End If
    
    End Select
    
    If Len(lpData(0)) = 20 Then
        If STEPLOGUSE Then
            Call OnLog("[STEP] SEND 1 : " & lpData(0))
        End If
        
        For i = 0 To 9
            nData(i) = Val("&H" & Mid(lpData(0), (i * 2) + 1, 2))
        Next
        
        frmMain.comStepping1.Output = nData
    Else
        If STEPLOGUSE Then
            Call OnLog("[STEP] ERR  1 : " & lpData(0))
        End If
    End If
        
    If Len(lpData(1)) = 20 Then
        If STEPLOGUSE Then
            Call OnLog("[STEP] SEND 2 : " & lpData(1))
        End If
        
        For i = 0 To 9
            nData(i) = Val("&H" & Mid(lpData(1), (i * 2) + 1, 2))
        Next
        
        frmMain.comStepping2.Output = nData
    Else
        If STEPLOGUSE Then
            Call OnLog("[STEP] ERR  2 : " & lpData(1))
        End If
    End If
    
    If RunVar.bAutoManual Then
        nSteppingMode = 0
    End If
End Sub

Private Sub btnLinActManual_Click(Index As Integer)
    Select Case Index
        Case 0:
            Call LinInit(BYTE1_ACT01)
            Call Delay(50)
            Call LinActWrite(8, &H22, BYTE1_ACT01, &H4, &H51, &H0, &H0, &H0, &H0, &H0)
            nLinReadSeq(0) = 18
        
        Case 1:
            Call LinInit(BYTE1_ACT01)
            Call Delay(50)
            Call LinActWrite(8, &H22, BYTE1_ACT01, &H4, &H51, &HFE, &HFF, &H0, &H0, &H0)
            nLinReadSeq(0) = 19
            
        Case 2:
            Call LinInit(BYTE1_ACT02)
            Call Delay(50)
            Call LinActWrite(8, &H22, BYTE1_ACT02, &H4, &H51, &H0, &H0, &H0, &H0, &H0)
            nLinReadSeq(1) = 18
        
        Case 3:
            Call LinInit(BYTE1_ACT02)
            Call Delay(50)
            Call LinActWrite(8, &H22, BYTE1_ACT02, &H4, &H51, &HFE, &HFF, &H0, &H0, &H0)
            nLinReadSeq(1) = 19
        
        Case 4:
            Call LinInit(BYTE1_ACT03)
            Call Delay(50)
            Call LinActWrite(8, &H22, BYTE1_ACT03, &H4, &H51, &H0, &H0, &H0, &H0, &H0)
            nLinReadSeq(2) = 18
        
        Case 5:
            Call LinInit(BYTE1_ACT03)
            Call Delay(50)
            Call LinActWrite(8, &H22, BYTE1_ACT03, &H4, &H51, &HFE, &HFF, &H0, &H0, &H0)
            nLinReadSeq(2) = 19
        
        Case 6:
            Call LinInit(BYTE1_ACT04)
            Call Delay(50)
            Call LinActWrite(8, &H22, BYTE1_ACT04, &H4, &H51, &H0, &H0, &H0, &H0, &H0)
            nLinReadSeq(3) = 18
        
        Case 7:
            Call LinInit(BYTE1_ACT04)
            Call Delay(50)
            Call LinActWrite(8, &H22, BYTE1_ACT04, &H4, &H51, &HFE, &HFF, &H0, &H0, &H0)
            nLinReadSeq(3) = 19
    
    End Select
    
    bAutoLinRead = IIf(bAutoLinRead, False, True)
End Sub

Private Sub btnLinInit_Click()
    Call LinInit
    
    nLinReadSeq(0) = 0
    nLinReadSeq(1) = 0
    nLinReadSeq(2) = 0
    nLinReadSeq(3) = 0
End Sub

Private Sub btnMS_Click(Index As Integer)
    MSFlag = Index + 1
End Sub

Private Sub btnSteppingInfo_Click()
    Call StepStartInfo
End Sub

Private Sub optIonManual_Click(Index As Integer)
    Dim nRes As Integer
    
    Select Case Index
        Case 0:
            nRes = 2
            
            Call DO_Control(O_ION_DIAG, False)
        Case 1:
            nRes = 3
            
            Call DO_Control(O_ION_DIAG, True)
    End Select
    
    ManuVar.nIonPos = nRes
End Sub

Private Sub pnlMessage_Click()
    nSteppingMode = STEP_MODE_READ
End Sub

Private Sub pnlNGCountValue_Click()
    Dim lpRes As String
    
    lpRes = InputBox("ºÒ·® ¼ö·®À» ÀÔ·ÂÇÏ¼¼¿ä.", frmRun.pnlNGLabel(1).Caption, frmRun.pnlNGCountValue.Caption)
    
    If lpRes = "" Then lpRes = frmRun.pnlNGCountValue.Caption
    
    SysVar.nContinueNGQty = Val(lpRes)
    
    Call SaveSystemFile
    
    frmRun.pnlNGCountValue.Caption = lpRes
End Sub

Private Sub btnNGCountReset_Click()
    frmRun.pnlNGCount.Caption = 0
End Sub

Private Sub cboCarType_Click()
    If RunVar.bLoading = True Then
        If SysVar.nMSUse = 2 And (lpNowModel <> Trim$(UCase(cboCarType.List(cboCarType.ListIndex)))) Then
            bMSTestStart = True
        End If
        
        lpNowModel = Trim$(UCase(cboCarType.List(cboCarType.ListIndex)))
        
        Call LoadSetupFile(lpNowModel)
        Call ScreenClear(True)
        Call MsgLog(MSG_READY)
    End If
End Sub

Private Sub btnClearCounter_Click()
    If MsgBox("Do you want to clear?", vbYesNo + vbQuestion, "Check") = vbYes Then
        SysVar.lTotalCounter = 0
        SysVar.lOkCounter = 0
        SysVar.lNgCounter = 0
        
        Call DispCounter
    End If
End Sub

Private Sub btnDO_Click(Index As Integer)
    Dim bMove As Boolean
    
    bMove = True
    
    If TABLETYPE Then
        bMove = MarkingInterlock(Index)
    End If
    
    ' ÁøÂ¥ ¿òÁ÷ÀÌ´Â ºÎºÐ
    If bMove Then
        If DOS(Index) Then
            Call DO_Control(Index, False)
        Else
            Call DO_Control(Index, True)
        End If
    End If
End Sub

Private Sub btnReturn_Click()
    bRunningGraphWin = False
    
'    If frmRun.pnlMessage.Caption = "OK" Or frmRun.pnlMessage.Caption = "NG" Then
'        bScreenLoad = True
'    End If
    
    SysVar.lpOldModelSave = Trim$(frmRun.pnlModelFullName.Caption)
    
    Call SaveSystemFile
    Unload Me
End Sub

Private Sub btnDiagnostic_Click()
    bRunningGraphWin = False
    RunVar.bMSStart = True
End Sub

Private Sub btnStatistical_Click()
    bRunningGraphWin = True
    frmGraph.Show vbModal
End Sub

Private Sub btnBarcodePrint_Click()
    Call BarCodePrint
End Sub

Private Sub chkLinActPower_Click()
    If frmRun.chkLinActPower.Value = 1 Then
        Call LinInit
        Call DO_Control(O_LIN_POWER, True)
        Call DO_Control(O_LIN_ACT_POWER, True)
        Call Delay(1000)
        
        If LinConnect(LIN_TEST_CHKSUM) <> 1 Then
            Call OnLog("LIN POWER CHECK...")
            
            frmRun.chkLinActPower.Value = 0
            
            Exit Sub
        Else
            Call OnLog("LIN OPEN...")
        End If
    Else
        bAutoLinRead = False
        
        Call LinInit
        Call Delay(1000)
        Call DO_Control(O_LIN_POWER, False)
        Call DO_Control(O_LIN_ACT_POWER, False)
        Call LinPortClose(1)
    End If
End Sub

Private Sub chkStart_Click()
    If txtDebugModelNo.Text <> "" Then
        lpNowModel = txtDebugModelNo.Text
    End If
    
    If chkStart.Value = 1 Then
        bPlcStartSig = True
        bGlobalStartSw = True
    Else
        bPlcStartSig = False
        bGlobalStartSw = False
    End If
End Sub

Private Sub chkStop_Click()
    If chkStop.Value = 1 Then
        bGlobalStopSw = True
    Else
        bGlobalStopSw = False
    End If
End Sub

Private Sub chkBlowerPower_Click()
    Dim i As Byte
    
    If chkBlowerPower.Value = 1 Then
        Call DO_Control(O_BLOWER_POWER, True)
    Else
        Call DO_Control(O_BLOWER_POWER, False)
    End If
    
    Select Case SetupVar.nBlowerType
        Case 1:
            If chkBlowerPower.Value = 1 Then
                Call DO_Control(O_BLOWER_PWM, True)
            Else
                Call DO_Control(O_BLOWER_PWM, False)
            End If
        
        Case 2:
            If chkBlowerPower.Value = 1 Then
                If LinConnect(LIN_TEST_CHKSUM) <> 1 Then
                    Call OnLog("LIN POWER CHECK...")
                    
                    frmRun.chkBlowerPower.Value = 0
                    
                    Exit Sub
                Else
                    Call OnLog("LIN OPEN...")
                End If
            Else
                Call LinPortClose(1)
                
                For i = 0 To 4
                    frmRun.optBlowerManual(i).Value = 0
                    frmRun.optBlowerManual(i).BackColor = &HFF80FF
                Next
            End If
        
    End Select
End Sub

Private Sub optBlowerManual_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 8
        If i <> Index Then
            frmRun.optBlowerManual(i).Value = 0
        End If
    Next
    
    Select Case SetupVar.nBlowerType
        Case 0, 1:
            If DOS(O_BLOWER_01) Then Call DO_Control(O_BLOWER_01, False)
            If DOS(O_BLOWER_02) Then Call DO_Control(O_BLOWER_02, False)
            If DOS(O_BLOWER_03) Then Call DO_Control(O_BLOWER_03, False)
            If DOS(O_BLOWER_04) Then Call DO_Control(O_BLOWER_04, False)
            If DOS(O_BLOWER_05) Then Call DO_Control(O_BLOWER_05, False)
            If DOS(O_BLOWER_06) Then Call DO_Control(O_BLOWER_06, False)
            If DOS(O_BLOWER_07) Then Call DO_Control(O_BLOWER_07, False)
            If DOS(O_BLOWER_08) Then Call DO_Control(O_BLOWER_08, False)
            
            Select Case Index
                Case 1: Call DO_Control(O_BLOWER_01, True)
                Case 2: Call DO_Control(O_BLOWER_02, True)
                Case 3: Call DO_Control(O_BLOWER_03, True)
                Case 4: Call DO_Control(O_BLOWER_04, True)
                Case 5: Call DO_Control(O_BLOWER_05, True)
                Case 6: Call DO_Control(O_BLOWER_06, True)
                Case 7: Call DO_Control(O_BLOWER_07, True)
                Case 8: Call DO_Control(O_BLOWER_08, True)
            End Select
        
        Case 2:
            Call LinBlrWrite(SetupVar.nLinSpeed(Index))
            
            For i = 0 To 4
                frmRun.optBlowerManual(i).Value = 0
                frmRun.optBlowerManual(i).BackColor = &HFF80FF
            Next
            
            frmRun.optBlowerManual(Index).BackColor = vbGreen
    End Select
    
    ManuVar.nBlowerPos = Index
End Sub

Private Sub chkActPower_Click(Index As Integer)
    Dim OPower As Integer
    
    OPower = ActNo(Index).O_POWER
    
    nSteppingMode = 0
    
    Call DO_Control(OPower, IIf(chkActPower(Index).Value = 1, True, False))
End Sub

Private Sub optAct01Manual_Click(Index As Integer)
    Dim i As Integer
    Dim ActDa As Integer
    Dim dActVolt As Double
    Dim lpRotation(1) As String
    Dim lpPos(3) As String
    Dim lpActNo As String
    Dim nPortNo As Integer
    Dim nActNo As Integer
    
    If SetupVar.nAct01TestType = 1 Then
        ActDa = ActNo(0).DA_NO
        dActVolt = SetupVar.dAct01SetVolt(Index)
        ManuVar.nAct01Pos = Index
        
        Call OutDa(ActDa, dActVolt, SysVar.bPercent(ActNo(0).AD_VOLT))
    Else
        ManuVar.nAct01Pos = Index
        nActNo = 0
        nSteppingMode = 0
        
        Select Case Index
            Case 0:
                bSteppingWrite(0) = True
                
                lpWriteStepRotation(0) = STEP_ROTA_P1
                lpWriteStepData(0) = STEP_POS1
                lpWriteStepData(1) = STEP_POS2
                
                lpWriteStepRotation(1) = STEP_NULL
                lpWriteStepData(2) = STEP_NULL
                lpWriteStepData(3) = STEP_NULL
            Case 1, 2, 3, 4, 5:
                bSteppingWrite(0) = True
                
                lpWriteStepRotation(0) = STEP_ROTA_P2
                lpWriteStepData(0) = Dec2Hex(Val2Byte(CM_LO, SetupVar.dAct01SetVolt(Index) - SetupVar.dAct01SetVolt(Index - 1)))
                lpWriteStepData(1) = Dec2Hex(Val2Byte(CM_HI, SetupVar.dAct01SetVolt(Index) - SetupVar.dAct01SetVolt(Index - 1)))
                
                lpWriteStepRotation(1) = STEP_NULL
                lpWriteStepData(2) = STEP_NULL
                lpWriteStepData(3) = STEP_NULL
            Case 6:
                bSteppingWrite(0) = True
                
                lpWriteStepRotation(0) = STEP_ROTA_P2
                lpWriteStepData(0) = STEP_POS1
                lpWriteStepData(1) = STEP_POS2
                
                lpWriteStepRotation(1) = STEP_NULL
                lpWriteStepData(2) = STEP_NULL
                lpWriteStepData(3) = STEP_NULL
        End Select
        
        nSteppingMode = STEP_MODE_MOVE
        
        frmRun.optAct01Manual(Index).Value = False
    End If
End Sub

Private Sub optAct02Manual_Click(Index As Integer)
    Dim i As Integer
    Dim ActDa As Integer
    Dim dActVolt As Double
    Dim lpRotation(1) As String
    Dim lpPos(3) As String
    Dim lpActNo As String
    Dim nPortNo As Integer
    Dim nActNo As Integer
    
    If SetupVar.nAct02TestType = 1 Then
        ActDa = ActNo(1).DA_NO
        dActVolt = SetupVar.dAct02SetVolt(Index)
        ManuVar.nAct02Pos = Index
        
        Call OutDa(ActDa, dActVolt, SysVar.bPercent(ActNo(1).AD_VOLT))
    Else
        ManuVar.nAct02Pos = Index
        nActNo = 1
        nSteppingMode = 0
        
        Select Case Index
            Case 0:
                bSteppingWrite(1) = True
                
                lpWriteStepRotation(0) = STEP_NULL
                lpWriteStepData(0) = STEP_NULL
                lpWriteStepData(1) = STEP_NULL
                
                lpWriteStepRotation(1) = STEP_ROTA_P1
                lpWriteStepData(2) = STEP_POS1
                lpWriteStepData(3) = STEP_POS2
            Case 1, 2:
                bSteppingWrite(1) = True
                
                lpWriteStepRotation(0) = STEP_NULL
                lpWriteStepData(0) = STEP_NULL
                lpWriteStepData(1) = STEP_NULL
                
                lpWriteStepRotation(1) = STEP_ROTA_P2
                lpWriteStepData(2) = Dec2Hex(Val2Byte(CM_LO, SetupVar.dAct02SetVolt(Index) - SetupVar.dAct02SetVolt(Index - 1)))
                lpWriteStepData(3) = Dec2Hex(Val2Byte(CM_HI, SetupVar.dAct02SetVolt(Index) - SetupVar.dAct02SetVolt(Index - 1)))
            Case 3:
                bSteppingWrite(1) = True
                
                lpWriteStepRotation(0) = STEP_NULL
                lpWriteStepData(0) = STEP_NULL
                lpWriteStepData(1) = STEP_NULL
                
                lpWriteStepRotation(1) = STEP_ROTA_P2
                lpWriteStepData(2) = STEP_POS1
                lpWriteStepData(3) = STEP_POS2
        End Select
        
        nSteppingMode = STEP_MODE_MOVE
        
        frmRun.optAct02Manual(Index).Value = False
    End If
End Sub

Private Sub optAct03Manual_Click(Index As Integer)
    Dim i As Integer
    Dim ActDa As Integer
    Dim dActVolt As Double
    Dim lpRotation(1) As String
    Dim lpPos(3) As String
    Dim lpActNo As String
    Dim nPortNo As Integer
    Dim nActNo As Integer
    
    If SetupVar.nAct03TestType = 1 Then
        ActDa = ActNo(2).DA_NO
        dActVolt = SetupVar.dAct03SetVolt(Index)
        ManuVar.nAct03Pos = Index
        
        Call OutDa(ActDa, dActVolt, SysVar.bPercent(ActNo(2).AD_VOLT))
    Else
        ManuVar.nAct03Pos = Index
        nActNo = 2
        nSteppingMode = 0
        
        Select Case Index
            Case 0:
                bSteppingWrite(0) = True
                
                lpWriteStepRotation(0) = STEP_ROTA_P1
                lpWriteStepData(0) = STEP_POS1
                lpWriteStepData(1) = STEP_POS2
                
                lpWriteStepRotation(1) = STEP_NULL
                lpWriteStepData(2) = STEP_NULL
                lpWriteStepData(3) = STEP_NULL
            Case 1, 2:
                bSteppingWrite(0) = True
                
                lpWriteStepRotation(0) = STEP_ROTA_P2
                lpWriteStepData(0) = Dec2Hex(Val2Byte(CM_LO, SetupVar.dAct03SetVolt(Index) - SetupVar.dAct03SetVolt(Index - 1)))
                lpWriteStepData(1) = Dec2Hex(Val2Byte(CM_HI, SetupVar.dAct03SetVolt(Index) - SetupVar.dAct03SetVolt(Index - 1)))
                
                lpWriteStepRotation(1) = STEP_NULL
                lpWriteStepData(2) = STEP_NULL
                lpWriteStepData(3) = STEP_NULL
            Case 3:
                bSteppingWrite(0) = True
                
                lpWriteStepRotation(0) = STEP_ROTA_P2
                lpWriteStepData(0) = STEP_POS1
                lpWriteStepData(1) = STEP_POS2
                
                lpWriteStepRotation(1) = STEP_NULL
                lpWriteStepData(2) = STEP_NULL
                lpWriteStepData(3) = STEP_NULL
        End Select
        
        nSteppingMode = STEP_MODE_MOVE
        
        frmRun.optAct03Manual(Index).Value = False
    End If
End Sub

Private Sub optAct04Manual_Click(Index As Integer)
    Dim i As Integer
    Dim ActDa As Integer
    Dim dActVolt As Double
    Dim lpRotation(1) As String
    Dim lpPos(3) As String
    Dim lpActNo As String
    Dim nPortNo As Integer
    Dim nActNo As Integer
    
    If SetupVar.nAct04TestType = 1 Then
        ActDa = ActNo(3).DA_NO
        dActVolt = SetupVar.dAct04SetVolt(Index)
        ManuVar.nAct04Pos = Index
        
        Call OutDa(ActDa, dActVolt, SysVar.bPercent(ActNo(3).AD_VOLT))
    Else
        ManuVar.nAct04Pos = Index
        nActNo = 3
        nSteppingMode = 0
        
        Select Case Index
            Case 0:
                bSteppingWrite(1) = True
                
                lpWriteStepRotation(0) = STEP_NULL
                lpWriteStepData(0) = STEP_NULL
                lpWriteStepData(1) = STEP_NULL
                
                lpWriteStepRotation(1) = STEP_ROTA_P1
                lpWriteStepData(2) = STEP_POS1
                lpWriteStepData(3) = STEP_POS2
            Case 1, 2:
                bSteppingWrite(1) = True
                
                lpWriteStepRotation(0) = STEP_NULL
                lpWriteStepData(0) = STEP_NULL
                lpWriteStepData(1) = STEP_NULL
                
                lpWriteStepRotation(1) = STEP_ROTA_P2
                lpWriteStepData(2) = Dec2Hex(Val2Byte(CM_LO, SetupVar.dAct04SetVolt(Index) - SetupVar.dAct04SetVolt(Index - 1)))
                lpWriteStepData(3) = Dec2Hex(Val2Byte(CM_HI, SetupVar.dAct04SetVolt(Index) - SetupVar.dAct04SetVolt(Index - 1)))
            Case 3:
                bSteppingWrite(1) = True
                
                lpWriteStepRotation(0) = STEP_NULL
                lpWriteStepData(0) = STEP_NULL
                lpWriteStepData(1) = STEP_NULL
                
                lpWriteStepRotation(1) = STEP_ROTA_P2
                lpWriteStepData(2) = STEP_POS1
                lpWriteStepData(3) = STEP_POS2
        End Select
        
        nSteppingMode = STEP_MODE_MOVE
        
        frmRun.optAct04Manual(Index).Value = False
    End If
End Sub

Private Sub chkSensorManual_Click(Index As Integer)
    Dim nRes As Integer
    
    nRes = 0
    
    If frmRun.chkSensorManual(0).Value = 1 Then nRes = nRes + &H1
    If frmRun.chkSensorManual(1).Value = 1 Then nRes = nRes + &H2
    If frmRun.chkSensorManual(2).Value = 1 Then nRes = nRes + &H4
    If frmRun.chkSensorManual(3).Value = 1 Then nRes = nRes + &H8
    If frmRun.chkSensorManual(4).Value = 1 Then nRes = nRes + &H10
    If frmRun.chkSensorManual(5).Value = 1 Then nRes = nRes + &H20
    
    If nRes = 0 Then
        Call DO_Control(O_SENSOR_POWER, False)
    Else
        Call DO_Control(O_SENSOR_POWER, True)
    End If
    
    ManuVar.nSensorPos = nRes
End Sub

Private Sub Form_Activate()
    nNowForm = FM_RUN
    
    Call LoadLangFile(FM_RUN)
    
    ' graph bug fix
    If RunVar.bLoading = True Then
        RunVar.bLoading = False
    ElseIf RunVar.bLoading = False Then
        If bRunningGraphWin = False Then Call OnStart
        RunVar.bLoading = True
    End If
End Sub

Private Sub Form_Load()
    RunVar.bLoading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call OnEnd
End Sub

Private Sub tmrLoop_Timer()
    Call RunMode
End Sub

Private Sub OnStart()
    fraDebug.Visible = DEBUGMODE
    
    lpPath = App.Path
    
    Call DIDOCtlArray
    Call LoadModelName
    Call ModelChange(cboCarType, Trim$(lpNowModel))
    Call ActBoardChange
    Call DO_Clear
    Call StartLog
    
    ' Auto/Manual ÇÑ¹øÀº ¼±ÅÃÇÏ±â À§ÇÔ
    RunVar.bAutoManual = Not DIS(I_AUTO_SW)
    
    Call ScreenClear(True)
    Call ZeroAD
    Call OnSolRelease
    Call SetTime(TM_PLCPROC)
    
    RunVar.bUpdate = True
    
    RunVar.nMSPos = 0
    RunVar.nMSUse = SysVar.nMSUse
    
    If SetupVar.bProductUse Then
        frmRun.btnPartProduct.Visible = True
        frmRun.btnPartProduct.Caption = Trim$(SetupVar.lpProductName)
    Else
        frmRun.btnPartProduct.Visible = False
    End If
    
    If SetupVar.bModelTypeUse Then
        frmRun.btnPartLR.Visible = True
        frmRun.btnPartLR.Caption = IIf(SetupVar.nModelType = 0, Trim$(SetupVar.lpLHDPartName), Trim$(SetupVar.lpRHDPartName))
    Else
        frmRun.btnPartLR.Visible = False
    End If
    
    frmRun.pnlModelFullName.Caption = Trim$(SysVar.lpOldModelSave)
    frmRun.pnlNGCountValue.Caption = SysVar.nContinueNGQty
    
    Erase bSteppingWrite
    Erase lStepActData
    
    tmrLoop.Enabled = True
    
    Call OnLog("ON START")
End Sub

Private Sub OnEnd()
    SysVar.lpModel = Trim$(cboCarType.Text)
    
    Call OnSolRelease
    Call DO_Clear
    Call Delay(100)
    
    tmrLoop.Enabled = False
    
    Call OnLog("ON END")
    Call EndLog
End Sub

Private Sub DIDOCtlArray()
    Dim i As Integer
    Dim bResDI(MAX_DIO_CHANNEL) As Boolean
    Dim bResDO(MAX_DIO_CHANNEL) As Boolean
    Dim btnCtl As BHImageButton
    
    For Each btnCtl In frmRun.btnDI
        bResDI(btnCtl.Index) = True
    Next
    
    For Each btnCtl In frmRun.btnDO
        bResDO(btnCtl.Index) = True
    Next
    
    For i = 0 To MAX_DIO_CHANNEL
        If bResDI(i) = False Then Call Load(frmRun.btnDI(i))
        If bResDO(i) = False Then Call Load(frmRun.btnDO(i))
    Next
End Sub

Private Sub RunMode()
    ' DISPLAY
    If RunVar.nDispCounter > DISP_TIME Then
        RunVar.nDispCounter = 0
        RunVar.bDispFlash = True
    End If
    
    If SysVar.bPlcCommUse And PLCUSE Then
        If ElapseTime(TM_PLCPROC) > 0.5 Then
            Call PlcStatus
            
            If RunVar.bRun = False Then
                Call PLC_Proc
            End If
            
            Call SetTime(TM_PLCPROC)
        End If
    End If
    
    Call AlwaysMode ' ALWAYS CHECKING
    Call CalSet ' mastersampletest
    
    If RunVar.bAutoManual <> DIS(I_AUTO_SW) Then
        RunVar.bAutoManual = DIS(I_AUTO_SW)
        
        If RunVar.bAutoManual Then
            Call OnLog("AUTO MODE...")
            Call ButtonManual(False)
            Call PartVisible(False)
        Else
            Call OnLog("MANUAL MODE...")
            Call ButtonManual(True)
            Call PartVisible(True)
            
            ManuVar.nBlowerPos = MANUAL_INIT
            ManuVar.nAct01Pos = MANUAL_INIT
            ManuVar.nAct02Pos = MANUAL_INIT
            ManuVar.nAct03Pos = MANUAL_INIT
            ManuVar.nAct04Pos = MANUAL_INIT
            ManuVar.nLinAct01Pos = MANUAL_INIT
            ManuVar.nLinAct02Pos = MANUAL_INIT
            ManuVar.nLinAct03Pos = MANUAL_INIT
            ManuVar.nLinAct04Pos = MANUAL_INIT
            ManuVar.nSensorPos = MANUAL_INIT
            ManuVar.nNvhPos = MANUAL_INIT
            ManuVar.nIonPos = MANUAL_INIT
        End If
        
        RunVar.nTestPos = 0
        RunVar.bRun = False
        
        Call ScreenClear(True, False)
        Call MsgLog(MSG_READY)
        Call DO_Clear
        Call OnSolRelease
    End If
    
    ' MASTER SAMPLE
    If RunVar.bMSStart Then
        Call MasterSampleTest
    End If
    
    ' EACH MODE
    If RunVar.bAutoManual Then
        If bPlcStartSig And PLCUSE Then
            If SysVar.nMSUse = 2 And (Trim$(cboCarType.Text) <> lpNowModel) Then
                Call MasterSampleTest
            End If
        End If
        
        Call AutoMode
    Else
        Call ManualMode
    End If
    
    ' DISPLAY
    If RunVar.bDispFlash Then
        RunVar.bDispFlash = False
    End If  ' bDisp
    
    RunVar.nDispCounter = RunVar.nDispCounter + 1
End Sub

Private Sub DispCounter()
    Call GetDisplayResultData
    Call GetStatistical(StcalVar)
    
    lblCounter(0).Caption = SysVar.lTotalCounter
    lblCounter(1).Caption = SysVar.lOkCounter
    lblCounter(2).Caption = SysVar.lNgCounter
    
    If SysVar.lTotalCounter <> 0 Then
        lblCounter(3).Caption = Format((1 - (SysVar.lNgCounter / SysVar.lTotalCounter)) * 100, "#0.0")
    Else
        lblCounter(3).Caption = Format("0.0")
    End If
End Sub

Private Sub UseableButton(ByVal bLock As Boolean)
    frmRun.cboCarType.Enabled = bLock
    frmRun.btnReturn.Enabled = bLock
    frmRun.btnClearCounter.Enabled = bLock
    frmRun.btnStatistical.Enabled = bLock
End Sub

Private Sub MasterSampleTest()
    Dim dCal(MAX_AD_CHANNEL) As Double
    Dim lBkColor As Long
    Dim dTime As Double
    Dim bRes As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim dCalTotal As Double
    Dim TempStr(10) As String
    
    Erase TempStr
    
    Select Case MASTERSAMPLELANG
        Case 0: TempStr$(0) = "MASTER SAMPLE"
        Case 1: TempStr$(0) = "CALIBRATION"
    End Select
    
    If RunVar.nMSPos = 0 Then
        Call OnLog(TempStr$(0) & " TEST START...")
        
        RunVar.bMSTest(0) = SysVar.bMSTest(0)
        RunVar.bMSTest(1) = SysVar.bMSTest(1)
        
        If SysVar.bLeakTest Then
            nLeakDummyPos = POS_INIT
            bLeakDummyUse = True
        End If
        
        RunVar.nMSPos = 10
    End If
    
    If RunVar.nMSPos = 10 Then
        Call SetVolt(Trim$(SysVar.dMSVolt))
        Call DO_Clear
        Call ZeroAD
        Call Sleep(300)
        
        RunVar.nMSPos = IIf(RunVar.bMSTest(0) Or RunVar.bMSTest(1), 20, 10002)
    End If
    
    If RunVar.nMSPos = 20 Then
        Call OutDa(ActNo(0).DA_NO, Trim$(SetupVar.dAct01SetVolt(0)), SysVar.bPercent(ActNo(0).AD_VOLT))
        Call OutDa(ActNo(1).DA_NO, Trim$(SetupVar.dAct02SetVolt(0)), SysVar.bPercent(ActNo(1).AD_VOLT))
        Call OutDa(ActNo(2).DA_NO, Trim$(SetupVar.dAct03SetVolt(0)), SysVar.bPercent(ActNo(2).AD_VOLT))
        Call OutDa(ActNo(3).DA_NO, Trim$(SetupVar.dAct04SetVolt(0)), SysVar.bPercent(ActNo(3).AD_VOLT))
        Call Delay(100)
        
        RunVar.nMSPos = 30
    End If
    
    If RunVar.nMSPos = 30 Then
        RunVar.bMSTotalResult = True
        
        Call ScreenClear(False)
        Call MsgLog(MSG_CAL)
        Call UseableButton(False)
        
        Call DO_Control(O_BLOWER_POWER, True)
        
        If SetupVar.nBlowerType = 1 Then
            Call DO_Control(O_BLOWER_PWM, True)
        End If
        
        Call DO_Control(ActNo(0).O_POWER, True)
        Call DO_Control(ActNo(1).O_POWER, True)
        Call DO_Control(ActNo(2).O_POWER, True)
        Call DO_Control(ActNo(3).O_POWER, True)
        Call DO_Control(O_SENSOR_POWER, True)
        
'        If RunVar.bMSTest(0) Then
'            Call DO_Control(O_MSOK, True)
'            Call DO_Control(O_MSNG, False)
'        ElseIf RunVar.bMSTest(1) Then
'            Call DO_Control(O_MSOK, False)
'            Call DO_Control(O_MSNG, True)
'        End If
        
        If RunVar.bMSTest(0) Then
            Call DO_Control(O_MSOK, True)
'            Call DO_Control(O_MSNG, True)
        End If
        
        For i = 0 To MAX_AD_CHANNEL
            RunVar.bMSResult(i) = SysVar.bOffsetMSUse(i)
        Next
        
        Call SetTime(TM_MS)
        RunVar.nMSPos = 100
    End If
    
    If RunVar.nMSPos = 100 Then
        dTime = ElapseTime(TM_MS)
        
        For i = 0 To MAX_AD_CHANNEL
            dCal(i) = ADRead(i)
        Next
        
        ' Display
        If RunVar.bDispFlash Then
            ' intake
            If RunVar.bMSResult(AD_ACT04_CURR) Then
                Me.pnlAct04Curr(0).Caption = Format(dCal(AD_ACT04_CURR), SysVar.lpUnit(AD_ACT04_CURR))
                Me.pnlAct04Volt(0).Caption = Format(dCal(AD_ACT04_VOLT), SysVar.lpUnit(AD_ACT04_VOLT))
            End If
        End If
        
        For i = 1 To MAX_AD_CHANNEL
            If RunVar.bMSResult(i) Then
                If dTime >= SysVar.dOffsetMSDelay(i) Then
                    RunVar.bMSResult(i) = False
                    
                    ' filtering = 10
                    dCalTotal = 0
                    
                    For j = 0 To SysVar.nFiltering - 1
                        dCalTotal = dCalTotal + dCal(i)
                    Next
                    
                    dCal(i) = dCalTotal / SysVar.nFiltering
                    dCal(i) = Format(dCal(i), SysVar.lpUnit(i))
                    
                    If RunVar.bMSTest(0) Then
                        If dCal(i) >= SysVar.dOffsetMSMin(i) And dCal(i) <= SysVar.dOffsetMSMax(i) Then
                            lBkColor = vbGreen
                        Else
                            lBkColor = vbRed
                            RunVar.bMSTotalResult = False
                        End If
                    ElseIf RunVar.bMSTest(1) Then
                        If dCal(i) < SysVar.dOffsetMSMin(i) Or dCal(i) > SysVar.dOffsetMSMax(i) Then
                            lBkColor = vbRed
                        Else
                            lBkColor = vbGreen
                            RunVar.bMSTotalResult = False
                        End If
                    End If
                    
                    Select Case i
                        Case AD_ACT04_CURR:     frmRun.pnlAct04Curr(0).BackColor = lBkColor
                        Case AD_ACT04_VOLT:     frmRun.pnlAct04Volt(0).BackColor = lBkColor
                    End Select
                End If
            End If
        Next
        
        bRes = True
        
        For i = 1 To MAX_AD_CHANNEL
            If RunVar.bMSResult(i) = True Then
                bRes = False
                Exit For
            End If
        Next
        
        If bRes Then
            If SysVar.bLeakTest Then
                RunVar.nMSPos = 200
            Else
                RunVar.nMSPos = 300
            End If
        End If
    End If
    
    If RunVar.nMSPos = 200 Then
        If nLeakDummyPos = 0 Then
            RunVar.nMSPos = 300
        End If
    End If
    
    If RunVar.nMSPos = 300 Then
        If RunVar.bMSTest(0) Then
            lBkColor = vbRed
        ElseIf RunVar.bMSTest(1) Then
            lBkColor = vbGreen
        End If
        
        Call GetDisplayResultData
        
        DataVar.lpSerialNo = "----"
        DataVar.lpModel = TempStr$(0)
        DataVar.lpTime = Format(time, "HH:MM:SS")
        
        If RunVar.bMSTest(0) Then
            RunVar.bMSJudge(0) = RunVar.bMSTotalResult
            RunVar.bMSTest(0) = False
            DataVar.lpResult = IIf(RunVar.bMSTotalResult, LangVar.Msg(MSG_OKPASS), LangVar.Msg(MSG_OKFAIL))
            
            Call MsgLog(IIf(RunVar.bMSTotalResult, MSG_OKPASS, MSG_OKFAIL))
        ElseIf RunVar.bMSTest(1) Then
            RunVar.bMSJudge(1) = RunVar.bMSTotalResult
            RunVar.bMSTest(1) = False
            DataVar.lpResult = IIf(RunVar.bMSTotalResult, LangVar.Msg(MSG_NGPASS), LangVar.Msg(MSG_NGFAIL))
            
            Call MsgLog(IIf(RunVar.bMSTotalResult, MSG_NGPASS, MSG_NGFAIL))
        End If
        
        Call SetTime(TM_MS)
        
        RunVar.nMSPos = 400
    End If

    If RunVar.nMSPos = 400 Then
        If ElapseTime(TM_MS) > SysVar.nMSAfterDelay Then
            Call SaveDataFile(SysVar.lpSaveFileName)
            
            RunVar.nMSPos = 900
        End If
    End If
    
    If RunVar.nMSPos = 900 Then
        RunVar.nMSPos = IIf(RunVar.bMSTest(0) Or RunVar.bMSTest(1), 10, 1000)
    End If
    
    If RunVar.nMSPos = 1000 Then
        Call OnLog(TempStr$(0) & " TEST END...")
        Call DO_Clear
        
        If SysVar.bMSTest(0) = False Then RunVar.bMSJudge(0) = True
        If SysVar.bMSTest(1) = False Then RunVar.bMSJudge(1) = True
        
        RunVar.nMSPos = IIf(RunVar.bMSJudge(0) = False Or RunVar.bMSJudge(1) = False, 2000, 10000)
    End If
    
    If RunVar.nMSPos = 2000 Then
        frmRun.pnlMS.Visible = True
        
        frmRun.btnMS(0).Caption = LangVar.MsgYes
        frmRun.btnMS(1).Caption = LangVar.MsgNo
        
        frmRun.MSOKFail.BackColor = IIf(RunVar.bMSJudge(0), vbGreen, vbRed)
        frmRun.MSOKFail.Caption = IIf(RunVar.bMSJudge(0), LangVar.Msg(MSG_OKPASS), LangVar.Msg(MSG_OKFAIL))
        frmRun.MSOKFail.Font.Size = 14
        frmRun.MSNGFail.BackColor = IIf(RunVar.bMSJudge(1), vbRed, vbGreen)
        frmRun.MSNGFail.Caption = IIf(RunVar.bMSJudge(1), LangVar.Msg(MSG_NGPASS), LangVar.Msg(MSG_NGFAIL))
        frmRun.MSNGFail.Font.Size = 14
        
        frmRun.pnlMS.Top = frmRun.Height / 2 - frmRun.pnlMS.Height / 2
        frmRun.pnlMS.Left = frmRun.Width / 2 - frmRun.pnlMS.Width / 2
        
        RunVar.nMSPos = 3000
    End If
    
    If RunVar.nMSPos = 3000 Then
        Select Case MSFlag
            Case 1: RunVar.nMSPos = 10001
            Case 2: RunVar.nMSPos = 10000
        End Select
    End If
    
    If RunVar.nMSPos >= 10000 Then
        MSFlag = 0
        RunVar.nMSUse = 0
        RunVar.bMSStart = False
        frmRun.pnlMS.Visible = False
        
        Select Case RunVar.nMSPos
            Case 10000:
                RunVar.nMSPos = 0
                
                Call ScreenClear(True)
                Call MsgLog(MSG_READY)
                Call UseableButton(True)
                
                bMSTestStart = False
            
            Case 10001:
                RunVar.nMSPos = 0
                
                Call Unload(frmRun)
                
                bMSTestStart = False
            
            Case 10002:
                RunVar.nMSPos = 0
                
                Call ScreenClear(True)
                Call MsgLog(MSG_READY)
                Call UseableButton(True)
                Call OnLog(TempStr$(0) & "TEST CANCEL...")
                
                bMSTestStart = False
        
        End Select
    End If
End Sub

Private Sub CalSet()
    nMSTime = Format(time, "HHMM")
    
    Select Case SysVar.nMSUse
        Case 1:
            If nMSTime >= nNowTime + (SysVar.nMS2Time * IIf(DEBUGMODE, 3, 100)) Then
                bMSTestStart = True
                nNowTime = Format(time, "HHMM")
            End If
        
        Case 3:
            If nMSTime >= 0 And nMSTime < 1200 Then
                If bMSTemp(0) = False Then
                    If nTime(0) = nMSTime Then
                        bMSTestStart = True
                        bMSTemp(0) = True
                    End If
                End If
                
                If nTime(0) <> nMSTime Then
                    bMSTemp(0) = False
                End If
            ElseIf nMSTime >= 1200 And nMSTime < 2400 Then
                If bMSTemp(1) = False Then
                    If nTime(1) = nMSTime Then
                        bMSTestStart = True
                        bMSTemp(1) = True
                    End If
                End If
                
                If nTime(1) <> nMSTime Then
                    bMSTemp(1) = False
                End If
            End If
    
    End Select
End Sub

Private Sub AlwaysMode()
    If RunVar.bDispFlash Then
        RunVar.lpNowDate = Format(Date, "YYYYMMDD")
        RunVar.bUpdateDate = IIf(RunVar.lpNowDate = SysVar.lpSaveDate, False, True)
        
        If RunVar.bUpdateDate And RunVar.bRun = False Then
            RunVar.bUpdateDate = False
            
            SysVar.lpSaveDate = RunVar.lpNowDate
            
'            SysVar.lTotalCounter = 0
'            SysVar.lOkCounter = 0
'            SysVar.lNgCounter = 0
            
            RunVar.bUpdate = True
        End If
        
        If RunVar.bUpdate Then
            RunVar.bUpdate = False
            
            SysVar.lpSaveFileName = MakeFolder & "\\" & SetupVar.lpFileName & "_" & SysVar.lpSaveDate & ".CSV"
            SysVar.lpSaveNgFileName = MakeFolder & "\\NGFiles\\" & SetupVar.lpFileName & "_NG_" & SysVar.lpSaveDate & ".CSV"
            
            Call ScreenClear(False)
            Call DispCounter
        End If
    End If
End Sub

Private Sub ScreenClear(ByVal bMessageClear As Boolean, Optional bLoad As Boolean = True)
    Dim lBkColor As Long
    Dim i As Integer
    
    ' Blower
    lBkColor = IIf(SetupVar.bBlowerUse, vbWhite, CO_NONE)
    
    For i = 0 To frmRun.pnlBlowerCurr.UBound
        frmRun.optBlowerManual(i).Caption = Trim$(SetupVar.lpBlowerName(i))
        frmRun.pnlBlowerName(i).Caption = Trim$(SetupVar.lpBlowerName(i))
        
        frmRun.pnlBlowerCurr(i).BackColor = lBkColor
        frmRun.pnlBlowerCurr(i).Caption = ""
        frmRun.pnlBlowerTime(i).BackColor = lBkColor
        frmRun.pnlBlowerTime(i).Caption = ""
        frmRun.pnlBlowerResult(i).BackColor = lBkColor
        frmRun.pnlBlowerResult(i).Caption = ""
    Next
    
    ' RPM
    lBkColor = IIf(SetupVar.bBlowerUse, vbWhite, CO_NONE)
    
    frmRun.pnlRpmName.Caption = Trim$(SetupVar.lpRpmName)
    frmRun.pnlRpmCurr.BackColor = lBkColor
    frmRun.pnlRpmCurr.Caption = ""
    frmRun.pnlRpmResult.BackColor = lBkColor
    frmRun.pnlRpmResult.Caption = ""
    
    ' Vib
    lBkColor = IIf(SetupVar.bVibUse, vbWhite, CO_NONE)
    
    frmRun.pnlVibName.Caption = Trim$(SetupVar.lpVibName)
    frmRun.pnlVibCurr.BackColor = lBkColor
    frmRun.pnlVibCurr.Caption = ""
    frmRun.pnlVibResult.BackColor = lBkColor
    frmRun.pnlVibResult.Caption = ""
    
    ' Act01
    frmRun.pnlAct01Item(0).Caption = Trim$(SetupVar.lpActName(0))
    lBkColor = IIf(SetupVar.bAct01Use, vbWhite, CO_NONE)
    
    If SetupVar.nAct01TestType = 0 Then
        frmRun.pnlAct01Item(2).Caption = "STEP"
    Else
        frmRun.pnlAct01Item(2).Caption = "VOLT (V)"
    End If
    
    If SetupVar.nAct01TestType = 0 Then
        For i = 0 To frmRun.pnlAct01Result.UBound
            frmRun.pnlAct01Result(i).Visible = True
        Next
        
        frmRun.pnlAct01StallDelta.Visible = False
        frmRun.pnlAct01Item(5).Visible = False
        frmRun.pnlAct01TotalTime.Visible = False
    Else
        For i = 0 To frmRun.pnlAct01Result.UBound
            frmRun.pnlAct01Result(i).Visible = False
        Next
        
        frmRun.pnlAct01StallDelta.Visible = True
        frmRun.pnlAct01Item(5).Visible = True
        frmRun.pnlAct01TotalTime.Visible = True
    End If
    
    frmRun.pnlAct01StallDelta.BackColor = lBkColor
    frmRun.pnlAct01StallDelta.Caption = ""
    
    frmRun.pnlAct01TotalTime.BackColor = lBkColor
    frmRun.pnlAct01TotalTime.Caption = ""
    
    For i = 0 To frmRun.optAct01Manual.UBound
        frmRun.optAct01Manual(i).Caption = Trim$(SetupVar.lpAct01Name(i))
    Next
    
    For i = 0 To frmRun.pnlAct01Curr.UBound
        frmRun.pnlAct01Name(i).Caption = Trim$(SetupVar.lpAct01Name(i))
        
        frmRun.pnlAct01Curr(i).BackColor = lBkColor
        frmRun.pnlAct01Curr(i).Caption = ""
        frmRun.pnlAct01Volt(i).BackColor = lBkColor
        frmRun.pnlAct01Volt(i).Caption = ""
        frmRun.pnlAct01Time(i).BackColor = lBkColor
        frmRun.pnlAct01Time(i).Caption = ""
        frmRun.pnlAct01Result(i).BackColor = lBkColor
        frmRun.pnlAct01Result(i).Caption = ""
    Next
    
    ' Act02
    frmRun.pnlAct02Item(0).Caption = Trim$(SetupVar.lpActName(1))
    lBkColor = IIf(SetupVar.bAct02Use, vbWhite, CO_NONE)
    
    If SetupVar.nAct02TestType = 0 Then
        frmRun.pnlAct02Item(2).Caption = "STEP"
    Else
        frmRun.pnlAct02Item(2).Caption = "VOLT (V)"
    End If
    
    If SetupVar.nAct02TestType = 0 Then
        For i = 0 To frmRun.pnlAct02Result.UBound
            frmRun.pnlAct02Result(i).Visible = True
        Next
        
        frmRun.pnlAct02StallDelta.Visible = False
    Else
        For i = 0 To frmRun.pnlAct02Result.UBound
            frmRun.pnlAct02Result(i).Visible = False
        Next
        
        frmRun.pnlAct02StallDelta.Visible = True
    End If
    
    frmRun.pnlAct02StallDelta.BackColor = lBkColor
    frmRun.pnlAct02StallDelta.Caption = ""
    
    For i = 0 To frmRun.optAct02Manual.UBound
        frmRun.optAct02Manual(i).Caption = Trim$(SetupVar.lpAct02Name(i))
    Next
    
    For i = 0 To frmRun.pnlAct02Curr.UBound
        frmRun.pnlAct02Name(i).Caption = Trim$(SetupVar.lpAct02Name(i))
        
        frmRun.pnlAct02Curr(i).BackColor = lBkColor
        frmRun.pnlAct02Curr(i).Caption = ""
        frmRun.pnlAct02Volt(i).BackColor = lBkColor
        frmRun.pnlAct02Volt(i).Caption = ""
        frmRun.pnlAct02Time(i).BackColor = lBkColor
        frmRun.pnlAct02Time(i).Caption = ""
        frmRun.pnlAct02Result(i).BackColor = lBkColor
        frmRun.pnlAct02Result(i).Caption = ""
    Next
    
    ' Act03
    frmRun.pnlAct03Item(0).Caption = Trim$(SetupVar.lpActName(2))
    lBkColor = IIf(SetupVar.bAct03Use, vbWhite, CO_NONE)
    
    If SetupVar.nAct03TestType = 0 Then
        frmRun.pnlAct03Item(2).Caption = "STEP"
    Else
        frmRun.pnlAct03Item(2).Caption = "VOLT (V)"
    End If
    
    If SetupVar.nAct03TestType = 0 Then
        For i = 0 To frmRun.pnlAct03Result.UBound
            frmRun.pnlAct03Result(i).Visible = True
        Next
        
        frmRun.pnlAct03StallDelta.Visible = False
    Else
        For i = 0 To frmRun.pnlAct03Result.UBound
            frmRun.pnlAct03Result(i).Visible = False
        Next
        
        frmRun.pnlAct03StallDelta.Visible = True
    End If
    
    frmRun.pnlAct03StallDelta.BackColor = lBkColor
    frmRun.pnlAct03StallDelta.Caption = ""
    
    For i = 0 To frmRun.optAct03Manual.UBound
        frmRun.optAct03Manual(i).Caption = Trim$(SetupVar.lpAct03Name(i))
    Next
    
    For i = 0 To frmRun.pnlAct03Curr.UBound
        frmRun.pnlAct03Name(i).Caption = Trim$(SetupVar.lpAct03Name(i))
        
        frmRun.pnlAct03Curr(i).BackColor = lBkColor
        frmRun.pnlAct03Curr(i).Caption = ""
        frmRun.pnlAct03Volt(i).BackColor = lBkColor
        frmRun.pnlAct03Volt(i).Caption = ""
        frmRun.pnlAct03Time(i).BackColor = lBkColor
        frmRun.pnlAct03Time(i).Caption = ""
        frmRun.pnlAct03Result(i).BackColor = lBkColor
        frmRun.pnlAct03Result(i).Caption = ""
    Next
    
    ' Act04
    frmRun.pnlAct04Item(0).Caption = Trim$(SetupVar.lpActName(3))
    lBkColor = IIf(SetupVar.bAct04Use, vbWhite, CO_NONE)
    
    If SetupVar.nAct04TestType = 0 Then
        frmRun.pnlAct04Item(2).Caption = "STEP"
    Else
        frmRun.pnlAct04Item(2).Caption = "VOLT (V)"
    End If
    
    If SetupVar.nAct04TestType = 0 Then
        For i = 0 To frmRun.pnlAct04Result.UBound
            frmRun.pnlAct04Result(i).Visible = True
        Next
        
        frmRun.pnlAct04StallDelta.Visible = False
    Else
        For i = 0 To frmRun.pnlAct04Result.UBound
            frmRun.pnlAct04Result(i).Visible = False
        Next
        
        frmRun.pnlAct04StallDelta.Visible = True
    End If
    
    frmRun.pnlAct04StallDelta.BackColor = lBkColor
    frmRun.pnlAct04StallDelta.Caption = ""
    
    For i = 0 To frmRun.optAct04Manual.UBound
        frmRun.optAct04Manual(i).Caption = Trim$(SetupVar.lpAct04Name(i))
    Next
    
    For i = 0 To frmRun.pnlAct04Curr.UBound
        frmRun.pnlAct04Name(i).Caption = Trim$(SetupVar.lpAct04Name(i))
        
        frmRun.pnlAct04Curr(i).BackColor = lBkColor
        frmRun.pnlAct04Curr(i).Caption = ""
        frmRun.pnlAct04Volt(i).BackColor = IIf(SetupVar.nAct042Pin = 0, lBkColor, CO_NONE)
        frmRun.pnlAct04Volt(i).Caption = ""
        frmRun.pnlAct04Time(i).BackColor = lBkColor
        frmRun.pnlAct04Time(i).Caption = ""
        frmRun.pnlAct04Result(i).BackColor = lBkColor
        frmRun.pnlAct04Result(i).Caption = ""
    Next
    
    ' Sensor
    For i = 0 To frmRun.pnlSensorVolt.UBound
        lBkColor = IIf(SetupVar.bSensorUse(i), vbWhite, CO_NONE)
        
        frmRun.chkSensorManual(i).Caption = Trim$(SetupVar.lpSensorName(i))
        frmRun.pnlSensorName(i).Caption = Trim$(SetupVar.lpSensorName(i))
        frmRun.pnlSensorVolt(i).BackColor = lBkColor
        frmRun.pnlSensorVolt(i).Caption = ""
    Next
    
    If bMessageClear Then
        Call OnLog("")
    End If
    
'    If bScreenLoad And bLoad Then
'        Call ScreenLoad
'    End If
    
    ' vib
'    Call DrawGrid(frmRun.picVib, GraphVar, 0)
End Sub

Private Sub ScreenLoad()
    Dim i As Integer
    
    Select Case DataVar.lpResult
        Case "OK": Call MsgLog(MSG_OK)
        Case "NG": Call MsgLog(MSG_NG)
    End Select
    
    frmRun.pnlSerial.Caption = DataVar.lpSerialNo
    frmRun.pnlCarType(3).Caption = DataVar.lpPallet
    
    ' blower
    For i = 0 To frmRun.pnlBlowerCurr.UBound
        If DataVar.lpBlowerCurr(i) <> "" And DataVar.lpBlowerTime(i) <> "" Then
            frmRun.pnlBlowerCurr(i).BackColor = IIf(InStr(1, DataVar.lpBlowerCurr(i), "#") > 0, vbRed, vbGreen)
            frmRun.pnlBlowerCurr(i).Caption = Replace(DataVar.lpBlowerCurr(i), "#", "")
            frmRun.pnlBlowerTime(i).BackColor = CO_NONE
            frmRun.pnlBlowerTime(i).Caption = Replace(DataVar.lpBlowerTime(i), "#", "")
            
            Select Case frmRun.pnlBlowerCurr(i).BackColor
                Case vbGreen:
                    frmRun.pnlBlowerResult(i).BackColor = vbGreen
                    frmRun.pnlBlowerResult(i).Caption = "OK"
                Case vbRed:
                    frmRun.pnlBlowerResult(i).BackColor = vbRed
                    frmRun.pnlBlowerResult(i).Caption = "NG"
            End Select
        End If
    Next
    
    If DataVar.lpRpm <> "" Then
        frmRun.pnlRpmCurr.BackColor = IIf(InStr(1, DataVar.lpRpm, "#") > 0, vbRed, vbGreen)
        frmRun.pnlRpmCurr.Caption = Replace(DataVar.lpRpm, "#", "")
        frmRun.pnlRpmResult.BackColor = IIf(InStr(1, DataVar.lpRpm, "#") > 0, vbRed, vbGreen)
        frmRun.pnlRpmResult.Caption = IIf(frmRun.pnlRpmResult.BackColor = vbGreen, "OK", "NG")
    End If
    
    If DataVar.lpVib <> "" Then
        frmRun.pnlVibCurr.BackColor = IIf(InStr(1, DataVar.lpVib, "#") > 0, vbRed, vbGreen)
        frmRun.pnlVibCurr.Caption = Replace(DataVar.lpVib, "#", "")
        frmRun.pnlVibResult.BackColor = IIf(InStr(1, DataVar.lpVib, "#") > 0, vbRed, vbGreen)
        frmRun.pnlVibResult.Caption = IIf(frmRun.pnlVibResult.BackColor = vbGreen, "OK", "NG")
    End If
End Sub

Private Sub StartLog()
    'Subclass the "Form", to Capture the Listbox Notification Messages ...
    lPrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SubClassedList)
End Sub

Private Sub EndLog()
    'Release the SubClassing, Very Important to Prevent Crashing!
    Call SetWindowLong(hwnd, GWL_WNDPROC, lPrevWndProc)
End Sub

Private Sub AutoMode()
    If DIS(I_START_SW) Or bGlobalStartSw Or bPlcStartSig Then
        Call SetTime(TM_TOTAL)
        
        If RunVar.bRun = False And RunVar.nMSPos = 0 Then
            Call DisplayWarning
            
            If ElapseTime(TM_AUTO) > 0.1 Then
                If DIS(I_START_SW) Or bGlobalStartSw Then
                    Call OnLog("INPUT START S/W")
                    
                    nStartSignal = 0
                ElseIf bPlcStartSig Then
                    Call OnLog("PLC START SIGNAL")
                    
                    nStartSignal = 1
                End If
                
                If bPlcStartSig Then
                    Call ModelChange(cboCarType, Trim$(lpNowModel))
                    
                    bPlcStartSig = False
                End If
                
                Call ActBoardChange
                
                RunVar.nTestPos = 0
                RunVar.bRun = True
            End If
        End If
    Else
        Call SetTime(TM_AUTO)
    End If
    
    If DIS(I_STOP_SW) Or bGlobalStopSw Or bPlcStopSig Then
        Call DisplayWarning
        
        ' ÀüÃ¼ ¹è°æÀÌ µ¤¿´´Ù ¾º¿© ÁüÀ¸·Î GRAPH°¡ Áö¿öÁø´Ù.
        ' ±×·¡ÇÁ¸¦ ´Ù½Ã ±×¸®±â À§ÇÑ ºÎºÐ
        'Call ScreenClear(True)
        
        If LINUSE Then
            bAutoLinRead = False
            
            Call DO_Control(O_LIN_ACT_POWER, False)
        End If
        
        If RunVar.bRun Then
            Call OnLog("STOP S/W")
            Call MsgLog(MSG_STOP)
            
            RunVar.bRun = False
            RunVar.bStopFlag = True
            
            Call UseableButton(False)
            Call SetVolt(SetupVar.dTestVolt)
            
            If DOS(O_WORK_ON) Then Call DO_Control(O_WORK_ON, False)
            If DOS(O_WORK_OFF) Then Call DO_Control(O_WORK_OFF, False)
            
            ' Stop Alarm
            If DOS(O_START_LAMP) Then Call DO_Control(O_START_LAMP, False)
            If DOS(O_STOP_LAMP) = False Then Call DO_Control(O_STOP_LAMP, True)
            If DOS(O_OK_LAMP) Then Call DO_Control(O_OK_LAMP, False)
            If DOS(O_NG_LAMP) Then Call DO_Control(O_NG_LAMP, False)
            If DOS(O_RUN_LAMP) Then Call DO_Control(O_RUN_LAMP, False)
            If DOS(O_BUZZER) Then Call DO_Control(O_BUZZER, False)
            
            Call DO_Clear
            Call UseableButton(True)
            
            If NVHUSE Then
                Call NvhSend(NVHREMOVE)
            End If
        Else
            If RunVar.bStopFlag = False Then
                RunVar.bStopFlag = True ' Stop S/W °è¼Ó µé¾î¿À´Â °ÍÀ» ¹æÁöÇÏ±â À§ÇØ ÇÑ¹ø¸¸ ´©¸£´Â ±â´É
                
                Call UseableButton(False)
                Call OnLog("STOP S/W RELEASE")
                
                ' ¸±¸®Áî ½ÇÁ¦·Î ½ÃÅ°´Â ºÎºÐ
                If DOS(O_START_LAMP) Then Call DO_Control(O_START_LAMP, False)
                If DOS(O_STOP_LAMP) = False Then Call DO_Control(O_STOP_LAMP, True)
                If DOS(O_OK_LAMP) Then Call DO_Control(O_OK_LAMP, False)
                If DOS(O_NG_LAMP) Then Call DO_Control(O_NG_LAMP, False)
                If DOS(O_RUN_LAMP) Then Call DO_Control(O_RUN_LAMP, False)
                If DOS(O_BUZZER) Then Call DO_Control(O_BUZZER, False)
                
                Call OnSolRelease
                Call MsgLog(MSG_READY)
                Call UseableButton(True)
            End If
        End If
    Else ' Stop S/W¸¦ Up ÇßÀ» ¶§ Release Á¶°ÇÀ» ¼³Á¤ÇÏ¿© ÁØ´Ù.
        RunVar.bStopFlag = False
    End If
    
    ' buzzer 3 sec
    If ElapseTime(TM_BUZZER) > 3 Then
        If DOS(O_BUZZER) Then Call DO_Control(O_BUZZER, False)
    End If
    
    If RunVar.bRun Then
        Call TestRun
    End If
End Sub

Private Sub ManualMode()
    Dim i As Integer
    Dim bDoRes(MAX_DIO_CHANNEL) As Boolean
    Dim bDiRes(MAX_DIO_CHANNEL) As Boolean
    Dim btnCtl As BHImageButton
    
    For Each btnCtl In btnDO
        bDoRes(btnCtl.Index) = True
    Next
    
    For Each btnCtl In btnDI
        bDiRes(btnCtl.Index) = True
    Next
    
    If RunVar.bDispFlash Then
        For i = 0 To MAX_DIO_CHANNEL
            If bDiRes(i) Then
                If DIS(i) Then
                    btnDI(i).BackColor = vbGreen
                Else
                    btnDI(i).BackColor = CO_NONE
                End If
            End If
            
            If bDoRes(i) Then
                If DOS(i) Then
                    btnDO(i).ForeColor = vbRed
                Else
                    btnDO(i).ForeColor = vbBlack
                End If
            End If
        Next
        
        Call ManualTest(ActNo(0).O_POWER)
        Call ManualTest(ActNo(1).O_POWER)
        Call ManualTest(ActNo(2).O_POWER)
        Call ManualTest(ActNo(3).O_POWER)
        Call ManualTest(O_SENSOR_POWER)
    End If
End Sub

Private Sub ManualTest(ByVal OPower As Integer)
    Dim ActPos As Integer
    Dim AdCurr As Integer
    Dim AdVolt As Integer
    Dim ActType As Integer
    Dim lblCurr As Object
    Dim lblVolt As Object
    Dim lVolt As Long
    
    Select Case OPower
        Case O_BLOWER_POWER:
            Set lblCurr = frmRun.pnlBlowerCurr
            
            ActPos = ManuVar.nBlowerPos
            AdCurr = AD_BLOWER_CURR
        
        Case ActNo(0).O_POWER:
            Set lblCurr = Me.pnlAct01Curr
            Set lblVolt = Me.pnlAct01Volt
            
            ActPos = ManuVar.nAct01Pos
            AdCurr = ActNo(0).AD_CURR
            AdVolt = IIf(SetupVar.nAct01TestType = 1, ActNo(0).AD_VOLT, 0)
            ActType = SetupVar.nAct01TestType
            
        Case ActNo(1).O_POWER:
            Set lblCurr = Me.pnlAct02Curr
            Set lblVolt = Me.pnlAct02Volt
            
            ActPos = ManuVar.nAct02Pos
            AdCurr = ActNo(1).AD_CURR
            AdVolt = IIf(SetupVar.nAct02TestType = 1, ActNo(1).AD_VOLT, 1)
            ActType = SetupVar.nAct02TestType
            
        Case ActNo(2).O_POWER:
            Set lblCurr = Me.pnlAct03Curr
            Set lblVolt = Me.pnlAct03Volt
            
            ActPos = ManuVar.nAct03Pos
            AdCurr = ActNo(2).AD_CURR
            AdVolt = IIf(SetupVar.nAct03TestType = 1, ActNo(2).AD_VOLT, 2)
            ActType = SetupVar.nAct03TestType
            
        Case ActNo(3).O_POWER:
            Set lblCurr = Me.pnlAct04Curr
            Set lblVolt = Me.pnlAct04Volt
            
            ActPos = ManuVar.nAct04Pos
            AdCurr = ActNo(3).AD_CURR
            AdVolt = IIf(SetupVar.nAct04TestType = 1, ActNo(3).AD_VOLT, 3)
            ActType = SetupVar.nAct04TestType
            
        Case O_SENSOR_POWER:
            Set lblVolt = Me.pnlSensorVolt
            
            ActPos = ManuVar.nSensorPos
    
    End Select
    
    If DOS(OPower) Then
        If ActPos <> MANUAL_INIT And ActPos >= 0 Then
            Select Case OPower
                Case O_SENSOR_POWER:
                    If (ActPos And &H1) = &H1 Then lblVolt(0).Caption = Format(ADRead(AD_SENSOR1), Trim$(SysVar.lpUnit(AD_SENSOR1)))
                    If (ActPos And &H2) = &H2 Then lblVolt(1).Caption = Format(ADRead(AD_SENSOR2), Trim$(SysVar.lpUnit(AD_SENSOR2)))
                    If (ActPos And &H4) = &H4 Then lblVolt(2).Caption = Format(ADRead(AD_SENSOR3), Trim$(SysVar.lpUnit(AD_SENSOR3)))
                    If (ActPos And &H8) = &H8 Then lblVolt(3).Caption = Format(ADRead(AD_SENSOR4), Trim$(SysVar.lpUnit(AD_SENSOR4)))
                    If (ActPos And &H10) = &H10 Then lblVolt(4).Caption = Format(ADRead(AD_SENSOR5), Trim$(SysVar.lpUnit(AD_SENSOR5)))
                    If (ActPos And &H20) = &H20 Then lblVolt(5).Caption = Format(ADRead(AD_SENSOR6), Trim$(SysVar.lpUnit(AD_SENSOR6)))
                Case ActNo(0).O_POWER, ActNo(1).O_POWER, ActNo(2).O_POWER, ActNo(3).O_POWER:
                    lblCurr(ActPos).Caption = Format(ADRead(AdCurr), Trim$(SysVar.lpUnit(AdCurr)))
                    
                    If ActType = 1 Then
                        lblVolt(ActPos).Caption = Format(ADRead(AdVolt), Trim$(SysVar.lpUnit(AdVolt)))
                    Else
'                        lVolt = Format(StepRead(AdVolt, ActPos), "0")
'                        lVolt = CalcStep(CDbl(lStepActData(AdVolt)), AdVolt, ActPos - 1)
'                        lblVolt(ActPos).Caption = Format(lVolt, "0")

                        If nSteppingMode = STEP_MODE_MOVE Then
                            nSteppingMode = STEP_MODE_READ
                        End If
                        
                        lblVolt(ActPos).Caption = Format(lStepActData(AdVolt), "0")
                    End If
            End Select
        End If
    End If
End Sub

Private Sub ButtonManual(ByVal bVisible As Boolean)
    Dim i As Integer
    
    For i = 0 To pnlManual.UBound
        pnlManual(i).ZOrder 0
        pnlManual(i).Visible = bVisible
    Next
End Sub

Private Sub MsgLog(ByVal nMsg As Integer)
    Dim nFontSize As Integer
    Dim lBkColor As Long
    Dim lpStr As String
    Dim lpTmp() As String
    
    ' default font size = 48
    
    Select Case nMsg
        Case MSG_STOP:      lBkColor = vbRed:       lpStr = "STOP"
        Case MSG_RUN:       lBkColor = vbYellow:    lpStr = "RUN"
        Case MSG_READY:     lBkColor = vbCyan:      lpStr = "READY"
        Case MSG_OK:        lBkColor = vbGreen:     lpStr = "OK"
        Case MSG_NG:        lBkColor = vbRed:       lpStr = "NG"
        Case MSG_CAL:       lBkColor = vbYellow:    lpStr = "CAL"
        Case MSG_ERR:       lBkColor = vbRed:       lpStr = "ERR"
        Case MSG_OKPASS:    lBkColor = vbGreen:     lpStr = "OKPASS"
        Case MSG_OKFAIL:    lBkColor = vbRed:       lpStr = "OKFAIL"
        Case MSG_NGPASS:    lBkColor = vbRed:       lpStr = "NGPASS"
        Case MSG_NGFAIL:    lBkColor = vbGreen:     lpStr = "NGFAIL"
        Case MSG_SCANNER:   lBkColor = vbRed:       lpStr = "SCANNER"
        Case MSG_NGTABLE:   lBkColor = vbRed:       lpStr = "NGTABLE"
    End Select
    
    If LangVar.Msg(nMsg) <> "" Then
        lpStr = LangVar.Msg(nMsg)
    End If
    
    Select Case nMsg
        Case MSG_SCANNER, MSG_NGTABLE:
            nFontSize = 48
            lpTmp = Split(lpStr, " ")
            lpStr = lpTmp(0) & vbCrLf & lpTmp(1)
        
        Case MSG_OKPASS, MSG_OKFAIL, MSG_NGPASS, MSG_NGFAIL:
            nFontSize = 40
            lpTmp = Split(lpStr, " ")
            lpStr = lpTmp(0) & vbCrLf & lpTmp(1) & " " & lpTmp(2) & " " & lpTmp(3)
        
        Case MSG_CAL:
            nFontSize = 40
            lpTmp = Split(lpStr, " ")
            lpStr = lpTmp(0) & vbCrLf & lpTmp(1)
        
        Case MSG_STOP, MSG_RUN, MSG_READY, MSG_ERR:
            nFontSize = 60
        
        Case MSG_OK, MSG_NG:
            nFontSize = 100
        
        Case Else:
            nFontSize = 48
    
    End Select
    
    pnlMessage.Font.Size = nFontSize
    pnlMessage.BackColor = lBkColor
    pnlMessage.Caption = lpStr
End Sub

Private Sub TestRun()
    If IsTestPos(TP_TEST, 0) Then
        nNowForm = FM_RUN
        
        Call ScreenClear(True)
        Call MsgLog(MSG_RUN)
        Call UseableButton(False)
        
        Call DO_Control(O_STEPPING1_RESET, True)
        Call DO_Control(O_STEPPING2_RESET, True)
        Call Delay(500)
        
        Call DO_Control(O_STEPPING1_RESET, False)
        Call DO_Control(O_STEPPING2_RESET, False)
        Call Delay(500)
        
        If SetupVar.nAct01TestType = 0 Or SetupVar.nAct02TestType = 0 Or SetupVar.nAct03TestType = 0 Or SetupVar.nAct04TestType = 0 Then
            Call StepStartInfo
        End If
        
        If PLCUSE And SysVar.bPlcCommUse Then
            Call PLC_TestRun
        End If
        
        Call DO_Clear
        
        nRetestCount = SysVar.nReTest
        
        Call SetTestPos(TP_TEST, 5)
    End If
    
    If IsTestPos(TP_TEST, 5) Then
        If bMSTestStart Then
            Call MasterSampleTest
            
            If bMSTestStart = False And bLeakDummyUse = False Then
                Call SetTestPos(TP_TEST, 10)
            End If
        Else
            Call SetTestPos(TP_TEST, 10)
        End If
    End If
    
    If IsTestPos(TP_TEST, 10) Then
        If TABLETYPE Then
            If DIS(I_WORK_ON) Then
                Call OnLog("RE-TEST START...")
                
                RunVar.bRestart = True
            Else
                Call OnLog("TEST START...")
                Call OnLog("DEPARTMENT TEST POSITION...")
                
                RunVar.bRestart = False
            End If
        End If
        
        If DOS(O_START_LAMP) = False Then Call DO_Control(O_START_LAMP, True)
        If DOS(O_STOP_LAMP) Then Call DO_Control(O_STOP_LAMP, False)
        If DOS(O_OK_LAMP) Then Call DO_Control(O_OK_LAMP, False)
        If DOS(O_NG_LAMP) Then Call DO_Control(O_NG_LAMP, False)
        If DOS(O_RUN_LAMP) = False Then Call DO_Control(O_RUN_LAMP, True)
        
        Call MsgLog(MSG_RUN)
        Call OnLog("MOVE ACTUATOR START POSITION.")
        Call OnLog("SUPPLY VOLT : " & Format(SetupVar.dTestVolt, "#0.00"))
        Call OnVarialbesInitialized
        Call ZeroAD
        Call SetVolt(SetupVar.dTestVolt)
        
        If TABLETYPE Then
            Call SetTestPos(TP_TEST, IIf(RunVar.bRestart, 200, 100))
        Else
            Call SetTestPos(TP_TEST, 2000)
        End If
        
        Call SetTime(TM_RUN)
    End If
    
    If IsTestPos(TP_TEST, 100) Then
        If ProductPartCheck Then
            Call SetTestPos(TP_TEST, 110)
        Else
            Call DisplayWarning("PRODUCT")
            Call SetTestPos(TP_TEST, 20000)
        End If
    End If
    
    If IsTestPos(TP_TEST, 110) Then
        Call DO_Control(O_ACT04_POWER, True)
        Call SetAct04Move(True)
        Call DO_Control(O_WORK_OFF, False)
        Call DO_Control(O_WORK_ON, True)
        Call SetTestPos(TP_TEST, 120)
    End If
    
    If IsTestPos(TP_TEST, 120) Then
        If DIS(I_WORK_ON) Then
            Call OnLog("ARRIVAL TEST POSITION...")
            
            Call DO_Control(O_WORK_ON, False)
            Call DO_Control(O_WORK_OFF, False)
            
            Call SetTestPos(TP_TEST, 2000)
        End If
        
        If ElapseTime(TM_RUN) > 10# Then
            Call OnLog("[ERROR] WORK SOL FAIL !!!")
            Call DO_Control(O_WORK_ON, False)
            Call DO_Control(O_WORK_OFF, False)
            Call MsgLog(MSG_ERR)
            
            RunVar.bRun = False
            Call SetTestPos(TP_TEST, 20000)
        End If
    End If
    
    If IsTestPos(TP_TEST, 200) Then
        Call SetTestPos(TP_TEST, 2000)
    End If
    
    ' ÆÄ·¿ µé¾î°¡°íºÎÅÍ ¾ÈÀü Ã¼Å©
    If RunVar.nTestPos >= 2000 And RunVar.nTestPos <= 6000 And SetupVar.bProductUse Then
        If ProductPartCheck(False) Then
            RunVar.bRun = False
            
            Call DisplayWarning("SAFETY")
        End If
    End If
    
    If IsTestPos(TP_TEST, 2000) Then
        Call OnPartCheck
        Call SetTestPos(TP_TEST, 4000)
    End If
    
    If IsTestPos(TP_TEST, 4000) Then
        Call OnLog("START TESTER...")
        Call SetStepAct01Move
        Call SetStepAct02Move
        Call SetStepAct03Move
        Call SetStepAct04Move
        
        nSteppingMode = STEP_MODE_MOVE
        
        Call Delay(1000)
        Call SetTestPos(TP_TEST, 5100)
    End If
    
    If IsTestPos(TP_TEST, 5100) Then
        Call OnAct01Test
        Call OnAct02Test
        Call OnAct03Test
        Call OnAct04Test
        Call OnSensorTest(6)
        
        If _
            RunVar.bTestEnd(TP_ACT01) And _
            RunVar.bTestEnd(TP_ACT02) And _
            RunVar.bTestEnd(TP_ACT03) And _
            RunVar.bTestEnd(TP_ACT04) And _
            RunVar.bTestEnd(TP_SENSOR) Then
            
            Call SetTestPos(TP_TEST, 6000)
        End If
    End If
    
    If IsTestPos(TP_TEST, 6000) Then
        Call NvhSend(NVHREMOVE)
        
        If RunVar.bFinal Then
            If SetupVar.bBarCodeUse Then
                Call OnLog("RUN BARCODE PRINT.")
                Call BarCodePrint(Format(Now, "DDMMYY HH:MM:SS") + "/" + Format(SysVar.lOkCounter + 1, "0000"))
                
                If TABLETYPE Then
                    frmRun.pnlSerial.Caption = "A" & Format(SysVar.lOkCounter + 1, "0000000")
                End If
            End If
            
            Call OnMarkRelease
            Call OnSolRelease
            Call SetTestPos(TP_TEST, 8000)
        Else
            Call SetTestPos(TP_TEST, 10000)
        End If
    End If
    
    If IsTestPos(TP_TEST, 8000) Then
        If RunVar.bFinal And SCANNERUSE And SetupVar.bScannerUse And SetupVar.bBarCodeUse Then
            If pnlMessage.Caption = "RUN" Then
                Call MsgLog(MSG_SCANNER)
            End If
            
            If ProductPartCheck(False, False) = False Then
                Call DisplayWarning("BARCODE")
            End If
            
            Select Case nScannerResult
                Case 1:
                    nScannerResult = 0
                    
                    Call DisplayWarning
                    Call MsgLog(MSG_RUN)
                    Call SetTestPos(TP_TEST, 10000)
                
                Case 2:
                    nScannerResult = 0
                    
                    If nScannerCount = 0 Then
                        If SetupVar.bDataSave Then
                            SysVar.lTotalCounter = SysVar.lTotalCounter - 1
                            SysVar.lNgCounter = SysVar.lNgCounter - 1
                        End If
                        
                        DataVar.lpMemo = "READ ERR"
                        
                        RunVar.bFinal = False
                        
                        Call DisplayWarning
                        Call MsgLog(MSG_RUN)
                        Call SetTestPos(TP_TEST, 10000)
                    End If
            
            End Select
        Else
            Call SetTestPos(TP_TEST, 10000)
        End If
    End If
    
    If IsTestPos(TP_TEST, 10000) Then
        Call DO_Clear
        
        If RunVar.bFinal Then
            Call OnLog("TOTAL RESULT : OK")
            Call MsgLog(MSG_OK)
            
            If DOS(O_START_LAMP) Then Call DO_Control(O_START_LAMP, False)
            If DOS(O_STOP_LAMP) Then Call DO_Control(O_STOP_LAMP, False)
            If DOS(O_OK_LAMP) = False Then Call DO_Control(O_OK_LAMP, True)
            If DOS(O_NG_LAMP) Then Call DO_Control(O_NG_LAMP, False)
            If DOS(O_RUN_LAMP) Then Call DO_Control(O_RUN_LAMP, False)
            
            If SetupVar.bDataSave Then
                SysVar.lOkCounter = SysVar.lOkCounter + 1
            End If
            
            frmRun.pnlNGCount.Caption = 0
        Else
            Call OnLog("TOTAL RESULT : NG")
            Call MsgLog(MSG_NG)
            
            If DOS(O_START_LAMP) Then Call DO_Control(O_START_LAMP, False)
            If DOS(O_STOP_LAMP) Then Call DO_Control(O_STOP_LAMP, False)
            If DOS(O_OK_LAMP) Then Call DO_Control(O_OK_LAMP, False)
            If DOS(O_NG_LAMP) = False Then Call DO_Control(O_NG_LAMP, True)
            If DOS(O_RUN_LAMP) Then Call DO_Control(O_RUN_LAMP, False)
            
            If DOS(O_BUZZER) = False Then Call DO_Control(O_BUZZER, True)
            
            Call SetTime(TM_BUZZER)
            
            If SetupVar.bDataSave Then
                SysVar.lNgCounter = SysVar.lNgCounter + 1
            End If
            
            frmRun.pnlNGCount.Caption = Val(frmRun.pnlNGCount.Caption) + 1
        End If
        
        If SetupVar.bDataSave Then
            SysVar.lTotalCounter = SysVar.lTotalCounter + 1
            
            Call PlcLeakGet
            
            Call DispCounter
            Call PlusStatistical(Trim$(UCase(Me.cboCarType.List(Me.cboCarType.ListIndex))), RunVar.bFinal)
            Call SaveDataFile(SysVar.lpSaveFileName) ' Total files.
            
            If RunVar.bFinal = False Then
                Call SaveDataFile(SysVar.lpSaveNgFileName) ' Another NG files.
            End If
        End If
        
        Call SetTime(TM_RUN)
        
        If nRetestCount > 0 And RunVar.bFinal = False Then
            nRetestCount = nRetestCount - 1
            
            Call ScreenClear(True)
            Call SetTestPos(TP_TEST, 10)
        Else
            Call SetTestPos(TP_TEST, 20000)
        End If
    End If
    
    ' Final
    If IsTestPos(TP_TEST, 20000) Then
        If PLCUSE And SysVar.bPlcCommUse Then Call PLC_TestEnd
        
        RunVar.bRun = False
        
        Call OnLog("TEST END...")
        Call OnLog("TEST TIME : " & Format(ElapseTime(TM_TOTAL), "#0.0"))
        
        If SysVar.bCaptureSend And RunVar.bFinal = False Then
            Call OnLog("SCREEN CAPTURE SEND...")
            Call Delay(500)
            Call SaveCapture
            Call ImageFileSend
        End If
        
        Call UseableButton(True)
    End If
End Sub

Private Sub OnVarialbesInitialized()
    Dim i As Integer
    
    ' À§Ä¡ ¹× ÆÇÁ¤ ÃÊ±âÈ­
    RunVar.nBlowerPos = POS_INIT
    RunVar.nAct01Pos = POS_INIT
    RunVar.nAct02Pos = POS_INIT
    RunVar.nAct03Pos = POS_INIT
    RunVar.nAct04Pos = POS_INIT
    RunVar.nSensorPos = POS_INIT
    RunVar.nIonPos = POS_INIT
    RunVar.nNvhPos = POS_INIT
    RunVar.nLinAct01Pos = POS_INIT
    RunVar.nLinAct02Pos = POS_INIT
    RunVar.nLinAct03Pos = POS_INIT
    RunVar.nLinAct04Pos = POS_INIT
    RunVar.nLinPtcPos = POS_INIT
    RunVar.nLinBlowerPos = POS_INIT
    
    ' º¯¼ö ÃÊ±âÈ­
    RunVar.bBlowerUse = False
    RunVar.bVibUse = False
    RunVar.bRpmUse = False
    RunVar.bAct01Use = False
    RunVar.bAct02Use = False
    RunVar.bAct03Use = False
    RunVar.bAct04Use = False
    RunVar.bIonUse = False
    RunVar.bNvhUse = False
    RunVar.bLinAct01Use = False
    RunVar.bLinAct02Use = False
    RunVar.bLinAct03Use = False
    RunVar.bLinAct04Use = False
    RunVar.bLinPtcUse = False
    RunVar.bLinBlowerUse = False
    
    Erase RunVar.bSensorUse
    Erase RunVar.bDoorUse
    Erase RunVar.bDoorStatus
    Erase RunVar.bDoorResult
    
    RunVar.bBlowerUse = SetupVar.bBlowerUse
    RunVar.bVibUse = SetupVar.bVibUse
    RunVar.bRpmUse = SetupVar.bBlowerUse
    RunVar.bAct01Use = SetupVar.bAct01Use
    RunVar.bAct02Use = SetupVar.bAct02Use
    RunVar.bAct03Use = SetupVar.bAct03Use
    RunVar.bAct04Use = SetupVar.bAct04Use
    RunVar.bIonUse = SetupVar.bIonUse
    RunVar.bNvhUse = SetupVar.bNvhUse
    RunVar.bLinAct01Use = SetupVar.bLinActUse(0)
    RunVar.bLinAct02Use = SetupVar.bLinActUse(1)
    RunVar.bLinAct03Use = SetupVar.bLinActUse(2)
    RunVar.bLinAct04Use = SetupVar.bLinActUse(3)
    RunVar.bLinPtcUse = SetupVar.bPTCUse
    RunVar.bLinBlowerUse = SetupVar.bLinBlowerUse
    
    For i = 0 To UBound(SetupVar.bSensorUse)
        RunVar.bSensorUse(i) = SetupVar.bSensorUse(i)
    Next
    
    ' ÀçÅ×½ºÆ® ÃÊ±âÈ­
    RunVar.bReBlowerUse = False
    RunVar.bReAct01Use = False
    RunVar.bReAct02Use = False
    RunVar.bReAct03Use = False
    RunVar.bReAct04Use = False
    RunVar.bReIonUse = False
    RunVar.bReVisionUse = False
    RunVar.bReNvhUse = False
    RunVar.bReLinAct01Use = False
    RunVar.bReLinAct02Use = False
    RunVar.bReLinAct03Use = False
    RunVar.bReLinAct04Use = False
    RunVar.bReLinPtcUse = False
    RunVar.bReLinBlowerUse = False
    
    Erase RunVar.bReSensorUse
    Erase RunVar.bReLeakUse
    
    Erase dAct01CurrBuf
    Erase dAct02CurrBuf
    Erase dAct03CurrBuf
    Erase dAct04CurrBuf
    Erase dBlowerCurrBuf
    
    Erase RunVar.bTestEnd
    
    nLinAutoAddressFlag = 0
    
    Erase RunVar.bLinDataResult
    Erase RunVar.lLinDataCP
    Erase RunVar.bLinCheckPoint
    Erase RunVar.dLinCPTime
    Erase RunVar.lLinDataFinal
    Erase RunVar.lLinDataMove
    Erase RunVar.bLinReadRes
    Erase RunVar.bLinRefPos
    
    Erase DataVar.lpAct04Curr
    Erase DataVar.lpAct04Volt
    Erase DataVar.lpAct04Time
    
    Erase DataVar.lpIon
    Erase DataVar.lpSensor
    
    DataVar.lpVision = ""
    
    DataVar.lpPartOK = ""
    DataVar.lpPartNG = ""
    
    Erase DataVar.lpDoor
    
    DataVar.lpMemo = ""
    
    Erase lpSteppingData
    Erase lpReadStepData
    Erase bActStallBit
    
    nSteppingMode = 0
    
    RunVar.bFinal = True
End Sub

Private Sub OnBlowerTest()
    If RunVar.bBlowerUse Then
        RunVar.bTestEnd(TP_BLOWER) = BlowerTest
    Else
        RunVar.bTestEnd(TP_BLOWER) = True
    End If
End Sub

Private Sub OnAct01Test()
    If RunVar.bAct01Use Then
        RunVar.bTestEnd(TP_ACT01) = Act01Test
    Else
        RunVar.bTestEnd(TP_ACT01) = True
    End If
End Sub

Private Sub OnAct02Test()
    If RunVar.bAct02Use Then
        RunVar.bTestEnd(TP_ACT02) = Act02Test
    Else
        RunVar.bTestEnd(TP_ACT02) = True
    End If
End Sub

Private Sub OnAct03Test()
    If RunVar.bAct03Use Then
        RunVar.bTestEnd(TP_ACT03) = Act03Test
    Else
        RunVar.bTestEnd(TP_ACT03) = True
    End If
End Sub

Private Sub OnAct04Test()
    If RunVar.bAct04Use Then
        RunVar.bTestEnd(TP_ACT04) = Act04Test
    Else
        RunVar.bTestEnd(TP_ACT04) = True
    End If
End Sub

Private Sub OnSensorTest(ByVal nSensorValue As Integer)
    Dim i As Integer
    Dim bRes As Boolean
    
    For i = 0 To nSensorValue - 1
        If RunVar.bSensorUse(i) Then
            bRes = True
            
            Exit For
        End If
    Next
    
    If bRes Then
        RunVar.bTestEnd(TP_SENSOR) = SensorTest(nSensorValue)
    Else
        RunVar.bTestEnd(TP_SENSOR) = True
    End If
End Sub

Private Sub OnIonTest()
    If RunVar.bIonUse Then
        RunVar.bTestEnd(TP_ION) = IonTest
    Else
        RunVar.bTestEnd(TP_ION) = True
    End If
End Sub

Private Sub GetDisplayResultData()
    Dim i As Integer
    
    ' GENERAL
    DataVar.lpModel = Trim$(frmRun.cboCarType.Text)
    DataVar.lpSerialNo = frmRun.pnlSerial.Caption
    DataVar.lpTime = Format(time, "HH:MM:SS")
    DataVar.lpResult = frmRun.pnlMessage.Caption
    
    ' BLOWER
    For i = 0 To 8
        DataVar.lpBlowerCurr(i) = frmRun.pnlBlowerCurr(i).Caption
        If frmRun.pnlBlowerCurr(i).BackColor = vbRed Then DataVar.lpBlowerCurr(i) = "#" & DataVar.lpBlowerCurr(i)
        DataVar.lpBlowerTime(i) = frmRun.pnlBlowerTime(i).Caption
        If frmRun.pnlBlowerTime(i).BackColor = vbRed Then DataVar.lpBlowerTime(i) = "#" & DataVar.lpBlowerTime(i)
    Next
    
    DataVar.lpRpm = frmRun.pnlRpmCurr.Caption
    If frmRun.pnlRpmCurr.BackColor = vbRed Then DataVar.lpRpm = "#" & DataVar.lpRpm
    
    ' ACT01
    For i = 0 To frmRun.pnlAct01Curr.UBound
        DataVar.lpAct01Curr(i) = frmRun.pnlAct01Curr(i).Caption
        If frmRun.pnlAct01Curr(i).BackColor = vbRed Then DataVar.lpAct01Curr(i) = "#" & DataVar.lpAct01Curr(i)
        
        DataVar.lpAct01Volt(i) = frmRun.pnlAct01Volt(i).Caption
        If frmRun.pnlAct01Volt(i).BackColor = vbRed Then DataVar.lpAct01Volt(i) = "#" & DataVar.lpAct01Volt(i)
        
        DataVar.lpAct01Time(i) = frmRun.pnlAct01Time(i).Caption
        If frmRun.pnlAct01Time(i).BackColor = vbRed Then DataVar.lpAct01Time(i) = "#" & DataVar.lpAct01Time(i)
    Next
    
    DataVar.lpAct01Stall(0) = frmRun.pnlAct01StallDelta.Caption
    If frmRun.pnlAct01StallDelta.BackColor = vbRed Then DataVar.lpAct01Stall(0) = "#" & DataVar.lpAct01Stall(0)
    
    ' ACT02
    For i = 0 To frmRun.pnlAct02Curr.UBound
        DataVar.lpAct02Curr(i) = frmRun.pnlAct02Curr(i).Caption
        If frmRun.pnlAct02Curr(i).BackColor = vbRed Then DataVar.lpAct02Curr(i) = "#" & DataVar.lpAct02Curr(i)
        
        DataVar.lpAct02Volt(i) = frmRun.pnlAct02Volt(i).Caption
        If frmRun.pnlAct02Volt(i).BackColor = vbRed Then DataVar.lpAct02Volt(i) = "#" & DataVar.lpAct02Volt(i)
        
        DataVar.lpAct02Time(i) = frmRun.pnlAct02Time(i).Caption
        If frmRun.pnlAct02Time(i).BackColor = vbRed Then DataVar.lpAct02Time(i) = "#" & DataVar.lpAct02Time(i)
    Next
    
    DataVar.lpAct02Stall(0) = frmRun.pnlAct02StallDelta.Caption
    If frmRun.pnlAct02StallDelta.BackColor = vbRed Then DataVar.lpAct02Stall(0) = "#" & DataVar.lpAct02Stall(0)
    
    ' ACT03
    For i = 0 To frmRun.pnlAct03Curr.UBound
        DataVar.lpAct03Curr(i) = frmRun.pnlAct03Curr(i).Caption
        If frmRun.pnlAct03Curr(i).BackColor = vbRed Then DataVar.lpAct03Curr(i) = "#" & DataVar.lpAct03Curr(i)
        
        DataVar.lpAct03Volt(i) = frmRun.pnlAct03Volt(i).Caption
        If frmRun.pnlAct03Volt(i).BackColor = vbRed Then DataVar.lpAct03Volt(i) = "#" & DataVar.lpAct03Volt(i)
        
        DataVar.lpAct03Time(i) = frmRun.pnlAct03Time(i).Caption
        If frmRun.pnlAct03Time(i).BackColor = vbRed Then DataVar.lpAct03Time(i) = "#" & DataVar.lpAct03Time(i)
    Next
    
    DataVar.lpAct03Stall(0) = frmRun.pnlAct03StallDelta.Caption
    If frmRun.pnlAct03StallDelta.BackColor = vbRed Then DataVar.lpAct03Stall(0) = "#" & DataVar.lpAct03Stall(0)
    
    ' ACT04
    For i = 0 To frmRun.pnlAct04Curr.UBound
        DataVar.lpAct04Curr(i) = frmRun.pnlAct04Curr(i).Caption
        If frmRun.pnlAct04Curr(i).BackColor = vbRed Then DataVar.lpAct04Curr(i) = "#" & DataVar.lpAct04Curr(i)
        
        DataVar.lpAct04Volt(i) = frmRun.pnlAct04Volt(i).Caption
        If frmRun.pnlAct04Volt(i).BackColor = vbRed Then DataVar.lpAct04Volt(i) = "#" & DataVar.lpAct04Volt(i)
        
        DataVar.lpAct04Time(i) = frmRun.pnlAct04Time(i).Caption
        If frmRun.pnlAct04Time(i).BackColor = vbRed Then DataVar.lpAct04Time(i) = "#" & DataVar.lpAct04Time(i)
    Next
    
    DataVar.lpAct04Stall(0) = frmRun.pnlAct04StallDelta.Caption
    If frmRun.pnlAct04StallDelta.BackColor = vbRed Then DataVar.lpAct04Stall(0) = "#" & DataVar.lpAct04Stall(0)
    
    ' SENSOR
    For i = 0 To 5
        DataVar.lpSensor(i) = frmRun.pnlSensorVolt(i).Caption
        If frmRun.pnlSensorVolt(i).BackColor = vbRed Then DataVar.lpSensor(i) = "#" & DataVar.lpSensor(i)
    Next
End Sub

Private Sub DisplayWarning(Optional lpStr As String = "")
    Dim i As Integer
    Dim lBkColor(9) As Long
    
    Select Case lpStr
        Case "NGCOUNT":
            lBkColor(0) = vbYellow
            lBkColor(1) = vbBlack
        Case Else:
            lBkColor(0) = vbRed
            lBkColor(1) = vbWhite
    End Select
    
    frmRun.picLog.BackColor = lBkColor(0)
    frmRun.picLog.Visible = IIf(lpStr = "", False, True)
    frmRun.picLog.ZOrder 0
    
    For i = 0 To 4
        frmRun.lblLog(i).Caption = ""
        frmRun.lblLog(i).ForeColor = lBkColor(1)
    Next
    
    Select Case lpStr
        Case "NGTABLE":
            lblLog(1).Caption = "ºÒ·®Ç°À»"
            lblLog(2).Caption = "NG Å×ÀÌºí·Î"
            lblLog(3).Caption = "ÀÌµ¿ÇØÁÖ¼¼¿ä."
        
        Case "STARTPART":
            lblLog(1).Caption = "CHECK"
            lblLog(2).Caption = "START PART"
            lblLog(3).Caption = "ERROR."
        
        Case "MODELTYPE":
            lblLog(1).Caption = "¸ðµ¨ Å¸ÀÔÀ»"
            lblLog(2).Caption = "È®ÀÎÇØÁÖ¼¼¿ä."
            lblLog(3).Caption = ""
        
        Case "MARKING":
            lblLog(1).Caption = ""
            lblLog(2).Caption = "MARKING NG"
            lblLog(3).Caption = ""
            
        Case "PRODUCT":
'            lblLog(1).Caption = "Á¦Ç°ÀÌ °¨ÁöµÇÁö"
'            lblLog(2).Caption = "¾Ê¾Ò½À´Ï´Ù."
'            lblLog(3).Caption = ""
            lblLog(1).Caption = "PRODUCT"
            lblLog(2).Caption = "CHECK"
            lblLog(3).Caption = "NG"
        
        Case "SCANNER":
            lblLog(1).Caption = "¹ÙÄÚµå µ¥ÀÌÅÍ°¡"
            lblLog(2).Caption = "¿Ã¹Ù¸£Áö ¾Ê½À´Ï´Ù."
            lblLog(3).Caption = ""
        
        Case "BARCODE":
            lblLog(1).Caption = ""
            lblLog(2).Caption = "USE BARCODE SCAN"
            lblLog(3).Caption = ""
            
        Case "SAFETY":
            lblLog(1).Caption = ""
            lblLog(2).Caption = "CHECK SAFETY"
            lblLog(3).Caption = ""
            
        Case "SIDEDOOR":
            lblLog(1).Caption = ""
            lblLog(2).Caption = "CHECK DOOR"
            lblLog(3).Caption = ""
            
        Case "NGCOUNT":
            lblLog(1).Caption = "ÇöÀç Àåºñ¿¡¼­ " + frmRun.pnlNGCount.Caption + "°³ÀÇ"
            lblLog(2).Caption = "Á¦Ç° ºÒ·®ÀÌ ¿¬¼ÓÀ¸·Î ¹ß»ý."
            lblLog(3).Caption = "È®ÀÎ (Á¤Áö¹öÆ°)"
        
    End Select
    
    If DOS(O_BUZZER) = False And lpStr <> "" Then Call DO_Control(O_BUZZER, True)
    
    Call SetTime(TM_BUZZER)
End Sub

Private Function StopRelease() As Boolean
    If DIS(I_STOP_SW) Or bGlobalStopSw Then
        StopRelease = True
    Else
        StopRelease = False
    End If
End Function

Private Sub OnMarkRelease()
    If LOCALTEST Or TABLETYPE = False Or SetupVar.bMarkingUse(0) = False Then Exit Sub
    
    Call OnLog("ON MARKING RELEASE...")
    Call DO_Control(O_MARKING1, True)
    Call SetTime(TM_MARKING)
    
    Do
        DoEvents
        
        If ElapseTime(TM_MARKING) > SetupVar.dMarkingTime(0) Then
            If DIS(I_MARKING1_ON) And DIS(I_MARKING1_OFF) = False Then
                Call DO_Control(O_MARKING1, False)
                
                Exit Do
            End If
        End If
        
        If ElapseTime(TM_MARKING) > 10# Then
            Call OnLog("[ERROR] MARKING OFF SENSOR FAIL (X" & I_MARKING1_OFF & ")")
            Call DO_Control(O_MARKING1, False)
            
            Exit Do
        End If
        
        If StopRelease Then
            Exit Do
        End If
    Loop
End Sub

Private Sub OnSolRelease()
    Dim MarkingOnOffFlag As Boolean
    
    If LOCALTEST Or TABLETYPE = False Then Exit Sub
    
    Call UseableButton(False)
    Call OnLog("ON SOL RELEASE...")
    
    Call DO_Control(O_BUZZER, False)
    Call DO_Control(O_MARKING1, False)
    Call DO_Control(O_VIB, False)
    
    Call DO_Control(O_WORK_ON, True)
    Call DO_Control(O_WORK_OFF, True)
    
'    Call DO_Control(O_LEAK01_STOP, True)
'    Call DO_Control(O_LEAK02_STOP, True)
    
    Call SetTime(TM_SOL)
    Do
        DoEvents
        
        If ElapseTime(TM_SOL) > 0.3 Then
            Call DO_Control(O_WORK_ON, False)
            Call DO_Control(O_WORK_OFF, False)
            
            Exit Do
        End If
    Loop
    
    Call SetTime(TM_SOL)
    Do
        DoEvents
        
        If (SetupVar.bMarkingUse(0) And DIS(I_MARKING1_OFF)) Or SetupVar.bMarkingUse(0) = False Then
            MarkingOnOffFlag = True
        End If
        
        If MarkingOnOffFlag Then
            Exit Do
        End If
        
        If ElapseTime(TM_SOL) > 10 Then
            If SetupVar.bMarkingUse(0) And Not DIS(I_MARKING1_OFF) Then
                Call OnLog("[ERROR] MARKING OFF SENSOR FAIL (X" & I_MARKING1_OFF & ")")
                
                MarkingOnOffFlag = False
            End If
            
            Exit Do
        End If
        
        If StopRelease Then
            Exit Do
        End If
    Loop
    
    If MarkingOnOffFlag Then
        Call DO_Control(O_WORK_ON, False)
        Call DO_Control(O_WORK_OFF, True)
    End If
    
    Call SetTime(TM_SOL)
    Do
        DoEvents
        
        If DIS(I_WORK_OFF) Then
            Exit Do
        End If
        
        If StopRelease And DIS(I_WORK_OFF) = False And DIS(I_WORK_ON) = False Then
            Exit Do
        End If
        
        If ElapseTime(TM_SOL) > 10 Then
            If DIS(I_WORK_OFF) = False Then
                Call OnLog("[ERROR] WORK OFF SENSOR FAIL (X" & I_WORK_OFF & ")")
            End If
            
            Exit Do
        End If
    Loop
    
    Call DO_Control(O_WORK_ON, False)
    Call DO_Control(O_WORK_OFF, False)
    Call UseableButton(True)
End Sub

