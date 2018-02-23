VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "1"
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15000
   ScaleWidth      =   19110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab HiddenTab 
      Height          =   12495
      Left            =   14700
      TabIndex        =   18
      Top             =   14280
      Visible         =   0   'False
      Width           =   19245
      _ExtentX        =   33946
      _ExtentY        =   22040
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      Tab             =   6
      TabsPerRow      =   9
      TabHeight       =   882
      TabCaption(0)   =   "HIDDEN 1"
      TabPicture(0)   =   "frmSetup.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraBox(35)"
      Tab(0).Control(1)=   "fraBox(47)"
      Tab(0).Control(2)=   "fraBox(3)"
      Tab(0).Control(3)=   "fraBox(34)"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "HIDDEN 2"
      TabPicture(1)   =   "frmSetup.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraBox(37)"
      Tab(1).Control(1)=   "fraBox(36)"
      Tab(1).Control(2)=   "fraBox(41)"
      Tab(1).Control(3)=   "fraBox(42)"
      Tab(1).Control(4)=   "fraBox(38)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "HIDDEN 3"
      TabPicture(2)   =   "frmSetup.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtPTCTime"
      Tab(2).Control(1)=   "chkPTC"
      Tab(2).Control(2)=   "txtPTCName"
      Tab(2).Control(3)=   "txtPTCCurrHi"
      Tab(2).Control(4)=   "txtPTCCurrLo"
      Tab(2).Control(5)=   "fraBox(2)"
      Tab(2).Control(6)=   "Frame7"
      Tab(2).Control(7)=   "fraBox(40)"
      Tab(2).Control(8)=   "fraBox(39)"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "HIDDEN 4"
      TabPicture(3)   =   "frmSetup.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pnlBarcodeFrame(0)"
      Tab(3).Control(1)=   "SSPanel1(7)"
      Tab(3).Control(2)=   "SSPanel1(6)"
      Tab(3).Control(3)=   "SSPanel1(5)"
      Tab(3).Control(4)=   "SSPanel1(3)"
      Tab(3).Control(5)=   "SSPanel1(2)"
      Tab(3).Control(6)=   "SSPanel1(1)"
      Tab(3).Control(7)=   "SSPanel1(0)"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frmSetup.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraBox(1)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "LEAK"
      TabPicture(5)   =   "frmSetup.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraBox(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "ACT"
      TabPicture(6)   =   "frmSetup.frx":00A8
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Label3(48)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label3(49)"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Label3(50)"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Label3(51)"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "pnlBox(20)"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "pnlBox(19)"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "pnlBox(18)"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "pnlBox(17)"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "txtActBoardNo(0)"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).Control(9)=   "txtActBoardNo(1)"
      Tab(6).Control(9).Enabled=   0   'False
      Tab(6).Control(10)=   "txtActBoardNo(2)"
      Tab(6).Control(10).Enabled=   0   'False
      Tab(6).Control(11)=   "txtActBoardNo(3)"
      Tab(6).Control(11).Enabled=   0   'False
      Tab(6).ControlCount=   12
      TabCaption(7)   =   "BLOWER"
      TabPicture(7)   =   "frmSetup.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label1(1)"
      Tab(7).Control(1)=   "fraBox(15)"
      Tab(7).Control(2)=   "txtLinSpeed(0)"
      Tab(7).Control(3)=   "txtLinSpeed(1)"
      Tab(7).Control(4)=   "txtLinSpeed(2)"
      Tab(7).Control(5)=   "txtLinSpeed(3)"
      Tab(7).Control(6)=   "txtLinSpeed(4)"
      Tab(7).Control(7)=   "fraBox(9)"
      Tab(7).Control(8)=   "txtLinSpeed(5)"
      Tab(7).Control(9)=   "txtLinSpeed(6)"
      Tab(7).Control(10)=   "txtLinSpeed(7)"
      Tab(7).Control(11)=   "txtLinSpeed(8)"
      Tab(7).Control(12)=   "fraBox(14)"
      Tab(7).Control(13)=   "fraBox(8)"
      Tab(7).ControlCount=   14
      TabCaption(8)   =   "LIN"
      TabPicture(8)   =   "frmSetup.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "chkCheckPoint"
      Tab(8).Control(1)=   "Frame(0)"
      Tab(8).Control(2)=   "fraBox(46)"
      Tab(8).Control(3)=   "fraBox(45)"
      Tab(8).Control(4)=   "fraBox(44)"
      Tab(8).Control(5)=   "fraBox(43)"
      Tab(8).Control(6)=   "Frame1(8)"
      Tab(8).Control(7)=   "Frame1(6)"
      Tab(8).Control(8)=   "Frame1(4)"
      Tab(8).Control(9)=   "Frame1(3)"
      Tab(8).Control(10)=   "fraBox(4)"
      Tab(8).Control(11)=   "Frame1(10)"
      Tab(8).Control(12)=   "Frame1(9)"
      Tab(8).Control(13)=   "Frame1(7)"
      Tab(8).Control(14)=   "Frame1(5)"
      Tab(8).Control(15)=   "Frame1(2)"
      Tab(8).ControlCount=   16
      Begin VB.Frame fraBox 
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4035
         Index           =   8
         Left            =   -74460
         TabIndex        =   885
         Top             =   600
         Width           =   18615
         Begin VB.TextBox txtBlowerName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   8
            Left            =   15240
            TabIndex        =   925
            Text            =   "8"
            Top             =   900
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   8
            Left            =   15240
            TabIndex        =   924
            Text            =   "6.00"
            Top             =   1560
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   8
            Left            =   15240
            TabIndex        =   923
            Text            =   "25.00"
            Top             =   2220
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   8
            Left            =   15240
            TabIndex        =   922
            Text            =   "3.0"
            Top             =   2880
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   7
            Left            =   13680
            TabIndex        =   921
            Text            =   "7"
            Top             =   900
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   7
            Left            =   13680
            TabIndex        =   920
            Text            =   "6.00"
            Top             =   1560
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   7
            Left            =   13680
            TabIndex        =   919
            Text            =   "25.00"
            Top             =   2220
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   7
            Left            =   13680
            TabIndex        =   918
            Text            =   "3.0"
            Top             =   2880
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   6
            Left            =   12120
            TabIndex        =   917
            Text            =   "6"
            Top             =   900
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   6
            Left            =   12120
            TabIndex        =   916
            Text            =   "6.00"
            Top             =   1560
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   6
            Left            =   12120
            TabIndex        =   915
            Text            =   "25.00"
            Top             =   2220
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   6
            Left            =   12120
            TabIndex        =   914
            Text            =   "3.0"
            Top             =   2880
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   10560
            TabIndex        =   913
            Text            =   "5"
            Top             =   900
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   10560
            TabIndex        =   912
            Text            =   "6.00"
            Top             =   1560
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   10560
            TabIndex        =   911
            Text            =   "25.00"
            Top             =   2220
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   10560
            TabIndex        =   910
            Text            =   "3.0"
            Top             =   2880
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   5880
            TabIndex        =   909
            Text            =   "1.5"
            Top             =   2880
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2760
            TabIndex        =   908
            Text            =   "0.5"
            Top             =   2880
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   4320
            TabIndex        =   907
            Text            =   "2.0"
            Top             =   2880
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   7440
            TabIndex        =   906
            Text            =   "1.5"
            Top             =   2880
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   9000
            TabIndex        =   905
            Text            =   "3.0"
            Top             =   2880
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   5880
            TabIndex        =   904
            Text            =   "7.00"
            Top             =   2220
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   5880
            TabIndex        =   903
            Text            =   "2.5"
            Top             =   1560
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2760
            TabIndex        =   902
            Text            =   "0.50"
            Top             =   2220
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2760
            TabIndex        =   901
            Text            =   "0.00"
            Top             =   1560
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   4320
            TabIndex        =   900
            Text            =   "5.00"
            Top             =   2220
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   4320
            TabIndex        =   899
            Text            =   "1.50"
            Top             =   1560
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   7440
            TabIndex        =   898
            Text            =   "12.00"
            Top             =   2220
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   7440
            TabIndex        =   897
            Text            =   "4.60"
            Top             =   1560
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   9000
            TabIndex        =   896
            Text            =   "25.00"
            Top             =   2220
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   9000
            TabIndex        =   895
            Text            =   "6.00"
            Top             =   1560
            Width           =   1395
         End
         Begin VB.TextBox txtRpmCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   16800
            TabIndex        =   894
            Text            =   "3500"
            Top             =   2220
            Width           =   1395
         End
         Begin VB.TextBox txtRpmCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   16800
            TabIndex        =   893
            Text            =   "0"
            Top             =   1560
            Width           =   1395
         End
         Begin VB.CheckBox chkBlower 
            Caption         =   "BLOWER"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   892
            Top             =   0
            Width           =   1395
         End
         Begin VB.TextBox txtRpmName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   16800
            TabIndex        =   891
            Text            =   "RPM"
            Top             =   900
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   4320
            TabIndex        =   890
            Text            =   "1"
            Top             =   900
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   5880
            TabIndex        =   889
            Text            =   "2"
            Top             =   900
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   7440
            TabIndex        =   888
            Text            =   "3"
            Top             =   900
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   9000
            TabIndex        =   887
            Text            =   "4"
            Top             =   900
            Width           =   1395
         End
         Begin VB.TextBox txtBlowerName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2760
            TabIndex        =   886
            Text            =   "0"
            Top             =   900
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MAX TIME (Sec)"
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   11
            Left            =   480
            TabIndex        =   929
            Top             =   2940
            Width           =   1875
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MIN CURR (A)"
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   9
            Left            =   720
            TabIndex        =   928
            Top             =   1620
            Width           =   1635
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MAX CIRR (A)"
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   10
            Left            =   720
            TabIndex        =   927
            Top             =   2280
            Width           =   1635
         End
         Begin VB.Label lblTemp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   20
            Left            =   1560
            TabIndex        =   926
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.Frame fraBox 
         Caption         =   "MOTOR TYPE"
         ForeColor       =   &H00808080&
         Height          =   3240
         Index           =   14
         Left            =   -74700
         TabIndex        =   732
         Top             =   5400
         Width           =   2325
         Begin VB.OptionButton optBlowerType 
            Caption         =   "FET"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   660
            TabIndex        =   735
            Top             =   1620
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optBlowerType 
            Caption         =   "RES"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   660
            TabIndex        =   734
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton optBlowerType 
            Caption         =   "LIN"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   660
            TabIndex        =   733
            Top             =   2520
            Width           =   975
         End
      End
      Begin VB.TextBox txtLinSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   8
         Left            =   -60540
         TabIndex        =   644
         Text            =   "0"
         Top             =   4500
         Width           =   1395
      End
      Begin VB.TextBox txtLinSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   7
         Left            =   -62040
         TabIndex        =   643
         Text            =   "0"
         Top             =   4500
         Width           =   1395
      End
      Begin VB.TextBox txtLinSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   6
         Left            =   -63540
         TabIndex        =   642
         Text            =   "0"
         Top             =   4500
         Width           =   1395
      End
      Begin VB.TextBox txtLinSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   5
         Left            =   -65040
         TabIndex        =   641
         Text            =   "0"
         Top             =   4500
         Width           =   1395
      End
      Begin VB.TextBox txtActBoardNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   3
         Left            =   1620
         TabIndex        =   639
         Text            =   "4"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtActBoardNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   2
         Left            =   1620
         TabIndex        =   637
         Text            =   "3"
         Top             =   2580
         Width           =   855
      End
      Begin VB.TextBox txtActBoardNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   1
         Left            =   1620
         TabIndex        =   635
         Text            =   "2"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtActBoardNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   0
         Left            =   1620
         TabIndex        =   633
         Text            =   "1"
         Top             =   1500
         Width           =   855
      End
      Begin VB.Frame fraBox 
         Height          =   3075
         Index           =   39
         Left            =   -66480
         TabIndex        =   597
         Top             =   6000
         Width           =   3795
         Begin VB.TextBox txtIonSubName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2280
            TabIndex        =   605
            Text            =   "ON"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtIonSubName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   960
            TabIndex        =   604
            Text            =   "OFF"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtIonName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Left            =   180
            TabIndex        =   603
            Text            =   "IONIZER"
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox chkIonUse 
            Caption         =   "IONIZER"
            Height          =   435
            Left            =   240
            TabIndex        =   602
            Top             =   -60
            Width           =   1275
         End
         Begin VB.TextBox txtIonHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   960
            TabIndex        =   601
            Text            =   "0.55"
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox txtIonHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2280
            TabIndex        =   600
            Text            =   "4.95"
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox txtIonLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2280
            TabIndex        =   599
            Text            =   "3.55"
            Top             =   1740
            Width           =   1215
         End
         Begin VB.TextBox txtIonLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   960
            TabIndex        =   598
            Text            =   "0"
            Top             =   1740
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MAX"
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   46
            Left            =   210
            TabIndex        =   607
            Top             =   2340
            Width           =   585
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MIN"
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   47
            Left            =   315
            TabIndex        =   606
            Top             =   1800
            Width           =   480
         End
      End
      Begin VB.Frame fraBox 
         Caption         =   "PARTS 1"
         Height          =   6675
         Index           =   34
         Left            =   -74460
         TabIndex        =   470
         Top             =   2820
         Width           =   9495
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   27
            Left            =   6240
            TabIndex        =   557
            Top             =   4980
            Width           =   555
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "27"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   27
            Left            =   4860
            TabIndex        =   556
            Top             =   4980
            Width           =   675
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   37
            Left            =   8400
            TabIndex        =   555
            Top             =   4980
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   36
            Left            =   8400
            TabIndex        =   554
            Top             =   4500
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   35
            Left            =   8400
            TabIndex        =   553
            Top             =   4020
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   34
            Left            =   8400
            TabIndex        =   552
            Top             =   3540
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   33
            Left            =   8400
            TabIndex        =   551
            Top             =   3060
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   32
            Left            =   8400
            TabIndex        =   550
            Top             =   2580
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   25
            Left            =   6240
            TabIndex        =   549
            Top             =   4020
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   24
            Left            =   6240
            TabIndex        =   548
            Top             =   3540
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   23
            Left            =   6240
            TabIndex        =   547
            Top             =   3060
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   22
            Left            =   6240
            TabIndex        =   546
            Top             =   2580
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   21
            Left            =   6240
            TabIndex        =   545
            Top             =   2100
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   20
            Left            =   6240
            TabIndex        =   544
            Top             =   1620
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   17
            Left            =   4080
            TabIndex        =   543
            Top             =   4980
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   16
            Left            =   4080
            TabIndex        =   542
            Top             =   4500
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   15
            Left            =   4080
            TabIndex        =   541
            Top             =   4020
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   14
            Left            =   4080
            TabIndex        =   540
            Top             =   3540
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   13
            Left            =   4080
            TabIndex        =   539
            Top             =   3060
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   12
            Left            =   4080
            TabIndex        =   538
            Top             =   2580
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   11
            Left            =   4080
            TabIndex        =   537
            Top             =   2100
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   10
            Left            =   4080
            TabIndex        =   536
            Top             =   1620
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   7
            Left            =   1920
            TabIndex        =   535
            Top             =   4980
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   6
            Left            =   1920
            TabIndex        =   534
            Top             =   4500
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   5
            Left            =   1920
            TabIndex        =   533
            Top             =   4020
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   4
            Left            =   1920
            TabIndex        =   532
            Top             =   3540
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   1920
            TabIndex        =   531
            Top             =   3060
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   1920
            TabIndex        =   530
            Top             =   2580
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   1920
            TabIndex        =   529
            Top             =   2100
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   1920
            TabIndex        =   528
            Top             =   1620
            Width           =   555
         End
         Begin VB.Frame fraPartsScript 
            Caption         =   " INFORMATION "
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   540
            TabIndex        =   519
            Top             =   360
            Width           =   8415
            Begin VB.Label lblTemp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   6360
               TabIndex        =   527
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblTemp 
               AutoSize        =   -1  'True
               Caption         =   "NAME ""#"" START PART"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   465
               Index           =   1
               Left            =   6900
               TabIndex        =   526
               Top             =   240
               Width           =   1140
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblTemp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "OFF"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   2340
               TabIndex        =   525
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblTemp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "O N"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   420
               TabIndex        =   524
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblTemp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00A0A0A0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   4440
               TabIndex        =   523
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblTemp 
               AutoSize        =   -1  'True
               Caption         =   "ON OK"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   5
               Left            =   960
               TabIndex        =   522
               Top             =   360
               Width           =   615
            End
            Begin VB.Label lblTemp 
               AutoSize        =   -1  'True
               Caption         =   "OFF OK"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   6
               Left            =   2880
               TabIndex        =   521
               Top             =   360
               Width           =   690
            End
            Begin VB.Label lblTemp 
               AutoSize        =   -1  'True
               Caption         =   "NOT USE"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   7
               Left            =   4980
               TabIndex        =   520
               Top             =   360
               Width           =   795
            End
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   10
            Left            =   2700
            TabIndex        =   518
            Top             =   1620
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "11"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   11
            Left            =   2700
            TabIndex        =   517
            Top             =   2100
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "12"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   12
            Left            =   2700
            TabIndex        =   516
            Top             =   2580
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "13"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   13
            Left            =   2700
            TabIndex        =   515
            Top             =   3060
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "14"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   14
            Left            =   2700
            TabIndex        =   514
            Top             =   3540
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   15
            Left            =   2700
            TabIndex        =   513
            Top             =   4020
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "16"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   16
            Left            =   2700
            TabIndex        =   512
            Top             =   4500
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "17"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   17
            Left            =   2700
            TabIndex        =   511
            Top             =   4980
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   540
            TabIndex        =   510
            Top             =   1620
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   540
            TabIndex        =   509
            Top             =   2100
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   540
            TabIndex        =   508
            Top             =   2580
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   540
            TabIndex        =   507
            Top             =   3060
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   4
            Left            =   540
            TabIndex        =   506
            Top             =   3540
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   5
            Left            =   540
            TabIndex        =   505
            Top             =   4020
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   6
            Left            =   540
            TabIndex        =   504
            Top             =   4500
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   7
            Left            =   540
            TabIndex        =   503
            Top             =   4980
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   20
            Left            =   4860
            TabIndex        =   502
            Top             =   1620
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "21"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   21
            Left            =   4860
            TabIndex        =   501
            Top             =   2100
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "22"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   22
            Left            =   4860
            TabIndex        =   500
            Top             =   2580
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "23"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   23
            Left            =   4860
            TabIndex        =   499
            Top             =   3060
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "24"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   24
            Left            =   4860
            TabIndex        =   498
            Top             =   3540
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "25"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   25
            Left            =   4860
            TabIndex        =   497
            Top             =   4020
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "32"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   32
            Left            =   7020
            TabIndex        =   496
            Top             =   2580
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "33"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   33
            Left            =   7020
            TabIndex        =   495
            Top             =   3060
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "34"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   34
            Left            =   7020
            TabIndex        =   494
            Top             =   3540
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "35"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   35
            Left            =   7020
            TabIndex        =   493
            Top             =   4020
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "36"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   36
            Left            =   7020
            TabIndex        =   492
            Top             =   4500
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   37
            Left            =   7020
            TabIndex        =   491
            Top             =   4980
            Width           =   675
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   30
            Left            =   8400
            TabIndex        =   490
            Top             =   1620
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   30
            Left            =   7020
            TabIndex        =   489
            Top             =   1620
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   31
            Left            =   8400
            TabIndex        =   488
            Top             =   2100
            Width           =   555
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   31
            Left            =   7020
            TabIndex        =   487
            Top             =   2100
            Width           =   675
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   39
            Left            =   8400
            TabIndex        =   486
            Top             =   5940
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   38
            Left            =   8400
            TabIndex        =   485
            Top             =   5460
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   29
            Left            =   6240
            TabIndex        =   484
            Top             =   5940
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   28
            Left            =   6240
            TabIndex        =   483
            Top             =   5460
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   19
            Left            =   4080
            TabIndex        =   482
            Top             =   5940
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   18
            Left            =   4080
            TabIndex        =   481
            Top             =   5460
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   9
            Left            =   1920
            TabIndex        =   480
            Top             =   5940
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   8
            Left            =   1920
            TabIndex        =   479
            Top             =   5460
            Width           =   555
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "39"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   39
            Left            =   7020
            TabIndex        =   478
            Top             =   5940
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "38"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   38
            Left            =   7020
            TabIndex        =   477
            Top             =   5460
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "29"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   29
            Left            =   4860
            TabIndex        =   476
            Top             =   5940
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "28"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   28
            Left            =   4860
            TabIndex        =   475
            Top             =   5460
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   9
            Left            =   540
            TabIndex        =   474
            Top             =   5940
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   8
            Left            =   540
            TabIndex        =   473
            Top             =   5460
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "19"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   19
            Left            =   2700
            TabIndex        =   472
            Top             =   5940
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "18"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   18
            Left            =   2700
            TabIndex        =   471
            Top             =   5460
            Width           =   675
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   10
            Left            =   3480
            TabIndex        =   596
            Top             =   1620
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   11
            Left            =   3480
            TabIndex        =   595
            Top             =   2100
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   12
            Left            =   3480
            TabIndex        =   594
            Top             =   2580
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   13
            Left            =   3480
            TabIndex        =   593
            Top             =   3060
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   14
            Left            =   3480
            TabIndex        =   592
            Top             =   3540
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   15
            Left            =   3480
            TabIndex        =   591
            Top             =   4020
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   16
            Left            =   3480
            TabIndex        =   590
            Top             =   4500
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   17
            Left            =   3480
            TabIndex        =   589
            Top             =   4980
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   588
            Top             =   1620
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   587
            Top             =   2100
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   586
            Top             =   2580
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   585
            Top             =   3060
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   1320
            TabIndex        =   584
            Top             =   3540
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   1320
            TabIndex        =   583
            Top             =   4020
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   1320
            TabIndex        =   582
            Top             =   4500
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   1320
            TabIndex        =   581
            Top             =   4980
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   20
            Left            =   5640
            TabIndex        =   580
            Top             =   1620
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   21
            Left            =   5640
            TabIndex        =   579
            Top             =   2100
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   22
            Left            =   5640
            TabIndex        =   578
            Top             =   2580
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   23
            Left            =   5640
            TabIndex        =   577
            Top             =   3060
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   24
            Left            =   5640
            TabIndex        =   576
            Top             =   3540
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   25
            Left            =   5640
            TabIndex        =   575
            Top             =   4020
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   27
            Left            =   5640
            TabIndex        =   574
            Top             =   4980
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   30
            Left            =   7800
            TabIndex        =   573
            Top             =   1620
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   31
            Left            =   7800
            TabIndex        =   572
            Top             =   2100
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   32
            Left            =   7800
            TabIndex        =   571
            Top             =   2580
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   33
            Left            =   7800
            TabIndex        =   570
            Top             =   3060
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   34
            Left            =   7800
            TabIndex        =   569
            Top             =   3540
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   35
            Left            =   7800
            TabIndex        =   568
            Top             =   4020
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   36
            Left            =   7800
            TabIndex        =   567
            Top             =   4500
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   37
            Left            =   7800
            TabIndex        =   566
            Top             =   4980
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   39
            Left            =   7800
            TabIndex        =   565
            Top             =   5940
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   38
            Left            =   7800
            TabIndex        =   564
            Top             =   5460
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   29
            Left            =   5640
            TabIndex        =   563
            Top             =   5940
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   28
            Left            =   5640
            TabIndex        =   562
            Top             =   5460
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   9
            Left            =   1320
            TabIndex        =   561
            Top             =   5940
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   8
            Left            =   1320
            TabIndex        =   560
            Top             =   5460
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   19
            Left            =   3480
            TabIndex        =   559
            Top             =   5940
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   18
            Left            =   3480
            TabIndex        =   558
            Top             =   5460
            Width           =   555
         End
      End
      Begin VB.Frame fraBox 
         Caption         =   "PRODUCT CHECK"
         Height          =   2655
         Index           =   3
         Left            =   -66060
         TabIndex        =   466
         Top             =   540
         Width           =   4875
         Begin VB.TextBox txtProductName 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   435
            Left            =   480
            TabIndex        =   469
            Top             =   1320
            Width           =   3855
         End
         Begin VB.TextBox txtProductList 
            Height          =   435
            Left            =   60
            TabIndex        =   468
            Text            =   "54"
            Top             =   2160
            Visible         =   0   'False
            Width           =   4755
         End
         Begin VB.CheckBox chkProductUse 
            Caption         =   "SIDE SENSOR"
            Height          =   495
            Left            =   540
            TabIndex        =   467
            Top             =   480
            Width           =   3795
         End
      End
      Begin VB.Frame fraBox 
         Height          =   1935
         Index           =   40
         Left            =   -67200
         TabIndex        =   462
         Top             =   3600
         Width           =   4875
         Begin VB.TextBox txtMarkingTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   0
            Left            =   1200
            TabIndex        =   464
            Text            =   "1.0"
            Top             =   1020
            Width           =   1935
         End
         Begin VB.CheckBox chkMarking 
            Caption         =   "USE MARKING 1"
            Height          =   435
            Index           =   0
            Left            =   1200
            TabIndex        =   463
            Top             =   540
            Width           =   2235
         End
         Begin VB.Label lblMarkingdd 
            AutoSize        =   -1  'True
            Caption         =   "Sec"
            Height          =   285
            Index           =   0
            Left            =   3300
            TabIndex        =   465
            Top             =   1140
            Width           =   390
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2535
         Left            =   -69360
         TabIndex        =   458
         Top             =   540
         Width           =   4875
         Begin VB.TextBox txtScannerValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   2760
            TabIndex        =   460
            Text            =   "1"
            Top             =   1080
            Width           =   735
         End
         Begin VB.CheckBox chkScannerUse 
            Caption         =   "USE SCANNER"
            Height          =   435
            Left            =   240
            TabIndex        =   459
            Top             =   -60
            Width           =   1995
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "NG COUNT"
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   461
            Top             =   1140
            Width           =   1290
         End
      End
      Begin VB.Frame fraBox 
         Height          =   3795
         Index           =   2
         Left            =   -74640
         TabIndex        =   447
         Top             =   1560
         Width           =   4515
         Begin VB.CheckBox chkBarCodePrintUse 
            Caption         =   "USE BARCODE PRINT"
            Height          =   315
            Left            =   240
            TabIndex        =   451
            Top             =   0
            Width           =   2715
         End
         Begin VB.OptionButton optBarcodeType 
            Caption         =   "2D"
            Height          =   435
            Index           =   0
            Left            =   780
            TabIndex        =   450
            Top             =   540
            Width           =   795
         End
         Begin VB.OptionButton optBarcodeType 
            Caption         =   "QR"
            Height          =   435
            Index           =   1
            Left            =   1920
            TabIndex        =   449
            Top             =   540
            Width           =   795
         End
         Begin VB.TextBox txtBarCode 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   11.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   2280
            TabIndex        =   448
            Text            =   "0000000"
            Top             =   3180
            Width           =   1875
         End
         Begin Threed.SSPanel pnlBarcodeFrame 
            Height          =   1875
            Index           =   1
            Left            =   360
            TabIndex        =   452
            Top             =   1080
            Width           =   3795
            _Version        =   65536
            _ExtentX        =   6694
            _ExtentY        =   3307
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
            Begin VB.TextBox txtBarCode 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   2
               Left            =   1620
               TabIndex        =   454
               Text            =   "MODEL"
               Top             =   300
               Width           =   2055
            End
            Begin VB.TextBox txtBarCode 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   3
               Left            =   1620
               TabIndex        =   453
               Text            =   "PART NO"
               Top             =   720
               Width           =   2055
            End
            Begin Threed.SSPanel pnlBarcode 
               Height          =   1215
               Left            =   120
               TabIndex        =   455
               Top             =   300
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   2143
               _StockProps     =   15
               Caption         =   "BARCODE"
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
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel1 
               Height          =   375
               Index           =   10
               Left            =   1620
               TabIndex        =   456
               Top             =   1140
               Width           =   2055
               _Version        =   65536
               _ExtentX        =   3625
               _ExtentY        =   661
               _StockProps     =   15
               Caption         =   "YYYYMMDD-SERIAL"
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
               Alignment       =   1
            End
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "SUPPLIER CODE"
            Height          =   285
            Index           =   2
            Left            =   360
            TabIndex        =   457
            Top             =   3180
            Width           =   1785
         End
      End
      Begin VB.Frame fraBox 
         Height          =   1935
         Index           =   47
         Left            =   -63540
         TabIndex        =   443
         Top             =   960
         Width           =   4515
         Begin VB.CheckBox chkPart 
            Caption         =   "CHECK DOOR (X26)"
            Height          =   345
            Index           =   26
            Left            =   960
            TabIndex        =   445
            Top             =   900
            Width           =   2775
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   26
            Left            =   720
            TabIndex        =   444
            Top             =   240
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   26
            Left            =   120
            TabIndex        =   446
            Top             =   240
            Visible         =   0   'False
            Width           =   555
         End
      End
      Begin VB.Frame fraBox 
         Height          =   1755
         Index           =   38
         Left            =   -72240
         TabIndex        =   440
         Top             =   1440
         Width           =   4515
         Begin VB.TextBox txtVisionName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   0
            Left            =   780
            TabIndex        =   442
            Text            =   "VISION"
            Top             =   780
            Width           =   2835
         End
         Begin VB.CheckBox chkVision 
            Caption         =   "USE VISION"
            Height          =   315
            Left            =   240
            TabIndex        =   441
            Top             =   0
            Width           =   1635
         End
      End
      Begin VB.Frame fraBox 
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2760
         Index           =   42
         Left            =   -74280
         TabIndex        =   423
         Top             =   5700
         Width           =   6675
         Begin VB.TextBox txtVisionDoorName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   2
            Left            =   1140
            TabIndex        =   425
            Text            =   "DOOR 1"
            Top             =   540
            Width           =   2835
         End
         Begin VB.CheckBox chkVisionDoor 
            Caption         =   "DOOR CHECK WITH ACT 03"
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   424
            Top             =   -120
            Width           =   3495
         End
         Begin Threed.SSPanel pnlVisionPoint 
            Height          =   675
            Index           =   4
            Left            =   180
            TabIndex        =   426
            Top             =   1080
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   1191
            _StockProps     =   15
            Caption         =   "  OPEN"
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
            RoundedCorners  =   0   'False
            Alignment       =   1
            Begin VB.OptionButton optOpenCameraPos3 
               Caption         =   "P1"
               Height          =   435
               Index           =   0
               Left            =   2460
               Style           =   1  'Graphical
               TabIndex        =   429
               Top             =   120
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.OptionButton optOpenCameraPos3 
               Caption         =   "P2"
               Height          =   435
               Index           =   1
               Left            =   3180
               Style           =   1  'Graphical
               TabIndex        =   428
               Top             =   120
               Width           =   675
            End
            Begin VB.ComboBox cboOpenCameraNo3 
               Height          =   405
               ItemData        =   "frmSetup.frx":00FC
               Left            =   960
               List            =   "frmSetup.frx":013C
               TabIndex        =   427
               Text            =   "OPEN"
               Top             =   120
               Width           =   1395
            End
         End
         Begin Threed.SSPanel pnlVisionPoint 
            Height          =   675
            Index           =   5
            Left            =   180
            TabIndex        =   430
            Top             =   1860
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   1191
            _StockProps     =   15
            Caption         =   "  CLOSE"
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
            RoundedCorners  =   0   'False
            Alignment       =   1
            Begin VB.OptionButton optCloseCameraPos3 
               Caption         =   "P2"
               Height          =   435
               Index           =   1
               Left            =   3180
               Style           =   1  'Graphical
               TabIndex        =   433
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optCloseCameraPos3 
               Caption         =   "P1"
               Height          =   435
               Index           =   0
               Left            =   2460
               Style           =   1  'Graphical
               TabIndex        =   432
               Top             =   120
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.ComboBox cboCloseCameraNo3 
               Height          =   405
               ItemData        =   "frmSetup.frx":01A4
               Left            =   960
               List            =   "frmSetup.frx":01E4
               TabIndex        =   431
               Text            =   "CLOSE"
               Top             =   120
               Width           =   1395
            End
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            Height          =   285
            Left            =   300
            TabIndex        =   434
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame fraBox 
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2760
         Index           =   41
         Left            =   -67020
         TabIndex        =   411
         Top             =   2280
         Width           =   6675
         Begin VB.CheckBox chkVisionDoor 
            Caption         =   "DOOR CHECK WITH ACT 02"
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   413
            Top             =   -120
            Width           =   3495
         End
         Begin VB.TextBox txtVisionDoorName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   1
            Left            =   1140
            TabIndex        =   412
            Text            =   "DOOR 1"
            Top             =   540
            Width           =   2835
         End
         Begin Threed.SSPanel pnlVisionPoint 
            Height          =   675
            Index           =   2
            Left            =   180
            TabIndex        =   414
            Top             =   1080
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   1191
            _StockProps     =   15
            Caption         =   "  OPEN"
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
            RoundedCorners  =   0   'False
            Alignment       =   1
            Begin VB.OptionButton optOpenCameraPos2 
               Caption         =   "P2"
               Height          =   435
               Index           =   1
               Left            =   3240
               Style           =   1  'Graphical
               TabIndex        =   417
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optOpenCameraPos2 
               Caption         =   "P1"
               Height          =   435
               Index           =   0
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   416
               Top             =   120
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.ComboBox cboOpenCameraNo2 
               Height          =   405
               ItemData        =   "frmSetup.frx":024C
               Left            =   960
               List            =   "frmSetup.frx":028C
               TabIndex        =   415
               Text            =   "OPEN"
               Top             =   120
               Width           =   1395
            End
         End
         Begin Threed.SSPanel pnlVisionPoint 
            Height          =   675
            Index           =   3
            Left            =   180
            TabIndex        =   418
            Top             =   1860
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   1191
            _StockProps     =   15
            Caption         =   "  CLOSE"
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
            RoundedCorners  =   0   'False
            Alignment       =   1
            Begin VB.OptionButton optCloseCameraPos2 
               Caption         =   "P1"
               Height          =   435
               Index           =   0
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   421
               Top             =   120
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.OptionButton optCloseCameraPos2 
               Caption         =   "P2"
               Height          =   435
               Index           =   1
               Left            =   3240
               Style           =   1  'Graphical
               TabIndex        =   420
               Top             =   120
               Width           =   675
            End
            Begin VB.ComboBox cboCloseCameraNo2 
               Height          =   405
               ItemData        =   "frmSetup.frx":02F4
               Left            =   960
               List            =   "frmSetup.frx":0334
               TabIndex        =   419
               Text            =   "CLOSE"
               Top             =   120
               Width           =   1395
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            Height          =   285
            Index           =   0
            Left            =   300
            TabIndex        =   422
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame fraBox 
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2760
         Index           =   36
         Left            =   -74280
         TabIndex        =   393
         Top             =   2280
         Width           =   6675
         Begin VB.CheckBox chkVisionDoor 
            Caption         =   "DOOR CHECK WITH ACT 01"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   395
            Top             =   -120
            Width           =   3495
         End
         Begin VB.TextBox txtVisionDoorName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   0
            Left            =   1140
            TabIndex        =   394
            Text            =   "DOOR 1"
            Top             =   540
            Width           =   2835
         End
         Begin Threed.SSPanel pnlVisionPoint 
            Height          =   675
            Index           =   0
            Left            =   180
            TabIndex        =   396
            Top             =   1080
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   1191
            _StockProps     =   15
            Caption         =   "  OPEN"
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
            RoundedCorners  =   0   'False
            Alignment       =   1
            Begin VB.OptionButton optOpenCameraPos1 
               Caption         =   "P5"
               Height          =   435
               Index           =   4
               Left            =   5400
               Style           =   1  'Graphical
               TabIndex        =   402
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optOpenCameraPos1 
               Caption         =   "P4"
               Height          =   435
               Index           =   3
               Left            =   4680
               Style           =   1  'Graphical
               TabIndex        =   401
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optOpenCameraPos1 
               Caption         =   "P3"
               Height          =   435
               Index           =   2
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   400
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optOpenCameraPos1 
               Caption         =   "P1"
               Height          =   435
               Index           =   0
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   399
               Top             =   120
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.OptionButton optOpenCameraPos1 
               Caption         =   "P2"
               Height          =   435
               Index           =   1
               Left            =   3240
               Style           =   1  'Graphical
               TabIndex        =   398
               Top             =   120
               Width           =   675
            End
            Begin VB.ComboBox cboOpenCameraNo1 
               Height          =   405
               ItemData        =   "frmSetup.frx":039C
               Left            =   960
               List            =   "frmSetup.frx":03DC
               TabIndex        =   397
               Text            =   "OPEN"
               Top             =   120
               Width           =   1395
            End
         End
         Begin Threed.SSPanel pnlVisionPoint 
            Height          =   675
            Index           =   1
            Left            =   180
            TabIndex        =   403
            Top             =   1860
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   1191
            _StockProps     =   15
            Caption         =   "  CLOSE"
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
            RoundedCorners  =   0   'False
            Alignment       =   1
            Begin VB.OptionButton optCloseCameraPos1 
               Caption         =   "P5"
               Height          =   435
               Index           =   4
               Left            =   5400
               Style           =   1  'Graphical
               TabIndex        =   409
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optCloseCameraPos1 
               Caption         =   "P4"
               Height          =   435
               Index           =   3
               Left            =   4680
               Style           =   1  'Graphical
               TabIndex        =   408
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optCloseCameraPos1 
               Caption         =   "P3"
               Height          =   435
               Index           =   2
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   407
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optCloseCameraPos1 
               Caption         =   "P2"
               Height          =   435
               Index           =   1
               Left            =   3240
               Style           =   1  'Graphical
               TabIndex        =   406
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optCloseCameraPos1 
               Caption         =   "P1"
               Height          =   435
               Index           =   0
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   405
               Top             =   120
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.ComboBox cboCloseCameraNo1 
               Height          =   405
               ItemData        =   "frmSetup.frx":0444
               Left            =   960
               List            =   "frmSetup.frx":0487
               TabIndex        =   404
               Text            =   "CLOSE"
               Top             =   120
               Width           =   1395
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            Height          =   285
            Index           =   0
            Left            =   300
            TabIndex        =   410
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame fraBox 
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2760
         Index           =   37
         Left            =   -66960
         TabIndex        =   381
         Top             =   5700
         Width           =   6495
         Begin VB.CheckBox chkVisionDoor 
            Caption         =   "DOOR CHECK WITH ACT 04"
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   383
            Top             =   -120
            Width           =   3495
         End
         Begin VB.TextBox txtVisionDoorName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   3
            Left            =   1140
            TabIndex        =   382
            Text            =   "DOOR 1"
            Top             =   540
            Width           =   2835
         End
         Begin Threed.SSPanel pnlVisionPoint 
            Height          =   675
            Index           =   6
            Left            =   180
            TabIndex        =   384
            Top             =   1080
            Width           =   6135
            _Version        =   65536
            _ExtentX        =   10821
            _ExtentY        =   1191
            _StockProps     =   15
            Caption         =   "  OPEN"
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
            RoundedCorners  =   0   'False
            Alignment       =   1
            Begin VB.ComboBox cboOpenCameraNo4 
               Height          =   405
               ItemData        =   "frmSetup.frx":04EF
               Left            =   960
               List            =   "frmSetup.frx":0517
               TabIndex        =   387
               Text            =   "OPEN"
               Top             =   120
               Width           =   1395
            End
            Begin VB.OptionButton optOpenCameraPos4 
               Caption         =   "P2"
               Height          =   435
               Index           =   1
               Left            =   3180
               Style           =   1  'Graphical
               TabIndex        =   386
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optOpenCameraPos4 
               Caption         =   "P1"
               Height          =   435
               Index           =   0
               Left            =   2460
               Style           =   1  'Graphical
               TabIndex        =   385
               Top             =   120
               Value           =   -1  'True
               Width           =   675
            End
         End
         Begin Threed.SSPanel pnlVisionPoint 
            Height          =   675
            Index           =   7
            Left            =   180
            TabIndex        =   388
            Top             =   1860
            Width           =   6135
            _Version        =   65536
            _ExtentX        =   10821
            _ExtentY        =   1191
            _StockProps     =   15
            Caption         =   "  CLOSE"
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
            RoundedCorners  =   0   'False
            Alignment       =   1
            Begin VB.OptionButton optCloseCameraPos4 
               Caption         =   "P1"
               Height          =   435
               Index           =   0
               Left            =   2460
               Style           =   1  'Graphical
               TabIndex        =   391
               Top             =   120
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.OptionButton optCloseCameraPos4 
               Caption         =   "P2"
               Height          =   435
               Index           =   1
               Left            =   3180
               Style           =   1  'Graphical
               TabIndex        =   390
               Top             =   120
               Width           =   675
            End
            Begin VB.ComboBox cboCloseCameraNo4 
               Height          =   405
               ItemData        =   "frmSetup.frx":0557
               Left            =   960
               List            =   "frmSetup.frx":057F
               TabIndex        =   389
               Text            =   "CLOSE"
               Top             =   120
               Width           =   1395
            End
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            Height          =   285
            Left            =   300
            TabIndex        =   392
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame fraBox 
         Height          =   4155
         Index           =   9
         Left            =   -74160
         TabIndex        =   359
         Top             =   5700
         Width           =   14475
         Begin VB.Frame fraBox 
            Caption         =   "DATA PROCESS"
            Height          =   2655
            Index           =   16
            Left            =   3660
            TabIndex        =   371
            Top             =   780
            Width           =   3555
            Begin VB.TextBox txtVibEnd 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   2580
               TabIndex        =   375
               Text            =   "20"
               Top             =   1860
               Width           =   735
            End
            Begin VB.TextBox txtVibStart 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   1680
               TabIndex        =   374
               Text            =   "10"
               Top             =   1860
               Width           =   735
            End
            Begin VB.OptionButton optVibResultType 
               Caption         =   "Peak"
               Height          =   285
               Index           =   0
               Left            =   540
               TabIndex        =   373
               Top             =   660
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.OptionButton optVibResultType 
               Caption         =   "RMS"
               Height          =   285
               Index           =   1
               Left            =   540
               TabIndex        =   372
               Top             =   1440
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "POSITION (%)"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   2
               Left            =   1680
               TabIndex        =   376
               Top             =   1500
               Width           =   1275
            End
         End
         Begin VB.CheckBox chkVib 
            Caption         =   "VIBRATION"
            Height          =   285
            Left            =   240
            TabIndex        =   370
            Top             =   0
            Width           =   1635
         End
         Begin VB.TextBox txtVibName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   1680
            TabIndex        =   369
            Text            =   "VIB"
            Top             =   900
            Width           =   1215
         End
         Begin VB.TextBox txtVibCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   1680
            TabIndex        =   368
            Text            =   "0.00"
            Top             =   1500
            Width           =   1215
         End
         Begin VB.TextBox txtVibCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   1680
            TabIndex        =   367
            Text            =   "0.50"
            Top             =   2100
            Width           =   1215
         End
         Begin VB.Frame fraBox 
            Caption         =   "TEST CONDITION"
            Height          =   2655
            Index           =   17
            Left            =   7980
            TabIndex        =   360
            Top             =   780
            Width           =   5715
            Begin VB.OptionButton optVibMethod 
               Caption         =   "Blower HI"
               Height          =   315
               Index           =   0
               Left            =   540
               TabIndex        =   364
               Top             =   660
               Value           =   -1  'True
               Width           =   1755
            End
            Begin VB.OptionButton optVibMethod 
               Caption         =   "User Define"
               Height          =   315
               Index           =   1
               Left            =   540
               TabIndex        =   363
               Top             =   1440
               Width           =   1755
            End
            Begin VB.TextBox txtVibVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   405
               Left            =   2460
               TabIndex        =   362
               Text            =   "12.0"
               Top             =   1860
               Width           =   1215
            End
            Begin VB.TextBox txtVibTime 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   3900
               TabIndex        =   361
               Text            =   "0.50"
               Top             =   1860
               Width           =   1215
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "TEST VOLT"
               Height          =   285
               Index           =   2
               Left            =   2460
               TabIndex        =   366
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "TEST TIME"
               Height          =   285
               Index           =   1
               Left            =   3900
               TabIndex        =   365
               Top             =   1440
               Width           =   1170
            End
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   5
            Left            =   795
            TabIndex        =   379
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MIN"
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   45
            Left            =   1080
            TabIndex        =   378
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MAX"
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   44
            Left            =   975
            TabIndex        =   377
            Top             =   2160
            Width           =   585
         End
      End
      Begin VB.TextBox txtLinSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   4
         Left            =   -66540
         TabIndex        =   357
         Text            =   "0"
         Top             =   4500
         Width           =   1395
      End
      Begin VB.TextBox txtLinSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   3
         Left            =   -68160
         TabIndex        =   356
         Text            =   "0"
         Top             =   4500
         Width           =   1395
      End
      Begin VB.TextBox txtLinSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   2
         Left            =   -69780
         TabIndex        =   355
         Text            =   "0"
         Top             =   4500
         Width           =   1395
      End
      Begin VB.TextBox txtLinSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   1
         Left            =   -71400
         TabIndex        =   354
         Text            =   "0"
         Top             =   4500
         Width           =   1395
      End
      Begin VB.TextBox txtLinSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   0
         Left            =   -73020
         TabIndex        =   353
         Text            =   "0"
         Top             =   4500
         Width           =   1395
      End
      Begin VB.Frame fraBox 
         Caption         =   "DIRECTION"
         ForeColor       =   &H00808080&
         Height          =   2460
         Index           =   15
         Left            =   -69360
         TabIndex        =   350
         Top             =   5100
         Width           =   2325
         Begin VB.OptionButton optBlowerDirection 
            Caption         =   "CW"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   660
            TabIndex        =   352
            Top             =   720
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optBlowerDirection 
            Caption         =   "CCW"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   660
            TabIndex        =   351
            Top             =   1620
            Width           =   975
         End
      End
      Begin VB.TextBox txtPTCCurrLo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   -69540
         TabIndex        =   349
         Text            =   "0.5"
         Top             =   4020
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtPTCCurrHi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   -69540
         TabIndex        =   348
         Text            =   "3"
         Top             =   4500
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtPTCName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   -69540
         TabIndex        =   347
         Text            =   "PTC"
         Top             =   3540
         Width           =   1695
      End
      Begin VB.CheckBox chkPTC 
         Caption         =   "PTC"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69540
         TabIndex        =   346
         Top             =   3060
         Width           =   1695
      End
      Begin VB.TextBox txtPTCTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   -69540
         TabIndex        =   345
         Text            =   "3"
         Top             =   4980
         Width           =   1695
      End
      Begin VB.CheckBox chkCheckPoint 
         Caption         =   "USE CHECK POINT"
         Height          =   315
         Left            =   -67200
         TabIndex        =   344
         Top             =   9420
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Frame Frame 
         Caption         =   "MODE REF POS"
         Height          =   1215
         Index           =   0
         Left            =   -68340
         TabIndex        =   341
         Top             =   7740
         Width           =   4395
         Begin VB.OptionButton optAct01RefPos 
            Caption         =   "STALL"
            Height          =   495
            Index           =   1
            Left            =   2460
            TabIndex        =   343
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optAct01RefPos 
            Caption         =   "LIMIT S/W"
            Height          =   495
            Index           =   0
            Left            =   600
            TabIndex        =   342
            Top             =   480
            Value           =   -1  'True
            Width           =   1512
         End
      End
      Begin VB.Frame fraBox 
         Height          =   4875
         Index           =   46
         Left            =   -67680
         TabIndex        =   323
         Top             =   2160
         Width           =   4395
         Begin VB.TextBox txtLinActAngle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2460
            TabIndex        =   332
            Text            =   "0"
            Top             =   2580
            Width           =   1635
         End
         Begin VB.CheckBox chkLinAct 
            Caption         =   "ACT 04 DISP."
            Height          =   315
            Index           =   3
            Left            =   300
            TabIndex        =   331
            Top             =   0
            Width           =   1815
         End
         Begin VB.TextBox txtLinActLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2460
            TabIndex        =   330
            Text            =   "0"
            Top             =   960
            Width           =   1635
         End
         Begin VB.TextBox txtLinActHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2460
            TabIndex        =   329
            Text            =   "999"
            Top             =   1500
            Width           =   1635
         End
         Begin VB.TextBox txtLinActFinal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2460
            TabIndex        =   328
            Text            =   "100"
            Top             =   2040
            Width           =   1635
         End
         Begin VB.TextBox txtLinActName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2460
            TabIndex        =   327
            Text            =   "ACT 04"
            Top             =   420
            Width           =   1635
         End
         Begin VB.TextBox txtLinActTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2460
            TabIndex        =   326
            Text            =   "20"
            Top             =   3120
            Width           =   1635
         End
         Begin VB.TextBox txtLinActCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2460
            TabIndex        =   325
            Text            =   "0"
            Top             =   3660
            Width           =   1635
         End
         Begin VB.TextBox txtLinActCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2460
            TabIndex        =   324
            Text            =   "999"
            Top             =   4200
            Width           =   1635
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DOOR STEP"
            Height          =   285
            Index           =   52
            Left            =   915
            TabIndex        =   340
            Top             =   2640
            Width           =   1320
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MIN ANGLE"
            Height          =   315
            Index           =   25
            Left            =   840
            TabIndex        =   339
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MAX ANGLE"
            Height          =   315
            Index           =   24
            Left            =   780
            TabIndex        =   338
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FINAL ANGLE"
            Height          =   276
            Index           =   23
            Left            =   684
            TabIndex        =   337
            Top             =   2100
            Width           =   1548
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            Height          =   315
            Index           =   21
            Left            =   1500
            TabIndex        =   336
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TEST TIME"
            Height          =   315
            Index           =   56
            Left            =   1020
            TabIndex        =   335
            Top             =   3180
            Width           =   1215
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CURR MIN"
            Height          =   285
            Index           =   63
            Left            =   1065
            TabIndex        =   334
            Top             =   3720
            Width           =   1185
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CURR MAX"
            Height          =   285
            Index           =   64
            Left            =   960
            TabIndex        =   333
            Top             =   4260
            Width           =   1290
         End
      End
      Begin VB.Frame fraBox 
         Height          =   4875
         Index           =   45
         Left            =   -67680
         TabIndex        =   305
         Top             =   1740
         Width           =   4395
         Begin VB.TextBox txtLinActAngle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   314
            Text            =   "0"
            Top             =   2580
            Width           =   1635
         End
         Begin VB.CheckBox chkLinAct 
            Caption         =   "ACT 03 DISP."
            Height          =   315
            Index           =   2
            Left            =   300
            TabIndex        =   313
            Top             =   0
            Width           =   1815
         End
         Begin VB.TextBox txtLinActLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   312
            Text            =   "0"
            Top             =   960
            Width           =   1635
         End
         Begin VB.TextBox txtLinActHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   311
            Text            =   "999"
            Top             =   1500
            Width           =   1635
         End
         Begin VB.TextBox txtLinActFinal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   310
            Text            =   "100"
            Top             =   2040
            Width           =   1635
         End
         Begin VB.TextBox txtLinActName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   309
            Text            =   "ACT 03"
            Top             =   420
            Width           =   1635
         End
         Begin VB.TextBox txtLinActTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   308
            Text            =   "20"
            Top             =   3120
            Width           =   1635
         End
         Begin VB.TextBox txtLinActCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   307
            Text            =   "0"
            Top             =   3660
            Width           =   1635
         End
         Begin VB.TextBox txtLinActCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   306
            Text            =   "999"
            Top             =   4200
            Width           =   1635
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DOOR STEP"
            Height          =   285
            Index           =   51
            Left            =   915
            TabIndex        =   322
            Top             =   2640
            Width           =   1320
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MIN ANGLE"
            Height          =   315
            Index           =   20
            Left            =   840
            TabIndex        =   321
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MAX ANGLE"
            Height          =   315
            Index           =   19
            Left            =   780
            TabIndex        =   320
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FINAL ANGLE"
            Height          =   276
            Index           =   18
            Left            =   684
            TabIndex        =   319
            Top             =   2100
            Width           =   1548
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            Height          =   315
            Index           =   16
            Left            =   1500
            TabIndex        =   318
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TEST TIME"
            Height          =   315
            Index           =   55
            Left            =   1020
            TabIndex        =   317
            Top             =   3180
            Width           =   1215
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CURR MIN"
            Height          =   285
            Index           =   61
            Left            =   1065
            TabIndex        =   316
            Top             =   3720
            Width           =   1185
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CURR MAX"
            Height          =   285
            Index           =   62
            Left            =   960
            TabIndex        =   315
            Top             =   4260
            Width           =   1290
         End
      End
      Begin VB.Frame fraBox 
         Height          =   4875
         Index           =   44
         Left            =   -67680
         TabIndex        =   287
         Top             =   1320
         Width           =   4395
         Begin VB.TextBox txtLinActAngle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2460
            TabIndex        =   296
            Text            =   "0"
            Top             =   2580
            Width           =   1635
         End
         Begin VB.CheckBox chkLinAct 
            Caption         =   "ACT 02 DISP."
            Height          =   315
            Index           =   1
            Left            =   300
            TabIndex        =   295
            Top             =   0
            Width           =   1815
         End
         Begin VB.TextBox txtLinActLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2460
            TabIndex        =   294
            Text            =   "0"
            Top             =   960
            Width           =   1635
         End
         Begin VB.TextBox txtLinActHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2460
            TabIndex        =   293
            Text            =   "999"
            Top             =   1500
            Width           =   1635
         End
         Begin VB.TextBox txtLinActFinal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2460
            TabIndex        =   292
            Text            =   "100"
            Top             =   2040
            Width           =   1635
         End
         Begin VB.TextBox txtLinActName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2460
            TabIndex        =   291
            Text            =   "ACT 02"
            Top             =   420
            Width           =   1635
         End
         Begin VB.TextBox txtLinActTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2460
            TabIndex        =   290
            Text            =   "20"
            Top             =   3120
            Width           =   1635
         End
         Begin VB.TextBox txtLinActCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2460
            TabIndex        =   289
            Text            =   "0"
            Top             =   3660
            Width           =   1635
         End
         Begin VB.TextBox txtLinActCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2460
            TabIndex        =   288
            Text            =   "999"
            Top             =   4200
            Width           =   1635
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DOOR STEP"
            Height          =   285
            Index           =   50
            Left            =   915
            TabIndex        =   304
            Top             =   2640
            Width           =   1320
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MIN ANGLE"
            Height          =   315
            Index           =   15
            Left            =   840
            TabIndex        =   303
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MAX ANGLE"
            Height          =   315
            Index           =   14
            Left            =   780
            TabIndex        =   302
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FINAL ANGLE"
            Height          =   276
            Index           =   13
            Left            =   684
            TabIndex        =   301
            Top             =   2100
            Width           =   1548
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            Height          =   315
            Index           =   11
            Left            =   1500
            TabIndex        =   300
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TEST TIME"
            Height          =   315
            Index           =   54
            Left            =   1020
            TabIndex        =   299
            Top             =   3180
            Width           =   1215
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CURR MIN"
            Height          =   285
            Index           =   59
            Left            =   1065
            TabIndex        =   298
            Top             =   3720
            Width           =   1185
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CURR MAX"
            Height          =   285
            Index           =   60
            Left            =   960
            TabIndex        =   297
            Top             =   4260
            Width           =   1290
         End
      End
      Begin VB.Frame fraBox 
         Height          =   4875
         Index           =   43
         Left            =   -67680
         TabIndex        =   269
         Top             =   900
         Width           =   4395
         Begin VB.TextBox txtLinActAngle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   278
            Text            =   "0"
            Top             =   2580
            Width           =   1635
         End
         Begin VB.TextBox txtLinActName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   277
            Text            =   "ACT 01"
            Top             =   420
            Width           =   1635
         End
         Begin VB.TextBox txtLinActFinal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   276
            Text            =   "100"
            Top             =   2040
            Width           =   1635
         End
         Begin VB.TextBox txtLinActHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   275
            Text            =   "999"
            Top             =   1500
            Width           =   1635
         End
         Begin VB.TextBox txtLinActLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   274
            Text            =   "0"
            Top             =   960
            Width           =   1635
         End
         Begin VB.CheckBox chkLinAct 
            Caption         =   "ACT 01 DISP."
            Height          =   315
            Index           =   0
            Left            =   300
            TabIndex        =   273
            Top             =   0
            Width           =   1815
         End
         Begin VB.TextBox txtLinActTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   272
            Text            =   "20"
            Top             =   3120
            Width           =   1635
         End
         Begin VB.TextBox txtLinActCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   271
            Text            =   "0"
            Top             =   3660
            Width           =   1635
         End
         Begin VB.TextBox txtLinActCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   270
            Text            =   "999"
            Top             =   4200
            Width           =   1635
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DOOR STEP"
            Height          =   285
            Index           =   49
            Left            =   915
            TabIndex        =   286
            Top             =   2640
            Width           =   1320
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            Height          =   315
            Index           =   3
            Left            =   1500
            TabIndex        =   285
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FINAL ANGLE"
            Height          =   276
            Index           =   4
            Left            =   684
            TabIndex        =   284
            Top             =   2100
            Width           =   1548
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MAX ANGLE"
            Height          =   315
            Index           =   2
            Left            =   780
            TabIndex        =   283
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MIN ANGLE"
            Height          =   315
            Index           =   1
            Left            =   840
            TabIndex        =   282
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TEST TIME"
            Height          =   315
            Index           =   53
            Left            =   1020
            TabIndex        =   281
            Top             =   3180
            Width           =   1215
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CURR MIN"
            Height          =   285
            Index           =   57
            Left            =   1050
            TabIndex        =   280
            Top             =   3720
            Width           =   1185
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CURR MAX"
            Height          =   285
            Index           =   58
            Left            =   945
            TabIndex        =   279
            Top             =   4260
            Width           =   1290
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "CHECK POINT 4"
         Height          =   4275
         Index           =   8
         Left            =   -70860
         TabIndex        =   251
         Top             =   2160
         Visible         =   0   'False
         Width           =   4395
         Begin VB.TextBox txtLinAct04CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   2460
            TabIndex        =   261
            Text            =   "0"
            Top             =   3480
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct04CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2460
            TabIndex        =   260
            Text            =   "0"
            Top             =   2820
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct04CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   259
            Text            =   "0"
            Top             =   2160
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct04CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2460
            TabIndex        =   258
            Text            =   "0"
            Top             =   1500
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct04CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   257
            Text            =   "0"
            Top             =   840
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct04Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   960
            TabIndex        =   256
            Text            =   "0"
            Top             =   3480
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct04Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   960
            TabIndex        =   255
            Text            =   "0"
            Top             =   2820
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct04Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   960
            TabIndex        =   254
            Text            =   "0"
            Top             =   2160
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct04Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   960
            TabIndex        =   253
            Text            =   "0"
            Top             =   1500
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct04Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   960
            TabIndex        =   252
            Text            =   "0"
            Top             =   840
            Width           =   1275
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "TIME"
            Height          =   285
            Index           =   48
            Left            =   2820
            TabIndex        =   268
            Top             =   420
            Width           =   570
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "PERCENT"
            Height          =   285
            Index           =   47
            Left            =   1020
            TabIndex        =   267
            Top             =   420
            Width           =   1170
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "5"
            Height          =   285
            Index           =   40
            Left            =   420
            TabIndex        =   266
            Top             =   3540
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "4"
            Height          =   285
            Index           =   39
            Left            =   420
            TabIndex        =   265
            Top             =   2880
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "3"
            Height          =   285
            Index           =   38
            Left            =   420
            TabIndex        =   264
            Top             =   2220
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "2"
            Height          =   285
            Index           =   37
            Left            =   420
            TabIndex        =   263
            Top             =   1560
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "1"
            Height          =   285
            Index           =   36
            Left            =   420
            TabIndex        =   262
            Top             =   900
            Width           =   150
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "CHECK POINT 3"
         Height          =   4275
         Index           =   6
         Left            =   -70860
         TabIndex        =   233
         Top             =   1740
         Visible         =   0   'False
         Width           =   4395
         Begin VB.TextBox txtLinAct03CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   2460
            TabIndex        =   243
            Text            =   "0"
            Top             =   3480
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct03CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2460
            TabIndex        =   242
            Text            =   "0"
            Top             =   2820
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct03CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   241
            Text            =   "0"
            Top             =   2160
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct03CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2460
            TabIndex        =   240
            Text            =   "0"
            Top             =   1500
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct03CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   239
            Text            =   "0"
            Top             =   840
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct03Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   960
            TabIndex        =   238
            Text            =   "0"
            Top             =   840
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct03Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   960
            TabIndex        =   237
            Text            =   "0"
            Top             =   1500
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct03Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   960
            TabIndex        =   236
            Text            =   "0"
            Top             =   2160
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct03Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   960
            TabIndex        =   235
            Text            =   "0"
            Top             =   2820
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct03Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   960
            TabIndex        =   234
            Text            =   "0"
            Top             =   3480
            Width           =   1275
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "TIME"
            Height          =   285
            Index           =   46
            Left            =   2820
            TabIndex        =   250
            Top             =   420
            Width           =   570
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "PERCENT"
            Height          =   285
            Index           =   45
            Left            =   1020
            TabIndex        =   249
            Top             =   420
            Width           =   1170
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "1"
            Height          =   285
            Index           =   35
            Left            =   420
            TabIndex        =   248
            Top             =   900
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "2"
            Height          =   285
            Index           =   34
            Left            =   420
            TabIndex        =   247
            Top             =   1560
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "3"
            Height          =   285
            Index           =   33
            Left            =   420
            TabIndex        =   246
            Top             =   2220
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "4"
            Height          =   285
            Index           =   32
            Left            =   420
            TabIndex        =   245
            Top             =   2880
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "5"
            Height          =   285
            Index           =   31
            Left            =   420
            TabIndex        =   244
            Top             =   3540
            Width           =   150
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "CHECK POINT 2"
         Height          =   4275
         Index           =   4
         Left            =   -70860
         TabIndex        =   215
         Top             =   1320
         Visible         =   0   'False
         Width           =   4395
         Begin VB.TextBox txtLinAct02CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   2460
            TabIndex        =   225
            Text            =   "0"
            Top             =   3480
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct02CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2460
            TabIndex        =   224
            Text            =   "0"
            Top             =   2820
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct02CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   223
            Text            =   "0"
            Top             =   2160
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct02CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2460
            TabIndex        =   222
            Text            =   "0"
            Top             =   1500
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct02CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   221
            Text            =   "0"
            Top             =   840
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct02Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   960
            TabIndex        =   220
            Text            =   "0"
            Top             =   3480
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct02Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   960
            TabIndex        =   219
            Text            =   "0"
            Top             =   2820
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct02Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   960
            TabIndex        =   218
            Text            =   "0"
            Top             =   2160
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct02Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   960
            TabIndex        =   217
            Text            =   "0"
            Top             =   1500
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct02Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   960
            TabIndex        =   216
            Text            =   "0"
            Top             =   840
            Width           =   1275
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "TIME"
            Height          =   285
            Index           =   44
            Left            =   2820
            TabIndex        =   232
            Top             =   420
            Width           =   570
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "PERCENT"
            Height          =   285
            Index           =   43
            Left            =   1020
            TabIndex        =   231
            Top             =   420
            Width           =   1170
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "5"
            Height          =   285
            Index           =   30
            Left            =   420
            TabIndex        =   230
            Top             =   3540
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "4"
            Height          =   285
            Index           =   29
            Left            =   420
            TabIndex        =   229
            Top             =   2880
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "3"
            Height          =   285
            Index           =   28
            Left            =   420
            TabIndex        =   228
            Top             =   2220
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "2"
            Height          =   285
            Index           =   27
            Left            =   420
            TabIndex        =   227
            Top             =   1560
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "1"
            Height          =   285
            Index           =   26
            Left            =   420
            TabIndex        =   226
            Top             =   900
            Width           =   150
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "CHECK POINT 1"
         Height          =   4275
         Index           =   3
         Left            =   -70860
         TabIndex        =   197
         Top             =   900
         Visible         =   0   'False
         Width           =   4395
         Begin VB.TextBox txtLinAct01CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   2460
            TabIndex        =   207
            Text            =   "0"
            Top             =   3480
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct01CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2460
            TabIndex        =   206
            Text            =   "0"
            Top             =   2820
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct01CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   205
            Text            =   "0"
            Top             =   2160
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct01CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2460
            TabIndex        =   204
            Text            =   "0"
            Top             =   1500
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct01CheckTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2460
            TabIndex        =   203
            Text            =   "0"
            Top             =   840
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct01Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   960
            TabIndex        =   202
            Text            =   "0"
            Top             =   840
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct01Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   960
            TabIndex        =   201
            Text            =   "0"
            Top             =   1500
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct01Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   960
            TabIndex        =   200
            Text            =   "0"
            Top             =   2160
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct01Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   960
            TabIndex        =   199
            Text            =   "0"
            Top             =   2820
            Width           =   1275
         End
         Begin VB.TextBox txtLinAct01Check 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   960
            TabIndex        =   198
            Text            =   "0"
            Top             =   3480
            Width           =   1275
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "TIME"
            Height          =   285
            Index           =   42
            Left            =   2820
            TabIndex        =   214
            Top             =   420
            Width           =   570
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "PERCENT"
            Height          =   285
            Index           =   41
            Left            =   1020
            TabIndex        =   213
            Top             =   420
            Width           =   1170
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "1"
            Height          =   285
            Index           =   6
            Left            =   420
            TabIndex        =   212
            Top             =   900
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "2"
            Height          =   285
            Index           =   7
            Left            =   420
            TabIndex        =   211
            Top             =   1560
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "3"
            Height          =   285
            Index           =   8
            Left            =   420
            TabIndex        =   210
            Top             =   2220
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "4"
            Height          =   285
            Index           =   9
            Left            =   420
            TabIndex        =   209
            Top             =   2880
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "5"
            Height          =   285
            Index           =   10
            Left            =   420
            TabIndex        =   208
            Top             =   3540
            Width           =   150
         End
      End
      Begin VB.Frame fraBox 
         Height          =   1935
         Index           =   4
         Left            =   -73860
         TabIndex        =   195
         Top             =   4620
         Width           =   4275
         Begin VB.CheckBox chkNvh 
            Caption         =   "NVH TEST"
            Height          =   315
            Left            =   1320
            TabIndex        =   196
            Top             =   900
            Width           =   1632
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "TEST OPTION"
         Height          =   3075
         Index           =   10
         Left            =   -73860
         TabIndex        =   190
         Top             =   6780
         Width           =   4275
         Begin VB.CheckBox chkStallUse 
            Caption         =   "USING STALL"
            Height          =   315
            Left            =   480
            TabIndex        =   194
            Top             =   2400
            Width           =   1935
         End
         Begin VB.CheckBox chkAutoAddress 
            Caption         =   "USING AUTOADDRESSING"
            Height          =   315
            Left            =   480
            TabIndex        =   193
            Top             =   1800
            Width           =   3372
         End
         Begin VB.OptionButton optLinTestType 
            Caption         =   "SEQUENCING TEST"
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   192
            Top             =   1080
            Width           =   2775
         End
         Begin VB.OptionButton optLinTestType 
            Caption         =   "COMBINING TEST"
            Height          =   495
            Index           =   0
            Left            =   480
            TabIndex        =   191
            Top             =   480
            Value           =   -1  'True
            Width           =   2775
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "START MOVE 4"
         Height          =   1875
         Index           =   9
         Left            =   -74280
         TabIndex        =   185
         Top             =   2160
         Visible         =   0   'False
         Width           =   4395
         Begin VB.OptionButton optLinAct04FirstMove 
            Caption         =   "OPEN"
            Height          =   495
            Index           =   1
            Left            =   2700
            TabIndex        =   188
            Top             =   420
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optLinAct04FirstMove 
            Caption         =   "CLOSE"
            Height          =   495
            Index           =   0
            Left            =   780
            TabIndex        =   187
            Top             =   420
            Width           =   1095
         End
         Begin VB.TextBox txtLinActMove 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   2130
            TabIndex        =   186
            Text            =   "0"
            Top             =   1080
            Width           =   1635
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MOVE ANGLE"
            Height          =   285
            Index           =   22
            Left            =   300
            TabIndex        =   189
            Top             =   1140
            Width           =   1605
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "START MOVE 3"
         Height          =   1875
         Index           =   7
         Left            =   -74280
         TabIndex        =   180
         Top             =   1740
         Visible         =   0   'False
         Width           =   4395
         Begin VB.OptionButton optLinAct03FirstMove 
            Caption         =   "CLOSE"
            Height          =   495
            Index           =   0
            Left            =   780
            TabIndex        =   183
            Top             =   420
            Width           =   1095
         End
         Begin VB.OptionButton optLinAct03FirstMove 
            Caption         =   "OPEN"
            Height          =   495
            Index           =   1
            Left            =   2700
            TabIndex        =   182
            Top             =   420
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.TextBox txtLinActMove 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   2160
            TabIndex        =   181
            Text            =   "0"
            Top             =   1080
            Width           =   1635
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MOVE ANGLE"
            Height          =   285
            Index           =   17
            Left            =   360
            TabIndex        =   184
            Top             =   1140
            Width           =   1605
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "START MOVE 2"
         Height          =   1875
         Index           =   5
         Left            =   -74280
         TabIndex        =   175
         Top             =   1320
         Visible         =   0   'False
         Width           =   4395
         Begin VB.OptionButton optLinAct02FirstMove 
            Caption         =   "OPEN"
            Height          =   495
            Index           =   1
            Left            =   2700
            TabIndex        =   178
            Top             =   420
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optLinAct02FirstMove 
            Caption         =   "CLOSE"
            Height          =   495
            Index           =   0
            Left            =   780
            TabIndex        =   177
            Top             =   420
            Width           =   1095
         End
         Begin VB.TextBox txtLinActMove 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   2310
            TabIndex        =   176
            Text            =   "0"
            Top             =   1080
            Width           =   1635
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MOVE ANGLE"
            Height          =   285
            Index           =   12
            Left            =   480
            TabIndex        =   179
            Top             =   1140
            Width           =   1605
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "START MOVE 1"
         Height          =   1875
         Index           =   2
         Left            =   -74280
         TabIndex        =   170
         Top             =   900
         Visible         =   0   'False
         Width           =   4395
         Begin VB.OptionButton optLinAct01FirstMove 
            Caption         =   "CLOSE"
            Height          =   495
            Index           =   0
            Left            =   780
            TabIndex        =   173
            Top             =   420
            Width           =   1095
         End
         Begin VB.OptionButton optLinAct01FirstMove 
            Caption         =   "OPEN"
            Height          =   495
            Index           =   1
            Left            =   2700
            TabIndex        =   172
            Top             =   420
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.TextBox txtLinActMove 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   2250
            TabIndex        =   171
            Text            =   "0"
            Top             =   1140
            Width           =   1635
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MOVE ANGLE"
            Height          =   285
            Index           =   5
            Left            =   420
            TabIndex        =   174
            Top             =   1200
            Width           =   1605
         End
      End
      Begin VB.Frame fraBox 
         Caption         =   " LEAK "
         Height          =   2535
         Index           =   5
         Left            =   -73920
         TabIndex        =   163
         Top             =   1140
         Width           =   6795
         Begin VB.CheckBox chkLeak 
            Caption         =   "EVA LEAK TEST"
            Height          =   315
            Index           =   1
            Left            =   600
            TabIndex        =   168
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox txtLeakName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   3180
            TabIndex        =   167
            Text            =   "EVA LEAK"
            Top             =   1140
            Width           =   2955
         End
         Begin VB.CheckBox chkLeak 
            Caption         =   "HTR LEAK TEST"
            Height          =   315
            Index           =   0
            Left            =   600
            TabIndex        =   166
            Top             =   660
            Width           =   2295
         End
         Begin VB.TextBox txtLeakName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   3180
            TabIndex        =   165
            Text            =   "HTR LEAK"
            Top             =   600
            Width           =   2955
         End
         Begin VB.TextBox txtLeakModel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Left            =   3180
            TabIndex        =   164
            Text            =   "0"
            Top             =   1680
            Width           =   2955
         End
         Begin VB.Label Label4 
            Caption         =   "LEAK GROUP"
            Height          =   315
            Index           =   0
            Left            =   1080
            TabIndex        =   169
            Top             =   1740
            Width           =   1515
         End
      End
      Begin VB.Frame fraBox 
         Caption         =   "MODEL TYPE"
         Height          =   2655
         Index           =   1
         Left            =   -74460
         TabIndex        =   155
         Top             =   1200
         Width           =   3795
         Begin VB.TextBox txtModelRHDList 
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
            Left            =   1980
            TabIndex        =   162
            Text            =   "55"
            Top             =   1980
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.CheckBox chkModelTypeUse 
            Caption         =   "BOTTOM SENSOR 1"
            Height          =   375
            Left            =   180
            TabIndex        =   161
            Top             =   540
            Width           =   3375
         End
         Begin VB.TextBox txtModelLHDList 
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
            Left            =   300
            TabIndex        =   160
            Text            =   "#55"
            Top             =   1980
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.OptionButton optModelType 
            Caption         =   "LHD"
            Height          =   495
            Index           =   0
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   159
            Top             =   1080
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.OptionButton optModelType 
            Caption         =   "RHD"
            Height          =   495
            Index           =   1
            Left            =   1980
            Style           =   1  'Graphical
            TabIndex        =   158
            Top             =   1080
            Width           =   1515
         End
         Begin VB.TextBox txtLHDPartName 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
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
            Left            =   300
            TabIndex        =   157
            Top             =   1560
            Width           =   1515
         End
         Begin VB.TextBox txtRHDPartName 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
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
            Left            =   1980
            TabIndex        =   156
            Top             =   1560
            Width           =   1515
         End
      End
      Begin VB.Frame fraBox 
         Caption         =   "PARTS 2"
         Height          =   6675
         Index           =   35
         Left            =   -64980
         TabIndex        =   19
         Top             =   2520
         Width           =   9495
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   79
            Left            =   6240
            TabIndex        =   108
            Top             =   4500
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   78
            Left            =   6240
            TabIndex        =   107
            Top             =   4980
            Width           =   555
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "27"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   79
            Left            =   4860
            TabIndex        =   106
            Top             =   4980
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "26"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   78
            Left            =   4860
            TabIndex        =   105
            Top             =   4500
            Width           =   675
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   77
            Left            =   8400
            TabIndex        =   104
            Top             =   4980
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   76
            Left            =   8400
            TabIndex        =   103
            Top             =   4500
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   75
            Left            =   8400
            TabIndex        =   102
            Top             =   4020
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   74
            Left            =   8400
            TabIndex        =   101
            Top             =   3540
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   73
            Left            =   8400
            TabIndex        =   100
            Top             =   3060
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   72
            Left            =   8400
            TabIndex        =   99
            Top             =   2580
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   71
            Left            =   6240
            TabIndex        =   98
            Top             =   4020
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   70
            Left            =   6240
            TabIndex        =   97
            Top             =   3540
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   69
            Left            =   6240
            TabIndex        =   96
            Top             =   3060
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   68
            Left            =   6240
            TabIndex        =   95
            Top             =   2580
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   67
            Left            =   6240
            TabIndex        =   94
            Top             =   2100
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   66
            Left            =   6240
            TabIndex        =   93
            Top             =   1620
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   65
            Left            =   4080
            TabIndex        =   92
            Top             =   4980
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   64
            Left            =   4080
            TabIndex        =   91
            Top             =   4500
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   63
            Left            =   4080
            TabIndex        =   90
            Top             =   4020
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   62
            Left            =   4080
            TabIndex        =   89
            Top             =   3540
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   61
            Left            =   4080
            TabIndex        =   88
            Top             =   3060
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   60
            Left            =   4080
            TabIndex        =   87
            Top             =   2580
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   59
            Left            =   4080
            TabIndex        =   86
            Top             =   2100
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   58
            Left            =   4080
            TabIndex        =   85
            Top             =   1620
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   57
            Left            =   1920
            TabIndex        =   84
            Top             =   4980
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   56
            Left            =   1920
            TabIndex        =   83
            Top             =   4500
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   55
            Left            =   1920
            TabIndex        =   82
            Top             =   4020
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   54
            Left            =   1920
            TabIndex        =   81
            Top             =   3540
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   53
            Left            =   1920
            TabIndex        =   80
            Top             =   3060
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   52
            Left            =   1920
            TabIndex        =   79
            Top             =   2580
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   51
            Left            =   1920
            TabIndex        =   78
            Top             =   2100
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   50
            Left            =   1920
            TabIndex        =   77
            Top             =   1620
            Width           =   555
         End
         Begin VB.Frame fraPartsScript 
            Caption         =   " INFORMATION "
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   1
            Left            =   540
            TabIndex        =   68
            Top             =   360
            Width           =   8415
            Begin VB.Label lblTemp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   6360
               TabIndex        =   76
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblTemp 
               AutoSize        =   -1  'True
               Caption         =   "NAME ""#"" START PART"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   465
               Index           =   16
               Left            =   6900
               TabIndex        =   75
               Top             =   240
               Width           =   1140
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblTemp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "OFF"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   2340
               TabIndex        =   74
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblTemp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "O N"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   420
               TabIndex        =   73
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblTemp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00A0A0A0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   4440
               TabIndex        =   72
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblTemp 
               AutoSize        =   -1  'True
               Caption         =   "ON OK"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   12
               Left            =   960
               TabIndex        =   71
               Top             =   360
               Width           =   615
            End
            Begin VB.Label lblTemp 
               AutoSize        =   -1  'True
               Caption         =   "OFF OK"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   11
               Left            =   2880
               TabIndex        =   70
               Top             =   360
               Width           =   690
            End
            Begin VB.Label lblTemp 
               AutoSize        =   -1  'True
               Caption         =   "NOT USE"
               BeginProperty Font 
                  Name            =   "³ª´®°íµñ"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   10
               Left            =   4980
               TabIndex        =   69
               Top             =   360
               Width           =   795
            End
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   77
            Left            =   2700
            TabIndex        =   67
            Top             =   1620
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "11"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   76
            Left            =   2700
            TabIndex        =   66
            Top             =   2100
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "12"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   75
            Left            =   2700
            TabIndex        =   65
            Top             =   2580
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "13"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   74
            Left            =   2700
            TabIndex        =   64
            Top             =   3060
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "14"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   73
            Left            =   2700
            TabIndex        =   63
            Top             =   3540
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   72
            Left            =   2700
            TabIndex        =   62
            Top             =   4020
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "16"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   71
            Left            =   2700
            TabIndex        =   61
            Top             =   4500
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "17"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   70
            Left            =   2700
            TabIndex        =   60
            Top             =   4980
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   69
            Left            =   540
            TabIndex        =   59
            Top             =   1620
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   68
            Left            =   540
            TabIndex        =   58
            Top             =   2100
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   67
            Left            =   540
            TabIndex        =   57
            Top             =   2580
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   66
            Left            =   540
            TabIndex        =   56
            Top             =   3060
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   65
            Left            =   540
            TabIndex        =   55
            Top             =   3540
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   64
            Left            =   540
            TabIndex        =   54
            Top             =   4020
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   63
            Left            =   540
            TabIndex        =   53
            Top             =   4500
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   62
            Left            =   540
            TabIndex        =   52
            Top             =   4980
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   61
            Left            =   4860
            TabIndex        =   51
            Top             =   1620
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "21"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   60
            Left            =   4860
            TabIndex        =   50
            Top             =   2100
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "22"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   59
            Left            =   4860
            TabIndex        =   49
            Top             =   2580
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "23"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   58
            Left            =   4860
            TabIndex        =   48
            Top             =   3060
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "24"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   57
            Left            =   4860
            TabIndex        =   47
            Top             =   3540
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "25"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   56
            Left            =   4860
            TabIndex        =   46
            Top             =   4020
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "32"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   55
            Left            =   7020
            TabIndex        =   45
            Top             =   2580
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "33"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   54
            Left            =   7020
            TabIndex        =   44
            Top             =   3060
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "34"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   53
            Left            =   7020
            TabIndex        =   43
            Top             =   3540
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "35"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   52
            Left            =   7020
            TabIndex        =   42
            Top             =   4020
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "36"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   51
            Left            =   7020
            TabIndex        =   41
            Top             =   4500
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   50
            Left            =   7020
            TabIndex        =   40
            Top             =   4980
            Width           =   675
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   49
            Left            =   8400
            TabIndex        =   39
            Top             =   1620
            Width           =   555
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   49
            Left            =   7020
            TabIndex        =   38
            Top             =   1620
            Width           =   675
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   48
            Left            =   8400
            TabIndex        =   37
            Top             =   2100
            Width           =   555
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   48
            Left            =   7020
            TabIndex        =   36
            Top             =   2100
            Width           =   675
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   47
            Left            =   8400
            TabIndex        =   35
            Top             =   5940
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   46
            Left            =   8400
            TabIndex        =   34
            Top             =   5460
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   45
            Left            =   6240
            TabIndex        =   33
            Top             =   5940
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   44
            Left            =   6240
            TabIndex        =   32
            Top             =   5460
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   43
            Left            =   4080
            TabIndex        =   31
            Top             =   5940
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   42
            Left            =   4080
            TabIndex        =   30
            Top             =   5460
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   41
            Left            =   1920
            TabIndex        =   29
            Top             =   5940
            Width           =   555
         End
         Begin VB.TextBox txtPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Index           =   40
            Left            =   1920
            TabIndex        =   28
            Top             =   5460
            Width           =   555
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "39"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   47
            Left            =   7020
            TabIndex        =   27
            Top             =   5940
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "38"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   46
            Left            =   7020
            TabIndex        =   26
            Top             =   5460
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "29"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   45
            Left            =   4860
            TabIndex        =   25
            Top             =   5940
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "28"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   44
            Left            =   4860
            TabIndex        =   24
            Top             =   5460
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   43
            Left            =   540
            TabIndex        =   23
            Top             =   5940
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   42
            Left            =   540
            TabIndex        =   22
            Top             =   5460
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "19"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   41
            Left            =   2700
            TabIndex        =   21
            Top             =   5940
            Width           =   675
         End
         Begin VB.CheckBox chkPart 
            Caption         =   "18"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   40
            Left            =   2700
            TabIndex        =   20
            Top             =   5460
            Width           =   675
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   79
            Left            =   3480
            TabIndex        =   148
            Top             =   1620
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   78
            Left            =   3480
            TabIndex        =   147
            Top             =   2100
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   77
            Left            =   3480
            TabIndex        =   146
            Top             =   2580
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   76
            Left            =   3480
            TabIndex        =   145
            Top             =   3060
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   75
            Left            =   3480
            TabIndex        =   144
            Top             =   3540
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   74
            Left            =   3480
            TabIndex        =   143
            Top             =   4020
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   73
            Left            =   3480
            TabIndex        =   142
            Top             =   4500
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   72
            Left            =   3480
            TabIndex        =   141
            Top             =   4980
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   71
            Left            =   1320
            TabIndex        =   140
            Top             =   1620
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   70
            Left            =   1320
            TabIndex        =   139
            Top             =   2100
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   69
            Left            =   1320
            TabIndex        =   138
            Top             =   2580
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   68
            Left            =   1320
            TabIndex        =   137
            Top             =   3060
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   67
            Left            =   1320
            TabIndex        =   136
            Top             =   3540
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   66
            Left            =   1320
            TabIndex        =   135
            Top             =   4020
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   65
            Left            =   1320
            TabIndex        =   134
            Top             =   4500
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   64
            Left            =   1320
            TabIndex        =   133
            Top             =   4980
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   63
            Left            =   5640
            TabIndex        =   132
            Top             =   1620
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   62
            Left            =   5640
            TabIndex        =   131
            Top             =   2100
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   61
            Left            =   5640
            TabIndex        =   130
            Top             =   2580
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   60
            Left            =   5640
            TabIndex        =   129
            Top             =   3060
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   59
            Left            =   5640
            TabIndex        =   128
            Top             =   3540
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   58
            Left            =   5640
            TabIndex        =   127
            Top             =   4020
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   57
            Left            =   5640
            TabIndex        =   126
            Top             =   4500
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   56
            Left            =   5640
            TabIndex        =   125
            Top             =   4980
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   55
            Left            =   7800
            TabIndex        =   124
            Top             =   1620
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   54
            Left            =   7800
            TabIndex        =   123
            Top             =   2100
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   53
            Left            =   7800
            TabIndex        =   122
            Top             =   2580
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   52
            Left            =   7800
            TabIndex        =   121
            Top             =   3060
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   51
            Left            =   7800
            TabIndex        =   120
            Top             =   3540
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   50
            Left            =   7800
            TabIndex        =   119
            Top             =   4020
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   49
            Left            =   7800
            TabIndex        =   118
            Top             =   4500
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   48
            Left            =   7800
            TabIndex        =   117
            Top             =   4980
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   47
            Left            =   7800
            TabIndex        =   116
            Top             =   5940
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   46
            Left            =   7800
            TabIndex        =   115
            Top             =   5460
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   45
            Left            =   5640
            TabIndex        =   114
            Top             =   5940
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   44
            Left            =   5640
            TabIndex        =   113
            Top             =   5460
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   43
            Left            =   1320
            TabIndex        =   112
            Top             =   5940
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   42
            Left            =   1320
            TabIndex        =   111
            Top             =   5460
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   41
            Left            =   3480
            TabIndex        =   110
            Top             =   5940
            Width           =   555
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "--"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   40
            Left            =   3480
            TabIndex        =   109
            Top             =   5460
            Width           =   555
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   615
         Index           =   0
         Left            =   -73260
         TabIndex        =   149
         Top             =   3360
         Width           =   3315
         _Version        =   65536
         _ExtentX        =   5847
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "BARCODE"
         BackColor       =   13160660
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
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Index           =   1
         Left            =   -70860
         TabIndex        =   150
         Top             =   3960
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "HANON"
         BackColor       =   13160660
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
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   2
         Left            =   -73260
         TabIndex        =   151
         Top             =   2580
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "MODEL :"
         BackColor       =   13160660
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
         Alignment       =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   3
         Left            =   -73260
         TabIndex        =   152
         Top             =   2940
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "HMC P/NO :"
         BackColor       =   13160660
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
         Alignment       =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   5
         Left            =   -73260
         TabIndex        =   153
         Top             =   4320
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "LOT NO :"
         BackColor       =   13160660
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
         Alignment       =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   6
         Left            =   -72360
         TabIndex        =   154
         Top             =   4320
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "YYYY.MM.DD.SERIAL"
         BackColor       =   13160660
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
         Alignment       =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Index           =   7
         Left            =   -67800
         TabIndex        =   380
         Top             =   3720
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "YYYYMMDDSERIAL"
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
         Alignment       =   1
      End
      Begin Threed.SSPanel pnlBarcodeFrame 
         Height          =   1875
         Index           =   0
         Left            =   -69540
         TabIndex        =   435
         Top             =   1020
         Width           =   3795
         _Version        =   65536
         _ExtentX        =   6694
         _ExtentY        =   3307
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
         Begin VB.TextBox txtBarCode 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   437
            Text            =   "CODE"
            Top             =   1380
            Width           =   1335
         End
         Begin VB.TextBox txtBarCode 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   436
            Text            =   "MODEL"
            Top             =   120
            Width           =   3315
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   795
            Index           =   4
            Left            =   240
            TabIndex        =   438
            Top             =   540
            Width           =   3315
            _Version        =   65536
            _ExtentX        =   5847
            _ExtentY        =   1402
            _StockProps     =   15
            Caption         =   "BARCODE"
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
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   315
            Index           =   8
            Left            =   1740
            TabIndex        =   439
            Top             =   1440
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "YYYYMMDDSERIAL"
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
            Alignment       =   1
         End
      End
      Begin Threed.SSPanel pnlBox 
         Height          =   615
         Index           =   17
         Left            =   3900
         TabIndex        =   930
         Top             =   1680
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
         _ExtentY        =   1085
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
         Begin VB.OptionButton optAct01Direction 
            Caption         =   "ONE-WAY"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   1980
            TabIndex        =   932
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton optAct01Direction 
            Caption         =   "BOTH-WAY"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   931
            Top             =   120
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "TEST TYPE"
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   34
            Left            =   120
            TabIndex        =   933
            Top             =   120
            Width           =   1215
         End
      End
      Begin Threed.SSPanel pnlBox 
         Height          =   615
         Index           =   18
         Left            =   3900
         TabIndex        =   934
         Top             =   2400
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
         _ExtentY        =   1085
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
         Begin VB.OptionButton optAct02Direction 
            Caption         =   "BOTH-WAY"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   936
            Top             =   120
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optAct02Direction 
            Caption         =   "ONE-WAY"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   1980
            TabIndex        =   935
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "TEST TYPE"
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   35
            Left            =   120
            TabIndex        =   937
            Top             =   120
            Width           =   1215
         End
      End
      Begin Threed.SSPanel pnlBox 
         Height          =   615
         Index           =   19
         Left            =   3900
         TabIndex        =   938
         Top             =   3120
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
         _ExtentY        =   1085
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
         Begin VB.OptionButton optAct03Direction 
            Caption         =   "ONE-WAY"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   1980
            TabIndex        =   940
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton optAct03Direction 
            Caption         =   "BOTH-WAY"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   939
            Top             =   120
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "TEST TYPE"
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   38
            Left            =   120
            TabIndex        =   941
            Top             =   120
            Width           =   1215
         End
      End
      Begin Threed.SSPanel pnlBox 
         Height          =   615
         Index           =   20
         Left            =   3900
         TabIndex        =   942
         Top             =   3900
         Width           =   4155
         _Version        =   65536
         _ExtentX        =   7329
         _ExtentY        =   1085
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
         Begin VB.OptionButton optAct04Direction 
            Caption         =   "BOTH-WAY"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   2640
            TabIndex        =   944
            Top             =   120
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.OptionButton optAct04Direction 
            Caption         =   "ONE-WAY"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   1260
            TabIndex        =   943
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "TEST TYPE"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   42
            Left            =   120
            TabIndex        =   945
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ACT NO."
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   51
         Left            =   480
         TabIndex        =   640
         Top             =   3180
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ACT NO."
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   50
         Left            =   480
         TabIndex        =   638
         Top             =   2640
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ACT NO."
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   49
         Left            =   480
         TabIndex        =   636
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ACT NO."
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   48
         Left            =   480
         TabIndex        =   634
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LIN SPEED"
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   1
         Left            =   -74640
         TabIndex        =   358
         Top             =   4560
         Width           =   1155
      End
   End
   Begin Threed.SSPanel SSPanel 
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18855
      _Version        =   65536
      _ExtentX        =   33258
      _ExtentY        =   4048
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
      Begin VB.ListBox lstSetupMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   10200
         TabIndex        =   7
         Top             =   120
         Width           =   5055
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
         Height          =   2055
         Left            =   15360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3375
      End
      Begin VB.ComboBox cboCarType 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   780
         ItemData        =   "frmSetup.frx":05BF
         Left            =   120
         List            =   "frmSetup.frx":05C1
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   120
         Width           =   9975
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   26.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   7080
         TabIndex        =   4
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   26.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4320
         TabIndex        =   3
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   2655
      End
      Begin VB.ComboBox cboLoad 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   630
         ItemData        =   "frmSetup.frx":05C3
         Left            =   120
         List            =   "frmSetup.frx":05C5
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   4155
      End
      Begin VB.CommandButton btnLoading 
         Caption         =   "LOAD MODEL"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   20.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1620
         UseMaskColor    =   -1  'True
         Width           =   4095
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   12255
      Left            =   120
      TabIndex        =   8
      Top             =   2580
      Width           =   18855
      _ExtentX        =   33258
      _ExtentY        =   21616
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   1235
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "GENERAL"
      TabPicture(0)   =   "frmSetup.frx":05C7
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraBox(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraBox(7)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraBox(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraBox(48)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "ACT 1 && 2"
      TabPicture(1)   =   "frmSetup.frx":05E3
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "pnlBox(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "pnlBox(4)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "ACT 3 && 4"
      TabPicture(2)   =   "frmSetup.frx":05FF
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pnlBox(10)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "pnlBox(7)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "OFFSET"
      TabPicture(3)   =   "frmSetup.frx":061B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSPanel(1)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fraBox 
         Caption         =   "STEPPING MOTOR SETTING"
         ForeColor       =   &H00000000&
         Height          =   3075
         Index           =   48
         Left            =   -74040
         TabIndex        =   645
         Top             =   8173
         Width           =   16875
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   15840
            TabIndex        =   726
            Text            =   "0"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   14880
            TabIndex        =   725
            Text            =   "0"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   13
            Left            =   13920
            TabIndex        =   724
            Text            =   "0"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   12960
            TabIndex        =   723
            Text            =   "10"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   12000
            TabIndex        =   722
            Text            =   "1100"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   11040
            TabIndex        =   721
            Text            =   "0"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   10080
            TabIndex        =   720
            Text            =   "0"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   9120
            TabIndex        =   719
            Text            =   "0"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   8160
            TabIndex        =   718
            Text            =   "0"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   7200
            TabIndex        =   717
            Text            =   "850"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   6240
            TabIndex        =   716
            Text            =   "0"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   5280
            TabIndex        =   715
            Text            =   "0"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   4320
            TabIndex        =   714
            Text            =   "0"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   713
            Text            =   "0"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2400
            TabIndex        =   712
            Text            =   "1"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct04Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   711
            Text            =   "300"
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   15840
            TabIndex        =   710
            Text            =   "0"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   14880
            TabIndex        =   709
            Text            =   "0"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   13
            Left            =   13920
            TabIndex        =   708
            Text            =   "0"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   12960
            TabIndex        =   707
            Text            =   "10"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   12000
            TabIndex        =   706
            Text            =   "1100"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   11040
            TabIndex        =   705
            Text            =   "0"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   10080
            TabIndex        =   704
            Text            =   "0"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   9120
            TabIndex        =   703
            Text            =   "0"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   8160
            TabIndex        =   702
            Text            =   "0"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   7200
            TabIndex        =   701
            Text            =   "850"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   6240
            TabIndex        =   700
            Text            =   "0"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   5280
            TabIndex        =   699
            Text            =   "0"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   4320
            TabIndex        =   698
            Text            =   "0"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   697
            Text            =   "0"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2400
            TabIndex        =   696
            Text            =   "1"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct03Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   695
            Text            =   "300"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   15840
            TabIndex        =   694
            Text            =   "0"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   14880
            TabIndex        =   693
            Text            =   "0"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   13
            Left            =   13920
            TabIndex        =   692
            Text            =   "0"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   12960
            TabIndex        =   691
            Text            =   "10"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   12000
            TabIndex        =   690
            Text            =   "1100"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   11040
            TabIndex        =   689
            Text            =   "0"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   10080
            TabIndex        =   688
            Text            =   "0"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   9120
            TabIndex        =   687
            Text            =   "0"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   8160
            TabIndex        =   686
            Text            =   "0"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   7200
            TabIndex        =   685
            Text            =   "850"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   6240
            TabIndex        =   684
            Text            =   "0"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   5280
            TabIndex        =   683
            Text            =   "0"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   4320
            TabIndex        =   682
            Text            =   "0"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   681
            Text            =   "0"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2400
            TabIndex        =   680
            Text            =   "1"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct02Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   679
            Text            =   "300"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   15840
            TabIndex        =   678
            Text            =   "0"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   14880
            TabIndex        =   677
            Text            =   "0"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   13
            Left            =   13920
            TabIndex        =   676
            Text            =   "0"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   12960
            TabIndex        =   675
            Text            =   "10"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   12000
            TabIndex        =   674
            Text            =   "1500"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   11040
            TabIndex        =   673
            Text            =   "0"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   10080
            TabIndex        =   672
            Text            =   "1350"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   9120
            TabIndex        =   664
            Text            =   "1000"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   8160
            TabIndex        =   663
            Text            =   "600"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   7200
            TabIndex        =   662
            Text            =   "130"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   6240
            TabIndex        =   661
            Text            =   "0"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   5280
            TabIndex        =   660
            Text            =   "0"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   4320
            TabIndex        =   659
            Text            =   "0"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   658
            Text            =   "1"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2400
            TabIndex        =   649
            Text            =   "1"
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtStepAct01Arr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   647
            Text            =   "300"
            Top             =   960
            Width           =   915
         End
         Begin Threed.SSPanel pnlSteppingTitle 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   646
            Top             =   960
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "ACT 01"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.01
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   3
            Left            =   1440
            TabIndex        =   648
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "FREQ"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   4
            Left            =   2400
            TabIndex        =   650
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "MODE"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   5
            Left            =   3360
            TabIndex        =   651
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "START DIR."
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   6
            Left            =   4320
            TabIndex        =   652
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "START STEP"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   7
            Left            =   5280
            TabIndex        =   653
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "DELIVERY STEP"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   8
            Left            =   6240
            TabIndex        =   654
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "POS1"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   9
            Left            =   7200
            TabIndex        =   655
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "POS2"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   10
            Left            =   8160
            TabIndex        =   656
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "POS3"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   11
            Left            =   9120
            TabIndex        =   657
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "POS4"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   12
            Left            =   10080
            TabIndex        =   665
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "POS5"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   13
            Left            =   11040
            TabIndex        =   666
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "POS6"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   14
            Left            =   12000
            TabIndex        =   667
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "CLOSE STEP MAX"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   15
            Left            =   12960
            TabIndex        =   668
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "CL VOLT MARGIN"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   16
            Left            =   13920
            TabIndex        =   669
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "DATA2"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   17
            Left            =   14880
            TabIndex        =   670
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "DATA3"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   18
            Left            =   15840
            TabIndex        =   671
            Top             =   360
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "DATA4"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   22
            Left            =   2400
            TabIndex        =   727
            Top             =   2400
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Uni(0) Bi(1)"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   23
            Left            =   3360
            TabIndex        =   728
            Top             =   2400
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "CW(0) CCW(1)"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
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
         Begin Threed.SSPanel pnlSteppingTitle 
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   729
            Top             =   1320
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "ACT 02"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.01
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
         Begin Threed.SSPanel pnlSteppingTitle 
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   730
            Top             =   1680
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "ACT 03"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.01
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
         Begin Threed.SSPanel pnlSteppingTitle 
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   731
            Top             =   2040
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "ACT 04"
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.01
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
      End
      Begin VB.Frame fraBox 
         Caption         =   "SENSOR"
         Height          =   3075
         Index           =   6
         Left            =   -73260
         TabIndex        =   608
         Top             =   4166
         Width           =   12735
         Begin VB.TextBox txtSensorTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   10680
            TabIndex        =   884
            Text            =   "3"
            Top             =   2340
            Width           =   1695
         End
         Begin VB.CheckBox chkSensor 
            Caption         =   "SENSOR 6"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   10680
            TabIndex        =   883
            Top             =   420
            Width           =   1695
         End
         Begin VB.TextBox txtSensorName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   10680
            TabIndex        =   882
            Text            =   "SENSOR 6"
            Top             =   900
            Width           =   1695
         End
         Begin VB.TextBox txtSensorCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   10680
            TabIndex        =   881
            Text            =   "3"
            Top             =   1860
            Width           =   1695
         End
         Begin VB.TextBox txtSensorCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   5
            Left            =   10680
            TabIndex        =   880
            Text            =   "0.5"
            Top             =   1380
            Width           =   1695
         End
         Begin VB.TextBox txtSensorTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   8940
            TabIndex        =   879
            Text            =   "3"
            Top             =   2340
            Width           =   1695
         End
         Begin VB.CheckBox chkSensor 
            Caption         =   "SENSOR 5"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   8940
            TabIndex        =   878
            Top             =   420
            Width           =   1695
         End
         Begin VB.TextBox txtSensorName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   8940
            TabIndex        =   877
            Text            =   "SENSOR 5"
            Top             =   900
            Width           =   1695
         End
         Begin VB.TextBox txtSensorCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   8940
            TabIndex        =   876
            Text            =   "3"
            Top             =   1860
            Width           =   1695
         End
         Begin VB.TextBox txtSensorCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   8940
            TabIndex        =   875
            Text            =   "0.5"
            Top             =   1380
            Width           =   1695
         End
         Begin VB.TextBox txtSensorCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   5460
            TabIndex        =   632
            Text            =   "0.5"
            Top             =   1380
            Width           =   1695
         End
         Begin VB.TextBox txtSensorCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   5460
            TabIndex        =   631
            Text            =   "3"
            Top             =   1860
            Width           =   1695
         End
         Begin VB.TextBox txtSensorName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   5460
            TabIndex        =   630
            Text            =   "SENSOR 3"
            Top             =   900
            Width           =   1695
         End
         Begin VB.CheckBox chkSensor 
            Caption         =   "SENSOR 3"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   5460
            TabIndex        =   629
            Top             =   420
            Width           =   1695
         End
         Begin VB.TextBox txtSensorTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   5460
            TabIndex        =   628
            Text            =   "3"
            Top             =   2340
            Width           =   1695
         End
         Begin VB.TextBox txtSensorCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   7200
            TabIndex        =   627
            Text            =   "0.5"
            Top             =   1380
            Width           =   1695
         End
         Begin VB.TextBox txtSensorCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   7200
            TabIndex        =   626
            Text            =   "3"
            Top             =   1860
            Width           =   1695
         End
         Begin VB.TextBox txtSensorName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   7200
            TabIndex        =   625
            Text            =   "SENSOR 4"
            Top             =   900
            Width           =   1695
         End
         Begin VB.CheckBox chkSensor 
            Caption         =   "SENSOR 4"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   7200
            TabIndex        =   624
            Top             =   420
            Width           =   1695
         End
         Begin VB.TextBox txtSensorTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   7200
            TabIndex        =   623
            Text            =   "3"
            Top             =   2340
            Width           =   1695
         End
         Begin VB.TextBox txtSensorCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   1980
            TabIndex        =   618
            Text            =   "0.5"
            Top             =   1380
            Width           =   1695
         End
         Begin VB.TextBox txtSensorCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   1980
            TabIndex        =   617
            Text            =   "3"
            Top             =   1860
            Width           =   1695
         End
         Begin VB.TextBox txtSensorName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   1980
            TabIndex        =   616
            Text            =   "SENSOR 1"
            Top             =   900
            Width           =   1695
         End
         Begin VB.TextBox txtSensorCurrLo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   3720
            TabIndex        =   615
            Text            =   "0.5"
            Top             =   1380
            Width           =   1695
         End
         Begin VB.TextBox txtSensorCurrHi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   3720
            TabIndex        =   614
            Text            =   "3"
            Top             =   1860
            Width           =   1695
         End
         Begin VB.TextBox txtSensorName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   3720
            TabIndex        =   613
            Text            =   "SENSOR 2"
            Top             =   900
            Width           =   1695
         End
         Begin VB.CheckBox chkSensor 
            Caption         =   "SENSOR 1"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1980
            TabIndex        =   612
            Top             =   420
            Width           =   1695
         End
         Begin VB.CheckBox chkSensor 
            Caption         =   "SENSOR 2"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   3720
            TabIndex        =   611
            Top             =   420
            Width           =   1695
         End
         Begin VB.TextBox txtSensorTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   1980
            TabIndex        =   610
            Text            =   "3"
            Top             =   2340
            Width           =   1695
         End
         Begin VB.TextBox txtSensorTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   3720
            TabIndex        =   609
            Text            =   "3"
            Top             =   2340
            Width           =   1695
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MIN"
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   19
            Left            =   1095
            TabIndex        =   622
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   621
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MAX"
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   32
            Left            =   990
            TabIndex        =   620
            Top             =   1920
            Width           =   585
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TIME (Sec)"
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   6
            Left            =   375
            TabIndex        =   619
            Top             =   2400
            Width           =   1200
         End
      End
      Begin VB.Frame fraBox 
         Height          =   1935
         Index           =   7
         Left            =   -68160
         TabIndex        =   13
         Top             =   1633
         Width           =   4635
         Begin VB.TextBox txtFileName 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            MaxLength       =   15
            TabIndex        =   15
            Top             =   1020
            Width           =   1755
         End
         Begin VB.CheckBox chkSaveUse 
            Caption         =   "FILE NAME FORM"
            Height          =   315
            Left            =   240
            TabIndex        =   14
            Top             =   0
            Value           =   1  'Checked
            Width           =   2352
         End
         Begin VB.Label lblTemp 
            AutoSize        =   -1  'True
            Caption         =   "SAVE AS :"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   18
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   900
         End
         Begin VB.Label lblTemp 
            AutoSize        =   -1  'True
            Caption         =   "_DATE(AUTOMATIC).CSV "
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   19
            Left            =   2040
            TabIndex        =   16
            Top             =   1080
            Width           =   2370
         End
      End
      Begin VB.Frame fraBox 
         Caption         =   "TEST VOLT"
         ForeColor       =   &H00000000&
         Height          =   1935
         Index           =   0
         Left            =   -73260
         TabIndex        =   9
         Top             =   1633
         Width           =   4875
         Begin VB.TextBox txtTestVolt 
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
            Left            =   2160
            TabIndex        =   10
            Text            =   "12.0"
            Top             =   840
            Width           =   1095
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
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   12
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label lblTemp 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   2
            Left            =   3360
            TabIndex        =   11
            Top             =   945
            Width           =   465
         End
      End
      Begin Threed.SSPanel pnlBox 
         Height          =   4335
         Index           =   4
         Left            =   120
         TabIndex        =   736
         Top             =   5280
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
         _ExtentY        =   7646
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
         Begin VB.TextBox txtActName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   120
            TabIndex        =   738
            Text            =   "TEMP1"
            Top             =   660
            Width           =   2175
         End
         Begin VB.CheckBox chkAct02 
            Caption         =   "ACT DISPLAY 02"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   737
            Top             =   120
            Width           =   2235
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   4035
            Index           =   6
            Left            =   6060
            TabIndex        =   739
            Top             =   180
            Width           =   7335
            _Version        =   65536
            _ExtentX        =   12938
            _ExtentY        =   7117
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
            Begin VB.TextBox txtAct02TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   779
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct02TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   778
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct02CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   777
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct02VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   776
               Text            =   "0"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct02CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   775
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct02VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   774
               Text            =   "0.3"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct02VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   773
               Text            =   "0.45"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct02CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   772
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct02VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   771
               Text            =   "0.15"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct02CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   770
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct02TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   769
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct02TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   768
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct02Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   767
               Text            =   "ST1"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct02Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   766
               Text            =   "P1"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct02SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   765
               Text            =   "0.3"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct02SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   764
               Text            =   "0"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct02VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   763
               Text            =   "4.85"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct02CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   762
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct02VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   761
               Text            =   "4.55"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct02CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   760
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct02TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   759
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct02TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   758
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct02Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   757
               Text            =   "P2"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct02SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   756
               Text            =   "4.7"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct02VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   755
               Text            =   "5"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct02CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   754
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct02VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   753
               Text            =   "4.7"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct02CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   752
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct02TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   751
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct02TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   750
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct02Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   749
               Text            =   "ST2"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct02SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   748
               Text            =   "5"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct02SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   747
               Text            =   "0.3"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct02Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   746
               Text            =   "FINAL"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct02TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   745
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct02TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   744
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct02CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   743
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct02VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   742
               Text            =   "0.15"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct02CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   741
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct02VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   740
               Text            =   "0.45"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.Label lblActTestTypeName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "VOLT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   787
               Top             =   660
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MIN TIME"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   17
               Left            =   120
               TabIndex        =   786
               Top             =   3060
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MAX TIME"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   18
               Left            =   120
               TabIndex        =   785
               Top             =   3540
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MAX CURR"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   23
               Left            =   120
               TabIndex        =   784
               Top             =   1620
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MIN CURR"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   24
               Left            =   120
               TabIndex        =   783
               Top             =   1140
               Width           =   1335
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "NAME"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   782
               Top             =   180
               Width           =   1335
            End
            Begin VB.Label lblActTestTypeMinName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MIN VOLT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   781
               Top             =   2100
               Width           =   1335
            End
            Begin VB.Label lblActTestTypeMaxName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MAX VOLT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   780
               Top             =   2580
               Width           =   1335
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   615
            Index           =   14
            Left            =   120
            TabIndex        =   872
            Top             =   2940
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   1085
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
            Begin VB.OptionButton optAct02TestType 
               Caption         =   "STEPPING"
               Height          =   375
               Index           =   0
               Left            =   1980
               TabIndex        =   874
               Top             =   120
               Width           =   1635
            End
            Begin VB.OptionButton optAct02TestType 
               Caption         =   "FEEDBACK"
               Height          =   375
               Index           =   1
               Left            =   3840
               TabIndex        =   873
               Top             =   120
               Value           =   -1  'True
               Width           =   1635
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   375
               Index           =   7
               Left            =   120
               TabIndex        =   957
               Top             =   120
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   661
               _StockProps     =   15
               Caption         =   "ACT TYPE"
               BackColor       =   13160660
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "³ª´®°íµñ"
                  Size            =   11.99
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   555
            Index           =   22
            Left            =   120
            TabIndex        =   953
            Top             =   3660
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   979
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
            Begin VB.TextBox txtAct02StallDeltaMaxVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Left            =   4500
               TabIndex        =   959
               Text            =   "5"
               Top             =   60
               Width           =   975
            End
            Begin VB.TextBox txtAct02StallDeltaMinVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Left            =   2580
               TabIndex        =   958
               Text            =   "0"
               Top             =   60
               Width           =   975
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   4
               Left            =   120
               TabIndex        =   954
               Top             =   60
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "STALL ¥Ä"
               BackColor       =   13160660
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
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   5
               Left            =   1860
               TabIndex        =   955
               Top             =   60
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "MIN"
               BackColor       =   13160660
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
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   6
               Left            =   3780
               TabIndex        =   956
               Top             =   60
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "MAX"
               BackColor       =   13160660
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
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   1635
            Index           =   5
            Left            =   120
            TabIndex        =   1084
            Top             =   1200
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
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
            Begin VB.TextBox txtActPeakCurrCount 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   435
               Index           =   1
               Left            =   3540
               TabIndex        =   1086
               Text            =   "10"
               Top             =   300
               Width           =   1755
            End
            Begin VB.TextBox txtActEndPosCount 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   435
               Index           =   1
               Left            =   3540
               TabIndex        =   1085
               Text            =   "10"
               Top             =   900
               Width           =   1755
            End
            Begin VB.Label Label3 
               Caption         =   "PEAK CURRENT COUNT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   10
               Left            =   540
               TabIndex        =   1088
               Top             =   360
               Width           =   2775
            End
            Begin VB.Label Label3 
               Caption         =   "END POSTION COUNT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   9
               Left            =   540
               TabIndex        =   1087
               Top             =   960
               Width           =   2775
            End
         End
      End
      Begin Threed.SSPanel pnlBox 
         Height          =   4335
         Index           =   0
         Left            =   120
         TabIndex        =   788
         Top             =   840
         Width           =   16935
         _Version        =   65536
         _ExtentX        =   29871
         _ExtentY        =   7646
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
         Begin VB.CheckBox chkAct01 
            Caption         =   "ACT DISPLAY 01"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   790
            Top             =   120
            Width           =   2235
         End
         Begin VB.TextBox txtActName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   120
            TabIndex        =   789
            Text            =   "MODE"
            Top             =   660
            Width           =   2175
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   1635
            Index           =   2
            Left            =   120
            TabIndex        =   791
            Top             =   1200
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
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
            Begin VB.TextBox txtActEndPosCount 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   435
               Index           =   0
               Left            =   3540
               TabIndex        =   793
               Text            =   "10"
               Top             =   900
               Width           =   1755
            End
            Begin VB.TextBox txtActPeakCurrCount 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   435
               Index           =   0
               Left            =   3540
               TabIndex        =   792
               Text            =   "10"
               Top             =   300
               Width           =   1755
            End
            Begin VB.Label Label3 
               Caption         =   "END POSTION COUNT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   12
               Left            =   540
               TabIndex        =   795
               Top             =   960
               Width           =   2775
            End
            Begin VB.Label Label3 
               Caption         =   "PEAK CURRENT COUNT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   11
               Left            =   540
               TabIndex        =   794
               Top             =   360
               Width           =   2775
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   4035
            Index           =   3
            Left            =   6060
            TabIndex        =   796
            Top             =   180
            Width           =   10755
            _Version        =   65536
            _ExtentX        =   18971
            _ExtentY        =   7117
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
            Begin VB.TextBox txtAct01TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   7
               Left            =   9540
               TabIndex        =   860
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   7
               Left            =   9540
               TabIndex        =   859
               Text            =   "0.15"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   7
               Left            =   9540
               TabIndex        =   858
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   7
               Left            =   9540
               TabIndex        =   857
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   7
               Left            =   9540
               TabIndex        =   856
               Text            =   "0.45"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   7
               Left            =   9540
               TabIndex        =   855
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct01Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   7
               Left            =   9540
               TabIndex        =   854
               Text            =   "FINAL"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct01SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   7
               Left            =   9540
               TabIndex        =   853
               Text            =   "0.3"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   6
               Left            =   8400
               TabIndex        =   852
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   6
               Left            =   8400
               TabIndex        =   851
               Text            =   "4.7"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   6
               Left            =   8400
               TabIndex        =   850
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   6
               Left            =   8400
               TabIndex        =   849
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   6
               Left            =   8400
               TabIndex        =   848
               Text            =   "5"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   6
               Left            =   8400
               TabIndex        =   847
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct01Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   6
               Left            =   8400
               TabIndex        =   846
               Text            =   "ST2"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct01SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   6
               Left            =   8400
               TabIndex        =   845
               Text            =   "5"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   5
               Left            =   7260
               TabIndex        =   844
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   5
               Left            =   7260
               TabIndex        =   843
               Text            =   "4.55"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   5
               Left            =   7260
               TabIndex        =   842
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   5
               Left            =   7260
               TabIndex        =   841
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   5
               Left            =   7260
               TabIndex        =   840
               Text            =   "4.85"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   5
               Left            =   7260
               TabIndex        =   839
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct01Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   5
               Left            =   7260
               TabIndex        =   838
               Text            =   "P5"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct01SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   5
               Left            =   7260
               TabIndex        =   837
               Text            =   "4.7"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   836
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   835
               Text            =   "0"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   834
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   833
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   832
               Text            =   "0.3"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   831
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct01Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   830
               Text            =   "ST1"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct01SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   829
               Text            =   "0"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   828
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   827
               Text            =   "0.15"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   826
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   825
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   824
               Text            =   "0.45"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   823
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct01Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   822
               Text            =   "P1"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct01SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   821
               Text            =   "0.3"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   820
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   819
               Text            =   "1.65"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   818
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   817
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   816
               Text            =   "1.95"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   815
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct01Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   814
               Text            =   "P2"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct01SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   813
               Text            =   "1.8"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   812
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   811
               Text            =   "2.35"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   810
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   809
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   808
               Text            =   "2.65"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   807
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct01Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   806
               Text            =   "P3"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct01SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   805
               Text            =   "2.5"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   804
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   803
               Text            =   "3.65"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   802
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct01CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   801
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct01VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   800
               Text            =   "3.95"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct01TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   799
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct01Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   798
               Text            =   "P4"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct01SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   797
               Text            =   "3.8"
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label lblActTestTypeName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "VOLT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   868
               Top             =   660
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MIN TIME"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   8
               Left            =   120
               TabIndex        =   867
               Top             =   3060
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MAX TIME"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   7
               Left            =   120
               TabIndex        =   866
               Top             =   3540
               Width           =   1335
            End
            Begin VB.Label lblActTestTypeMinName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MIN VOLT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   865
               Top             =   2100
               Width           =   1335
            End
            Begin VB.Label lblActTestTypeMaxName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MAX VOLT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   864
               Top             =   2580
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MAX CURR"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   863
               Top             =   1620
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MIN CURR"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   862
               Top             =   1140
               Width           =   1335
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "NAME"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   861
               Top             =   180
               Width           =   1335
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   615
            Index           =   13
            Left            =   120
            TabIndex        =   869
            Top             =   2940
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   1085
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
            Begin VB.OptionButton optAct01TestType 
               Caption         =   "FEEDBACK"
               Height          =   375
               Index           =   1
               Left            =   3840
               TabIndex        =   871
               Top             =   120
               Value           =   -1  'True
               Width           =   1635
            End
            Begin VB.OptionButton optAct01TestType 
               Caption         =   "STEPPING"
               Height          =   375
               Index           =   0
               Left            =   1980
               TabIndex        =   870
               Top             =   120
               Width           =   1635
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   948
               Top             =   120
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   661
               _StockProps     =   15
               Caption         =   "ACT TYPE"
               BackColor       =   13160660
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "³ª´®°íµñ"
                  Size            =   11.99
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   555
            Index           =   21
            Left            =   120
            TabIndex        =   946
            Top             =   3660
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   979
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
            Begin VB.TextBox txtAct01StallDeltaMaxVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Left            =   4500
               TabIndex        =   952
               Text            =   "5"
               Top             =   60
               Width           =   975
            End
            Begin VB.TextBox txtAct01StallDeltaMinVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Left            =   2580
               TabIndex        =   947
               Text            =   "0"
               Top             =   60
               Width           =   975
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   1
               Left            =   120
               TabIndex        =   949
               Top             =   60
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "STALL ¥Ä"
               BackColor       =   13160660
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
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   2
               Left            =   1860
               TabIndex        =   950
               Top             =   60
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "MIN"
               BackColor       =   13160660
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
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   3
               Left            =   3780
               TabIndex        =   951
               Top             =   60
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "MAX"
               BackColor       =   13160660
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
            End
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   11295
         Index           =   1
         Left            =   -74880
         TabIndex        =   960
         Top             =   840
         Width           =   18615
         _Version        =   65536
         _ExtentX        =   32835
         _ExtentY        =   19923
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
            Left            =   9660
            TabIndex        =   963
            Text            =   "0.00"
            Top             =   1080
            Visible         =   0   'False
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
            Left            =   11280
            TabIndex        =   962
            Text            =   "1.00"
            Top             =   1080
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CheckBox chkAdjustUse 
            Caption         =   "CH 00"
            Height          =   285
            Index           =   0
            Left            =   4920
            TabIndex        =   961
            Top             =   1140
            Visible         =   0   'False
            Width           =   1095
         End
         Begin Threed.SSPanel pnlAdChName 
            Height          =   435
            Index           =   0
            Left            =   6240
            TabIndex        =   964
            Top             =   1080
            Visible         =   0   'False
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
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
         Begin VB.Label lblTitleTemp 
            Alignment       =   2  'Center
            Caption         =   "CHANNEL"
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
            Left            =   6480
            TabIndex        =   967
            Top             =   540
            Width           =   1815
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
            Height          =   315
            Index           =   1
            Left            =   9660
            TabIndex        =   966
            Top             =   540
            Width           =   1335
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
            Height          =   315
            Index           =   0
            Left            =   11280
            TabIndex        =   965
            Top             =   540
            Width           =   1395
         End
      End
      Begin Threed.SSPanel pnlBox 
         Height          =   4335
         Index           =   7
         Left            =   -74880
         TabIndex        =   968
         Top             =   840
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
         _ExtentY        =   7646
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
         Begin VB.TextBox txtActName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   2
            Left            =   120
            TabIndex        =   970
            Text            =   "TEMP2"
            Top             =   660
            Width           =   2175
         End
         Begin VB.CheckBox chkAct03 
            Caption         =   "ACT DISPLAY 03"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   969
            Top             =   120
            Width           =   2235
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   4035
            Index           =   9
            Left            =   6060
            TabIndex        =   971
            Top             =   180
            Width           =   7335
            _Version        =   65536
            _ExtentX        =   12938
            _ExtentY        =   7117
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
            Begin VB.TextBox txtAct03SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1011
               Text            =   "0.3"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct03Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1010
               Text            =   "FINAL"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct03TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1009
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct03TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1008
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct03CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1007
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct03VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1006
               Text            =   "0.15"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct03CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1005
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct03VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1004
               Text            =   "0.45"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct03TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   1003
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct03TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   1002
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct03CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   1001
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct03VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   1000
               Text            =   "0"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct03CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   999
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct03VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   998
               Text            =   "0.3"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct03VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   997
               Text            =   "0.45"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct03CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   996
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct03VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   995
               Text            =   "0.15"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct03CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   994
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct03TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   993
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct03TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   992
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct03Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   991
               Text            =   "ST1"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct03Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   990
               Text            =   "P1"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct03SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   989
               Text            =   "0.3"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct03SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   988
               Text            =   "0"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct03VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   987
               Text            =   "4.85"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct03CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   986
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct03VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   985
               Text            =   "4.55"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct03CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   984
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct03TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   983
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct03TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   982
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct03Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   981
               Text            =   "P2"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct03SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   980
               Text            =   "4.7"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct03VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   979
               Text            =   "5"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct03CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   978
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct03VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   977
               Text            =   "4.7"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct03CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   976
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct03TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   975
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct03TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   974
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct03Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   973
               Text            =   "ST2"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct03SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   972
               Text            =   "5"
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label lblActTestTypeMaxName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MAX VOLT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   1019
               Top             =   2580
               Width           =   1335
            End
            Begin VB.Label lblActTestTypeMinName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MIN VOLT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   1018
               Top             =   2100
               Width           =   1335
            End
            Begin VB.Label lblActTestTypeName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "VOLT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   1017
               Top             =   660
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MIN TIME"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   37
               Left            =   120
               TabIndex        =   1016
               Top             =   3060
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MAX TIME"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   36
               Left            =   120
               TabIndex        =   1015
               Top             =   3540
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MAX CURR"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   33
               Left            =   120
               TabIndex        =   1014
               Top             =   1620
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MIN CURR"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   31
               Left            =   120
               TabIndex        =   1013
               Top             =   1140
               Width           =   1335
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "NAME"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   1012
               Top             =   180
               Width           =   1335
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   615
            Index           =   15
            Left            =   120
            TabIndex        =   1020
            Top             =   2940
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   1085
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
            Begin VB.OptionButton optAct03TestType 
               Caption         =   "STEPPING"
               Height          =   375
               Index           =   0
               Left            =   1980
               TabIndex        =   1022
               Top             =   120
               Width           =   1635
            End
            Begin VB.OptionButton optAct03TestType 
               Caption         =   "FEEDBACK"
               Height          =   375
               Index           =   1
               Left            =   3840
               TabIndex        =   1021
               Top             =   120
               Value           =   -1  'True
               Width           =   1635
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   375
               Index           =   8
               Left            =   120
               TabIndex        =   1023
               Top             =   120
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   661
               _StockProps     =   15
               Caption         =   "ACT TYPE"
               BackColor       =   13160660
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "³ª´®°íµñ"
                  Size            =   11.99
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   1635
            Index           =   8
            Left            =   120
            TabIndex        =   1089
            Top             =   1200
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
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
            Begin VB.TextBox txtActPeakCurrCount 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   435
               Index           =   2
               Left            =   3540
               TabIndex        =   1091
               Text            =   "10"
               Top             =   300
               Width           =   1755
            End
            Begin VB.TextBox txtActEndPosCount 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   435
               Index           =   2
               Left            =   3540
               TabIndex        =   1090
               Text            =   "10"
               Top             =   900
               Width           =   1755
            End
            Begin VB.Label Label3 
               Caption         =   "PEAK CURRENT COUNT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   4
               Left            =   540
               TabIndex        =   1093
               Top             =   360
               Width           =   2775
            End
            Begin VB.Label Label3 
               Caption         =   "END POSTION COUNT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   2
               Left            =   540
               TabIndex        =   1092
               Top             =   960
               Width           =   2775
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   555
            Index           =   23
            Left            =   120
            TabIndex        =   1099
            Top             =   3660
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   979
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
            Begin VB.TextBox txtAct03StallDeltaMaxVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Left            =   4500
               TabIndex        =   1104
               Text            =   "5"
               Top             =   60
               Width           =   975
            End
            Begin VB.TextBox txtAct03StallDeltaMinVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Left            =   2580
               TabIndex        =   1103
               Text            =   "0"
               Top             =   60
               Width           =   975
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   9
               Left            =   120
               TabIndex        =   1100
               Top             =   60
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "STALL ¥Ä"
               BackColor       =   13160660
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
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   10
               Left            =   1860
               TabIndex        =   1101
               Top             =   60
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "MIN"
               BackColor       =   13160660
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
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   11
               Left            =   3780
               TabIndex        =   1102
               Top             =   60
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "MAX"
               BackColor       =   13160660
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
            End
         End
      End
      Begin Threed.SSPanel pnlBox 
         Height          =   4335
         Index           =   10
         Left            =   -74880
         TabIndex        =   1024
         Top             =   5280
         Width           =   15315
         _Version        =   65536
         _ExtentX        =   27014
         _ExtentY        =   7646
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
         Begin VB.CheckBox chkAct04 
            Caption         =   "ACT DISPLAY 04"
            BeginProperty Font 
               Name            =   "³ª´®°íµñ"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   1029
            Top             =   120
            Width           =   2235
         End
         Begin VB.TextBox txtActName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   3
            Left            =   120
            TabIndex        =   1028
            Text            =   "REC"
            Top             =   660
            Width           =   2175
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   1935
            Index           =   1
            Left            =   13500
            TabIndex        =   1025
            Top             =   180
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
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
            Begin VB.OptionButton opt2Pin 
               Caption         =   "5 PIN"
               Height          =   375
               Index           =   0
               Left            =   360
               TabIndex        =   1027
               Top             =   660
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton opt2Pin 
               Caption         =   "2 PIN"
               Height          =   375
               Index           =   1
               Left            =   360
               TabIndex        =   1026
               Top             =   1200
               Width           =   975
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   16
               Left            =   120
               TabIndex        =   1112
               Top             =   60
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "PIN TYPE"
               BackColor       =   13160660
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
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   4035
            Index           =   12
            Left            =   6060
            TabIndex        =   1030
            Top             =   180
            Width           =   7335
            _Version        =   65536
            _ExtentX        =   12938
            _ExtentY        =   7117
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
            Begin VB.TextBox txtAct04VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1072
               Text            =   "0.45"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct04VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1071
               Text            =   "0.15"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.OptionButton opt2PinPos 
               Caption         =   "P1"
               Height          =   495
               Index           =   0
               Left            =   6120
               TabIndex        =   1070
               Top             =   2040
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton opt2PinPos 
               Caption         =   "P2"
               Height          =   435
               Index           =   1
               Left            =   6120
               TabIndex        =   1069
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct04SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1068
               Text            =   "0.3"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct04Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1067
               Text            =   "FINAL"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct04TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1066
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct04TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1065
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct04CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1064
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct04CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   4
               Left            =   6120
               TabIndex        =   1063
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct04TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   1062
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct04TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   1061
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct04CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   1060
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct04VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   1059
               Text            =   "0"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct04CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   1058
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct04VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   1057
               Text            =   "0.3"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct04VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   1056
               Text            =   "0.45"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct04CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   1055
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct04VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   1054
               Text            =   "0.15"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct04CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   1053
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct04TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   1052
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct04TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   1051
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct04Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   1050
               Text            =   "ST1"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct04Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   1049
               Text            =   "P1"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct04SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   1
               Left            =   2700
               TabIndex        =   1048
               Text            =   "0.3"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct04SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   1560
               TabIndex        =   1047
               Text            =   "0"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct04VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   1046
               Text            =   "4.85"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct04CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   1045
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct04VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   1044
               Text            =   "4.55"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct04CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   1043
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct04TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   1042
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct04TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   1041
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct04Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   1040
               Text            =   "P2"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct04SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   2
               Left            =   3840
               TabIndex        =   1039
               Text            =   "4.7"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtAct04VoltHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   1038
               Text            =   "5"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtAct04CurrHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   1037
               Text            =   "80"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtAct04VoltLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   1036
               Text            =   "4.7"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtAct04CurrLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   1035
               Text            =   "5"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtAct04TimeLo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   1034
               Text            =   "0.1"
               Top             =   3000
               Width           =   1095
            End
            Begin VB.TextBox txtAct04TimeHi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   1033
               Text            =   "10.0"
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtAct04Name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   1032
               Text            =   "ST2"
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtAct04SetVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   3
               Left            =   4980
               TabIndex        =   1031
               Text            =   "5"
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label lblActTestTypeMaxName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MAX VOLT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   1080
               Top             =   2580
               Width           =   1335
            End
            Begin VB.Label lblActTestTypeMinName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MIN VOLT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   1079
               Top             =   2100
               Width           =   1335
            End
            Begin VB.Label lblActTestTypeName 
               Alignment       =   1  'Right Justify
               Caption         =   "VOLT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   1078
               Top             =   660
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "MIN TIME"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   39
               Left            =   120
               TabIndex        =   1077
               Top             =   3060
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "MAX TIME"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   40
               Left            =   120
               TabIndex        =   1076
               Top             =   3540
               Width           =   1335
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               Caption         =   "MAX CURR"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   1075
               Top             =   1620
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "MIN CURR"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   41
               Left            =   120
               TabIndex        =   1074
               Top             =   1140
               Width           =   1335
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "NAME"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   1073
               Top             =   180
               Width           =   1335
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   615
            Index           =   16
            Left            =   120
            TabIndex        =   1081
            Top             =   2940
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   1085
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
            Begin VB.OptionButton optAct04TestType 
               Caption         =   "STEPPING"
               Height          =   375
               Index           =   0
               Left            =   1980
               TabIndex        =   1083
               Top             =   120
               Width           =   1635
            End
            Begin VB.OptionButton optAct04TestType 
               Caption         =   "FEEDBACK"
               Height          =   375
               Index           =   1
               Left            =   3840
               TabIndex        =   1082
               Top             =   120
               Value           =   -1  'True
               Width           =   1635
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   375
               Index           =   15
               Left            =   120
               TabIndex        =   1109
               Top             =   120
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   661
               _StockProps     =   15
               Caption         =   "ACT TYPE"
               BackColor       =   13160660
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "³ª´®°íµñ"
                  Size            =   11.99
                  Charset         =   129
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   1635
            Index           =   11
            Left            =   120
            TabIndex        =   1094
            Top             =   1200
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
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
            Begin VB.TextBox txtActEndPosCount 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   435
               Index           =   3
               Left            =   3540
               TabIndex        =   1096
               Text            =   "10"
               Top             =   900
               Width           =   1755
            End
            Begin VB.TextBox txtActPeakCurrCount 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   435
               Index           =   3
               Left            =   3540
               TabIndex        =   1095
               Text            =   "10"
               Top             =   300
               Width           =   1755
            End
            Begin VB.Label Label3 
               Caption         =   "END POSTION COUNT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   13
               Left            =   540
               TabIndex        =   1098
               Top             =   960
               Width           =   2775
            End
            Begin VB.Label Label3 
               Caption         =   "PEAK CURRENT COUNT"
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   5
               Left            =   540
               TabIndex        =   1097
               Top             =   360
               Width           =   2775
            End
         End
         Begin Threed.SSPanel pnlBox 
            Height          =   555
            Index           =   24
            Left            =   120
            TabIndex        =   1105
            Top             =   3660
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   979
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
            Begin VB.TextBox txtAct04StallDeltaMaxVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Left            =   4500
               TabIndex        =   1111
               Text            =   "5"
               Top             =   60
               Width           =   975
            End
            Begin VB.TextBox txtAct04StallDeltaMinVolt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H00000000&
               Height          =   435
               Left            =   2580
               TabIndex        =   1110
               Text            =   "0"
               Top             =   60
               Width           =   975
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   12
               Left            =   120
               TabIndex        =   1106
               Top             =   60
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "STALL ¥Ä"
               BackColor       =   13160660
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
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   13
               Left            =   1860
               TabIndex        =   1107
               Top             =   60
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "MIN"
               BackColor       =   13160660
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
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   435
               Index           =   14
               Left            =   3780
               TabIndex        =   1108
               Top             =   60
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "MAX"
               BackColor       =   13160660
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
            End
         End
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim bLoading As Boolean

Private Sub Form_Activate()
    nNowForm = FM_SETUP
    
    Call LoadLangFile(FM_SETUP)
    Call opt2Pin_Click(SetupVar.nAct042Pin)
    Call optBlowerType_Click(SetupVar.nBlowerType)
    Call optAct01TestType_Click(SetupVar.nAct01TestType)
    Call optAct02TestType_Click(SetupVar.nAct02TestType)
    Call optAct03TestType_Click(SetupVar.nAct03TestType)
    Call optAct04TestType_Click(SetupVar.nAct04TestType)
    
    If bLoading = False Then
        Call OnStart
        
        bLoading = True
    End If
End Sub

Private Sub Form_Load()
    bLoading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call OnEnd
End Sub

Private Sub btnLoading_Click()
    If bLoading = True Then
        Call LoadSetupFile(Trim$(cboLoad.Text))
        Call SetupMem2Disp
    End If
End Sub

Private Sub btnReturn_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    Call OnSave
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    Dim lpStr As String
    
    On Error Resume Next
    
    If cboCarType.ListIndex = -1 Then
        For i = 0 To (cboCarType.ListCount - 1)
            If UCase(Trim$(cboCarType.List(i))) = UCase(Trim$(cboCarType.Text)) Then
                Exit For
            End If
        Next
        If i = cboCarType.ListCount Then
            Call MsgBox("Model not found.", vbInformation, "Checking")
            Exit Sub
        End If
    Else
        i = cboCarType.ListIndex
    End If
    lpStr = Trim$(cboCarType.List(i))
    
    If Trim$(lpNowModel) = Trim$(cboCarType.Text) Then
        Call MsgBox("Now model name is " + lpStr + ".", vbCritical, "Warning")
        Exit Sub
    End If
    
    If MsgBox(lpStr + " Delete." + vbCrLf + "Do you want to delete ? ", vbYesNo + vbQuestion + vbApplicationModal, "Checking") = vbYes Then
        Call INIModelDelete(Trim$(cboCarType.Text))
        
        cboCarType.RemoveItem i
        cboCarType.ListIndex = 0
        
        Call MsgBox("Deleted.", vbOKOnly + vbExclamation, "chkLinActUse")
        Call LoadSetupFile(Trim$(cboCarType.List(cboCarType.ListIndex)))
        Call SetupMem2Disp
    End If
End Sub

Private Sub cboCarType_Click()
    If bLoading = True Then
        Call LoadSetupFile(Trim$(cboCarType.List(cboCarType.ListIndex)))
        Call PartVisible(False)
        Call SetupMem2Disp
    End If
End Sub

Private Sub lblPart_Click(Index As Integer)
    If chkPart(Index).Value = 1 Then
        If lblPart(Index).BackColor = vbRed Then
            lblPart(Index).BackColor = vbGreen
            lblPart(Index).Caption = "ON"
        Else
            lblPart(Index).BackColor = vbRed
            lblPart(Index).Caption = "OFF"
        End If
    End If
End Sub

Private Sub chkPart_Click(Index As Integer)
    If chkPart(Index).Value = 1 Then
        lblPart(Index).BackColor = vbRed
        lblPart(Index).Caption = "OFF"
    Else
        lblPart(Index).BackColor = CO_NONE
        lblPart(Index).Caption = "--"
    End If
End Sub

Private Sub opt2Pin_Click(Index As Integer)
    Dim i As Integer
    Dim bRes As Boolean
    
    Select Case Index
        Case 0: bRes = True
        Case 1: bRes = False
    End Select
    
'    frmSetup.fraFinalPos.Caption = IIf(bRes, "5 PIN FINAL POSITION", "2 PIN FINAL POSITION")
    
    For i = 0 To 1
        frmSetup.opt2PinPos(i).Visible = Not bRes
    Next
    
    frmSetup.lblActTestTypeName(3).Visible = bRes
    frmSetup.lblActTestTypeMinName(3).Visible = bRes
    frmSetup.lblActTestTypeMaxName(3).Visible = bRes
    
    For i = 0 To frmSetup.txtAct04Name.UBound
        frmSetup.txtAct04SetVolt(i).Visible = bRes
        frmSetup.txtAct04VoltHi(i).Visible = bRes
        frmSetup.txtAct04VoltLo(i).Visible = bRes
    Next
End Sub

Private Sub optAct01TestType_Click(Index As Integer)
    Select Case Index
        Case 0:
            frmSetup.lblActTestTypeName(0).Caption = "STEP"
            frmSetup.lblActTestTypeMinName(0).Caption = "MIN STEP"
            frmSetup.lblActTestTypeMaxName(0).Caption = "MAX STEP"
        Case 1:
            frmSetup.lblActTestTypeName(0).Caption = "VOLT"
            frmSetup.lblActTestTypeMinName(0).Caption = "MIN VOLT"
            frmSetup.lblActTestTypeMaxName(0).Caption = "MAX VOLT"
    End Select
End Sub

Private Sub optAct02TestType_Click(Index As Integer)
    Select Case Index
        Case 0:
            frmSetup.lblActTestTypeName(1).Caption = "STEP"
            frmSetup.lblActTestTypeMinName(1).Caption = "MIN STEP"
            frmSetup.lblActTestTypeMaxName(1).Caption = "MAX STEP"
        Case 1:
            frmSetup.lblActTestTypeName(1).Caption = "VOLT"
            frmSetup.lblActTestTypeMinName(1).Caption = "MIN VOLT"
            frmSetup.lblActTestTypeMaxName(1).Caption = "MAX VOLT"
    End Select
End Sub

Private Sub optAct03TestType_Click(Index As Integer)
    Select Case Index
        Case 0:
            frmSetup.lblActTestTypeName(2).Caption = "STEP"
            frmSetup.lblActTestTypeMinName(2).Caption = "MIN STEP"
            frmSetup.lblActTestTypeMaxName(2).Caption = "MAX STEP"
        Case 1:
            frmSetup.lblActTestTypeName(2).Caption = "VOLT"
            frmSetup.lblActTestTypeMinName(2).Caption = "MIN VOLT"
            frmSetup.lblActTestTypeMaxName(2).Caption = "MAX VOLT"
    End Select
End Sub

Private Sub optAct04TestType_Click(Index As Integer)
    Select Case Index
        Case 0:
            frmSetup.lblActTestTypeName(3).Caption = "STEP"
            frmSetup.lblActTestTypeMinName(3).Caption = "MIN STEP"
            frmSetup.lblActTestTypeMaxName(3).Caption = "MAX STEP"
        Case 1:
            frmSetup.lblActTestTypeName(3).Caption = "VOLT"
            frmSetup.lblActTestTypeMinName(3).Caption = "MIN VOLT"
            frmSetup.lblActTestTypeMaxName(3).Caption = "MAX VOLT"
    End Select
End Sub

Private Sub optBlowerType_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To frmSetup.txtBlowerName.UBound
        frmSetup.txtLinSpeed(i).BackColor = IIf(Index < 2, CO_NONE, &H80C0FF)
        frmSetup.txtLinSpeed(i).Enabled = IIf(Index < 2, False, True)
    Next
End Sub

Private Sub OnStart()
    Dim bRes As Boolean
    Dim nRes As Integer
    
    SSTab.Tab = 0
    lpPath = App.Path
    
    Call SetupChCtlArray
    
    bRes = ModelChange(cboLoad, Trim$(lpNowModel))
    bRes = ModelChange(cboCarType, Trim$(lpNowModel))
    
    If bRes Then
        Call SetupMem2Disp
    Else
        If SETUPSELECT Then
            cboCarType.Text = Format(nNowModelNo, "0000") & "_" & SelectCar(0).ModelName & "_" & SelectCar(0).ModelNameSub(0)
        Else
            cboCarType.Text = "0000_DEFAULT"
        End If
    End If
    
    frmSetup.txtProductList.Visible = DEBUGMODE
    frmSetup.chkModelTypeUse.Visible = DEBUGMODE
    frmSetup.txtModelLHDList.Visible = DEBUGMODE
    frmSetup.txtModelRHDList.Visible = DEBUGMODE
    
    Select Case SetupVar.nModelType
        Case 0:
            If SetupVar.lpModelLHDList <> "" Then
                If Left$(SetupVar.lpModelLHDList, 1) = "#" Then
                    nRes = Val(Right(SetupVar.lpModelLHDList, 2))
                Else
                    nRes = Val(SetupVar.lpModelLHDList)
                End If
                
                frmSetup.chkModelTypeUse.Caption = frmSetup.chkModelTypeUse.Caption & " " & "(X" & Format(nRes, "00") & ")"
            Else
                frmSetup.chkModelTypeUse.Caption = ""
            End If
        
        Case 1:
            If SetupVar.lpModelRHDList <> "" Then
                If Left$(SetupVar.lpModelRHDList, 1) = "#" Then
                    nRes = Val(Right(SetupVar.lpModelRHDList, 2))
                Else
                    nRes = Val(SetupVar.lpModelRHDList)
                End If
                
                frmSetup.chkModelTypeUse.Caption = frmSetup.chkModelTypeUse.Caption & " " & "(X" & Format(nRes, "00") & ")"
            Else
                frmSetup.chkModelTypeUse.Caption = ""
            End If
    
    End Select
    
    If SetupVar.lpProductList <> "" Then
        If Left$(SetupVar.lpProductList, 1) = "#" Then
            nRes = Val(Right(SetupVar.lpProductList, 2))
        Else
            nRes = Val(SetupVar.lpProductList)
        End If
        
        frmSetup.chkProductUse.Caption = frmSetup.chkProductUse.Caption & " " & "(X" & Format(nRes, "00") & ")"
    End If
    
    Call SetupPartVisible
    Call OnLog("Setup File Load.")
End Sub

Private Sub OnEnd()
    Call OnLog("Setup End")
End Sub

Private Sub OnSave()
    Call SetupDisp2Mem
    Call SaveSetupFile(Trim$(cboCarType.Text))
    Call ModelChange(cboCarType, Trim$(cboCarType.Text))
    Call OnLog("Setup File Save")
End Sub

Private Sub SetupPartVisible()
    Dim i As Integer
    Dim lpStr() As String
    Dim lpRes As String
    
    ' Product check pass
    lpStr = Split(SetupVar.lpProductList, ",")
    
    For i = LBound(lpStr) To UBound(lpStr)
        frmSetup.chkPart(Val(lpStr(i))).Visible = False
        frmSetup.lblPart(Val(lpStr(i))).Visible = False
        frmSetup.txtPart(Val(lpStr(i))).Visible = False
    Next
    
    ' Model check pass
    lpRes = SetupVar.lpModelLHDList & "," & SetupVar.lpModelRHDList
    lpStr = Split(lpRes, ",")
    
    For i = LBound(lpStr) To UBound(lpStr)
        frmSetup.chkPart(Val(lpStr(i))).Visible = False
        frmSetup.lblPart(Val(lpStr(i))).Visible = False
        frmSetup.txtPart(Val(lpStr(i))).Visible = False
    Next
    
    For i = 0 To MAX_DIO_CHANNEL
        Select Case i
            Case I_START_SW, I_STOP_SW, I_AUTO_SW:
                frmSetup.chkPart(i).Visible = False
                frmSetup.lblPart(i).Visible = False
                frmSetup.txtPart(i).Visible = False
            
            Case I_WORK_ON, I_WORK_OFF, I_MARKING1_ON, I_MARKING1_OFF:
                frmSetup.chkPart(i).Visible = False
                frmSetup.lblPart(i).Visible = False
                frmSetup.txtPart(i).Visible = False
        
        End Select
        
        If i Mod 10 = 8 Or i Mod 10 = 9 Then
            frmSetup.chkPart(i).Visible = False
            frmSetup.lblPart(i).Visible = False
            frmSetup.txtPart(i).Visible = False
        End If
    Next
End Sub

Private Sub SetupChCtlArray()
    Dim i As Integer
    Dim Index As Integer
    
    frmSetup.pnlAdChName(0).Outline = False
    
    ' ¹è¿­ 0 ÄÁÆ®·Ñ Á¦¿Ü
    For i = 1 To MAX_AD_CHANNEL
        Index = frmSetup.chkAdjustUse.Count
        
        Call Load(frmSetup.chkAdjustUse(Index))
        Call Load(frmSetup.pnlAdChName(Index))
        Call Load(frmSetup.txtAdd(Index))
        Call Load(frmSetup.txtMulti(Index))
        
        frmSetup.chkAdjustUse(Index).Top = frmSetup.chkAdjustUse(Index - 1).Top + 480
        frmSetup.pnlAdChName(Index).Top = frmSetup.pnlAdChName(Index - 1).Top + 480
        frmSetup.txtAdd(Index).Top = frmSetup.txtAdd(Index - 1).Top + 480
        frmSetup.txtMulti(Index).Top = frmSetup.txtMulti(Index - 1).Top + 480
        
        frmSetup.chkAdjustUse(Index).Caption = "CH " & Format(Index, "00")
        
        frmSetup.chkAdjustUse(Index).Visible = True
        frmSetup.pnlAdChName(Index).Visible = True
        frmSetup.txtAdd(Index).Visible = True
        frmSetup.txtMulti(Index).Visible = True
    Next
End Sub
