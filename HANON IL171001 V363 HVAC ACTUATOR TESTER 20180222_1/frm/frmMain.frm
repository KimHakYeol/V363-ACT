VERSION 5.00
Object = "{F36F6844-D389-11D1-8968-006097AA579E}#1.0#0"; "HiResTimer.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{A306B168-AE98-11D3-83AE-00A024BDBF2B}#3.0#0"; "ActEther.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MAIN"
   ClientHeight    =   15000
   ClientLeft      =   150
   ClientTop       =   -915
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
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   15000
   ScaleWidth      =   19110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnMenu 
      BackColor       =   &H00F0F0F0&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   36
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Index           =   5
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   11460
      Width           =   6555
   End
   Begin VB.CommandButton btnMenu 
      BackColor       =   &H00F0F0F0&
      Caption         =   "DATABASE"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   36
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Index           =   4
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9480
      Width           =   6555
   End
   Begin VB.CommandButton btnMenu 
      BackColor       =   &H00F0F0F0&
      Caption         =   "SYSTEM SET"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   36
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Index           =   3
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   6555
   End
   Begin VB.CommandButton btnMenu 
      BackColor       =   &H00F0F0F0&
      Caption         =   "I/O TEST"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   36
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Index           =   2
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5580
      Width           =   6555
   End
   Begin VB.CommandButton btnMenu 
      BackColor       =   &H00F0F0F0&
      Caption         =   "SETUP"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   36
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Index           =   1
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   6555
   End
   Begin VB.CommandButton btnMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      Caption         =   "AUTO / MANUAL"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   36
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Index           =   0
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1620
      Width           =   6555
   End
   Begin MSCommLib.MSComm Power 
      Left            =   660
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   6
      RThreshold      =   1
      RTSEnable       =   -1  'True
      EOFEnable       =   -1  'True
   End
   Begin HIRESTIMERLib.HiResTimer tmrPower 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   64
      Enabled         =   0   'False
      Interval        =   1
   End
   Begin MSCommLib.MSComm comNvh 
      Left            =   1260
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   19200
      SThreshold      =   1
   End
   Begin ACTETHERLibCtl.ActQJ71E71UDP ActPlc 
      Left            =   120
      OleObjectBlob   =   "frmMain.frx":9BA42
      Top             =   900
   End
   Begin MSCommLib.MSComm comStepping1 
      Left            =   2520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   19200
      SThreshold      =   1
      EOFEnable       =   -1  'True
      InputMode       =   1
   End
   Begin MSCommLib.MSComm comStepping2 
      Left            =   3180
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   19200
      SThreshold      =   1
      EOFEnable       =   -1  'True
      InputMode       =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim bLoading As Boolean

Private Sub Form_Activate()
    nNowForm = FM_MAIN
    
    Call LoadLangFile(FM_MAIN)
    
    ' 15360 X 19200
    If bLoading = False Then
        bLoading = True
        
        If bRunningGraphWin = False Then Call OnStart
    End If
End Sub

Private Sub Form_Load()
    bLoading = False
End Sub

Private Sub btnMenu_Click(Index As Integer)
    If Index = 1 Then Index = 10
    
    Call MenuSelect(Index)
End Sub

' Function List ===============================================================

Private Sub OnStart()
    Dim bRes As Boolean
    Dim lpRes As String
    
    nNowTime = Format(time, "HHMM")
    lpPath = App.Path
    
    Call MakeFolder
    
    bRes = LoadModelName
    bRes = LoadPlcFile
    bRes = LoadSystemFile
    
    If bRes = False Then
        Call MenuSelect(FM_SPLASH)
        
        bLogin = True
        Call MenuSelect(FM_SYSTEM)
        
        bLogin = True
        Call MenuSelect(FM_SETUP)
        End
    Else
        If Len(Trim$(SysVar.lpModel)) = 0 Then
            If SETUPSELECT Then
                lpRes = SelectCar(0).ModelNameSub(0)
                
                If SelectCar(0).ModelName <> "" And lpRes <> "" Then
                    lpNowModel = Format(nNowModelNo, "0000") & "_" & SelectCar(0).ModelName & "_" & lpRes
                Else
                    lpNowModel = "0000_DEFAULT"
                End If
            Else
                lpNowModel = "0000_DEFAULT"
            End If
            
            SysVar.lpModel = lpNowModel
        Else
            lpNowModel = Trim$(SysVar.lpModel)
        End If
        
        Call MenuSelect(FM_SPLASH)
        
        tmrPower.Enabled = POWERGET
        
        Call MenuSelect(FM_RUN)
    End If
End Sub

Private Sub MenuSelect(ByVal nMenu As Integer)
    bLogin = DEBUGMODE
    nMenu = IIf(SETUPSELECT = False And nMenu = FM_SETUPSELECT, FM_SETUP, nMenu)
    
    Select Case nMenu
        Case FM_RUN:
            bRunningGraphWin = False
            frmRun.Show vbModal
            
        Case FM_SETUP:
            If bLogin = False Then
                frmLogin.Show vbModal
            End If
            
            If bLogin Then
                frmSetup.Show vbModal
            End If
            
        Case FM_SETUPSELECT:
            If bLogin = False Then
                frmLogin.Show vbModal
            End If
            
            If bLogin Then
                frmSetupSelect.Show vbModal
            End If
            
        Case FM_IO:
            If bLogin = False Then
                frmLogin.Show vbModal
            End If
        
            If bLogin Then
                Do
                    DoEvents
                    bSplash = False
                    frmIO.Show vbModal
                    
                    If bSplash Then
                        frmSplash.Show vbModal
                    Else
                        Exit Do
                    End If
                Loop
            End If
            
        Case FM_SYSTEM:
            If bLogin = False Then
                frmLogin.Show vbModal
            End If
            
            If bLogin Then
                frmSystem.Show vbModal
            End If
        
        Case FM_DATABASE:
            frmDatabase.Show vbModal
        
        Case FM_EXIT:
            frmExit.Show vbModal
        
        Case FM_SPLASH:
            frmSplash.Show vbModal
    
    End Select
    
    bLogin = False
End Sub

Private Sub tmrPower_Timer()
    Call SetPowerProcess
End Sub

Private Sub Power_OnComm()
    Call GetPowerProcess
End Sub

Private Sub comNvh_OnComm()
    Call NvhReceived
End Sub

Private Sub comStepping1_OnComm()
    Call Step1Received
End Sub

Private Sub comStepping2_OnComm()
    Call Step2Received
End Sub


