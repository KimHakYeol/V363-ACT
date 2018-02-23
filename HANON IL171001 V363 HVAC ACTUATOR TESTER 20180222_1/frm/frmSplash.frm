VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SPLASH"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11745
   BeginProperty Font 
      Name            =   "³ª´®°íµñ"
      Size            =   12
      Charset         =   129
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHold 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10020
      TabIndex        =   3
      Top             =   1020
      Width           =   1215
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   480
      ScaleHeight     =   585
      ScaleWidth      =   10725
      TabIndex        =   2
      Top             =   4680
      Width           =   10755
   End
   Begin VB.ListBox lstMsg 
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
      Height          =   2580
      Left            =   480
      TabIndex        =   1
      Top             =   1620
      Width           =   10755
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "2018.02.20 START"
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
      Height          =   225
      Index           =   0
      Left            =   540
      TabIndex        =   6
      Top             =   5460
      Width           =   1710
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "IL171001"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10155
      TabIndex        =   5
      Top             =   4320
      Width           =   1080
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   18
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   4260
      Width           =   1695
   End
   Begin VB.Label lblProgramTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "V363 HVAC ACTUATOR TESTER"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   27.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   540
      TabIndex        =   0
      Top             =   480
      Width           =   8220
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim bLoading As Boolean

Private Sub Form_Activate()
    nNowForm = FM_SPLASH
    
    bLoading = False
    
    If bLoading = False Then
        bLoading = True
        
        Call InitStart
    End If
End Sub

Private Sub Form_Load()
    bLoading = False
End Sub

Private Sub chkHold_Click()
    '
End Sub

Private Function InitStart()
    Dim i As Integer
    Dim bRes As Boolean
    
    bRes = True
    
    If bRes Then
        Call OnLog("CHECKING SYSTEM PERFORMANCE...")
        
        bRes = InitPerformanceLevelUp
        
        If bRes = False Then
            Call OnLog("ERROR : SYSTEM PERFORMANCE FAILED....")
            chkHold.Value = 1
        End If
        
        Call Drawing(1000)
        Call Drawing(1000)
        Call Drawing(1000)
        Call Drawing(1000)
        Call Drawing(1000)
    End If
    
    If POWERTYPE > 0 Then
        If bRes Then
            Call OnLog("POWER SUPPLY COMMUNICATION INITIALIZE...")
            bRes = PowerOpen
            
            If bRes = False Then
                Call OnLog("ERROR : POWER SUPPLY COMMUNICATION FAILED !!!")
                chkHold.Value = 1
            End If
        End If
    End If
    
    If PLCUSE Then
        If bRes Then
            i = 0
            
            Call OnLog("PLC COMMUNICATION INITIALIZE...")
            
PLC_RESTART:
            
            bRes = PLC_Open
            
            Call Drawing(1000)
            
            If bRes = False Then
                If i > 5 Then
                    Call OnLog("ERROR : PLC COMMUNICATION !!!")
                    chkHold.Value = 1
                Else
                    i = i + 1
                    
                    GoTo PLC_RESTART
                End If
            End If
        End If
    End If
    
    If NVHUSE Then
        If bRes Then
            Call OnLog("NVH COMMUNICATION INITIALIZE...")
            bRes = NvhOpen
            Call Drawing(1000)
            
            If bRes = False Then
                Call OnLog("ERROR : NVH COMMUNICATION FAILED !!!")
                chkHold.Value = 1
            End If
        End If
    End If
    
    If bRes Then
        Call OnLog("DIO CARD INITIALIZE...")
        Call Drawing(1000)
        Call InitMapping
    End If
    
    If bRes Then
        Call OnLog("A/D CARD INITIALIZE...")
        Call Drawing(1000)
    End If
    
    If bRes Then
        bRes = InitAD
        
        Call Drawing(1000)
        
        If bRes = False Then
            Call OnLog("ERROR : A/D CARD FAILED !!!")
            chkHold.Value = 1
        End If
        
        If bRes Then
            bRes = InitDA
            
            If bRes = False Then
                Call OnLog("ERROR : D/A CARD FAILED !!!")
                chkHold.Value = 1
            End If
        End If
    End If
    
    If STEPUSE And bRes Then
        Call OnLog("STEPPING BOARD INITIALIZE...")
        
        Call DO_Control(O_STEPPING1_RESET, True)
        Call DO_Control(O_STEPPING2_RESET, True)
        Call Delay(2000)
        
        Call DO_Control(O_STEPPING1_RESET, False)
        Call DO_Control(O_STEPPING2_RESET, False)
        Call Delay(2000)
        
        bRes = StepOpen
        
        Call Drawing(1000)
        
        If bRes = False Then
            Call OnLog("ERROR : STEPPING BOARD FAILED !!!")
            chkHold.Value = 1
        End If
    End If
    
    If bRes Then
        If DOS(O_LIN_POWER) = False Then Call DO_Control(O_LIN_POWER, True)
    End If
    
    If chkHold.Value = 1 Then
        chkHold.ForeColor = vbRed
        Call DrawBox(100, True)
        Do
            DoEvents
            If chkHold.Value = 0 Then Exit Do
        Loop
        
        If MsgBox("Do you want continue?", vbYesNo + vbQuestion, "Question") = vbNo Then
            End
        End If
    End If
    
    Unload Me
End Function

Private Sub Drawing(ByVal nWaitTime As Integer)
    Dim dTime   As Double
    
    If DEBUGMODE Then Exit Sub
    
    Call SetGraph
    
    Call SetTime(TM_SPLASH)
    Do
        DoEvents
        dTime = ElapseTime(TM_SPLASH)
        
        If dTime > (nWaitTime / 1000) Then Exit Do
        
        Call DrawBox(dTime * 100)
    Loop
    
    Call Delay(350)
End Sub

Private Sub SetGraph()
    lblPercent.Caption = "0%"
    picProgress.Scale (0, 0)-(100, 100)
    picProgress.Cls
End Sub

Private Sub DrawBox(ByVal nValue As Integer, Optional ByVal bError As Boolean = False)
    lblPercent.Caption = Format(nValue, "#0") & "%"
    
    If bError = False Then
        picProgress.Line (0, 0)-(nValue, 100), RGB(0, 255, 0), BF
    Else
        picProgress.Line (0, 0)-(nValue, 100), RGB(255, 0, 0), BF
    End If
End Sub

