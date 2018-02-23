VERSION 5.00
Begin VB.Form frmExit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "EXIT"
   ClientHeight    =   7740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11460
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
   Picture         =   "frmExit.frx":0000
   ScaleHeight     =   7740
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExitMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      Caption         =   "WINDOWS SHUT DOWN"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   36
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Index           =   0
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   9495
   End
   Begin VB.CommandButton btnExitMenu 
      BackColor       =   &H00F0F0F0&
      Caption         =   "TEST FINISH"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   36
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Index           =   1
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3060
      Width           =   9495
   End
   Begin VB.CommandButton btnExitMenu 
      BackColor       =   &H00F0F0F0&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   36
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Index           =   2
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5100
      Width           =   9495
   End
   Begin VB.Shape Shape1 
      Height          =   7695
      Left            =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' SHUTDOWN API ================================================================

Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As Long, ReturnLength As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Const ERROR_SUCCESS             As Long = 0&
Private Const TOKEN_ADJUST_PRIVILEGES   As Long = &H20
Private Const TOKEN_QUERY               As Long = &H8
Private Const SE_SHUTDOWN_NAME          As String = "SeShutdownPrivilege"
Private Const SE_PRIVILEGE_ENABLED      As Long = &H2
Private Const EWX_LOGOFF                As Long = 0
Private Const EWX_REBOOT                As Long = 2
Private Const EWX_FORCE                 As Long = 4
Private Const ANYSIZE_ARRAY             As Long = 1

Private Const EWX_SHUTDOWN              As Integer = 1

Private Type LUID
   LowPart      As Long
   HighPart     As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   pLuid        As LUID
   Attributes   As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount               As Long
   Privileges(ANYSIZE_ARRAY)    As LUID_AND_ATTRIBUTES
End Type


' START =======================================================================

Const MENU_SHUTDOWN As Integer = 0
Const MENU_EXIT     As Integer = 1
Const MENU_RETURN   As Integer = 2


Dim bLoading As Boolean

Private Sub Form_Activate()
    nNowForm = FM_EXIT
    
    Call LoadLangFile(FM_EXIT)
    
    If bLoading = False Then
        bLoading = True
    End If
End Sub

Private Sub Form_Load()
    bLoading = False
End Sub

Private Sub btnExitMenu_Click(Index As Integer)
    Select Case Index
        Case MENU_SHUTDOWN:
            If MsgBox("Do you want shutdown?", vbYesNo + vbCritical, "Checking") = vbYes Then
                Call SaveSystemFile
                
                frmMain.tmrPower.Enabled = False
                DoEvents
                Call Sleep(270)
                Call AD_Close
                Call StepClose
                
                If DOS(O_LIN_POWER) Then Call DO_Control(O_LIN_POWER, False)
                
                If POWERTYPE > 0 Then Call PowerClose
                If PLCUSE Then Call PLC_Close

                Call DoShutDown(EWX_SHUTDOWN)
            End If
            
        Case MENU_EXIT:
            Call SaveSystemFile
            
            frmMain.tmrPower.Enabled = False
            
            DoEvents
            Call Sleep(270)
            Call AD_Close
            Call StepClose
            
            If DOS(O_LIN_POWER) Then Call DO_Control(O_LIN_POWER, False)
            
            If POWERTYPE > 0 Then Call PowerClose
            If PLCUSE Then Call PLC_Close

            Unload Me
            End
            
        Case MENU_RETURN:
            Unload Me
    End Select
End Sub

' Function List ===============================================================

Private Function DoShutDown(ByVal nFlag As Integer)
    Dim hToken  As Long
    Dim G_Tkp   As TOKEN_PRIVILEGES
    Dim lRes    As Long
    
    On Error GoTo Err_SBCrt_ShutDown
    
    lRes = GetCurrentProcess()   ' -1 Á¤»ó
    lRes = OpenProcessToken(lRes, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken)
    
    If lRes = 0 Then
        Call MsgBox("OpenProcessToken Error")
        Exit Function
    End If
    
    Call LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, G_Tkp.Privileges(0).pLuid)
    
    G_Tkp.PrivilegeCount = 1
    G_Tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    
    Call AdjustTokenPrivileges(hToken, False, G_Tkp, 0, 0, 0)
    
    If (GetLastError() <> ERROR_SUCCESS) Then
        Call MsgBox("AdjustTokenPrivileges Error")
        Exit Function
    End If
    
    lRes = ExitWindowsEx(nFlag Or EWX_FORCE, 0)
    If lRes = 0 Then
        Call MsgBox("ExitWindowsEx Error")
        Exit Function
    End If
    
Exit_SBCrt_ShutDown:
    Exit Function
Err_SBCrt_ShutDown:
    Call MsgBox("SBCrt_ShutDown : ( " & Trim$(Err) & " )" & vbCrLf & Trim$(ERROR), vbCritical + vbOKOnly, "SHUTDOWN Error !!!")
    Resume Exit_SBCrt_ShutDown
End Function

