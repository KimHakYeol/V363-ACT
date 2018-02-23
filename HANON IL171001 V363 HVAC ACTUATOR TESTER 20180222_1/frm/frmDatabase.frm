VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDatabase 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DATABASE"
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15000
   ScaleWidth      =   19110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnDebugDelete 
      Caption         =   "Today Delete"
      Height          =   435
      Left            =   13740
      TabIndex        =   11
      Top             =   1140
      Visible         =   0   'False
      Width           =   1695
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
      Height          =   1395
      Left            =   15600
      TabIndex        =   10
      Top             =   180
      Width           =   3315
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "OPEN"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   27.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   12120
      TabIndex        =   9
      Top             =   180
      Width           =   3315
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   180
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   12975
      Left            =   120
      TabIndex        =   8
      Top             =   1740
      Width           =   18795
      _ExtentX        =   33152
      _ExtentY        =   22886
      _Version        =   393216
      Cols            =   7
      FixedCols       =   6
      BackColorFixed  =   14737632
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.0%"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   21.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   9390
      TabIndex        =   7
      Top             =   900
      Width           =   1185
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERCENT"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Index           =   4
      Left            =   9195
      TabIndex        =   3
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label lblNg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   21.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   7035
      TabIndex        =   6
      Top             =   900
      Width           =   285
   End
   Begin VB.Label lblOk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   21.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   4530
      TabIndex        =   5
      Top             =   900
      Width           =   285
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   21.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   1935
      TabIndex        =   4
      Top             =   900
      Width           =   285
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NG"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Index           =   3
      Left            =   6900
      TabIndex        =   2
      Top             =   420
      Width           =   555
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Index           =   2
      Left            =   4410
      TabIndex        =   1
      Top             =   420
      Width           =   525
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Index           =   1
      Left            =   1515
      TabIndex        =   0
      Top             =   420
      Width           =   1125
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim bLoading As Boolean

Private Sub Form_Activate()
    nNowForm = FM_DATABASE
    
    Call LoadLangFile(FM_DATABASE)
    
    frmDatabase.btnDebugDelete.Visible = DEBUGMODE
    
    If bLoading = False Then
        bLoading = True
        Call OpenDatabase(True)
    End If
End Sub

Private Sub Form_Load()
    bLoading = False
End Sub

Private Sub btnDebugDelete_Click()
    On Error Resume Next
    
    If MsgBox("Are you sure you want to delete?", vbYesNo) = vbYes Then
        Call Kill(SysVar.lpSaveFileName)
    End If
End Sub

Private Sub btnReturn_Click()
    Unload Me
End Sub

Private Sub btnOpen_Click()
    Call OpenDatabase(False)
End Sub

' Function List ===============================================================

Private Sub OpenDatabase(ByVal bStart As Boolean)
    Dim lpFileName As String
    
    On Error GoTo ErrHandler_OpenDatabase:
    
    If bStart = False Then
        dlgFile.InitDir = App.Path ' MakeFolder
        dlgFile.Filter = "Database Files (*.CSV)|*.CSV"
        dlgFile.FileName = ""
        dlgFile.CancelError = True                      ' ÀúÀå Dialog ¿¡¼­ Cancel ´­·¶À» ¶§´Â ±×³É ³ª°£´Ù.
        dlgFile.ShowOpen
        lpFileName = dlgFile.FileName
    Else
        lpFileName = Trim$(SysVar.lpSaveFileName)
    End If
    
    If Len(lpFileName) = 0 Then
        Call MsgBox("File not found.")
        Exit Sub
    End If

    Call OpenDataFile(lpFileName)
    
    Exit Sub
    
ErrHandler_OpenDatabase:
    Call MsgBox("Data file open error." & vbCrLf & "Location : OpenDatabase()", vbOKOnly, "")
End Sub

Private Sub OpenDataFile(ByVal lpFileName As String)
    Const TOTAL_        As Integer = 0
    Const OK_           As Integer = 1
    Const NG_           As Integer = 2
    Const POS_RESULT    As Integer = 5
    
    Dim lCounter(3) As Long
    Dim lpData As String
    Dim lLine As Long
    Dim dPercent As Double
    Dim i As Integer
    Dim nFileNo As Integer
    
    On Error GoTo ErrHandler_OpenDataFile
    
    nFileNo = FreeFile
    
    MSGrid.Visible = False
    
    MSGrid.Cols = SysVar.nDBCol
    MSGrid.Rows = SysVar.nDBRow
    
    For i = 0 To SysVar.nDBCol - 1
        MSGrid.ColWidth(i) = 1600
    Next
    
    MSGrid.ColWidth(0) = 800
    
    MSGrid.ColWidth(1) = 1400
    MSGrid.ColWidth(2) = 1800
    
    MSGrid.ColWidth(5) = 2000
    
    MSGrid.Clear
    
    Me.Caption = lpFileName
    
    For i = 0 To 3
        lCounter(i) = 0
    Next
    
    lLine = 0
    i = 0
    
    MSGrid.TextMatrix(0, 0) = "TEST NO."
    
    Open lpFileName For Input As #nFileNo
        Do Until EOF(nFileNo)
            Input #nFileNo, lpData
            i = i + 1
            
            If i >= SysVar.nDBCol Then ' Add Column
                i = 1
                lLine = lLine + 1
                
                If lLine >= MSGrid.Rows Then
                    MSGrid.Rows = MSGrid.Rows + 1
                End If
                
                MSGrid.TextMatrix(lLine, 0) = str(lLine) ' Column Number
            End If
            
            MSGrid.TextMatrix(lLine, i) = lpData ' Write Data
            
            ' Coodinate for Color and Align
            MSGrid.Col = i
            MSGrid.Row = lLine
            MSGrid.RowHeight(lLine) = 300
            
            If MSGrid.Row > 0 Then
                If i = POS_RESULT Then
                    If lpData Like UCase("*OK*") Or lpData Like UCase("*ÇÕ°Ý*") Or lpData Like UCase("*¾çÇ°*") Then
                        lCounter(OK_) = lCounter(OK_) + 1
                        lCounter(TOTAL_) = lCounter(TOTAL_) + 1
                    ElseIf lpData Like UCase("*NG*") Or lpData Like UCase("*ºÒÇÕ°Ý*") Or lpData Like UCase("*ºÒ·®*") Then
                        lCounter(NG_) = lCounter(NG_) + 1
                        lCounter(TOTAL_) = lCounter(TOTAL_) + 1
                    End If
                End If
                
                If i >= POS_RESULT Then
                    If lpData Like UCase("*OK*") Or lpData Like UCase("*ÇÕ°Ý*") Or lpData Like UCase("*¾çÇ°*") Then
                        MSGrid.CellBackColor = vbGreen
                    ElseIf lpData Like UCase("*NG*") Or lpData Like UCase("*ºÒÇÕ°Ý*") Or lpData Like UCase("*ºÒ·®*") Then
                        MSGrid.CellBackColor = vbRed
                    End If
                    
                    If Left(Trim$(lpData), 1) = "#" Then
                        MSGrid.CellBackColor = vbRed
                        MSGrid.TextMatrix(lLine, i) = Right(lpData, Len(lpData) - 1)
                    End If
                End If
            End If
            
            MSGrid.CellAlignment = flexAlignCenterCenter
        Loop
        
        MSGrid.Visible = True
        
        If lLine = 0 Then
            Call MsgBox("Empty data...")
        Else
            dPercent = (1 - (lCounter(NG_) / (lCounter(OK_) + lCounter(NG_)))) * 100
            
            lblTotal.Caption = lCounter(TOTAL_)
            lblOk.Caption = lCounter(OK_)
            lblNg.Caption = lCounter(NG_)
            
            lblPercent.Caption = Format(dPercent, "#0.0") & "%"
        End If
    
    Close #nFileNo
    Exit Sub

ErrHandler_OpenDataFile:
    MSGrid.Visible = True
    
    Call MsgBox("Data file open error" & vbCrLf & "Location : OpenDataFile()", vbOKOnly, "")
    
    Close #nFileNo
End Sub
