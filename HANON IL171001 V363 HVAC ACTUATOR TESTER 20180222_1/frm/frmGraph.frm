VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmGraph 
   BorderStyle     =   4  '°íÁ¤ µµ±¸ Ã¢
   Caption         =   "GRAPH"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15210
   BeginProperty Font 
      Name            =   "³ª´®°íµñ"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin MSChart20Lib.MSChart MSChart 
      Height          =   5175
      Left            =   300
      OleObjectBlob   =   "frmGraph.frx":0000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4740
      Width           =   14655
   End
   Begin VB.CommandButton btnReturn 
      Caption         =   "RETURN"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   26.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10860
      TabIndex        =   7
      Top             =   180
      Width           =   4275
   End
   Begin MSACAL.Calendar calCalender 
      Height          =   3015
      Left            =   10860
      TabIndex        =   0
      Top             =   1140
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   5318
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2008
      Month           =   7
      Day             =   27
      DayLength       =   0
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   3015
      Left            =   300
      TabIndex        =   2
      Top             =   900
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5318
      _Version        =   393216
      Rows            =   52
      Cols            =   5
      FixedRows       =   2
      FixedCols       =   0
      BackColorFixed  =   -2147483643
      BackColorBkg    =   12632256
      Appearance      =   0
   End
   Begin VB.Shape Shape1 
      Height          =   3855
      Index           =   0
      Left            =   180
      Top             =   180
      Width           =   10575
   End
   Begin VB.Label lbTemp 
      Caption         =   "Day ratio"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   300
      TabIndex        =   6
      Top             =   420
      Width           =   1335
   End
   Begin VB.Label lblDate 
      Caption         =   "20070805"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1620
      TabIndex        =   5
      Top             =   420
      Width           =   1815
   End
   Begin VB.Label lblTemp 
      Caption         =   "Month ratio"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   4
      Top             =   4500
      Width           =   1335
   End
   Begin VB.Label lblDate 
      Caption         =   "200708¿ù"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1620
      TabIndex        =   3
      Top             =   4500
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      Height          =   5775
      Index           =   1
      Left            =   180
      Top             =   4260
      Width           =   14895
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim lpOldDate As String

Private Sub Form_Activate()
    nNowForm = FM_GRAPH
    
    Call LoadLangFile(FM_GRAPH)
End Sub

Private Sub Form_Load()
    Call ClearViewC
    
    calCalender.Today
    
    Call DisplayCalender(calCalender.Value, StcalVar())
    Call ChartLoading(calCalender.Value)
End Sub

Private Sub btnReturn_Click()
    Unload Me
End Sub

Private Sub calCalender_Click()
    Call DataLoading(calCalender.Value)
    
    If Format(calCalender.Value, "yyyymm") <> Format(lpOldDate, "yyyymm") Then
        Call ChartLoading(calCalender.Value)
    End If
End Sub

' =============================================================================
' ÇÔ¼öµé
Private Sub ClearViewC()
    Dim i As Integer
    
    With MSGrid
        .Rows = 52
        .Cols = 5
        .FixedRows = 2
        .FixedCols = 0
        .Clear
        For i = 1 To 4
            .ColAlignment(i) = 1
            .ColWidth(i) = (.Width / 15) * 2.25
        Next
        .ColWidth(0) = (.Width / 15) * 5.5
        .Clear
    End With
End Sub

Private Sub DisplayCalender(ByVal lpValue As String, ByRef SSD() As StatisticalVariables)
    Dim i As Integer
    
    With MSGrid
        .Visible = False
        
        .Clear
        
        .Row = 0
        .Col = 0
        .Text = "Model"
        .Col = 1
        .Text = "Total"
        .Col = 2
        .Text = "OK"
        .Col = 3
        .Text = "NG"
        .Col = 4
        .Text = "Ratio"
        .Row = 1
        .Col = 0
        .Text = "ALL"
        .Col = 1
        .Text = str(SSD(0).lAllCounter)
        .Col = 2
        .Text = str(SSD(0).lOkCounter)
        .Col = 3
        .Text = str(SSD(0).lNgCounter)
        .Col = 4
        .Text = Format(SSD(0).dPercent, "#0.0")
        
        
        For i = 1 To 50
            If Trim$(SSD(i).lpModelName) = "" Then Exit For
            .Row = i
            .Col = 0
            .Text = Trim$(SSD(i).lpModelName)
            .Col = 1
            .Text = str(SSD(i).lAllCounter)
            .Col = 2
            .Text = str(SSD(i).lOkCounter)
            .Col = 3
            .Text = str(SSD(i).lNgCounter)
            .Col = 4
            .Text = Format(SSD(i).dPercent, "#0.0")
        Next
        
        .Visible = True
    End With
    
    lblDate(0).Caption = lpValue
End Sub

Private Sub DataLoading(ByVal DateV As String)
    Dim vFileNumver As Variant
    Dim lpFileName As String
    Dim lpString As String
    Dim i As Integer
    Dim nFileNo As Integer
    
    On Error Resume Next
    
    nFileNo = FreeFile
    
    i = 0
    
    lpFileName = lpPath + "\StatiDataFile\" + Format(DateV, "yyyy") + "\" + Format(DateV, "yyyymm") + "\" + Format(DateV, "yyyymmdd") + ".csv"
    
    MSGrid.Visible = False
    If SearchFile(lpFileName) Then
        MSGrid.Clear
        Open lpFileName For Input As #nFileNo
        vFileNumver = LOF(nFileNo)
        If vFileNumver <> 0 Then
            Do While Not EOF(nFileNo)
                
                MSGrid.Row = i
                For vFileNumver = 0 To 4
                    Input #nFileNo, lpString
                    MSGrid.Col = vFileNumver
                    MSGrid.Text = Trim$(lpString)
                Next
                i = i + 1
                If EOF(nFileNo) = True Then
                    Exit Do
                End If
            Loop
        End If
        Close #nFileNo
        lblDate(0).Caption = DateV
    Else
        MSGrid.Clear
        MSGrid.Row = 0
        MSGrid.Col = 0
        MSGrid.Text = "Model"
        MSGrid.Col = 1
        MSGrid.Text = "Total"
        MSGrid.Col = 2
        MSGrid.Text = "OK"
        MSGrid.Col = 3
        MSGrid.Text = "NG"
        MSGrid.Col = 4
        MSGrid.Text = "Ratio"
        
        lblDate(0).Caption = "Empty data."
    End If
    
    On Error GoTo 0
    MSGrid.Visible = True
End Sub

Private Sub ChartLoading(ByVal DateV As String)
    Dim vFileNumber As Variant
    Dim lpDateInfo As String
    Dim lpFileName As String
    Dim lpString As String
    Dim i As Integer
    Dim nFileNo As Integer
    
    On Error Resume Next
    
    nFileNo = FreeFile
    
    lpOldDate = DateV
    lblDate(1).Caption = Format(DateV, "yyyymm") + " "
    
    For i = 1 To 31
        lpDateInfo = Format(DateV, "yyyymm") + Format(i, "00")
        lpFileName = lpPath + "\StatiDataFile\" + Format(DateV, "yyyy") + "\" + Format(DateV, "yyyymm") + "\" + lpDateInfo + ".csv"
        
        If SearchFile(lpFileName) Then
            lpString = "0"
            Open lpFileName For Input As #nFileNo
            vFileNumber = LOF(nFileNo)
            If vFileNumber <> 0 Then
                For vFileNumber = 0 To 9
                    Input #nFileNo, lpString
                    If EOF(nFileNo) = True Then
                        Exit For
                    End If
                Next
            End If
            Close #nFileNo
            
            MSChart.Row = i
            MSChart.RowLabel = Trim$(str(i))
            MSChart.Data = Val(lpString)
        Else
            MSChart.Row = i
            MSChart.RowLabel = Trim$(str(i))
            MSChart.Data = 0
        End If
    Next
    On Error GoTo 0
End Sub

