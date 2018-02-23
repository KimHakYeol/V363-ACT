Attribute VB_Name = "mdlLeak"
Option Explicit

Public Function Leak01Open() As Boolean
    If LEAKUSE = False Then
        Leak01Open = True
        
        Exit Function
    End If
    
    If SysVar.nLeakPort(0) = 0 Then
        Call MsgBox("LEAK 01 PORT CHECK...")
        
        Leak01Open = False
        
        Exit Function
    End If
    
    frmMain.comLeak01.CommPort = SysVar.nLeakPort(0)
    
    If frmMain.comLeak01.PortOpen = False Then
        frmMain.comLeak01.PortOpen = True
        Leak01Open = True
    End If
    
    Call Delay(20)
End Function

Public Function Leak02Open() As Boolean
    If LEAKUSE = False Then
        Leak02Open = True
        
        Exit Function
    End If
    
    If SysVar.nLeakPort(1) = 0 Then
        Call MsgBox("LEAK 02 PORT CHECK...")
        
        Leak02Open = False
        
        Exit Function
    End If
    
    frmMain.comLeak02.CommPort = SysVar.nLeakPort(1)
    
    If frmMain.comLeak02.PortOpen = False Then
        frmMain.comLeak02.PortOpen = True
        Leak02Open = True
    End If
    
    Call Delay(20)
End Function

Public Sub Leak01Close()
    If LEAKUSE = False Then
        Exit Sub
    End If
    
    If frmMain.comLeak01.PortOpen = True Then
        frmMain.comLeak01.PortOpen = False
    End If
    
    Call Delay(20)
End Sub

Public Sub Leak02Close()
    If LEAKUSE = False Then
        Exit Sub
    End If
    
    If frmMain.comLeak02.PortOpen = True Then
        frmMain.comLeak02.PortOpen = False
    End If
    
    Call Delay(20)
End Sub

Public Sub Leak01Received()
    Dim A As String
    Dim B As Integer
    Dim lpRes As String
    
    If LEAKUSE = False Then
        Exit Sub
    End If
    
    A = frmMain.comLeak01.Input
    
    If A <> "" Then
        B = Asc(A)
        
        If B = 13 Then
            lpLeakData(0) = ""
        ElseIf B = 3 Then
            lpRes = LeakParsing(lpLeakData(0))
            
            If Mid$(lpRes, 2, 1) = "0" Then
                lpRes = Replace(lpRes, "0", "", 1, 1)
            End If
            
            frmRun.pnlLeakData(0).Caption = lpRes
        Else
            lpLeakData(0) = lpLeakData(0) & A
        End If
    End If
End Sub

Public Sub Leak02Received()
    Dim A As String
    Dim B As Integer
    Dim lpRes As String
    
    If LEAKUSE = False Then
        Exit Sub
    End If
    
    A = frmMain.comLeak02.Input
    
    If A <> "" Then
        B = Asc(A)
        
        If B = 13 Then
            lpLeakData(1) = ""
        ElseIf B = 3 Then
            lpRes = LeakParsing(lpLeakData(1))
            
            If Mid$(lpRes, 2, 1) = "0" Then
                lpRes = Replace(lpRes, "0", "", 1, 1)
            End If
            
            frmRun.pnlLeakData(1).Caption = lpRes
        Else
            lpLeakData(1) = lpLeakData(1) & A
        End If
    End If
End Sub

Public Function LeakTest(ByVal nType As Integer) As Boolean
    Dim lBkColor As Long
    Dim bRes As Boolean
    Dim dTime As Double
    
    lBkColor = vbWhite
    
    If RunVar.nLeakPos(nType) = POS_INIT Then
        RunVar.bLeakResult(nType) = False
        
        Call OnLog("Send : Leak Model")
        Call SendLeakModel(SetupVar.nLeakModel)
        
        ' Clear Screen
        frmRun.pnlLeakData(nType).BackColor = lBkColor
        frmRun.pnlLeakData(nType).Caption = ""
        frmRun.pnlLeakResult(nType).BackColor = lBkColor
        frmRun.pnlLeakResult(nType).Caption = ""
        
        Select Case nType
            Case 0: Call DO_Control(O_LEAK01_STOP, True)
            Case 1: Call DO_Control(O_LEAK02_STOP, True)
        End Select
        
        Call DO_Control(O_LEAK_CLAMP_SOL, True)
        Call SetTime(TM_LEAKTEST)
        
        RunVar.nLeakPos(nType) = POS_START_INIT
    End If
    
    If RunVar.nLeakPos(nType) = POS_START_INIT Then
        If ElapseTime(TM_LEAKTEST) > 1 Then
            Select Case nType
                Case 0:
                    Call DO_Control(O_LEAK01_START, True)
                    Call DO_Control(O_LEAK01_STOP, False)
                
                Case 1:
                    Call DO_Control(O_LEAK02_START, True)
                    Call DO_Control(O_LEAK02_STOP, False)
                
            End Select
            
            Call SetTime(TM_LEAKTEST)
            
            RunVar.nLeakPos(nType) = POS_START_RUN
        End If
    End If
    
    If RunVar.nLeakPos(nType) = POS_START_RUN Then
        If ElapseTime(TM_LEAKTEST) > 2 Then
            Select Case nType
                Case 0: Call DO_Control(O_LEAK01_START, False)
                Case 1: Call DO_Control(O_LEAK02_START, False)
            End Select
            
            RunVar.nLeakPos(nType) = POS_RUN_INIT
        End If
    End If
    
    If RunVar.nLeakPos(nType) = POS_RUN_INIT Then
        Select Case nType
            Case 0: Call SetTime(TM_LEAK01)
            Case 1: Call SetTime(TM_LEAK02)
        End Select
        
        RunVar.nLeakPos(nType) = POS_RUN
    End If
    
    If RunVar.nLeakPos(nType) = POS_RUN Then
        Select Case nType
            Case 0:
                If ElapseTime(TM_LEAK01) > 30 Then
                    Call OnLog("[ERROR] LEAK 01 TIME OVER !!!")
                    Call LeakResult(0, False)
                    
                    RunVar.bReLeakUse(0) = True
                    RunVar.bLeakUse(nType) = False
                    RunVar.bFinal = False
                    RunVar.nLeakPos(nType) = POS_END
                End If
                
                If DIS(I_LEAK01_OK) Then RunVar.bLeakResult(0) = True: RunVar.nLeakPos(nType) = POS_END_INIT
                If DIS(I_LEAK01_NG) Then RunVar.bLeakResult(0) = False: RunVar.nLeakPos(nType) = POS_END_INIT
            
            Case 1:
                If ElapseTime(TM_LEAK02) > 30 Then
                    Call OnLog("[ERROR] LEAK 02 TIME OVER !!!")
                    Call LeakResult(1, False)
                    
                    RunVar.bReLeakUse(1) = True
                    RunVar.bLeakUse(nType) = False
                    RunVar.bFinal = False
                    RunVar.nLeakPos(nType) = POS_END
                End If
                
                If DIS(I_LEAK02_OK) Then RunVar.bLeakResult(1) = True: RunVar.nLeakPos(nType) = POS_END_INIT
                If DIS(I_LEAK02_NG) Then RunVar.bLeakResult(1) = False: RunVar.nLeakPos(nType) = POS_END_INIT
        
        End Select
    End If
    
    If RunVar.nLeakPos(nType) = POS_END_INIT Then
        Call LeakResult(nType, RunVar.bLeakResult(nType))
        
        If frmRun.pnlLeakData(nType).Caption <> "" Then
            RunVar.nLeakPos(nType) = POS_END_RUN
        End If
    End If
    
    If RunVar.nLeakPos(nType) = POS_END_RUN Then
        Select Case nType
            Case 0: bRes = DIS(I_LEAK01_END): Call SetTime(TM_LEAK01)
            Case 1: bRes = DIS(I_LEAK02_END): Call SetTime(TM_LEAK02)
        End Select
        
        If bRes Then
            RunVar.nLeakPos(nType) = POS_END_RUN + 1
        End If
    End If
    
    If RunVar.nLeakPos(nType) = POS_END_RUN + 1 Then
        Select Case nType
            Case 0: dTime = ElapseTime(TM_LEAK01)
            Case 1: dTime = ElapseTime(TM_LEAK02)
        End Select
        
        If dTime > 3 Then
            RunVar.nLeakPos(nType) = POS_END
        End If
    End If
    
    If RunVar.nLeakPos(nType) = POS_END Then
        Call DO_Control(O_LEAK_CLAMP_SOL, False)
        
        bRes = True
    End If
    
    LeakTest = bRes
End Function

Public Sub LeakResult(ByVal nType As Integer, ByVal bResult As Boolean)
    Dim lpData As String
    Dim lBkColor As Long
    
    If bResult Then
        lpData = "OK"
        lBkColor = vbGreen
    Else
        lpData = "NG"
        lBkColor = vbRed
        
        RunVar.bReLeakUse(nType) = True
        RunVar.bFinal = False
    End If
    
    frmRun.pnlLeakData(nType).BackColor = lBkColor
    frmRun.pnlLeakResult(nType).BackColor = lBkColor
    frmRun.pnlLeakResult(nType).Caption = lpData
End Sub

Private Sub SendLeakModel(ByVal nLeakNo As Integer)
    Dim bRes(4) As Boolean
    
    bRes(0) = IIf((nLeakNo And &H1) = &H1, True, False)
    bRes(1) = IIf((nLeakNo And &H2) = &H2, True, False)
    bRes(2) = IIf((nLeakNo And &H4) = &H4, True, False)
    bRes(3) = IIf((nLeakNo And &H8) = &H8, True, False)
    bRes(4) = IIf((nLeakNo And &H10) = &H10, True, False)
    
    Call DO_Control(O_LEAK_MODE_ID1, bRes(0))
    Call DO_Control(O_LEAK_MODE_ID2, bRes(1))
    Call DO_Control(O_LEAK_MODE_ID4, bRes(2))
    Call DO_Control(O_LEAK_MODE_ID8, bRes(3))
    
    Call SetTime(TM_SOL)
    Do
        DoEvents
        If ElapseTime(TM_SOL) > 0.5 Then
            Exit Do
        End If
    Loop
End Sub

Public Sub LeakDummyTest()
    Dim bRes(1) As Boolean
    
    nNowForm = FM_RUN
    
    If nLeakDummyPos = POS_INIT Then
        Call OnLog("LEAK CAL : START")
        Call SendLeakModel(SysVar.nLeakGroup)
        Call SetTime(TM_LEAKSOL)
        
        nLeakDummyPos = POS_START_INIT
    End If
    
    If nLeakDummyPos = POS_START_INIT Then
        If ElapseTime(TM_LEAKSOL) > 1 Then
            Call DO_Control(O_LEAK_CLAMP_SOL, True)
            
            If SetupVar.bLeakUse(0) Then Call DO_Control(O_LEAK01_CAL, True)
            If SetupVar.bLeakUse(1) Then Call DO_Control(O_LEAK02_CAL, True)
            
            If SetupVar.bLeakUse(0) Then Call DO_Control(O_LEAK01_STOP, True)
            If SetupVar.bLeakUse(1) Then Call DO_Control(O_LEAK02_STOP, True)
            
            Call SetTime(TM_LEAKSOL)
            
            nLeakDummyPos = POS_START_INIT + 1
        End If
    End If
    
    If nLeakDummyPos = POS_START_INIT + 1 Then
        If ElapseTime(TM_LEAKSOL) > 1 Then
            If SetupVar.bLeakUse(0) Then Call DO_Control(O_LEAK01_STOP, False)
            If SetupVar.bLeakUse(1) Then Call DO_Control(O_LEAK02_STOP, False)
            
            If SetupVar.bLeakUse(0) Then Call DO_Control(O_LEAK01_START, True)
            If SetupVar.bLeakUse(1) Then Call DO_Control(O_LEAK02_START, True)
            
            Call SetTime(TM_LEAKSOL)
            
            nLeakDummyPos = POS_START_RUN
        End If
    End If
    
    If nLeakDummyPos = POS_START_RUN Then
        If ElapseTime(TM_LEAKSOL) > 1 Then
            If SetupVar.bLeakUse(0) Then Call DO_Control(O_LEAK01_START, False)
            If SetupVar.bLeakUse(1) Then Call DO_Control(O_LEAK02_START, False)
            
            Call SetTime(TM_LEAKSOL)
            
            nLeakDummyPos = POS_RUN
        End If
    End If
    
    If nLeakDummyPos = POS_RUN Then
        If SetupVar.bLeakUse(0) Then
            bRes(0) = DIS(I_LEAK01_NG)
        Else
            bRes(0) = True
        End If
        
        If SetupVar.bLeakUse(1) Then
            bRes(1) = DIS(I_LEAK02_NG)
        Else
            bRes(1) = True
        End If
        
        If DIS(I_LEAK01_NG) Then
            RunVar.bLeakResult(0) = True
        ElseIf DIS(I_LEAK01_OK) Then
            RunVar.bLeakResult(0) = False
        End If
        
        If DIS(I_LEAK02_NG) Then
            RunVar.bLeakResult(1) = True
        ElseIf DIS(I_LEAK02_OK) Then
            RunVar.bLeakResult(1) = False
        End If
        
        If bRes(0) And bRes(1) Then
            nLeakDummyPos = POS_END_INIT
        End If
        
        If DIS(I_LEAK01_OK) Or DIS(I_LEAK02_OK) Then
            nLeakDummyPos = POS_END_INIT + 1
        End If
        
        If ElapseTime(TM_LEAKSOL) > 30 Then
            nLeakDummyPos = POS_END_INIT + 1
        End If
    End If
    
    If nLeakDummyPos = POS_END_INIT Then
        Call SetTime(TM_LEAKSOL)
        
        nLeakDummyPos = POS_END
    End If
    
    If nLeakDummyPos = POS_END_INIT + 1 Then
        Call DO_Control(O_BUZZER, True)
        Call SetTime(TM_LEAKSOL)
        
        nLeakDummyPos = POS_END
    End If
    
    If nLeakDummyPos = POS_END Then
        If ElapseTime(TM_LEAKSOL) > 1 Then
            If SetupVar.bLeakUse(0) Then Call DO_Control(O_LEAK01_SOL, True)
            If SetupVar.bLeakUse(1) Then Call DO_Control(O_LEAK02_SOL, True)
            
            If SetupVar.bLeakUse(0) Then Call DO_Control(O_LEAK01_CAL, False)
            If SetupVar.bLeakUse(1) Then Call DO_Control(O_LEAK02_CAL, False)
            Call SetTime(TM_LEAKSOL)
            
            nLeakDummyPos = POS_END + 1
        End If
    End If
    
    If nLeakDummyPos = POS_END + 1 Then
        If ElapseTime(TM_LEAKSOL) > 1 Then
            If SetupVar.bLeakUse(0) Then Call LeakResult(0, RunVar.bLeakResult(0))
            If SetupVar.bLeakUse(1) Then Call LeakResult(1, RunVar.bLeakResult(1))
            
            If SetupVar.bLeakUse(0) Then Call DO_Control(O_LEAK01_SOL, False)
            If SetupVar.bLeakUse(1) Then Call DO_Control(O_LEAK02_SOL, False)
            
            Call DO_Control(O_BUZZER, False)
            Call SetTime(TM_LEAKSOL)
            
            nLeakDummyPos = POS_END + 2
        End If
    End If
    
    If nLeakDummyPos = POS_END + 2 Then
        If ElapseTime(TM_LEAKSOL) > 3 Then
            Call OnLog("LEAK CAL : END")
            Call DO_Control(O_LEAK_CLAMP_SOL, False)
            Call SetTime(TM_LEAKSOL)
            
            nLeakDummyPos = 0
            bLeakDummyUse = False
        End If
    End If
End Sub

Public Sub DispLeakCal(ByVal bVisible As Boolean)
    Dim i As Integer
    
    frmRun.picLog.BackColor = vbBlue
    frmRun.picLog.ZOrder 0
    frmRun.picLog.Visible = bVisible
    
    For i = 0 To frmRun.lblLog.UBound
        frmRun.lblLog(i).Caption = ""
        frmRun.lblLog(i).ForeColor = vbWhite
    Next
    
    If bVisible Then
        frmRun.lblLog(1).Caption = "- LEAK CALIBRATION - "
        frmRun.lblLog(2).Caption = "리크 캘리브레이션을"
        frmRun.lblLog(3).Caption = "실행하십시오."
    End If
End Sub

Private Function LeakParsing(ByVal lpCmd As String) As String
    Dim Var As Variant
    
    Var = Split(lpCmd, ",")
    
    LeakParsing = Var(3)
End Function

Public Sub LeakSolRelease()
    Call DO_Control(O_LEAK01_STOP, True)
    Call DO_Control(O_LEAK02_STOP, True)
    
    Call SetTime(TM_LEAKSOL)
    Do
        DoEvents
        
        If ElapseTime(TM_LEAKSOL) > 2 Then
            Call DO_Control(O_LEAK_CLAMP_SOL, False)
            
            Exit Do
        End If
    Loop
    
    Call DO_Control(O_LEAK01_START, False)
    Call DO_Control(O_LEAK02_START, False)
    Call DO_Control(O_LEAK01_STOP, False)
    Call DO_Control(O_LEAK02_STOP, False)
    
    bLeakSol = False
End Sub

Public Sub LeakIOClear()
    Call DO_Control(O_LEAK01_STOP, False)
    Call DO_Control(O_LEAK02_STOP, False)
    Call DO_Control(O_LEAK01_START, False)
    Call DO_Control(O_LEAK02_START, False)
    Call DO_Control(O_LEAK01_CAL, False)
    Call DO_Control(O_LEAK02_CAL, False)
End Sub
