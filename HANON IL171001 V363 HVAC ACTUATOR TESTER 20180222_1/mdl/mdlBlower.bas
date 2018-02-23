Attribute VB_Name = "mdlBlower"
Option Explicit

Public Function BlowerTest() As Boolean
    Dim bRes As Boolean
    Dim dCurr As Double
    Dim dRpm As Double
    Dim i As Integer
    
    bRes = False
    
    If IsTestPos(TP_BLOWER, POS_INIT) Then
        Erase RunVar.bBlowerAddr
        
        Call DO_Control(O_BLOWER_POWER, True)
        Call DO_Control(O_BLOWER_DIRECTION, IIf(SetupVar.nBlowerDirection = 1, True, False))
        
        If SetupVar.nBlowerType = 1 Then
            Call DO_Control(O_BLOWER_PWM, True)
        End If
        
        For i = 0 To POS_BLOWER_HI
            RunVar.bBlowerAddr(i) = True
        Next
        
        ' 이름이 비었을 경우 띄위 테스트하기 위함
        For i = 0 To POS_BLOWER_HI ' 컨트롤 갯수
            If Len(Trim$(SetupVar.lpBlowerName(i))) = 0 Then RunVar.bBlowerAddr(i) = False
        Next
        
        
        RunVar.nBlowerPos = POS_RUN_INIT
    End If
    
    If IsTestPos(TP_BLOWER, POS_RUN_INIT) Then
        Call SetTime(TM_BLOWER)
        
        If DOS(O_BLOWER_01) Then Call DO_Control(O_BLOWER_01, False)
        If DOS(O_BLOWER_02) Then Call DO_Control(O_BLOWER_02, False)
        If DOS(O_BLOWER_03) Then Call DO_Control(O_BLOWER_03, False)
        If DOS(O_BLOWER_04) Then Call DO_Control(O_BLOWER_04, False)
        If DOS(O_BLOWER_05) Then Call DO_Control(O_BLOWER_05, False)
        If DOS(O_BLOWER_06) Then Call DO_Control(O_BLOWER_06, False)
        If DOS(O_BLOWER_07) Then Call DO_Control(O_BLOWER_07, False)
        If DOS(O_BLOWER_08) Then Call DO_Control(O_BLOWER_08, False)
        
        For i = 0 To POS_BLOWER_HI
            If RunVar.bBlowerAddr(i) Then
                Select Case SetupVar.nBlowerType
                    Case 0, 1:
                        Select Case i
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
                        If LINUSE Then
                            If LinBlrWrite(SetupVar.nLinSpeed(i)) Then
                                Call OnLog("[LIN] SPEED : " & SetupVar.nLinSpeed(i))
                            End If
                        End If
                
                End Select
                
                Erase dBlowerCurrBuf
                
                RunVar.nBlowerCount = 0
                RunVar.nBlowerPos = i
                RunVar.bBlowerAddr(i) = False
                
                
                
                If RunVar.nBlowerPos = 1 Then Call NvhSend(NVHLOW)
                If RunVar.nBlowerPos = 5 Then Call NvhSend(NVHHD)
                If RunVar.nBlowerPos = 7 Then Call NvhSend(NVHHV)
                
                
                
                
                
                Exit For
            End If
            
            Call SetTestPos(TP_BLOWER, POS_END)
        Next
    End If
    
    If (RunVar.nBlowerPos >= 0) And (RunVar.nBlowerPos <= POS_BLOWER_HI) Then
        dCurr = Format(ADRead(AD_BLOWER_CURR), SysVar.lpUnit(AD_BLOWER_CURR))
        
        If RunVar.nBlowerPos = POS_BLOWER_HI Then
            If RunVar.bRpmUse Then
                dRpm = Format(ADRead(AD_BLOWER_RPM), SysVar.lpUnit(AD_BLOWER_RPM))
                
'                dRpm = Abs(dRpm - 1.08)
'                dRpm = dRpm / 0.00054
'                dRpm = dRpm * 12.5
'                dRpm = dRpm / 60
'                dRpm = Abs((dRpm - 50) * 30)
            End If
        End If
        
        If ElapseTime(TM_BLOWER) > 0.1 Then
            If RunVar.bDispFlash Then
                frmRun.pnlBlowerCurr(RunVar.nBlowerPos).Caption = Format(dCurr, SysVar.lpUnit(AD_BLOWER_CURR))
                frmRun.pnlBlowerTime(RunVar.nBlowerPos).Caption = Format(ElapseTime(TM_BLOWER), "#0.0")
                
                If RunVar.nBlowerCount < MAX_INT Then
                    dBlowerCurrBuf(RunVar.nBlowerCount) = dCurr
                    RunVar.nBlowerCount = RunVar.nBlowerCount + 1
                End If
                
                If RunVar.nBlowerPos = POS_BLOWER_HI Then
                    If RunVar.bRpmUse Then
                        frmRun.pnlRpmCurr.Caption = Format(dRpm, "#0")
                    End If
                End If
            End If
        End If
        
        If ElapseTime(TM_BLOWER) >= SetupVar.dBlowerTime(RunVar.nBlowerPos) Then
            Call BlowerResult(RunVar.nBlowerPos, dCurr)
            
            If RunVar.nBlowerPos = 1 Then Call NvhSend(NVHSTOP)
            If RunVar.nBlowerPos = 5 Then Call NvhSend(NVHSTOP)
            If RunVar.nBlowerPos = 7 Then Call NvhSend(NVHSTOP)
            
            If RunVar.nBlowerPos = POS_BLOWER_HI Then
                If RunVar.bRpmUse Then
                    Call BlowerRpmResult(dRpm)
                End If
            End If
            
            RunVar.nBlowerPos = POS_RUN_INIT
        End If
        
        bRes = False
    End If
    
    If IsTestPos(TP_BLOWER, POS_END) Then
        Call SetVolt(SetupVar.dTestVolt)
        
        If SetupVar.nBlowerType = 2 And LINUSE Then
            Call LinBlrWrite(252)
            Call Delay(10)
            Call LinBlrWrite(252)
            Call Delay(10)
        End If
        
        Call DO_Control(O_BLOWER_POWER, False)
        Call DO_Control(O_BLOWER_PWM, False)
        Call DO_Control(O_BLOWER_DIRECTION, False)
        
        Call DO_Control(O_BLOWER_01, False)
        Call DO_Control(O_BLOWER_02, False)
        Call DO_Control(O_BLOWER_03, False)
        Call DO_Control(O_BLOWER_04, False)
        Call DO_Control(O_BLOWER_05, False)
        Call DO_Control(O_BLOWER_06, False)
        Call DO_Control(O_BLOWER_07, False)
        Call DO_Control(O_BLOWER_08, False)
        
        RunVar.bBlowerUse = False
        bRes = True
    End If
    
    BlowerTest = bRes
End Function

Private Function BlowerResult(ByVal nPos As Integer, ByVal dCurr As Double) As Boolean
    Dim lBkColor    As Long
    Dim bRes        As Boolean
    
    bRes = False
    
    If dCurr >= SetupVar.dBlowerCurrLo(nPos) And dCurr <= SetupVar.dBlowerCurrHi(nPos) Then
        frmRun.pnlBlowerResult(nPos).Caption = "OK"
        lBkColor = vbGreen
        
        bRes = True
    Else
        frmRun.pnlBlowerResult(nPos).Caption = "NG"
        lBkColor = vbRed
        
        RunVar.bReBlowerUse = True
        RunVar.bFinal = False
        
        Call SetPlc(PLC_BLOWER_CURR)
    End If
    
    frmRun.pnlBlowerCurr(nPos).Caption = Format(dCurr, SysVar.lpUnit(AD_BLOWER_CURR))
    frmRun.pnlBlowerCurr(nPos).BackColor = lBkColor
    frmRun.pnlBlowerTime(nPos).BackColor = CO_NONE
    frmRun.pnlBlowerResult(nPos).BackColor = lBkColor
    
    BlowerResult = bRes
End Function

Private Function BlowerRpmResult(ByVal dRpm As Double) As Boolean
    Dim lBkColor    As Long
    Dim bRes        As Boolean
    
    bRes = False
    
    If dRpm >= SetupVar.dRpmCurrLo And dRpm <= SetupVar.dRpmCurrHi Then
        frmRun.pnlRpmResult.Caption = "OK"
        
        lBkColor = vbGreen
        bRes = True
    Else
        frmRun.pnlRpmResult.Caption = "NG"
        
        lBkColor = vbRed
        
        RunVar.bReBlowerUse = True
        RunVar.bFinal = False
        
        Call SetPlc(PLC_BLOWER_RPM)
    End If
    
    frmRun.pnlRpmCurr.Caption = Format(dRpm, "#0")
    frmRun.pnlRpmCurr.BackColor = lBkColor
    frmRun.pnlRpmResult.BackColor = lBkColor
    
    BlowerRpmResult = bRes
End Function

Public Function BlowerVibResult() As Boolean
    Dim lBkColor    As Long
    Dim bRes        As Boolean
    Dim dVib        As Double
    
    bRes = False
    
    If SetupVar.nVibResultType = 0 Then ' PEAK
        dVib = FindVibPeak
    Else                                ' RMS
        dVib = FindVibRMS
    End If
    
    Call ReDraw(dVib, frmRun.picVib, GraphVar, dVibBuf(), RunVar.nVibCount)
    
    If dVib >= SetupVar.dVibCurrLo And dVib <= SetupVar.dVibCurrHi Then
        frmRun.pnlVibResult.Caption = "OK"
        
        lBkColor = vbGreen
        bRes = True
    Else
        frmRun.pnlVibResult.Caption = "NG"
        
        lBkColor = vbRed
        
        RunVar.bReBlowerUse = True
        RunVar.bFinal = False
        
        Call SetPlc(PLC_BLOWER_VIB)
    End If
    
    frmRun.pnlVibCurr.Caption = Format(dVib, SysVar.lpUnit(AD_VIB))
    frmRun.pnlVibCurr.BackColor = lBkColor
    frmRun.pnlVibResult.BackColor = lBkColor
    
    BlowerVibResult = bRes
End Function

Private Function FindVibPeak() As Double
    Dim i       As Integer
    Dim dMax    As Double
    
    Call OnLog("VIBRATION PEAK RESULT...")
    
    dMax = 0
    For i = 0 To RunVar.nVibCount
        If dMax < dVibBuf(1, i) Then
            dMax = dVibBuf(1, i)
            GraphVar.nPeak = i
        End If
    Next
    
    FindVibPeak = dMax
End Function

Private Function FindVibRMS() As Double
    Dim dPos    As Double
    Dim dAvg    As Double
    Dim i       As Integer
    
    Call BubbleSort(dVibBuf(), RunVar.nVibCount)
    
    dPos = RunVar.nVibCount / 100
    
    If dPos <> 0 Then
        GraphVar.PS = Int(dPos * SetupVar.dVibStart)
        GraphVar.PE = Int(dPos * SetupVar.dVibEnd)
    Else
        GraphVar.PS = 0
        GraphVar.PE = 0
    End If
    
    ' 입력값 퍼센트로 분리후 Average
     dAvg = 0
     
     For i = GraphVar.PS To GraphVar.PE
         dAvg = dAvg + dVibBuf(2, i)
     Next
     
     If (GraphVar.PE - GraphVar.PS) = 0 Then
        dAvg = 0
     Else
        dAvg = dAvg / (GraphVar.PE - GraphVar.PS)
    End If
    
    FindVibRMS = dAvg
End Function

