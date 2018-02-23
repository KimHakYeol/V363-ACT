Attribute VB_Name = "mdlAct02"
Option Explicit

Private Const ACT_DA_NO As Integer = 1
Private Const ACT_FINAL_POS As Integer = 4
Private Const ACT_STALL1 As Integer = 0
Private Const ACT_STALL2 As Integer = 3

Public Sub SetAct02Move(ByVal bStart As Boolean)
    Dim nRes As Integer
    
    RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
    RunVar.nActEndPosCount(ACT_DA_NO) = 0
    RunVar.nActCount(ACT_DA_NO) = 0
    Erase dAct02CurrBuf
    
    If SetupVar.nAct02TestType = 1 Then
        nRes = IIf(SetupVar.dAct02SetVolt(ACT_FINAL_POS) < (((SetupVar.dAct02SetVolt(2) - SetupVar.dAct02SetVolt(1)) / 2) + SetupVar.dAct02SetVolt(1)), 1, 2)
        
        If SetupVar.nAct02Direction = 0 And bStart Then
            nRes = Abs(nRes - 1)
            
            Call OutDa(ActNo(ACT_DA_NO).DA_NO, SetupVar.dAct02SetVolt(nRes), SysVar.bPercent(ActNo(ACT_DA_NO).AD_VOLT))
        Else
            Call OutDa(ActNo(ACT_DA_NO).DA_NO, SetupVar.dAct02SetVolt(ACT_FINAL_POS), SysVar.bPercent(ActNo(ACT_DA_NO).AD_VOLT))
        End If
    End If
    
    Call DO_Control(ActNo(ACT_DA_NO).O_POWER, True)
    Call SetTime(TM_ACT02)
End Sub

Public Sub SetStepAct02Move()
    If SetupVar.nAct02TestType <> 0 Then Exit Sub
    
    Call DO_Control(ActNo(ACT_DA_NO).O_POWER, True)
    
    bSteppingWrite(1) = SetupVar.bAct02Use
    
    If SetupVar.bAct02Use Then
        lpWriteStepRotation(1) = STEP_ROTA_P2
        lpWriteStepData(2) = STEP_LITTLE1
        lpWriteStepData(3) = STEP_LITTLE2
    Else
        lpWriteStepRotation(1) = STEP_NULL
        lpWriteStepData(2) = STEP_NULL
        lpWriteStepData(3) = STEP_NULL
    End If
End Sub

Public Function GetAct02Move(ByVal bStart As Boolean) As Boolean
    Dim dCurr As Double
    Dim dVolt As Double
    Dim dTime As Double
    Dim bRes As Boolean
    Dim nRes As Integer
    Dim dMaxTime As Double
    Dim dVoltLo As Double
    Dim dVoltHi As Double
    
    If SetupVar.nAct02TestType = 0 Then
        GetAct02Move = True
        
        Exit Function
    End If
    
    bRes = False
    
    dCurr = Format(ADRead(ActNo(ACT_DA_NO).AD_CURR), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
    dVolt = Format(ADRead(ActNo(ACT_DA_NO).AD_VOLT), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
    dTime = ElapseTime(TM_ACT02)
    
    nRes = IIf(SetupVar.dAct02SetVolt(ACT_FINAL_POS) < (((SetupVar.dAct02SetVolt(2) - SetupVar.dAct02SetVolt(1)) / 2) + SetupVar.dAct02SetVolt(1)), 1, 2)
    
    If SetupVar.nAct02Direction = 0 And bStart Then
        Select Case nRes
            Case 1: nRes = 2
            Case 2: nRes = 1
        End Select
    End If
    
    dVoltLo = SetupVar.dAct02VoltLo(ACT_FINAL_POS)
    dVoltHi = SetupVar.dAct02VoltHi(ACT_FINAL_POS)
    dMaxTime = SetupVar.dAct02TimeHi(ACT_FINAL_POS)
    
    If dTime > 0.1 Then
        If RunVar.bDispFlash Then
            If bStart Then
                frmRun.pnlAct02Curr(nRes).Caption = Format(dCurr, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
                frmRun.pnlAct02Volt(nRes).Caption = Format(dVolt, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
                frmRun.pnlAct02Time(nRes).Caption = Format(dTime, "#0.0")
            End If
            
            If dVolt >= dVoltLo And dVolt <= dVoltHi Then
                RunVar.nActEndPosCount(ACT_DA_NO) = RunVar.nActEndPosCount(ACT_DA_NO) + 1
                
                If RunVar.nActEndPosCount(ACT_DA_NO) > SetupVar.nActEndPosCount(ACT_DA_NO) Then
                    bRes = True
                End If
            Else
                If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                    dAct02CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                    RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                End If
                
                If dCurr > SetupVar.dAct02CurrHi(nRes) Then
                    RunVar.nActPeakCurrCount(ACT_DA_NO) = RunVar.nActPeakCurrCount(ACT_DA_NO) + 1
                Else
                    RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                End If
                
                If RunVar.nActPeakCurrCount(ACT_DA_NO) > SetupVar.nActPeakCurrCount(ACT_DA_NO) Then
                    bRes = True
                End If
            End If
            
            If dTime >= dMaxTime Then
                If bStart Then
                    Call Act02Result(nRes, dCurr, dVolt, dTime)
                    
                    bRes = True
                Else
                    bRes = True
                End If
            End If
        End If
    End If
    
    GetAct02Move = bRes
End Function

Public Function Act02Test() As Boolean
    Dim bRes As Boolean
    Dim bResult As Boolean
    Dim bTestExist As Boolean
    Dim dCurr As Double
    Dim dVolt As Double
    Dim dTime As Double
    Dim i As Integer
    Dim j As Integer
    Dim dSetVolt As Double
    Dim dOldVolt As Double
    
    bRes = False
    
    If IsTestPos(TP_ACT02, POS_INIT) Then
        Call SetAct02Move(True)
        Call SetTestPos(TP_ACT02, POS_START_RUN)
    End If
    
    If IsTestPos(TP_ACT02, POS_START_RUN) Then
        If GetAct02Move(True) Then
            Call SetTestPos(TP_ACT02, POS_RUN_INIT)
        End If
    End If
    
    If IsTestPos(TP_ACT02, POS_RUN_INIT) Then
        If RunVar.bReAct02Use Then
            Call SetTestPos(TP_ACT02, POS_END)
            Exit Function
        End If
        
        RunVar.nAct02MaxLoop = 4
        
        If SetupVar.nAct02TestType = 1 Then
            RunVar.sAct02Addr = IIf(SetupVar.dAct02SetVolt(ACT_FINAL_POS) < (((SetupVar.dAct02SetVolt(2) - SetupVar.dAct02SetVolt(1)) / 2) + SetupVar.dAct02SetVolt(1)), Array(0, 1, 2, 3, 4), Array(1, 0, 2, 3, 4))
        Else
            RunVar.sAct02Addr = Array(0, 1, 2, 3, 4)
        End If
        
        For i = 0 To RunVar.nAct02MaxLoop
            If SetupVar.bAct02Use = False Then
                RunVar.sAct02Addr(i) = EMPTY_STACK_ADDR
            End If
        Next
        
        Call SetTestPos(TP_ACT02, POS_RUN)
    End If
    
    If IsTestPos(TP_ACT02, POS_RUN) Then
        bTestExist = False
        
        For i = 0 To RunVar.nAct02MaxLoop
            If RunVar.sAct02Addr(i) <> EMPTY_STACK_ADDR Then
                RunVar.nAct02Pos = RunVar.sAct02Addr(i)
                
                If SetupVar.nAct02TestType = 1 Then
                    Call OutDa(ActNo(ACT_DA_NO).DA_NO, Trim$(SetupVar.dAct02SetVolt(RunVar.nAct02Pos)), SysVar.bPercent(ActNo(ACT_DA_NO).AD_VOLT))
                Else
                    nSteppingMode = STEP_MODE_MOVE
                    bSteppingWrite(ACT_DA_NO) = True
                    
                    If RunVar.nAct02Pos > 1 Then
                        For j = RunVar.nAct02Pos To 1 Step -1
                            If Val(frmRun.pnlAct02Volt(j - 1).Caption) > 0 Then
                                dOldVolt = Val(frmRun.pnlAct02Volt(j - 1).Caption)
                                
                                GoTo FOREND
                            End If
                        Next

FOREND:
                    
                    End If
                    
                    Select Case RunVar.nAct02Pos
                        Case 0: ' stall 1
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P1
                            dSetVolt = Val("&H" & STEP_POS1 & STEP_POS2)
                        Case 3: ' stall 2
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P2
                            dSetVolt = Val("&H" & STEP_POS1 & STEP_POS2)
                        Case 4: ' final
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P1
                            dSetVolt = Abs(SetupVar.dAct02SetVolt(RunVar.nAct02Pos) - dOldVolt)
                        Case 1, 2:
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P2
                            dSetVolt = Abs(SetupVar.dAct02SetVolt(RunVar.nAct02Pos) - dOldVolt)
                    End Select
                    
                    lpWriteStepData(2) = Dec2Hex(Val2Byte(CM_LO, Abs(dSetVolt)))
                    lpWriteStepData(3) = Dec2Hex(Val2Byte(CM_HI, Abs(dSetVolt)))
                End If
                
                Call Delay(100)
                Call SetTime(TM_ACT02)
                
                RunVar.sAct02Addr(i) = EMPTY_STACK_ADDR
                RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                RunVar.nActEndPosCount(ACT_DA_NO) = 0
                RunVar.nActCount(ACT_DA_NO) = 0
                Erase dAct02CurrBuf
                
                lpReadStepData(2) = ""
                lpReadStepData(3) = ""
                lStepActData(ACT_DA_NO) = 0
                
                bTestExist = True
                Exit For
            End If
        Next
        
        If bTestExist = False Then
            Call SetTestPos(TP_ACT02, POS_END_INIT)
        End If
    End If
    
    If RunVar.nAct02Pos >= 0 And RunVar.nAct02Pos <= 10 Then
        dCurr = Format(ADRead(ActNo(ACT_DA_NO).AD_CURR), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
        
        If SetupVar.nAct02TestType = 1 Then
            dVolt = Format(ADRead(ActNo(ACT_DA_NO).AD_VOLT), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
        Else
            dVolt = Format(lStepActData(ACT_DA_NO), "0")
        End If
        
        dTime = ElapseTime(TM_ACT02)
        
        If dTime > 0.1 Then
            If RunVar.bDispFlash Then
                frmRun.pnlAct02Curr(RunVar.nAct02Pos).Caption = Format(dCurr, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
                
                If SetupVar.nAct02TestType = 1 Then
                    frmRun.pnlAct02Volt(RunVar.nAct02Pos).Caption = Format(dVolt, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
                Else
                    nSteppingMode = STEP_MODE_READ
                    dVolt = CalcStep(CDbl(lStepActData(ACT_DA_NO)), ACT_DA_NO, RunVar.nAct02Pos - 1)
                    frmRun.pnlAct02Volt(RunVar.nAct02Pos).Caption = Format(dVolt, "0")
                End If
                
                frmRun.pnlAct02Time(RunVar.nAct02Pos).Caption = Format(dTime, "#0.0")
                
                If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                    dAct02CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                    RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                End If
                
                If SetupVar.nAct02TestType = 1 Then
                    If dVolt >= SetupVar.dAct02VoltLo(RunVar.nAct02Pos) And dVolt <= SetupVar.dAct02VoltHi(RunVar.nAct02Pos) Then
                        RunVar.nActEndPosCount(ACT_DA_NO) = RunVar.nActEndPosCount(ACT_DA_NO) + 1
                        
                        If RunVar.nActEndPosCount(ACT_DA_NO) > SetupVar.nActEndPosCount(ACT_DA_NO) Then
                            bResult = True
                        End If
                    Else
                        If RunVar.nAct02Pos = ACT_STALL1 Or RunVar.nAct02Pos = ACT_STALL2 Then
                            If dCurr > SetupVar.dAct02CurrLo(RunVar.nAct02Pos) Then
                                RunVar.nActPeakCurrCount(ACT_DA_NO) = RunVar.nActPeakCurrCount(ACT_DA_NO) + 1
                            Else
                                RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                            End If
                        Else
                            If dCurr > SetupVar.dAct02CurrHi(RunVar.nAct02Pos) Then
                                RunVar.nActPeakCurrCount(ACT_DA_NO) = RunVar.nActPeakCurrCount(ACT_DA_NO) + 1
                            Else
                                RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                            End If
                        End If
                        
                        If RunVar.nActPeakCurrCount(ACT_DA_NO) > SetupVar.nActPeakCurrCount(ACT_DA_NO) Then
                            bResult = True
                        End If
                    End If
                Else
                    If RunVar.nAct02Pos = ACT_STALL1 Or RunVar.nAct02Pos = ACT_STALL2 Then
                        If dCurr > SetupVar.dAct02CurrHi(RunVar.nAct02Pos) Then
                            RunVar.nActPeakCurrCount(ACT_DA_NO) = RunVar.nActPeakCurrCount(ACT_DA_NO) + 1
                        Else
                            RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                        End If
                        
                        If RunVar.nActPeakCurrCount(ACT_DA_NO) > SetupVar.nActPeakCurrCount(ACT_DA_NO) Then
                            bResult = True
                        End If
                        
                        If bActStallBit(ACT_DA_NO) Then
                            bResult = True
                            bActStallBit(ACT_DA_NO) = False
                        End If
                    Else
                        If dVolt >= SetupVar.dAct02VoltLo(RunVar.nAct02Pos) And dVolt <= SetupVar.dAct02VoltHi(RunVar.nAct02Pos) Then
                            RunVar.nActEndPosCount(ACT_DA_NO) = RunVar.nActEndPosCount(ACT_DA_NO) + 1
                            
                            If RunVar.nActEndPosCount(ACT_DA_NO) > SetupVar.nActEndPosCount(ACT_DA_NO) Then
                                bResult = True
                            End If
                        End If
                    End If
                End If
            End If
            
            If dTime >= SetupVar.dAct02TimeHi(RunVar.nAct02Pos) Then ' Time Over
                Call OnLog("[ERROR] ACT02 TIME OVER : " & Format(dTime, "#0.00") & " !!!")
                
                bResult = True
            End If
        End If
        
        If bResult Then
            If Act02Result(RunVar.nAct02Pos, dCurr, dVolt, dTime) Then
                Call SetTestPos(TP_ACT02, POS_RUN)
            Else
'                Call SetTestPos(TP_ACT02, POS_END)
                Call SetTestPos(TP_ACT02, POS_RUN)
            End If
        End If
    End If
    
    If IsTestPos(TP_ACT02, POS_END_INIT) Then
        dVolt = Val(frmRun.pnlAct02Volt(2).Caption) - Val(frmRun.pnlAct02Volt(1).Caption)
        
        Call Act02StallDeltaResult(dVolt)
        
        If RunVar.bReAct02Use Then
            Call SetTestPos(TP_ACT02, POS_END)
        Else
            Call SetAct02Move(False)
            Call SetTestPos(TP_ACT02, POS_END_RUN)
        End If
    End If
    
    If IsTestPos(TP_ACT02, POS_END_RUN) Then
        If GetAct02Move(False) Then
            Call SetTestPos(TP_ACT02, POS_END)
        End If
    End If
    
    If IsTestPos(TP_ACT02, POS_END) Then
        Call DO_Control(ActNo(ACT_DA_NO).O_POWER, False)
        
        RunVar.bAct02Use = False
        bRes = True
    End If
    
    Act02Test = bRes
End Function

Public Function Act02Result(ByVal nPos As Integer, ByVal dCurr As Double, ByVal dVolt As Double, ByVal dTime As Double) As Boolean
    Dim lBkColor As Long
    Dim bRes(2) As Boolean
    Dim bTotalRes As Boolean
    Dim i As Integer
    Dim dRes As Double
    
    Erase bRes
    
    Select Case nPos
        Case ACT_STALL1, ACT_STALL2:
            Call DataSort(dAct02CurrBuf, RunVar.nActCount(ACT_DA_NO))
            
            dCurr = dAct02CurrBuf(0)
        Case Else:
            If RunVar.nActCount(ACT_DA_NO) > 0 Then
                For i = 0 To RunVar.nActCount(ACT_DA_NO)
                    If dAct02CurrBuf(i) >= SetupVar.dAct02CurrLo(nPos) And dAct02CurrBuf(i) <= SetupVar.dAct02CurrHi(nPos) Then
                        dRes = dRes + dAct02CurrBuf(i)
                        
                        Debug.Print dAct02CurrBuf(i)
                    Else
                        RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) - 1
                    End If
                Next
                
                If RunVar.nActCount(ACT_DA_NO) < 1 Then
                    RunVar.nActCount(ACT_DA_NO) = 1
                End If
                
                dRes = dRes / RunVar.nActCount(ACT_DA_NO)
            Else
                dRes = dCurr
            End If
            
            dCurr = Format(dRes, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
    End Select
    
    If dCurr >= SetupVar.dAct02CurrLo(nPos) And dCurr <= SetupVar.dAct02CurrHi(nPos) Then
        lBkColor = vbGreen
        bRes(0) = True
    Else
        lBkColor = vbRed
    End If
    
    frmRun.pnlAct02Curr(nPos).Caption = Format(dCurr, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
    frmRun.pnlAct02Curr(nPos).BackColor = lBkColor
    
    If SetupVar.nAct02TestType = 0 Then
        Select Case nPos
            Case 0:
                dVolt = 0
            Case Else:
        
        End Select
    End If
    
    If dVolt >= SetupVar.dAct02VoltLo(nPos) And dVolt <= SetupVar.dAct02VoltHi(nPos) Then
        lBkColor = vbGreen
        bRes(1) = True
    Else
        lBkColor = vbRed
    End If
    
    If SetupVar.nAct02TestType = 1 Then
        frmRun.pnlAct02Volt(nPos).Caption = Format(dVolt, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
    Else
        frmRun.pnlAct02Volt(nPos).Caption = Format(dVolt, "0")
    End If
    
    frmRun.pnlAct02Volt(nPos).BackColor = lBkColor
    
    If dTime >= SetupVar.dAct02TimeLo(nPos) And dTime <= SetupVar.dAct02TimeHi(nPos) Then
        lBkColor = vbGreen
        bRes(2) = True
    Else
        lBkColor = vbRed
    End If
    
    frmRun.pnlAct02Time(nPos).BackColor = lBkColor
    
    If bRes(0) And bRes(1) And bRes(2) Then
        frmRun.pnlAct02Result(nPos).Caption = "OK"
        lBkColor = vbGreen
        bTotalRes = True
    Else
        frmRun.pnlAct02Result(nPos).Caption = "NG"
        lBkColor = vbRed
        bTotalRes = False
        RunVar.bReAct02Use = True
        RunVar.bFinal = False
        
        If bRes(0) = False Then Call SetPlc(PLC_ACT02_CURR)
        If bRes(1) = False Then Call SetPlc(PLC_ACT02_VOLT)
    End If
    
    frmRun.pnlAct02Result(nPos).BackColor = lBkColor
    
    Act02Result = bTotalRes
End Function

Private Function Act02StallDeltaResult(ByVal dVolt As Double) As Boolean
    Dim lBkColor As Long
    
    If dVolt >= SetupVar.dAct02StallDeltaVoltLo And dVolt <= SetupVar.dAct02StallDeltaVoltHi Then
        lBkColor = vbGreen
    Else
        lBkColor = vbRed
        
        RunVar.bReAct02Use = True
        RunVar.bFinal = False
        
        Call SetPlc(PLC_ACT02_VOLT)
    End If
    
    frmRun.pnlAct02StallDelta.Caption = Format(dVolt, "#0.0")
    frmRun.pnlAct02StallDelta.BackColor = lBkColor
End Function

