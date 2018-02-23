Attribute VB_Name = "mdlAct01"
Option Explicit

Private Const ACT_DA_NO As Integer = 0
Private Const ACT_FINAL_POS As Integer = 7
Private Const ACT_STALL1 As Integer = 0
Private Const ACT_STALL2 As Integer = 6

Public Sub SetAct01Move()
    RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
    RunVar.nActEndPosCount(ACT_DA_NO) = 0
    RunVar.nActCount(ACT_DA_NO) = 0
    Erase dAct01CurrBuf
    
    If SetupVar.nAct01TestType = 1 Then
        Call OutDa(ActNo(ACT_DA_NO).DA_NO, SetupVar.dAct01SetVolt(ACT_FINAL_POS), SysVar.bPercent(ActNo(ACT_DA_NO).AD_VOLT))
    End If
    
    Call DO_Control(ActNo(ACT_DA_NO).O_POWER, True)
    Call SetTime(TM_ACT01)
End Sub

Public Sub SetStepAct01Move()
    If SetupVar.nAct01TestType <> 0 Then Exit Sub
    
    Call DO_Control(ActNo(ACT_DA_NO).O_POWER, True)
    
    bSteppingWrite(0) = SetupVar.bAct01Use
    
    If SetupVar.bAct01Use Then
        lpWriteStepRotation(0) = STEP_ROTA_P2
        lpWriteStepData(0) = STEP_LITTLE1
        lpWriteStepData(1) = STEP_LITTLE2
    Else
        lpWriteStepRotation(0) = STEP_NULL
        lpWriteStepData(0) = STEP_NULL
        lpWriteStepData(1) = STEP_NULL
    End If
End Sub

Public Function GetAct01Move(ByVal bStart As Boolean) As Boolean
    Dim dCurr As Double
    Dim dVolt As Double
    Dim dTime As Double
    Dim bRes As Boolean
    Dim nRes As Integer
    Dim dMaxTime As Double
    Dim dVoltLo As Double
    Dim dVoltHi As Double
    
    bRes = False
    
    If SetupVar.nAct01TestType = 0 Then
        If ElapseTime(TM_ACT01) > 1 Then
            Call SetTime(TM_ACT01)
            
            GetAct01Move = True
        End If
        
        Exit Function
    End If
    
    dCurr = Format(ADRead(ActNo(0).AD_CURR), SysVar.lpUnit(ActNo(0).AD_CURR))
    dVolt = Format(ADRead(ActNo(0).AD_VOLT), SysVar.lpUnit(ActNo(0).AD_VOLT))
    dTime = ElapseTime(TM_ACT01)
    
    nRes = IIf(SetupVar.dAct01SetVolt(ACT_FINAL_POS) < (((SetupVar.dAct01SetVolt(2) - SetupVar.dAct01SetVolt(1)) / 2) + SetupVar.dAct01SetVolt(1)), 1, 2)
    
    If SetupVar.nAct01Direction = 0 And bStart Then
        Select Case nRes
            Case 1: nRes = 2
            Case 2: nRes = 1
        End Select
    End If
    
    dVoltLo = SetupVar.dAct01VoltLo(ACT_FINAL_POS)
    dVoltHi = SetupVar.dAct01VoltHi(ACT_FINAL_POS)
    dMaxTime = SetupVar.dAct01TimeHi(ACT_FINAL_POS)
    
    If dTime > 0.1 Then
        If RunVar.bDispFlash Then
            If bStart Then
                frmRun.pnlAct01Curr(nRes).Caption = Format(dCurr, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
                frmRun.pnlAct01Volt(nRes).Caption = Format(dVolt, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
                frmRun.pnlAct01Time(nRes).Caption = Format(dTime, "#0.0")
            End If
            
            If dVolt >= dVoltLo And dVolt <= dVoltHi Then
                RunVar.nActEndPosCount(ACT_DA_NO) = RunVar.nActEndPosCount(ACT_DA_NO) + 1
                
                If RunVar.nActEndPosCount(ACT_DA_NO) > SetupVar.nActEndPosCount(ACT_DA_NO) Then
                    bRes = True
                End If
            Else
                If RunVar.nActCount(ACT_DA_NO) < MAX_INT - 1 Then
                    dAct01CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                    RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                End If
                
                If dCurr > SetupVar.dAct01CurrHi(nRes) Then
                    If RunVar.nActPeakCurrCount(ACT_DA_NO) < MAX_INT - 1 Then
                        RunVar.nActPeakCurrCount(ACT_DA_NO) = RunVar.nActPeakCurrCount(ACT_DA_NO) + 1
                    End If
                Else
                    RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                End If
                
                If RunVar.nActPeakCurrCount(ACT_DA_NO) > SetupVar.nActPeakCurrCount(ACT_DA_NO) Then
                    bRes = True
                End If
            End If
            
            If dTime >= dMaxTime Then
                If bStart Then
                    Call Act01Result(nRes, dCurr, dVolt, dTime)
                    
                    bRes = True
                Else
                    bRes = True
                End If
            End If
        End If
    End If
    
    GetAct01Move = bRes
End Function

Public Function Act01Test() As Boolean
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
    
    If IsTestPos(TP_ACT01, POS_INIT) Then
        Call NvhSend(NVHACT)
        Call SetAct01Move
        Call SetTestPos(TP_ACT01, POS_START_RUN)
    End If
    
    If IsTestPos(TP_ACT01, POS_START_RUN) Then
        If GetAct01Move(True) Then
            Call SetTestPos(TP_ACT01, POS_RUN_INIT)
        End If
    End If
    
    If IsTestPos(TP_ACT01, POS_RUN_INIT) Then
        If RunVar.bReAct01Use Then
            Call SetTestPos(TP_ACT01, POS_END)
            Exit Function
        End If
        
        RunVar.nAct01MaxLoop = 7
        
        If SetupVar.nAct02TestType = 1 Then
            If SetupVar.dAct01SetVolt(ACT_FINAL_POS) < SetupVar.dAct01SetVolt(1) Then
                RunVar.sAct01Addr = Array(2, 3, 4, 5, 6, 5, 4, 3, 2, 1, 0, 7)
            ElseIf SetupVar.dAct01SetVolt(ACT_FINAL_POS) < SetupVar.dAct01SetVolt(2) Then
                RunVar.sAct01Addr = Array(3, 4, 5, 6, 5, 4, 3, 2, 1, 0, 1, 7)
            ElseIf SetupVar.dAct01SetVolt(ACT_FINAL_POS) < SetupVar.dAct01SetVolt(3) Then
                RunVar.sAct01Addr = Array(4, 5, 6, 5, 4, 3, 2, 1, 0, 1, 2, 7)
            ElseIf SetupVar.dAct01SetVolt(ACT_FINAL_POS) < SetupVar.dAct01SetVolt(4) Then
                RunVar.sAct01Addr = Array(5, 6, 5, 4, 3, 2, 1, 0, 1, 2, 3, 7)
            Else
                RunVar.sAct01Addr = Array(4, 3, 2, 1, 0, 1, 2, 3, 4, 5, 6, 7)
            End If
        Else
            RunVar.sAct01Addr = Array(0, 1, 2, 3, 4, 5, 6, 7)
        End If
        
        ' 이름이 비었을 경우 띄위 테스트하기 위함
        For i = 0 To RunVar.nAct01MaxLoop
            If Len(Trim$(SetupVar.lpAct01Name(i))) = 0 Then
                For j = 0 To RunVar.nAct01MaxLoop
                    If RunVar.sAct01Addr(j) = i Then
                        RunVar.sAct01Addr(j) = EMPTY_STACK_ADDR
                    End If
                Next
            End If
        Next
        
        For i = 0 To RunVar.nAct01MaxLoop
            If SetupVar.bAct01Use = False Then
                RunVar.sAct01Addr(i) = EMPTY_STACK_ADDR
            End If
        Next
        
        Call SetTestPos(TP_ACT01, POS_RUN)
    End If
    
    If IsTestPos(TP_ACT01, POS_RUN) Then
        bTestExist = False
        
        For i = 0 To RunVar.nAct01MaxLoop
            If RunVar.sAct01Addr(i) <> EMPTY_STACK_ADDR Then
                RunVar.nAct01Pos = RunVar.sAct01Addr(i)
                
                If SetupVar.nAct01TestType = 1 Then
                    Call OutDa(ActNo(ACT_DA_NO).DA_NO, Trim$(SetupVar.dAct01SetVolt(RunVar.nAct01Pos)), SysVar.bPercent(ActNo(ACT_DA_NO).AD_VOLT))
                Else
                    nSteppingMode = STEP_MODE_MOVE
                    bSteppingWrite(ACT_DA_NO) = True
                    
                    If RunVar.nAct01Pos > 1 Then
                        For j = RunVar.nAct01Pos To 1 Step -1
                            If Val(frmRun.pnlAct01Volt(j - 1).Caption) > 0 Then
                                dOldVolt = Val(frmRun.pnlAct01Volt(j - 1).Caption)
                                
                                GoTo FOREND
                            End If
                        Next

FOREND:
                    
                    End If
                    
                    Select Case RunVar.nAct01Pos
                        Case 0: ' stall 1
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P1
                            dSetVolt = Val("&H" & STEP_POS1 & STEP_POS2)
                        Case 6: ' stall 2
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P2
                            dSetVolt = Val("&H" & STEP_POS1 & STEP_POS2)
                        Case 7: ' final
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P1
                            dSetVolt = Abs(SetupVar.dAct01SetVolt(RunVar.nAct01Pos) - dOldVolt)
                        Case 1, 2, 3, 4, 5:
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P2
                            dSetVolt = Abs(SetupVar.dAct01SetVolt(RunVar.nAct01Pos) - dOldVolt)
                    End Select
                    
                    lpWriteStepData(0) = Dec2Hex(Val2Byte(CM_LO, Abs(dSetVolt)))
                    lpWriteStepData(1) = Dec2Hex(Val2Byte(CM_HI, Abs(dSetVolt)))
                End If
                
                Call Delay(100)
                Call SetTime(TM_ACT01)
                
                RunVar.sAct01Addr(i) = EMPTY_STACK_ADDR
                RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                RunVar.nActEndPosCount(ACT_DA_NO) = 0
                RunVar.nActCount(ACT_DA_NO) = 0
                Erase dAct01CurrBuf
                
                lpReadStepData(0) = ""
                lpReadStepData(1) = ""
                lStepActData(ACT_DA_NO) = 0
                
                bTestExist = True
                Exit For
            End If
        Next
        
        If bTestExist = False Then
            Call SetTestPos(TP_ACT01, POS_END_INIT)
        End If
    End If
    
    If RunVar.nAct01Pos >= 0 And RunVar.nAct01Pos <= RunVar.nAct01MaxLoop Then
        dCurr = Format(ADRead(ActNo(ACT_DA_NO).AD_CURR), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
        
        If SetupVar.nAct01TestType = 1 Then
            dVolt = Format(ADRead(ActNo(ACT_DA_NO).AD_VOLT), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
        Else
            dVolt = Format(lStepActData(ACT_DA_NO), "0")
        End If
        
        dTime = ElapseTime(TM_ACT01)
        
        If dTime > 0.1 Then
            If RunVar.bDispFlash Then
                frmRun.pnlAct01Curr(RunVar.nAct01Pos).Caption = Format(dCurr, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
                
                If SetupVar.nAct01TestType = 1 Then
                    frmRun.pnlAct01Volt(RunVar.nAct01Pos).Caption = Format(dVolt, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
                Else
                    nSteppingMode = STEP_MODE_READ
                    dVolt = CalcStep(CDbl(lStepActData(ACT_DA_NO)), ACT_DA_NO, RunVar.nAct01Pos - 1)
                    frmRun.pnlAct01Volt(RunVar.nAct01Pos).Caption = Format(dVolt, "0")
                End If
                
                frmRun.pnlAct01Time(RunVar.nAct01Pos).Caption = Format(dTime, "#0.0")
                
                If SetupVar.nAct01TestType = 1 Then
                    If dVolt >= SetupVar.dAct01VoltLo(RunVar.nAct01Pos) And dVolt <= SetupVar.dAct01VoltHi(RunVar.nAct01Pos) Then
                        RunVar.nActEndPosCount(ACT_DA_NO) = RunVar.nActEndPosCount(ACT_DA_NO) + 1
                        
                        If RunVar.nActEndPosCount(ACT_DA_NO) > SetupVar.nActEndPosCount(ACT_DA_NO) Then
                            bResult = True
                        End If
                    Else
                        If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                            dAct01CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                            RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                        End If
                        
                        If RunVar.nAct01Pos = ACT_STALL1 Or RunVar.nAct01Pos = ACT_STALL2 Then
                            If dCurr > SetupVar.dAct01CurrLo(RunVar.nAct01Pos) Then
                                RunVar.nActPeakCurrCount(ACT_DA_NO) = RunVar.nActPeakCurrCount(ACT_DA_NO) + 1
                            Else
                                RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                            End If
                        Else
                            If dCurr > SetupVar.dAct01CurrHi(RunVar.nAct01Pos) Then
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
                    If RunVar.nAct01Pos = ACT_STALL1 Or RunVar.nAct01Pos = ACT_STALL2 Then
                        If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                            dAct01CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                            RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                        End If
                        
                        If dCurr > SetupVar.dAct01CurrHi(RunVar.nAct01Pos) Then
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
                        If dVolt >= SetupVar.dAct01VoltLo(RunVar.nAct01Pos) And dVolt <= SetupVar.dAct01VoltHi(RunVar.nAct01Pos) Then
                            RunVar.nActEndPosCount(ACT_DA_NO) = RunVar.nActEndPosCount(ACT_DA_NO) + 1
                            
                            If RunVar.nActEndPosCount(ACT_DA_NO) > SetupVar.nActEndPosCount(ACT_DA_NO) Then
                                bResult = True
                            End If
                        Else
                            If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                                dAct01CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                                RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                            End If
                        End If
                    End If
                End If
            End If
            
            If dTime >= SetupVar.dAct01TimeHi(RunVar.nAct01Pos) Then ' Time Over
                Call OnLog("[ERROR] ACT01 TIME OVER : " & Format(dTime, "#0.00") & " !!!")
                
                bResult = True
            End If
        End If
        
        If bResult Then
            If Act01Result(RunVar.nAct01Pos, dCurr, dVolt, dTime) Then
                Call SetTestPos(TP_ACT01, POS_RUN)
            Else
'                Call SetTestPos(TP_ACT01, POS_END)
                Call SetTestPos(TP_ACT01, POS_RUN)
            End If
        End If
    End If
    
    If IsTestPos(TP_ACT01, POS_END_INIT) Then
        dVolt = Val(frmRun.pnlAct01Volt(5).Caption) - Val(frmRun.pnlAct01Volt(1).Caption)
        
        Call Act01StallDeltaResult(dVolt)
        
        If RunVar.bReAct01Use Then
            Call SetTestPos(TP_ACT01, POS_END)
        Else
            Call SetAct01Move
            Call SetTestPos(TP_ACT01, POS_END_RUN)
        End If
    End If
    
    If IsTestPos(TP_ACT01, POS_END_RUN) Then
        If GetAct01Move(False) Then
            Call SetTestPos(TP_ACT01, POS_END)
        End If
    End If
    
    If IsTestPos(TP_ACT01, POS_END) Then
        Call NvhSend(NVHSTOP)
        Call DO_Control(ActNo(ACT_DA_NO).O_POWER, False)
        
        RunVar.bAct01Use = False
        bRes = True
    End If
    
    Act01Test = bRes
End Function

Public Function Act01Result(ByVal nPos As Integer, ByVal dCurr As Double, ByVal dVolt As Double, ByVal dTime As Double) As Boolean
    Dim lBkColor As Long
    Dim bRes(2) As Boolean
    Dim bTotalRes As Boolean
    Dim i As Integer
    Dim dRes As Double
    
    Erase bRes
    
    Select Case nPos
        Case ACT_STALL1, ACT_STALL2:
            Call DataSort(dAct01CurrBuf, RunVar.nActCount(ACT_DA_NO))
            
            dCurr = dAct01CurrBuf(0)
        Case Else:
            If RunVar.nActCount(ACT_DA_NO) > 0 Then
                For i = 0 To RunVar.nActCount(ACT_DA_NO)
                    If dAct01CurrBuf(i) >= SetupVar.dAct01CurrLo(nPos) And dAct01CurrBuf(i) <= SetupVar.dAct01CurrHi(nPos) Then
                        dRes = dRes + dAct01CurrBuf(i)
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
    
    If dCurr >= SetupVar.dAct01CurrLo(nPos) And dCurr <= SetupVar.dAct01CurrHi(nPos) Then
        lBkColor = vbGreen
        bRes(0) = True
    Else
        lBkColor = vbRed
    End If
    
    frmRun.pnlAct01Curr(nPos).Caption = Format(dCurr, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
    frmRun.pnlAct01Curr(nPos).BackColor = lBkColor
    
    If SetupVar.nAct01TestType = 0 Then
        Select Case nPos
            Case 0:
                dVolt = 0
            Case Else:
        
        End Select
    End If
    
    If dVolt >= SetupVar.dAct01VoltLo(nPos) And dVolt <= SetupVar.dAct01VoltHi(nPos) Then
        lBkColor = vbGreen
        bRes(1) = True
    Else
        lBkColor = vbRed
    End If
    
    If SetupVar.nAct01TestType = 1 Then
        frmRun.pnlAct01Volt(nPos).Caption = Format(dVolt, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
    Else
        frmRun.pnlAct01Volt(nPos).Caption = Format(dVolt, "0")
    End If
    
    frmRun.pnlAct01Volt(nPos).BackColor = lBkColor
    
    If dTime >= SetupVar.dAct01TimeLo(nPos) And dTime <= SetupVar.dAct01TimeHi(nPos) Then
        lBkColor = vbGreen
        bRes(2) = True
    Else
        lBkColor = vbRed
    End If
    
    frmRun.pnlAct01Time(nPos).BackColor = lBkColor
    
    If bRes(0) And bRes(1) And bRes(2) Then
        frmRun.pnlAct01Result(nPos).Caption = "OK"
        lBkColor = vbGreen
        bTotalRes = True
    Else
        frmRun.pnlAct01Result(nPos).Caption = "NG"
        lBkColor = vbRed
        
        bTotalRes = False
        RunVar.bReAct01Use = True
        RunVar.bFinal = False
        
        If bRes(0) = False Then Call SetPlc(PLC_ACT01_CURR)
        If bRes(1) = False Then Call SetPlc(PLC_ACT01_VOLT)
    End If
    
    frmRun.pnlAct01Result(nPos).BackColor = lBkColor
    
    Act01Result = bTotalRes
End Function

Private Function Act01StallDeltaResult(ByVal dVolt As Double) As Boolean
    Dim lBkColor As Long
    
    If dVolt >= SetupVar.dAct01StallDeltaVoltLo And dVolt <= SetupVar.dAct01StallDeltaVoltHi Then
        lBkColor = vbGreen
    Else
        lBkColor = vbRed
        
        RunVar.bReAct01Use = True
        RunVar.bFinal = False
        
        Call SetPlc(PLC_ACT01_VOLT)
    End If
    
    frmRun.pnlAct01StallDelta.Caption = Format(dVolt, "#0.0")
    frmRun.pnlAct01StallDelta.BackColor = lBkColor
End Function

