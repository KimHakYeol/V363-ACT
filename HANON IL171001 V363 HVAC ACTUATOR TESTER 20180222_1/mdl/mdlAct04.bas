Attribute VB_Name = "mdlAct04"
Option Explicit

Private Const ACT_DA_NO As Integer = 3
Private Const ACT_FINAL_POS As Integer = 4
Private Const ACT_STALL1 As Integer = 0
Private Const ACT_STALL2 As Integer = 3

Public Sub SetAct04Move(ByVal bStart As Boolean)
    Dim nRes As Integer
    
    RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
    RunVar.nActEndPosCount(ACT_DA_NO) = 0
    RunVar.nActCount(ACT_DA_NO) = 0
    Erase dAct04CurrBuf
    
    If SetupVar.nAct04TestType = 1 Then
        If SetupVar.nAct042Pin = 0 Then
            nRes = IIf(SetupVar.dAct04SetVolt(ACT_FINAL_POS) < (((SetupVar.dAct04SetVolt(1) - SetupVar.dAct04SetVolt(0)) / 2) + SetupVar.dAct04SetVolt(0)), 0, 1)
            
            If SetupVar.nAct04Direction = 0 And bStart Then
                nRes = Abs(nRes - 1)
                
                Call OutDa(ActNo(ACT_DA_NO).DA_NO, SetupVar.dAct04SetVolt(nRes), SysVar.bPercent(ActNo(ACT_DA_NO).AD_VOLT))
            Else
                Call OutDa(ActNo(ACT_DA_NO).DA_NO, SetupVar.dAct04SetVolt(ACT_FINAL_POS), SysVar.bPercent(ActNo(ACT_DA_NO).AD_VOLT))
            End If
        Else
            nRes = SetupVar.nAct042PinPos
            
            If SetupVar.nAct04Direction = 0 And bStart Then
                nRes = Abs(SetupVar.nAct042PinPos - 1)
            End If
            
            Call OutDa(ActNo(ACT_DA_NO).DA_NO, IIf(nRes = 0, 0, 5))
        End If
    End If
    
    Call DO_Control(ActNo(ACT_DA_NO).O_POWER, True)
    Call SetTime(TM_ACT04)
End Sub

Public Sub SetStepAct04Move()
    If SetupVar.nAct04TestType <> 0 Then Exit Sub
    
    Call DO_Control(ActNo(ACT_DA_NO).O_POWER, True)
    
    bSteppingWrite(3) = SetupVar.bAct04Use
    
    If SetupVar.bAct04Use Then
        lpWriteStepRotation(3) = STEP_ROTA_P2
        lpWriteStepData(6) = STEP_LITTLE1
        lpWriteStepData(7) = STEP_LITTLE2
    Else
        lpWriteStepRotation(3) = STEP_NULL
        lpWriteStepData(6) = STEP_NULL
        lpWriteStepData(7) = STEP_NULL
    End If
End Sub

Public Function GetAct04Move(ByVal bStart As Boolean) As Boolean
    Dim dCurr As Double
    Dim dVolt As Double
    Dim dTime As Double
    Dim bRes As Boolean
    Dim bNGFlag As Boolean
    Dim nRes As Integer
    Dim dMaxTime As Double
    Dim dVoltLo As Double
    Dim dVoltHi As Double
    
    If SetupVar.nAct04TestType = 0 Then
        GetAct04Move = True
        
        Exit Function
    End If
    
    bRes = False
    bNGFlag = False
    
    dCurr = Format(ADRead(ActNo(ACT_DA_NO).AD_CURR), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
    
    If SetupVar.nAct042Pin = 0 Then
        dVolt = Format(ADRead(ActNo(ACT_DA_NO).AD_VOLT), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
    End If
    
    dTime = ElapseTime(TM_ACT04)
    
    If SetupVar.nAct042Pin = 0 Then
        nRes = IIf(SetupVar.dAct04SetVolt(ACT_FINAL_POS) < (((SetupVar.dAct04SetVolt(2) - SetupVar.dAct04SetVolt(1)) / 2) + SetupVar.dAct04SetVolt(1)), 1, 2)
    Else
        nRes = SetupVar.nAct042PinPos + 1
    End If
    
    If SetupVar.nAct04Direction = 0 And bStart Then
        Select Case nRes
            Case 1: nRes = 2
            Case 2: nRes = 1
        End Select
    End If
    
    dVoltLo = SetupVar.dAct04VoltLo(ACT_FINAL_POS)
    dVoltHi = SetupVar.dAct04VoltHi(ACT_FINAL_POS)
    dMaxTime = SetupVar.dAct04TimeHi(ACT_FINAL_POS)
    
    If dTime > 0.1 Then
        If RunVar.bDispFlash Then
            If bStart Then
                frmRun.pnlAct04Curr(nRes).Caption = Format(dCurr, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
                
                If SetupVar.nAct042Pin = 0 Then
                    frmRun.pnlAct04Volt(nRes).Caption = Format(dVolt, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
                End If
                
                frmRun.pnlAct04Time(nRes).Caption = Format(dTime, "#0.0")
            End If
            
            If dVolt >= dVoltLo And dVolt <= dVoltHi Then
                RunVar.nActEndPosCount(ACT_DA_NO) = RunVar.nActEndPosCount(ACT_DA_NO) + 1
                
                If RunVar.nActEndPosCount(ACT_DA_NO) > SetupVar.nActEndPosCount(ACT_DA_NO) Then
                    bRes = True
                End If
            Else
                If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                    dAct04CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                    RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                End If
                
                If dCurr > SetupVar.dAct04CurrHi(nRes) Then
                    RunVar.nActPeakCurrCount(ACT_DA_NO) = RunVar.nActPeakCurrCount(ACT_DA_NO) + 1
                Else
                    RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                End If
                
                If RunVar.nActPeakCurrCount(ACT_DA_NO) > SetupVar.nActPeakCurrCount(ACT_DA_NO) Then
                    bRes = True
                End If
            End If
            
            If dTime >= dMaxTime Then
                bRes = True
                bNGFlag = True
            End If
            
            If bRes And bStart And bNGFlag Then
                RunVar.nActCount(ACT_DA_NO) = 0
                
                Call Act04Result(nRes, dCurr, dVolt, dTime, bNGFlag)
            End If
        End If
    End If
    
    GetAct04Move = bRes
End Function

Public Function Act04Test() As Boolean
    Dim bRes As Boolean
    Dim bResult As Boolean
    Dim bTestExist As Boolean
    Dim bNGFlag As Boolean
    Dim dCurr As Double
    Dim dVolt As Double
    Dim dTime As Double
    Dim i As Integer
    Dim j As Integer
    Dim dSetVolt As Double
    Dim dOldVolt As Double
    
    bRes = False
    bNGFlag = False
    
    If IsTestPos(TP_ACT04, POS_INIT) Then
        Call SetAct04Move(True)
        Call SetTestPos(TP_ACT04, POS_START_RUN)
    End If
    
    If IsTestPos(TP_ACT04, POS_START_RUN) Then
        If GetAct04Move(True) Then
            Call SetTestPos(TP_ACT04, POS_RUN_INIT)
        End If
    End If
    
    If IsTestPos(TP_ACT04, POS_RUN_INIT) Then
        If RunVar.bReAct04Use Then
            Call SetTestPos(TP_ACT04, POS_END)
            Exit Function
        End If
        
        If SetupVar.nAct042Pin = 0 Then
            RunVar.nAct04MaxLoop = 4
            
            If SetupVar.nAct04TestType = 1 Then
                RunVar.sAct04Addr = IIf(SetupVar.dAct04SetVolt(ACT_FINAL_POS) < (((SetupVar.dAct04SetVolt(2) - SetupVar.dAct04SetVolt(1)) / 2) + SetupVar.dAct04SetVolt(1)), Array(0, 1, 2, 3, 4), Array(1, 0, 2, 3, 4))
            Else
                RunVar.sAct04Addr = Array(0, 1, 2, 3, 4)
            End If
        Else
            RunVar.nAct04MaxLoop = 1
            
            If SetupVar.nAct042PinPos = 0 Then
                RunVar.sAct04Addr = Array(3, 0)
            Else
                RunVar.sAct04Addr = Array(0, 3)
            End If
        End If
        
        For i = 0 To RunVar.nAct04MaxLoop
            If SetupVar.bAct04Use = False Then
                RunVar.sAct04Addr(i) = EMPTY_STACK_ADDR
            End If
        Next
        
        Call SetTestPos(TP_ACT04, POS_RUN)
    End If
    
    If IsTestPos(TP_ACT04, POS_RUN) Then
        bTestExist = False
        
        For i = 0 To RunVar.nAct04MaxLoop
            If RunVar.sAct04Addr(i) <> EMPTY_STACK_ADDR Then
                RunVar.nAct04Pos = RunVar.sAct04Addr(i)
                
                If SetupVar.nAct04TestType = 1 Then
                    If SetupVar.nAct042Pin = 0 Then
                        Call OutDa(ActNo(ACT_DA_NO).DA_NO, Trim$(SetupVar.dAct04SetVolt(RunVar.nAct04Pos)), SysVar.bPercent(ActNo(ACT_DA_NO).AD_VOLT))
                    Else
                        Call OutDa(ActNo(ACT_DA_NO).DA_NO, IIf(RunVar.nAct04Pos = 0, 0, 5))
                    End If
                Else
                    nSteppingMode = STEP_MODE_MOVE
                    bSteppingWrite(ACT_DA_NO) = True
                    
                    If RunVar.nAct04Pos > 1 Then
                        For j = RunVar.nAct04Pos To 1 Step -1
                            If Val(frmRun.pnlAct04Volt(j - 1).Caption) > 0 Then
                                dOldVolt = Val(frmRun.pnlAct04Volt(j - 1).Caption)
                                
                                GoTo FOREND
                            End If
                        Next

FOREND:
                    
                    End If
                    
                    Select Case RunVar.nAct04Pos
                        Case 0: ' stall 1
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P1
                            dSetVolt = Val("&H" & STEP_POS1 & STEP_POS2)
                        Case 3: ' stall 2
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P2
                            dSetVolt = Val("&H" & STEP_POS1 & STEP_POS2)
                        Case 4: ' final
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P1
                            dSetVolt = Abs(SetupVar.dAct04SetVolt(RunVar.nAct04Pos) - dOldVolt)
                        Case 1, 2:
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P2
                            dSetVolt = Abs(SetupVar.dAct04SetVolt(RunVar.nAct04Pos) - dOldVolt)
                    End Select
                    
                    lpWriteStepData(6) = Dec2Hex(Val2Byte(CM_LO, Abs(dSetVolt)))
                    lpWriteStepData(7) = Dec2Hex(Val2Byte(CM_HI, Abs(dSetVolt)))
                End If
                
                Call Delay(100)
                Call SetTime(TM_ACT04)
                
                RunVar.sAct04Addr(i) = EMPTY_STACK_ADDR
                RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                RunVar.nActEndPosCount(ACT_DA_NO) = 0
                RunVar.nActCount(ACT_DA_NO) = 0
                Erase dAct04CurrBuf
                
                lpReadStepData(6) = ""
                lpReadStepData(7) = ""
                lStepActData(ACT_DA_NO) = 0
                
                bTestExist = True
                Exit For
            End If
        Next
        
        If bTestExist = False Then
            Call SetTestPos(TP_ACT04, POS_END_INIT)
        End If
    End If
    
    If RunVar.nAct04Pos >= 0 And RunVar.nAct04Pos <= 10 Then
        dCurr = Format(ADRead(ActNo(ACT_DA_NO).AD_CURR), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
        
        If SetupVar.nAct04TestType = 1 Then
            If SetupVar.nAct042Pin = 0 Then
                dVolt = Format(ADRead(ActNo(ACT_DA_NO).AD_VOLT), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
            End If
        Else
            dVolt = Format(lStepActData(ACT_DA_NO), "0")
        End If
        
        dTime = ElapseTime(TM_ACT04)
        
        If dTime > 0.1 Then
            If RunVar.bDispFlash Then
                frmRun.pnlAct04Curr(RunVar.nAct04Pos).Caption = Format(dCurr, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
                
                If SetupVar.nAct04TestType = 1 Then
                    If SetupVar.nAct042Pin = 0 Then
                        frmRun.pnlAct04Volt(RunVar.nAct04Pos).Caption = Format(dVolt, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
                    End If
                Else
                    nSteppingMode = STEP_MODE_READ
                    dVolt = CalcStep(CDbl(lStepActData(ACT_DA_NO)), ACT_DA_NO, RunVar.nAct04Pos - 1)
                    frmRun.pnlAct04Volt(RunVar.nAct04Pos).Caption = Format(dVolt, "0")
                End If
                
                frmRun.pnlAct04Time(RunVar.nAct04Pos).Caption = Format(dTime, "#0.0")
                
                If SetupVar.nAct04TestType = 1 Then
                    If dVolt >= SetupVar.dAct04VoltLo(RunVar.nAct04Pos) And dVolt <= SetupVar.dAct04VoltHi(RunVar.nAct04Pos) Then
                        RunVar.nActEndPosCount(ACT_DA_NO) = RunVar.nActEndPosCount(ACT_DA_NO) + 1
                        
                        If RunVar.nActEndPosCount(ACT_DA_NO) > SetupVar.nActEndPosCount(ACT_DA_NO) Then
                            bResult = True
                        End If
                    Else
                        If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                            dAct04CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                            RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                        End If
                        
                        If RunVar.nAct04Pos = ACT_STALL1 Or RunVar.nAct04Pos = ACT_STALL2 Then
                            If dCurr > SetupVar.dAct04CurrLo(RunVar.nAct04Pos) Then
                                RunVar.nActPeakCurrCount(ACT_DA_NO) = RunVar.nActPeakCurrCount(ACT_DA_NO) + 1
                            Else
                                RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                            End If
                        Else
                            If dCurr > SetupVar.dAct04CurrHi(RunVar.nAct04Pos) Then
                                RunVar.nActPeakCurrCount(ACT_DA_NO) = RunVar.nActPeakCurrCount(ACT_DA_NO) + 1
                            Else
                                RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                            End If
                        End If
                        
                        If RunVar.nActPeakCurrCount(ACT_DA_NO) > SetupVar.nActPeakCurrCount(ACT_DA_NO) Then
                            bNGFlag = True
                            bResult = True
                        End If
                    End If
                Else
                    If RunVar.nAct04Pos = 0 Or RunVar.nAct04Pos = 3 Then
                        If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                            dAct04CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                            RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                        End If
                        
                        If dCurr > SetupVar.dAct04CurrHi(RunVar.nAct04Pos) Then
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
                        If dVolt >= SetupVar.dAct04VoltLo(RunVar.nAct04Pos) And dVolt <= SetupVar.dAct04VoltHi(RunVar.nAct04Pos) Then
                            RunVar.nActEndPosCount(ACT_DA_NO) = RunVar.nActEndPosCount(ACT_DA_NO) + 1
                            
                            If RunVar.nActEndPosCount(ACT_DA_NO) > SetupVar.nActEndPosCount(ACT_DA_NO) Then
                                bResult = True
                            End If
                        Else
                            If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                                dAct04CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                                RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                            End If
                        End If
                    End If
                End If
            End If
            
            If dTime >= SetupVar.dAct04TimeHi(RunVar.nAct04Pos) Then ' Time Over
                Call OnLog("[ERROR] ACT04 TIME OVER : " & Format(dTime, "#0.00") & " !!!")
                
                bResult = True
            End If
        End If
        
        If bResult Then
            If Act04Result(RunVar.nAct04Pos, dCurr, dVolt, dTime, bNGFlag) Then
                Call SetTestPos(TP_ACT04, POS_RUN)
            Else
'                Call SetTestPos(TP_ACT04, POS_END)
                Call SetTestPos(TP_ACT04, POS_RUN)
            End If
        End If
    End If
    
    If IsTestPos(TP_ACT04, POS_END_INIT) Then
        dVolt = Val(frmRun.pnlAct04Volt(2).Caption) - Val(frmRun.pnlAct04Volt(1).Caption)
        
        Call Act04StallDeltaResult(dVolt)
        
        If RunVar.bReAct04Use Then
            Call SetTestPos(TP_ACT04, POS_END)
        Else
            Call SetAct04Move(False)
            Call SetTestPos(TP_ACT04, POS_END_RUN)
        End If
    End If
    
    If IsTestPos(TP_ACT04, POS_END_RUN) Then
        If GetAct04Move(False) Then
            Call SetTestPos(TP_ACT04, POS_END)
        End If
    End If
    
    If IsTestPos(TP_ACT04, POS_END) Then
        Call DO_Control(ActNo(ACT_DA_NO).O_POWER, False)
        RunVar.bAct04Use = False
        bRes = True
    End If
    
    Act04Test = bRes
End Function

Public Function Act04Result(ByVal nPos As Integer, ByVal dCurr As Double, ByVal dVolt As Double, ByVal dTime As Double, Optional ByVal bFlag As Boolean = False) As Boolean
    Dim lBkColor As Long
    Dim bRes(2) As Boolean
    Dim bTotalRes As Boolean
    Dim i As Integer
    Dim dRes As Double
    
    Erase bRes
    
    Select Case nPos
        Case ACT_STALL1, ACT_STALL2:
            Call DataSort(dAct04CurrBuf, RunVar.nActCount(ACT_DA_NO))
            
            dCurr = dAct04CurrBuf(0)
        Case Else:
            If RunVar.nActCount(ACT_DA_NO) > 0 Then
                If bFlag = False Or SetupVar.nAct042Pin = 1 Then
                    For i = 0 To RunVar.nActCount(ACT_DA_NO)
                        If dAct04CurrBuf(i) >= SetupVar.dAct04CurrLo(nPos) And dAct04CurrBuf(i) <= SetupVar.dAct04CurrHi(nPos) Then
                            dRes = dRes + dAct04CurrBuf(i)
                        Else
                            RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) - 1
                        End If
                    Next
                    
                    If RunVar.nActCount(ACT_DA_NO) < 1 Then
                        RunVar.nActCount(ACT_DA_NO) = 1
                    End If
                    
                    dRes = dRes / RunVar.nActCount(ACT_DA_NO)
                End If
            Else
                dRes = dCurr
            End If
            
            dCurr = Format(dRes, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
    End Select
    
    If dCurr >= SetupVar.dAct04CurrLo(nPos) And dCurr <= SetupVar.dAct04CurrHi(nPos) Then
        lBkColor = vbGreen
        bRes(0) = True
    Else
        lBkColor = vbRed
    End If
    
    frmRun.pnlAct04Curr(nPos).Caption = Format(dCurr, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
    frmRun.pnlAct04Curr(nPos).BackColor = lBkColor

CURRPASS:
    
    If SetupVar.nAct04TestType = 0 Then
        Select Case nPos
            Case 0:
                dVolt = 0
            Case Else:
        
        End Select
    End If
    
    If SetupVar.nAct042Pin = 1 Then
        bRes(1) = True
    Else
        If dVolt >= SetupVar.dAct04VoltLo(nPos) And dVolt <= SetupVar.dAct04VoltHi(nPos) Then
            lBkColor = vbGreen
            bRes(1) = True
        Else
            lBkColor = vbRed
        End If
        
        If SetupVar.nAct04TestType = 1 Then
            frmRun.pnlAct04Volt(nPos).Caption = Format(dVolt, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
        Else
            frmRun.pnlAct04Volt(nPos).Caption = Format(dVolt, "0")
        End If
        
        frmRun.pnlAct04Volt(nPos).BackColor = lBkColor
    End If
    
    If dTime >= SetupVar.dAct04TimeLo(nPos) And dTime <= SetupVar.dAct04TimeHi(nPos) Then
        lBkColor = vbGreen
        bRes(2) = True
    Else
        lBkColor = vbRed
    End If
    
    frmRun.pnlAct04Time(nPos).BackColor = lBkColor
    
    If bRes(0) And bRes(1) And bRes(2) Then
        frmRun.pnlAct04Result(nPos).Caption = "OK"
        lBkColor = vbGreen
        bTotalRes = True
    Else
        frmRun.pnlAct04Result(nPos).Caption = "NG"
        lBkColor = vbRed
        bTotalRes = False
        RunVar.bReAct04Use = True
        RunVar.bFinal = False
        
        If bRes(0) = False Then Call SetPlc(PLC_ACT04_CURR)
        If bRes(1) = False Then Call SetPlc(PLC_ACT04_VOLT)
    End If
    
    frmRun.pnlAct04Result(nPos).BackColor = lBkColor
    
    Act04Result = bTotalRes
End Function

Private Function Act04StallDeltaResult(ByVal dVolt As Double) As Boolean
    Dim lBkColor As Long
    
    If dVolt >= SetupVar.dAct04StallDeltaVoltLo And dVolt <= SetupVar.dAct04StallDeltaVoltHi Then
        lBkColor = vbGreen
    Else
        lBkColor = vbRed
        
        RunVar.bReAct04Use = True
        RunVar.bFinal = False
        
        Call SetPlc(PLC_ACT04_VOLT)
    End If
    
    frmRun.pnlAct04StallDelta.Caption = Format(dVolt, "#0.0")
    frmRun.pnlAct04StallDelta.BackColor = lBkColor
End Function

