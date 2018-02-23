Attribute VB_Name = "mdlAct03"
Option Explicit

Private Const ACT_DA_NO As Integer = 2
Private Const ACT_FINAL_POS As Integer = 4
Private Const ACT_STALL1 As Integer = 0
Private Const ACT_STALL2 As Integer = 3

Public Sub SetAct03Move(ByVal bStart As Boolean)
    Dim nRes As Integer
    
    RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
    RunVar.nActEndPosCount(ACT_DA_NO) = 0
    RunVar.nActCount(ACT_DA_NO) = 0
    Erase dAct03CurrBuf
    
    If SetupVar.nAct03TestType = 1 Then
        nRes = IIf(SetupVar.dAct03SetVolt(ACT_FINAL_POS) < (((SetupVar.dAct03SetVolt(1) - SetupVar.dAct03SetVolt(0)) / 2) + SetupVar.dAct03SetVolt(0)), 0, 1)
        
        If SetupVar.nAct03Direction = 0 And bStart Then
            nRes = Abs(nRes - 1)
            
            Call OutDa(ActNo(ACT_DA_NO).DA_NO, SetupVar.dAct03SetVolt(nRes), SysVar.bPercent(ActNo(ACT_DA_NO).AD_VOLT))
        Else
            Call OutDa(ActNo(ACT_DA_NO).DA_NO, SetupVar.dAct03SetVolt(ACT_FINAL_POS), SysVar.bPercent(ActNo(ACT_DA_NO).AD_VOLT))
        End If
    End If
    
    Call DO_Control(ActNo(ACT_DA_NO).O_POWER, True)
    Call SetTime(TM_ACT03)
End Sub

Public Sub SetStepAct03Move()
    If SetupVar.nAct03TestType <> 0 Then Exit Sub
    
    Call DO_Control(ActNo(ACT_DA_NO).O_POWER, True)
    
    bSteppingWrite(2) = SetupVar.bAct03Use
    
    If SetupVar.bAct03Use Then
        lpWriteStepRotation(2) = STEP_ROTA_P2
        lpWriteStepData(4) = STEP_LITTLE1
        lpWriteStepData(5) = STEP_LITTLE2
    Else
        lpWriteStepRotation(2) = STEP_NULL
        lpWriteStepData(4) = STEP_NULL
        lpWriteStepData(5) = STEP_NULL
    End If
End Sub

Public Function GetAct03Move(ByVal bStart As Boolean) As Boolean
    Dim dCurr As Double
    Dim dVolt As Double
    Dim dTime As Double
    Dim bRes As Boolean
    Dim nRes As Integer
    Dim dMaxTime As Double
    Dim dVoltLo As Double
    Dim dVoltHi As Double
    
    If SetupVar.nAct03TestType = 0 Then
        GetAct03Move = True
        
        Exit Function
    End If
    
    bRes = False
    
    dCurr = Format(ADRead(ActNo(ACT_DA_NO).AD_CURR), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
    dVolt = Format(ADRead(ActNo(ACT_DA_NO).AD_VOLT), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
    dTime = ElapseTime(TM_ACT03)
    
    nRes = IIf(SetupVar.dAct03SetVolt(ACT_FINAL_POS) < (((SetupVar.dAct03SetVolt(2) - SetupVar.dAct03SetVolt(1)) / 2) + SetupVar.dAct03SetVolt(1)), 1, 2)
    
    If SetupVar.nAct03Direction = 0 And bStart Then
        Select Case nRes
            Case 1: nRes = 2
            Case 2: nRes = 1
        End Select
    End If
    
    dVoltLo = SetupVar.dAct03VoltLo(ACT_FINAL_POS)
    dVoltHi = SetupVar.dAct03VoltHi(ACT_FINAL_POS)
    dMaxTime = SetupVar.dAct03TimeHi(ACT_FINAL_POS)
    
    If dTime > 0.1 Then
        If RunVar.bDispFlash Then
            If bStart Then
                frmRun.pnlAct03Curr(nRes).Caption = Format(dCurr, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
                frmRun.pnlAct03Volt(nRes).Caption = Format(dVolt, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
                frmRun.pnlAct03Time(nRes).Caption = Format(dTime, "#0.0")
            End If
            
            If dVolt >= dVoltLo And dVolt <= dVoltHi Then
                RunVar.nActEndPosCount(ACT_DA_NO) = RunVar.nActEndPosCount(ACT_DA_NO) + 1
                
                If RunVar.nActEndPosCount(ACT_DA_NO) > SetupVar.nActEndPosCount(ACT_DA_NO) Then
                    bRes = True
                End If
            Else
                If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                    dAct03CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                    RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                End If
                
                If dCurr > SetupVar.dAct03CurrHi(nRes) Then
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
                    Call Act03Result(nRes, dCurr, dVolt, dTime)
                    
                    bRes = True
                Else
                    bRes = True
                End If
            End If
        End If
    End If
    
    GetAct03Move = bRes
End Function

Public Function Act03Test() As Boolean
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
    
    If IsTestPos(TP_ACT03, POS_INIT) Then
        Call SetAct03Move(True)
        Call SetTestPos(TP_ACT03, POS_START_RUN)
    End If
    
    If IsTestPos(TP_ACT03, POS_START_RUN) Then
        If GetAct03Move(True) Then
            Call SetTestPos(TP_ACT03, POS_RUN_INIT)
        End If
    End If
    
    If IsTestPos(TP_ACT03, POS_RUN_INIT) Then
        If RunVar.bReAct03Use Then
            Call SetTestPos(TP_ACT03, POS_END)
            Exit Function
        End If
        
        RunVar.nAct03MaxLoop = 4
        
        If SetupVar.nAct03TestType = 1 Then
            RunVar.sAct03Addr = IIf(SetupVar.dAct03SetVolt(ACT_FINAL_POS) < (((SetupVar.dAct03SetVolt(2) - SetupVar.dAct03SetVolt(1)) / 2) + SetupVar.dAct03SetVolt(1)), Array(0, 1, 2, 3, 4), Array(1, 0, 2, 3, 4))
        Else
            RunVar.sAct03Addr = Array(0, 1, 2, 3, 4)
        End If
        
        For i = 0 To RunVar.nAct03MaxLoop
            If SetupVar.bAct03Use = False Then
                RunVar.sAct03Addr(i) = EMPTY_STACK_ADDR
            End If
        Next
        
        Call SetTestPos(TP_ACT03, POS_RUN)
    End If
    
    If IsTestPos(TP_ACT03, POS_RUN) Then
        bTestExist = False
        
        For i = 0 To RunVar.nAct03MaxLoop
            If RunVar.sAct03Addr(i) <> EMPTY_STACK_ADDR Then
                RunVar.nAct03Pos = RunVar.sAct03Addr(i)
                
                If SetupVar.nAct03TestType = 1 Then
                    Call OutDa(ActNo(ACT_DA_NO).DA_NO, Trim$(SetupVar.dAct03SetVolt(RunVar.nAct03Pos)), SysVar.bPercent(ActNo(ACT_DA_NO).AD_VOLT))
                Else
                    nSteppingMode = STEP_MODE_MOVE
                    bSteppingWrite(ACT_DA_NO) = True
                    
                    If RunVar.nAct03Pos > 1 Then
                        For j = RunVar.nAct03Pos To 1 Step -1
                            If Val(frmRun.pnlAct03Volt(j - 1).Caption) > 0 Then
                                dOldVolt = Val(frmRun.pnlAct03Volt(j - 1).Caption)
                                
                                GoTo FOREND
                            End If
                        Next

FOREND:
                    
                    End If
                    
                    Select Case RunVar.nAct03Pos
                        Case 0: ' stall 1
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P1
                            dSetVolt = Val("&H" & STEP_POS1 & STEP_POS2)
                        Case 3: ' stall 2
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P2
                            dSetVolt = Val("&H" & STEP_POS1 & STEP_POS2)
                        Case 4: ' final
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P1
                            dSetVolt = Abs(SetupVar.dAct03SetVolt(RunVar.nAct03Pos) - dOldVolt)
                        Case 1, 2:
                            lpWriteStepRotation(ACT_DA_NO) = STEP_ROTA_P2
                            dSetVolt = Abs(SetupVar.dAct03SetVolt(RunVar.nAct03Pos) - dOldVolt)
                    End Select
                    
                    lpWriteStepData(4) = Dec2Hex(Val2Byte(CM_LO, Abs(dSetVolt)))
                    lpWriteStepData(5) = Dec2Hex(Val2Byte(CM_HI, Abs(dSetVolt)))
                End If
                
                Call Delay(100)
                Call SetTime(TM_ACT03)
                
                RunVar.sAct03Addr(i) = EMPTY_STACK_ADDR
                RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                RunVar.nActEndPosCount(ACT_DA_NO) = 0
                RunVar.nActCount(ACT_DA_NO) = 0
                Erase dAct03CurrBuf
                
                lpReadStepData(4) = ""
                lpReadStepData(5) = ""
                lStepActData(ACT_DA_NO) = 0
                
                bTestExist = True
                Exit For
            End If
        Next
        
        If bTestExist = False Then
            Call SetTestPos(TP_ACT03, POS_END_INIT)
        End If
    End If
    
    If RunVar.nAct03Pos >= 0 And RunVar.nAct03Pos <= 10 Then
        dCurr = Format(ADRead(ActNo(ACT_DA_NO).AD_CURR), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
        
        If SetupVar.nAct03TestType = 1 Then
            dVolt = Format(ADRead(ActNo(ACT_DA_NO).AD_VOLT), SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
        Else
            dVolt = Format(lStepActData(ACT_DA_NO), "0")
        End If
        
        dTime = ElapseTime(TM_ACT03)
        
        If dTime > 0.1 Then
            If RunVar.bDispFlash Then
                frmRun.pnlAct03Curr(RunVar.nAct03Pos).Caption = Format(dCurr, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
                
                If SetupVar.nAct03TestType = 1 Then
                    frmRun.pnlAct03Volt(RunVar.nAct03Pos).Caption = Format(dVolt, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
                Else
                    nSteppingMode = STEP_MODE_READ
                    dVolt = CalcStep(CDbl(lStepActData(ACT_DA_NO)), ACT_DA_NO, RunVar.nAct03Pos - 1)
                    frmRun.pnlAct03Volt(RunVar.nAct03Pos).Caption = Format(dVolt, "0")
                End If
                
                frmRun.pnlAct03Time(RunVar.nAct03Pos).Caption = Format(dTime, "#0.0")
                
                If SetupVar.nAct03TestType = 1 Then
                    If dVolt >= SetupVar.dAct03VoltLo(RunVar.nAct03Pos) And dVolt <= SetupVar.dAct03VoltHi(RunVar.nAct03Pos) Then
                        RunVar.nActEndPosCount(ACT_DA_NO) = RunVar.nActEndPosCount(ACT_DA_NO) + 1
                        
                        If RunVar.nActEndPosCount(ACT_DA_NO) > SetupVar.nActEndPosCount(ACT_DA_NO) Then
                            bResult = True
                        End If
                    Else
                        If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                            dAct03CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                            RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                        End If
                        
                        If RunVar.nAct03Pos = ACT_STALL1 Or RunVar.nAct03Pos = ACT_STALL2 Then
                            If dCurr > SetupVar.dAct03CurrLo(RunVar.nAct03Pos) Then
                                RunVar.nActPeakCurrCount(ACT_DA_NO) = RunVar.nActPeakCurrCount(ACT_DA_NO) + 1
                            Else
                                RunVar.nActPeakCurrCount(ACT_DA_NO) = 0
                            End If
                        Else
                            If dCurr > SetupVar.dAct03CurrHi(RunVar.nAct03Pos) Then
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
                    If RunVar.nAct03Pos = ACT_STALL1 Or RunVar.nAct03Pos = ACT_STALL2 Then
                        If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                            dAct03CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                            RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                        End If
                        
                        If dCurr > SetupVar.dAct03CurrHi(RunVar.nAct03Pos) Then
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
                        If dVolt >= SetupVar.dAct03VoltLo(RunVar.nAct03Pos) And dVolt <= SetupVar.dAct03VoltHi(RunVar.nAct03Pos) Then
                            RunVar.nActEndPosCount(ACT_DA_NO) = RunVar.nActEndPosCount(ACT_DA_NO) + 1
                            
                            If RunVar.nActEndPosCount(ACT_DA_NO) > SetupVar.nActEndPosCount(ACT_DA_NO) Then
                                bResult = True
                            End If
                        Else
                            If RunVar.nActCount(ACT_DA_NO) < MAX_INT Then
                                dAct03CurrBuf(RunVar.nActCount(ACT_DA_NO)) = dCurr
                                RunVar.nActCount(ACT_DA_NO) = RunVar.nActCount(ACT_DA_NO) + 1
                            End If
                        End If
                    End If
                End If
            End If
            
            If dTime >= SetupVar.dAct03TimeHi(RunVar.nAct03Pos) Then ' Time Over
                Call OnLog("[ERROR] ACT03 TIME OVER : " & Format(dTime, "#0.00") & " !!!")
                
                bResult = True
            End If
        End If
        
        If bResult Then
            If Act03Result(RunVar.nAct03Pos, dCurr, dVolt, dTime) Then
                Call SetTestPos(TP_ACT03, POS_RUN)
            Else
'                Call SetTestPos(TP_ACT03, POS_END)
                Call SetTestPos(TP_ACT03, POS_RUN)
            End If
        End If
    End If
    
    If IsTestPos(TP_ACT03, POS_END_INIT) Then
        dVolt = Val(frmRun.pnlAct03Volt(2).Caption) - Val(frmRun.pnlAct03Volt(1).Caption)
        
        Call Act03StallDeltaResult(dVolt)
        
        If RunVar.bReAct03Use Then
            Call SetTestPos(TP_ACT03, POS_END)
        Else
            Call SetAct03Move(False)
            Call SetTestPos(TP_ACT03, POS_END_RUN)
        End If
    End If
    
    If IsTestPos(TP_ACT03, POS_END_RUN) Then
        If GetAct03Move(False) Then
            Call SetTestPos(TP_ACT03, POS_END)
        End If
    End If
    
    If IsTestPos(TP_ACT03, POS_END) Then
        Call DO_Control(ActNo(ACT_DA_NO).O_POWER, False)
        RunVar.bAct03Use = False
        bRes = True
    End If
    
    Act03Test = bRes
End Function

Public Function Act03Result(ByVal nPos As Integer, ByVal dCurr As Double, ByVal dVolt As Double, ByVal dTime As Double) As Boolean
    Dim lBkColor As Long
    Dim bRes(2) As Boolean
    Dim bTotalRes As Boolean
    Dim i As Integer
    Dim dRes As Double
    
    Erase bRes
    
    Select Case nPos
        Case ACT_STALL1, ACT_STALL2:
            Call DataSort(dAct03CurrBuf, RunVar.nActCount(ACT_DA_NO))
            
            dCurr = dAct03CurrBuf(0)
        Case Else:
            If RunVar.nActCount(ACT_DA_NO) > 0 Then
                For i = 0 To RunVar.nActCount(ACT_DA_NO)
                    If dAct03CurrBuf(i) >= SetupVar.dAct03CurrLo(nPos) And dAct03CurrBuf(i) <= SetupVar.dAct03CurrHi(nPos) Then
                        dRes = dRes + dAct03CurrBuf(i)
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
    
    If dCurr >= SetupVar.dAct03CurrLo(nPos) And dCurr <= SetupVar.dAct03CurrHi(nPos) Then
        lBkColor = vbGreen
        bRes(0) = True
    Else
        lBkColor = vbRed
    End If
    
    frmRun.pnlAct03Curr(nPos).Caption = Format(dCurr, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_CURR))
    frmRun.pnlAct03Curr(nPos).BackColor = lBkColor
    
    If SetupVar.nAct03TestType = 0 Then
        Select Case nPos
            Case 0:
                dVolt = 0
            Case Else:
        
        End Select
    End If
    
    If dVolt >= SetupVar.dAct03VoltLo(nPos) And dVolt <= SetupVar.dAct03VoltHi(nPos) Then
        lBkColor = vbGreen
        bRes(1) = True
    Else
        lBkColor = vbRed
    End If
    
    If SetupVar.nAct03TestType = 1 Then
        frmRun.pnlAct03Volt(nPos).Caption = Format(dVolt, SysVar.lpUnit(ActNo(ACT_DA_NO).AD_VOLT))
    Else
        frmRun.pnlAct03Volt(nPos).Caption = Format(dVolt, "0")
    End If
    
    frmRun.pnlAct03Volt(nPos).BackColor = lBkColor
    
    If dTime >= SetupVar.dAct03TimeLo(nPos) And dTime <= SetupVar.dAct03TimeHi(nPos) Then
        lBkColor = vbGreen
        bRes(2) = True
    Else
        lBkColor = vbRed
    End If
    
    frmRun.pnlAct03Time(nPos).BackColor = lBkColor
    
    If bRes(0) And bRes(1) And bRes(2) Then
        frmRun.pnlAct03Result(nPos).Caption = "OK"
        lBkColor = vbGreen
        bTotalRes = True
    Else
        frmRun.pnlAct03Result(nPos).Caption = "NG"
        lBkColor = vbRed
        bTotalRes = False
        RunVar.bReAct03Use = True
        RunVar.bFinal = False
        
        If bRes(0) = False Then Call SetPlc(PLC_ACT03_CURR)
        If bRes(1) = False Then Call SetPlc(PLC_ACT03_VOLT)
    End If
    
    frmRun.pnlAct03Result(nPos).BackColor = lBkColor
    
    Act03Result = bTotalRes
End Function

Private Function Act03StallDeltaResult(ByVal dVolt As Double) As Boolean
    Dim lBkColor As Long
    
    If dVolt >= SetupVar.dAct03StallDeltaVoltLo And dVolt <= SetupVar.dAct03StallDeltaVoltHi Then
        lBkColor = vbGreen
    Else
        lBkColor = vbRed
        
        RunVar.bReAct03Use = True
        RunVar.bFinal = False
        
        Call SetPlc(PLC_ACT03_VOLT)
    End If
    
    frmRun.pnlAct03StallDelta.Caption = Format(dVolt, "#0.0")
    frmRun.pnlAct03StallDelta.BackColor = lBkColor
End Function

