Attribute VB_Name = "mdlLinAct01"
Option Explicit

Public Function LinAct01Test() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim bRes As Boolean
    Dim nRes As Integer
    Dim dRes As Double
    Dim lTmp(1) As Long
    Dim lpTmp As String
    Dim lMoveData As Long
    
    bRes = False
    
    If IsTestPos(TP_LINACT01, POS_INIT) Then
        RunVar.nAct01Count = 0
        
        Erase dAct01CurrBuf
        
        Call SetTestPos(TP_LINACT01, POS_START_RUN)
    End If
    
    If RunVar.nLinAct01Pos >= POS_START_RUN And RunVar.nLinAct01Pos < POS_END_INIT Then
        frmRun.pnlLinActCurr(0).Caption = Format(ADRead(AD_LIN_ACT_CURR), SysVar.lpUnit(AD_LIN_ACT_CURR))
        
        If RunVar.nAct01Count < MAX_INT Then
            dAct01CurrBuf(RunVar.nAct01Count) = Val(frmRun.pnlLinActCurr(0).Caption)
            RunVar.nAct01Count = RunVar.nAct01Count + 1
        End If
    End If
    
    If IsTestPos(TP_LINACT01, POS_START_RUN) Then
        Call SetTime(TM_LINACT01)
        
        For i = 0 To 4
            RunVar.nLinAct01CheckPos(i) = SetupVar.nLinAct01Check(i)
            RunVar.dLinCPTime(0, i) = SetupVar.dLinAct01CheckTime(i)
            
            If RunVar.nLinAct01CheckPos(i) > 0 And RunVar.dLinCPTime(0, i) > 0 Then
                RunVar.bLinCheckPoint(0, i) = True
            End If
        Next
        
        Call OnLog("[LIN] ACT 01 START...")
        Call SetTestPos(TP_LINACT01, POS_RUN_INIT)
    End If
    
    If IsTestPos(TP_LINACT01, POS_RUN_INIT) Then
        Call SetTestPos(TP_LINACT01, POS_RUN_INIT + 10)
    End If
    
    ' stall 1
    If IsTestPos(TP_LINACT01, POS_RUN_INIT + 10) Then
        Call LinInit(BYTE1_ACT01)
        Call LinActWrite(8, &H22, BYTE1_ACT01, &H4, &H51, &H0, &H0, &H0, &H0, &H0)
        
        nLinReadSeq(0) = 1
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(0), i) = False
        Next
        
        Call SetTestPos(TP_LINACT01, POS_RUN)
    End If
    
    ' stall 2
    If IsTestPos(TP_LINACT01, POS_RUN_INIT + 20) Then
        Call LinInit(BYTE1_ACT01)
        Call LinActWrite(8, &H22, BYTE1_ACT01, &H4, &H51, &HFE, &HFF, &H0, &H0, &H0)
        
        nLinReadSeq(0) = 2
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(0), i) = False
        Next
        
        Call SetTestPos(TP_LINACT01, POS_RUN)
    End If
    
    ' stall result
    If IsTestPos(TP_LINACT01, POS_RUN_INIT + 25) Then
        If SetupVar.bStallUse = False Then
            If frmRun.pnlLinActOpen(0).BackColor = vbGreen Then
                frmRun.pnlLinActClose(0).BackColor = vbGreen
                frmRun.pnlLinShipp(0).BackColor = vbGreen
                frmRun.pnlLinShipp(0).Caption = "OK"
            Else
                frmRun.pnlLinActClose(0).BackColor = vbRed
                frmRun.pnlLinShipp(0).BackColor = vbRed
                frmRun.pnlLinShipp(0).Caption = "NG"
            End If
        Else
            If frmRun.pnlLinActClose(0).BackColor = vbGreen And frmRun.pnlLinActOpen(0).BackColor = vbGreen Then
                frmRun.pnlLinShipp(0).BackColor = vbGreen
                frmRun.pnlLinShipp(0).Caption = "OK"
            Else
                frmRun.pnlLinShipp(0).BackColor = vbRed
                frmRun.pnlLinShipp(0).Caption = "NG"
            End If
        End If
        
        Call SetTestPos(TP_LINACT01, POS_RUN_INIT + 26)
    End If
    
    ' move result
    If IsTestPos(TP_LINACT01, POS_RUN_INIT + 26) Then
        lMoveData = 0
        lMoveData = RunVar.lLinDataMove(2, 0) - RunVar.lLinDataMove(1, 0)
        lMoveData = lMoveData * LIN_1STEP_ANGLE
        
        frmRun.pnlLinActMove(0).Caption = Format(lMoveData, "0")
        
        Call OnLog("[" & Format(nLinReadSeq(0), "00") & "] ACT01 " & RunVar.lLinDataMove(2, 0) & " " & RunVar.lLinDataMove(1, 0) & " " & RunVar.lLinDataMove(2, 0) - RunVar.lLinDataMove(1, 0) & " " & lMoveData)
        
        If lMoveData >= SetupVar.dLinActLo(0) And lMoveData <= SetupVar.dLinActHi(0) Then
            'ok
            frmRun.pnlLinActMove(0).BackColor = vbGreen
        Else
            'ng
            frmRun.pnlLinActMove(0).BackColor = vbRed
            
            RunVar.bFinal = False
            
            Call SetPlc(PLC_ACT01_ANGLE)
            Call SetTestPos(TP_LINACT01, POS_END)
            
            Exit Function
        End If
        
        Call SetTestPos(TP_LINACT01, POS_RUN_INIT + IIf(SetupVar.bCheckPointUse, 30, 40))
    End If
    
    ' checkpoint
    If IsTestPos(TP_LINACT01, POS_RUN_INIT + 30) Then
        For i = 0 To 4
            If RunVar.bLinCheckPoint(0, i) Then
                lMoveData = RunVar.lLinDataMove(2, 0) - RunVar.lLinDataMove(1, 0)
                lMoveData = lMoveData * (RunVar.nLinAct01CheckPos(i) / 100)
                lMoveData = RunVar.lLinDataMove(1, 0) + lMoveData
                
                lpTmp = Replace(Format(Hex$(lMoveData), "@@@@"), " ", "0")
                RunVar.lLinDataCP(0, i) = lMoveData
                
                lTmp(0) = Val("&H" & Left(lpTmp, 2) & "&")
                lTmp(1) = Val("&H" & Mid(lpTmp, 3) & "&")
                
                nLinTimeNo = j
                
                Exit For
            End If
        Next
        
        Call LinActWrite(8, &H22, BYTE1_ACT01, &H4, &H51, lTmp(1), lTmp(0), &H0, &H0, &H0)
        
        nLinReadSeq(0) = 3
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(0), i) = False
        Next
        
        Call SetTestPos(TP_LINACT01, POS_RUN)
    End If
    
    ' final
    If IsTestPos(TP_LINACT01, POS_RUN_INIT + 40) Then
        lMoveData = RunVar.lLinDataMove(2, 0) - RunVar.lLinDataMove(1, 0)
        lMoveData = lMoveData * (SetupVar.nLinActFinal(0) / 100)
        lMoveData = RunVar.lLinDataMove(1, 0) + lMoveData
        
        lpTmp = Replace(Format(Hex$(lMoveData), "@@@@"), " ", "0")
        RunVar.lLinDataFinal(0) = lMoveData
        
        lTmp(0) = Val("&H" & Left(lpTmp, 2) & "&")
        lTmp(1) = Val("&H" & Right(lpTmp, 2) & "&")
                                                                                                            
        Call LinInit(BYTE1_ACT01)
        Call LinActWrite(8, &H22, BYTE1_ACT01, &H4, &H51, lTmp(1), lTmp(0), &H0, &H0, &H0)
        
        nLinReadSeq(0) = 4
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(0), i) = False
        Next
        
        Call SetTestPos(TP_LINACT01, POS_RUN)
    End If
    
    ' final result
    If IsTestPos(TP_LINACT01, POS_RUN_INIT + 45) Then
        lMoveData = (RunVar.lLinDataMove(4, 0) - RunVar.lLinDataMove(1, 0)) * LIN_1STEP_ANGLE
        frmRun.pnlLinActFinal(0).Caption = Format(lMoveData, "0")
        
        If RunVar.nAct01Count = 0 Then
            RunVar.nAct01Count = 1
        End If
        
        For i = 0 To RunVar.nAct01Count
            dRes = dRes + dAct01CurrBuf(i)
        Next
        
        dRes = dRes / RunVar.nAct01Count
        
        frmRun.pnlLinActCurr(0).Caption = Format(dRes, SysVar.lpUnit(AD_LIN_ACT_CURR))
        
        If dRes >= SetupVar.dLinActCurrLo(0) And dRes <= SetupVar.dLinActCurrHi(0) Then
            frmRun.pnlLinActCurr(0).BackColor = vbGreen
        Else
            frmRun.pnlLinActCurr(0).BackColor = vbRed
            
            RunVar.bFinal = False
            
            Call SetPlc(PLC_ACT01_CURR)
            Call SetTestPos(TP_LINACT01, POS_END)
            
            Exit Function
        End If
        
        Call SetTestPos(TP_LINACT01, POS_RUN_INIT + 70)
    End If
    
    If IsTestPos(TP_LINACT01, POS_RUN_INIT + 70) Then
        Call SetTestPos(TP_LINACT01, POS_END_INIT)
    End If
    
    ' time
    If RunVar.nLinAct01Pos >= POS_RUN_INIT And RunVar.nLinAct01Pos <= POS_END_INIT Then
        frmRun.pnlLinActTime(0).Caption = Format(ElapseTime(TM_LINACT01), "#0.0")
        
        If ElapseTime(TM_LINACT01) > SetupVar.dLinActTime(0) Then
            RunVar.bFinal = False
            bLinCommError = False
            
            frmRun.pnlLinActTime(0).BackColor = vbRed
            
            Call SetPlc(PLC_ACT01_ANGLE)
            Call SetTestPos(TP_LINACT01, POS_END_INIT)
        End If
    End If
    
    ' lin data read
    If IsTestPos(TP_LINACT01, POS_RUN) Then
        If ElapseTime(TM_LINACT01) * 1000 Mod 200 > 1 Then
            If RunVar.bLinAct01Use And RunVar.bLinDataResult(nLinReadSeq(0), 0) = False Then
                Call LinIDWrite(BYTE1_ACT01)
            End If
        End If
        
        Call SetTestPos(TP_LINACT01, POS_RUN + 10)
    End If
    
    ' seq next
    If IsTestPos(TP_LINACT01, POS_RUN + 10) Then
        Select Case nLinReadSeq(0)
            Case 1:
                If SetupVar.bStallUse Then
                    Call SetTestPos(TP_LINACT01, POS_RUN_INIT + 20)
                Else
                    Call SetTestPos(TP_LINACT01, POS_RUN_INIT + 25)
                End If
            
            Case 2:
                Call SetTestPos(TP_LINACT01, POS_RUN_INIT + 25)
            
            Case 3:
                Call SetTestPos(TP_LINACT01, POS_RUN_INIT + 40)
            
            Case 4:
                Call SetTestPos(TP_LINACT01, POS_RUN_INIT + 45)
        
        End Select
        
        If RunVar.bLinDataResult(nLinReadSeq(0), 0) = False Then
            Call SetTestPos(TP_LINACT01, POS_RUN)
        End If
    End If
    
    If IsTestPos(TP_LINACT01, POS_END_INIT) Then
        If bLinCommError Then
            If nLinActNo = 99 Then
                Exit Function
            End If
            
            Select Case nLinReadSeq(0)
                Case 1, 2:
                    frmRun.pnlLinShipp(0).Caption = "ERROR"
                    frmRun.pnlLinShipp(0).BackColor = vbRed
                
                Case 3:
                    frmRun.pnlLinActMove(0).Caption = "ERROR"
                    frmRun.pnlLinActMove(0).BackColor = vbRed
                
                Case 4:
                    frmRun.pnlLinActFinal(0).Caption = "ERROR"
                    frmRun.pnlLinActFinal(0).BackColor = vbRed
            
            End Select
            
            bLinCommError = False
            RunVar.bFinal = False
            
            Call SetPlc(PLC_ACT01_ANGLE)
            Call SetTestPos(TP_LINACT01, POS_END)
            
            Exit Function
        End If
        
        If _
            frmRun.pnlLinShipp(0).BackColor = vbRed Or _
            frmRun.pnlLinActResult(0).BackColor = vbRed Or _
            frmRun.pnlLinActFinal(0).BackColor = vbRed Or _
            frmRun.pnlLinActTime(0).BackColor = vbRed Then
            
            frmRun.pnlLinActFinal(0).BackColor = vbRed
            frmRun.pnlLinActTime(0).BackColor = vbRed
            frmRun.pnlLinActResult(0).BackColor = vbRed
            frmRun.pnlLinActResult(0).Caption = "NG"
        Else
            frmRun.pnlLinActFinal(0).BackColor = vbGreen
            frmRun.pnlLinActTime(0).BackColor = vbGreen
            frmRun.pnlLinActResult(0).BackColor = vbGreen
            frmRun.pnlLinActResult(0).Caption = "OK"
        End If
        
        Call LinInit(BYTE1_ACT01)
        Call SetTestPos(TP_LINACT01, POS_END)
    End If
    
    If IsTestPos(TP_LINACT01, POS_END) Then
        bRes = True
    End If
    
    LinAct01Test = bRes
End Function

