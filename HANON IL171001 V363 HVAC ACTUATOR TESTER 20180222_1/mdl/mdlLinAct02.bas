Attribute VB_Name = "mdlLinAct02"
Option Explicit

Public Function LinAct02Test() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim bRes As Boolean
    Dim nRes As Integer
    Dim dRes As Double
    Dim lTmp(1) As Long
    Dim lpTmp As String
    Dim lMoveData As Long
    
    bRes = False
    
    If IsTestPos(TP_LINACT02, POS_INIT) Then
        RunVar.nAct02Count = 0
        
        Erase dAct02CurrBuf
        
        Call SetTestPos(TP_LINACT02, POS_START_RUN)
    End If
    
    If RunVar.nLinAct02Pos >= POS_START_RUN And RunVar.nLinAct02Pos < POS_END_INIT And SetupVar.nLinTestType = 1 Then
        frmRun.pnlLinActCurr(1).Caption = Format(ADRead(AD_LIN_ACT_CURR), SysVar.lpUnit(AD_LIN_ACT_CURR))
        
        If RunVar.nAct02Count < MAX_INT Then
            dAct02CurrBuf(RunVar.nAct02Count) = Val(frmRun.pnlLinActCurr(1).Caption)
            RunVar.nAct02Count = RunVar.nAct02Count + 1
        End If
    End If
    
    If IsTestPos(TP_LINACT02, POS_START_RUN) Then
        Call SetTime(TM_LINACT02)
        
        For i = 0 To 4
            RunVar.nLinAct02CheckPos(i) = SetupVar.nLinAct02Check(i)
            RunVar.dLinCPTime(1, i) = SetupVar.dLinAct02CheckTime(i)
            
            If RunVar.nLinAct02CheckPos(i) > 0 And RunVar.dLinCPTime(1, i) > 0 Then
                RunVar.bLinCheckPoint(1, i) = True
            End If
        Next
        
        Call OnLog("[LIN] ACT 02 START...")
        Call SetTestPos(TP_LINACT02, POS_RUN_INIT)
    End If
    
    If IsTestPos(TP_LINACT02, POS_RUN_INIT) Then
        Call SetTestPos(TP_LINACT02, POS_RUN_INIT + 10)
    End If
    
    ' stall 1
    If IsTestPos(TP_LINACT02, POS_RUN_INIT + 10) Then
        Call LinInit(BYTE1_ACT02)
        Call LinActWrite(8, &H22, BYTE1_ACT02, &H4, &H51, &H0, &H0, &H0, &H0, &H0)
        
        nLinReadSeq(1) = 1
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(1), i) = False
        Next
        
        Call SetTestPos(TP_LINACT02, POS_RUN)
    End If
    
    ' stall 2
    If IsTestPos(TP_LINACT02, POS_RUN_INIT + 20) Then
        Call LinInit(BYTE1_ACT02)
        Call LinActWrite(8, &H22, BYTE1_ACT02, &H4, &H51, &HFE, &HFF, &H0, &H0, &H0)
        
        nLinReadSeq(1) = 2
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(1), i) = False
        Next
        
        Call SetTestPos(TP_LINACT02, POS_RUN)
    End If
    
    ' stall result
    If IsTestPos(TP_LINACT02, POS_RUN_INIT + 25) Then
        If SetupVar.bStallUse = False Then
            If frmRun.pnlLinActOpen(1).BackColor = vbGreen Then
                frmRun.pnlLinActClose(1).BackColor = vbGreen
                frmRun.pnlLinShipp(1).BackColor = vbGreen
                frmRun.pnlLinShipp(1).Caption = "OK"
            Else
                frmRun.pnlLinActClose(1).BackColor = vbRed
                frmRun.pnlLinShipp(1).BackColor = vbRed
                frmRun.pnlLinShipp(1).Caption = "NG"
            End If
        Else
            If frmRun.pnlLinActClose(1).BackColor = vbGreen And frmRun.pnlLinActOpen(1).BackColor = vbGreen Then
                frmRun.pnlLinShipp(1).BackColor = vbGreen
                frmRun.pnlLinShipp(1).Caption = "OK"
            Else
                frmRun.pnlLinShipp(1).BackColor = vbRed
                frmRun.pnlLinShipp(1).Caption = "NG"
            End If
        End If
        
        Call SetTestPos(TP_LINACT02, POS_RUN_INIT + 26)
    End If
    
    ' move result
    If IsTestPos(TP_LINACT02, POS_RUN_INIT + 26) Then
        lMoveData = 0
        lMoveData = RunVar.lLinDataMove(2, 1) - RunVar.lLinDataMove(1, 1)
        lMoveData = lMoveData * LIN_1STEP_ANGLE
        
        frmRun.pnlLinActMove(1).Caption = Format(lMoveData, "0")
        
        Call OnLog("[" & Format(nLinReadSeq(1), "00") & "] ACT02 " & RunVar.lLinDataMove(2, 1) & " " & RunVar.lLinDataMove(1, 1) & " " & RunVar.lLinDataMove(2, 1) - RunVar.lLinDataMove(1, 1) & " " & lMoveData)
        
        If lMoveData >= SetupVar.dLinActLo(1) And lMoveData <= SetupVar.dLinActHi(1) Then
            'ok
            frmRun.pnlLinActMove(1).BackColor = vbGreen
        Else
            'ng
            frmRun.pnlLinActMove(1).BackColor = vbRed
            
            RunVar.bFinal = False
            
            Call SetPlc(PLC_ACT02_ANGLE)
            Call SetTestPos(TP_LINACT02, POS_END)
            
            Exit Function
        End If
        
        Call SetTestPos(TP_LINACT02, POS_RUN_INIT + IIf(SetupVar.bCheckPointUse, 30, 40))
    End If
    
    ' checkpoint
    If IsTestPos(TP_LINACT02, POS_RUN_INIT + 30) Then
        For i = 0 To 4
            If RunVar.bLinCheckPoint(1, i) Then
                lMoveData = RunVar.lLinDataMove(2, 1) - RunVar.lLinDataMove(1, 1)
                lMoveData = lMoveData * (RunVar.nLinAct02CheckPos(i) / 100)
                lMoveData = RunVar.lLinDataMove(1, 1) + lMoveData
                
                Call OnLog(RunVar.lLinDataMove(2, 1) & "     " & RunVar.lLinDataMove(1, 1) & "     " & RunVar.lLinDataMove(2, 1) - RunVar.lLinDataMove(1, 1) & "     " & lMoveData)
                
                lpTmp = Replace(Format(Hex$(lMoveData), "@@@@"), " ", "0")
                RunVar.lLinDataCP(0, i) = lMoveData
                
                Call OnLog("HEX : " & lpTmp & "    " & RunVar.nLinAct02CheckPos(i) & " %")
                
                lTmp(0) = Val("&H" & Left(lpTmp, 2) & "&")
                lTmp(1) = Val("&H" & Mid(lpTmp, 3) & "&")
                
                nLinTimeNo = j
                
                Exit For
            End If
        Next
        
        Call LinActWrite(8, &H22, BYTE1_ACT02, &H4, &H51, lTmp(1), lTmp(0), &H0, &H0, &H0)
        
        nLinReadSeq(1) = 3
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(1), i) = False
        Next
        
        Call SetTestPos(TP_LINACT02, POS_RUN)
    End If
    
    ' final
    If IsTestPos(TP_LINACT02, POS_RUN_INIT + 40) Then
        lMoveData = RunVar.lLinDataMove(2, 1) - RunVar.lLinDataMove(1, 1)
        lMoveData = lMoveData * (SetupVar.nLinActFinal(1) / 100)
        lMoveData = RunVar.lLinDataMove(1, 1) + lMoveData
        
        lpTmp = Replace(Format(Hex$(lMoveData), "@@@@"), " ", "0")
        RunVar.lLinDataFinal(1) = lMoveData
        
        lTmp(0) = Val("&H" & Left(lpTmp, 2) & "&")
        lTmp(1) = Val("&H" & Right(lpTmp, 2) & "&")
        
        Call LinInit(BYTE1_ACT02)
        Call LinActWrite(8, &H22, BYTE1_ACT02, &H4, &H51, lTmp(1), lTmp(0), &H0, &H0, &H0)
        
        nLinReadSeq(1) = 4
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(1), i) = False
        Next
        
        Call SetTestPos(TP_LINACT02, POS_RUN)
    End If
    
    ' final result
    If IsTestPos(TP_LINACT02, POS_RUN_INIT + 45) Then
        lMoveData = (RunVar.lLinDataMove(4, 1) - RunVar.lLinDataMove(1, 1)) * LIN_1STEP_ANGLE
        frmRun.pnlLinActFinal(1).Caption = Format(lMoveData, "0")
        
        If RunVar.nAct02Count = 0 Then
            RunVar.nAct02Count = 1
        End If
        
        For i = 0 To RunVar.nAct02Count
            dRes = dRes + dAct02CurrBuf(i)
        Next
        
        dRes = dRes / RunVar.nAct02Count
        
        frmRun.pnlLinActCurr(1).Caption = Format(dRes, SysVar.lpUnit(AD_LIN_ACT_CURR))
        
        If dRes >= SetupVar.dLinActCurrLo(1) And dRes <= SetupVar.dLinActCurrHi(1) Then
            frmRun.pnlLinActCurr(1).BackColor = vbGreen
        Else
            frmRun.pnlLinActCurr(1).BackColor = vbRed
            
            RunVar.bFinal = False
            
            Call SetPlc(PLC_ACT02_CURR)
            Call SetTestPos(TP_LINACT02, POS_END)
            
            Exit Function
        End If
        
        Call SetTestPos(TP_LINACT02, POS_RUN_INIT + 70)
    End If
    
    If IsTestPos(TP_LINACT02, POS_RUN_INIT + 70) Then
        Call SetTestPos(TP_LINACT02, POS_END_INIT)
    End If
    
    ' time
    If RunVar.nLinAct02Pos >= POS_RUN_INIT And RunVar.nLinAct02Pos <= POS_END_INIT Then
        If nLinReadSeq(1) <= 4 Then
            frmRun.pnlLinActTime(1).Caption = Format(ElapseTime(TM_LINACT02), "#0.0")
        End If
        
        If ElapseTime(TM_LINACT02) > SetupVar.dLinActTime(1) Then
            RunVar.bFinal = False
            bLinCommError = False
            
            frmRun.pnlLinActTime(1).BackColor = vbRed
            
            Call SetPlc(PLC_ACT02_ANGLE)
            Call SetTestPos(TP_LINACT02, POS_END_INIT)
        End If
    End If
    
    ' lin data read
    If IsTestPos(TP_LINACT02, POS_RUN) Then
        If ElapseTime(TM_LINACT02) * 1000 Mod 200 > 1 Then
            If RunVar.bLinAct02Use And RunVar.bLinDataResult(nLinReadSeq(1), 1) = False Then
                Call LinIDWrite(BYTE1_ACT02)
            End If
        End If
        
        Call SetTestPos(TP_LINACT02, POS_RUN + 10)
    End If
    
    ' seq next
    If IsTestPos(TP_LINACT02, POS_RUN + 10) Then
        Select Case nLinReadSeq(1)
            Case 1:
                If SetupVar.bStallUse Then
                    Call SetTestPos(TP_LINACT02, POS_RUN_INIT + 20)
                Else
                    Call SetTestPos(TP_LINACT02, POS_RUN_INIT + 25)
                End If
            
            Case 2:
                Call SetTestPos(TP_LINACT02, POS_RUN_INIT + 25)
            
            Case 3:
                Call SetTestPos(TP_LINACT02, POS_RUN_INIT + 40)
            
            Case 4:
                Call SetTestPos(TP_LINACT02, POS_RUN_INIT + 45)
        
        End Select
        
        If RunVar.bLinDataResult(nLinReadSeq(1), 1) = False Then
            Call SetTestPos(TP_LINACT02, POS_RUN)
        End If
    End If
    
    If IsTestPos(TP_LINACT02, POS_END_INIT) Then
        If bLinCommError Then
            If nLinActNo = 99 Then
                Exit Function
            End If
            
            Select Case nLinReadSeq(1)
                Case 1, 2:
                    frmRun.pnlLinShipp(1).Caption = "ERROR"
                    frmRun.pnlLinShipp(1).BackColor = vbRed
                
                Case 3:
                    frmRun.pnlLinActMove(1).Caption = "ERROR"
                    frmRun.pnlLinActMove(1).BackColor = vbRed
                
                Case 4:
                    frmRun.pnlLinActFinal(1).Caption = "ERROR"
                    frmRun.pnlLinActFinal(1).BackColor = vbRed
            
            End Select
            
            bLinCommError = False
            RunVar.bFinal = False
            
            Call SetPlc(PLC_ACT01_ANGLE)
            Call SetTestPos(TP_LINACT02, POS_END)
            
            Exit Function
        End If
        
        If _
            frmRun.pnlLinShipp(1).BackColor = vbRed Or _
            frmRun.pnlLinActResult(1).BackColor = vbRed Or _
            frmRun.pnlLinActFinal(1).BackColor = vbRed Or _
            frmRun.pnlLinActTime(1).BackColor = vbRed Then
            
            frmRun.pnlLinActFinal(1).BackColor = vbRed
            frmRun.pnlLinActTime(1).BackColor = vbRed
            frmRun.pnlLinActResult(1).BackColor = vbRed
            frmRun.pnlLinActResult(1).Caption = "NG"
        Else
            frmRun.pnlLinActFinal(1).BackColor = vbGreen
            frmRun.pnlLinActTime(1).BackColor = vbGreen
            frmRun.pnlLinActResult(1).BackColor = vbGreen
            frmRun.pnlLinActResult(1).Caption = "OK"
        End If
        
        Call LinInit(BYTE1_ACT02)
        Call SetTestPos(TP_LINACT02, POS_END)
    End If
    
    If IsTestPos(TP_LINACT02, POS_END) Then
        bRes = True
    End If
    
    LinAct02Test = bRes
End Function

