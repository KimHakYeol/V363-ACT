Attribute VB_Name = "mdlLinAct03"
Option Explicit

Public Function LinAct03Test() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim bRes As Boolean
    Dim nRes As Integer
    Dim dRes As Double
    Dim lTmp(1) As Long
    Dim lpTmp As String
    Dim lMoveData As Long
    
    bRes = False
    
    If IsTestPos(TP_LINACT03, POS_INIT) Then
        RunVar.nAct03Count = 0
        
        Erase dAct03CurrBuf
        
        Call SetTestPos(TP_LINACT03, POS_START_RUN)
    End If
    
    If RunVar.nLinAct03Pos >= POS_START_RUN And RunVar.nLinAct03Pos < POS_END_INIT And SetupVar.nLinTestType = 1 Then
        frmRun.pnlLinActCurr(2).Caption = Format(ADRead(AD_LIN_ACT_CURR), SysVar.lpUnit(AD_LIN_ACT_CURR))
        
        If RunVar.nAct03Count < MAX_INT Then
            dAct03CurrBuf(RunVar.nAct03Count) = Val(frmRun.pnlLinActCurr(2).Caption)
            RunVar.nAct03Count = RunVar.nAct03Count + 1
        End If
    End If
    
    If IsTestPos(TP_LINACT03, POS_START_RUN) Then
        Call SetTime(TM_LINACT03)
        
        For i = 0 To 4
            RunVar.nLinAct03CheckPos(i) = SetupVar.nLinAct03Check(i)
            RunVar.dLinCPTime(2, i) = SetupVar.dLinAct03CheckTime(i)
            
            If RunVar.nLinAct03CheckPos(i) > 0 And RunVar.dLinCPTime(2, i) > 0 Then
                RunVar.bLinCheckPoint(2, i) = True
            End If
        Next
        
        Call OnLog("[LIN] ACT 03 START...")
        Call SetTestPos(TP_LINACT03, POS_RUN_INIT)
    End If
    
    If IsTestPos(TP_LINACT03, POS_RUN_INIT) Then
        Call SetTestPos(TP_LINACT03, POS_RUN_INIT + 10)
    End If
    
    ' stall 1
    If IsTestPos(TP_LINACT03, POS_RUN_INIT + 10) Then
        Call LinInit(BYTE1_ACT03)
        Call LinActWrite(8, &H22, BYTE1_ACT03, &H4, &H51, &H0, &H0, &H0, &H0, &H0)
        
        nLinReadSeq(2) = 1
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(2), i) = False
        Next
        
        Call SetTestPos(TP_LINACT03, POS_RUN)
    End If
    
    ' stall 2
    If IsTestPos(TP_LINACT03, POS_RUN_INIT + 20) Then
        Call LinInit(BYTE1_ACT03)
        Call LinActWrite(8, &H22, BYTE1_ACT03, &H4, &H51, &HFE, &HFF, &H0, &H0, &H0)
        
        nLinReadSeq(2) = 2
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(2), i) = False
        Next
        
        Call SetTestPos(TP_LINACT03, POS_RUN)
    End If
    
    ' stall result
    If IsTestPos(TP_LINACT03, POS_RUN_INIT + 25) Then
        If SetupVar.bStallUse = False Then
            If frmRun.pnlLinActOpen(2).BackColor = vbGreen Then
                frmRun.pnlLinActClose(2).BackColor = vbGreen
                frmRun.pnlLinShipp(2).BackColor = vbGreen
                frmRun.pnlLinShipp(2).Caption = "OK"
            Else
                frmRun.pnlLinActClose(2).BackColor = vbRed
                frmRun.pnlLinShipp(2).BackColor = vbRed
                frmRun.pnlLinShipp(2).Caption = "NG"
            End If
        Else
            If frmRun.pnlLinActClose(2).BackColor = vbGreen And frmRun.pnlLinActOpen(2).BackColor = vbGreen Then
                frmRun.pnlLinShipp(2).BackColor = vbGreen
                frmRun.pnlLinShipp(2).Caption = "OK"
            Else
                frmRun.pnlLinShipp(2).BackColor = vbRed
                frmRun.pnlLinShipp(2).Caption = "NG"
            End If
        End If
        
        Call SetTestPos(TP_LINACT03, POS_RUN_INIT + 26)
    End If
    
    ' move result
    If IsTestPos(TP_LINACT03, POS_RUN_INIT + 26) Then
        lMoveData = 0
        lMoveData = RunVar.lLinDataMove(2, 2) - RunVar.lLinDataMove(1, 2)
        lMoveData = lMoveData * LIN_1STEP_ANGLE
        
        frmRun.pnlLinActMove(2).Caption = Format(lMoveData, "0")
        
        Call OnLog("[" & Format(nLinReadSeq(2), "00") & "] ACT03 " & RunVar.lLinDataMove(2, 2) & " " & RunVar.lLinDataMove(1, 2) & " " & RunVar.lLinDataMove(2, 2) - RunVar.lLinDataMove(1, 2) & " " & lMoveData)
        
        If lMoveData >= SetupVar.dLinActLo(2) And lMoveData <= SetupVar.dLinActHi(2) Then
            'ok
            frmRun.pnlLinActMove(2).BackColor = vbGreen
        Else
            'ng
            frmRun.pnlLinActMove(2).BackColor = vbRed
            
            RunVar.bFinal = False
            
            Call SetPlc(PLC_ACT03_ANGLE)
            Call SetTestPos(TP_LINACT03, POS_END)
            
            Exit Function
        End If
        
        Call SetTestPos(TP_LINACT03, POS_RUN_INIT + IIf(SetupVar.bCheckPointUse, 30, 40))
    End If
    
    ' checkpoint
    If IsTestPos(TP_LINACT03, POS_RUN_INIT + 30) Then
        For i = 0 To 4
            If RunVar.bLinCheckPoint(2, i) Then
                lMoveData = RunVar.lLinDataMove(2, 2) - RunVar.lLinDataMove(1, 2)
                lMoveData = lMoveData * (RunVar.nLinAct03CheckPos(i) / 100)
                lMoveData = RunVar.lLinDataMove(1, 2) + lMoveData
                
                Call OnLog(RunVar.lLinDataMove(2, 2) & "     " & RunVar.lLinDataMove(1, 2) & "     " & RunVar.lLinDataMove(2, 2) - RunVar.lLinDataMove(1, 2) & "     " & lMoveData)
                
                lpTmp = Replace(Format(Hex$(lMoveData), "@@@@"), " ", "0")
                RunVar.lLinDataCP(0, i) = lMoveData
                
                Call OnLog("HEX : " & lpTmp & "    " & RunVar.nLinAct03CheckPos(i) & " %")
                
                lTmp(0) = Val("&H" & Left(lpTmp, 2) & "&")
                lTmp(1) = Val("&H" & Mid(lpTmp, 3) & "&")
                
                nLinTimeNo = j
                
                Exit For
            End If
        Next
        
        Call LinActWrite(8, &H22, BYTE1_ACT03, &H4, &H51, lTmp(1), lTmp(0), &H0, &H0, &H0)
        
        nLinReadSeq(2) = 3
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(2), i) = False
        Next
        
        Call SetTestPos(TP_LINACT03, POS_RUN)
    End If
    
    ' final
    If IsTestPos(TP_LINACT03, POS_RUN_INIT + 40) Then
        lMoveData = RunVar.lLinDataMove(2, 2) - RunVar.lLinDataMove(1, 2)
        lMoveData = lMoveData * (SetupVar.nLinActFinal(2) / 100)
        lMoveData = RunVar.lLinDataMove(1, 2) + lMoveData
        
        lpTmp = Replace(Format(Hex$(lMoveData), "@@@@"), " ", "0")
        RunVar.lLinDataFinal(2) = lMoveData
        
        lTmp(0) = Val("&H" & Left(lpTmp, 2) & "&")
        lTmp(1) = Val("&H" & Right(lpTmp, 2) & "&")
        
        Call LinInit(BYTE1_ACT03)
        Call LinActWrite(8, &H22, BYTE1_ACT03, &H4, &H51, lTmp(1), lTmp(0), &H0, &H0, &H0)
        
        nLinReadSeq(2) = 4
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(2), i) = False
        Next
        
        Call SetTestPos(TP_LINACT03, POS_RUN)
    End If
    
    ' final result
    If IsTestPos(TP_LINACT03, POS_RUN_INIT + 45) Then
        lMoveData = (RunVar.lLinDataMove(4, 2) - RunVar.lLinDataMove(1, 2)) * LIN_1STEP_ANGLE
        frmRun.pnlLinActFinal(2).Caption = Format(lMoveData, "0")
        
        If RunVar.nAct03Count = 0 Then
            RunVar.nAct03Count = 1
        End If
        
        For i = 0 To RunVar.nAct03Count
            dRes = dRes + dAct03CurrBuf(i)
        Next
        
        dRes = dRes / RunVar.nAct03Count
        
        frmRun.pnlLinActCurr(2).Caption = Format(dRes, SysVar.lpUnit(AD_LIN_ACT_CURR))
        
        If dRes >= SetupVar.dLinActCurrLo(2) And dRes <= SetupVar.dLinActCurrHi(2) Then
            frmRun.pnlLinActCurr(2).BackColor = vbGreen
        Else
            frmRun.pnlLinActCurr(2).BackColor = vbRed
            
            RunVar.bFinal = False
            
            Call SetPlc(PLC_ACT03_CURR)
            Call SetTestPos(TP_LINACT03, POS_END)
            
            Exit Function
        End If
        
        Call SetTestPos(TP_LINACT03, POS_RUN_INIT + 70)
    End If
    
    If IsTestPos(TP_LINACT03, POS_RUN_INIT + 70) Then
        Call SetTestPos(TP_LINACT03, POS_END_INIT)
    End If
    
    ' time
    If RunVar.nLinAct03Pos >= POS_RUN_INIT And RunVar.nLinAct03Pos <= POS_END_INIT Then
        If nLinReadSeq(2) <= 4 Then
            frmRun.pnlLinActTime(2).Caption = Format(ElapseTime(TM_LINACT03), "#0.0")
        End If
        
        If ElapseTime(TM_LINACT03) > SetupVar.dLinActTime(2) Then
            RunVar.bFinal = False
            bLinCommError = False
            
            frmRun.pnlLinActTime(2).BackColor = vbRed
            
            Call SetPlc(PLC_ACT03_ANGLE)
            Call SetTestPos(TP_LINACT03, POS_END_INIT)
        End If
    End If
    
    ' lin data read
    If IsTestPos(TP_LINACT03, POS_RUN) Then
        If ElapseTime(TM_LINACT03) * 1000 Mod 200 > 1 Then
            If RunVar.bLinAct03Use And RunVar.bLinDataResult(nLinReadSeq(2), 2) = False Then
                Call LinIDWrite(BYTE1_ACT03)
            End If
        End If
        
        Call SetTestPos(TP_LINACT03, POS_RUN + 10)
    End If
    
    ' seq next
    If IsTestPos(TP_LINACT03, POS_RUN + 10) Then
        Select Case nLinReadSeq(2)
            Case 1:
                If SetupVar.bStallUse Then
                    Call SetTestPos(TP_LINACT03, POS_RUN_INIT + 20)
                Else
                    Call SetTestPos(TP_LINACT03, POS_RUN_INIT + 25)
                End If
            
            Case 2:
                Call SetTestPos(TP_LINACT03, POS_RUN_INIT + 25)
            
            Case 3:
                Call SetTestPos(TP_LINACT03, POS_RUN_INIT + 40)
            
            Case 4:
                Call SetTestPos(TP_LINACT03, POS_RUN_INIT + 45)
        
        End Select
        
        If RunVar.bLinDataResult(nLinReadSeq(2), 2) = False Then
            Call SetTestPos(TP_LINACT03, POS_RUN)
        End If
    End If
    
    If IsTestPos(TP_LINACT03, POS_END_INIT) Then
        If bLinCommError Then
            If nLinActNo = 99 Then
                Exit Function
            End If
            
            Select Case nLinReadSeq(2)
                Case 1, 2:
                    frmRun.pnlLinShipp(2).Caption = "ERROR"
                    frmRun.pnlLinShipp(2).BackColor = vbRed
                
                Case 3:
                    frmRun.pnlLinActMove(2).Caption = "ERROR"
                    frmRun.pnlLinActMove(2).BackColor = vbRed
                
                Case 4:
                    frmRun.pnlLinActFinal(2).Caption = "ERROR"
                    frmRun.pnlLinActFinal(2).BackColor = vbRed
            
            End Select
            
            bLinCommError = False
            RunVar.bFinal = False
            
            Call SetPlc(PLC_ACT03_ANGLE)
            Call SetTestPos(TP_LINACT03, POS_END)
            
            Exit Function
        End If
        
        If _
            frmRun.pnlLinShipp(2).BackColor = vbRed Or _
            frmRun.pnlLinActResult(2).BackColor = vbRed Or _
            frmRun.pnlLinActFinal(2).BackColor = vbRed Or _
            frmRun.pnlLinActTime(2).BackColor = vbRed Then
            
            frmRun.pnlLinActFinal(2).BackColor = vbRed
            frmRun.pnlLinActTime(2).BackColor = vbRed
            frmRun.pnlLinActResult(2).BackColor = vbRed
            frmRun.pnlLinActResult(2).Caption = "NG"
        Else
            frmRun.pnlLinActFinal(2).BackColor = vbGreen
            frmRun.pnlLinActTime(2).BackColor = vbGreen
            frmRun.pnlLinActResult(2).BackColor = vbGreen
            frmRun.pnlLinActResult(2).Caption = "OK"
        End If
        
        Call LinInit(BYTE1_ACT03)
        Call SetTestPos(TP_LINACT03, POS_END)
    End If
    
    If IsTestPos(TP_LINACT03, POS_END) Then
        bRes = True
    End If
    
    LinAct03Test = bRes
End Function

