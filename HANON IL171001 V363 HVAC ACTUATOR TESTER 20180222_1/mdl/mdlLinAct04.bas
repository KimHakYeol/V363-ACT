Attribute VB_Name = "mdlLinAct04"
Option Explicit

Public Function LinAct04Test() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim bRes As Boolean
    Dim nRes As Integer
    Dim dRes As Double
    Dim lTmp(1) As Long
    Dim lpTmp As String
    Dim lMoveData As Long
    
    bRes = False
    
    If IsTestPos(TP_LINACT04, POS_INIT) Then
        RunVar.nAct04Count = 0
        
        Erase dAct04CurrBuf
        
        Call SetTestPos(TP_LINACT04, POS_START_RUN)
    End If
    
    If RunVar.nLinAct04Pos >= POS_START_RUN And RunVar.nLinAct04Pos < POS_END_INIT And SetupVar.nLinTestType = 1 Then
        frmRun.pnlLinActCurr(3).Caption = Format(ADRead(AD_LIN_ACT_CURR), SysVar.lpUnit(AD_LIN_ACT_CURR))
        
        If RunVar.nAct04Count < MAX_INT Then
            dAct04CurrBuf(RunVar.nAct04Count) = Val(frmRun.pnlLinActCurr(3).Caption)
            RunVar.nAct04Count = RunVar.nAct04Count + 1
        End If
    End If
    
    If IsTestPos(TP_LINACT04, POS_START_RUN) Then
        Call SetTime(TM_LINACT04)
        
        For i = 0 To 4
            RunVar.nLinAct04CheckPos(i) = SetupVar.nLinAct04Check(i)
            RunVar.dLinCPTime(3, i) = SetupVar.dLinAct04CheckTime(i)
            
            If RunVar.nLinAct04CheckPos(i) > 0 And RunVar.dLinCPTime(3, i) > 0 Then
                RunVar.bLinCheckPoint(3, i) = True
            End If
        Next
        
        Call OnLog("[LIN] ACT 04 START...")
        Call SetTestPos(TP_LINACT04, POS_RUN_INIT)
    End If
    
    If IsTestPos(TP_LINACT04, POS_RUN_INIT) Then
        Call SetTestPos(TP_LINACT04, POS_RUN_INIT + 10)
    End If
    
    ' stall 1
    If IsTestPos(TP_LINACT04, POS_RUN_INIT + 10) Then
        Call LinInit(BYTE1_ACT04)
        Call LinActWrite(8, &H22, BYTE1_ACT04, &H4, &H51, &H0, &H0, &H0, &H0, &H0)
        
        nLinReadSeq(3) = 1
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(3), i) = False
        Next
        
        Call SetTestPos(TP_LINACT04, POS_RUN)
    End If
    
    ' stall 2
    If IsTestPos(TP_LINACT04, POS_RUN_INIT + 20) Then
        Call LinInit(BYTE1_ACT04)
        Call LinActWrite(8, &H22, BYTE1_ACT04, &H4, &H51, &HFE, &HFF, &H0, &H0, &H0)
        
        nLinReadSeq(3) = 2
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(3), i) = False
        Next
        
        Call SetTestPos(TP_LINACT04, POS_RUN)
    End If
    
    ' stall result
    If IsTestPos(TP_LINACT04, POS_RUN_INIT + 25) Then
        If SetupVar.bStallUse = False Then
            If frmRun.pnlLinActOpen(3).BackColor = vbGreen Then
                frmRun.pnlLinActClose(3).BackColor = vbGreen
                frmRun.pnlLinShipp(3).BackColor = vbGreen
                frmRun.pnlLinShipp(3).Caption = "OK"
            Else
                frmRun.pnlLinActClose(3).BackColor = vbRed
                frmRun.pnlLinShipp(3).BackColor = vbRed
                frmRun.pnlLinShipp(3).Caption = "NG"
            End If
        Else
            If frmRun.pnlLinActClose(3).BackColor = vbGreen And frmRun.pnlLinActOpen(3).BackColor = vbGreen Then
                frmRun.pnlLinShipp(3).BackColor = vbGreen
                frmRun.pnlLinShipp(3).Caption = "OK"
            Else
                frmRun.pnlLinShipp(3).BackColor = vbRed
                frmRun.pnlLinShipp(3).Caption = "NG"
            End If
        End If
        
        Call SetTestPos(TP_LINACT04, POS_RUN_INIT + 26)
    End If
    
    ' move result
    If IsTestPos(TP_LINACT04, POS_RUN_INIT + 26) Then
        lMoveData = 0
        lMoveData = RunVar.lLinDataMove(2, 3) - RunVar.lLinDataMove(1, 3)
        lMoveData = lMoveData * LIN_1STEP_ANGLE
        
        frmRun.pnlLinActMove(3).Caption = Format(lMoveData, "0")
        
        Call OnLog("[" & Format(nLinReadSeq(3), "00") & "] ACT04 " & RunVar.lLinDataMove(2, 3) & " " & RunVar.lLinDataMove(1, 3) & " " & RunVar.lLinDataMove(2, 3) - RunVar.lLinDataMove(1, 3) & " " & lMoveData)
        
        If lMoveData >= SetupVar.dLinActLo(3) And lMoveData <= SetupVar.dLinActHi(3) Then
            'ok
            frmRun.pnlLinActMove(3).BackColor = vbGreen
        Else
            'ng
            frmRun.pnlLinActMove(3).BackColor = vbRed
            
            RunVar.bFinal = False
            
            Call SetPlc(PLC_ACT04_ANGLE)
            Call SetTestPos(TP_LINACT04, POS_END)
            
            Exit Function
        End If
        
        Call SetTestPos(TP_LINACT04, POS_RUN_INIT + IIf(SetupVar.bCheckPointUse, 30, 40))
    End If
    
    ' checkpoint
    If IsTestPos(TP_LINACT04, POS_RUN_INIT + 30) Then
        For i = 0 To 4
            If RunVar.bLinCheckPoint(3, i) Then
                lMoveData = RunVar.lLinDataMove(2, 3) - RunVar.lLinDataMove(1, 3)
                lMoveData = lMoveData * (RunVar.nLinAct04CheckPos(i) / 100)
                lMoveData = RunVar.lLinDataMove(1, 3) + lMoveData
                
                Call OnLog(RunVar.lLinDataMove(2, 3) & "     " & RunVar.lLinDataMove(1, 3) & "     " & RunVar.lLinDataMove(2, 3) - RunVar.lLinDataMove(1, 3) & "     " & lMoveData)
                
                lpTmp = Replace(Format(Hex$(lMoveData), "@@@@"), " ", "0")
                RunVar.lLinDataCP(0, i) = lMoveData
                
                Call OnLog("HEX : " & lpTmp & "    " & RunVar.nLinAct04CheckPos(i) & " %")
                
                lTmp(0) = Val("&H" & Left(lpTmp, 2) & "&")
                lTmp(1) = Val("&H" & Mid(lpTmp, 3) & "&")
                
                nLinTimeNo = j
                
                Exit For
            End If
        Next
        
        Call LinActWrite(8, &H22, BYTE1_ACT04, &H4, &H51, lTmp(1), lTmp(0), &H0, &H0, &H0)
        
        nLinReadSeq(3) = 3
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(3), i) = False
        Next
        
        Call SetTestPos(TP_LINACT04, POS_RUN)
    End If
    
    ' final
    If IsTestPos(TP_LINACT04, POS_RUN_INIT + 40) Then
        lMoveData = RunVar.lLinDataMove(2, 3) - RunVar.lLinDataMove(1, 3)
        lMoveData = lMoveData * (SetupVar.nLinActFinal(3) / 100)
        lMoveData = RunVar.lLinDataMove(1, 3) + lMoveData
        
        lpTmp = Replace(Format(Hex$(lMoveData), "@@@@"), " ", "0")
        RunVar.lLinDataFinal(3) = lMoveData
        
        lTmp(0) = Val("&H" & Left(lpTmp, 2) & "&")
        lTmp(1) = Val("&H" & Right(lpTmp, 2) & "&")
        
        Call LinInit(BYTE1_ACT04)
        Call LinActWrite(8, &H22, BYTE1_ACT04, &H4, &H51, lTmp(1), lTmp(0), &H0, &H0, &H0)
        
        nLinReadSeq(3) = 4
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(3), i) = False
        Next
        
        Call SetTestPos(TP_LINACT04, POS_RUN)
    End If
    
    ' final result
    If IsTestPos(TP_LINACT04, POS_RUN_INIT + 45) Then
        lMoveData = (RunVar.lLinDataMove(4, 3) - RunVar.lLinDataMove(1, 3)) * LIN_1STEP_ANGLE
        frmRun.pnlLinActFinal(3).Caption = Format(lMoveData, "0")
        
        If RunVar.nAct04Count = 0 Then
            RunVar.nAct04Count = 1
        End If
        
        For i = 0 To RunVar.nAct04Count
            dRes = dRes + dAct04CurrBuf(i)
        Next
        
        dRes = dRes / RunVar.nAct04Count
        
        frmRun.pnlLinActCurr(3).Caption = Format(dRes, SysVar.lpUnit(AD_LIN_ACT_CURR))
        
        If dRes >= SetupVar.dLinActCurrLo(3) And dRes <= SetupVar.dLinActCurrHi(3) Then
            frmRun.pnlLinActCurr(3).BackColor = vbGreen
        Else
            frmRun.pnlLinActCurr(3).BackColor = vbRed
            
            RunVar.bFinal = False
            
            Call SetPlc(PLC_ACT04_CURR)
            Call SetTestPos(TP_LINACT04, POS_END)
            
            Exit Function
        End If
        
        Call SetTestPos(TP_LINACT04, POS_RUN_INIT + 70)
    End If
    
    If IsTestPos(TP_LINACT04, POS_RUN_INIT + 70) Then
        Call SetTestPos(TP_LINACT04, POS_END_INIT)
    End If
    
    ' time
    If RunVar.nLinAct04Pos >= POS_RUN_INIT And RunVar.nLinAct04Pos <= POS_END_INIT Then
        If nLinReadSeq(3) <= 4 Then
            frmRun.pnlLinActTime(3).Caption = Format(ElapseTime(TM_LINACT04), "#0.0")
        End If
        
        If ElapseTime(TM_LINACT04) > SetupVar.dLinActTime(3) Then
            RunVar.bFinal = False
            bLinCommError = False
            
            frmRun.pnlLinActTime(3).BackColor = vbRed
            
            Call SetPlc(PLC_ACT04_ANGLE)
            Call SetTestPos(TP_LINACT04, POS_END_INIT)
        End If
    End If
    
    ' lin data read
    If IsTestPos(TP_LINACT04, POS_RUN) Then
        If ElapseTime(TM_LINACT04) * 1000 Mod 200 > 1 Then
            If RunVar.bLinAct04Use And RunVar.bLinDataResult(nLinReadSeq(3), 3) = False Then
                Call LinIDWrite(BYTE1_ACT04)
            End If
        End If
        
        Call SetTestPos(TP_LINACT04, POS_RUN + 10)
    End If
    
    ' seq next
    If IsTestPos(TP_LINACT04, POS_RUN + 10) Then
        Select Case nLinReadSeq(3)
            Case 1:
                If SetupVar.bStallUse Then
                    Call SetTestPos(TP_LINACT04, POS_RUN_INIT + 20)
                Else
                    Call SetTestPos(TP_LINACT04, POS_RUN_INIT + 25)
                End If
            
            Case 2:
                Call SetTestPos(TP_LINACT04, POS_RUN_INIT + 25)
            
            Case 3:
                Call SetTestPos(TP_LINACT04, POS_RUN_INIT + 40)
            
            Case 4:
                Call SetTestPos(TP_LINACT04, POS_RUN_INIT + 45)
        
        End Select
        
        If RunVar.bLinDataResult(nLinReadSeq(3), 3) = False Then
            Call SetTestPos(TP_LINACT04, POS_RUN)
        End If
    End If
    
    If IsTestPos(TP_LINACT04, POS_END_INIT) Then
        If bLinCommError Then
            If nLinActNo = 99 Then
                Exit Function
            End If
            
            Select Case nLinReadSeq(3)
                Case 1, 2:
                    frmRun.pnlLinShipp(3).Caption = "ERROR"
                    frmRun.pnlLinShipp(3).BackColor = vbRed
                
                Case 3:
                    frmRun.pnlLinActMove(3).Caption = "ERROR"
                    frmRun.pnlLinActMove(3).BackColor = vbRed
                
                Case 4:
                    frmRun.pnlLinActFinal(3).Caption = "ERROR"
                    frmRun.pnlLinActFinal(3).BackColor = vbRed
            
            End Select
            
            bLinCommError = False
            RunVar.bFinal = False
            
            Call SetPlc(PLC_ACT04_ANGLE)
            Call SetTestPos(TP_LINACT04, POS_END)
            
            Exit Function
        End If
        
        If _
            frmRun.pnlLinShipp(3).BackColor = vbRed Or _
            frmRun.pnlLinActResult(3).BackColor = vbRed Or _
            frmRun.pnlLinActFinal(3).BackColor = vbRed Or _
            frmRun.pnlLinActTime(3).BackColor = vbRed Then
            
            frmRun.pnlLinActFinal(3).BackColor = vbRed
            frmRun.pnlLinActTime(3).BackColor = vbRed
            frmRun.pnlLinActResult(3).BackColor = vbRed
            frmRun.pnlLinActResult(3).Caption = "NG"
        Else
            frmRun.pnlLinActFinal(3).BackColor = vbGreen
            frmRun.pnlLinActTime(3).BackColor = vbGreen
            frmRun.pnlLinActResult(3).BackColor = vbGreen
            frmRun.pnlLinActResult(3).Caption = "OK"
        End If
        
        Call LinInit(BYTE1_ACT04)
        Call SetTestPos(TP_LINACT04, POS_END)
    End If
    
    If IsTestPos(TP_LINACT04, POS_END) Then
        bRes = True
    End If
    
    LinAct04Test = bRes
End Function

