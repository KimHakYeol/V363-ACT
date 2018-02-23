Attribute VB_Name = "mdlLinPTC"
Option Explicit

Public Function LinPtcTest() As Boolean
    Dim i As Integer
    Dim bRes As Boolean
    
    bRes = False
    
    If IsTestPos(TP_LINPTC, POS_INIT) Then
        Call SetTestPos(TP_LINPTC, POS_START_RUN)
    End If
    
    If IsTestPos(TP_LINPTC, POS_START_RUN) Then
        Call SetTime(TM_PTC)
        Call OnLog("[LIN] PTC START...")
        Call SetTestPos(TP_LINPTC, POS_RUN_INIT)
    End If
    
    If IsTestPos(TP_LINPTC, POS_RUN_INIT) Then
        nLinReadSeq(4) = 5
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(4), i) = False
        Next
        
        Call LinInit(BYTE1_PTC)
        Call SetTestPos(TP_LINPTC, POS_RUN)
    End If
    
    ' time
    If RunVar.nLinPtcPos >= POS_RUN_INIT And RunVar.nLinPtcPos <= POS_END_INIT Then
        If ElapseTime(TM_PTC) > SetupVar.dPTCTime Then
            RunVar.bFinal = False
            bLinCommError = False
            
            frmRun.pnlPTCVolt.Caption = "NG"
            frmRun.pnlPTCVolt.BackColor = vbRed
            
            Call SetPlc(PLC_PTC)
            Call LinInit(BYTE1_PTC)
            Call SetTestPos(TP_LINPTC, POS_END)
            
            Exit Function
        End If
    End If
    
    ' lin data read
    If IsTestPos(TP_LINPTC, POS_RUN) Then
        If ElapseTime(TM_PTC) * 1000 Mod 200 > 1 Then
            If RunVar.bLinDataResult(nLinReadSeq(4), 4) = False Then
                Call LinIDWrite(BYTE1_PTC)
            End If
        End If
        
        Call SetTestPos(TP_LINPTC, POS_RUN + 10)
    End If
    
    ' seq next
    If IsTestPos(TP_LINPTC, POS_RUN + 10) Then
        Call SetTestPos(TP_LINPTC, POS_END_INIT)
        
        If RunVar.bLinDataResult(nLinReadSeq(4), 4) = False Then
            Call SetTestPos(TP_LINPTC, POS_RUN)
        End If
    End If
    
    If IsTestPos(TP_LINPTC, POS_END_INIT) Then
        If bLinCommError Then
            If nLinActNo = 99 Then
                Exit Function
            End If
            
            frmRun.pnlPTCVolt.Caption = "ERROR"
            frmRun.pnlPTCVolt.BackColor = vbRed
            
            bLinCommError = False
            RunVar.bFinal = False
            
            Call SetPlc(PLC_PTC)
            Call SetTestPos(TP_LINPTC, POS_END)
            
            Exit Function
        End If
        
        If RunVar.lLinDataMove(nLinReadSeq(4), 4) > 0 Then
            ' ok
            frmRun.pnlPTCVolt.Caption = "OK"
            frmRun.pnlPTCVolt.BackColor = vbGreen
        Else
            ' ng
            frmRun.pnlPTCVolt.Caption = "NG"
            frmRun.pnlPTCVolt.BackColor = vbRed
        End If
        
        Call LinInit(BYTE1_PTC)
        Call SetTestPos(TP_LINPTC, POS_END)
    End If
    
    If IsTestPos(TP_LINPTC, POS_END) Then
        bRes = True
    End If
    
    LinPtcTest = bRes
End Function

