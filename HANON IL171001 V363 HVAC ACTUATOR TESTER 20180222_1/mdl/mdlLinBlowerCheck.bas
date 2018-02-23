Attribute VB_Name = "mdlLinBlowerCheck"
Option Explicit

Public Function LinBlowerCheckTest() As Boolean
    Dim i As Integer
    Dim bRes As Boolean
    
    bRes = False
    
    If IsTestPos(TP_LINBLOWERCHECK, POS_INIT) Then
        Call SetTestPos(TP_LINBLOWERCHECK, POS_START_RUN)
    End If
    
    If IsTestPos(TP_LINBLOWERCHECK, POS_START_RUN) Then
        Call SetTime(TM_LINBLOWERCHECK)
        Call OnLog("[LIN] BLOWER CHECK START...")
        Call SetTestPos(TP_LINBLOWERCHECK, POS_RUN_INIT)
    End If
    
    If IsTestPos(TP_LINBLOWERCHECK, POS_RUN_INIT) Then
        nLinReadSeq(5) = 6
        
        For i = 0 To 9
            RunVar.bLinDataResult(nLinReadSeq(5), i) = False
        Next
        
        Call LinInit(BYTE1_BLOWER)
        Call SetTestPos(TP_LINBLOWERCHECK, POS_RUN)
    End If
    
    ' time
    If RunVar.nLinBlowerPos >= POS_RUN_INIT And RunVar.nLinBlowerPos <= POS_END_INIT Then
        If ElapseTime(TM_LINBLOWERCHECK) > SetupVar.dLinBlowerTime Then
            RunVar.bFinal = False
            bLinCommError = False
            
            frmRun.pnlLinBlowerData.Caption = "NG"
            frmRun.pnlLinBlowerData.BackColor = CO_RED
            
            Call LinInit(BYTE1_BLOWER)
            Call SetTestPos(TP_LINBLOWERCHECK, POS_END)
            
            Exit Function
        End If
    End If
    
    ' lin data read
    If IsTestPos(TP_LINBLOWERCHECK, POS_RUN) Then
        If ElapseTime(TM_LINBLOWERCHECK) * 1000 Mod 200 > 1 Then
            If RunVar.bLinDataResult(nLinReadSeq(5), 5) = False Then
                Call LinIDWrite(BYTE1_BLOWER)
            End If
        End If
        
        Call SetTestPos(TP_LINBLOWERCHECK, POS_RUN + 10)
    End If
    
    ' seq next
    If IsTestPos(TP_LINBLOWERCHECK, POS_RUN + 10) Then
        Call SetTestPos(TP_LINBLOWERCHECK, POS_END_INIT)
        
        If RunVar.bLinDataResult(nLinReadSeq(5), 5) = False Then
            Call SetTestPos(TP_LINBLOWERCHECK, POS_RUN)
        End If
    End If
    
    If IsTestPos(TP_LINBLOWERCHECK, POS_END_INIT) Then
        If bLinCommError Then
            If nLinActNo = 99 Then
                Exit Function
            End If
            
            frmRun.pnlLinBlowerData.Caption = "ERROR"
            frmRun.pnlLinBlowerData.BackColor = CO_RED
            
            bLinCommError = False
            RunVar.bFinal = False
            
            Call SetTestPos(TP_LINBLOWERCHECK, POS_END)
            
            Exit Function
        End If
        
        If RunVar.lLinDataMove(nLinReadSeq(5), 5) > 0 Then
            ' ok
            frmRun.pnlLinBlowerData.Caption = "OK"
            frmRun.pnlLinBlowerData.BackColor = CO_GREEN
        Else
            ' ng
            frmRun.pnlLinBlowerData.Caption = "NG"
            frmRun.pnlLinBlowerData.BackColor = CO_RED
        End If
        
        Call LinInit(BYTE1_BLOWER)
        Call SetTestPos(TP_LINBLOWERCHECK, POS_END)
    End If
    
    If IsTestPos(TP_LINBLOWERCHECK, POS_END) Then
        bRes = True
    End If
    
    LinBlowerCheckTest = bRes
End Function

