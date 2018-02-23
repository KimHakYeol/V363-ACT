Attribute VB_Name = "mdlIonizer"
Option Explicit

Public Function IonTest() As Boolean
    Dim bRes As Boolean
    Dim dCurr As Double
    Dim dVolt As Double
    Dim dTime As Double
    Dim bTestExist As Boolean
    Dim i As Integer
    Dim j As Integer
    
    bRes = False
    
    If IsTestPos(TP_ION, POS_INIT) Then
        Call DO_Control(O_ION_POWER, True)
        Call SetTestPos(TP_ION, POS_RUN)
    End If
    
    If IsTestPos(TP_ION, POS_RUN) Then
        Call DO_Control(O_ION_DIAG, False)
        
        Call SetTime(TM_ION)
        Call SetTestPos(TP_ION, 0)
    End If
    
    If IsTestPos(TP_ION, POS_RUN + 1) Then
        Call DO_Control(O_ION_DIAG, True)
        
        Call SetTime(TM_ION)
        Call SetTestPos(TP_ION, 1)
    End If
    
    If IsTestPos(TP_ION, 0) Or IsTestPos(TP_ION, 1) Then
        dCurr = ADRead(AD_ION)
        dTime = ElapseTime(TM_ION)
        
        If RunVar.bDispFlash Then
            frmRun.pnlIonData(RunVar.nIonPos).Caption = Format(dCurr, SysVar.lpUnit(AD_ION))
        End If
        
        If dTime >= 1 Then
            If IsTestPos(TP_ION, 0) Then
                Call IonResult(0, dCurr)
                
                Call SetTime(TM_ION)
                Call SetTestPos(TP_ION, POS_RUN + 1)
            End If
            
            If IsTestPos(TP_ION, 1) Then
                Call IonResult(1, dCurr)
                
                Call SetTime(TM_ION)
                Call SetTestPos(TP_ION, POS_END)
            End If
        End If
    End If
    
    If IsTestPos(TP_ION, POS_END) Then
        Call DO_Control(O_ION_POWER, False)
        Call DO_Control(O_ION_DIAG, False)
        
        bRes = True
    End If
    
    IonTest = bRes
End Function

Public Function IonResult(ByVal nPos As Integer, ByVal dValue As Double)
    Dim dLo         As Double
    Dim dHi         As Double
    Dim lBkColor    As Long
    
    dLo = SetupVar.dIonLo(nPos)
    dHi = SetupVar.dIonHi(nPos)
    
    If dValue >= dLo And dValue <= dHi Then
        lBkColor = vbGreen
    Else
        lBkColor = vbRed
        RunVar.bReIonUse = True
        RunVar.bFinal = False
    End If
        
    frmRun.pnlIonData(nPos).Caption = Format(dValue, SysVar.lpUnit(AD_ION))
    frmRun.pnlIonData(nPos).BackColor = lBkColor
End Function
