Attribute VB_Name = "mdlSensor"
Option Explicit

Public Function SensorTest(ByVal nSensorValue As Integer) As Boolean
    Dim dVolt(9) As Double
    Dim dTime As Double
    Dim bRes As Boolean
    Dim i As Integer
    
    bRes = False
    
    If RunVar.nSensorPos = POS_INIT Then
        For i = 0 To nSensorValue - 1
            If RunVar.bSensorUse(i) Then
                Call DO_Control(O_SENSOR_POWER, True)
                Call SetTime(TM_SENSOR)
                
                RunVar.nSensorPos = POS_RUN
                
                Exit For
            End If
        Next
    End If
    
    If RunVar.nSensorPos = POS_RUN Then
        dTime = ElapseTime(TM_SENSOR)
        
        For i = 0 To nSensorValue - 1
            dVolt(i) = ADRead(AD_SENSOR1 + i)
            dVolt(i) = OhmCalc(dVolt(i), AD_SENSOR1 + i)
        Next
        
        If RunVar.bDispFlash Then
            For i = 0 To nSensorValue - 1
                If RunVar.bSensorUse(i) Then frmRun.pnlSensorVolt(i).Caption = Format(dVolt(i), SysVar.lpUnit(AD_SENSOR1 + i))
            Next
        End If
        
        For i = 0 To nSensorValue - 1
            If dTime >= SetupVar.dSensorTime(i) And RunVar.bSensorUse(i) Then Call SensorResult(i, dVolt(i), AD_SENSOR1 + i)
        Next
    End If
    
    RunVar.nSensorPos = POS_END
    
    For i = 0 To nSensorValue - 1
        If RunVar.bSensorUse(i) Then
            RunVar.nSensorPos = POS_RUN
            
            Exit For
        End If
    Next
    
    If RunVar.nSensorPos = POS_END Then
        Call DO_Control(O_SENSOR_POWER, False)
        
        bRes = True
    End If
    
    SensorTest = bRes
End Function

Public Function SensorResult(ByVal nPos As Integer, ByVal dVolt As Double, ByVal nCh As Integer) As Boolean
    Dim lBkColor As Long
    Dim lpStr As String
    Dim bRes As Boolean
    
    If dVolt >= SetupVar.dSensorCurrLo(nPos) And dVolt <= SetupVar.dSensorCurrHi(nPos) Then
        bRes = True
    End If
    
    If bRes Then
        lBkColor = vbGreen
        lpStr = "OK"
    Else
        lBkColor = vbRed
        lpStr = "NG"
        
        RunVar.bReSensorUse(nPos) = True
        RunVar.bFinal = False
        
        If bRes = False Then Call SetPlc(Choose(nPos + 1, PLC_SENSOR1, PLC_SENSOR2, PLC_SENSOR3, PLC_SENSOR4, PLC_SENSOR5, PLC_SENSOR6))
    End If
    
    frmRun.pnlSensorVolt(nPos).Caption = Format(dVolt, SysVar.lpUnit(nCh))
    frmRun.pnlSensorVolt(nPos).BackColor = lBkColor
    
    RunVar.bSensorUse(nPos) = False
    
    SensorResult = bRes
End Function

Public Function OhmCalc(ByVal dVout As Double, ByVal AD_CH As Integer) As Double
    Dim dB As Double
    Dim dItotal As Double
    Dim dIB As Double
    Dim dIA As Double
    Dim dThr As Double
    
    If AD_CH = 0 Then
        OhmCalc = 0
        
        Exit Function
    End If
    
    dB = dVout * 2
    dItotal = (SetupVar.dTestVolt - dB) / 100000
    dIB = dB / 100000
    dIA = dItotal - dIB
    
    If dIA = 0 Then
        dIA = dIA + 0.00001
    End If
    
    dThr = dB / dIA
    dThr = dThr / 1000
    dThr = Abs(dThr)
    
    If dThr > 99.99 Then
        dThr = 99.99
    End If
    
    OhmCalc = Format(dThr, "0.00")
End Function

