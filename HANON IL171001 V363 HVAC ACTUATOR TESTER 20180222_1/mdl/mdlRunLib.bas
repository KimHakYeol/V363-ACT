Attribute VB_Name = "mdlRunLib"
Option Explicit

Public Function SetTestPos(ByVal TestPos As Integer, ByVal nPos As Integer)
    Select Case (TestPos)
        Case TP_TEST: RunVar.nTestPos = nPos
        Case TP_BLOWER: RunVar.nBlowerPos = nPos
        Case TP_ACT01: RunVar.nAct01Pos = nPos
        Case TP_ACT02: RunVar.nAct02Pos = nPos
        Case TP_ACT03: RunVar.nAct03Pos = nPos
        Case TP_ACT04: RunVar.nAct04Pos = nPos
        Case TP_SENSOR: RunVar.nSensorPos = nPos
        Case TP_ION: RunVar.nIonPos = nPos
        Case TP_LINACT01: RunVar.nLinAct01Pos = nPos
        Case TP_LINACT02: RunVar.nLinAct02Pos = nPos
        Case TP_LINACT03: RunVar.nLinAct03Pos = nPos
        Case TP_LINACT04: RunVar.nLinAct04Pos = nPos
        Case TP_LINPTC: RunVar.nLinPtcPos = nPos
    End Select
End Function

Public Function IsTestPos(ByVal TestPos As Integer, ByVal nPos As Integer) As Boolean
    Dim bRes As Boolean
    
    bRes = False
    
    Select Case (TestPos)
        Case TP_TEST: If RunVar.nTestPos = nPos Then bRes = True
        Case TP_BLOWER: If RunVar.nBlowerPos = nPos Then bRes = True
        Case TP_ACT01: If RunVar.nAct01Pos = nPos Then bRes = True
        Case TP_ACT02: If RunVar.nAct02Pos = nPos Then bRes = True
        Case TP_ACT03: If RunVar.nAct03Pos = nPos Then bRes = True
        Case TP_ACT04: If RunVar.nAct04Pos = nPos Then bRes = True
        Case TP_SENSOR: If RunVar.nSensorPos = nPos Then bRes = True
        Case TP_ION: If RunVar.nIonPos = nPos Then bRes = True
        Case TP_LINACT01: If RunVar.nLinAct01Pos = nPos Then bRes = True
        Case TP_LINACT02: If RunVar.nLinAct02Pos = nPos Then bRes = True
        Case TP_LINACT03: If RunVar.nLinAct03Pos = nPos Then bRes = True
        Case TP_LINACT04: If RunVar.nLinAct04Pos = nPos Then bRes = True
        Case TP_LINPTC: If RunVar.nLinPtcPos = nPos Then bRes = True
    End Select
    
    IsTestPos = bRes
End Function
