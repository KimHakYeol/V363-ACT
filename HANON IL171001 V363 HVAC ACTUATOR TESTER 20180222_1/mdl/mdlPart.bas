Attribute VB_Name = "mdlPart"
Option Explicit

Public Sub PartVisible(ByVal bVisible As Boolean)
    Dim i As Integer
    Dim bRes(MAX_DIO_CHANNEL) As Boolean
    Dim bProductRes(MAX_DIO_CHANNEL) As Boolean
    Dim btnCtl As BHImageButton
    Dim lpStr() As String
    
    DataVar.lpPartOK = ""
    DataVar.lpPartNG = ""
    
    For Each btnCtl In frmRun.btnDI
        bRes(btnCtl.Index) = True
    Next
    
    If SetupVar.bProductUse Then
        lpStr = Split(SetupVar.lpProductList, ",")
        
        For i = LBound(lpStr) To UBound(lpStr)
            bProductRes(Val(lpStr(i))) = True
        Next
    End If
    
    If SetupVar.bModelTypeUse Then
        lpStr = Split(Switch(SetupVar.nModelType = 0, SetupVar.lpModelLHDList, SetupVar.nModelType = 1, SetupVar.lpModelRHDList), ",")
        
        For i = LBound(lpStr) To UBound(lpStr)
            bProductRes(Val(lpStr(i))) = True
        Next
    End If
    
    For i = 0 To MAX_DIO_CHANNEL
        If bRes(i) And bProductRes(i) = False Then
            Select Case i
                Case I_WORK_ON, I_WORK_OFF, I_MARKING1_ON, I_MARKING1_OFF:
                    frmRun.btnDI(i).Visible = True
                    frmRun.btnDI(i).ZOrder 0
                
                Case Else:
                    If SetupVar.bPartUse(i) Then
                        frmRun.btnDI(i).Visible = True
                        frmRun.btnDI(i).BackColor = CO_NONE
                        frmRun.btnDI(i).ZOrder 0
                    Else
                        frmRun.btnDI(i).Visible = bVisible
                        frmRun.btnDI(i).ZOrder 0
                    End If
                    
                    If SetupVar.lpPartName(i) <> "" And SetupVar.lpPartName(i) <> "#" Then frmRun.btnDI(i).Caption = Trim$(SetupVar.lpPartName(i))
            
            End Select
        End If
        
        If bRes(i) And bProductRes(i) Then
            frmRun.btnDI(i).Visible = True
            frmRun.btnDI(i).BackColor = CO_NONE
            frmRun.btnDI(i).ZOrder 0
        End If
    Next
End Sub

Public Function PartDIS(ByVal nCh As Integer) As Boolean
    Dim bRes As Boolean
    
    If SetupVar.bPartStatus(nCh) Then   ' ON Check
        bRes = DIS(nCh)
    Else                                ' OFF Check
        bRes = Not DIS(nCh)
    End If
    
    PartDIS = bRes
End Function

Public Function PartCheck(ByVal nStartPos As Integer, ByVal nEndPos As Integer, ByVal lpStr As String) As Boolean
    Dim bRes    As Boolean
    Dim i       As Integer
    
    bRes = True
    
    For i = nStartPos To nEndPos
        DoEvents
        
        If SetupVar.bPartUse(i) Then
            frmRun.btnDI(i).Visible = True
            frmRun.btnDI(i).ZOrder 0
            
            If PartDIS(i) Then
                frmRun.btnDI(i).BackColor = vbGreen
                DataVar.lpPartOK = Trim$(DataVar.lpPartOK) + str(i) + "/"
            Else
                frmRun.btnDI(i).BackColor = vbRed
                DataVar.lpPartNG = Trim$(DataVar.lpPartNG) + str(i) + "/"
                
                Call OnLog((Space(1) & Trim$(str(i)) & Space(4 - Len(Trim$(str(i))))))
                bRes = False
            End If
        End If
    Next
    
    If bRes = False Then
        Call OnLog(lpStr)
    End If
    
    PartCheck = bRes
End Function

Public Function ModelTypeCheck() As Boolean
    Dim lpStr() As String
    Dim nRes As Integer
    Dim bPartRes(MAX_DIO_CHANNEL) As Boolean
    Dim i As Integer
    Dim lpModelType As String
    
    ModelTypeCheck = True
    
    Select Case SetupVar.nModelType
        Case 0: If SetupVar.lpModelLHDList = "" Or SetupVar.bModelTypeUse = False Then Exit Function
        Case 1: If SetupVar.lpModelRHDList = "" Or SetupVar.bModelTypeUse = False Then Exit Function
    End Select
    
    lpModelType = Switch(SetupVar.nModelType = 0, "LHD", SetupVar.nModelType = 1, "RHD")
    
    Call OnLog("MODEL " & lpModelType & " CHECK...")
    
    lpStr = Split(Switch(SetupVar.nModelType = 0, SetupVar.lpModelLHDList, SetupVar.nModelType = 1, SetupVar.lpModelRHDList), ",")
    
    For i = LBound(lpStr) To UBound(lpStr)
        If Left$(lpStr(i), 1) = "#" Then
            nRes = Val(Mid(lpStr(i), 2))
            SetupVar.bPartStatus(nRes) = False
        Else
            nRes = Val(lpStr(i))
            SetupVar.bPartStatus(nRes) = True
        End If
        
        frmRun.btnDI(nRes).Visible = True
        frmRun.btnDI(nRes).ZOrder 0
        
        If PartDIS(nRes) Then
            frmRun.btnDI(nRes).BackColor = vbGreen
            DataVar.lpPartOK = Trim$(DataVar.lpPartOK) & CStr(nRes) & "/"
            
            bPartRes(i) = True
        Else
            frmRun.btnDI(nRes).BackColor = vbRed
            DataVar.lpPartNG = Trim$(DataVar.lpPartNG) & CStr(nRes) & "/"
            
            Call OnLog((Space(1) & CStr(nRes) & Space(4 - Len(CStr(nRes)))))
            bPartRes(i) = False
        End If
    Next
    
    For i = LBound(lpStr) To UBound(lpStr)
        If bPartRes(i) = False Then
            ModelTypeCheck = False
            RunVar.bFinal = False
            
            Exit For
        End If
    Next
End Function

Public Function ProductPartCheck(Optional bMessage As Boolean = True, Optional bTextColor As Boolean = True) As Boolean
    Dim lpStr() As String
    Dim nRes As Integer
    Dim bPartRes(MAX_DIO_CHANNEL) As Boolean
    Dim i As Integer
    
    ProductPartCheck = True
    
    If SetupVar.lpProductList = "" Or SetupVar.bProductUse = False Then
        Exit Function
    End If
    
    If bMessage Then
        Call OnLog("PRODUCT CHECK PART...")
    End If
    
    lpStr = Split(SetupVar.lpProductList, ",")
    
    For i = LBound(lpStr) To UBound(lpStr)
        If Left$(lpStr(i), 1) = "#" Then
            nRes = Val(Mid(lpStr(i), 2))
            SetupVar.bPartStatus(nRes) = False
        Else
            nRes = Val(lpStr(i))
            SetupVar.bPartStatus(nRes) = True
        End If
        
        frmRun.btnDI(nRes).Visible = True
        frmRun.btnDI(nRes).ZOrder 0
        
        If PartDIS(nRes) Then
            frmRun.btnDI(nRes).BackColor = vbGreen
            
            bPartRes(i) = True
        Else
            frmRun.btnDI(nRes).BackColor = vbRed
            
            If bMessage Then
                Call OnLog((Space(1) & CStr(nRes) & Space(4 - Len(CStr(nRes)))))
            End If
            
            bPartRes(i) = False
        End If
    Next
    
    For i = LBound(lpStr) To UBound(lpStr)
        If bPartRes(i) = False Then
            ProductPartCheck = False
            
            If bMessage Then
                RunVar.bFinal = False
            End If
            
            Exit For
        End If
    Next
End Function

Public Function StartPartCheck() As Boolean
    Dim lpStr() As String
    Dim bPartRes(MAX_DIO_CHANNEL) As Boolean
    Dim bRes As Boolean
    Dim i As Integer
    Dim lpRes As String
    
    bRes = True
    
    Call OnLog("START CHECK PART...")
    
    If SetupVar.bProductUse Then ' 제품 감지 제외
        lpStr = Split(SetupVar.lpProductList, ",")
        
        For i = LBound(lpStr) To UBound(lpStr)
            bPartRes(Val(lpStr(i))) = True
        Next
    End If
    
    If SetupVar.bModelTypeUse Then ' 모델 감지 제외
        lpRes = SetupVar.lpModelLHDList & "," & SetupVar.lpModelRHDList
        lpStr = Split(lpRes, ",")
        
        For i = LBound(lpStr) To UBound(lpStr)
            bPartRes(Val(lpStr(i))) = True
        Next
    End If
    
    For i = 0 To MAX_DIO_CHANNEL
        If RunVar.bDoorUse(i) Then bPartRes(i) = True ' 도어 체크 제외
        
        If bPartRes(i) = False Then ' start parts pass
            If SetupVar.bPartUse(i) And SetupVar.lpPartName(i) = "#" Then
                bPartRes(i) = PartCheck(i, i, "Part Check...")
                
                If bPartRes(i) = False Then
                    bRes = False
                    
                    Exit For
                End If
            End If
        End If
    Next
    
    If bRes Then
        StartPartCheck = True
    Else
        StartPartCheck = False
        RunVar.bFinal = False
    End If
End Function

Public Sub OnPartCheck()
    Dim lpStr() As String
    Dim bPartRes(MAX_DIO_CHANNEL) As Boolean
    Dim bRes As Boolean
    Dim i As Integer
    Dim lpRes As String
    
    bRes = True
    
    Call OnLog("ON CHECK PART...")
    
    If SetupVar.bProductUse Then ' 제품 감지 제외
        lpStr = Split(SetupVar.lpProductList, ",")
        
        For i = LBound(lpStr) To UBound(lpStr)
            bPartRes(Val(lpStr(i))) = True
        Next
    End If
    
    If SetupVar.bModelTypeUse Then ' 모델 감지 제외
        lpRes = SetupVar.lpModelLHDList & "," & SetupVar.lpModelRHDList
        lpStr = Split(lpRes, ",")
        
        For i = LBound(lpStr) To UBound(lpStr)
            bPartRes(Val(lpStr(i))) = True
        Next
    End If
    
    For i = 0 To MAX_DIO_CHANNEL
        If RunVar.bDoorUse(i) Then bPartRes(i) = True ' 도어 체크 제외
        
        If bPartRes(i) = False Then ' start parts pass
            If SetupVar.bPartUse(i) And SetupVar.lpPartName(i) = "" Then
                bPartRes(i) = PartCheck(i, i, "Part Check...")
                
                If bPartRes(i) = False Then
                    bRes = False
                End If
            End If
        End If
    Next
    
    If bRes = False Then
        RunVar.bFinal = False
    End If
End Sub

Public Function DoorTest(ByVal nSelect As Integer, ByVal nDoorNo As Integer)
    If nDoorNo = 999 Then Exit Function
    
    Select Case nSelect
        Case 0: ' start
            If RunVar.bDoorUse(nDoorNo) Then RunVar.bDoorStatus(nDoorNo) = DIS(nDoorNo)
        
        Case 1: ' check
            frmRun.btnDI(nDoorNo).BackColor = IIf(DIS(nDoorNo), vbGreen, CO_NONE)
            
            If RunVar.bDoorStatus(nDoorNo) <> DIS(nDoorNo) And RunVar.bDoorResult(nDoorNo) = False Then
                RunVar.bDoorResult(nDoorNo) = True
            End If
        
        Case 2: ' result
            If RunVar.bDoorUse(nDoorNo) And SetupVar.bPartUse(nDoorNo) Then
                If RunVar.bDoorResult(nDoorNo) Then
                    frmRun.btnDI(nDoorNo).BackColor = vbGreen
                    frmRun.pnlDoorResult.BackColor = vbGreen
                Else
                    frmRun.btnDI(nDoorNo).BackColor = vbRed
                    frmRun.pnlDoorResult.BackColor = vbRed
                    
                    RunVar.bReAct04Use = True
                    RunVar.bFinal = False
                End If
            End If
        
        Case 9: ' check & ok result
            If RunVar.bDoorUse(nDoorNo) And RunVar.bDoorResult(nDoorNo) = False Then
                If RunVar.bDoorStatus(nDoorNo) <> DIS(nDoorNo) Then
                    frmRun.btnDI(nDoorNo).BackColor = vbGreen
                    frmRun.pnlDoorResult.BackColor = vbGreen
                    
                    RunVar.bDoorResult(nDoorNo) = True
                    RunVar.bDoorUse(nDoorNo) = False
                End If
            End If
        
    End Select
End Function

