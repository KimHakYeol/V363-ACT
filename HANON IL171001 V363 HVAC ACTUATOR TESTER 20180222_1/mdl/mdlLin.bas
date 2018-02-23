Attribute VB_Name = "mdlLin"
Option Explicit

Public Function LinIDWrite(ByVal IDNO As Integer) As Integer
    Dim lStatus As Long
    Dim Transmit As NCTYPE_CAN_FRAME
    
    Transmit.ArbitrationId = IDNO And (&H3F)
    Transmit.DataLength = 0
    Transmit.IsRemote = 17 ' 1
    
    lStatus = ncWrite(LinTxRx, Len(Transmit), Transmit)
    
    If (CheckStatus(lStatus, "ncWrite") = True) Then GoTo ERROR
    
    Call Delay(10)
    
    Exit Function

ERROR:
    
    Status = ncCloseObject(LinTxRx)
End Function

Public Function LinActWrite(ByVal ByteNo As Integer, ByVal IDNO As Integer, ByVal Data0 As Integer, ByVal Data1 As Integer, ByVal Data2 As Integer, ByVal Data3 As Integer, ByVal Data4 As Integer, ByVal Data5 As Integer, ByVal Data6 As Integer, ByVal Data7 As Integer) As Integer
    Dim Transmit As NCTYPE_CAN_FRAME
    Dim lStatus As Long
   
    Transmit.ArbitrationId = CLng(IDNO And &H3F)
    Transmit.DataLength = ByteNo
    Transmit.IsRemote = 18 ' full

    ' idÇÊ·¯
    Transmit.Data(0) = CLng(Data0)
    Transmit.Data(1) = CLng(Data1)
    Transmit.Data(2) = CLng(Data2)
    Transmit.Data(3) = CLng(Data3)
    Transmit.Data(4) = CLng(Data4)
    Transmit.Data(5) = CLng(Data5)
    Transmit.Data(6) = CLng(Data6)
    Transmit.Data(7) = CLng(Data7)
    
    lStatus = ncWrite(LinTxRx, Len(Transmit), Transmit)
    
    If (CheckStatus(lStatus, "ncWrite") = True) Then GoTo ERROR
    
    Call Delay(30)
    
    Exit Function

ERROR:
    
    Status = ncCloseObject(LinTxRx)
End Function

Public Function LinInit(Optional ByVal lID As Long = &HFF)
    If lID = &HFF Then
        Call LinActWrite(8, &H22, &H7F, &HE5, &HF9, &HFF, &HFF, &HFF, &HFF, &H40)
        Call LinActWrite(8, &H22, &H7F, &HE5, &H51, &HFF, &HFF, &HFF, &HFF, &H0)
    Else
        Call LinActWrite(8, &H22, CInt(lID), &HE5, &HF9, &HFF, &HFF, &HFF, &HFF, &H40)
        Call LinActWrite(8, &H22, CInt(lID), &HE5, &H51, &HFF, &HFF, &HFF, &HFF, &H0)
    End If
    
    Call Delay(50)
End Function

Public Sub LinAutoAddress()
    Dim i As Integer
    
    Call OnLog("[LIN] AUTOADDRESS NO : " & nLinAutoAddressFlag)
    
    Select Case nLinAutoAddressFlag
        ' BSM init
        Case 30: Call LinActWrite(8, &H3C, &H7F, &H6, &HB5, &HFF, &H7F, &H1, &HF1, &HFF)
        
        ' Assign Nad to slave n (New NAD -> 7,6,5,4)
        Case 40: Call LinActWrite(8, &H3C, &H7F, &H6, &HB5, &HFF, &H7F, &H2, &HF1, &H7)
        Case 41: Call LinActWrite(8, &H3C, &H7F, &H6, &HB5, &HFF, &H7F, &H2, &HF1, &H6)
        Case 42: Call LinActWrite(8, &H3C, &H7F, &H6, &HB5, &HFF, &H7F, &H2, &HF1, &H5)
        Case 43: Call LinActWrite(8, &H3C, &H7F, &H6, &HB5, &HFF, &H7F, &H2, &HF1, &H4)
        
        ' Store NAD in slave
        Case 44: Call LinActWrite(8, &H3C, &H7F, &H6, &HB5, &HFF, &H7F, &H3, &HF1, &HFF)
        
        ' Assign NAD finish
        Case 45: Call LinActWrite(8, &H3C, &H7F, &H6, &HB5, &HFF, &H7F, &H4, &HF1, &HFF)
        
        ' Assign ID range   (New NAD :7,6,5,4)
        Case 46: Call LinActWrite(8, &H3C, &H7, &H6, &HB7, &H0, &HE2, &H47, &HFF, &HFF)
        Case 47: Call LinActWrite(8, &H3C, &H6, &H6, &HB7, &H0, &HE2, &H6, &HFF, &HFF)
        Case 48: Call LinActWrite(8, &H3C, &H5, &H6, &HB7, &H0, &HE2, &H85, &HFF, &HFF)
        Case 49: Call LinActWrite(8, &H3C, &H4, &H6, &HB7, &H0, &HE2, &HC4, &HFF, &HFF)
        
        Case 50:
            frmRun.pnlLinAddrResult.Caption = "VERIFY ADDRESS."
            
            bAutoLinRead = True
        
        Case 60 To 70:
            If SetupVar.bLinActUse(0) Then
                Call LinIDWrite(BYTE1_ACT01)
                
                nLinReadSeq(0) = 10
            End If
        
        Case 80 To 90:
            If SetupVar.bLinActUse(1) Then
                Call LinIDWrite(BYTE1_ACT02)
                
                nLinReadSeq(1) = 11
            End If
        
        Case 100 To 110:
            If SetupVar.bLinActUse(2) Then
                Call LinIDWrite(BYTE1_Act03)
                
                nLinReadSeq(2) = 12
            End If
        
        Case 120 To 130:
            If SetupVar.bLinActUse(3) Then
                Call LinIDWrite(BYTE1_ACT04)
                
                nLinReadSeq(3) = 13
            End If
        
        Case 140:
            frmRun.pnlLinAddrResult.BackColor = vbGreen
            
            For i = 0 To 3
                If frmRun.pnlLinAddr(i).BackColor = vbRed Then
                    frmRun.pnlLinAddrResult.BackColor = vbRed
                    
                    Exit For
                End If
            Next
            
            bAutoLinRead = False
    
    End Select
    
    DoEvents
End Sub
