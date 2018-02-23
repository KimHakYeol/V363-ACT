Attribute VB_Name = "mdlBarcode"
Option Explicit


' =============================================================================
' PRINTER API
' -----------------------------------------------------------------------------
Private Type DOCINFO
    lpDocName       As String
    lpOutputFile    As String
    lpDatatype      As String
End Type

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long

Public Function BarCodePrint(Optional ByVal lpBarcode As String = "") As Boolean
    Dim lhPrinter As Long
    Dim Nil As Long
    Dim lRes As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim MyDocInfo As DOCINFO
    Dim lpCmd As String
    Dim lpRes As String
    Dim lpPrinterModel As String
    Dim lpActData(1) As String
    Dim i As Integer
    
    lpPrinterModel = "ZEBRA_GT420T"
    
    If SysVar.lpPrintName <> Printer.DeviceName And DEBUGMODE Then
        Debug.Print "NAME 1 : " & SysVar.lpPrintName ' ini printer name
        Debug.Print "NAME 2 : " & Printer.DeviceName ' window default printer
        
        Exit Function
    End If
    
    If Trim$(lpBarcode) = "" Then lpBarcode = Format(SysVar.lOkCounter, "0000")
    
    Nil = 0
    lRes = OpenPrinter(SysVar.lpPrintName, lhPrinter, Nil)
    
    If lRes = 0 Then
        Call MsgBox("Printer Not Exist.")
        Exit Function
    End If
    
    MyDocInfo.lpDocName = "ILHO"
    MyDocInfo.lpOutputFile = vbNullString
    MyDocInfo.lpDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)
    
    lpCmd = ""
    
    Select Case lpPrinterModel
        Case "DATAMAX":
            lpCmd = lpCmd + Chr(2) & "L" & Chr(13)
            lpCmd = lpCmd + "H18" & Chr(13)
            lpCmd = lpCmd + "D11" & Chr(13)
            
            lpRes = Left(SetupVar.lpBarcode(0), 1)
            
            Select Case lpRes
                Case "A", "B", "C": lpRes = "&C"
                Case Else: lpRes = ""
            End Select
            
            lpCmd = lpCmd + "1e02030" & "0050" & "0020" & lpRes & Trim$(SetupVar.lpBarcode(0)) & Chr(13)
            lpCmd = lpCmd + "1311000" & "0010" & "0020" & Trim$(SetupVar.lpBarcode(0)) & Chr(13)
            lpCmd = lpCmd + "1111000" & "0010" & "0060" & Format(Now, "YYYYMMDD") & Format(SysVar.lOkCounter + 1, "0000") & Chr(13)
            lpCmd = lpCmd + "Q0001" & Chr(13)
            lpCmd = lpCmd + "E" & Chr(13)
        
        Case "ZEBRA":
            lpCmd = lpCmd + "CT~~CD,~CC^~CT~" & Chr$(13)
            lpCmd = lpCmd + "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR6,6~SD15^JUS^LRN^CI0^XZ" & Chr$(13)
            lpCmd = lpCmd + "^XA" & Chr$(13)
            lpCmd = lpCmd + "^MMT" & Chr$(13)
            lpCmd = lpCmd + "^PW531" & Chr$(13)
            lpCmd = lpCmd + "^LL0295" & Chr$(13)
            lpCmd = lpCmd + "^LS0" & Chr$(13)
            
            lpCmd = lpCmd + "^FT60,80" & "^A0N,40,40" & "^FH\" & "^FD" & "MODEL : " & Trim$(SetupVar.lpBarcode(0)) & "^FS" & Chr$(13)
            lpCmd = lpCmd + "^FT60,120" & "^A0N,40,40" & "^FH\" & "^FD" & "HMC P/NO : " & Trim$(SetupVar.lpBarcode(1)) & "^FS" & Chr$(13)
            lpCmd = lpCmd + "^BY4,4,90" & "^FT60,220" & "^BCN,,Y,N" & "^FD>:" & Trim$(SetupVar.lpBarcode(2)) & "^FS" & Chr$(13)
            lpCmd = lpCmd + "^FT390,255" & "^A0N,40,40" & "^FH\" & "^FD" & "HANON" & "^FS" & Chr$(13)
            lpCmd = lpCmd + "^FT60,295" & "^A0N,40,40" & "^FH\" & "^FD" & "LOT NO : " & Format(Now, "YYYY.MM.DD") & "." & lpBarcode & "^FS" & Chr$(13)
            
            lpCmd = lpCmd + "^PQ1,0,1,Y^XZ" & Chr$(13)
        
        Case "ZEBRA2":
            lpCmd = lpCmd + "CT~~CD,~CC^~CT~" & Chr$(13)
            lpCmd = lpCmd + "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR6,6~SD15^JUS^LRN^CI0^XZ" & Chr$(13)
            lpCmd = lpCmd + "^XA" & Chr$(13)
            lpCmd = lpCmd + "^MMT" & Chr$(13)
            lpCmd = lpCmd + "^PW531" & Chr$(13)
            lpCmd = lpCmd + "^LL0295" & Chr$(13)
            lpCmd = lpCmd + "^LS0" & Chr$(13)
            
            ' barcode - code128
            lpCmd = lpCmd + "^BY5,4,140" & "^FT100,300" & "^BCN,,N,N" & "^FD>:" & Trim$(SetupVar.lpBarcode(3)) & "^FS" & Chr$(13)
            
            ' text
            lpCmd = lpCmd + "^FT100,380" & "^A0N,80,80" & "^FH\" & "^FD" & Trim$(SetupVar.lpBarcode(3)) & "^FS" & Chr$(13)
            lpCmd = lpCmd + "^FT320,340" & "^A0N,36,36" & "^FH\" & "^FD" & Trim$(SetupVar.lpBarcode(4)) & "^FS" & Chr$(13)
            lpCmd = lpCmd + "^FT320,380" & "^A0N,36,36" & "^FH\" & "^FD" & Format(Now, "YYYYMMDD") & lpBarcode & "^FS" & Chr$(13)
            
            lpCmd = lpCmd + "^PQ1,0,1,Y^XZ" & Chr$(13)
        
        Case "ZEBRA_GT420T":
            lpCmd = lpCmd + "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR6,6~SD15^JUS^LRN^CI0^XZ" & Chr$(13)
            lpCmd = lpCmd + "^XA" & Chr$(13)
            lpCmd = lpCmd + "^MMT" & Chr$(13)
            lpCmd = lpCmd + "^PW531" & Chr$(13)
            lpCmd = lpCmd + "^LL0295" & Chr$(13)
            lpCmd = lpCmd + "^LS0" & Chr$(13)
            
            ' ^BY - barcode setting
            ' ^FO - data position X, Y
            ' ^A0 - 0 (default font) 1~?
            
            For i = 0 To 1
                If frmRun.pnlAct04Volt(i).Caption = "" Then
                    lpActData(i) = "0.00"
                Else
                    lpActData(i) = frmRun.pnlAct04Volt(i).Caption
                End If
            Next
            
            lpBarcode = Trim$(SetupVar.lpBarcode(4)) & " " & Trim$(SetupVar.lpBarcode(2)) & " " & Trim$(SetupVar.lpBarcode(3)) & " " & Format(Now, "YYYYMMDD") & "-" & Format(SysVar.lOkCounter + 1, "0000") & " " & Format(lpActData(0), "0.00") & " " & Format(lpActData(1), "0.00")
            
            Select Case SetupVar.nBarcodeType
                Case 0: ' datamatrix
                    ' barcode - datamatrix
                    lpCmd = lpCmd + "^FO" & "90,30" & "^BY" & "2,2,80" & "^BXN,5,200" & "^FD" & lpBarcode & "^FS" & Chr$(13)
                    
                    ' top text
                    lpCmd = lpCmd + "^FO" & "230,40" & "^A0" & "N,30,30" & "^FH\" & "^FD" & Trim$(SetupVar.lpBarcode(2)) & "^FS" & Chr$(13)
                    ' mid text
                    lpCmd = lpCmd + "^FO" & "230,80" & "^A0" & "N,30,30" & "^FH\" & "^FD" & Trim$(SetupVar.lpBarcode(3)) & "^FS" & Chr$(13)
                    ' end text
                    lpCmd = lpCmd + "^FO" & "230,120" & "^A0" & "N,30,30" & "^FH\" & "^FD" & Format(Now, "YYYYMMDD") & "-" & Format(SysVar.lOkCounter + 1, "0000") & "^FS" & Chr$(13)
                Case 1:
                    ' barcode - qr
                    lpCmd = lpCmd + "^FT" & "90,160" & "^BQN,2,4" & "^FDMM,A" & lpBarcode & "^FS" & Chr$(13)
                    
                    ' top text
                    lpCmd = lpCmd + "^FO" & "230,40" & "^A0" & "N,30,30" & "^FH\" & "^FD" & Trim$(SetupVar.lpBarcode(2)) & "^FS" & Chr$(13)
                    ' mid text
                    lpCmd = lpCmd + "^FO" & "230,80" & "^A0" & "N,30,30" & "^FH\" & "^FD" & Trim$(SetupVar.lpBarcode(3)) & "^FS" & Chr$(13)
                    ' end text
                    lpCmd = lpCmd + "^FO" & "230,120" & "^A0" & "N,30,30" & "^FH\" & "^FD" & Format(Now, "YYYYMMDD") & "-" & Format(SysVar.lOkCounter + 1, "0000") & "^FS" & Chr$(13)
            End Select
            
            lpCmd = lpCmd + "^PQ1,0,1,Y^XZ" & Chr$(13)
    
    End Select
    
    lRes = WritePrinter(lhPrinter, ByVal lpCmd, Len(lpCmd), lpcWritten)
    lRes = EndPagePrinter(lhPrinter)
    lRes = EndDocPrinter(lhPrinter)
    lRes = ClosePrinter(lhPrinter)
End Function
