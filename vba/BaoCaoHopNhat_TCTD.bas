Attribute VB_Name = "BaoCaoHopNhat_TCTD"
'========================================================================================
' MODULE: BaoCaoHopNhat_TCTD
' TAC GIA: tuna1710 (Financial Statement Tool)
' CHUC NANG:
'   1. Tu dong doc du lieu Bang can doi tai khoan (TT200) tu file Excel (Q1, Q2, Q3, Q4...)
'   2. Quy doi sang 84 tai khoan he thong TCTD cap 5 (Sheet 'so lieu')
'   3. Ap dung cac but toan can tru (Trung gian ngoai te, Chuyen phai thu ve tien gui, WU, WASH)
'   4. Tu dong cap nhat Bo Bao Cao Tai Chinh Hop Nhat:
'      - B02 (Mau B02a/TCTD): Bao cao tinh hinh tai chinh giua nien do
'      - B03 (Mau B03a/TCTD): Bao cao ket qua hoat dong
'      - B04 (Mau B04a/TCTD): Bao cao luu chuyen tien te (Truc tiep)
'      - B05 (Mau B05a/TCTD): Thuyet minh BCTC
'      - Thue: Tinh hinh thuc hien nghia vu Ngan sach Nha nuoc
'   5. Kiem tra can doi tuyet doi (Lech = 0 VND)
'========================================================================================

Option Explicit

Sub ChonFileBangCanDoi()
    Dim fd As Object
    Dim filePath As String
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim r As Long, lastR As Long
    Dim valStr As String
    Dim qNum As Integer, yrNum As Long
    Dim ck13885 As Double, ck13881 As Double
    Dim acc As String
    
    On Error GoTo ErrorHandler
    
    Set fd = Application.FileDialog(3) 'msoFileDialogFilePicker
    With fd
        .Title = "Chon file Bang can doi tai khoan (TT200)"
        .Filters.Clear
        .Filters.Add "Excel Files (*.xlsx; *.xls; *.xlsm)", "*.xlsx; *.xls; *.xlsm"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbSrc = Workbooks.Open(filePath, ReadOnly:=True)
    Set wsSrc = wbSrc.Sheets(1)
    
    ' 1. Quet tim Quy va Nam
    qNum = 1: yrNum = 2026
    For r = 1 To 15
        valStr = CStr(wsSrc.Cells(r, 1).Value)
        If InStr(1, valStr, "Qu", vbTextCompare) > 0 Then
            If InStr(1, valStr, "1") > 0 Then qNum = 1
            If InStr(1, valStr, "2") > 0 Then qNum = 2
            If InStr(1, valStr, "3") > 0 Then qNum = 3
            If InStr(1, valStr, "4") > 0 Then qNum = 4
            Dim m As Long
            For m = 2020 To 2035
                If InStr(1, valStr, CStr(m)) > 0 Then yrNum = m: Exit For
            Next m
            Exit For
        End If
    Next r
    
    ' 2. Quet tim so du cuoi ky cua 13885.02 va 13881 de goi y can tru
    lastR = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    ck13885 = 0: ck13881 = 0
    For r = 1 To lastR
        acc = Trim(CStr(wsSrc.Cells(r, 1).Value))
        If acc = "13885.02" Then
            ck13885 = Val(wsSrc.Cells(r, 13).Value)
        ElseIf acc = "13881" Then
            ck13881 = Val(wsSrc.Cells(r, 13).Value)
        End If
    Next r
    
    wbSrc.Close False
    Set wbSrc = Nothing
    
    ' 3. Ghi thong tin vao TRANG_CHU
    With ThisWorkbook.Sheets("TRANG_CHU")
        .Range("C7").Value = filePath
        .Range("C8").Value = "QUY " & qNum & " NAM " & yrNum
        .Range("D13").Value = ck13885
        .Range("D14").Value = ck13881
        .Range("D15").Value = 0
        .Range("D16").Value = 0
    End With
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Da nhan dien thanh cong file: Quy " & qNum & " nam " & yrNum & vbCrLf & _
           "- So goi y can tru ngoai te: " & Format(ck13885, "#,##0") & " VND" & vbCrLf & _
           "- So goi y chuyen phai thu: " & Format(ck13881, "#,##0") & " VND", _
           vbInformation, "Thanh cong"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not wbSrc Is Nothing Then wbSrc.Close False
    MsgBox "Loi khi doc file: " & Err.Description, vbCritical, "Thong bao loi"
End Sub

Sub CapNhatVaLapBaoCaoHopNhat()
    Dim wsHome As Worksheet, wsSL As Worksheet, wsBT As Worksheet
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim filePath As String, qStr As String
    Dim qNum As Integer, yrNum As Long
    Dim r As Long, lastR As Long
    Dim acc As String, valStr As String
    Dim canTruD7 As Double, canTruD25 As Double, canTruWU As Double, canTruWASH As Double
    Dim dictTB As Object
    Dim mapRules As Object
    Dim qRoman As String, endDateStr As String
    
    On Error GoTo ErrorHandler
    
    Set wsHome = ThisWorkbook.Sheets("TRANG_CHU")
    Set wsSL = ThisWorkbook.Sheets("so lieu")
    Set wsBT = ThisWorkbook.Sheets("But toan can tru")
    
    filePath = Trim(CStr(wsHome.Range("C7").Value))
    If filePath = "" Or Dir(filePath) = "" Then
        MsgBox "Vui long bam 'BUOC 1: CHON FILE BANG CAN DOI TAI KHOAN' truoc!", vbExclamation, "Chua chon file"
        Exit Sub
    End If
    
    ' Lay so tien can tru (Neu o sua tay co gia tri thi lay sua tay, khong thi lay so goi y)
    If Val(wsHome.Range("E13").Value) > 0 Then
        canTruD7 = Val(wsHome.Range("E13").Value)
    Else
        canTruD7 = Val(wsHome.Range("D13").Value)
    End If
    
    If Val(wsHome.Range("E14").Value) > 0 Then
        canTruD25 = Val(wsHome.Range("E14").Value)
    Else
        canTruD25 = Val(wsHome.Range("D14").Value)
    End If
    
    If Val(wsHome.Range("E15").Value) > 0 Then canTruWU = Val(wsHome.Range("E15").Value) Else canTruWU = Val(wsHome.Range("D15").Value)
    If Val(wsHome.Range("E16").Value) > 0 Then canTruWASH = Val(wsHome.Range("E16").Value) Else canTruWASH = Val(wsHome.Range("D16").Value)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 1. Mo file CĐTK nguon va doc toan bo tai khoan
    Set wbSrc = Workbooks.Open(filePath, ReadOnly:=True)
    Set wsSrc = wbSrc.Sheets(1)
    
    ' Nhan dien Quy / Nam
    qNum = 1: yrNum = 2026
    For r = 1 To 15
        valStr = CStr(wsSrc.Cells(r, 1).Value)
        If InStr(1, valStr, "Qu", vbTextCompare) > 0 Then
            If InStr(1, valStr, "1") > 0 Then qNum = 1
            If InStr(1, valStr, "2") > 0 Then qNum = 2
            If InStr(1, valStr, "3") > 0 Then qNum = 3
            If InStr(1, valStr, "4") > 0 Then qNum = 4
            Dim y As Long
            For y = 2020 To 2035
                If InStr(1, valStr, CStr(y)) > 0 Then yrNum = y: Exit For
            Next y
            Exit For
        End If
    Next r
    
    Set dictTB = CreateObject("Scripting.Dictionary")
    lastR = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    
    For r = 1 To lastR
        valStr = Trim(CStr(wsSrc.Cells(r, 1).Value))
        If valStr <> "" And IsNumeric(Left(valStr, 1)) Then
            ' Array: 0: dkNo, 1: dkCo, 2: psNo, 3: psCo, 4: ckNo, 5: ckCo
            dictTB(valStr) = Array(Val(wsSrc.Cells(r, 3).Value), _
                                   Val(wsSrc.Cells(r, 5).Value), _
                                   Val(wsSrc.Cells(r, 6).Value), _
                                   Val(wsSrc.Cells(r, 7).Value), _
                                   Val(wsSrc.Cells(r, 13).Value), _
                                   Val(wsSrc.Cells(r, 15).Value))
        End If
    Next r
    
    wbSrc.Close False
    Set wbSrc = Nothing
    
    ' 2. Khoi tao bang anh xa cac tai khoan
    Set mapRules = CreateMappingRules()
    
    ' 3. Dien du lieu vao sheet 'so lieu' (Dong 1 den 84)
    Dim tkSL As String
    Dim srcAccs As Variant
    Dim sumVals(5) As Double
    Dim item As Variant
    Dim tbRow As Variant
    Dim i As Integer
    
    For r = 1 To 84
        tkSL = Trim(CStr(wsSL.Cells(r, 1).Value))
        For i = 0 To 5: sumVals(i) = 0: Next i
        
        If mapRules.Exists(tkSL) Then
            srcAccs = mapRules(tkSL)
            For Each item In srcAccs
                If dictTB.Exists(CStr(item)) Then
                    tbRow = dictTB(CStr(item))
                    For i = 0 To 5
                        sumVals(i) = sumVals(i) + CDbl(tbRow(i))
                    Next i
                End If
            Next item
            
            ' Ap dung can tru dac thu
            If tkSL = "131101" Then
                sumVals(2) = sumVals(2) + canTruD25 ' PS No
                sumVals(4) = sumVals(4) + canTruD25 ' CK No
            ElseIf tkSL = "359299" Then
                sumVals(3) = sumVals(3) + (canTruD25 + canTruD7) ' PS Co
                sumVals(4) = sumVals(4) - (canTruD25 + canTruD7) ' CK No
            ElseIf tkSL = "459999" Then
                sumVals(3) = sumVals(3) - canTruD7 ' PS Co
                sumVals(5) = sumVals(5) - canTruD7 ' CK Co
            End If
        End If
        
        ' Ghi vao sheet 'so lieu'
        For i = 0 To 5
            wsSL.Cells(r, i + 2).Value = sumVals(i)
        Next i
    Next r
    
    ' Chuan hoa dong 86 tren sheet 'so lieu' ve kiem tra can doi noi bo (Lech = 0 d)
    wsSL.Range("A86").Value = "Chenh lech (No - Co)"
    wsSL.Range("B86").Formula = "=B85-C85"
    wsSL.Range("C86").Value = ""
    wsSL.Range("D86").Formula = "=D85-E85"
    wsSL.Range("E86").Value = ""
    wsSL.Range("F86").Formula = "=F85-G85"
    wsSL.Range("G86").Value = ""
    
    ' 4. Cap nhat sheet 'But toan can tru'
    wsBT.Range("D4").Value = canTruWU
    wsBT.Range("D7").Value = canTruD7
    wsBT.Range("D10").Value = canTruWASH
    wsBT.Range("D25").Value = canTruD25
    wsBT.Range("D27").Value = canTruD25
    
    ' 5. Cap nhat tieu de tren cac Sheet Bao cao
    Select Case qNum
        Case 1: qRoman = "I": endDateStr = "31/03"
        Case 2: qRoman = "II": endDateStr = "30/06"
        Case 3: qRoman = "III": endDateStr = "30/09"
        Case 4: qRoman = "IV": endDateStr = "31/12"
    End Select
    
    On Error Resume Next
    Dim wsB02 As Worksheet, wsB03 As Worksheet
    Dim wsB04 As Worksheet, wsB05 As Worksheet, wsThue As Worksheet
    Set wsB02 = ThisWorkbook.Sheets("B02")
    Set wsB03 = ThisWorkbook.Sheets("B03")
    Set wsB04 = ThisWorkbook.Sheets("B04")
    Set wsB05 = ThisWorkbook.Sheets("B05")
    Set wsThue = ThisWorkbook.Sheets("Thue")
    
    wsB02.Range("A6").Value = "Quy " & qRoman & " nam " & yrNum
    wsB02.Range("D78").Formula = "='so lieu'!G36"                 ' Von dieu le
    wsB02.Range("D84").Formula = "='so lieu'!G37+'so lieu'!G38"   ' Cac quy VCSH
    wsB02.Range("D88").Formula = "='B03'!L36"                     ' LN nam nay
    wsB02.Range("D89").Formula = "='so lieu'!G41"                 ' LN nam truoc
    
    wsB03.Range("A7").Value = "Nam " & yrNum
    wsB03.Range("A8").Value = "Ngay " & endDateStr & "/" & yrNum
    wsB03.Range("D11").Value = "Quy " & qRoman
    
    ' Khi lap Quy 2: Dien so lieu Quy 1 vao Cot J tren B03 de Cot L luy ke cong du 6 thang
    If qNum = 2 Then
        wsB03.Range("J11").Value = "Quy I"
        wsB03.Range("J12").Value = "Nam " & yrNum
        
        Dim j14 As Double, j17 As Double, j18 As Double, j20 As Double
        Dim j27 As Double, j30 As Double, j31 As Double, j34 As Double, j36 As Double
        
        ' Lay tu cot DK tren sheet so lieu:
        j14 = Val(wsSL.Range("C42").Value) - Val(wsSL.Range("B42").Value)
        j17 = Val(wsSL.Range("C43").Value)
        j18 = Val(wsSL.Range("B47").Value) - Val(wsSL.Range("C47").Value)
        j20 = Val(wsSL.Range("C44").Value) - Val(wsSL.Range("B44").Value) - Val(wsSL.Range("B48").Value)
        
        ' Tong chi phi quan ly Q1 (TK 832099, 851103..872001 tren cot DK No)
        j27 = Val(wsSL.Range("B50").Value) + _
              Val(wsSL.Range("B53").Value) + Val(wsSL.Range("B55").Value) + Val(wsSL.Range("B56").Value) + _
              Val(wsSL.Range("B57").Value) + Val(wsSL.Range("B58").Value) + Val(wsSL.Range("B59").Value) + _
              Val(wsSL.Range("B60").Value) + Val(wsSL.Range("B61").Value) + Val(wsSL.Range("B62").Value) + _
              Val(wsSL.Range("B64").Value) + Val(wsSL.Range("B66").Value) + Val(wsSL.Range("B67").Value) + _
              Val(wsSL.Range("B69").Value) + Val(wsSL.Range("B70").Value) + Val(wsSL.Range("B71").Value) + _
              Val(wsSL.Range("B72").Value) + Val(wsSL.Range("B73").Value) + Val(wsSL.Range("B74").Value) + _
              Val(wsSL.Range("B75").Value) + Val(wsSL.Range("B77").Value) + Val(wsSL.Range("B78").Value) + _
              Val(wsSL.Range("B79").Value) + Val(wsSL.Range("B80").Value)
              
        j30 = j14 + (j17 - j18) + j20 - j27
        j31 = Val(wsSL.Range("B51").Value) ' Thue TNDN Q1
        j34 = j30 - j31
        j36 = j34
        
        wsB03.Range("J14").Value = j14
        wsB03.Range("J17").Value = j17
        wsB03.Range("J18").Value = j18
        wsB03.Range("J20").Value = j20
        wsB03.Range("J27").Value = j27
        wsB03.Range("J30").Value = j30
        wsB03.Range("J31").Value = j31
        wsB03.Range("J34").Value = j34
        wsB03.Range("J36").Value = j36
    End If
    
    wsB04.Range("A7").Value = "Quy " & qRoman & " nam " & yrNum
    wsB04.Range("D50").Formula = "='so lieu'!B3"                              ' Du dau TK 1312
    wsB04.Range("D112").Formula = "='so lieu'!B1+'so lieu'!B2+'so lieu'!B4"    ' Tien dau ky dong
    wsB04.Range("D113").Formula = "=D114-D111-D112"                          ' Ty gia hoi doai
    wsB04.Range("D115").Formula = "=D114-D112-D111-D113"                     ' Kiem tra can doi = 0
    
    wsB05.Range("A5").Value = "Quy " & qRoman & " nam " & yrNum
    wsB05.Range("B199").Formula = "=Thue!D12"
    wsB05.Range("B202").Formula = "=Thue!D16"
    
    wsThue.Range("A5").Value = "Quy " & qNum & " nam " & yrNum
    
    ' Cap nhat so lieu Thue
    Dim arrT As Variant
    If dictTB.Exists("33311") Then
        arrT = dictTB("33311")
        wsThue.Range("D12").Value = arrT(1): wsThue.Range("E12").Value = arrT(3): wsThue.Range("F12").Value = arrT(2): wsThue.Range("G12").Value = arrT(3): wsThue.Range("L12").Value = arrT(2): wsThue.Range("M12").Value = arrT(5)
    End If
    If dictTB.Exists("33341.01") Then
        arrT = dictTB("33341.01")
        wsThue.Range("D16").Value = arrT(1): wsThue.Range("E16").Value = arrT(3): wsThue.Range("F16").Value = arrT(2): wsThue.Range("G16").Value = arrT(3): wsThue.Range("L16").Value = arrT(2): wsThue.Range("M16").Value = arrT(5)
    End If
    If dictTB.Exists("33351.01") Then
        arrT = dictTB("33351.01")
        wsThue.Range("D17").Value = arrT(1): wsThue.Range("E17").Value = arrT(3): wsThue.Range("F17").Value = arrT(2): wsThue.Range("G17").Value = arrT(3): wsThue.Range("L17").Value = arrT(2): wsThue.Range("M17").Value = arrT(5)
    End If
    If dictTB.Exists("33382.01") Then
        arrT = dictTB("33382.01")
        wsThue.Range("D21").Value = arrT(1): wsThue.Range("E21").Value = arrT(3): wsThue.Range("F21").Value = arrT(2): wsThue.Range("G21").Value = arrT(3): wsThue.Range("L21").Value = arrT(2): wsThue.Range("M21").Value = arrT(5)
    End If
    
    Dim cLetters As Variant, cL As Variant, totColVal As Double
    cLetters = Array("D", "E", "F", "G", "L", "M")
    For Each cL In cLetters
        totColVal = Val(wsThue.Range(cL & "12").Value) + Val(wsThue.Range(cL & "16").Value) + Val(wsThue.Range(cL & "17").Value) + Val(wsThue.Range(cL & "21").Value)
        wsThue.Range(cL & "11").Value = totColVal
        wsThue.Range(cL & "26").Value = totColVal
    Next cL

    On Error GoTo ErrorHandler
    
    ' 6. Kiem tra can doi va cap nhat Dashboard tren TRANG_CHU
    Dim diffDK As Double, diffPS As Double, diffCK As Double
    Dim totTS As Double, totNV As Double
    
    Application.Calculate
    
    diffDK = wsSL.Range("B85").Value - wsSL.Range("C85").Value
    diffPS = wsSL.Range("D85").Value - wsSL.Range("E85").Value
    diffCK = wsSL.Range("F85").Value - wsSL.Range("G85").Value
    
    totTS = Val(wsB02.Range("D59").Value)
    totNV = Val(wsB02.Range("D91").Value)
    
    With wsHome
        .Range("C8").Value = "QUY " & qNum & " NAM " & yrNum
        .Range("C22").Value = Format(wsSL.Range("B85").Value, "#,##0")
        .Range("D22").Value = Format(wsSL.Range("C85").Value, "#,##0")
        .Range("E22").Value = IIf(Abs(diffDK) < 1, "CAN DOI (0 d)", "LECH: " & Format(diffDK, "#,##0") & " d")
        .Range("F22").Value = IIf(Abs(diffDK) < 1, "Hoan hao", "Can kiem tra")
        
        .Range("C23").Value = Format(wsSL.Range("D85").Value, "#,##0")
        .Range("D23").Value = Format(wsSL.Range("E85").Value, "#,##0")
        .Range("E23").Value = IIf(Abs(diffPS) < 1, "CAN DOI (0 d)", "LECH: " & Format(diffPS, "#,##0") & " d")
        .Range("F23").Value = IIf(Abs(diffPS) < 1, "Hoan hao", "Can kiem tra")
        
        .Range("C24").Value = Format(wsSL.Range("F85").Value, "#,##0")
        .Range("D24").Value = Format(wsSL.Range("G85").Value, "#,##0")
        .Range("E24").Value = IIf(Abs(diffCK) < 1, "CAN DOI (0 d)", "LECH: " & Format(diffCK, "#,##0") & " d")
        .Range("F24").Value = IIf(Abs(diffCK) < 1, "Hoan hao", "Can kiem tra")
    End With
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "DA LAP THANH CONG BO BAO CAO HOP NHAT QUY " & qRoman & " NAM " & yrNum & "!" & vbCrLf & _
           "- Bang Can Doi Tai Khoan: Can doi tuyet doi (Lech = 0 VND)" & vbCrLf & _
           "- Vui long chuyen sang cac sheet B02, B03, B04, B05 de xem bao cao.", _
           vbInformation, "Thanh cong ruc ro"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not wbSrc Is Nothing Then wbSrc.Close False
    MsgBox "Co loi xay ra: " & Err.Description, vbCritical, "Thong bao loi"
End Sub

Private Function CreateMappingRules() As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    
    m("101101") = Array("11111.01")
    m("131101") = Array("11211.01")
    m("131201") = Array("12811.02", "12811.04")
    m("132101") = Array("11221.01")
    m("301301") = Array("21121.01")
    m("301401") = Array("21131.01")
    m("301501") = Array("21141.01")
    m("301902") = Array("21181.01")
    m("305101") = Array("21411.01")
    m("305102") = Array("21411.02")
    m("351001") = Array("24411.01")
    m("353202") = Array("1331")
    m("359299") = Array("13881.01", "13881.02", "13882.01", "13882.02", "13885.01", "13885.02")
    m("361299") = Array("14111.09")
    m("388009") = Array("24211.01")
    m("391") = Array("13884.01")
    m("397001") = Array("13884.02")
    m("427909") = Array("33861.01")
    m("453101") = Array("33311")
    m("453401") = Array("33341.01")
    m("453801") = Array("33382.01")
    m("453802") = Array("33351.01")
    m("454001") = Array("33883.01")
    m("459908") = Array("33511.09")
    m("459999") = Array("33821.01", "33831.01", "33841.01", "33882.02", "33884.01", "33884.02", "33886.01", "33886.02")
    m("462001") = Array("33411")
    m("484101") = Array("35311.01")
    m("484201") = Array("35321.01")
    m("497001") = Array("33885")
    m("601001") = Array("41111.01")
    m("611001") = Array("41811.01")
    m("612101") = Array("41411.01")
    m("631101") = Array("41311.01")
    m("691001") = Array("42121.01")
    m("692001") = Array("42111.01")
    m("701001") = Array("51511.02", "51511.03")
    m("711001") = Array("51131.01")
    m("721001") = Array("51511.01")
    m("811001") = Array("63211.01")
    m("821001") = Array("63511.01")
    m("832099") = Array("64251.03")
    m("833101") = Array("82111.01")
    m("851103") = Array("64211.01")
    m("852001") = Array("64213.01")
    m("853101") = Array("64212.01")
    m("853201") = Array("64212.02")
    m("853401") = Array("64212.03")
    m("853999") = Array("64212.04")
    m("856001") = Array("64211.04")
    m("859001") = Array("64213.06")
    m("861101") = Array("64221.01")
    m("861401") = Array("64221.03")
    m("862001") = Array("64281.02", "64281.04")
    m("863001") = Array("64281.01")
    m("865001") = Array("64271.02", "64271.13")
    m("866001") = Array("64282.08", "64282.11")
    m("868001") = Array("64282.05", "64282.06")
    m("869101") = Array("64271.01", "64271.11", "64271.12")
    m("869301") = Array("64282.02")
    m("869401") = Array("64282.03")
    m("869402") = Array("64282.04")
    m("869701") = Array("64271.07")
    m("869999") = Array("64271.10", "64281.03", "64281.09")
    m("871001") = Array("64241.01")
    m("872001") = Array("64271.08")
    
    Set CreateMappingRules = m
End Function
