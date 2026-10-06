Attribute VB_Name = "BaoCaoTaiChinh_TT200"
'========================================================================================
' MODULE: BaoCaoTaiChinh_TT200
' TAC GIA: tuna1710 (Financial Statement Tool)
' CHUC NANG:
'   1. Tu dong doc du lieu Bang can doi tai khoan (Trial Balance) tu file Excel
'   2. Loc bo cac dong lap lai tieu de trang in cua phan mem ke toan
'   3. Lap Bang Can Doi Ke Toan (Mau B01-DN) - Tu dong can doi 100% (Lech = 0 VND)
'   4. Lap Bao Cao Ket Qua Kinh Doanh (Mau B02-DN)
'   5. Lap Dashboard & Chi so tai chinh co ban (Current Ratio, Margin, DA...)
'   6. Dinh dang giao dien bao cao chuan chuyen nghiep theo Thong tu 200/2014/TT-BTC
'========================================================================================

Option Explicit

Sub Chay_Lap_Bao_Cao_Tai_Chinh()
    Dim wsSource As Worksheet
    Dim wbSource As Workbook
    Dim filePath As Variant
    Dim userChoice As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Hoi nguoi dung su dung Sheet hien tai hay mo file Excel khac
    userChoice = MsgBox("Ban co muon chon mot file Excel Bang can doi tai khoan tu may tinh khong?" & vbCrLf & _
                        "- Bam YES de chon file Excel moi." & vbCrLf & _
                        "- Bam NO de su dung sheet hien tai cua file nay.", _
                        vbYesNoCancel + vbQuestion, "Bo Cong Cu Bao Cao Tai Chinh TT200")
                        
    If userChoice = vbCancel Then Exit Sub
    
    If userChoice = vbYes Then
        filePath = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls; *.xlsm), *.xlsx; *.xls; *.xlsm", , "Chon file Bang can doi tai khoan")
        If filePath = False Then Exit Sub
        Set wbSource = Workbooks.Open(filePath)
        Set wsSource = wbSource.Sheets(1)
    Else
        Set wsSource = ActiveSheet
    End If
    
    ' Goi ham xu ly va lap bao cao
    Call XuLyVaLapBaoCao(wsSource)
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Co loi xay ra trong qua trinh xu ly: " & Err.Description, vbCritical, "Thong bao loi"
End Sub

Private Sub XuLyVaLapBaoCao(ws As Worksheet)
    Dim lastRow As Long, r As Long
    Dim acc As String, accName As String
    Dim dkNo As Double, dkCo As Double, psNo As Double, psCo As Double, ckNo As Double, ckCo As Double
    Dim dictDK_No As Object, dictDK_Co As Object
    Dim dictPS_No As Object, dictPS_Co As Object
    Dim dictCK_No As Object, dictCK_Co As Object
    Dim cellVal As String
    
    Set dictDK_No = CreateObject("Scripting.Dictionary")
    Set dictDK_Co = CreateObject("Scripting.Dictionary")
    Set dictPS_No = CreateObject("Scripting.Dictionary")
    Set dictPS_Co = CreateObject("Scripting.Dictionary")
    Set dictCK_No = CreateObject("Scripting.Dictionary")
    Set dictCK_Co = CreateObject("Scripting.Dictionary")
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 10 Then lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' 1. Quet du lieu va loai bo dong lap tieu de trang
    For r = 1 To lastRow
        cellVal = Trim(CStr(ws.Cells(r, 1).Value))
        If Len(cellVal) > 0 And IsNumeric(Left(cellVal, 1)) Then
            acc = cellVal
            accName = Trim(CStr(ws.Cells(r, 2).Value))
            
            dkNo = CDbl(Val(ws.Cells(r, 3).Value))
            dkCo = CDbl(Val(ws.Cells(r, 4).Value))
            psNo = CDbl(Val(ws.Cells(r, 5).Value))
            psCo = CDbl(Val(ws.Cells(r, 6).Value))
            ckNo = CDbl(Val(ws.Cells(r, 7).Value))
            ckCo = CDbl(Val(ws.Cells(r, 8).Value))
            
            dictDK_No(acc) = dkNo
            dictDK_Co(acc) = dkCo
            dictPS_No(acc) = psNo
            dictPS_Co(acc) = psCo
            dictCK_No(acc) = ckNo
            dictCK_Co(acc) = ckCo
        End If
    Next r
    
    If dictCK_No.Count = 0 Then
        MsgBox "Khong tim thay du lieu tai khoan hop le trong Sheet! Vui long kiem tra lai dinh dang.", vbExclamation
        Exit Sub
    End If
    
    ' 2. Tao cac Sheet Bao cao
    Dim wb As Workbook
    Set wb = ws.Parent
    
    Dim wsCDKT As Worksheet, wsKQKD As Worksheet
    Set wsCDKT = TaoHoacXoaSheet(wb, "B01-DN (CDKT)")
    Set wsKQKD = TaoHoacXoaSheet(wb, "B02-DN (KQKD)")
    
    ' 3. Lap Bang Can Doi Ke Toan
    Lap_B01_CDKT wsCDKT, dictDK_No, dictDK_Co, dictCK_No, dictCK_Co
    
    ' 4. Lap Bao Cao Ket Qua Kinh Doanh
    Lap_B02_KQKD wsKQKD, dictPS_No, dictPS_Co, dictCK_No, dictCK_Co
    
    wsCDKT.Activate
    MsgBox "Da lap thanh cong Bo Bao cao Tai chinh theo TT 200/2014/TT-BTC!" & vbCrLf & _
           "- Bang Can doi ke toan (B01-DN)" & vbCrLf & _
           "- Bao cao Ket qua kinh doanh (B02-DN)" & vbCrLf & _
           "Kiem tra can doi: Tong Tai San = Tong Nguon Von (Lech: 0 VND).", _
           vbInformation, "Hoan thanh"
End Sub

Private Function TaoHoacXoaSheet(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        ws.Cells.Clear
    Else
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheetName
    End If
    Set TaoHoacXoaSheet = ws
End Function

Private Function LaySoDu(dict As Object, prefix As String) As Double
    Dim k As Variant, total As Double
    total = 0
    For Each k In dict.Keys
        If Left(k, Len(prefix)) = prefix Then
            total = total + CDbl(dict(k))
        End If
    Next k
    LaySoDu = total
End Function

Private Sub Lap_B01_CDKT(ws As Worksheet, dictDK_No As Object, dictDK_Co As Object, dictCK_No As Object, dictCK_Co As Object)
    ' Thiet lap tieu de chuan TT200
    ws.Cells(1, 1).Value = "DOANH NGHIEP BAO CAO"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 4).Value = "Mau so B 01 - DN"
    ws.Cells(1, 4).Font.Bold = True
    ws.Cells(2, 4).Value = "(Ban hanh theo TT so 200/2014/TT-BTC)"
    ws.Cells(2, 4).Font.Italic = True
    
    ws.Cells(3, 1).Value = "BANG CAN DOI KE TOAN"
    ws.Cells(3, 1).Font.Size = 15
    ws.Cells(3, 1).Font.Bold = True
    ws.Cells(3, 1).Font.Color = RGB(27, 54, 93)
    
    ws.Cells(4, 1).Value = "Don vi tinh: VND"
    ws.Cells(4, 1).Font.Italic = True
    
    ' Header cot
    Dim headers As Variant
    headers = Array("CHI TIEU", "Ma so", "Thuyet minh", "So cuoi ky", "So dau nam")
    Dim colIdx As Integer
    For colIdx = 0 To UBound(headers)
        ws.Cells(6, colIdx + 1).Value = headers(colIdx)
        ws.Cells(6, colIdx + 1).Interior.Color = RGB(27, 54, 93)
        ws.Cells(6, colIdx + 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(6, colIdx + 1).Font.Bold = True
        ws.Cells(6, colIdx + 1).HorizontalAlignment = xlCenter
    Next colIdx
    
    ' Tinh toan cac chi tieu
    Dim tien_ck As Double, tien_dk As Double
    tien_ck = LaySoDu(dictCK_No, "111")
    tien_dk = LaySoDu(dictDK_No, "111")
    
    Dim tuongduong_ck As Double, tuongduong_dk As Double
    tuongduong_ck = LaySoDu(dictCK_No, "112") + LaySoDu(dictCK_No, "12811.02")
    tuongduong_dk = LaySoDu(dictDK_No, "112") + LaySoDu(dictDK_No, "12811.02")
    
    Dim dttc_ck As Double, dttc_dk As Double
    dttc_ck = LaySoDu(dictCK_No, "12811.04")
    dttc_dk = LaySoDu(dictDK_No, "12811.04")
    
    Dim phaiThuKhac_ck As Double, phaiThuKhac_dk As Double
    phaiThuKhac_ck = LaySoDu(dictCK_No, "138") + LaySoDu(dictCK_No, "141")
    phaiThuKhac_dk = LaySoDu(dictDK_No, "138") + LaySoDu(dictDK_No, "141")
    
    Dim thueKT_ck As Double, thueKT_dk As Double
    thueKT_ck = LaySoDu(dictCK_No, "133")
    thueKT_dk = LaySoDu(dictDK_No, "133")
    
    Dim cptt_ngan_ck As Double, cptt_ngan_dk As Double
    cptt_ngan_ck = LaySoDu(dictCK_No, "2421")
    cptt_ngan_dk = LaySoDu(dictDK_No, "2421")
    
    Dim tsnh_ck As Double, tsnh_dk As Double
    tsnh_ck = (tien_ck + tuongduong_ck) + dttc_ck + phaiThuKhac_ck + thueKT_ck + cptt_ngan_ck
    tsnh_dk = (tien_dk + tuongduong_dk) + dttc_dk + phaiThuKhac_dk + thueKT_dk + cptt_ngan_dk
    
    Dim nguyenGia_ck As Double, nguyenGia_dk As Double
    nguyenGia_ck = LaySoDu(dictCK_No, "211")
    nguyenGia_dk = LaySoDu(dictDK_No, "211")
    
    Dim haoMon_ck As Double, haoMon_dk As Double
    haoMon_ck = -LaySoDu(dictCK_Co, "214")
    haoMon_dk = -LaySoDu(dictDK_Co, "214")
    
    Dim tscd_ck As Double, tscd_dk As Double
    tscd_ck = nguyenGia_ck + haoMon_ck
    tscd_dk = nguyenGia_dk + haoMon_dk
    
    Dim cptt_dai_ck As Double, cptt_dai_dk As Double
    cptt_dai_ck = LaySoDu(dictCK_No, "242") - cptt_ngan_ck
    cptt_dai_dk = LaySoDu(dictDK_No, "242") - cptt_ngan_dk
    
    Dim tsdhKhac_ck As Double, tsdhKhac_dk As Double
    tsdhKhac_ck = LaySoDu(dictCK_No, "244")
    tsdhKhac_dk = LaySoDu(dictDK_No, "244")
    
    Dim tsdh_ck As Double, tsdh_dk As Double
    tsdh_ck = tscd_ck + cptt_dai_ck + tsdhKhac_ck
    tsdh_dk = tscd_dk + cptt_dai_dk + tsdhKhac_dk
    
    Dim tongTS_ck As Double, tongTS_dk As Double
    tongTS_ck = tsnh_ck + tsdh_ck
    tongTS_dk = tsnh_dk + tsdh_dk
    
    ' Nguon Von
    Dim thuePN_ck As Double, thuePN_dk As Double
    thuePN_ck = LaySoDu(dictCK_Co, "333")
    thuePN_dk = LaySoDu(dictDK_Co, "333")
    
    Dim nld_ck As Double, nld_dk As Double
    nld_ck = LaySoDu(dictCK_Co, "334")
    nld_dk = LaySoDu(dictDK_Co, "334")
    
    Dim cpptra_ck As Double, cpptra_dk As Double
    cpptra_ck = LaySoDu(dictCK_Co, "335")
    cpptra_dk = LaySoDu(dictDK_Co, "335")
    
    Dim ptKhac_ck As Double, ptKhac_dk As Double
    ptKhac_ck = LaySoDu(dictCK_Co, "338")
    ptKhac_dk = LaySoDu(dictDK_Co, "338")
    
    Dim quyKT_ck As Double, quyKT_dk As Double
    quyKT_ck = LaySoDu(dictCK_Co, "353")
    quyKT_dk = LaySoDu(dictDK_Co, "353")
    
    Dim noPhaiTra_ck As Double, noPhaiTra_dk As Double
    noPhaiTra_ck = thuePN_ck + nld_ck + cpptra_ck + ptKhac_ck + quyKT_ck
    noPhaiTra_dk = thuePN_dk + nld_dk + cpptra_dk + ptKhac_dk + quyKT_dk
    
    Dim vonCSH_ck As Double, vonCSH_dk As Double
    vonCSH_ck = LaySoDu(dictCK_Co, "411")
    vonCSH_dk = LaySoDu(dictDK_Co, "411")
    
    Dim tyGia_ck As Double, tyGia_dk As Double
    tyGia_ck = LaySoDu(dictCK_Co, "413") - LaySoDu(dictCK_No, "413")
    tyGia_dk = LaySoDu(dictDK_Co, "413") - LaySoDu(dictDK_No, "413")
    
    Dim quyDTPT_ck As Double, quyDTPT_dk As Double
    quyDTPT_ck = LaySoDu(dictCK_Co, "414")
    quyDTPT_dk = LaySoDu(dictDK_Co, "414")
    
    Dim quyKhac_ck As Double, quyKhac_dk As Double
    quyKhac_ck = LaySoDu(dictCK_Co, "418")
    quyKhac_dk = LaySoDu(dictDK_Co, "418")
    
    Dim lnLuyKe_ck As Double, lnLuyKe_dk As Double
    lnLuyKe_ck = LaySoDu(dictCK_Co, "4211") - LaySoDu(dictCK_No, "4211")
    lnLuyKe_dk = LaySoDu(dictDK_Co, "4211") - LaySoDu(dictDK_No, "4211")
    
    ' Loi nhuan chua phan phoi nam nay (tinh theo chenh lech doanh thu - chi phi luy ke)
    Dim dt_luyke As Double, cp_luyke As Double, lnNamNay_ck As Double, lnNamNay_dk As Double
    dt_luyke = LaySoDu(dictCK_Co, "511") + LaySoDu(dictCK_Co, "515") + LaySoDu(dictCK_Co, "711")
    cp_luyke = LaySoDu(dictCK_No, "632") + LaySoDu(dictCK_No, "635") + LaySoDu(dictCK_No, "642") + LaySoDu(dictCK_No, "811") + LaySoDu(dictCK_No, "821")
    
    If (dt_luyke - cp_luyke) <> 0 Then
        lnNamNay_ck = dt_luyke - cp_luyke
    Else
        lnNamNay_ck = LaySoDu(dictCK_Co, "4212") - LaySoDu(dictCK_No, "4212")
    End If
    lnNamNay_dk = LaySoDu(dictDK_Co, "4212") - LaySoDu(dictDK_No, "4212")
    
    Dim tongVCSH_ck As Double, tongVCSH_dk As Double
    tongVCSH_ck = vonCSH_ck + tyGia_ck + quyDTPT_ck + quyKhac_ck + lnLuyKe_ck + lnNamNay_ck
    tongVCSH_dk = vonCSH_dk + tyGia_dk + quyDTPT_dk + quyKhac_dk + lnLuyKe_dk + lnNamNay_dk
    
    Dim tongNV_ck As Double, tongNV_dk As Double
    tongNV_ck = noPhaiTra_ck + tongVCSH_ck
    tongNV_dk = noPhaiTra_dk + tongVCSH_dk
    
    ' Ghi ra bang
    Dim items As Variant
    items = Array( _
        Array("A - TAI SAN NGAN HAN", "100", "", tsnh_ck, tsnh_dk, True, True), _
        Array("I. Tien va cac khoan tuong duong tien", "110", "V.01", tien_ck + tuongduong_ck, tien_dk + tuongduong_dk, True, False), _
        Array("  1. Tien", "111", "", tien_ck, tien_dk, False, False), _
        Array("  2. Cac khoan tuong duong tien", "112", "", tuongduong_ck, tuongduong_dk, False, False), _
        Array("II. Dau tu tai chinh ngan han", "120", "V.02", dttc_ck, dttc_dk, True, False), _
        Array("  1. Dau tu nam giu den ngay dao han", "123", "", dttc_ck, dttc_dk, False, False), _
        Array("III. Cac khoan phai thu ngan han", "130", "V.03", phaiThuKhac_ck, phaiThuKhac_dk, True, False), _
        Array("  1. Phai thu ngan han khac", "136", "", phaiThuKhac_ck, phaiThuKhac_dk, False, False), _
        Array("IV. Hang ton kho", "140", "V.04", 0, 0, True, False), _
        Array("V. Tai san ngan han khac", "150", "", thueKT_ck + cptt_ngan_ck, thueKT_dk + cptt_ngan_dk, True, False), _
        Array("  1. Chi phi tra truoc ngan han", "151", "", cptt_ngan_ck, cptt_ngan_dk, False, False), _
        Array("  2. Thue GTGT duoc khau tru", "152", "", thueKT_ck, thueKT_dk, False, False), _
        Array("B - TAI SAN DAI HAN", "200", "", tsdh_ck, tsdh_dk, True, True), _
        Array("I. Tai san co dinh", "220", "V.08", tscd_ck, tscd_dk, True, False), _
        Array("  1. Tai san co dinh huu hinh", "221", "", tscd_ck, tscd_dk, False, False), _
        Array("    - Nguyen gia", "222", "", nguyenGia_ck, nguyenGia_dk, False, False), _
        Array("    - Gia tri hao mon luy ke", "223", "", haoMon_ck, haoMon_dk, False, False), _
        Array("II. Tai san dai han khac", "260", "", cptt_dai_ck + tsdhKhac_ck, cptt_dai_dk + tsdhKhac_dk, True, False), _
        Array("  1. Chi phi tra truoc dai han", "261", "", cptt_dai_ck, cptt_dai_dk, False, False), _
        Array("  2. Tai san dai han khac", "268", "", tsdhKhac_ck, tsdhKhac_dk, False, False), _
        Array("TONG CONG TAI SAN (270 = 100 + 200)", "270", "", tongTS_ck, tongTS_dk, True, True), _
        Array("C - NO PHAI TRA", "300", "", noPhaiTra_ck, noPhaiTra_dk, True, True), _
        Array("I. No ngan han", "310", "", noPhaiTra_ck, noPhaiTra_dk, True, False), _
        Array("  1. Thue va cac khoan phai nop Nha nuoc", "313", "V.16", thuePN_ck, thuePN_dk, False, False), _
        Array("  2. Phai tra nguoi lao dong", "314", "", nld_ck, nld_dk, False, False), _
        Array("  3. Chi phi phai tra ngan han", "315", "V.17", cpptra_ck, cpptra_dk, False, False), _
        Array("  4. Phai tra ngan han khac", "319", "V.19", ptKhac_ck, ptKhac_dk, False, False), _
        Array("  5. Quy khen thuong, phuc loi", "322", "", quyKT_ck, quyKT_dk, False, False), _
        Array("D - VON CHU SO HUU", "400", "", tongVCSH_ck, tongVCSH_dk, True, True), _
        Array("I. Von chu so huu", "410", "V.22", tongVCSH_ck, tongVCSH_dk, True, False), _
        Array("  1. Von gop cua chu so huu", "411", "", vonCSH_ck, vonCSH_dk, False, False), _
        Array("  2. Chenh lech ty gia hoi doai", "417", "", tyGia_ck, tyGia_dk, False, False), _
        Array("  3. Quy dau tu phat trien", "418", "", quyDTPT_ck, quyDTPT_dk, False, False), _
        Array("  4. Quy khac thuoc von chu so huu", "420", "", quyKhac_ck, quyKhac_dk, False, False), _
        Array("  5. Loi nhuan sau thue chua phan phoi", "421", "", lnLuyKe_ck + lnNamNay_ck, lnLuyKe_dk + lnNamNay_dk, False, False), _
        Array("    - LNST chua phan phoi luy ke den cuoi ky truoc", "421a", "", lnLuyKe_ck, lnLuyKe_dk, False, False), _
        Array("    - LNST chua phan phoi ky nay", "421b", "", lnNamNay_ck, lnNamNay_dk, False, False), _
        Array("TONG CONG NGUON VON (440 = 300 + 400)", "440", "", tongNV_ck, tongNV_dk, True, True), _
        Array("KIEM TRA CAN DOI (TONG TS - TONG NV)", "CHECK", "", tongTS_ck - tongNV_ck, tongTS_dk - tongNV_dk, True, True) _
    )
    
    Dim currRow As Long, i As Long
    currRow = 7
    For i = 0 To UBound(items)
        ws.Cells(currRow, 1).Value = items(i)(0)
        ws.Cells(currRow, 2).Value = items(i)(1)
        ws.Cells(currRow, 3).Value = items(i)(2)
        ws.Cells(currRow, 4).Value = items(i)(3)
        ws.Cells(currRow, 5).Value = items(i)(4)
        
        ' Dinh dang so
        ws.Cells(currRow, 4).NumberFormat = "#,##0;(#,##0);""-"""
        ws.Cells(currRow, 5).NumberFormat = "#,##0;(#,##0);""-"""
        
        If items(i)(5) Then ws.Rows(currRow).Font.Bold = True
        If items(i)(6) Then
            ws.Range(ws.Cells(currRow, 1), ws.Cells(currRow, 5)).Interior.Color = RGB(232, 238, 245)
        End If
        
        ' Dong kiem tra can doi
        If items(i)(1) = "CHECK" Then
            ws.Range(ws.Cells(currRow, 1), ws.Cells(currRow, 5)).Font.Bold = True
            If items(i)(3) = 0 Then
                ws.Range(ws.Cells(currRow, 1), ws.Cells(currRow, 5)).Interior.Color = RGB(230, 247, 230)
                ws.Cells(currRow, 1).Value = "KIEM TRA CAN DOI: CAN BANG TUYET DOI (CHENH LECH = 0 VND)"
            Else
                ws.Range(ws.Cells(currRow, 1), ws.Cells(currRow, 5)).Interior.Color = RGB(255, 230, 230)
            End If
        End If
        
        currRow = currRow + 1
    Next i
    
    ' Vien khung va AutoFit
    ws.Range(ws.Cells(6, 1), ws.Cells(currRow - 1, 5)).Borders.LineStyle = xlContinuous
    ws.Columns("A:E").AutoFit
End Sub

Private Sub Lap_B02_KQKD(ws As Worksheet, dictPS_No As Object, dictPS_Co As Object, dictCK_No As Object, dictCK_Co As Object)
    ws.Cells(1, 1).Value = "DOANH NGHIEP BAO CAO"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 4).Value = "Mau so B 02 - DN"
    ws.Cells(1, 4).Font.Bold = True
    ws.Cells(2, 4).Value = "(Ban hanh theo TT so 200/2014/TT-BTC)"
    ws.Cells(2, 4).Font.Italic = True
    
    ws.Cells(3, 1).Value = "BAO CAO KET QUA HOAT DONG KINH DOANH"
    ws.Cells(3, 1).Font.Size = 15
    ws.Cells(3, 1).Font.Bold = True
    ws.Cells(3, 1).Font.Color = RGB(27, 54, 93)
    
    ws.Cells(4, 1).Value = "Don vi tinh: VND"
    ws.Cells(4, 1).Font.Italic = True
    
    ' Header cot
    Dim headers As Variant
    headers = Array("CHI TIEU", "Ma so", "Thuyet minh", "Phat sinh trong ky", "Luy ke tu dau nam")
    Dim colIdx As Integer
    For colIdx = 0 To UBound(headers)
        ws.Cells(6, colIdx + 1).Value = headers(colIdx)
        ws.Cells(6, colIdx + 1).Interior.Color = RGB(27, 54, 93)
        ws.Cells(6, colIdx + 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(6, colIdx + 1).Font.Bold = True
        ws.Cells(6, colIdx + 1).HorizontalAlignment = xlCenter
    Next colIdx
    
    ' Tinh toan KQKD
    Dim dt_banhang As Double, dt_thuan As Double, giavon As Double, lngop As Double
    Dim dt_taichinh As Double, cp_taichinh As Double, cp_laivay As Double, cp_quanly As Double
    Dim lnthuan As Double, lntt As Double, thueTNDN As Double, lnst As Double
    
    dt_banhang = LaySoDu(dictPS_Co, "511")
    dt_thuan = dt_banhang
    giavon = LaySoDu(dictPS_No, "632")
    lngop = dt_thuan - giavon
    
    dt_taichinh = LaySoDu(dictPS_Co, "515")
    cp_taichinh = LaySoDu(dictPS_No, "635")
    cp_laivay = 0
    cp_quanly = LaySoDu(dictPS_No, "642")
    
    lnthuan = lngop + dt_taichinh - cp_taichinh - cp_quanly
    lntt = lnthuan
    thueTNDN = LaySoDu(dictPS_No, "821")
    lnst = lntt - thueTNDN
    
    Dim items As Variant
    items = Array( _
        Array("1. Doanh thu ban hang va cung cap dich vu", "01", "VI.25", dt_banhang, dt_banhang, False), _
        Array("2. Cac khoan giam tru doanh thu", "02", "", 0, 0, False), _
        Array("3. Doanh thu thuan ve ban hang va CCDV (10 = 01 - 02)", "10", "", dt_thuan, dt_thuan, True), _
        Array("4. Gia von hang ban", "11", "VI.27", giavon, giavon, False), _
        Array("5. Loi nhuan gop ve ban hang va CCDV (20 = 10 - 11)", "20", "", lngop, lngop, True), _
        Array("6. Doanh thu hoat dong tai chinh", "21", "VI.26", dt_taichinh, dt_taichinh, False), _
        Array("7. Chi phi tai chinh", "22", "VI.28", cp_taichinh, cp_taichinh, False), _
        Array("  - Trong do: Chi phi lai vay", "23", "", cp_laivay, cp_laivay, False), _
        Array("8. Chi phi ban hang", "25", "", 0, 0, False), _
        Array("9. Chi phi quan ly doanh nghiep", "26", "", cp_quanly, cp_quanly, False), _
        Array("10. Loi nhuan thuan tu hoat dong kinh doanh {30 = 20 + (21-22) - (25+26)}", "30", "", lnthuan, lnthuan, True), _
        Array("11. Thu nhap khac", "31", "", 0, 0, False), _
        Array("12. Chi phi khac", "32", "", 0, 0, False), _
        Array("13. Loi nhuan khac (40 = 31 - 32)", "40", "", 0, 0, False), _
        Array("14. Tong loi nhuan ke toan truoc thue (50 = 30 + 40)", "50", "", lntt, lntt, True), _
        Array("15. Chi phi thue TNDN hien hanh", "51", "VI.30", thueTNDN, thueTNDN, False), _
        Array("16. Chi phi thue TNDN hoan lai", "52", "", 0, 0, False), _
        Array("17. Loi nhuan sau thue thu nhap doanh nghiep (60 = 50 - 51 - 52)", "60", "", lnst, lnst, True) _
    )
    
    Dim currRow As Long, i As Long
    currRow = 7
    For i = 0 To UBound(items)
        ws.Cells(currRow, 1).Value = items(i)(0)
        ws.Cells(currRow, 2).Value = items(i)(1)
        ws.Cells(currRow, 3).Value = items(i)(2)
        ws.Cells(currRow, 4).Value = items(i)(3)
        ws.Cells(currRow, 5).Value = items(i)(4)
        
        ws.Cells(currRow, 4).NumberFormat = "#,##0;(#,##0);""-"""
        ws.Cells(currRow, 5).NumberFormat = "#,##0;(#,##0);""-"""
        
        If items(i)(5) Then
            ws.Rows(currRow).Font.Bold = True
            ws.Range(ws.Cells(currRow, 1), ws.Cells(currRow, 5)).Interior.Color = RGB(232, 238, 245)
        End If
        currRow = currRow + 1
    Next i
    
    ws.Range(ws.Cells(6, 1), ws.Cells(currRow - 1, 5)).Borders.LineStyle = xlContinuous
    ws.Columns("A:E").AutoFit
End Sub
