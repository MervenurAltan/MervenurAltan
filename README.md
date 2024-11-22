Sub FiltreleVeAc()
    Dim ws As Worksheet
    Dim kriter As String
    Dim degerler() As String
    Dim veriAraligi As Range
    Dim satir As Range
    Dim sutunHarfi As String
    Dim sutunNumarasi As Integer
    Dim i As Integer
    Dim eslesti As Boolean
    
    ' Çalışma sayfasını seç
    Set ws = ThisWorkbook.Sheets("Sayfa1") ' Sayfa adını kontrol edin
    
    ' Kullanıcıdan filtrelemek istediği sütunu sor (A-M arası)
    sutunHarfi = InputBox("Hangi sütunda arama yapmak istiyorsunuz? (A-M arası sütun harfini girin)", "Sütun Seçimi")
    
    ' Girilen sütun harfini büyük harfe çevir
    sutunHarfi = UCase(sutunHarfi)
    
    ' Eğer kullanıcı geçerli bir sütun harfi girmezse çıkış yap
    If Len(sutunHarfi) <> 1 Or Not (sutunHarfi Like "[A-M]") Then
        MsgBox "Geçerli bir sütun harfi girmelisiniz (A-M).", vbExclamation
        Exit Sub
    End If
    
    ' Sütun harfini sütun numarasına dönüştür
    sutunNumarasi = Asc(sutunHarfi) - Asc("A") + 1
    
    ' Kullanıcıdan filtrelemek istediği değerleri alın
    kriter = InputBox("Hangi değerleri aratmak istersiniz? Virgülle ayırarak giriniz:", "Filtreleme Kriteri")
    
    ' Eğer kullanıcı bir değer girmezse çıkış yap
    If kriter = "" Then
        MsgBox "Bir değer girmelisiniz.", vbExclamation
        Exit Sub
    End If
    
    ' Virgülle ayrılmış değerleri diziye dönüştür
    degerler = Split(kriter, ",")
    
    ' Girilen değerlerdeki boşlukları temizle
    For i = LBound(degerler) To UBound(degerler)
        degerler(i) = Trim(degerler(i))
    Next i
    
    ' Veri aralığını tanımlayın (ilk satır başlık, veri aralığı sütun için dinamik)
    Set veriAraligi = ws.Range(ws.Cells(2, sutunNumarasi), ws.Cells(ws.Cells(ws.Rows.Count, sutunNumarasi).End(xlUp).Row, sutunNumarasi))
    
    ' Tüm satırları gizlemeden önce aç
    ws.Rows.Hidden = False
    
    ' Her bir satır için kontrol yap
    For Each satir In veriAraligi
        eslesti = False ' Her satır için eşleşmeyi sıfırla
        For i = LBound(degerler) To UBound(degerler)
            ' Hücredeki değer girilen kriterlerden biriyle eşleşiyor mu?
            If satir.Value = degerler(i) Then
                eslesti = True
                Exit For ' Eşleşme bulunduysa daha fazla kontrol gerekmez
            End If
        Next i
        
        ' Eşleşme bulunmadıysa satırı gizle
        If Not eslesti Then
            satir.EntireRow.Hidden = True
        End If
    Next satir
End Sub

--------------
Sub FiltreyiTemizle()
    Dim ws As Worksheet
    Dim sutunHarfi As String
    Dim sutunNumarasi As Integer
    Dim veriAraligi As Range
    
    ' Çalışma sayfasını seç
    Set ws = ThisWorkbook.Sheets("Sayfa1") ' Sayfa adını kontrol edin
    
    ' Kullanıcıdan filtreyi temizlemek istediği sütunu sor (A-M arası)
    sutunHarfi = InputBox("Hangi sütundaki filtreyi temizlemek istiyorsunuz? (A-M arası sütun harfini girin)", "Filtreyi Temizle")
    
    ' Girilen sütun harfini büyük harfe çevir
    sutunHarfi = UCase(sutunHarfi)
    
    ' Eğer kullanıcı geçerli bir sütun harfi girmezse çıkış yap
    If Len(sutunHarfi) <> 1 Or Not (sutunHarfi Like "[A-M]") Then
        MsgBox "Geçerli bir sütun harfi girmelisiniz (A-M).", vbExclamation
        Exit Sub
    End If
    
    ' Sütun harfini sütun numarasına dönüştür
    sutunNumarasi = Asc(sutunHarfi) - Asc("A") + 1
    
    ' Veri aralığını tanımlayın (ilk satır başlık, veri aralığı sütun için dinamik)
    Set veriAraligi = ws.Range(ws.Cells(2, sutunNumarasi), ws.Cells(ws.Cells(ws.Rows.Count, sutunNumarasi).End(xlUp).Row, sutunNumarasi))
    
    ' Tüm satırları yeniden görünür yap
    ws.Rows.Hidden = False
    
    ' Kullanıcıya bilgi ver
    MsgBox sutunHarfi & " sütunundaki filtre temizlendi.", vbInformation
End Sub
