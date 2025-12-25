Attribute VB_Name = "DeepL_Translator_v1"
'===================================================================================
' PROJE / PROJECT     : Excel VBA Translator Macro using DeepL API
' SÜRÜM / VERSION     : 1 (Free API Version)
' TARİH / DATE        : 25.12.2025
' YAZAR / AUTHOR      : [https://ahmetmersin.com]
'===================================================================================
'
' [TR] SÜRÜM NOTLARI VE ÖZELLİKLER (KURUMSAL)
' ----------------------------------------------------------------------------------
' 1. Kapsamlı İşlem Yeteneği:
'    Aktif çalışma kitabında yer alan tüm sayfalar (gizli sayfalar dahil olmak üzere)
'    kapsama alınarak sırasıyla tercüme edilmektedir.
'
' 2. Performans Optimizasyonu (Batch Processing):
'    Verimliliği artırmak amacıyla veri işleme mimarisi güncellenmiştir. Hücreler
'    50'li gruplar halinde paketlenerek tek bir API isteği ile çevrilmekte,
'    bu sayede işlem hızı önemli ölçüde artırılmıştır. Bu optimizasyon ile,
'    ücretsiz (Free) API planı kullanılarak dahi büyük ölçekli Excel dosyalarının
'    çevrilmesi mümkün hale gelmiştir.
'
' 3. Süreç Takibi:
'    İşlem süreci boyunca anlık ilerleme durumu ve işlenen sayfa bilgisi,
'    Excel durum çubuğu (Status Bar) üzerinden takip edilebilmektedir.
'
'===================================================================================
'
' [EN] RELEASE NOTES AND FEATURES
' ----------------------------------------------------------------------------------
' 1. Comprehensive Scope:
'    Automatically translates all sheets within the active workbook sequentially,
'    including hidden sheets.
'
' 2. Performance Optimization:
'    Uses "Batch Processing" to increase efficiency. Cells are grouped into sets
'    of 50 and sent in a single API request to speed up the process. With this
'    optimization, it is now possible to translate large-scale Excel files even
'    using the Free API plan.
'
' 3. Process Tracking:
'    You can track the progress and the current sheet name on the Status Bar
'    during the operation.
'===================================================================================

Sub TumKitabiIngilizceyeCevir_FreeAPI()
    On Error GoTo HataYakalayici

    ' --- KULLANICI AYARLARI ---
    Dim apiKey As String
    apiKey = "*****---------API_KEY_BURAYA_GELECEK---------******" '

    ' --- API AYARLARI ---
    ' Ücretsiz API planı kullanıldığı için endpoint aşağıdaki gibi olmalıdır.
    Dim endpoint As String
    endpoint = "https://api-free.deepl.com/v2/translate" ' <-- ÜCRETSİZ PLAN
    ' -------------------------

    If apiKey = "YENI_API_ANAHTARINIZI_BURAYA_YAPISTIRIN" Or apiKey = "" Then
        MsgBox "Lütfen kodun içindeki 'apiKey' değişkenini kendi yeni DeepL anahtarınızla güncelleyin.", vbCritical, "Ayar Gerekli"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim allCellsToTranslate As Collection
    Set allCellsToTranslate = New Collection
    
    ' Çeviri öncesi çevrilecek tüm hücreleri bir koleksiyonda topla
    Application.StatusBar = "Çevrilecek hücreler hesaplanıyor..."
    For Each ws In ThisWorkbook.Worksheets
        Dim cell As Range
        For Each cell In ws.UsedRange.Cells
            If Not IsEmpty(cell.Value) And Not IsNumeric(cell.Value) And Not IsDate(cell.Value) And Len(Trim(cell.Value)) > 0 Then
                allCellsToTranslate.Add cell
            End If
        Next cell
    Next ws
    
    If allCellsToTranslate.Count = 0 Then
        MsgBox "Çalışma kitabında çevrilecek metin içeren hücre bulunamadı.", vbInformation, "Boş Kitap"
        Exit Sub
    End If
    
    Dim onay As VbMsgBoxResult
    onay = MsgBox(ThisWorkbook.Worksheets.Count & " adet sayfada (gizliler dahil) toplam " & allCellsToTranslate.Count & " hücre İngilizce'ye çevrilecektir." & vbCrLf & vbCrLf & _
                  "Bu işlem uzun sürebilir ve geri alınamaz. Devam etmek istiyor musunuz?", vbYesNo + vbQuestion, "Tüm Kitap için Çeviri Onayı")
    
    If onay = vbNo Then
        MsgBox "İşlem iptal edildi.", vbInformation, "İptal"
        Exit Sub
    End If

    ' Performans için Excel ayarları
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    Dim i As Long
    Dim batchCells As Collection
    Dim islenenToplamHücre As Long
    islenenToplamHücre = 0

    For i = 1 To allCellsToTranslate.Count Step 50 ' 50'li gruplar halinde ilerle
        Set batchCells = New Collection
        Dim requestBody As String
        requestBody = "{""text"": ["
        
        Dim j As Long
        For j = i To IIf(i + 49 > allCellsToTranslate.Count, allCellsToTranslate.Count, i + 49)
            batchCells.Add allCellsToTranslate(j)
            requestBody = requestBody & """" & EscapeJson(allCellsToTranslate(j).Value) & ""","
        Next j
        
        If Right(requestBody, 1) = "," Then
            requestBody = Left(requestBody, Len(requestBody) - 1)
        End If
        
        ' <-- HEDEF DİL EN OLARAK AYARLANDI
        requestBody = requestBody & "], ""target_lang"": ""EN""}"
        
        http.Open "POST", endpoint, False
        http.setRequestHeader "Authorization", "DeepL-Auth-Key " & apiKey
        http.setRequestHeader "Content-Type", "application/json"
        http.send requestBody

        islenenToplamHücre = islenenToplamHücre + batchCells.Count
        Application.StatusBar = "Çeviri yapılıyor... (" & islenenToplamHücre & " / " & allCellsToTranslate.Count & ")"

        Select Case http.Status
            Case 200 ' Başarılı
                Dim responseText As String
                responseText = http.responseText
                
                Dim translations() As String
                translations = Split(responseText, """text"":""")
                
                Dim k As Long
                For k = 1 To batchCells.Count
                    Dim translatedText As String
                    translatedText = Split(translations(k), """")(0)
                    batchCells(k).Value = DecodeJsonChars(translatedText)
                Next k
                
            Case 403 ' Yetkilendirme hatası
                batchCells(1).Value = "Hata: API Anahtarı Geçersiz."
                MsgBox "API Anahtarı geçersiz veya yetkiniz yok. İşlem durduruldu.", vbCritical, "Yetki Hatası"
                GoTo TemizlikVeCikis
                
            Case 429, 456 ' Kota aşıldı
                batchCells(1).Value = "Hata: Aylık Kota Aşıldı."
                MsgBox "DeepL ÜCRETSİZ API kotası aşıldı. İşlem durduruldu.", vbCritical, "Kota Aşıldı"
                GoTo TemizlikVeCikis
                
            Case Else ' Diğer hatalar
                batchCells(1).Value = "Hata: Kod " & http.Status
                 MsgBox "Bilinmeyen bir API hatası oluştu. HTTP Durum Kodu: " & http.Status & vbCrLf & http.responseText, vbCritical, "API Hatası"
                GoTo TemizlikVeCikis
        End Select
        
        Application.Wait (Now + TimeValue("00:00:01"))
    Next i

TemizlikVeCikis:
    ' Excel ayarlarını normale döndür
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    If Err.Number = 0 Then
        MsgBox "Tüm sayfalardaki çeviri işlemi tamamlandı!", vbInformation, "İşlem Bitti"
    End If
    Exit Sub

HataYakalayici:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "Beklenmedik bir VBA hatası oluştu: " & Err.Description, vbCritical, "VBA Kodu Hatası"
End Sub

Private Function EscapeJson(text As String) As String
    EscapeJson = Replace(text, "\", "\\")
    EscapeJson = Replace(EscapeJson, """", "\""")
    EscapeJson = Replace(EscapeJson, vbLf, "\n")
    EscapeJson = Replace(EscapeJson, vbCr, "\r")
    EscapeJson = Replace(EscapeJson, vbTab, "\t")
End Function

Private Function DecodeJsonChars(text As String) As String
    DecodeJsonChars = Replace(text, "\""", """")
    DecodeJsonChars = Replace(DecodeJsonChars, "\\", "\")
    DecodeJsonChars = Replace(DecodeJsonChars, "\n", vbLf)
    DecodeJsonChars = Replace(DecodeJsonChars, "\r", vbCr)
    DecodeJsonChars = Replace(DecodeJsonChars, "\t", vbTab)
End Function