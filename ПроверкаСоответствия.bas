Attribute VB_Name = "NewMacros"
Sub ПроверкаСоответствияТребованиям()

    Dim doc As Document
    Set doc = ActiveDocument

    Dim issues As String
    issues = "Проверка требований:" & vbCrLf

    ' 1. Проверка шрифта и размера
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        With para.Range.Font
            If .Name <> "Times New Roman" Then
                issues = issues & "- Обнаружен другой шрифт: " & .Name & vbCrLf
            End If
            If .Size <> 12 Then
                issues = issues & "- Обнаружен другой размер шрифта: " & .Size & vbCrLf
            End If
        End With
    Next para

    ' 2. Проверка межстрочного интервала
    For Each para In doc.Paragraphs
        If para.LineSpacingRule <> wdLineSpaceSingle Then
            issues = issues & "- Обнаружен не одинарный межстрочный интервал" & vbCrLf
            Exit For
        End If
    Next para

    ' 3. Проверка полей страницы
    With doc.PageSetup
        If .TopMargin <> CentimetersToPoints(1.5) Then issues = issues & "- Верхнее поле не 15 мм" & vbCrLf
        If .BottomMargin <> CentimetersToPoints(1.5) Then issues = issues & "- Нижнее поле не 15 мм" & vbCrLf
        If .LeftMargin <> CentimetersToPoints(2) Then issues = issues & "- Левое поле не 20 мм" & vbCrLf
        If .RightMargin <> CentimetersToPoints(2) Then issues = issues & "- Правое поле не 20 мм" & vbCrLf
        If .Orientation <> wdOrientPortrait Then issues = issues & "- Не книжная ориентация" & vbCrLf
    End With

    ' 4. Абзацный отступ
    For Each para In doc.Paragraphs
        If para.LeftIndent <> CentimetersToPoints(0.75) And para.Style = "Обычный" Then
            issues = issues & "- Абзацный отступ не 0.75 см в стиле 'Обычный'" & vbCrLf
            Exit For
        End If
    Next para

    ' 5. Проверка на колонтитулы, сноски, нумерацию
    If doc.Footnotes.Count > 0 Then
        issues = issues & "- Документ содержит сноски" & vbCrLf
    End If
    If doc.Sections(1).Footers(wdHeaderFooterPrimary).Exists Then
        If doc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text <> "" Then
            issues = issues & "- Документ содержит нижние колонтитулы" & vbCrLf
        End If
    End If
    If doc.Sections(1).Headers(wdHeaderFooterPrimary).Exists Then
        If doc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text <> "" Then
            issues = issues & "- Документ содержит верхние колонтитулы" & vbCrLf
        End If
    End If

    ' 6. Нумерация страниц
    If doc.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.Count > 0 Then
        issues = issues & "- Обнаружена нумерация страниц" & vbCrLf
    End If

    MsgBox issues, vbInformation, "Результаты проверки"
End Sub
