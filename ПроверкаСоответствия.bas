Attribute VB_Name = "NewMacros"
Sub �������������������������������()

    Dim doc As Document
    Set doc = ActiveDocument

    Dim issues As String
    issues = "�������� ����������:" & vbCrLf

    ' 1. �������� ������ � �������
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        With para.Range.Font
            If .Name <> "Times New Roman" Then
                issues = issues & "- ��������� ������ �����: " & .Name & vbCrLf
            End If
            If .Size <> 12 Then
                issues = issues & "- ��������� ������ ������ ������: " & .Size & vbCrLf
            End If
        End With
    Next para

    ' 2. �������� ������������ ���������
    For Each para In doc.Paragraphs
        If para.LineSpacingRule <> wdLineSpaceSingle Then
            issues = issues & "- ��������� �� ��������� ����������� ��������" & vbCrLf
            Exit For
        End If
    Next para

    ' 3. �������� ����� ��������
    With doc.PageSetup
        If .TopMargin <> CentimetersToPoints(1.5) Then issues = issues & "- ������� ���� �� 15 ��" & vbCrLf
        If .BottomMargin <> CentimetersToPoints(1.5) Then issues = issues & "- ������ ���� �� 15 ��" & vbCrLf
        If .LeftMargin <> CentimetersToPoints(2) Then issues = issues & "- ����� ���� �� 20 ��" & vbCrLf
        If .RightMargin <> CentimetersToPoints(2) Then issues = issues & "- ������ ���� �� 20 ��" & vbCrLf
        If .Orientation <> wdOrientPortrait Then issues = issues & "- �� ������� ����������" & vbCrLf
    End With

    ' 4. �������� ������
    For Each para In doc.Paragraphs
        If para.LeftIndent <> CentimetersToPoints(0.75) And para.Style = "�������" Then
            issues = issues & "- �������� ������ �� 0.75 �� � ����� '�������'" & vbCrLf
            Exit For
        End If
    Next para

    ' 5. �������� �� �����������, ������, ���������
    If doc.Footnotes.Count > 0 Then
        issues = issues & "- �������� �������� ������" & vbCrLf
    End If
    If doc.Sections(1).Footers(wdHeaderFooterPrimary).Exists Then
        If doc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text <> "" Then
            issues = issues & "- �������� �������� ������ �����������" & vbCrLf
        End If
    End If
    If doc.Sections(1).Headers(wdHeaderFooterPrimary).Exists Then
        If doc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text <> "" Then
            issues = issues & "- �������� �������� ������� �����������" & vbCrLf
        End If
    End If

    ' 6. ��������� �������
    If doc.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.Count > 0 Then
        issues = issues & "- ���������� ��������� �������" & vbCrLf
    End If

    MsgBox issues, vbInformation, "���������� ��������"
End Sub
