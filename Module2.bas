Attribute VB_Name = "Module2"
Option Explicit

' �����ֶ�����
Private m_ExifDateIndex As Integer
Private m_ExifDateIndexFound As Boolean

Public ZuiZaoDate As Integer


Public Function GetFileDateParts(ByVal filePath As String, ByVal useExifFirst As Boolean, _
                          ByRef outYear As Integer, ByRef outMonth As Integer, ByRef outDay As Integer) As Boolean

    Dim objShell As Object
    Dim objFolder As Object
    Dim objFile As Object
    
    Dim dtExif As Variant
    Dim dtFileName As Variant
    Dim dtModified As Variant
    Dim dtCreated As Variant
    Dim validDates As Collection
    Dim oldDates As Collection   ' ���������ڱ���С��2000�������
    Dim minDate As Date
    Dim i As Integer
    Dim hasExif As Boolean
    Dim fileName As String
    Dim Yyear As Integer, Ymonth As Integer, Yday As Integer
    Dim isValid As Boolean

    On Error GoTo ErrorHandler

    ' ���� Shell ����
    Set objShell = CreateObject("Shell.Application")
    
    ' ��ȡ�ļ�����Ŀ¼
    Set objFolder = objShell.NameSpace(Left(filePath, InStrRev(filePath, "\") - 1))
    
    ' ��ȡ�����ļ�����
    Set objFile = objFolder.ParseName(Right(filePath, Len(filePath) - InStrRev(filePath, "\")))
    
    ' ��ʼ�����ϱ�����Ч����
    Set validDates = New Collection
    Set oldDates = New Collection  ' �������ϣ���¼С��2000�������
        
    ' ���Դ� EXIF ��ȡ����
    'Dim dtExif As Variant
    'Dim hasExif As Boolean
     Dim attrName As String
     Dim attrValue As String
    
    If Not m_ExifDateIndexFound Then
        ' ��һ�����У����ҡ��������ڡ��򡰴���ý�����ڡ�������
        For i = 0 To 300
            attrName = objFolder.GetDetailsOf(objFolder.Items, i)
            If InStr(attrName, "��������") > 0 Or InStr(attrName, "����ý������") > 0 Then
                m_ExifDateIndex = i
                m_ExifDateIndexFound = True
                Exit For
            End If
        Next i
        
        ' ���û�ҵ��ֶΣ����� EXIF ����
        If Not m_ExifDateIndexFound Then
            GoTo SkipExif
        End If
    End If

    ' ʹ�û����������ȡ EXIF ֵ
    attrValue = objFolder.GetDetailsOf(objFile, m_ExifDateIndex)
    
    If attrValue <> "" Then
        dtExif = ExtractDateParts(attrValue)
        If Not IsNull(dtExif) Then
            validDates.Add dtExif
            hasExif = True
        End If
    End If
    
SkipExif:
    ' ��� useExifFirst Ϊ True �� EXIF �ɹ���ȡ����� �� 2000����ֱ�ӷ���
    If useExifFirst And hasExif And Year(dtExif) >= ZuiZaoDate Then
        outYear = Year(dtExif)
        outMonth = Month(dtExif)
        outDay = Day(dtExif)
        GetFileDateParts = True
        Exit Function
    End If

    ' ���Դ��ļ�����ȡ����
    fileName = objFile.Name
    isValid = ExtractDateFromFileName(fileName, Yyear, Ymonth, Yday, "")
    If isValid Then
        dtFileName = DateSerial(Yyear, Ymonth, Yday)
        If Year(dtFileName) >= ZuiZaoDate Then
            validDates.Add dtFileName
        Else
            oldDates.Add dtFileName
        End If
    End If

    ' ��ȡ�޸�ʱ��ʹ���ʱ��
    On Error Resume Next
    dtModified = CDate(objFolder.GetDetailsOf(objFile, 3)) ' �޸�ʱ��
    dtCreated = CDate(objFolder.GetDetailsOf(objFile, 4))   ' ����ʱ��
    On Error GoTo ErrorHandler
    
    If Year(dtModified) >= ZuiZaoDate Then
        validDates.Add dtModified
    Else
        oldDates.Add dtModified
    End If
    
    If Year(dtCreated) >= ZuiZaoDate Then
        validDates.Add dtCreated
    Else
        oldDates.Add dtCreated
    End If

    ' �ҳ�������Ч���������������
    Dim dt As Variant
    Dim useCollection As Collection
    
    If validDates.Count > 0 Then
        Set useCollection = validDates
    Else
        Set useCollection = oldDates
    End If

    If useCollection.Count = 0 Then
        GetFileDateParts = False
        Exit Function
    End If

    minDate = Now
    For Each dt In useCollection
        If dt < minDate Then minDate = dt
    Next dt

    ' ������
    outYear = Year(minDate)
    outMonth = Month(minDate)
    outDay = Day(minDate)
    GetFileDateParts = True
    Exit Function

ErrorHandler:
    GetFileDateParts = False
End Function

Function ExtractDateParts(dateString As String) As Variant
    On Error Resume Next
    Dim cleanStr As String
    Dim spacePos As Integer
    Dim parts() As String
    Dim y As Integer, m As Integer, d As Integer

    cleanStr = Trim(dateString)
    spacePos = InStr(cleanStr, " ")
    If spacePos > 0 Then cleanStr = Left(cleanStr, spacePos - 1)
    cleanStr = Replace(Replace(cleanStr, "-", "/"), ".", "/")
    parts = Split(cleanStr, "/")

    If UBound(parts) >= 2 Then
        y = Val(OnlyDigits(parts(0)))
        m = Val(OnlyDigits(parts(1)))
        d = Val(OnlyDigits(parts(2)))

        If IsDate(DateSerial(y, m, d)) Then
            ExtractDateParts = DateSerial(y, m, d)
        Else
            ExtractDateParts = Null ' ����ʧ�ܣ����� Null
        End If
    Else
        ExtractDateParts = Null ' ��ʽ���ԣ����� Null
    End If
End Function


Function ExtractDateFromFileName(ByVal fileName As String, ByRef outYear As Integer, ByRef outMonth As Integer, ByRef outDay As Integer, ByRef errMsg As String) As Boolean
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim dtStr As String
    Dim dt As Date
    Dim y$, m$, d$
    
    On Error GoTo ErrorHandler
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Pattern = "\d{4}-\d{1,2}-\d{1,2}|" & _
                   "\d{4}\.\d{1,2}\.\d{1,2}|" & _
                   "\d{4}_\d{1,2}_\d{1,2}|" & _
                   "\d{8}|" & _
                   "\d{2}\d{2}\d{2}|" & _
                   "\d{2}-\d{1,2}-\d{4}"
        .Global = True
    End With

    If regEx.Test(fileName) Then
        Set matches = regEx.Execute(fileName)
        For Each match In matches
            dtStr = match.Value
            
            Select Case True
                Case InStr(dtStr, "-") > 0
                    Dim parts: parts = Split(Replace(dtStr, "-", "/"), "/")
                    If UBound(parts) < 2 Then GoTo NextMatch
                    If Len(parts(2)) = 4 Then
                        m = parts(0): d = parts(1): y = parts(2)
                    Else
                        y = parts(0): m = parts(1): d = parts(2)
                    End If
                    
                Case InStr(dtStr, ".") > 0
                    parts = Split(Replace(dtStr, ".", "/"), "/")
                    If UBound(parts) < 2 Then GoTo NextMatch
                    y = parts(0): m = parts(1): d = parts(2)
                    
                Case InStr(dtStr, "_") > 0
                    parts = Split(Replace(dtStr, "_", "/"), "/")
                    If UBound(parts) < 2 Then GoTo NextMatch
                    y = parts(0): m = parts(1): d = parts(2)
                    
                Case Len(dtStr) = 8
                    y = Left(dtStr, 4)
                    m = Mid(dtStr, 5, 2)
                    d = Mid(dtStr, 7, 2)
                    
                Case Len(dtStr) = 6 And IsNumeric(dtStr)
                    y = "20" & Left(dtStr, 2)
                    m = Mid(dtStr, 3, 2)
                    d = Mid(dtStr, 5, 2)
                    
                Case Else
                    GoTo NextMatch
            End Select
            
            dtStr = y & "/" & m & "/" & d
            If IsDate(dtStr) Then
                dt = CDate(dtStr)
                outYear = Year(dt)
                outMonth = Month(dt)
                outDay = Day(dt)
                ExtractDateFromFileName = True
                errMsg = "Success"
                Exit Function
            End If

NextMatch:
        Next match
    End If

    ExtractDateFromFileName = False
    errMsg = "No valid date found in the filename."
    Exit Function

ErrorHandler:
    ExtractDateFromFileName = False
    errMsg = "Error: " & Err.Description
End Function

Function OnlyDigits(inputStr As String) As String
    Dim result As String
    Dim i As Integer

    result = ""
    For i = 1 To Len(inputStr)
        If IsNumeric(Mid(inputStr, i, 1)) Then
            result = result & Mid(inputStr, i, 1)
        End If
    Next i

    OnlyDigits = result
End Function

