Attribute VB_Name = "Module2"
Option Explicit

' 缓存字段索引
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
    Dim oldDates As Collection   ' 新增：用于保存小于2000年的日期
    Dim minDate As Date
    Dim i As Integer
    Dim hasExif As Boolean
    Dim fileName As String
    Dim Yyear As Integer, Ymonth As Integer, Yday As Integer
    Dim isValid As Boolean

    On Error GoTo ErrorHandler

    ' 创建 Shell 对象
    Set objShell = CreateObject("Shell.Application")
    
    ' 获取文件所在目录
    Set objFolder = objShell.NameSpace(Left(filePath, InStrRev(filePath, "\") - 1))
    
    ' 获取具体文件对象
    Set objFile = objFolder.ParseName(Right(filePath, Len(filePath) - InStrRev(filePath, "\")))
    
    ' 初始化集合保存有效日期
    Set validDates = New Collection
    Set oldDates = New Collection  ' 新增集合：记录小于2000年的日期
        
    ' 尝试从 EXIF 获取日期
    'Dim dtExif As Variant
    'Dim hasExif As Boolean
     Dim attrName As String
     Dim attrValue As String
    
    If Not m_ExifDateIndexFound Then
        ' 第一次运行：查找“拍摄日期”或“创建媒体日期”的索引
        For i = 0 To 300
            attrName = objFolder.GetDetailsOf(objFolder.Items, i)
            If InStr(attrName, "拍摄日期") > 0 Or InStr(attrName, "创建媒体日期") > 0 Then
                m_ExifDateIndex = i
                m_ExifDateIndexFound = True
                Exit For
            End If
        Next i
        
        ' 如果没找到字段，跳过 EXIF 部分
        If Not m_ExifDateIndexFound Then
            GoTo SkipExif
        End If
    End If

    ' 使用缓存的索引获取 EXIF 值
    attrValue = objFolder.GetDetailsOf(objFile, m_ExifDateIndex)
    
    If attrValue <> "" Then
        dtExif = ExtractDateParts(attrValue)
        If Not IsNull(dtExif) Then
            validDates.Add dtExif
            hasExif = True
        End If
    End If
    
SkipExif:
    ' 如果 useExifFirst 为 True 且 EXIF 成功提取且年份 ≥ 2000，则直接返回
    If useExifFirst And hasExif And Year(dtExif) >= ZuiZaoDate Then
        outYear = Year(dtExif)
        outMonth = Month(dtExif)
        outDay = Day(dtExif)
        GetFileDateParts = True
        Exit Function
    End If

    ' 尝试从文件名提取日期
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

    ' 获取修改时间和创建时间
    On Error Resume Next
    dtModified = CDate(objFolder.GetDetailsOf(objFile, 3)) ' 修改时间
    dtCreated = CDate(objFolder.GetDetailsOf(objFile, 4))   ' 创建时间
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

    ' 找出所有有效日期中最早的日期
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

    ' 输出结果
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
            ExtractDateParts = Null ' 解析失败，返回 Null
        End If
    Else
        ExtractDateParts = Null ' 格式不对，返回 Null
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

