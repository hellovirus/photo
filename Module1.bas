Attribute VB_Name = "Module1"
Option Explicit

' 在模块顶部声明全局变量
Public fso As New FileSystemObject ' 确保在整个模块中都可以访问到 fso

Public gDictAllMedia As New Dictionary   ' 第1个：固定内容的所有媒体类型字典
Public gDictAllExt As New Dictionary    ' 第2个：遍历后生成的所有文件扩展名字典
'Public dictExtensions As New Dictionary    ' 第2个：遍历后生成的所有文件扩展名字典
Public gDictIntersect As New Dictionary ' 第3个：第1个和第3个/第4个的交集
Public gDictUserSel As New Dictionary   ' 第4个：用户输入的扩展名字典

Public gDictUserUse As New Dictionary   ' 第4个：用户输入的扩展名字典

Public ShuLiang As Long

Public Function MsgboxDict()
    DisplayDictionaryContentMsg gDictAllMedia
    DisplayDictionaryContentMsg gDictAllExt
    DisplayDictionaryContentMsg gDictUserUse

End Function

Public Sub InitGlobalDictAllMedia()
    Dim exts As String
    Dim arr() As String
    Dim i As Integer
    
    ' 固定内容字符串
    exts = "BMP,JPG,JPEG,PNG,TIF,GIF,PCX,TGA,MP4,AVI,MOV,MKV,FLV,WMV,MPEG,3GP,MP3,WMA,WAV"
    arr = Split(Trim(LCase(exts)), ",")
    
    ' 清空旧数据
    If Not gDictAllMedia Is Nothing Then
        If gDictAllMedia.Count > 0 Then gDictAllMedia.RemoveAll
    Else
        Set gDictAllMedia = New Dictionary
    End If
    
    ' 添加到字典
    For i = 0 To UBound(arr)
        If Not gDictAllMedia.Exists(arr(i)) Then
            gDictAllMedia.Add arr(i), True
        End If
    Next i
End Sub

Public Sub BuildGlobalDictIntersect()
   Dim key As Variant
   Dim KeyValue As Variant
    ' 清空旧数据
    If Not gDictUserUse Is Nothing Then
        If gDictUserUse.Count > 0 Then gDictUserUse.RemoveAll
    Else
        Set gDictUserUse = New Dictionary
    End If
    
    ' 如果用户自定义字典为空，则使用 gDictAllExt
    'If gDictUserSel.Count = 0 Then
        For Each key In gDictAllMedia.Keys
            If gDictAllExt.Exists(key) Then
                KeyValue = gDictAllExt.item(key)
                gDictUserUse.Add key, KeyValue
            End If
        Next key
    'Else
        ' 否则取 gDictAllMedia 和 gDictUserSel 的交集
    '    For Each key In gDictAllMedia.Keys
    '        If gDictUserSel.Exists(key) Then
    '            gDictUserUse.Add key, True
    '        End If
    '    Next key
    'End If
End Sub

Public Sub UpdateGlobalDictUserSel(userInput As String)
    Dim arr() As String
    Dim i As Integer
    
    ' 清空旧数据
    If Not gDictUserUse Is Nothing Then
        If gDictUserUse.Count > 0 Then gDictUserUse.RemoveAll
    Else
        Set gDictUserUse = New Dictionary
    End If
    
    If Trim(userInput) = "" Then Exit Sub
    
    arr = Split(Trim(LCase(userInput)), ",")
    
    For i = 0 To UBound(arr)
        Dim ext As String
        ext = Trim(arr(i))
        If ext <> "" Then
            If Not gDictUserUse.Exists(ext) Then
                gDictUserUse.Add ext, True
            End If
        End If
    Next i
End Sub

'判读msflexgrid 是否为空
Function IsGridEmpty(grid As MSFlexGrid) As Boolean
    IsGridEmpty = True
    Dim r As Long, c As Long
    With grid
        For r = .FixedRows To .Rows - 1       ' 跳过标题行
            For c = .FixedCols To .Cols - 1    ' 跳过标题列
                If .TextMatrix(r, c) <> "" Then
                    IsGridEmpty = False
                    Exit Function
                End If
            Next c
        Next r
    End With
End Function


Function IntersectExtensions(str1 As String, str2 As String) As String
    Dim arr1() As String
    Dim arr2() As String
    Dim dict As Object
    Dim i As Integer
    Dim result As String
    
        Dim item As String
        
    ' 创建字典对象来存储第一个数组中的元素
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 拆分两个字符串为数组
    If InStr(str1, ",") > 0 Then
        arr1 = Split(str1, ",")
    Else
        If str1 <> "" Then
            ReDim arr1(0)
            arr1(0) = str1
        Else
            ReDim arr1(-1 To 0) ' 空数组
        End If
    End If

    If InStr(str2, ",") > 0 Then
        arr2 = Split(str2, ",")
    Else
        If str2 <> "" Then
            ReDim arr2(0)
            arr2(0) = str2
        Else
            ReDim arr2(-1 To 0) ' 空数组
        End If
    End If

    ' 将第一个数组的元素加入字典
    For i = LBound(arr1) To UBound(arr1)
        'Dim item As String
        item = Trim(arr1(i))
        If item <> "" Then
            If Not dict.Exists(item) Then
                dict.Add item, Nothing
            End If
        End If
    Next i

    ' 遍历第二个数组，查找存在于字典中的项（即交集）
    result = ""
    For i = LBound(arr2) To UBound(arr2)
        'Dim item As String
        item = Trim(arr2(i))
        If item <> "" Then
            If dict.Exists(item) Then
                If result = "" Then
                    result = item
                Else
                    result = result & "," & item
                End If
            End If
        End If
    Next i

    ' 返回结果
    IntersectExtensions = result

    ' 清理
    Set dict = Nothing
End Function

' 递归创建多级目录（支持路径中包含多个不存在的文件夹）
Public Sub MkDirRecursive(ByVal fullPath As String)
    Dim pathParts() As String
    Dim currentPath As String
    Dim i As Integer

    fullPath = Trim(fullPath)
    If fullPath = "" Or Dir(fullPath, vbDirectory) <> "" Then Exit Sub

    pathParts = Split(fullPath, "\")
    currentPath = pathParts(0)
    For i = 1 To UBound(pathParts)
        currentPath = currentPath & "\" & pathParts(i)
        If Dir(currentPath, vbDirectory) = "" Then
            MkDir currentPath
        End If
    Next i
End Sub

' 复制文件到目标路径，成功返回 True，失败返回 False
Function CopyFileToTarget(srcPath As String, destPath As String) As Boolean
    On Error GoTo ErrHandler
    FileCopy srcPath, destPath
    CopyFileToTarget = True
    Exit Function

ErrHandler:
    CopyFileToTarget = False
End Function

' 移动文件（复制后删除源文件），成功返回 True，失败返回 False
Function MoveFileToTarget(srcPath As String, destPath As String) As Boolean
    On Error GoTo ErrHandler
    FileCopy srcPath, destPath
    Kill srcPath
    MoveFileToTarget = True
    Exit Function
ErrHandler:
    MoveFileToTarget = False
End Function

' 将操作记录写入日志文件
Sub LogFileAction(srcPath As String, destPath As String, success As Boolean, logFile As TextStream)
    Dim action As String
    If success Then
        action = "成功"
    Else
        action = "失败"
    End If
    'logFile.WriteLine Now & " | " & action & " | 源路径: " & srcPath & " | 目标路径: " & destPath
    logFile.WriteLine " | " & action & " | 源路径: " & srcPath & " | 目标路径: " & destPath
End Sub

'' 更新界面日志显示（比如在一个 ListBox 中显示）
Sub UpdateLogDisplay(text As String)
    Form1.lstlog.AddItem text
    Form1.lstlog.ListIndex = Form1.lstlog.NewIndex ' 滚动到最后一条
End Sub

' 根据文件日期生成目标路径
Function BuildTargetPath(ByVal fileDate As Date, ByVal basePath As String, ByVal structure As Integer) As String
    Dim yearStr As String, monthStr As String, dayStr As String
    yearStr = Year(fileDate)
    monthStr = Right("0" & Month(fileDate), 2)
    dayStr = Right("0" & Day(fileDate), 2)

    Select Case structure
        Case 1
            BuildTargetPath = basePath & "\" & yearStr
        Case 2
            BuildTargetPath = basePath & "\" & yearStr & "\" & monthStr
        Case 3
            BuildTargetPath = basePath & "\" & yearStr & "\" & monthStr & "\" & dayStr
        Case Else
            BuildTargetPath = basePath
    End Select

    ' 确保目标路径存在
    If Dir(BuildTargetPath, vbDirectory) = "" Then
        MkDirRecursive BuildTargetPath
    End If
End Function

' 判断目标路径下是否存在同名文件，并返回合适的文件路径（仅通过文件大小判断）
'Function GetDestinationFilePath(srcPath As String, targetPath As String) As String
'    Dim srcFileName As String
'    Dim baseName As String
'    Dim ext As String
'    Dim counter As Integer
'    Dim testPath As String
'    Dim existingSize As Long
'    Dim currentSize As Long
'    Dim recoveryDir As String
'
'    On Error GoTo ErrorHandler
'
'    ' 获取源文件名、基本名称和扩展名
'    srcFileName = Mid$(srcPath, InStrRev(srcPath, "\") + 1)
'    baseName = Left$(srcFileName, InStrRev(srcFileName, ".") - 1)
'    ext = Mid$(srcFileName, InStrRev(srcFileName, "."))
'
'    ' 构建初始的目标文件路径
'    GetDestinationFilePath = targetPath & "\" & srcFileName
'
'    ' 如果文件不存在，直接返回
'    If Dir(GetDestinationFilePath) = "" Then Exit Function
'
'    ' 获取目标文件和源文件的大小
'    existingSize = FileLen(GetDestinationFilePath)
'    currentSize = FileLen(srcPath)
'
'    ' 如果大小一致，认为是重复文件，移动到回收站目录
'    If existingSize = currentSize Then
'        ' 设置回收站目录为：目标文件夹根目录下的“同名文件回收站”
'        recoveryDir = BuildRecoveryPath(targetPath)
'
'        ' 返回回收站目录下的文件路径
'        GetDestinationFilePath = recoveryDir & "\" & srcFileName
'        Exit Function
'    End If
'
'    ' 否则尝试添加 -2、-3 等后缀避免重名
'    counter = 2
'    Do
'        testPath = targetPath & "\" & baseName & "-" & CStr(counter) & ext
'        If Dir(testPath) = "" Then
'            GetDestinationFilePath = testPath
'            Exit Do
'        End If
'        counter = counter + 1
'    Loop
'
'Exit Function
'ErrorHandler:
'    MsgBox "发生错误：" & Err.Description
'    GetDestinationFilePath = ""
'
'End Function
Function GetFileHash(ByVal filePath As String) As String
    Dim objStream As Object
    Dim objMD5 As Object
    Dim hashBytes As Variant
    Dim i As Integer
    Dim hexHash As String
    
    On Error GoTo ErrorHandler

    ' 创建 ADODB.Stream 和 MD5 加密对象
    Set objStream = CreateObject("ADODB.Stream")
    Set objMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

    ' 打开流并加载文件
    objStream.Type = 1 ' adTypeBinary
    objStream.Open
    objStream.LoadFromFile filePath

    ' 计算哈希值
    hashBytes = objMD5.ComputeHash_2(objStream.Read)

    ' 判断是否成功获取到哈希字节数组
    If Not isEmpty(hashBytes) And TypeName(hashBytes) = "Byte()" Then
        ' 遍历每个字节，转换为十六进制字符串
        For i = LBound(hashBytes) To UBound(hashBytes)
            hexHash = hexHash & Right$("0" & Hex(AscB(MidB(hashBytes, i + 1, 1))), 2)
        Next i
    Else
        hexHash = ""
    End If

    ' 返回小写的哈希值
    GetFileHash = LCase(hexHash)
    Exit Function

ErrorHandler:
    GetFileHash = ""
    MsgBox "计算文件哈希出错：" & Err.Description, vbCritical
End Function
' 判断目标路径下是否存在同名文件，并返回合适的文件路径（根据文件大小判断）
Function GetDestinationFilePath( _
    srcPath As String, _
    targetPath As String, _
    Optional useSizeCompare As Boolean = True) As String
    
    Dim srcFileName As String
    Dim baseName As String
    Dim ext As String
    Dim counter As Integer
    Dim testPath As String
    Dim currentSize As Long
    Dim existingSize As Long
    Dim currentHash As String
    Dim existingHash As String
    Dim recoveryDir As String
    
    On Error GoTo ErrorHandler

    ' 获取源文件名、基本名称和扩展名
    srcFileName = Mid$(srcPath, InStrRev(srcPath, "\") + 1)
    If InStrRev(srcFileName, ".") > 0 Then
        baseName = Left$(srcFileName, InStrRev(srcFileName, ".") - 1)
        ext = Mid$(srcFileName, InStrRev(srcFileName, "."))
    Else
        baseName = srcFileName
        ext = ""
    End If

    ' 构建初始的目标文件路径
    GetDestinationFilePath = targetPath & "\" & srcFileName

    ' 如果目标路径下不存在该文件，直接返回
    If Dir(GetDestinationFilePath) = "" Then Exit Function

    ' 根据比较方式选择判断逻辑
    If useSizeCompare Then
        ' 使用文件大小比对
        currentSize = FileLen(srcPath)
        existingSize = FileLen(GetDestinationFilePath)
        
        If existingSize = currentSize Then
            recoveryDir = BuildRecoveryPath(targetPath)
            GetDestinationFilePath = recoveryDir & "\" & srcFileName
            Exit Function
        End If
    Else
        ' 使用哈希比对
        currentHash = GetFileHash(srcPath)
        existingHash = GetFileHash(GetDestinationFilePath)
        
        If existingHash = currentHash And existingHash <> "" Then
            recoveryDir = BuildRecoveryPath(targetPath)
            GetDestinationFilePath = recoveryDir & "\" & srcFileName
            Exit Function
        End If
    End If

    ' 否则尝试添加 -2、-3 等后缀避免重名，并每次检查是否重复
    counter = 2
    Do
        testPath = targetPath & "\" & baseName & "-" & CStr(counter) & ext
        
        If Dir(testPath) = "" Then
            GetDestinationFilePath = testPath
            Exit Do
        Else
            If useSizeCompare Then
                existingSize = FileLen(testPath)
                currentSize = FileLen(srcPath)
                If existingSize = currentSize Then
                    recoveryDir = BuildRecoveryPath(targetPath)
                    GetDestinationFilePath = recoveryDir & "\" & srcFileName
                    Exit Do
                End If
            Else
                existingHash = GetFileHash(testPath)
                currentHash = GetFileHash(srcPath)
                If existingHash = currentHash And existingHash <> "" Then
                    recoveryDir = BuildRecoveryPath(targetPath)
                    GetDestinationFilePath = recoveryDir & "\" & srcFileName
                    Exit Do
                End If
            End If
            
            ' 不重复，则继续增加计数器
            counter = counter + 1
        End If
    Loop

Exit Function

ErrorHandler:
    MsgBox "发生错误：" & Err.Description, vbCritical
    GetDestinationFilePath = ""
End Function

' 构建回收站目录路径并确保其存在
Function BuildRecoveryPath(basePath As String) As String
    Dim recoveryPath As String
    recoveryPath = Form1.txtDestPath.text & "\同名文件回收站"
    
    ' 检查并创建回收站目录
    If Dir(recoveryPath, vbDirectory) = "" Then
        MkDir recoveryPath
    End If
    
    BuildRecoveryPath = recoveryPath
End Function

'显示字典的内容
Function DisplayDictionaryContent(dict As Dictionary)
    Dim key As Variant
    Dim content As String
   ' Dim Shuliang As Long
    
    content = "需要处理的文件类型:"
    ShuLiang = 0
    
'    If dict Is Nothing Then
'        MsgBox "字典为空或未初始化", vbExclamation
'        Exit Sub
'    End If
    
    If dict.Count = 0 Then
        'Form1.txtResult.text = " 没有符合条件并需要处理的文件！ "
        Form1.lblStatus.Caption = " 没有符合条件并需要处理的文件！ "
        Exit Function
    End If
    
    For Each key In dict.Keys
'        content = content & "键: " & key & ", 值: " & dict(key) & vbCrLf
         content = content & key & " "
         ShuLiang = ShuLiang + dict(key)
        ' MsgBox key & " " & dict(key)
    Next key
    
'    MsgBox content, vbInformation
    'Form1.txtResult.text = content
     Form1.lblStatus.Caption = content & "共计文件数：" & ShuLiang
End Function

'显示字典的内容
Public Sub DisplayDictionaryContentMsg(dict As Dictionary)
    Dim key As Variant
    Dim content As String
    
    content = "需要处理的文件类型:" & vbCrLf
    
    If dict Is Nothing Then
        MsgBox "字典为空或未初始化", vbExclamation
        Exit Sub
    End If
    
    If dict.Count = 0 Then
        MsgBox " 没有符合条件并需要处理的文件！ "
        Exit Sub
    End If
    
    For Each key In dict.Keys
         content = content & "键: " & key & ", 值: " & dict(key) & vbCrLf
         'content = content & key & ", "
    Next key
    
    MsgBox content, vbInformation
    'Form1.txtResult.text = content
End Sub

