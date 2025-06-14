Attribute VB_Name = "Module1"
Option Explicit

' ��ģ�鶥������ȫ�ֱ���
Public fso As New FileSystemObject ' ȷ��������ģ���ж����Է��ʵ� fso

Public gDictAllMedia As New Dictionary   ' ��1�����̶����ݵ�����ý�������ֵ�
Public gDictAllExt As New Dictionary    ' ��2�������������ɵ������ļ���չ���ֵ�
'Public dictExtensions As New Dictionary    ' ��2�������������ɵ������ļ���չ���ֵ�
Public gDictIntersect As New Dictionary ' ��3������1���͵�3��/��4���Ľ���
Public gDictUserSel As New Dictionary   ' ��4�����û��������չ���ֵ�

Public gDictUserUse As New Dictionary   ' ��4�����û��������չ���ֵ�

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
    
    ' �̶������ַ���
    exts = "BMP,JPG,JPEG,PNG,TIF,GIF,PCX,TGA,MP4,AVI,MOV,MKV,FLV,WMV,MPEG,3GP,MP3,WMA,WAV"
    arr = Split(Trim(LCase(exts)), ",")
    
    ' ��վ�����
    If Not gDictAllMedia Is Nothing Then
        If gDictAllMedia.Count > 0 Then gDictAllMedia.RemoveAll
    Else
        Set gDictAllMedia = New Dictionary
    End If
    
    ' ��ӵ��ֵ�
    For i = 0 To UBound(arr)
        If Not gDictAllMedia.Exists(arr(i)) Then
            gDictAllMedia.Add arr(i), True
        End If
    Next i
End Sub

Public Sub BuildGlobalDictIntersect()
   Dim key As Variant
   Dim KeyValue As Variant
    ' ��վ�����
    If Not gDictUserUse Is Nothing Then
        If gDictUserUse.Count > 0 Then gDictUserUse.RemoveAll
    Else
        Set gDictUserUse = New Dictionary
    End If
    
    ' ����û��Զ����ֵ�Ϊ�գ���ʹ�� gDictAllExt
    'If gDictUserSel.Count = 0 Then
        For Each key In gDictAllMedia.Keys
            If gDictAllExt.Exists(key) Then
                KeyValue = gDictAllExt.item(key)
                gDictUserUse.Add key, KeyValue
            End If
        Next key
    'Else
        ' ����ȡ gDictAllMedia �� gDictUserSel �Ľ���
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
    
    ' ��վ�����
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

'�ж�msflexgrid �Ƿ�Ϊ��
Function IsGridEmpty(grid As MSFlexGrid) As Boolean
    IsGridEmpty = True
    Dim r As Long, c As Long
    With grid
        For r = .FixedRows To .Rows - 1       ' ����������
            For c = .FixedCols To .Cols - 1    ' ����������
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
        
    ' �����ֵ�������洢��һ�������е�Ԫ��
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' ��������ַ���Ϊ����
    If InStr(str1, ",") > 0 Then
        arr1 = Split(str1, ",")
    Else
        If str1 <> "" Then
            ReDim arr1(0)
            arr1(0) = str1
        Else
            ReDim arr1(-1 To 0) ' ������
        End If
    End If

    If InStr(str2, ",") > 0 Then
        arr2 = Split(str2, ",")
    Else
        If str2 <> "" Then
            ReDim arr2(0)
            arr2(0) = str2
        Else
            ReDim arr2(-1 To 0) ' ������
        End If
    End If

    ' ����һ�������Ԫ�ؼ����ֵ�
    For i = LBound(arr1) To UBound(arr1)
        'Dim item As String
        item = Trim(arr1(i))
        If item <> "" Then
            If Not dict.Exists(item) Then
                dict.Add item, Nothing
            End If
        End If
    Next i

    ' �����ڶ������飬���Ҵ������ֵ��е����������
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

    ' ���ؽ��
    IntersectExtensions = result

    ' ����
    Set dict = Nothing
End Function

' �ݹ鴴���༶Ŀ¼��֧��·���а�����������ڵ��ļ��У�
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

' �����ļ���Ŀ��·�����ɹ����� True��ʧ�ܷ��� False
Function CopyFileToTarget(srcPath As String, destPath As String) As Boolean
    On Error GoTo ErrHandler
    FileCopy srcPath, destPath
    CopyFileToTarget = True
    Exit Function

ErrHandler:
    CopyFileToTarget = False
End Function

' �ƶ��ļ������ƺ�ɾ��Դ�ļ������ɹ����� True��ʧ�ܷ��� False
Function MoveFileToTarget(srcPath As String, destPath As String) As Boolean
    On Error GoTo ErrHandler
    FileCopy srcPath, destPath
    Kill srcPath
    MoveFileToTarget = True
    Exit Function
ErrHandler:
    MoveFileToTarget = False
End Function

' ��������¼д����־�ļ�
Sub LogFileAction(srcPath As String, destPath As String, success As Boolean, logFile As TextStream)
    Dim action As String
    If success Then
        action = "�ɹ�"
    Else
        action = "ʧ��"
    End If
    'logFile.WriteLine Now & " | " & action & " | Դ·��: " & srcPath & " | Ŀ��·��: " & destPath
    logFile.WriteLine " | " & action & " | Դ·��: " & srcPath & " | Ŀ��·��: " & destPath
End Sub

'' ���½�����־��ʾ��������һ�� ListBox ����ʾ��
Sub UpdateLogDisplay(text As String)
    Form1.lstlog.AddItem text
    Form1.lstlog.ListIndex = Form1.lstlog.NewIndex ' ���������һ��
End Sub

' �����ļ���������Ŀ��·��
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

    ' ȷ��Ŀ��·������
    If Dir(BuildTargetPath, vbDirectory) = "" Then
        MkDirRecursive BuildTargetPath
    End If
End Function

' �ж�Ŀ��·�����Ƿ����ͬ���ļ��������غ��ʵ��ļ�·������ͨ���ļ���С�жϣ�
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
'    ' ��ȡԴ�ļ������������ƺ���չ��
'    srcFileName = Mid$(srcPath, InStrRev(srcPath, "\") + 1)
'    baseName = Left$(srcFileName, InStrRev(srcFileName, ".") - 1)
'    ext = Mid$(srcFileName, InStrRev(srcFileName, "."))
'
'    ' ������ʼ��Ŀ���ļ�·��
'    GetDestinationFilePath = targetPath & "\" & srcFileName
'
'    ' ����ļ������ڣ�ֱ�ӷ���
'    If Dir(GetDestinationFilePath) = "" Then Exit Function
'
'    ' ��ȡĿ���ļ���Դ�ļ��Ĵ�С
'    existingSize = FileLen(GetDestinationFilePath)
'    currentSize = FileLen(srcPath)
'
'    ' �����Сһ�£���Ϊ���ظ��ļ����ƶ�������վĿ¼
'    If existingSize = currentSize Then
'        ' ���û���վĿ¼Ϊ��Ŀ���ļ��и�Ŀ¼�µġ�ͬ���ļ�����վ��
'        recoveryDir = BuildRecoveryPath(targetPath)
'
'        ' ���ػ���վĿ¼�µ��ļ�·��
'        GetDestinationFilePath = recoveryDir & "\" & srcFileName
'        Exit Function
'    End If
'
'    ' ��������� -2��-3 �Ⱥ�׺��������
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
'    MsgBox "��������" & Err.Description
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

    ' ���� ADODB.Stream �� MD5 ���ܶ���
    Set objStream = CreateObject("ADODB.Stream")
    Set objMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

    ' �����������ļ�
    objStream.Type = 1 ' adTypeBinary
    objStream.Open
    objStream.LoadFromFile filePath

    ' �����ϣֵ
    hashBytes = objMD5.ComputeHash_2(objStream.Read)

    ' �ж��Ƿ�ɹ���ȡ����ϣ�ֽ�����
    If Not isEmpty(hashBytes) And TypeName(hashBytes) = "Byte()" Then
        ' ����ÿ���ֽڣ�ת��Ϊʮ�������ַ���
        For i = LBound(hashBytes) To UBound(hashBytes)
            hexHash = hexHash & Right$("0" & Hex(AscB(MidB(hashBytes, i + 1, 1))), 2)
        Next i
    Else
        hexHash = ""
    End If

    ' ����Сд�Ĺ�ϣֵ
    GetFileHash = LCase(hexHash)
    Exit Function

ErrorHandler:
    GetFileHash = ""
    MsgBox "�����ļ���ϣ����" & Err.Description, vbCritical
End Function
' �ж�Ŀ��·�����Ƿ����ͬ���ļ��������غ��ʵ��ļ�·���������ļ���С�жϣ�
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

    ' ��ȡԴ�ļ������������ƺ���չ��
    srcFileName = Mid$(srcPath, InStrRev(srcPath, "\") + 1)
    If InStrRev(srcFileName, ".") > 0 Then
        baseName = Left$(srcFileName, InStrRev(srcFileName, ".") - 1)
        ext = Mid$(srcFileName, InStrRev(srcFileName, "."))
    Else
        baseName = srcFileName
        ext = ""
    End If

    ' ������ʼ��Ŀ���ļ�·��
    GetDestinationFilePath = targetPath & "\" & srcFileName

    ' ���Ŀ��·���²����ڸ��ļ���ֱ�ӷ���
    If Dir(GetDestinationFilePath) = "" Then Exit Function

    ' ���ݱȽϷ�ʽѡ���ж��߼�
    If useSizeCompare Then
        ' ʹ���ļ���С�ȶ�
        currentSize = FileLen(srcPath)
        existingSize = FileLen(GetDestinationFilePath)
        
        If existingSize = currentSize Then
            recoveryDir = BuildRecoveryPath(targetPath)
            GetDestinationFilePath = recoveryDir & "\" & srcFileName
            Exit Function
        End If
    Else
        ' ʹ�ù�ϣ�ȶ�
        currentHash = GetFileHash(srcPath)
        existingHash = GetFileHash(GetDestinationFilePath)
        
        If existingHash = currentHash And existingHash <> "" Then
            recoveryDir = BuildRecoveryPath(targetPath)
            GetDestinationFilePath = recoveryDir & "\" & srcFileName
            Exit Function
        End If
    End If

    ' ��������� -2��-3 �Ⱥ�׺������������ÿ�μ���Ƿ��ظ�
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
            
            ' ���ظ�����������Ӽ�����
            counter = counter + 1
        End If
    Loop

Exit Function

ErrorHandler:
    MsgBox "��������" & Err.Description, vbCritical
    GetDestinationFilePath = ""
End Function

' ��������վĿ¼·����ȷ�������
Function BuildRecoveryPath(basePath As String) As String
    Dim recoveryPath As String
    recoveryPath = Form1.txtDestPath.text & "\ͬ���ļ�����վ"
    
    ' ��鲢��������վĿ¼
    If Dir(recoveryPath, vbDirectory) = "" Then
        MkDir recoveryPath
    End If
    
    BuildRecoveryPath = recoveryPath
End Function

'��ʾ�ֵ������
Function DisplayDictionaryContent(dict As Dictionary)
    Dim key As Variant
    Dim content As String
   ' Dim Shuliang As Long
    
    content = "��Ҫ������ļ�����:"
    ShuLiang = 0
    
'    If dict Is Nothing Then
'        MsgBox "�ֵ�Ϊ�ջ�δ��ʼ��", vbExclamation
'        Exit Sub
'    End If
    
    If dict.Count = 0 Then
        'Form1.txtResult.text = " û�з�����������Ҫ������ļ��� "
        Form1.lblStatus.Caption = " û�з�����������Ҫ������ļ��� "
        Exit Function
    End If
    
    For Each key In dict.Keys
'        content = content & "��: " & key & ", ֵ: " & dict(key) & vbCrLf
         content = content & key & " "
         ShuLiang = ShuLiang + dict(key)
        ' MsgBox key & " " & dict(key)
    Next key
    
'    MsgBox content, vbInformation
    'Form1.txtResult.text = content
     Form1.lblStatus.Caption = content & "�����ļ�����" & ShuLiang
End Function

'��ʾ�ֵ������
Public Sub DisplayDictionaryContentMsg(dict As Dictionary)
    Dim key As Variant
    Dim content As String
    
    content = "��Ҫ������ļ�����:" & vbCrLf
    
    If dict Is Nothing Then
        MsgBox "�ֵ�Ϊ�ջ�δ��ʼ��", vbExclamation
        Exit Sub
    End If
    
    If dict.Count = 0 Then
        MsgBox " û�з�����������Ҫ������ļ��� "
        Exit Sub
    End If
    
    For Each key In dict.Keys
         content = content & "��: " & key & ", ֵ: " & dict(key) & vbCrLf
         'content = content & key & ", "
    Next key
    
    MsgBox content, vbInformation
    'Form1.txtResult.text = content
End Sub

