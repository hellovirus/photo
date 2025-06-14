VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "图片音视频整理大师 - 52pojie - hellovirus"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14010
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   14010
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   29
      Text            =   "2000"
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CheckBox chkIncludeSubFolders 
      Caption         =   "包含子文件夹"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   27
      Top             =   120
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.ListBox lstlog 
      Height          =   1320
      Left            =   120
      TabIndex        =   24
      Top             =   6720
      Width           =   5535
   End
   Begin VB.CommandButton btnSelectFolder 
      BackColor       =   &H00C0FFFF&
      Caption         =   "选择 源文件夹："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C0FFC0&
      Caption         =   "开始处理"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6720
      Width           =   1130
   End
   Begin VB.Frame Frame5 
      Caption         =   "同名文件处理 ： "
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   5280
      Width           =   5295
      Begin VB.OptionButton Option10 
         Caption         =   " 首  选： 比较文件大小 (执行速度快)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   4575
      End
      Begin VB.OptionButton Option11 
         Caption         =   " 不推荐： 比较文件MD5值(执行速度慢)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   4575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " 时间标签选择 ： "
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   9375
      Begin VB.OptionButton Option9 
         Caption         =   "按照 exif 时间、文件名包含时间、修改日期、创建日期 四项中的最早时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   8895
      End
      Begin VB.OptionButton Option8 
         Caption         =   "首选exif时间，如无选文件名包含时间 、修改日期、创建日期三项中的最早时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   8895
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " 操作方式 ： "
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   6760
      Begin VB.OptionButton Option6 
         Caption         =   " 复制 （保留源文件）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton Option7 
         Caption         =   " 移动 （删除源文件）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   13
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " 目标文件夹格式 ： "
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   9375
      Begin VB.OptionButton Option5 
         Caption         =   "年月日：如 2006/04/22"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
      Begin VB.OptionButton Option4 
         Caption         =   " 年月：如 2017/06"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton Option3 
         Caption         =   " 年：如 2024"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "选择需要处理的文件类型"
      Height          =   7815
      Left            =   9600
      TabIndex        =   3
      Top             =   240
      Width           =   4335
      Begin VB.CheckBox Check1 
         Caption         =   "反选"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   28
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "全选"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "自己按需选择："
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "默  认： 常用图片,视频,音频"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   $"Form1.frx":3AFA
         Top             =   360
         Value           =   -1  'True
         Width           =   3975
      End
      Begin MSFlexGridLib.MSFlexGrid fgFiles 
         Height          =   6015
         Left            =   60
         TabIndex        =   26
         Top             =   1320
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   10610
         _Version        =   393216
         Enabled         =   0   'False
         HighLight       =   0
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "详细情况栏"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   7400
         Width           =   4095
      End
   End
   Begin VB.TextBox txtDestPath 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   7335
   End
   Begin VB.CommandButton cmdSelectDest 
      BackColor       =   &H00FFFFC0&
      Caption         =   "选择目标文件夹："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtSourcePath 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   7080
      Picture         =   "Form1.frx":3B74
      ScaleHeight     =   3975
      ScaleWidth      =   2415
      TabIndex        =   25
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "年"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6660
      TabIndex        =   31
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "照片时间不早于："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   30
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "请选择 源文件夹 和 目标文件夹 ..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   8160
      Width           =   13935
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' 引用：
' 1. Microsoft Scripting Runtime
' 2. Microsoft Shell Controls And Automation
' 3. ActiveX Data Objects 2.8 Library

' API声明
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Const STR_SORT_ASC = "↑"
Private Const STR_SORT_DESC = "↓"

Private Const FuhaoXuanzhe = "√"
Private Const FuhaoWeixuanzhe = "□"


Private Type BROWSEINFO
    hWndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

' 排序相关变量
Private m_lngSortColumn As Long    ' 当前排序列
Private m_bolSortAscending As Boolean  ' 排序方向
Private m_arrColumnType() As String    ' 列数据类型

' 选择状态数组
Private m_arrSelected() As Boolean  ' 记录每一行的选择状态

'Dim TotalFilesNum As Long

Public TotalFilesNum As Long      ' 总文件数

Private SucFileLog As String  '定义成功日志路径
Private LostFileLog As String  '定义失败日志路径

Private FilePathOld As String
Private FileNameOld As String

Private FileNum100 As Integer


Private Sub UpdateUIExtensionTable()
    Dim i As Long
    Dim key As Variant

    ' 清空表格
    fgFiles.Rows = 1  ' 至少保留表头
    i = 1

    ReDim m_arrSelected(1 To gDictAllExt.Count)

    For Each key In gDictAllExt.Keys
        If i >= fgFiles.Rows Then
            fgFiles.Rows = fgFiles.Rows + 1
        End If

        fgFiles.TextMatrix(i, 0) = "□"
        fgFiles.TextMatrix(i, 1) = i
        fgFiles.TextMatrix(i, 2) = key
        fgFiles.TextMatrix(i, 3) = gDictAllExt(key)

        m_arrSelected(i) = False

        i = i + 1
    Next key

    Option1.Value = True
    Option2.Enabled = True
End Sub

' 新增函数：根据路径和是否包含子文件夹来索引文件
Public Sub IndexFilesInFolder(ByVal folderPath As String, Optional ByVal includeSubFolders As Boolean = True)
    Dim fso As New FileSystemObject
    Dim objFolder As folder
    
    Set objFolder = fso.GetFolder(folderPath)
    
    ' 清空字典
    If Not gDictAllExt Is Nothing Then
        If gDictAllExt.Count > 0 Then
            gDictAllExt.RemoveAll
        End If
    Else
        Set gDictAllExt = New Dictionary
    End If

    TotalFilesNum = 0

    ' 调用 TraverseFolder，并传入 includeSubFolders 参数
    TraverseFolder fso, objFolder, gDictAllExt, TotalFilesNum, includeSubFolders


    Label1.Caption = "选中文件夹共有" & TotalFilesNum & "个文件! "

    If gDictAllExt.Count = 0 Then
        MsgBox "所选择文件夹，为空文件夹，退出..."
        Exit Sub
    End If

    ' 更新界面表格
    UpdateUIExtensionTable

    ' 其他 UI 状态更新
    CheckAndEnableStartButton
    BuildGlobalDictIntersect
    
    DisplayDictionaryContent gDictUserUse
    
    
    UpdateSelectMorenStatus
End Sub


'Private Sub btnSelectFolder_Click()
'    Dim folderPath As String
'
'    TotalFilesNum = 0
'
'    folderPath = BrowseForFolder("请选择要处理的源文件夹")
'    If folderPath = "" Then
'        MsgBox "未选择文件夹...", vbCritical
'        Exit Sub
'    End If
'
'    txtSourcePath.text = folderPath
'
'    ' 清空表格
'    fgFiles.Rows = 1  ' 至少保留表头和1行数据
'
'    Dim fso As New FileSystemObject
'    Dim objFolder As folder
'    Set objFolder = fso.GetFolder(folderPath)
'
'    If Not gDictAllExt Is Nothing Then
'        If gDictAllExt.Count > 0 Then
'            gDictAllExt.RemoveAll
'        End If
'    Else
'        Set gDictAllExt = New Dictionary ' 如果字典未初始化，则重新初始化
'    End If
'
'    'TraverseFolder fso, objFolder, gDictAllExt, TotalFilesNum,
'     TraverseFolder fso, objFolder, gDictAllExt, TotalFilesNum, Nothing, chkIncludeSubFolders.Value
'
'    Label1.Caption = "选中文件夹共有" & TotalFilesNum & "个文件! "
'
'    If gDictAllExt.Count = 0 Then
'       MsgBox "所选择文件夹，为空文件夹，退出..."
'       Exit Sub
'    End If
'    ' 初始化选择状态数组
'    ReDim m_arrSelected(1 To gDictAllExt.Count)
'
'    ' 填充表格
'    Dim key As Variant
'    Dim i As Long
'    i = 1  ' 从第1行开始（0行是表头）
'
'    For Each key In gDictAllExt.Keys
'        ' 确保有足够的行
'        If i >= fgFiles.Rows Then
'            fgFiles.Rows = fgFiles.Rows + 1
'        End If
'
'        fgFiles.TextMatrix(i, 0) = "□"  ' 选择框（初始为未选中）
'        fgFiles.TextMatrix(i, 1) = i    ' 序号
'        fgFiles.TextMatrix(i, 2) = key  ' 扩展名
'        fgFiles.TextMatrix(i, 3) = gDictAllExt(key)  ' 数量
'
'        ' 初始化选择状态
'        m_arrSelected(i) = False
'
'        i = i + 1
'    Next key
'
'    ' 重置全选框状态
'    'chkSelectAll.Value = 1
'
'    Option1.Value = True
'
'    Option2.Enabled = True
'
'    CheckAndEnableStartButton
'    '定义好默认文件后缀名与所有文件名的交集
'    BuildGlobalDictIntersect
'
'    DisplayDictionaryContent gDictUserUse
'
'    UpdateSelectMorenStatus
'
'    'UpdateFileCount
'
'End Sub

Private Sub btnSelectFolder_Click()
    Dim folderPath As String
    
    TotalFilesNum = 0


    folderPath = BrowseForFolder("请选择要处理的源文件夹")
    If folderPath = "" Then
        MsgBox "未选择文件夹...", vbCritical
        Exit Sub
    End If

    txtSourcePath.text = folderPath
    chkIncludeSubFolders.Enabled = True

    ' 调用新函数进行索引
    IndexFilesInFolder folderPath, chkIncludeSubFolders.Value
    
End Sub


Private Sub Check1_Click()
    Dim i As Long
    If Not IsGridEmpty(fgFiles) Then
        For i = 1 To fgFiles.Rows - 1
            m_arrSelected(i) = Not m_arrSelected(i)
            
            If m_arrSelected(i) Then
                fgFiles.TextMatrix(i, 0) = FuhaoXuanzhe
            Else
                fgFiles.TextMatrix(i, 0) = FuhaoWeixuanzhe
            End If
        Next i
        
        UpdateFileCount
        DisplayDictionaryContent gDictUserUse
    End If
End Sub

Private Sub chkIncludeSubFolders_Click()
    If txtSourcePath.text <> "" Then
        IndexFilesInFolder txtSourcePath.text, chkIncludeSubFolders.Value
    End If
End Sub

Private Sub chkSelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
   Dim i As Long
    If Not IsGridEmpty(fgFiles) Then
        For i = 1 To fgFiles.Rows - 1
            m_arrSelected(i) = (chkSelectAll.Value = 1)
            
            ' 更新显示
            If m_arrSelected(i) Then
'            If chkSelectAll.Value Then
                fgFiles.TextMatrix(i, 0) = FuhaoXuanzhe
            Else
                fgFiles.TextMatrix(i, 0) = FuhaoWeixuanzhe
            End If
        Next i
        ' 更新结果文本
        'UpdateResultText
        UpdateFileCount
        DisplayDictionaryContent gDictUserUse
    End If

End Sub

Private Sub fgFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then  ' 左键点击
        If Not IsGridEmpty(fgFiles) Then
            ' 如果点击的是选择框列
            If fgFiles.Col = 0 And y > fgFiles.RowHeight(0) And y < fgFiles.RowHeight(0) * fgFiles.Rows Then
                ' 切换选择状态
                m_arrSelected(fgFiles.Row) = Not m_arrSelected(fgFiles.Row)

                ' 更新显示
                If m_arrSelected(fgFiles.Row) Then
                    fgFiles.TextMatrix(fgFiles.Row, 0) = FuhaoXuanzhe
                Else
                    fgFiles.TextMatrix(fgFiles.Row, 0) = FuhaoWeixuanzhe
                End If
                
            End If
        End If
        
        Sort fgFiles, y ' 呼叫msflexgrid 排序 函数
        
        ' 统一在这里更新全选框状态
        UpdateSelectAllStatus
        UpdateFileCount
        DisplayDictionaryContent gDictUserUse
    End If
End Sub

Private Sub Form_Load()
    ' 设置窗体基本属性
   ' Me.Caption = "图片音视频按日期分类工具"
    
    ' 初始化FlexGrid
    With fgFiles
        .Cols = 4  ' 4列：选择框、序号、扩展名、数量
        .FixedCols = 0  ' 没有固定列
        .FixedRows = 1 ' 1个固定行（表头）
        
        ' 设置表头
        .TextMatrix(0, 0) = "选择"
        .TextMatrix(0, 1) = "序号"
        .TextMatrix(0, 2) = "扩展名"
        .TextMatrix(0, 3) = "数量"
        
        ' 设置列宽
        .ColWidth(0) = 700  ' 选择框列
        .ColWidth(1) = 900  ' 序号列
        .ColWidth(2) = 1200 ' 扩展名列
        .ColWidth(3) = 1050  ' 数量列
        
    End With
    
    ' 初始化列数据类型
    ReDim m_arrColumnType(0 To fgFiles.Cols - 1)
    m_arrColumnType(0) = "text"  ' 选择框列
    m_arrColumnType(1) = "text"  ' 序号列
    m_arrColumnType(2) = "text"  ' 扩展名列
    m_arrColumnType(3) = "number" ' 数量列
    
    ' 初始化文本框
    'txtResult.text = "默认类型：BMP，JPG，PNG，TIF，GIF，PCX，TGA，MP4、AVI、MOV、MKV、FLV、WMV，MPEG，3GP，MP3，WMA，WAV"
    
    InitGlobalDictAllMedia ' 初始化 默认文件类型 字典
    ' DisplayDictionaryContent gDictAllMedia
    ZuiZaoDate = CInt(Text1.text)
    
End Sub

' 源文件夹选择按钮事件处理
Private Sub cmdSelectSource_Click()
    Dim folderPath As String
    folderPath = BrowseForFolder("请选择要处理的源文件夹")
    If folderPath <> "" Then
        txtSourcePath.text = folderPath
        CheckAndEnableStartButton
    End If
End Sub

' 目标文件夹选择按钮事件处理
Private Sub cmdSelectDest_Click()
    Dim folderPath As String
    folderPath = BrowseForFolder("请选择文件分类的目标文件夹")
    If folderPath <> "" Then
        txtDestPath.text = folderPath
        'recoveryPath = folderPath & "\同名文件回收站"
        CheckAndEnableStartButton

 ' 创建日志文件（如果不存在）
        CreateLogFilesIfNotExists folderPath
    End If
End Sub

' 检查并创建日志文件
Sub CreateLogFilesIfNotExists(folderPath As String)
   
    SucFileLog = folderPath & "\执行日志.txt"
    'LostFileLog = folderPath & "\失败日志.txt"
    
    ' 创建“成功日志.txt”（如果不存在）
    If Dir(SucFileLog) = "" Then
        Open SucFileLog For Output As #1
    Else
        Open SucFileLog For Append As #1
    End If
        If txtSourcePath.text <> "" Then Print #1, vbCrLf & "源文件夹：" & txtSourcePath.text & "，文件执行日志 - " & Now()
        Close #1
    
    ' 创建“失败日志.txt”（如果不存在）
'    If Dir(LostFileLog) = "" Then
'        Open LostFileLog For Output As #1
'    Else
'        Open LostFileLog For Append As #1
'    End If
'        If txtSourcePath.text <> "" Then Print #1, vbCrLf & "源文件夹：" & txtSourcePath.text & "，文件执行失败日志 - " & Now()
'        Close #1
End Sub

' 检查并启用开始按钮
Private Sub CheckAndEnableStartButton()
    cmdStart.Enabled = (txtSourcePath.text <> "" And txtDestPath.text <> "")
End Sub

' 开始处理按钮事件处理
Private Sub cmdStart_Click()
    Dim fso As New FileSystemObject
    Dim sourceFolder As folder
    Dim logFile As TextStream
    Dim processedFiles As Long
    Dim copiedFiles As Long
    Dim startTime As Date
    
    Dim fileExtensions() As String ' 存储允许的文件后缀名
    'Dim extensionList As String
    
    ' 检查源文件夹是否存在
    If Not fso.FolderExists(txtSourcePath.text) Then
        MsgBox "源文件夹不存在!", vbExclamation
        Exit Sub
    End If
    
    ' 检查目标文件夹是否存在，不存在则创建
    If Not fso.FolderExists(txtDestPath.text) Then
        On Error Resume Next
        fso.CreateFolder txtDestPath.text
        If Err.Number <> 0 Then
            MsgBox "无法创建目标文件夹: " & Err.Description, vbCritical
            Exit Sub
        End If
        On Error GoTo 0
    End If
    
    ' 打开或创建日志文件
    On Error Resume Next
    Set logFile = fso.CreateTextFile(SucFileLog, True)
    If Err.Number <> 0 Then
        MsgBox "无法创建或打开日志文件: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    cmdStart.Enabled = False
    FileNum100 = 0
    'FileNum10 = 0


    ' 写入日志头
    logFile.WriteLine "========== 文件分类开始: " & Now() & " =========="
    logFile.WriteLine "源文件夹: " & txtSourcePath.text
    logFile.WriteLine "目标文件夹: " & txtDestPath.text
    logFile.WriteLine
    
    ' 初始化计数器与时间
    startTime = Now()
    processedFiles = 0
    copiedFiles = 0
    
    If ShuLiang > 100 Then
        FileNum100 = ShuLiang \ 100
    ElseIf ShuLiang > 10 Then
        FileNum100 = ShuLiang \ 10
    Else
        FileNum100 = 1
    End If
    
    ' 设置状态提示
    lblStatus.Caption = "正在处理文件..."
    lblStatus.Refresh
    
    ' 获取源文件夹对象
    Set sourceFolder = fso.GetFolder(txtSourcePath.text)
    
    If Option3.Value = True Then ProcessFolder sourceFolder, logFile, processedFiles, copiedFiles, 1, Option6.Value
    If Option4.Value = True Then ProcessFolder sourceFolder, logFile, processedFiles, copiedFiles, 2, Option6.Value
    If Option5.Value = True Then ProcessFolder sourceFolder, logFile, processedFiles, copiedFiles, 3, Option6.Value
    ' 递归处理所有文件
    'ProcessFolder sourceFolder, logFile, processedFiles, copiedFiles, 2, True, fileExtensions
    
    ' 关闭日志文件并写入统计信息
    logFile.WriteLine
    logFile.WriteLine "========== 文件分类完成 =========="
    logFile.WriteLine "处理文件总数: " & processedFiles
    logFile.WriteLine "成功复制文件数: " & copiedFiles
    logFile.WriteLine "处理用时: " & Format(Now() - startTime, "hh:mm:ss")
    logFile.Close
    
    ' 更新界面状态显示
    lblStatus.Caption = "处理完成! 共处理 " & processedFiles & " 个文件，成功复制 " & copiedFiles & " 个文件。用时: " & Format(Now() - startTime, "hh:mm:ss")
    MsgBox lblStatus.Caption
    cmdStart.Enabled = True
    
End Sub

'
' 递归处理文件夹
' 参数说明：
'   folder - 要处理的文件夹对象
'   logFile - 日志流对象
'   processedFiles - 已处理文件数（引用传递）
'   copiedFiles - 已复制/剪切文件数（引用传递）
'   targetStructure - 目标路径结构（1=年；2=年/月；3=年/月/日）
'   isCopy - True=复制，False=剪切
'
Public Sub ProcessFolder( _
    folder As Object, _
    logFile As TextStream, _
    ByRef processedFiles As Long, _
    ByRef copiedFiles As Long, _
    ByVal targetStructure As Integer, _
    ByVal isCopy As Boolean)
    
    Dim subFolder As Object     ' 子文件夹对象
    Dim file As Object          ' 当前遍历的文件对象
    Dim fileDate As Date        ' 文件的时间信息
    Dim targetPath As String    ' 构建的目标路径
    Dim success As Boolean      ' 操作是否成功
    Dim destFilePath As String  ' 目标文件完整路径
    
    Dim fileYear As Integer     ' 文件所属年份（替代 year）
    Dim fileMonth As Integer    ' 文件所属月份（替代 month）
    Dim fileDay As Integer      ' 文件所属日期（替代 day）
    Dim fileExt As String       ' 文件扩展名（小写）

    'Dim FilePathOld As String
    
    ' 遍历当前文件夹中的所有文件
    For Each file In folder.Files
        
        FilePathOld = file.Path
        FileNameOld = file.Name
        
        ' 获取文件扩展名，并转换为小写以便比较
        fileExt = LCase$(fso.GetExtensionName(FilePathOld))
        
        ' 如果没有扩展名，跳过该文件
        If fileExt = "" Then
            GoTo SkipThisFile
        End If
        
        ' 检查扩展名是否在允许的字典中
        If Not gDictUserUse.Exists(fileExt) Then
            GoTo SkipThisFile
        End If

        ' 增加已处理文件计数
        processedFiles = processedFiles + 1

        ' 每处理10个文件更新一次状态栏（提升用户体验）
'        If processedFiles Mod 10 = 0 Then
        If processedFiles Mod FileNum100 = 0 Then
            lblStatus.Caption = "正在处理文件: " & processedFiles & "，已处理: " & copiedFiles
            lblStatus.Refresh  ' 可选刷新界面
        End If

        ' 获取文件时间信息（优先使用 EXIF）
'        If GetFileDateParts(FilePathOld, True, fileYear, fileMonth, fileDay) Then
        If GetFileDateParts(FilePathOld, Option8.Value, fileYear, fileMonth, fileDay) Then
            fileDate = DateSerial(fileYear, fileMonth, fileDay)
        Else
            ' 如果获取失败，使用文件最后修改时间作为备用
            fileDate = FileDateTime(FilePathOld)
        End If

        ' 根据文件日期构建目标路径
        targetPath = BuildTargetPath(fileDate, txtDestPath.text, targetStructure)

        ' 获取目标文件完整路径（包括文件名）
        destFilePath = GetDestinationFilePath(FilePathOld, targetPath, Option10.Value)

        ' 执行复制或剪切操作
        If isCopy Then
            ' 复制文件
            success = CopyFileToTarget(FilePathOld, destFilePath)
        Else
            ' 剪切文件（先复制再删除原文件）
            success = MoveFileToTarget(FilePathOld, destFilePath)
        End If

        ' 记录日志
        LogFileAction FilePathOld, destFilePath, success, logFile

        ' 每处理5个文件更新一次日志显示（可选调试输出）
'        If processedFiles Mod 5 = 0 Then
        If processedFiles Mod FileNum100 = 0 Then
            UpdateLogDisplay "处理文件: " & FileNameOld & " --> " & IIf(success, "成功", "失败")
        End If

        ' 如果操作成功，增加成功计数
        If success Then copiedFiles = copiedFiles + 1

SkipThisFile:
    Next file

    ' 递归处理子文件夹
    On Error Resume Next
    For Each subFolder In folder.SubFolders
        ProcessFolder subFolder, logFile, processedFiles, copiedFiles, targetStructure, isCopy
    Next subFolder
End Sub

' 显示文件夹选择对话框
Private Function BrowseForFolder(title As String) As String
    Dim objShell As Object
    Dim objFolder As Object
    
    On Error Resume Next
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, title, 0)
    
    If Not objFolder Is Nothing Then
        BrowseForFolder = objFolder.Self.Path
    Else
        BrowseForFolder = ""
    End If
    On Error GoTo 0
End Function

' 显示保存文件对话框
Private Function ShowSaveDialog(title As String, filter As String) As String
    Dim dlg As Object
    
    On Error Resume Next
    Set dlg = CreateObject("MSComDlg.CommonDialog")
    
    With dlg
        .DialogTitle = title
        .filter = filter
        .Flags = &H800 ' OFN_OVERWRITEPROMPT
        .ShowSave
        
        If Err.Number = 0 Then
            ShowSaveDialog = .fileName
        Else
            ShowSaveDialog = ""
        End If
    End With
    On Error GoTo 0
End Function


'排序函数
Public Sub Sort(sgrd As MSFlexGrid, y As Single)
    With sgrd
    ' 检查是否在网格区域内
        If y >= .RowPos(0) And y < (.RowPos(0) + .RowHeight(0)) Then
            If .Tag <> "" Then
            'If .Tag <> "" And CInt(.Tag) <> .Col Then
                If .Tag <> .Col Then
                .TextMatrix(0, .Tag) = Left(.TextMatrix(0, .Tag), Len(.TextMatrix(0, .Tag)) - 1)
                End If
            End If
        
            If Right(.TextMatrix(0, .Col), 1) = STR_SORT_ASC Then
                .Sort = flexSortGenericAscending
                .TextMatrix(0, .Col) = Replace(.TextMatrix(0, .Col), STR_SORT_ASC, STR_SORT_DESC)
            ElseIf Right(.TextMatrix(0, .Col), 1) = STR_SORT_DESC Then
                .Sort = flexSortGenericDescending
                .TextMatrix(0, .Col) = Replace(.TextMatrix(0, .Col), STR_SORT_DESC, STR_SORT_ASC)
            Else
            .Sort = flexSortGenericDescending
            .TextMatrix(0, .Col) = .TextMatrix(0, .Col) & STR_SORT_ASC
            End If
        .Tag = CStr(.Col)
        End If
    End With
    
End Sub

' 递归遍历文件夹，统计文件扩展名、空文件夹和文件总数
Private Sub TraverseFolder(fso As FileSystemObject, _
                            ByVal objFolder As folder, _
                            dictExtensions As Object, _
                            Optional ByRef totalFiles As Long = 0, _
                            Optional ByVal includeSubFolders As Boolean = True, _
                            Optional dictEmptyFolders As Object = Nothing)

    Dim isEmpty As Boolean
    
    ' 检查文件夹是否为空（无文件和无子文件夹）
    isEmpty = (objFolder.Files.Count = 0 And objFolder.SubFolders.Count = 0)
    
    ' 如果提供了空文件夹字典，则记录空文件夹
    If Not dictEmptyFolders Is Nothing And isEmpty Then
        dictEmptyFolders.Add objFolder.Path, True
    End If
    
    ' 遍历当前文件夹中的文件并统计扩展名
    Dim objFile As file
    Dim strExt As String
    
    For Each objFile In objFolder.Files
        strExt = LCase$(fso.GetExtensionName(objFile.Path))
        
        If strExt <> "" Then
            If dictExtensions.Exists(strExt) Then
                dictExtensions(strExt) = dictExtensions(strExt) + 1
            Else
                dictExtensions.Add strExt, 1
            End If
        End If
        
        ' 累计文件总数
        totalFiles = totalFiles + 1
        If totalFiles Mod 10 = 0 Then
            lblStatus.Caption = "已统计文件数: " & totalFiles
            lblStatus.Refresh  ' 可选刷新界面
        End If
    Next objFile
    
    ' 如果允许递归子文件夹，则继续深入
    If includeSubFolders Then
        Dim objSubFolder As folder
        For Each objSubFolder In objFolder.SubFolders
            TraverseFolder fso, objSubFolder, dictExtensions, totalFiles, includeSubFolders
        Next objSubFolder
    End If
End Sub
' 获取文件夹路径
Private Function GetFolder(hWnd As Long, strTitle As String) As String
    Dim bi As BROWSEINFO
    With bi
        .hWndOwner = hWnd
        .lpszTitle = strTitle
        .ulFlags = &H1 ' 允许选择文件夹
    End With
    
    Dim pidl As Long
    pidl = SHBrowseForFolder(bi)
    If pidl = 0 Then Exit Function
    
    Dim buffer As String
    buffer = String$(512, 0)
    If SHGetPathFromIDList(pidl, buffer) Then
        GetFolder = Left$(buffer, InStr(buffer, vbNullChar) - 1)
    End If
End Function

Private Sub UpdateSelectAllStatus()
    Dim i As Long
    Dim allSelected As Boolean
    
    allSelected = True
    
    For i = 1 To fgFiles.Rows - 1
        If fgFiles.TextMatrix(i, 0) = FuhaoWeixuanzhe Then
            allSelected = False
            Exit For
        End If
    Next i
    
    chkSelectAll.Value = IIf(allSelected, 1, 0)
End Sub
' 更新默认和全部交集的选中状态
Private Sub UpdateSelectMorenStatus()
   Dim i As Long
   Dim XuanzhongFiles As Long
   
   XuanzhongFiles = 0
   
    If Not IsGridEmpty(fgFiles) Then
        For i = 1 To fgFiles.Rows - 1
            'm_arrSelected(i) = (chkSelectAll.Value = 1)
            
            ' 更新显示
            If gDictUserUse.Exists(fgFiles.TextMatrix(i, 2)) Then
                fgFiles.TextMatrix(i, 0) = FuhaoXuanzhe
                 m_arrSelected(i) = True
                 

            Else
                fgFiles.TextMatrix(i, 0) = FuhaoWeixuanzhe
                 m_arrSelected(i) = False
            End If
        
        Next i
    End If
End Sub


' 更新结果文本 - 按行显示选中内容
Private Sub UpdateResultText()
    Dim strResult As String
    Dim i As Long
    
    strResult = "需要处理的文件类型:"
    
    For i = 1 To fgFiles.Rows - 1
        If m_arrSelected(i) Then
            ' 使用换行符连接选中项
            strResult = strResult & fgFiles.TextMatrix(i, 2) & " "
        End If
    Next i
    
    'txtResult.text = strResult
    lblStatus.Caption = strResult
    UpdateGlobalDictUserSel (strResult)
  '  DisplayDictionaryContent gDictUserUse
 
End Sub


' 更新文件总数显示
Private Sub UpdateFileCount()
    Dim SelFiles As Long
    Dim SelFilesExt As Long
    Dim i As Long
    Dim StrSelExt As String
    
    
    SelFiles = 0
    SelFilesExt = 0
        ' 清空旧数据
    If Not gDictUserUse Is Nothing Then
        If gDictUserUse.Count > 0 Then gDictUserUse.RemoveAll
    Else
        Set gDictUserUse = New Dictionary
    End If
    
'    ' 计算选中的文件总数
'    For i = 1 To fgFiles.Rows - 1
'        If m_arrSelected(i) Then
'            StrSelExt = CStr(fgFiles.TextMatrix(i, 2))
'            SelFiles = SelFiles + CLng(fgFiles.TextMatrix(i, 3))
'            SelFilesExt = SelFilesExt + 1
'
'            If Not gDictUserUse.Exists(StrSelExt) Then
'                gDictUserUse.Add StrSelExt, True
'            End If
'
'        End If
'    Next i
    ' 计算选中的文件总数
    For i = 1 To fgFiles.Rows - 1
        If fgFiles.TextMatrix(i, 0) = FuhaoXuanzhe Then
            StrSelExt = CStr(fgFiles.TextMatrix(i, 2))
            SelFiles = SelFiles + CLng(fgFiles.TextMatrix(i, 3))
            SelFilesExt = SelFilesExt + 1
            
            If Not gDictUserUse.Exists(StrSelExt) Then
                gDictUserUse.Add StrSelExt, CLng(fgFiles.TextMatrix(i, 3))
            End If

        End If
    Next i

    
    ' 更新标签显示
    Label1.Caption = "共" & TotalFilesNum & "个文件，选中 " & SelFilesExt & "类 " & SelFiles & "个。"
    
End Sub


Private Sub Option1_Click()

    BuildGlobalDictIntersect

    'txtResult.text = "BMP，JPG，PNG，TIF，GIF，PCX，TGA，MP4、AVI、MOV、MKV、FLV、WMV，MPEG，3GP，MP3，WMA，WAV"
    chkSelectAll.Enabled = False
    Check1.Enabled = False
    fgFiles.Enabled = False
    
    DisplayDictionaryContent gDictUserUse
    
    UpdateSelectMorenStatus
    UpdateFileCount
    
End Sub

Private Sub Option2_Click()
    chkSelectAll.Enabled = True
    Check1.Enabled = True
    
    fgFiles.Enabled = True
    'chkSelectAll_Click
    'UpdateResultText
    DisplayDictionaryContent gDictUserUse
         UpdateFileCount
   
End Sub


Private Sub Picture1_DblClick()
MsgBox "图片音视频整理大师 - 52pojie - hellovirus"
End Sub

Private Sub Text1_Change()
    Dim result As String
    Dim i As Integer
    
    result = ""
    For i = 1 To Len(Text1.text)
        Dim ch As String
        ch = Mid(Text1.text, i, 1)
        If ch >= "0" And ch <= "9" Then
            result = result + ch
        End If
    Next i

    ' 限制最多4位
    If Len(result) > 4 Then
        result = Left(result, 4)
    End If

    ' 如果内容被修改了，则更新文本框
    If Text1.text <> result Then
        Text1.text = result
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    ' 允许数字和退格键
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = Asc(vbBack) Then
        ' 如果是数字输入，并且当前文本长度 >=4，则禁止输入
        If Len(Text1.text) >= 4 And KeyAscii <> Asc(vbBack) Then
            KeyAscii = 0
        End If
    Else
        ' 非法字符，屏蔽输入
        KeyAscii = 0
    End If
End Sub


Private Sub Text1_LostFocus()
    If Text1.text <> "" And CInt(Text1.text) < Year(Now) Then
       ZuiZaoDate = CInt(Text1.text)
    Else
       MsgBox "输入时间有误，超过现在的年份，请修改！"
       Text1.text = 2000
       Text1.SetFocus
    End If

End Sub
