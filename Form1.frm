VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "ͼƬ����Ƶ�����ʦ - 52pojie - hellovirus"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14010
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   14010
   StartUpPosition =   1  '����������
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�������ļ���"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ѡ�� Դ�ļ��У�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ʼ����"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ͬ���ļ����� �� "
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   5280
      Width           =   5295
      Begin VB.OptionButton Option10 
         Caption         =   " ��  ѡ�� �Ƚ��ļ���С (ִ���ٶȿ�)"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   " ���Ƽ��� �Ƚ��ļ�MD5ֵ(ִ���ٶ���)"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   " ʱ���ǩѡ�� �� "
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   9375
      Begin VB.OptionButton Option9 
         Caption         =   "���� exif ʱ�䡢�ļ�������ʱ�䡢�޸����ڡ��������� �����е�����ʱ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��ѡexifʱ�䣬����ѡ�ļ�������ʱ�� ���޸����ڡ��������������е�����ʱ��"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   " ������ʽ �� "
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   6760
      Begin VB.OptionButton Option6 
         Caption         =   " ���� ������Դ�ļ���"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   " �ƶ� ��ɾ��Դ�ļ���"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   " Ŀ���ļ��и�ʽ �� "
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   9375
      Begin VB.OptionButton Option5 
         Caption         =   "�����գ��� 2006/04/22"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   " ���£��� 2017/06"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   " �꣺�� 2024"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "ѡ����Ҫ������ļ�����"
      Height          =   7815
      Left            =   9600
      TabIndex        =   3
      Top             =   240
      Width           =   4335
      Begin VB.CheckBox Check1 
         Caption         =   "��ѡ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ȫѡ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�Լ�����ѡ��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "Ĭ  �ϣ� ����ͼƬ,��Ƶ,��Ƶ"
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "��ϸ�����"
         BeginProperty Font 
            Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "ѡ��Ŀ���ļ��У�"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��Ƭʱ�䲻���ڣ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ѡ�� Դ�ļ��� �� Ŀ���ļ��� ..."
      BeginProperty Font 
         Name            =   "����"
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


' ���ã�
' 1. Microsoft Scripting Runtime
' 2. Microsoft Shell Controls And Automation
' 3. ActiveX Data Objects 2.8 Library

' API����
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Const STR_SORT_ASC = "��"
Private Const STR_SORT_DESC = "��"

Private Const FuhaoXuanzhe = "��"
Private Const FuhaoWeixuanzhe = "��"


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

' ������ر���
Private m_lngSortColumn As Long    ' ��ǰ������
Private m_bolSortAscending As Boolean  ' ������
Private m_arrColumnType() As String    ' ����������

' ѡ��״̬����
Private m_arrSelected() As Boolean  ' ��¼ÿһ�е�ѡ��״̬

'Dim TotalFilesNum As Long

Public TotalFilesNum As Long      ' ���ļ���

Private SucFileLog As String  '����ɹ���־·��
Private LostFileLog As String  '����ʧ����־·��

Private FilePathOld As String
Private FileNameOld As String

Private FileNum100 As Integer


Private Sub UpdateUIExtensionTable()
    Dim i As Long
    Dim key As Variant

    ' ��ձ��
    fgFiles.Rows = 1  ' ���ٱ�����ͷ
    i = 1

    ReDim m_arrSelected(1 To gDictAllExt.Count)

    For Each key In gDictAllExt.Keys
        If i >= fgFiles.Rows Then
            fgFiles.Rows = fgFiles.Rows + 1
        End If

        fgFiles.TextMatrix(i, 0) = "��"
        fgFiles.TextMatrix(i, 1) = i
        fgFiles.TextMatrix(i, 2) = key
        fgFiles.TextMatrix(i, 3) = gDictAllExt(key)

        m_arrSelected(i) = False

        i = i + 1
    Next key

    Option1.Value = True
    Option2.Enabled = True
End Sub

' ��������������·�����Ƿ�������ļ����������ļ�
Public Sub IndexFilesInFolder(ByVal folderPath As String, Optional ByVal includeSubFolders As Boolean = True)
    Dim fso As New FileSystemObject
    Dim objFolder As folder
    
    Set objFolder = fso.GetFolder(folderPath)
    
    ' ����ֵ�
    If Not gDictAllExt Is Nothing Then
        If gDictAllExt.Count > 0 Then
            gDictAllExt.RemoveAll
        End If
    Else
        Set gDictAllExt = New Dictionary
    End If

    TotalFilesNum = 0

    ' ���� TraverseFolder�������� includeSubFolders ����
    TraverseFolder fso, objFolder, gDictAllExt, TotalFilesNum, includeSubFolders


    Label1.Caption = "ѡ���ļ��й���" & TotalFilesNum & "���ļ�! "

    If gDictAllExt.Count = 0 Then
        MsgBox "��ѡ���ļ��У�Ϊ���ļ��У��˳�..."
        Exit Sub
    End If

    ' ���½�����
    UpdateUIExtensionTable

    ' ���� UI ״̬����
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
'    folderPath = BrowseForFolder("��ѡ��Ҫ�����Դ�ļ���")
'    If folderPath = "" Then
'        MsgBox "δѡ���ļ���...", vbCritical
'        Exit Sub
'    End If
'
'    txtSourcePath.text = folderPath
'
'    ' ��ձ��
'    fgFiles.Rows = 1  ' ���ٱ�����ͷ��1������
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
'        Set gDictAllExt = New Dictionary ' ����ֵ�δ��ʼ���������³�ʼ��
'    End If
'
'    'TraverseFolder fso, objFolder, gDictAllExt, TotalFilesNum,
'     TraverseFolder fso, objFolder, gDictAllExt, TotalFilesNum, Nothing, chkIncludeSubFolders.Value
'
'    Label1.Caption = "ѡ���ļ��й���" & TotalFilesNum & "���ļ�! "
'
'    If gDictAllExt.Count = 0 Then
'       MsgBox "��ѡ���ļ��У�Ϊ���ļ��У��˳�..."
'       Exit Sub
'    End If
'    ' ��ʼ��ѡ��״̬����
'    ReDim m_arrSelected(1 To gDictAllExt.Count)
'
'    ' �����
'    Dim key As Variant
'    Dim i As Long
'    i = 1  ' �ӵ�1�п�ʼ��0���Ǳ�ͷ��
'
'    For Each key In gDictAllExt.Keys
'        ' ȷ�����㹻����
'        If i >= fgFiles.Rows Then
'            fgFiles.Rows = fgFiles.Rows + 1
'        End If
'
'        fgFiles.TextMatrix(i, 0) = "��"  ' ѡ��򣨳�ʼΪδѡ�У�
'        fgFiles.TextMatrix(i, 1) = i    ' ���
'        fgFiles.TextMatrix(i, 2) = key  ' ��չ��
'        fgFiles.TextMatrix(i, 3) = gDictAllExt(key)  ' ����
'
'        ' ��ʼ��ѡ��״̬
'        m_arrSelected(i) = False
'
'        i = i + 1
'    Next key
'
'    ' ����ȫѡ��״̬
'    'chkSelectAll.Value = 1
'
'    Option1.Value = True
'
'    Option2.Enabled = True
'
'    CheckAndEnableStartButton
'    '�����Ĭ���ļ���׺���������ļ����Ľ���
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


    folderPath = BrowseForFolder("��ѡ��Ҫ�����Դ�ļ���")
    If folderPath = "" Then
        MsgBox "δѡ���ļ���...", vbCritical
        Exit Sub
    End If

    txtSourcePath.text = folderPath
    chkIncludeSubFolders.Enabled = True

    ' �����º�����������
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
            
            ' ������ʾ
            If m_arrSelected(i) Then
'            If chkSelectAll.Value Then
                fgFiles.TextMatrix(i, 0) = FuhaoXuanzhe
            Else
                fgFiles.TextMatrix(i, 0) = FuhaoWeixuanzhe
            End If
        Next i
        ' ���½���ı�
        'UpdateResultText
        UpdateFileCount
        DisplayDictionaryContent gDictUserUse
    End If

End Sub

Private Sub fgFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then  ' ������
        If Not IsGridEmpty(fgFiles) Then
            ' ����������ѡ�����
            If fgFiles.Col = 0 And y > fgFiles.RowHeight(0) And y < fgFiles.RowHeight(0) * fgFiles.Rows Then
                ' �л�ѡ��״̬
                m_arrSelected(fgFiles.Row) = Not m_arrSelected(fgFiles.Row)

                ' ������ʾ
                If m_arrSelected(fgFiles.Row) Then
                    fgFiles.TextMatrix(fgFiles.Row, 0) = FuhaoXuanzhe
                Else
                    fgFiles.TextMatrix(fgFiles.Row, 0) = FuhaoWeixuanzhe
                End If
                
            End If
        End If
        
        Sort fgFiles, y ' ����msflexgrid ���� ����
        
        ' ͳһ���������ȫѡ��״̬
        UpdateSelectAllStatus
        UpdateFileCount
        DisplayDictionaryContent gDictUserUse
    End If
End Sub

Private Sub Form_Load()
    ' ���ô����������
   ' Me.Caption = "ͼƬ����Ƶ�����ڷ��๤��"
    
    ' ��ʼ��FlexGrid
    With fgFiles
        .Cols = 4  ' 4�У�ѡ�����š���չ��������
        .FixedCols = 0  ' û�й̶���
        .FixedRows = 1 ' 1���̶��У���ͷ��
        
        ' ���ñ�ͷ
        .TextMatrix(0, 0) = "ѡ��"
        .TextMatrix(0, 1) = "���"
        .TextMatrix(0, 2) = "��չ��"
        .TextMatrix(0, 3) = "����"
        
        ' �����п�
        .ColWidth(0) = 700  ' ѡ�����
        .ColWidth(1) = 900  ' �����
        .ColWidth(2) = 1200 ' ��չ����
        .ColWidth(3) = 1050  ' ������
        
    End With
    
    ' ��ʼ������������
    ReDim m_arrColumnType(0 To fgFiles.Cols - 1)
    m_arrColumnType(0) = "text"  ' ѡ�����
    m_arrColumnType(1) = "text"  ' �����
    m_arrColumnType(2) = "text"  ' ��չ����
    m_arrColumnType(3) = "number" ' ������
    
    ' ��ʼ���ı���
    'txtResult.text = "Ĭ�����ͣ�BMP��JPG��PNG��TIF��GIF��PCX��TGA��MP4��AVI��MOV��MKV��FLV��WMV��MPEG��3GP��MP3��WMA��WAV"
    
    InitGlobalDictAllMedia ' ��ʼ�� Ĭ���ļ����� �ֵ�
    ' DisplayDictionaryContent gDictAllMedia
    ZuiZaoDate = CInt(Text1.text)
    
End Sub

' Դ�ļ���ѡ��ť�¼�����
Private Sub cmdSelectSource_Click()
    Dim folderPath As String
    folderPath = BrowseForFolder("��ѡ��Ҫ�����Դ�ļ���")
    If folderPath <> "" Then
        txtSourcePath.text = folderPath
        CheckAndEnableStartButton
    End If
End Sub

' Ŀ���ļ���ѡ��ť�¼�����
Private Sub cmdSelectDest_Click()
    Dim folderPath As String
    folderPath = BrowseForFolder("��ѡ���ļ������Ŀ���ļ���")
    If folderPath <> "" Then
        txtDestPath.text = folderPath
        'recoveryPath = folderPath & "\ͬ���ļ�����վ"
        CheckAndEnableStartButton

 ' ������־�ļ�����������ڣ�
        CreateLogFilesIfNotExists folderPath
    End If
End Sub

' ��鲢������־�ļ�
Sub CreateLogFilesIfNotExists(folderPath As String)
   
    SucFileLog = folderPath & "\ִ����־.txt"
    'LostFileLog = folderPath & "\ʧ����־.txt"
    
    ' �������ɹ���־.txt������������ڣ�
    If Dir(SucFileLog) = "" Then
        Open SucFileLog For Output As #1
    Else
        Open SucFileLog For Append As #1
    End If
        If txtSourcePath.text <> "" Then Print #1, vbCrLf & "Դ�ļ��У�" & txtSourcePath.text & "���ļ�ִ����־ - " & Now()
        Close #1
    
    ' ������ʧ����־.txt������������ڣ�
'    If Dir(LostFileLog) = "" Then
'        Open LostFileLog For Output As #1
'    Else
'        Open LostFileLog For Append As #1
'    End If
'        If txtSourcePath.text <> "" Then Print #1, vbCrLf & "Դ�ļ��У�" & txtSourcePath.text & "���ļ�ִ��ʧ����־ - " & Now()
'        Close #1
End Sub

' ��鲢���ÿ�ʼ��ť
Private Sub CheckAndEnableStartButton()
    cmdStart.Enabled = (txtSourcePath.text <> "" And txtDestPath.text <> "")
End Sub

' ��ʼ����ť�¼�����
Private Sub cmdStart_Click()
    Dim fso As New FileSystemObject
    Dim sourceFolder As folder
    Dim logFile As TextStream
    Dim processedFiles As Long
    Dim copiedFiles As Long
    Dim startTime As Date
    
    Dim fileExtensions() As String ' �洢������ļ���׺��
    'Dim extensionList As String
    
    ' ���Դ�ļ����Ƿ����
    If Not fso.FolderExists(txtSourcePath.text) Then
        MsgBox "Դ�ļ��в�����!", vbExclamation
        Exit Sub
    End If
    
    ' ���Ŀ���ļ����Ƿ���ڣ��������򴴽�
    If Not fso.FolderExists(txtDestPath.text) Then
        On Error Resume Next
        fso.CreateFolder txtDestPath.text
        If Err.Number <> 0 Then
            MsgBox "�޷�����Ŀ���ļ���: " & Err.Description, vbCritical
            Exit Sub
        End If
        On Error GoTo 0
    End If
    
    ' �򿪻򴴽���־�ļ�
    On Error Resume Next
    Set logFile = fso.CreateTextFile(SucFileLog, True)
    If Err.Number <> 0 Then
        MsgBox "�޷����������־�ļ�: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    cmdStart.Enabled = False
    FileNum100 = 0
    'FileNum10 = 0


    ' д����־ͷ
    logFile.WriteLine "========== �ļ����࿪ʼ: " & Now() & " =========="
    logFile.WriteLine "Դ�ļ���: " & txtSourcePath.text
    logFile.WriteLine "Ŀ���ļ���: " & txtDestPath.text
    logFile.WriteLine
    
    ' ��ʼ����������ʱ��
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
    
    ' ����״̬��ʾ
    lblStatus.Caption = "���ڴ����ļ�..."
    lblStatus.Refresh
    
    ' ��ȡԴ�ļ��ж���
    Set sourceFolder = fso.GetFolder(txtSourcePath.text)
    
    If Option3.Value = True Then ProcessFolder sourceFolder, logFile, processedFiles, copiedFiles, 1, Option6.Value
    If Option4.Value = True Then ProcessFolder sourceFolder, logFile, processedFiles, copiedFiles, 2, Option6.Value
    If Option5.Value = True Then ProcessFolder sourceFolder, logFile, processedFiles, copiedFiles, 3, Option6.Value
    ' �ݹ鴦�������ļ�
    'ProcessFolder sourceFolder, logFile, processedFiles, copiedFiles, 2, True, fileExtensions
    
    ' �ر���־�ļ���д��ͳ����Ϣ
    logFile.WriteLine
    logFile.WriteLine "========== �ļ�������� =========="
    logFile.WriteLine "�����ļ�����: " & processedFiles
    logFile.WriteLine "�ɹ������ļ���: " & copiedFiles
    logFile.WriteLine "������ʱ: " & Format(Now() - startTime, "hh:mm:ss")
    logFile.Close
    
    ' ���½���״̬��ʾ
    lblStatus.Caption = "�������! ������ " & processedFiles & " ���ļ����ɹ����� " & copiedFiles & " ���ļ�����ʱ: " & Format(Now() - startTime, "hh:mm:ss")
    MsgBox lblStatus.Caption
    cmdStart.Enabled = True
    
End Sub

'
' �ݹ鴦���ļ���
' ����˵����
'   folder - Ҫ������ļ��ж���
'   logFile - ��־������
'   processedFiles - �Ѵ����ļ��������ô��ݣ�
'   copiedFiles - �Ѹ���/�����ļ��������ô��ݣ�
'   targetStructure - Ŀ��·���ṹ��1=�ꣻ2=��/�£�3=��/��/�գ�
'   isCopy - True=���ƣ�False=����
'
Public Sub ProcessFolder( _
    folder As Object, _
    logFile As TextStream, _
    ByRef processedFiles As Long, _
    ByRef copiedFiles As Long, _
    ByVal targetStructure As Integer, _
    ByVal isCopy As Boolean)
    
    Dim subFolder As Object     ' ���ļ��ж���
    Dim file As Object          ' ��ǰ�������ļ�����
    Dim fileDate As Date        ' �ļ���ʱ����Ϣ
    Dim targetPath As String    ' ������Ŀ��·��
    Dim success As Boolean      ' �����Ƿ�ɹ�
    Dim destFilePath As String  ' Ŀ���ļ�����·��
    
    Dim fileYear As Integer     ' �ļ�������ݣ���� year��
    Dim fileMonth As Integer    ' �ļ������·ݣ���� month��
    Dim fileDay As Integer      ' �ļ��������ڣ���� day��
    Dim fileExt As String       ' �ļ���չ����Сд��

    'Dim FilePathOld As String
    
    ' ������ǰ�ļ����е������ļ�
    For Each file In folder.Files
        
        FilePathOld = file.Path
        FileNameOld = file.Name
        
        ' ��ȡ�ļ���չ������ת��ΪСд�Ա�Ƚ�
        fileExt = LCase$(fso.GetExtensionName(FilePathOld))
        
        ' ���û����չ�����������ļ�
        If fileExt = "" Then
            GoTo SkipThisFile
        End If
        
        ' �����չ���Ƿ���������ֵ���
        If Not gDictUserUse.Exists(fileExt) Then
            GoTo SkipThisFile
        End If

        ' �����Ѵ����ļ�����
        processedFiles = processedFiles + 1

        ' ÿ����10���ļ�����һ��״̬���������û����飩
'        If processedFiles Mod 10 = 0 Then
        If processedFiles Mod FileNum100 = 0 Then
            lblStatus.Caption = "���ڴ����ļ�: " & processedFiles & "���Ѵ���: " & copiedFiles
            lblStatus.Refresh  ' ��ѡˢ�½���
        End If

        ' ��ȡ�ļ�ʱ����Ϣ������ʹ�� EXIF��
'        If GetFileDateParts(FilePathOld, True, fileYear, fileMonth, fileDay) Then
        If GetFileDateParts(FilePathOld, Option8.Value, fileYear, fileMonth, fileDay) Then
            fileDate = DateSerial(fileYear, fileMonth, fileDay)
        Else
            ' �����ȡʧ�ܣ�ʹ���ļ�����޸�ʱ����Ϊ����
            fileDate = FileDateTime(FilePathOld)
        End If

        ' �����ļ����ڹ���Ŀ��·��
        targetPath = BuildTargetPath(fileDate, txtDestPath.text, targetStructure)

        ' ��ȡĿ���ļ�����·���������ļ�����
        destFilePath = GetDestinationFilePath(FilePathOld, targetPath, Option10.Value)

        ' ִ�и��ƻ���в���
        If isCopy Then
            ' �����ļ�
            success = CopyFileToTarget(FilePathOld, destFilePath)
        Else
            ' �����ļ����ȸ�����ɾ��ԭ�ļ���
            success = MoveFileToTarget(FilePathOld, destFilePath)
        End If

        ' ��¼��־
        LogFileAction FilePathOld, destFilePath, success, logFile

        ' ÿ����5���ļ�����һ����־��ʾ����ѡ���������
'        If processedFiles Mod 5 = 0 Then
        If processedFiles Mod FileNum100 = 0 Then
            UpdateLogDisplay "�����ļ�: " & FileNameOld & " --> " & IIf(success, "�ɹ�", "ʧ��")
        End If

        ' ��������ɹ������ӳɹ�����
        If success Then copiedFiles = copiedFiles + 1

SkipThisFile:
    Next file

    ' �ݹ鴦�����ļ���
    On Error Resume Next
    For Each subFolder In folder.SubFolders
        ProcessFolder subFolder, logFile, processedFiles, copiedFiles, targetStructure, isCopy
    Next subFolder
End Sub

' ��ʾ�ļ���ѡ��Ի���
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

' ��ʾ�����ļ��Ի���
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


'������
Public Sub Sort(sgrd As MSFlexGrid, y As Single)
    With sgrd
    ' ����Ƿ�������������
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

' �ݹ�����ļ��У�ͳ���ļ���չ�������ļ��к��ļ�����
Private Sub TraverseFolder(fso As FileSystemObject, _
                            ByVal objFolder As folder, _
                            dictExtensions As Object, _
                            Optional ByRef totalFiles As Long = 0, _
                            Optional ByVal includeSubFolders As Boolean = True, _
                            Optional dictEmptyFolders As Object = Nothing)

    Dim isEmpty As Boolean
    
    ' ����ļ����Ƿ�Ϊ�գ����ļ��������ļ��У�
    isEmpty = (objFolder.Files.Count = 0 And objFolder.SubFolders.Count = 0)
    
    ' ����ṩ�˿��ļ����ֵ䣬���¼���ļ���
    If Not dictEmptyFolders Is Nothing And isEmpty Then
        dictEmptyFolders.Add objFolder.Path, True
    End If
    
    ' ������ǰ�ļ����е��ļ���ͳ����չ��
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
        
        ' �ۼ��ļ�����
        totalFiles = totalFiles + 1
        If totalFiles Mod 10 = 0 Then
            lblStatus.Caption = "��ͳ���ļ���: " & totalFiles
            lblStatus.Refresh  ' ��ѡˢ�½���
        End If
    Next objFile
    
    ' �������ݹ����ļ��У����������
    If includeSubFolders Then
        Dim objSubFolder As folder
        For Each objSubFolder In objFolder.SubFolders
            TraverseFolder fso, objSubFolder, dictExtensions, totalFiles, includeSubFolders
        Next objSubFolder
    End If
End Sub
' ��ȡ�ļ���·��
Private Function GetFolder(hWnd As Long, strTitle As String) As String
    Dim bi As BROWSEINFO
    With bi
        .hWndOwner = hWnd
        .lpszTitle = strTitle
        .ulFlags = &H1 ' ����ѡ���ļ���
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
' ����Ĭ�Ϻ�ȫ��������ѡ��״̬
Private Sub UpdateSelectMorenStatus()
   Dim i As Long
   Dim XuanzhongFiles As Long
   
   XuanzhongFiles = 0
   
    If Not IsGridEmpty(fgFiles) Then
        For i = 1 To fgFiles.Rows - 1
            'm_arrSelected(i) = (chkSelectAll.Value = 1)
            
            ' ������ʾ
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


' ���½���ı� - ������ʾѡ������
Private Sub UpdateResultText()
    Dim strResult As String
    Dim i As Long
    
    strResult = "��Ҫ������ļ�����:"
    
    For i = 1 To fgFiles.Rows - 1
        If m_arrSelected(i) Then
            ' ʹ�û��з�����ѡ����
            strResult = strResult & fgFiles.TextMatrix(i, 2) & " "
        End If
    Next i
    
    'txtResult.text = strResult
    lblStatus.Caption = strResult
    UpdateGlobalDictUserSel (strResult)
  '  DisplayDictionaryContent gDictUserUse
 
End Sub


' �����ļ�������ʾ
Private Sub UpdateFileCount()
    Dim SelFiles As Long
    Dim SelFilesExt As Long
    Dim i As Long
    Dim StrSelExt As String
    
    
    SelFiles = 0
    SelFilesExt = 0
        ' ��վ�����
    If Not gDictUserUse Is Nothing Then
        If gDictUserUse.Count > 0 Then gDictUserUse.RemoveAll
    Else
        Set gDictUserUse = New Dictionary
    End If
    
'    ' ����ѡ�е��ļ�����
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
    ' ����ѡ�е��ļ�����
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

    
    ' ���±�ǩ��ʾ
    Label1.Caption = "��" & TotalFilesNum & "���ļ���ѡ�� " & SelFilesExt & "�� " & SelFiles & "����"
    
End Sub


Private Sub Option1_Click()

    BuildGlobalDictIntersect

    'txtResult.text = "BMP��JPG��PNG��TIF��GIF��PCX��TGA��MP4��AVI��MOV��MKV��FLV��WMV��MPEG��3GP��MP3��WMA��WAV"
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
MsgBox "ͼƬ����Ƶ�����ʦ - 52pojie - hellovirus"
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

    ' �������4λ
    If Len(result) > 4 Then
        result = Left(result, 4)
    End If

    ' ������ݱ��޸��ˣ�������ı���
    If Text1.text <> result Then
        Text1.text = result
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    ' �������ֺ��˸��
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = Asc(vbBack) Then
        ' ������������룬���ҵ�ǰ�ı����� >=4�����ֹ����
        If Len(Text1.text) >= 4 And KeyAscii <> Asc(vbBack) Then
            KeyAscii = 0
        End If
    Else
        ' �Ƿ��ַ�����������
        KeyAscii = 0
    End If
End Sub


Private Sub Text1_LostFocus()
    If Text1.text <> "" And CInt(Text1.text) < Year(Now) Then
       ZuiZaoDate = CInt(Text1.text)
    Else
       MsgBox "����ʱ�����󣬳������ڵ���ݣ����޸ģ�"
       Text1.text = 2000
       Text1.SetFocus
    End If

End Sub
