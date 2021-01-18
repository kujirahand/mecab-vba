Attribute VB_Name = "mecab"
Option Explicit

'
' MeCab for Excel VBA (Windows)
'

Dim MeCabPath As String
Dim MeCabCharset As String
Dim MeCabOptions As String
Dim MeCabDictDir As String
Const MeCabCharsetDefault = "Shift_JIS"

' MeCab�̌��ʃf�[�^��ێ����郆�[�U�[�^
Public Type MeCabItem
    �\�w�` As String
    �i�� As String
    �i���ڍ�1 As String
    �i���ڍ�2 As String
    �i���ڍ�3 As String
    ���p�` As String
    ���p�^ As String
    ���` As String
    ���~ As String
    ���� As String
End Type

Public Sub SetMeCabPath(ByVal Path)
    MeCabPath = Path
End Sub

Public Sub SetMeCabCharset(ByVal Charset As String)
    MeCabCharset = Charset
End Sub

Public Sub SetMeCabOptions(ByVal Options As String)
    MeCabOptions = Options
End Sub

Public Sub SetMeCabDictDir(ByVal DictDir As String)
    MeCabDictDir = DictDir
End Sub

Private Sub MeCabInit()
    ' Find MeCab
    If MeCabPath = "" Then
        MeCabPath = "C:\Program Files (x86)\MeCab\bin\mecab.exe"
        If FileExists(MeCabPath) = False Then
            MeCabPath = "C:\Program Files\MeCab\bin\mecab.exe"
            If FileExists(MeCabPath) = False Then
                MeCabPath = ThisWorkbook.Path & "\mecab.exe"
                If FileExists(MeCabPath) = False Then
                    MeCabPath = ThisWorkbook.Path & "\bin\mecab.exe"
                    If FileExists(MeCabPath) = False Then
                        MsgBox "MeCab���C���X�g�[������Ă��܂���B" & vbCrLf & _
                            "���邢�́AMeCab�̃p�X���w�肵�Ă��������B"
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub MeCabExecToSheet(ByVal InText As String, ByRef Sheet As Worksheet, ByVal Top As Integer)
    Dim Res As String
    Res = MeCabExec(InText)
    If Res = "" Then Exit Sub
    Dim Lines, y, x, row As Integer
    ' On Error GoTo ERR_OVERFLOW
    row = Top
    Lines = Split(Res, vbCrLf)
    For y = 0 To UBound(Lines)
        Dim Line As String
        Line = Lines(y)
        If Line = "" Then GoTo Y_CONTINUE
        
        Dim Tabs
        Tabs = Split(Line, Chr(9))
        If UBound(Tabs) = 0 Then
            DoEvents
            GoTo Y_CONTINUE
        End If
        Dim Word, Desc
        Word = Tabs(0)
        Desc = Tabs(1)
        
        Dim Cm
        Cm = Split(Desc, ",")
        If UBound(Cm) < 8 Then GoTo Y_CONTINUE
                
        Sheet.Cells(row, 1) = Word
        For x = 0 To UBound(Cm)
            Sheet.Cells(row, 2 + x) = Cm(x)
        Next
        row = row + 1
Y_CONTINUE:
    Next
    Exit Sub
ERR_OVERFLOW:
    Debug.Print "[MeCabExecToSheet] " & Err.Description & " : " & row & "�s�ڂŃG���["
End Sub


Public Function MeCabExecToItems(ByVal InText As String) As MeCabItem()
    Dim Res As String
    Res = Trim(MeCabExec(InText))
    Dim i, Lines, Line, Cm, Word, Desc, da, wa
    Lines = Split(Res, vbCrLf)
    Dim items() As MeCabItem
    ReDim items(UBound(Lines) + 1)
    For i = 0 To UBound(Lines)
        ' �s�𓾂�
        Line = Lines(i)
        If Line = "" Then GoTo I_CONTINUE
        ' �^�u�ŋ�؂�
        wa = Split(Line, Chr(9))
        If UBound(wa) = 0 Then
            DoEvents
            GoTo I_CONTINUE
        End If
        Word = wa(0)
        Desc = wa(1)
        ' �J���}�ŋ�؂�
        da = Split(Desc, ",")
        If UBound(da) < 8 Then GoTo I_CONTINUE
        ' comma : �i��,�i���ו���1,�i���ו���2,�i���ו���3,���p�^,���p�`,���`,�ǂ�,����
        items(i).�\�w�` = Word
        items(i).�i�� = da(0)
        items(i).�i���ڍ�1 = da(1)
        items(i).�i���ڍ�2 = da(2)
        items(i).�i���ڍ�3 = da(3)
        items(i).���p�^ = da(4)
        items(i).���p�` = da(5)
        items(i).���` = da(6)
        items(i).���~ = da(7)
        items(i).���� = da(8)
I_CONTINUE:
    Next
    MeCabExecToItems = items
End Function


Public Function MeCabExec(ByVal InText As String) As String
    Dim InFile As String, ResultFile As String, Cmd As String, Res As String
    Dim BatFile As String, Opt As String
    
    ' MeCab�̏�����
    Call MeCabInit
    
    BatFile = GetTempPath(".bat")
    InFile = GetTempPath(".txt")
    ResultFile = GetTempPath(".txt")
    
    ' ���̓e�L�X�g���t�@�C���ɕۑ�
    MeCabSaveText InFile, InText ' ���̓e�L�X�g�̓C���X�g�[�������̃R�[�h
    
    ' �I�v�V�����𔽉f
    Opt = "" & MeCabOptions
    If MeCabDictDir <> "" Then
        Opt = Opt & " -d " & MeCabDictDir
    End If
    
    ' �o�b�`���쐬
    Cmd = "type " & qq(InFile) & " | " & qq(MeCabPath) & " " & MeCabOptions & " > " & qq(ResultFile) & vbCrLf
    ' Cmd = Cmd & "pause" & vbCrLf
    SaveToFile BatFile, Cmd, "Shift_JIS" ' �o�b�`�t�@�C����Shift_JIS�K�{
    Debug.Print Cmd
    
    ' �o�b�`�����s
    If ShellWait(BatFile) Then
        Res = MeCabLoadText(ResultFile)
        ' Debug.Print Res
        MeCabExec = Res
    Else
        MeCabExec = ""
    End If
End Function


Private Function GetTempPath(Ext As String) As String
    Dim FSO As Object, tmp As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    tmp = FSO.GetSpecialFolder(2) & "\" & FSO.GetBaseName(FSO.GetTempName) & Ext
    GetTempPath = tmp
End Function


' Clear sheet
Public Sub ClearSheet(ByRef Sheet As Worksheet, ByVal TopRow As Integer)
    Dim EndCol, EndRow, row, Col
    With Sheet.UsedRange
        EndRow = .Rows(.Rows.Count).row
        EndCol = .Columns(.Columns.Count).Column
    End With
    For row = TopRow To EndRow
        For Col = 1 To EndCol
            Sheet.Cells(row, Col) = ""
        Next
    Next
End Sub

' TSV to Sheet
Public Sub TSVToSheet(ByRef Sheet As Worksheet, ByVal tsv As String, TopRow As Integer)
    Dim Rows As Variant, Cols As Variant
    Dim i, j
    Rows = Split(tsv, Chr(10))
    For i = 0 To UBound(Rows)
        Cols = Split(Rows(i), Chr(9))
        For j = 0 To UBound(Cols)
            Dim v
            v = Cols(j)
            v = Replace(v, "��", vbCrLf)
            Sheet.Cells(i + TopRow, j + 1) = v
        Next
    Next
End Sub


' Convert Sheet to TSV
Public Function ToTSV(ByRef Sheet As Worksheet) As String
    Dim s As String
    s = ""
    ' �V�[�g�͈̔͂��擾
    Dim BottomRow As Integer, RightCol As Integer
    BottomRow = Sheet.Range("A1").End(xlDown).row
    RightCol = Sheet.Range("A1").End(xlToRight).Column
    ' �V�[�g�͈͂����ォ�珇�Ɏ擾
    Dim y, x, v
    For y = 1 To BottomRow
        For x = 1 To RightCol
            v = Sheet.Cells(y, x)
            ' �Z�����̉��s�����͒u�����Ă���
            v = Replace(v, vbCrLf, "��")
            s = s & v & Chr(9)
        Next
        s = s & vbCrLf
    Next
    ToTSV = s
End Function

' �N�H�[�g����
Public Function qq(str) As String
    qq = """" & str & """"
End Function

' �R�}���h�����s���ďI���܂őҋ@����
Public Function ShellWait(ByVal Command As String) As Boolean
    On Error GoTo SHELL_ERROR
    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    Dim Res As Integer
    Res = wsh.Run(Command, 7, True) ' �ŏ������ďI���܂őҋ@
    ShellWait = (Res = 0)
    Exit Function
SHELL_ERROR:
    ShellWait = False
End Function

Public Sub MeCabSaveText(ByVal Filename, ByVal Text)
    If MeCabCharset = "" Then MeCabCharset = MeCabCharsetDefault
    SaveToFile Filename, Text, MeCabCharset
End Sub

Public Function MeCabLoadText(ByVal Filename) As String
    If MeCabCharset = "" Then MeCabCharset = MeCabCharsetDefault
    MeCabLoadText = LoadFromFile(Filename, MeCabCharset)
End Function

' �C�ӂ̕����G���R�[�f�B���O���w�肵�ăe�L�X�g�t�@�C����ǂ�
Private Function LoadFromFile(Filename, Charset) As String
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' text
    stream.Charset = Charset
    stream.Open
    stream.LoadFromFile Filename
    LoadFromFile = stream.ReadText
    stream.Close
End Function

' �e�L�X�g���w�蕶���R�[�h�Ńt�@�C���ɕۑ�
Private Sub SaveToFile(ByVal Filename, ByVal Text, ByVal Charset)
    ' UTF-8 �̏ꍇ BOM�͕s�v
    If LCase(Charset) = "utf-8" Or LCase(Charset) = "utf-8n" Or LCase(Charset) = "utf8" Then
        Call SaveToFileUTF8N(Filename, Text)
        Exit Sub
    End If
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = Charset
    stream.Open
    stream.WriteText Text
    stream.SaveToFile Filename, 2
    stream.Close
End Sub

' BOM�Ȃ���UTF-8�Ńt�@�C���Ƀe�L�X�g����������
Private Sub SaveToFileUTF8N(Filename, Text)
    Dim stream, buf
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2 ' �e�L�X�g���[�h���w�� --- (*1)
        .Charset = "UTF-8"
        .Open
        .WriteText Text ' �e�L�X�g����������
        .Position = 0 ' �J�[�\�����t�@�C���擪�� --- (*2)
        .Type = 1 ' �o�C�i�����[�h�ɕύX
        .Position = 3 ' BOM(3�o�C�g)���΂�
        buf = .Read() ' ���e��ǂݍ���
        .Position = 0 ' �J�[�\����擪�� --- (*3)
        .Write buf ' BOM�Ȃ��̃e�L�X�g����������
        .SetEOS
        .SaveToFile Filename, 2
        .Close
    End With
End Sub

Private Function FileExists(ByVal Filename) As Boolean
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FileExists = FSO.FileExists(Filename)
    Set FSO = Nothing
End Function

