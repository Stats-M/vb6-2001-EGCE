VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowser 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10815
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "������"
      Height          =   350
      Left            =   9345
      TabIndex        =   23
      Top             =   1995
      Width           =   1170
   End
   Begin SHDocVwCtl.WebBrowser Inet1 
      Height          =   540
      Left            =   4935
      TabIndex        =   15
      Top             =   1995
      Width           =   750
      ExtentX         =   1323
      ExtentY         =   952
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   5985
      TabIndex        =   13
      Text            =   "3"
      Top             =   2310
      Width           =   750
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   6930
      Top             =   1995
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "frmBrowser.frx":0442
      Top             =   1785
      Width           =   4740
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4110
      Left            =   105
      TabIndex        =   7
      Top             =   2730
      Width           =   10620
      ExtentX         =   18732
      ExtentY         =   7250
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "����"
      Height          =   350
      Left            =   9345
      TabIndex        =   10
      Top             =   630
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   350
      Left            =   9345
      TabIndex        =   9
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   225
      Left            =   2205
      TabIndex        =   22
      Top             =   1470
      Width           =   645
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   225
      Left            =   2205
      TabIndex        =   21
      Top             =   1155
      Width           =   645
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   225
      Left            =   2205
      TabIndex        =   20
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label15 
      Caption         =   "������ ������� "
      Height          =   225
      Left            =   210
      TabIndex        =   19
      Top             =   1470
      Width           =   2010
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   435
      Left            =   4935
      TabIndex        =   18
      Top             =   525
      Width           =   4320
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label13 
      Caption         =   "������ �������"
      Height          =   225
      Left            =   210
      TabIndex        =   17
      Top             =   1155
      Width           =   2010
   End
   Begin VB.Label Label12 
      Caption         =   "����� ����������:"
      Height          =   225
      Left            =   3255
      TabIndex        =   16
      Top             =   525
      Width           =   1590
   End
   Begin VB.Label Label11 
      Caption         =   "PAGE NUM"
      Height          =   225
      Left            =   5775
      TabIndex        =   14
      Top             =   1995
      Width           =   1065
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   435
      Left            =   7455
      TabIndex        =   12
      Top             =   2100
      Width           =   1275
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   $"frmBrowser.frx":0448
      Height          =   645
      Left            =   3150
      TabIndex        =   8
      Top             =   1155
      Width           =   7470
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   225
      Left            =   4935
      TabIndex        =   6
      Top             =   210
      Width           =   4215
   End
   Begin VB.Line Line1 
      X1              =   3045
      X2              =   3045
      Y1              =   105
      Y2              =   1785
   End
   Begin VB.Label Label7 
      Caption         =   "������� N1 - ������:"
      Height          =   225
      Left            =   3255
      TabIndex        =   5
      Top             =   210
      Width           =   1590
   End
   Begin VB.Label Label5 
      Caption         =   "�� ��� ����������"
      Height          =   225
      Left            =   210
      TabIndex        =   4
      Top             =   840
      Width           =   1905
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   225
      Left            =   2205
      TabIndex        =   3
      Top             =   525
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "������� ������"
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   525
      Width           =   1905
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   225
      Left            =   2205
      TabIndex        =   1
      Top             =   210
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "������������ ��������"
      Height          =   225
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1905
   End
   Begin VB.Menu mnuFile 
      Caption         =   "����"
      Begin VB.Menu mnuWriteAuthors 
         Caption         =   "�������� �������"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�����"
      End
   End
   Begin VB.Menu mnuDownload 
      Caption         =   "�������"
      Begin VB.Menu mnuConnect 
         Caption         =   "������������"
      End
      Begin VB.Menu mnuHyp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "������"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "����������"
      End
      Begin VB.Menu mnuView 
         Caption         =   "�������� HTML"
      End
      Begin VB.Menu mnuHyp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "���������..."
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "���������"
      Begin VB.Menu mnuProse 
         Caption         =   "�����"
         Begin VB.Menu mnuF002 
            Caption         =   "������"
         End
      End
      Begin VB.Menu mnuOfftop 
         Caption         =   "���������"
         Begin VB.Menu mnuF001 
            Caption         =   "������� ���������"
         End
      End
      Begin VB.Menu mnuF003 
         Caption         =   "�������� �����"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "�������"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "����� �������"
      End
      Begin VB.Menu mnuHyp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "� ���������"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Const MaxRetryAttempts = 3  '������������ ����� ��������� ������� ������� ��������

'�������, � ���� ���������� �� ������ P-II? :) ��� ��� ��, � �� �������!
'���������� ��� P-II � ����� ����� ������������ ������ 32-� ������� �����������.
Private StartPos As Long        '��������� ������� ��� ��������� ��������
Private EndPos As Long          '�������� ������� ��� ��������� ��������
Private TopicNum As Long        '����� �����
Private email As String         'e-mail ������ (��������� � ��������� ������)
Private Nickname As String      '��� ������
Private msgHREF As String       'URL ���������� ���������
Private msgHead As String       '��������� ���������
Private Indent As Long          '������ ���������
Private tokenname As String     '������������� ��������
Private lResult As Long         '�������������� ���������� Long
''LEGACY Private FirstTime(2) As Boolean '��� ����������� ���������������� Inet
Private strMess As String       '��� ������� ������� � ��������� ����������
Private vtData As Variant       '������������� Microsoft, ��� ���������� �� �������
Private bDone As Boolean        '����� ������ �������
Private bCancel As Boolean      'TRUE, ���� ���� ������ ������ Cancel
Private GettingURLs As Boolean  '��������� ��� �������, ��� ������������ ������� ��������
Private Loading As Boolean      '��� �������� ����� ������ ��������� �������
Private lRetry As Long           '��������� �������
Private HTMLDoc As HTMLDocument

'���������� ��� ������ Feeder - ����� ���� � �.�.
Private COLOR As String         '���� ���� ��� HTML-��������
Private bColor As Boolean       '�������� ���� �� ������ ����
Private TopicChanged As Boolean '����������� ��������� ������ ���� ����������
Private LastTopicNum As Long    '����� ���� ���������� ��������� (�������� ����� ������� ��� ��� � HTML)

Private Sub Test()
Dim HistoryLen As Long       '���������� ������� � �������
Dim hHFile As Long  '��������� �� ���� �������
Dim hFile As Long   '��������� �� �� �������
Dim hAuthor As tAuthor
Dim hHistory As tHistory
Dim i As Long

    hHFile = FreeFile
    Open (App.Path & DataPath & HistoryFile) For Random As hHFile Len = Len(hHistory)
    hFile = FreeFile
    Open (App.Path & DataPath & AuthorsFile) For Random As hFile Len = Len(hAuthor)
    '���������� ���������� ������� � �������
    HistoryLen = LOF(hHFile) \ Len(hHistory)
    Label10.Caption = Str(HistoryLen)
''    Text1.Text = strEmpty
''    If HistoryLen > 0 Then   '������ ����������
''        For i = 1 To HistoryLen
''            '������ ��������� ������ � �������
''            Get hHFile, i, hHistory
''            Text1.Text = Text1.Text & hHistory.AuthorID & "-" & hHistory.MsgSize & " " & hHistory.URLvalue & " "
''            If hHistory.AuthorID > 0 Then
''                Get hFile, hHistory.AuthorID, hAuthor
''                Text1.Text = Text1.Text & Left(hAuthor.Nickname, hAuthor.NicknameLen) & vbCrLf
''            End If
''        Next i
''    End If
''    Text1.Text = Text1.Text + vbCrLf
''    For i = 1 To Authors.LastRecNum
''        Get hFile, i, hAuthor
''        Text1.Text = Text1.Text + Left(hAuthor.Nickname, hAuthor.NicknameLen) + vbCrLf
''    Next i
    Close hHFile
    Close hFile
    ''bResult = Authors.UpdateDB(ucHistoryMode)
End Sub


'*****************************************************
'*      ����� � �������� ����������� ��������        *
'* �� ��������� ����� � �������� �� ������           *
'*****************************************************
Private Function Splitter(ByRef lpString As String) As Boolean
Dim tmpstr As String
    Splitter = False
    '�������� ������, ���� �� �����!
    Timer1.Enabled = False
    Label14.Caption = "���� ������ ��������-����������"
    Label14.Refresh
    '��������, ��� ��� ��������� ��, ��� ���� � � ������ ���� <dl>
    StartPos = InStr(1, lpString, "<DL>", vbTextCompare)
    If StartPos = 0 Then
        '��� ��������!!!
        'TO DO smth
        Label4.Caption = "������ ������� ����������"
        Label4.Refresh
        Label14.Caption = "������ ������� ����������"
        Label14.Refresh
        If InStr(1, lpString, "The requested URL could not be retrieved", vbTextCompare) > 0 Then
            Label14.Caption = "������������� �������� �� ����� ���� ��������"
            Label14.Refresh
        End If
        Exit Function
    End If
    '��������������� �������� ������! ����� ����� ������� ��������� �� ������� ����� ����� ���������� ��� ���������.
    '������ ���� ������ (���� ����� "��������")
    lResult = InStr(1, lpString, "��������", vbTextCompare)
    '����� ���� ������ (���� ����� "������")
    EndPos = InStr(lResult, lpString, "������", vbTextCompare)
    '�������� ������
    BaseHREF = Mid(lpString, lResult, EndPos - lResult)
    '����� � ���� :) ������ ������:
    ''lResult = InStr(1, BaseHREF, "href=`", vbTextCompare)book.cgi?book=
    lResult = InStr(1, BaseHREF, "book.cgi?book=", vbTextCompare)
    '���� ���� ��������� � ������ (book=Elitegames...) �
    '���������� ��������� ������� (���� ��������� ������ �������� � ������)
    EndPos = InStr(lResult, BaseHREF, "=", vbTextCompare) + 1
    '�������� � �������� ��� ������� �����
'TO DO �������� ������ ������: ������ �� ������ book.cgi �������� ��� �������� �������!!!
    BaseHREF = Replace(Mid(BaseHREF, lResult, EndPos - lResult), "`", strEmpty, 1, -1, vbBinaryCompare)
    
    '���������� ��������� ��������: �������� ��� ������
    lpString = Right(lpString, Len(lpString) - StartPos)
    Label2.Caption = Str(CurrPage)
    '������� �� ������ ��� ���� � NavigateComplete2
    
    '���� ����� �����
    StartPos = InStr(1, lpString, "<p>", vbTextCompare) + Len("<p>") '����� ������������ ���
    EndPos = InStr(StartPos, lpString, ".", vbTextCompare)
    '�� ����� ���������: ��� ������ ��� ����� ������ �� �������, � ��������
    '� ��� ������ ������� ��������, ��� � �����
    Do While EndPos - StartPos < 8
        '���� �� ���������� ����� (���������, ������ 5-�������� ������ ��� �������?)
        '8 = <p> + XXXXX
        If EndPos = 0 Then
            '<p> ���, �� ��� �� ����� �����... ����� ����� ����?
            'TO DO smth
            Exit Do
        End If
        TopicNum = CLng(Val(Mid(lpString, StartPos, EndPos - StartPos)))
        '������� ����� �����
        EndPos = InStr(EndPos, lpString, "<p>", vbTextCompare)
        If EndPos = 0 Then
            '��������� <p> �� ������. ������������, ���� ��� ������ ���� � ����� �������� ���� �����������
            'TO DO smth
            Exit Do
        End If
        '�������� ����� � ���������� ����� <p>...
        tmpstr = Mid(lpString, StartPos - 3, EndPos)
        '... � ������� �� �� �������� ������: ����� ������� ����� ������ ����-���� �� ��������
        '�������������� ������ ����� ��� �������� ������������ ������ � ����������
        lpString = Right(lpString, Len(lpString) - EndPos + 1)
        '�������� ����� ������ �� �����������
        If ParserLevel1(TopicNum, tmpstr) Then
            'TO DO
        End If
        '���� ����� ��������� �����
        StartPos = InStr(1, lpString, "<p>", vbTextCompare) + Len("<p>") '����� ������������ ���
        EndPos = InStr(StartPos, lpString, ".", vbTextCompare)
    Loop
    '�������� ������ ��� � InitFeeder!
    Label14.Caption = "��������-���������� ����������������"
    Label14.Refresh
    Splitter = True
End Function

'******************************************************
'*      ����������� ����� �� ��������� ������         *
'* �� ������ ����� ������ �� ������ ��������� �����   *
'* � ���������� �� � MsgMap (��� ���� ����������)     *
'*      TopicNum - ����� ���� (��� ������ � �� ������)*
'******************************************************
Private Function ParserLevel1(ByVal TopicNum As Long, ByVal lpString As String) As Boolean
Dim StartPos As Long
Dim EndPos As Long
    '''EndPos = InStr(1, lpString, "A href=`http://book.by.ru/cgi-bin/book.cgi?book=", vbTextCompare)
    ''EndPos = InStr(1, lpString, "a href=`book.cgi?book=", vbTextCompare)
    'EndPos = InStr(1, lpString, "book.cgi?book=", vbTextCompare)
    EndPos = InStr(1, lpString, BaseHREF, vbTextCompare)
    Debug.Print "���� � ParserLevel1"
''    Indent = 0 '������ ��������� �����, ������ = 0
    '���� � ���, ��� ��� Indent = 0 ������ ��� <DL> ������� ������ ������� �
    '������� ������ �� ������, ��� ������������ ��������� ������� ��� ������
    '������� ������. ��� ���������� ����� ���������� ������ ������ � 1.
    Indent = 1 '������ ��������� �����, ������ = 1
    Do While EndPos <> 0
        '���������� e-mail ������
        email = strEmpty
        '���� ��������� ������ �� e-mail
        StartPos = InStr(1, lpString, "<a href=`mailto:", vbTextCompare)
        '�������� - ����� ����������, �.�. �� ������ ��� ����� e-mail � ����� ����������
        '''lResult = EndPos + Len("A href=`http://book.by.ru/cgi-bin/book.cgi?book=")
        ''lResult = EndPos + Len("a href=`book.cgi?book=")
        'lResult = EndPos + Len("book.cgi?book=")
        lResult = EndPos + Len(BaseHREF)
        If EndPos > StartPos Then
            '����������� ��������� ������ ������� (������� e-mail, ����� ����������)
            If Not (StartPos = 0) Then
                '���� e-mail ������
                StartPos = StartPos + Len("<A href=`mailto:")
                EndPos = InStr(StartPos, lpString, "`>", vbTextCompare)
                email = Mid(lpString, StartPos, EndPos - StartPos)
            End If  '� ��������� ������ e-mail ������ �� ������
        End If  '� ��������� ������ e-mail ������, �� �� ��������� ��� � ����������
                '��������� -> ���������� ���
        '���������� ��� ������
        Nickname = strEmpty
        StartPos = InStr(1, lpString, "<B><I>", vbTextCompare) + Len("<B><I>")
        EndPos = InStr(StartPos, lpString, "</I></B>", vbTextCompare)
        If Not (StartPos >= EndPos) Then
            '��� ������ ������� (���� ��, ������� �����/�� ������� ������� ��� 8-) )
            Nickname = Mid(lpString, StartPos, EndPos - StartPos)
            Nickname = RemoveSpaces(Nickname)
        End If
        '���������� URL ���������
        msgHREF = strEmpty
        EndPos = InStr(lResult, lpString, "`>", vbTextCompare)
        '����� �����, ������ ��� �� &amp; , �������� ������ &
        msgHREF = Replace(Mid(lpString, lResult, EndPos - lResult), "amp;", strEmpty, 1, -1, vbBinaryCompare)
        '������������� �����! �������� ���� ��� ������������ ���������.
        '������ ��� ����� �������� ����� ���������� ������� Book.ru
        StartPos = EndPos + 2   '��������� "`>"
        EndPos = InStr(StartPos, lpString, "</A>", vbTextCompare)
        msgHead = Mid(lpString, StartPos, EndPos - StartPos)
'��������! ������ ���� �������� (���� ����) ���������� "(-)",
'������� ������ ��������� ������� � ����������, ����������� �� ������ ����.
'����� ������� "(+)"
'���������� ���������� lResult � StartPos ��� �������� �� ������������
        msgHead = RemoveSpaces(msgHead)
        StartPos = InStr(1, msgHead, "(-)", vbTextCompare)
        Do While StartPos > 0
        ''If StartPos > 0 Then
            '(-) ������. �������� (�����), ��������� �� �� � ������...
            If StartPos > Len(msgHead) - 5 Then
                '(-) � ����� ������ - ������ ������� �����
                msgHead = Left(msgHead, Len(msgHead) - 3)
            Else
                '(-) � ��������. ������� ������ ��� 3 �������
                msgHead = Left(msgHead, StartPos - 1) + Right(msgHead, Len(msgHead) - StartPos - 2)
            End If
        ''End If
            '������� ������� � ������ � ����� ������
            msgHead = RemoveSpaces(msgHead)
            StartPos = InStr(1, msgHead, "(-)", vbTextCompare)
        Loop
        StartPos = InStr(1, msgHead, "(+)", vbTextCompare)
        Do While StartPos > 0
        ''If StartPos > 0 Then
            '(+) ������. �������� (�����), ��������� �� �� � ������...
            If StartPos > Len(msgHead) - 5 Then
                '(+) � ����� ������ - ������ ������� �����
                msgHead = Left(msgHead, Len(msgHead) - 3)
            Else
                '(+) � ��������. ������� ������ ��� 3 �������
                msgHead = Left(msgHead, StartPos - 1) + Right(msgHead, Len(msgHead) - StartPos - 2)
            End If
        ''End If
            '������� ������� � ������ � ����� ������
            msgHead = RemoveSpaces(msgHead)
            StartPos = InStr(1, msgHead, "(+)", vbTextCompare)
        Loop
'������ � ��� ���������
LogEvent ("[Parser Level 1] ��������� ����: " + msgHead)
'{Fixed} TO DO ��� �� ������? msgHead = RemoveSpaces(msgHead)
'������� ������ ���������� ���������, ��� ������ ���� ��
        '���� � ����� ��������� ����� �������� ��� � PareserLevel2
        '�������� ������������ �����
        lpString = Right(lpString, Len(lpString) - EndPos)
        '��������� ������
'������ � ��� ���������
LogEvent ("[Parser Level 1] ������ ����� ���������� ��� ������ " + Nickname)
        With MsgMap
            .TopicNum = TopicNum
            .Author = Nickname
            .AuthorLen = Len(Nickname)
            .email = email
            .emailLen = Len(email)
            .msgURL = msgHREF
            .msgURLLen = Len(msgHREF)
            .msgHead = msgHead
            .msgHeadLen = Len(msgHead)
            .MsgSize = 0
            .msgDate = strEmpty
            .msgDateLen = 0
            .Indent = Indent
        End With
        Debug.Print Str(MsgMap.TopicNum) & " - " & MsgMap.Author & " - " & MsgMap.msgURL
        '��������� � �� � �������� ����� ������
        Label4.Caption = Str(MsgMap.Save)
        Label4.Refresh
'������ � ��� ���������
LogEvent ("[Parser Level 1] ���� ��������� � ����� ������ ��� ������� " + Label4.Caption)
        '���� ��������� ������ �� ��������� (������������� �����)+����� <dl> � </dl>
        '''EndPos = InStr(1, lpString, "A href=`http://book.by.ru/cgi-bin/book.cgi?book=", vbTextCompare)
        ''EndPos = InStr(1, lpString, "a href=`book.cgi?book=", vbTextCompare)
        'EndPos = InStr(1, lpString, "book.cgi?book=", vbTextCompare)
        EndPos = InStr(1, lpString, BaseHREF, vbTextCompare)
        '��������� ������ ���������� ���������. ��������! EndPos ����� ���������!!!
        '����� ����� �� ������ � �� EndPos, ���� �� �� ����.
        '� ����� ����� <dl> ��� </dl>
        If Not (EndPos = 0) Then
            '����� ������������ ���������� email � �������� ���������
            '������ ����� ������ ��� �� �����, � �� �������� ��������� � ������ �����
            email = Left(lpString, EndPos)
            '���������� ���� ����� ������ ������������
            '������������ <dl> � </dl> ����������� �� �����, ������� ���� "� ���"
            lResult = InStr(1, email, "</dl>", vbTextCompare)
            Do While lResult > 0
                '���� ���� </dl>, �������� �� �� ������ � ��������� �������
                email = Right(email, Len(email) - lResult - Len("</dl>"))
                Indent = Indent - 1
                lResult = InStr(1, email, "</dl>", vbTextCompare)
            Loop
            lResult = InStr(1, email, "<dl>", vbTextCompare)
            Do While lResult > 0
                '���� ���� <dl>, �������� �� �� ������ � ����������� �������
                email = Right(email, Len(email) - lResult - Len("<dl>"))
                Indent = Indent + 1
                lResult = InStr(1, email, "<dl>", vbTextCompare)
            Loop
            '�������� ������� �� ���� ����� ��������� �� ���������
''            If Indent < 0 Then
''                Indent = 0
''            End If
            '��������� ���������� �������. ������ 0 ������ ����� 1.
            If Indent < 1 Then
                Indent = 1
            End If
        End If
    Loop
End Function

'******************************************************
'*      �������������� ���������� ��� Feeder �        *
'*              �������� HTML-��������                *
'* ����� ��������� ������� ��������� �������          *
'******************************************************
Private Function InitFeeder() As Boolean
    Label14.Caption = "���������� � ������� ������"
    Label14.Refresh
    lRetry = 0 '������� ������� ������� �����������
    '������ ������ - ��������� ������������� ������ ������ ���� ����������
    TopicChanged = True
    bDone = False
    strMess = strEmpty
    '������ �������� ����� �� ����� ����
    bColor = True
    COLOR = ColorLight
    '������� HTML-����, �������������� ������ (���� ����) ������������
    EraseHTML (CurrPage)
    bResult = Module1.WriteHTML(0, CurrPage, COLOR, strEmpty, RecNum)  '�����
    '��������� ������ ����� ��� ������������� �������� ���� � Feeder
    bResult = Module1.WriteHTML(5, CurrPage, COLOR, strEmpty, RecNum)  '����� �������
    bResult = Module1.WriteNavigator(0, CurrPage, RecNum, Indent, strEmpty) '����� ���������� � ���������� �������
    If MsgMap.LastRecNum > 0 Then
        '�������� ����� ������ � �� ������. ����� (� ParserLevel2) ��� ������
        '����� ����� ��������, ����� ��� �������� � ����������
        RecNum = MsgMap.GetNext
        If Not (RecNum = 0) Then
            '��, ������ ���������, ��������� �������
            '���������� ��� ��������� ��������� �� MsgMap ���� �� �� ������������...
            With hMsg
                .Author = MsgMap.Author
                .AuthorLen = MsgMap.AuthorLen
                .email = MsgMap.email
                .emailLen = MsgMap.emailLen
                '.msgDate = strEmpty
                .msgDate = MsgMap.msgDate
                '.msgDateLen = 0
                .msgDateLen = MsgMap.msgDateLen
                .msgHead = MsgMap.msgHead
                .msgHeadLen = MsgMap.msgHeadLen
                '.MsgSize = 0
                .MsgSize = MsgMap.MsgSize
                .msgURL = MsgMap.msgURL
                .msgURLLen = MsgMap.msgURLLen
                .TopicNum = MsgMap.TopicNum
                .Indent = MsgMap.Indent
            End With
            '�������� ����� ���� ������ ������� ���������
            LastTopicNum = hMsg.TopicNum
            '��������� ������� ������� ��� �������� �� ���������� (��� ����� �����)
            InitFeeder = True
            Call StartInet(Left(hMsg.msgURL, hMsg.msgURLLen))
        End If
    Else
        InitFeeder = False
    End If
    Label14.Caption = "���� ������� ��������� �� �������"
    Label14.Refresh
    '�������� ������
    Timer1.Enabled = True
End Function

'*****************************************************
'* ����������� ������ ������ �������                 *
'* ��������� ������ �������� ����������              *
'* ��������� ��������� ��������                      *
'*      Feeder = 0 �������� ����������               *
'*      Feeder <>0 ����� ������, ��������� ������    *
'*****************************************************
Private Function Feeder() As Long
Dim tmpString As String
    '��������� ������ �� ����� ������
    Timer1.Enabled = False
    '����� ������ �� ������� � ����������� �� ParserLev2, �������� ����� ���������
    tmpString = ParserLevel2(strMess)
    If lRetry > 0 Then
        If lRetry > MaxRetryAttempts Then
            '�������� �������� �� ������� ����� MaxRetryAttempts �������
            Label16.Caption = Str(Val(Label16.Caption) + 1)
            Label14.Caption = "������ �������. ������� ������� ������"
            Label14.Refresh
            tmpString = "<i>���������� <b>EliteGames conference extractor</b></i>. ���������� ��������� ��������. ���������� ������� ��������: " & Format(MaxRetryAttempts)
            '����� �������� ��������
            lRetry = 0
        Else
            '������������� ����������
            bDone = False
            strMess = strEmpty
            '��������� ������ ��������
            Label14.Caption = "��������� ������ �������� �� ������� ������"
            Label14.Refresh
            Call StartInet(Left(hMsg.msgURL, hMsg.msgURLLen))
            Timer1.Enabled = True
            Exit Function
        End If
    End If
    If LastTopicNum <> hMsg.TopicNum Then
        '������� ��������� ��������� ��� � ��������� ����
        TopicChanged = True
    End If
    '�������� ����� ������� ����
    LastTopicNum = hMsg.TopicNum
    '���� ���� ����������, ��...
    If TopicChanged Then
        '������ ���� ������������� �� ������� ����
        COLOR = ColorLight
        '����� �������� ���� ����� �����, ������ ����
        bColor = True
        '��������� ���� � HTML � ��������� �����
        bResult = WriteHTML(6, CurrPage, COLOR, strEmpty, RecNum)  '��������� �������
        bResult = WriteHTML(5, CurrPage, COLOR, strEmpty, RecNum)  '����� ������� � ����
        bResult = WriteHTML(1, CurrPage, COLOR, Left(hMsg.msgHead, hMsg.msgHeadLen), RecNum)
        '�������� ������ � ����������: ����� ����
        bResult = WriteNavigator(3, CurrPage, RecNum, 0, strEmpty)
        '�������� ���� ����� ����, ����� ������ ������ ����� ������������ ��� �����
        TopicChanged = False
    End If
    '����������� ���� ����������
    '���� � ����� ������ ����������
    If Not (hMsg.emailLen = 0) And (FullLogging Or bFullLogging) Then
        '����� ��� ��������� ����������
        bResult = WriteHTML(9, CurrPage, COLOR, Left(hMsg.email, hMsg.emailLen), RecNum)
        bResult = WriteHTML(10, CurrPage, COLOR, Left(hMsg.Author, hMsg.AuthorLen), RecNum)
    Else
        '������ ��� �������� e-mail
        bResult = WriteHTML(2, CurrPage, COLOR, Left(hMsg.Author, hMsg.AuthorLen), RecNum)
    End If
'{Fixed} TO DO WHAT ARE HELL&&& DateLen = -1!!!!! RecNum=12, 72
    If Not (hMsg.msgDateLen < 0) Then
        bResult = WriteHTML(3, CurrPage, COLOR, Left(hMsg.msgDate, hMsg.msgDateLen), RecNum)
    Else
        bResult = WriteHTML(3, CurrPage, COLOR, strEmpty, RecNum)
    End If
    bResult = WriteHTML(8, CurrPage, COLOR, Left(hMsg.msgHead, hMsg.msgHeadLen), RecNum)
    bResult = WriteHTML(4, CurrPage, COLOR, tmpString, RecNum)
    '����� ���������
    bResult = WriteNavigator(1, CurrPage, RecNum, hMsg.Indent, Left(hMsg.Author, hMsg.AuthorLen))
    bResult = WriteNavigator(2, CurrPage, RecNum, 0, Left(hMsg.msgHead, hMsg.msgHeadLen))
    '������ ���� ���� ��� ���������� ���������
    bColor = Not bColor
    COLOR = IIf(bColor, ColorLight, ColorDark)
    ''Select Case bColor
    ''    Case True:
    ''        Color = ColorLight
    ''    Case False:
    ''        Color = ColorDark
    ''End Select
    '�������� ��������� ��������� ������
    RecNum = MsgMap.GetNext
    If RecNum = 0 Then
        Call StopFeeder(ucDLcomplete)
        '������� ��� ������� �������: ������ ��������
        Exit Function
    Else
        '��������� ������, ��������� �������!
        '���������� ��� ��������� ��������� �� MsgMap ���� �� �� ������������...
        With hMsg
            .Author = MsgMap.Author
            .AuthorLen = MsgMap.AuthorLen
            .email = MsgMap.email
            .emailLen = MsgMap.emailLen
            .msgDate = strEmpty
            .msgDateLen = 0
            .msgHead = MsgMap.msgHead
            .msgHeadLen = MsgMap.msgHeadLen
            .MsgSize = 0
            .msgURL = MsgMap.msgURL
            .msgURLLen = MsgMap.msgURLLen
            .TopicNum = MsgMap.TopicNum
            .Indent = MsgMap.Indent
        End With
        '������������� ����������
        bDone = False
        strMess = strEmpty
        '��������� ������� ������� ��� �������� �� ����������
        Call StartInet(Left(hMsg.msgURL, hMsg.msgURLLen))
    End If
    '���������� ������, ���� �� ������ ������ ����
    If Not bCancel Then
        Timer1.Enabled = True
'��� ���� ����������� � mnuStop? ����� ���?
''    Else
''        Call StopFeeder
    End If
End Function

'*****************************************************
'* ���������� ����������� �������� ����� ����������  *
'* (��������) �������)                               *
'*      Mode=ucDLcomplete - ������� ���������        *
'*      Mode=ucDLstopped  - ������� �����������      *
'*****************************************************
Private Sub StopFeeder(ByVal Mode As enDLresult)
    '����� ������� (��� ���� �������)
    '��������� ��������
    bResult = Module1.WriteHTML(6, CurrPage, COLOR, strEmpty, RecNum)
    bResult = Module1.WriteHTML(7, CurrPage, COLOR, strEmpty, RecNum)
    '��������� ���������
    bResult = WriteNavigator(4, CurrPage, RecNum, 0, strEmpty)
    Select Case Mode
        Case ucDLcomplete:
            '�������� ������: �������� �������
            Label14.Caption = "���� ���������� ������� �������..."
            Label14.Refresh
            bResult = MsgMap.UpdateDB(ucHistoryMode)
            '�������� �� �������
            ''TO DO BUG � ���� �� ��� ������ ������ ���? � ��� ���� �����������...
                    ''Label14.Caption = "���� ���������� �������������� ����������..."
                    ''Label14.Refresh
                    ''bResult = Authors.UpdateDB(ucHistoryMode)
            '���������� ������� ������ ��� ������ ������� ������� ��������
            Label14.Caption = "������ �������� �������..."
            Label14.Refresh
            bResult = UpdateTOC(CurrPage)
        Case ucDLstopped:
            '�������� ������: ������� ����������� �������������
            bResult = MsgMap.UpdateDB(ucUseCurrRec)
    End Select
    '���������� ���������� �� �������
    Label14.Caption = "������ ���������� � HTML..."
    Label14.Refresh
    bResult = Authors.WriteAHTMLs
    '������� ��
    MsgMap.Clear
    Label14.Caption = "�������� �������."
    Label14.Refresh
    Label8.Caption = "������."
    Label8.Refresh
    mnuExit.Enabled = True
    mnuStop.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = False
End Sub

'*****************************************************
'* ��������� ����� �������� � ���������� �� �������  *
'* � ���������� �� ��������                          *
'*      msgNum - ����� ������ � �� ������            *
'*****************************************************
'lpString �������� ByRef (ByVal �� �������� �����)
Private Function ParserLevel2(ByRef lpString As String) As String
Dim StartPos As Long
Dim EndPos As Long
Dim tmpstr As String
Dim i As Long
Dim tmpLen As Long      '��� ��������� ���� ���������� ��� �������� ����������

''������� ������ �� ���������� ����������
''1)� �������� ��� ����������� ����������
''2)���� ������� ������ (������ ���������), �� � ���������� ��������� ������� (-)
''���������� - ����� ���, ��� �������� �� 2 ������ :( (-)
''�������    - �����: ���, ��� �������� �� 2 ������ :(
    '������ � ��� ���������
    LogEvent ("[Parser Level 2] enter routine")
    '��� ������ ����������� �������������� � ������ ��� �������
    Label14.Caption = "������ ���������� ��������"
    Label14.Refresh
    StartPos = InStr(1, lpString, "��������", vbTextCompare)
    If StartPos = 0 Then
        '�������� ������� - ��������� ������
        lRetry = lRetry + 1
        Exit Function
    End If
    '�������� ��� ������, ����� �� ��������
    '+1 ����� ����� "�" �������� :) ��� �� �����, �� ��� �������� :)
    lpString = Right(lpString, Len(lpString) - StartPos + 1)
    '��������� ����� �� ������ ���������
    ''{FIXED} BUG Invalid call, add 1 for continue
    StartPos = InStr(1, lpString, "</center>", vbTextCompare) + Len("</center>")
    If StartPos = 0 Then
        '�������� ������� - ��������� ������
        lRetry = lRetry + 1
        Exit Function
    End If
    '������ ���� ����� ���������
    EndPos = InStr(StartPos, lpString, ">��������</A>", vbTextCompare)
    If EndPos = 0 Then
        '�������� ������� - ��������� ������
        lRetry = lRetry + 1
        Exit Function
    End If
    '�������� ��� ������ (������� ����������� ������ (+1))
    lpString = Mid(lpString, StartPos, EndPos - StartPos + 1)
    'OK, ������ � ����������� �� ������� e-mail ������ �������� ���������
    If hMsg.emailLen = 0 Then
        StartPos = InStr(1, lpString, "</I></B>:", vbTextCompare) + Len("</I></B>:")
    Else
        StartPos = InStr(1, lpString, "</I></B></A>:", vbTextCompare) + Len("</I></B></A>:")
    End If
    ''EndPos = InStr(StartPos, lpString, "<P>", vbTextCompare)
    '��� ���������� ����� ���� ��� ���� ������� �� ������
    EndPos = InStr(StartPos, lpString, "</FONT> <P>", vbTextCompare)
    '�������, �� � ����� ������� ������ ������� ������ ������. ��� ���� ��������.
    If EndPos = 0 Then
        '�������� ������� - ��������� ������
        lRetry = lRetry + 1
        Exit Function
    End If
    ''TO DO BUG lpString="" ������ ��������� ����� ������ �� �������, ����������
    msgHead = Mid(lpString, StartPos, EndPos - StartPos)
    '�������� ������� (-), ������� �������� �������� �������. ����� ����� ��������
    '���������� ���������� lResult � StartPos ��� �������� �� ������� ��������
    '�-���! �� �� ����� � � (+)
    msgHead = RemoveSpaces(msgHead)
    StartPos = InStr(1, msgHead, "(-)", vbTextCompare)
    Do While StartPos > 0
    ''If StartPos > 0 Then
        '(-) ������. �������� (�����), ��������� �� �� � ������...
        If StartPos > Len(msgHead) - 5 Then
            '(-) � ����� ������ - ������ ������� �����
            msgHead = Left(msgHead, Len(msgHead) - 3)
        Else
            '(-) � ��������. ������� ������ ��� 3 �������
            msgHead = Left(msgHead, StartPos - 1) + Right(msgHead, Len(msgHead) - StartPos - 2)
        End If
    ''End If
        '������� �������
        msgHead = RemoveSpaces(msgHead)
        StartPos = InStr(1, msgHead, "(-)", vbTextCompare)
    Loop
    StartPos = InStr(1, msgHead, "(+)", vbTextCompare)
    Do While StartPos > 0
    ''If StartPos > 0 Then
        '(+) ������. �������� (�����), ��������� �� �� � ������...
        If StartPos > Len(msgHead) - 5 Then
            '(+) � ����� ������ - ������ ������� �����
            msgHead = Left(msgHead, Len(msgHead) - 3)
        Else
            '(+) � ��������. ������� ������ ��� 3 �������
            msgHead = Left(msgHead, StartPos - 1) + Right(msgHead, Len(msgHead) - StartPos - 2)
        End If
    ''End If
        '������� �������
        msgHead = RemoveSpaces(msgHead)
        StartPos = InStr(1, msgHead, "(+)", vbTextCompare)
    Loop
    '������ � ��� ���������
    LogEvent ("[Parser Level 2] ��������� ����: " + msgHead)
    '������������� � ������ ���������
    StartPos = EndPos + Len("</FONT> <P>")
    '��������: � ��������� ����� ���� ��������� ����� <P>, �� �� ������
    '������ ���� � ����� ���������.
    '����������� ������� ������������ ����������� ����, ��, � ���������, �� ����� 0:
    '���� �� ���� ������ � ������ ����������� ���� � ���������...
    '����� ������ �������: <P><I>27 ������ 2001, 01:04:49</I> <FONT size=2><B>
    '����� ���, ����������� ���� � ��, ��� �� ���
    lResult = InStr(StartPos, lpString, "</I> <FONT size=2><B>", vbTextCompare)
    '��� ������������ ��������, 30 �������� ����� ����������� ���� ��� <I> ���
    '�� ������ ��� ��������, ������� ������ �� ������� ������� ������ ��, ��� ����
    lResult = lResult - 30
    '�������� ������ ������������ ����� ���������
    EndPos = InStr(lResult, lpString, "<I>", vbTextCompare)
    '����� ������ ��������� ������� ���� ������������ ������������ ����������
    '����� ������� ����� ������
    '����� ������� - ��������� ���������� ����������: ��� �� �� � ��������� book.ru?
    ''TO DO �������� �������� �������� If Left(hMsg.msgHead, hMsg.msgHeadLen) = msgHead Then
    ''�������� ���������� ���� ����
    tmpLen = Len(msgHead)
    If Not (hMsg.msgHeadLen = tmpLen) Then
        '������ � ��� ���������
        LogEvent ("[Parser Level 2] ����� ��������� ���� ������� �� ���������� � Parser Level 1")
        If hMsg.msgHeadLen > tmpLen Then
            '��������� ������, ��� ���������� � ParserLevel1
            '�������� ��������� � �����
            msgHead = msgHead & String(hMsg.msgHeadLen - tmpLen, Chr(32))
            '������ � ��� ���������
            LogEvent ("[Parser Level 2] ��������� ���� ������, ��� ���������� � Parser Level 1")
        'Else
            '��������� �������, ��� ���������� � ParserLevel1
            '�������� �������� tmpLen, �������� �� �������� �� hMsg
            '(����� End If, �.�. ��� ����� ������� � ����� �������)
        End If
        tmpLen = hMsg.msgHeadLen
    End If
    If Left(hMsg.msgHead, hMsg.msgHeadLen) = Left(msgHead, tmpLen) Then
    '��, ��� ���� � �� �� ���������
        '�������� ����� ������ ���������
        tmpstr = Mid(lpString, StartPos, EndPos - StartPos)
        '����� ��������� ����������� ����� ��������� <P>. ������, ����� �� �������
        '������ ����� �� �������� ��������
        Do While UCase(Right(tmpstr, 3)) = "<P>"
            tmpstr = Left(tmpstr, Len(tmpstr) - 3)
        Loop
        ParserLevel2 = tmpstr
        hMsg.MsgSize = Len(tmpstr)
        '��� ���� ��������...
        StartPos = EndPos + Len("<I>")
        EndPos = InStr(StartPos, lpString, "</i>", vbTextCompare)
        tmpstr = Mid(lpString, StartPos, EndPos - StartPos)
        hMsg.msgDate = tmpstr
        hMsg.msgDateLen = Len(tmpstr)
        '��������� ��������� �������� � �������
        Label14.Caption = "���������� ���������� ����������"
        Label14.Refresh
                '�����! ���������� ���� � �����! �����������...
                '��! ������ �� ���� ��������� ����� �������� (����� ���������, ����, �� �����)
'������ � ��� ���������
LogEvent ("[Parser Level 2] ���� ���������� ���������� ��� ������ " + Left(hMsg.Author, hMsg.AuthorLen))
        With hMsg
            MsgMap.AuthorLen = .AuthorLen
            If .AuthorLen > 0 Then
                MsgMap.Author = Left(.Author, .AuthorLen)
            Else
                MsgMap.Author = strEmpty
            End If
            MsgMap.emailLen = .emailLen
            If .emailLen > 0 Then
                MsgMap.email = Left(.email, .emailLen)
            Else
                MsgMap.email = strEmpty
            End If
            MsgMap.msgDateLen = .msgDateLen
            If .msgDateLen > 0 Then
                MsgMap.msgDate = Left(.msgDate, .msgDateLen)
            Else
                MsgMap.msgDate = strEmpty
            End If
            MsgMap.msgHeadLen = .msgHeadLen
            If .msgHeadLen > 0 Then
                MsgMap.msgHead = Left(.msgHead, .msgHeadLen)
            Else
                MsgMap.msgHead = strEmpty
            End If
            MsgMap.MsgSize = .MsgSize
            MsgMap.msgURLLen = .msgURLLen
            If .msgURLLen > 0 Then
                MsgMap.msgURL = Left(.msgURL, .msgURLLen)
            Else
                MsgMap.msgURL = strEmpty
            End If
            MsgMap.TopicNum = .TopicNum
            MsgMap.Indent = .Indent
        End With
'������ � ��� ���������
LogEvent ("[Parser Level 2] ���������� ���������� ���������� - ����� MsgMap.UpdateDB(ucMsgURLsMode, " + Str(RecNum) + ")")
        bResult = MsgMap.UpdateDB(ucMsgURLsMode, RecNum)
    Else
        '������ � ��� ���������
        LogEvent ("[Parser Level 2] ��������� ���� �� ������������� ���������, ����������� � Parser Level 1. ������ ����� " + Str(RecNum))
        '��������� �����������! ������...
        hMsg.MsgSize = 0
        hMsg.msgDate = strEmpty
        hMsg.msgDateLen = -1
        ParserLevel2 = "<i>���������� <b>EliteGames conference extractor</b></i>. ������ ���������� �� ������� www.book.by.ru ��������� ������������ ������ � �� ����� ���� ���������"
        Label17.Caption = Str(Val(Label17.Caption) + 1)
    End If
    With Authors
        '������ � ��� ���������
        LogEvent ("[Parser Level 2] ������ �������� ���������� �� ������� � cAuthors")
        If hMsg.emailLen > 0 Then
            .email = Left(hMsg.email, hMsg.emailLen)
        Else
            .email = strEmpty
        End If
        .emailLen = hMsg.emailLen
        If hMsg.AuthorLen > 0 Then
            .Nickname = Left(hMsg.Author, hMsg.AuthorLen)
        Else
            .Nickname = strEmpty
        End If
        .NicknameLen = hMsg.AuthorLen
        '��������� ���� �� ����� �������� ����������, �������� (? ����???)
        .TotalNums = 1
        .TotalSize = hMsg.MsgSize
        If Not (hMsg.msgDateLen <= 0) Then
            '��������� ���� � �������� ������
            '������ ������� YYYY.MM.DD (HH:MM:SS)
            
            '�������� �����
            tmpstr = strEmpty
            tmpLen = hMsg.msgDateLen
            For i = 3 To tmpLen
                If Not (Mid(hMsg.msgDate, i, 1) Like "[0-9]") Then
                    tmpstr = tmpstr & Mid(hMsg.msgDate, i, 1)
                Else
                    i = tmpLen
                End If
            Next i
            tmpstr = RemoveSpaces(tmpstr)
            Select Case tmpstr
                Case "������":      tmpstr = "01"
                Case "�������":     tmpstr = "02"
                Case "�����":       tmpstr = "03"
                Case "������":      tmpstr = "04"
                Case "���":         tmpstr = "05"
                Case "����":        tmpstr = "06"
                Case "����":        tmpstr = "07"
                Case "�������":     tmpstr = "08"
                Case "��������":    tmpstr = "09"
                Case "�������":     tmpstr = "10"
                Case "������":      tmpstr = "11"
                Case "�������":     tmpstr = "12"
                '�� ������ �������� ��� ��� ����
                Case Else:          tmpstr = "13"
            End Select
            '�������� �����, �������� ����������� .MM.DD
            .FirstDate = "." & tmpstr & "." & Left(hMsg.msgDate, 2)
            '�������� ���
            For i = 3 To tmpLen
                If Mid(hMsg.msgDate, i, 1) Like "[0-9]" Then
                    tmpstr = Mid(hMsg.msgDate, i, 4)
                    i = tmpLen
                End If
            Next i
            '�������� YYYY.MM.DD_
            .FirstDate = tmpstr & .FirstDate & Chr(32)
            '�������� �����
            tmpLen = InStr(1, hMsg.msgDate, "(", vbTextCompare)
            tmpstr = Mid(hMsg.msgDate, tmpLen, 10)
            .FirstDate = .FirstDate & tmpstr
            .FirstDateLen = 21
        Else
'������ � ��� ���������
LogEvent ("[Parser Level 2] � �������� ������ ���� <=0")
            .FirstDate = strEmpty
            .FirstDateLen = 0
        End If
    End With
'������ � ��� ���������
LogEvent ("[Parser Level 2] ���������� ���������� �� ������ - ����� Authors.UpdateDB(ucPrimaryMode)")
    If Authors.UpdateDB(ucPrimaryMode) Then
        '�������� �������� �������� ������������ �������
        Label6.Caption = Str(Val(Label6.Caption) + 1)
        '����� � ������ eMail-������ ������, ���� ��� ���� ���������� � UpdateDB:
        hMsg.emailLen = Authors.emailLen
        hMsg.email = Left(Authors.email, Authors.emailLen)
    Else
'������ � ��� ���������
LogEvent ("[Parser Level 2] ������ ������ Authors.UpdateDB(ucPrimaryMode)")
        'TO DO �������� Label(Thread) +1 ���������� + ������!!!
        Label6.Caption = Str(Val(Label6.Caption) + 1)
        Label6.Caption = Str(Val(Label16.Caption) + 1)
    End If
    '������ ������� ����������: ��������� ������� �� ��������� (lRetry = 0)
    lRetry = 0
    '���!!!
End Function

Private Sub Command1_Click()
    Call mnuStart_Click
End Sub

'*****************************************************
'* ������ ������ ���� - ������������� �������        *
'*****************************************************
Private Sub Command2_Click()
    Call mnuStop_Click
End Sub

Private Sub Command3_Click()
    '��������� ���� "������� �������� �����������", ����� ������
    '�����������, ��������� ��� ���
    bDone = True
''{NET ERROR}    Call Inet1_DocumentComplete(Nothing, "dfg")
End Sub

Private Sub Form_Load()
    Loading = True  '�������� - �� ������������ ����������
    strMess = strEmpty
    Inet1.Offline = WorkOffline
    WebBrowser1.Offline = WorkOffline
    '����� ������� ����� ����������� ��� � 2 �������
    'TO DO � ������� - �� ��������
    Timer1.Interval = 1500
    Timer1.Enabled = False
    Command2.Enabled = False
    mnuStop.Enabled = False
    mnuStart.Enabled = False
    '������ ������ ��������� ������ � ������ ������
    Command3.Enabled = False
End Sub

Private Sub Form_Resize()
''    If frmHistory.ScaleWidth < 10500 Then
''        frmHistory.Width = 10500
''    End If
''    If frmHistory.ScaleHeight < 5000 Then
''        frmHistory.Height = 5000
''    End If
''    Command1.Top = frmHistory.ScaleHeight - 600
''    Command2.Top = frmHistory.ScaleHeight - 600
''    Command3.Top = frmHistory.ScaleHeight - 600
''    Command4.Top = frmHistory.ScaleHeight - 600
''    Command1.Left = (frmHistory.ScaleWidth - 10410) \ 2 + 210
''    Command2.Left = Command1.Left + 2100
''    Command3.Left = Command2.Left + 3360
''    Command4.Left = Command3.Left + 3045
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'TO DO ��������� ��������� ������ � ������� ����
    Set MsgMap = Nothing
    Set Authors = Nothing
    Set HTMLDoc = Nothing
    Inet1.Stop
    Set fBrw = Nothing
End Sub

Private Sub Inet1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If Not Loading Then
        '����� ����������� ������� DocumentComplete
        If (pDisp Is Inet1.Object) Then  'Is ��������� ������������ ��������
''{NET ERROR}        If (pDisp Is Inet1.Object) Or (pDisp Is Nothing) Then  'Is ��������� ������������ ��������
            '������ HTML-�������� ���������
            Set HTMLDoc = Inet1.Document
            If HTMLDoc Is Nothing Then
            ''TO DO ���������� ����� (����� ���������)
            ' Not an HTLM document
                Exit Sub
            End If
            vtData = HTMLDoc.body.innerHTML
            '������� ��� �������, � �� ���������� � ����
                ''TO DO ���������� �����
                Label8.Caption = "������ ���������� �� ������"
                Label8.Refresh
            strMess = Replace(CStr(vtData), """", "`", 1, -1, vbBinaryCompare)
                ''TO DO ���������� �����
                Label8.Caption = "������ ����������� ������"
                Label8.Refresh
            strMess = ClearGarbage(strMess)
            '������ ������ ��������� ������ � ������ ������
            Command3.Enabled = False
            Command3.Refresh
            bDone = True    '�������� ��������
        End If
    End If
End Sub

Private Sub Inet1_DownloadBegin()
    ''TO DO ���������� �����
    Label8.Caption = "������� �����..."
    Label8.Refresh
End Sub

Private Sub Inet1_DownloadComplete()
'TO DO ��������!!! ���� DocumentComplete ����� ��������, �� ��� ��������� ���� ����� �������
Exit Sub
    '������ HTML-�������� ���������
    Set HTMLDoc = Inet1.Document
    If HTMLDoc Is Nothing Then
    ''TO DO ���������� ����� (����� ���������)
    ' Not an HTLM document
        Exit Sub
    End If
    strMess = Replace(HTMLDoc.body.innerHTML, """", "`", 1, -1, vbBinaryCompare)
''''''    vtData = HTMLDoc.body.innerHTML
    '������� ��� �������, � �� ���������� � ����
        ''TO DO ���������� �����
        Label8.Caption = "������ ���������� �� ������"
''''''    strMess = Replace(CStr(vtData), """", "`", 1, -1, vbBinaryCompare)
        ''TO DO ���������� �����
        Label8.Caption = "������ ����������� ������"
    strMess = ClearGarbage(strMess)
    '������ ������ ��������� ������ � ������ ������
    Command3.Enabled = False
    Command3.Refresh
    bDone = True    '�������� ��������
End Sub

Private Sub mnuAbout_Click()
    Load frmAbout
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuConnect_Click()
'������� ������� ������������

Dim LastRecNum As Long
Dim i As Long
Dim blValue As Long '���� Boolean, �� ���� ����� �� ����� ��������, �������� �������� :(
Dim hFile As Long
Dim NextPage As Long    '����� ��������� ������������ ��������
    
    '��������� ������ �������
    hFile = FreeFile
    Open (App.Path & DataPath & PagesFile) For Random As hFile Len = Len(blValue)
    LastRecNum = LOF(hFile) \ Len(blValue)
    NextPage = 0
    If LastRecNum > 0 Then
        '��� ���-�� ��������, ��������, ��� �� ��� ������ ������������?
        For i = 1 To LastRecNum
            Get hFile, i, blValue
            If blValue = 0 Then
                '���, ���� �����, ���������� ����� ���� �� ����� (������ ����������)
                NextPage = i
                i = LastRecNum
            End If
        Next i
        If NextPage = 0 Then
            '� ��������� (1, LastRecNum) ��� �������� ��������
            If Not (LastRecNum >= 80) Then
                '���� �� ��� 80 ������� ��� ��������, �� ���������� ������ ��������
                '������ � ���� �����, ��� ��������� � ������� ��� (LastRecNum)
                NextPage = LastRecNum + 1
            End If
        End If
    End If
    Close hFile
    Text2.Text = Format(NextPage)
    
    '���������� ��������� �������� ������ (Win ������������� �������� ������ �� ����������)
''DEBUG mnuStart.Enabled = True
    WorkOffline = False
    WebBrowser1.Offline = WorkOffline
    WebBrowser1.Navigate ("http://book.by.ru/cgi-bin/book.cgi?book=Elitegames")
End Sub

Private Sub mnuExit_Click()
    '��������� ������ ���� ���������
    Call StopLogging
    Unload Me
    End
End Sub

Private Sub mnuF003_Click()
    Call Test
End Sub

Private Sub mnuHelpContents_Click()
   Call ShowCHMHelp
End Sub

'*****************************************************
'* ���������� ����� ������������� ��������           *
'*****************************************************
Private Function GetPageNum() As Long
    'TO DO �������� ����������� ����������� ������ �������� (� ������ ������!!!)
    GetPageNum = CLng(Val(Text2.Text))
    If GetPageNum > LastPageAvailable Then
        'TO DO Warning message
        GetPageNum = LastPageAvailable
    ElseIf GetPageNum < 1 Then
        GetPageNum = 1
    End If
End Function

'*****************************************************
'* ������ ����� ���� �����                           *
'* ���������� �������� � ���������� ��������         *
'*****************************************************
Private Sub mnuStart_Click()
Dim strURL As String
    Loading = False '������ ������������ �������� �������
    Inet1.Offline = WorkOffline
    '�������� ����� �������� ��� �������
    CurrPage = GetPageNum
    ''TO DO ���� ����� ����� �����������, �� �����!!!
    '��������� ������� ����� � ���� � �� ������
    Command1.Enabled = False
    mnuStart.Enabled = False
    '�������� ����������� ������� ������ Cancel
    bCancel = False
    Command2.Enabled = True
    mnuStop.Enabled = True
    '������ �������� ��� ���������� �������
    mnuExit.Enabled = False
    '����� ���������� �� �������
    Label16.Caption = "0"
    Label17.Caption = "0"
    Label2.Caption = "0"
    Label4.Caption = "0"
    Label6.Caption = "0"
    '��������� �������� ������
    ''TO DO ����� �������� ����� (��������� �����������)
    strURL = "/book.cgi?book=Elitegames&p=" & Format(CurrPage) & "&ac=1"
    ''strURL = "C:\tmp\test8.htm"
    '������ ������� ���� �� �����.
    ''Command3.Enabled = False
    ''Command3.Refresh
    bDone = False
    GettingURLs = True
    StartInet (strURL)
    '���������� ��������� �� StartInet
    Label8.Caption = "������ ���������� ����"
    Label8.Refresh
    '���� �������!
    Timer1.Enabled = True
    Debug.Print "������ ���������� ����!"
End Sub

'*****************************************************
'* ��������� ������� � �������� ������ ��  msgURL    *
'* �������� ���������� � ��� ��������� - �           *
'* Inet_DocumentComplete � Feeder ��������������     *
'*****************************************************
Private Sub StartInet(ByVal msgURL As String)
    '���� ����������� ���������� �������
    Command3.Enabled = True
    Command3.Refresh
    bDone = False
'' msgURL = "C:\tmp\test2.htm"
'' msgURL = "C:\tmp\test3.htm"
'' msgURL = "C:\tmp\test5.htm"
''Inet1.Navigate (msgURL)
''Exit Sub
    
    'Inet1.Navigate ("http://book.by.ru/cgi-bin/book.cgi?book=" & msgURL)
    Inet1.Navigate (BaseHREFprefix & BaseHREF & msgURL)
    Debug.Print "������ ����, " & BaseHREFprefix & BaseHREF & msgURL
    Label8.Caption = "������ ����"
    Label8.Refresh
End Sub

'*****************************************************
'* ������ ����� ���� ���� - ������������� �������    *
'*****************************************************
Private Sub mnuStop_Click()
    '���� ������� ������ ���� - �� ������ ������
    bCancel = True
    '��������� ������ ����
    Command1.Enabled = True
    '��������� ������ ���� ������ � �����
    mnuStart.Enabled = True
    mnuExit.Enabled = True
    '���������� ������� ��� ������� Cancel (Stop)
    Inet1.Stop
    Label8.Caption = "������� �����������"
    Label8.Refresh
    Call StopFeeder(ucDLstopped)
    'TO DO ����������� ��� � ���������� ������
    Result = MsgBox("������� ������� �������� ��������. ��� �� �����," + vbCrLf + "�� ������ ����������� ��� ����������, �������" + vbCrLf + "������� �������� �� ������� ���������." + vbCrLf + "�������������� �������� ""��������� �������"" ����������" + vbCrLf + "������� ��� ��������� ����������.", vbOKOnly + vbExclamation, "������� �������� " & Format(CurrPage) & " �����������")
End Sub

'*****************************************************
'* ��������� HTML ���� � ������������ �������        *
'*****************************************************
Private Sub mnuView_Click()
    Call ShowHTML
End Sub

'*****************************************************
'*   ��������� HTML-����� �� ���������� �� �������   *
'*****************************************************
Private Sub mnuWriteAuthors_Click()
    '���������� ���������� �� �������
    Label14.Caption = "������ ���������� � HTML..."
    Label14.Refresh
    bResult = Authors.WriteAHTMLs
    Label14.Caption = "�������� �������."
    Label14.Refresh
    Label8.Caption = "������."
    Label8.Refresh
End Sub

Private Sub Timer1_Timer()
    If GettingURLs Then
    '������: ������ ���� ��������� ����������
        If bDone Then
            '�� ������ ������ ��������� ������� �� (���� �� ��� �������� �� ����������� ����)
            MsgMap.Clear
            ''Call Inet1_DocumentComplete(Nothing, "dfg")
            If Splitter(strMess) Then
                GettingURLs = False
                bResult = InitFeeder
            End If
        End If
    Else
    '������ �������� ����� ���������
        If bDone Then
            lResult = Feeder
        End If
    End If
End Sub

Private Sub WebBrowser1_DownloadBegin()
    '���� ������� ����� �������� ������, �� �� � ���� - ����� ��������� �������
    Command1.Enabled = True
    mnuStart.Enabled = True
End Sub
