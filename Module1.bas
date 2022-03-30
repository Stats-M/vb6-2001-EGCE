Attribute VB_Name = "Module1"
Option Explicit

'���������� �������� FullLogging = 1 ��� ��������� ������ e-mail ������� � HTML-����
Public Const FullLogging = 1

Public Type tAuthor
    Nickname As String * 255
    NicknameLen As Long
    TotalSize As Long
    TotalNums As Long
    email As String * 255
    emailLen As Long
    'Add-on
    FirstDate As String * 255
    FirstDateLen As Long
End Type

Public Type tMessage
    TopicNum As Long
    msgURL As String * 255
    msgURLLen As Long
    Author As String * 255
    AuthorLen As Long
    email As String * 255
    emailLen As Long
    msgHead As String * 255
    msgHeadLen As Long
    MsgSize As Long
    msgDate As String * 255
    msgDateLen As Long
    Indent As Long
End Type

Public Type tHistory
    URLvalue As Long
    TopicNum As Long
    AuthorID As Long
    MsgSize As Long
    Reserve As Long
End Type

Public Const AuthorsFile = "authors.db"     '��� ���� ������ �� �������
Public Const MsgFile = "msg.db"             '��� ���� ������ ������ �� ���������
Public Const HistoryFile = "history.db"     '��� �������
Public Const PagesFile = "pages.db"         '��� ���� ������ �� ���������
Public Const IndexFile = "index.html"       '������� ������
Public Const TOCFile = "pagesTOC.html"      '�������� ����������
Public Const StartupFile = "Startup.html"   '��������� ��������
Public Const FAQFile = "FAQ.html"           '����
Public Const PagesPrefix = "page"           '������� ����� �������� ���������
Public Const NavPrefix = "pTOC"             '������� ����� ��������-����������
Public Const IndexPrefix = "pIndex"         '������� ����� ������� �������� ���������
Public Const HTMLext = ".html"              '���������� ���� �������
Public Const ANamesPage = "authors1.html"   '������ �� ��������
Public Const ANumsPage = "authors2.html"    '������ �� ���-�� ���������
Public Const ASizePage = "authors3.html"    '������ �� ������� ���������
Public Const ATimePage = "authors4.html"    '������ �� ������� ��������� � �����
Public Const HelpFile = "\EGCEhelp.chm"     '��� ����� �������
Public Const CSSFile = "EGCE.css"           '������� ������ HTML
Public Const PaneFile = "pane.js"           '������� ������ HTML
Public Const DataPath = "\EGCEdata\"        '���� � ������ ������
Public Const strEmpty = ""
Public Const ColorLight = "ffffff"          '��������� ����� ��� HTML-��������
Public Const ColorDark = "ddddff"
Public Const LastPageAvailable% = 80        '����� ��������� �������� ������ �����
Public Const iMaxSize% = 255                '������������ ������ ������
Public Const strSepURLDir = "/"             '����������� URL-�������
Public Const strSepDir = "\"                '����������� ����������
Public Const strHHelpEXEname = "hh.exe"     '��� ��������� ��������� ������� *.CHM
Public Const strExplorer = "explorer.exe"
Public Const strFullLogging = "full"

''HKEY_LOCAL_MACHINE\Software\CLASSES\chm.file\shell\open\command

Public PauseTime As Variant, Start As Variant    '��� ���������� ���������
Public fBrw As frmBrowser           '���������� ���� ���������
Public Result As VbMsgBoxResult     '��� ������ ���������
Public bResult As Boolean           '��� ����������� �������
Public WorkOffline As Boolean       '�������� � ���-���� ��-���������
Public hMsg As tMessage             '������ ��� �������
Public MsgMap As cMsgMap            '����� ��� ������ � ����� ������ ���������
Public Authors As cAuthors          '����� ��� ������ � ����� ������ �������
Public hFile As Long                '��������� ��� ������ � �������
Public CurrPage As Long             '����� ������������ ��������
Public RecNum As Long               '����� ������ � �� ������ (��� �������� �����)
Public BaseHREF As String           '������� ����� �����������
Public BaseHREFprefix As String     '������� �������� ������ �����������
Public TRindex As Long              '������ �������� ���� TR (��� ���������� ���������)
Public bFullLogging As Boolean

'������� DLL-�������
Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'********************************************
'* ��������� ���� ��������� Windows �����   *
'* Win API                                  *
'* ���������� ���� �������� �����������     *
'* ����������� ���������� \                 *
'********************************************
Function GetWindowsDir() As String
Dim strBuf As String
Dim iZeroPos As Integer

    '��������� ����� ���������
    strBuf = Space(iMaxSize)
    If GetWindowsDirectory(strBuf, iMaxSize) > 0 Then
        '���� ���������� ������
        iZeroPos = InStr(strBuf, Chr$(0))
        '���� ���������� ����, �� ������� ���
        If iZeroPos > 0 Then
            strBuf = Left$(strBuf, iZeroPos - 1)
        End If
        '���� �� ����� ������ ��� ����������� ����������, ��������� ���
        If Right(Trim(strBuf), Len(strSepURLDir)) <> strSepURLDir And _
           Right(Trim(strBuf), Len(strSepDir)) <> strSepDir Then
            strBuf = RTrim$(strBuf) & strSepDir
        End If
        GetWindowsDir = strBuf
    Else
        GetWindowsDir = vbNullString
    End If
End Function

'************************************************************
'* ������ ���������� ������� Windows (������ ������� *.CHM) *
'* ����� ����� hh.exe ����� ������ ������������� �� �����,  *
'* ��������� �� ��, ��� ���� ���� � ����������� �������     *
'* ����� � ����� Windows                                    *
'************************************************************
Public Sub ShowCHMHelp()
Dim RetValue As Double
    '�������� ���� � ����� Windows ����� DLL call
    RetValue = Shell(GetWindowsDir & strHHelpEXEname & Chr(32) & App.Path & HelpFile, vbMaximizedFocus)
End Sub

'************************************************************
'* ���������� HTML �������� � ����������� ���� ����������   *
'* �������.                                                 *
'* �������� �������� �������������� ����� �������� �� ����� *
'* ���������� � �������� ���������                          *
'************************************************************
Public Sub ShowHTML()
Dim RetValue As Double
    '�������� ��������
    RetValue = Shell(GetWindowsDir & strExplorer & Chr(32) & App.Path & DataPath & IndexFile, vbMaximizedFocus)
End Sub

Function GetCommandLine() As Boolean
Dim CmdLine As String
    On Error Resume Next
    
    GetCommandLine = False
    CmdLine = Command()
    CmdLine = LCase(CmdLine)
    If Len(CmdLine) > 0 Then
        If InStr(1, CmdLine, strFullLogging, vbTextCompare) <> 0 Then
            GetCommandLine = True
        End If
    End If
    
End Function

Sub Main()
    
    frmSplash.Show
    frmSplash.Refresh
    PauseTime = 2   ' Set duration.
    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
        DoEvents   ' Yield to other processes.
    Loop
    '����� ���� ���� ������ �� �������� ���������... :))
    
    '���� ������� ��������, ������ ������� ������
    Set fBrw = New frmBrowser
    '������������� ������� (�������� �� frmBrowser_Load)
    Set MsgMap = New cMsgMap
    Set Authors = New cAuthors
    
    '��������� ������ ���� ���������
    Call StartLogging
    
    '������������� ��������� ����������
    RecNum = 0
    CurrPage = 1
    TRindex = 0
    bFullLogging = GetCommandLine   '���������� ��� ��� � ����������� ��������� ������
    
    If Dir(App.Path & "\EGCEdata\", vbDirectory) = strEmpty Then
        MkDir App.Path & "\EGCEdata"
    End If
    '���� ������� ������ ��� ���, ������� ��
    If Dir(App.Path & DataPath & CSSFile) = strEmpty Then
        WriteCSS
    End If
    '���� ������� ���������� ������� ��� ���, ������ ���
    If Dir(App.Path & DataPath & PaneFile) = strEmpty Then
        WritePane
    End If
    '��������� ��� ����� �������
    App.HelpFile = App.Path & HelpFile
    
    WorkOffline = True  '��� ������������� ������ �� �����
    'TO DO ���� ���, ����� ���������...
    BaseHREFprefix = "http://book.by.ru/cgi-bin//"
    
    Load fBrw
    Unload frmSplash

    fBrw.Show
End Sub

'*****************************************************
'*      ���������� ���������� �������� � HTML        *
'*  Mode = 0 �������� ������ ��������                *
'*  Mode = 1 �������� ��������� ���������            *
'*           (��� ������ ����� �����)                *
'*  Mode = 2 �������� ������ ���������               *
'*  Mode = 3 �������� ���� ���������                 *
'*  Mode = 4 �������� ��������� ���������            *
'*           (�������� ������ ����� Mode=8)          *
'*  Mode = 5 ����� ����                              *
'*  Mode = 6 ����� ����                              *
'*  Mode = 7 ����� ��������                          *
'*  Mode = 8 �������� ��������� � ������ ������      *
'*           (����� Mode=4)                          *
'*  Mode = 9 ������ ������ �� e-mail ������          *
'*  Mode =10 ������ ������ (����� ������ 9)          *
'*      Page    - ����� ��������                     *
'*      Color   - ���� ���� ������ ���������         *
'*      sMsg    - ������������ ��������              *
'*      RecNum  - ����� ������ � �� ������           *
'*****************************************************
Public Function WriteHTML(ByVal Mode As Long, ByVal Page As Long, ByVal COLOR As String, ByVal sMsg As String, ByVal RecNum As Long) As Boolean
    WriteHTML = False
    hFile = FreeFile
    Open (App.Path & DataPath & PagesPrefix & Format(Page) & HTMLext) For Append As hFile
    Select Case Mode
        Case 0: '�������� ������ ��������
            Print #hFile, "<HTML><HEAD><TITLE>EGCE - �������� " & Format(Page) & " - ����������� Elite Games �� WWW.BOOK.BY.RU</TITLE>"
            Print #hFile, "<LINK HREF=""" & CSSFile & """ REL=stylesheet TYPE=""text/css"">"
            Print #hFile, "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; CHARSET=Windows-1251"">"
            Print #hFile, "<META http-equiv=Pragma content=no-cache>"
            Print #hFile, "<META name=GENERATOR content=""EGCE"">"
            ''Print #hFile, "<STYLE type=text/css>"
            ''Print #hFile, "BODY{"
            ''Print #hFile, "a           {text-decoration:none; color:#006600}"
            ''Print #hFile, "a:link      {text-decoration:none; color:#006600}"
            ''Print #hFile, "a:active    {text-decoration:none; color:#00AA00}"
            ''Print #hFile, "a:hover     {text-decoration:underline; color:#00AA00}"
            ''Print #hFile, "a:visited   {text-decoration:none; color:#0000AA}"
            ''Print #hFile, "}</STYLE>"
            Print #hFile, "<SCRIPT language=JavaScript>"
            Print #hFile, "function NavigatorScroll(tgt)"
            Print #hFile, "{"
            Print #hFile, "    var target=window.parent.frames;"
            Print #hFile, "    var coll = target(0).document.all.tags(""P"");"
            Print #hFile, "    if (coll.length>0)"
            Print #hFile, "    {"
            Print #hFile, "        coll(eval(tgt-1)).scrollIntoView (false);"
            Print #hFile, "    }"
            Print #hFile, "}"
            Print #hFile, "</SCRIPT>"
            Print #hFile, "</HEAD><BODY><CENTER>"
        Case 1: '�������� ��������� ��������� (��� ������ ����� �����)
            ''Print #hFile, "<TR><TD vAlign=center width=""20%"" bgColor=#000090><FONT color=#FFFFFF face=""Verdana, Arial, Helvetica, Geneva"" size=1><B>�����, ����</B></FONT></TD>"
            ''Print #hFile, "<TD vAlign=center bgColor=#000090><FONT color=#FFFFFF face=""Verdana, Arial, Helvetica, Geneva"" size=1><B>����� " & Str(hMsg.TopicNum) & ": " & sMsg & "</B></FONT></TD></TR>"
            Print #hFile, "<TR><TD vAlign=center width=""20%"" bgColor=#000090><span class=""Info"">�����, ����</span></TD>"
            Print #hFile, "<TD vAlign=center bgColor=#000090><span class=""Info"">����� " & Str(hMsg.TopicNum) & ": " & sMsg & "</span></TD></TR>"
            TRindex = TRindex + 1
        Case 2: '�������� ������ ���������
            ''Print #hFile, "<TR ID=" & Format(RecNum) & " bgColor=#" & COLOR & " onmouseover=""NavigatorScroll(" & Format(RecNum) & ")""><TD vAlign=top width=""20%""><FONT color=#000090 face=""Verdana, Arial, Helvetica, Geneva"" size=2><center><B>" & sMsg & "</B></center></FONT><BR>"
            Print #hFile, "<TR ID=" & Format(RecNum) & " bgColor=#" & COLOR & " onmouseover=""NavigatorScroll(" & Format(RecNum) & ")"">"
            Print #hFile, "<TD vAlign=top><center><span class=""Nick""><B>" & sMsg & "</B></span><BR><BR>"
            TRindex = TRindex + 1
        Case 3: '�������� ���� ���������
            ''Print #hFile, "<FONT color=#000090 face=""Verdana, Arial, Helvetica, Geneva"" size=1><center><B>����: " & sMsg & "</B></center></FONT></TD>"
            Print #hFile, "<span class=""Date""><B>����: " & sMsg & "</B></span></center></TD>"
        Case 4: '�������� ��������� ��������� (�������� ������ ����� Mode=8)
            ''Print #hFile, sMsg & "</FONT></TD></TR>"
            Print #hFile, sMsg & "</span></TD></TR>"
        Case 5: '����� ����
            Print #hFile, "<TABLE style='text-align:justify' border=0 cellPadding=10 cellSpacing=1 width=""95%""><TBODY>"
        Case 6: '����� ���� + �����������
            Print #hFile, "</TBODY></TABLE>"
        Case 7: '����� ��������
            Print #hFile, "<hr width=85%>"
            Print #hFile, "<font size=1>"
            Print #hFile, "This page was generated by:<br>"
            'Add version information
            Print #hFile, "Elite Games Conference Extractor (EGCE)<br>v. " & App.Major & "." & App.Minor & "." & App.Revision & " "
            Print #hFile, "EGCE &copy Copyright 2001 <b>Rade</b><br>"
            Print #hFile, "Elite Games &copy 1999-2001 ������ ���������� a.k.a. Ranger"
            Print #hFile, "</font>"
            Print #hFile, "</CENTER></BODY></HTML>"
        Case 8: '�������� ��������� � ������ ������ (����� Mode=4)
            ''Print #hFile, "<TD vAlign=top><FONT color=#000090 face=""Verdana, Arial, Helvetica, Geneva"" size=2><B>" & sMsg & "</B><br>"
            Print #hFile, "<TD vAlign=top><span class=""Nick""><B>" & sMsg & "</B><br>"
        Case 9: '������ ������ �� e-mail ������
            ''Print #hFile, "<TR ID=" & Format(RecNum) & " bgColor=#" & COLOR & " onmouseover=""NavigatorScroll(" & Format(RecNum) & ")""><TD vAlign=top width=""20%""><FONT color=#000090 face=""Verdana, Arial, Helvetica, Geneva"" size=2><center><B><A href=""mailto:" & sMsg & """>"
            Print #hFile, "<TR ID=" & Format(RecNum) & " bgColor=#" & COLOR & " onmouseover=""NavigatorScroll(" & Format(RecNum) & ")"">"
            Print #hFile, "<TD vAlign=top><center><span class=""Nick""><B><A href=""mailto:" & sMsg & "?Subject=Elite Games Old Conference"">"
            TRindex = TRindex + 1
        Case 10: '������ ������ (����� ������ 9)
            ''Print #hFile, sMsg & "</A></B></center></FONT><BR>"
            Print #hFile, sMsg & "</A></B></span><BR><BR>"
    End Select
    Close hFile
    WriteHTML = True
End Function

'*****************************************************
'*      ���������� ��������� � ��������� ������      *
'*  Mode = 0 �������� ���������, ������ + �����      *
'*  Mode = 1 �������� ������ � ��������              *
'*  Mode = 2 �������� ��������� ���������            *
'*  Mode = 3 ��������� � �������� ������� ��� ������ *
'*           ����� ����                              *
'*  Mode = 4 ������� ��������� (������ ��������)     *
'*      Page    - ����� ��������                     *
'*      Color   - ���� ���� ������ ���������         *
'*      RecNum  - ����� ������ � �� ������           *
'*      Indent  - �������� ������� ���������         *
'*      sMsg    - ������������ ��������              *
'*****************************************************
Public Function WriteNavigator(ByVal Mode As Long, ByVal Page As Long, ByVal RecNum As Long, ByVal Indent As Long, ByVal sMsg As String) As Boolean
Dim hFile As Long
Static LastIndent As Long   '������ ����������� ���������. Static ���������
                            '������� ���������� ���������� ����� ����������
                            '������ ������� (���������� ������)
    WriteNavigator = False
    hFile = FreeFile
    Open (App.Path & DataPath & NavPrefix & Format(Page) & HTMLext) For Append As hFile
    Select Case Mode
        Case 0: '����� ����������, ���������� ������� + ����� ��������� (��������� LastIndent)
'            LastIndent = 0
            '��������� ������ ������ ����� 1, � �� 0.
            LastIndent = 1
            '������� ���������
            Print #hFile, "<HTML><HEAD><TITLE>��������� EGCE - �������� " & Format(Page) & "</TITLE>"
            Print #hFile, "<LINK HREF=""" & CSSFile & """ REL=stylesheet TYPE=""text/css"">"
            Print #hFile, "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; CHARSET=Windows-1251"">"
            Print #hFile, "<META http-equiv=Pragma content=no-cache>"
            Print #hFile, "<META name=GENERATOR content=""EGCE"">"
            Print #hFile, "<SCRIPT language=JavaScript src=""" & PaneFile & """></SCRIPT>"
            Print #hFile, "<SCRIPT language=JavaScript>"
            Print #hFile, "function MainScroll(tgt)"
            Print #hFile, "{"
            Print #hFile, "    var target=window.parent.frames;"
            Print #hFile, "    var coll = target(1).document.all.tags(""TR"");"
            Print #hFile, "    coll(eval(tgt-1)).scrollIntoView (true);"
            Print #hFile, "}"
            Print #hFile, "</SCRIPT>"
            Print #hFile, "</HEAD><BODY bgcolor=#F8F8F8 onclick=""DoIt()""><br>"
            '<br> ����� ��� ����� ��� ������ ������ 1 ��������� ��� ����� � ���� ����������
            Close hFile
            '������ ���������� � ��������� ��������
            hFile = FreeFile
            Open (App.Path & DataPath & IndexPrefix & Format(Page) & HTMLext) For Append As hFile
            Print #hFile, "<HTML><HEAD><TITLE>������ EGCE - �������� " & Format(Page) & "</TITLE>"
            Print #hFile, "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; CHARSET=Windows-1251"">"
            Print #hFile, "</HEAD>"
            ''Print #hFile, "<frameset bordercolor=green rows=""150,1*"">"
            ''��� �����? � ����������...
            Print #hFile, "<frameset bordercolor=green rows=""60,1*"">"
            Print #hFile, " <frame name=UpperFrame src=""" & NavPrefix & Format(Page) & HTMLext & """ style='mso-linked-frame:auto'>"
            Print #hFile, " <frame name=LowerFrame src=""" & PagesPrefix & Format(Page) & HTMLext & """ style='mso-linked-frame:auto'>"
            Print #hFile, "<noframes>"
            Print #hFile, "<BODY>"
            Print #hFile, "��� ������� �� ������������ ������. ��������� ����������. ��������"
            Print #hFile, "<a href=""" & PagesPrefix & Format(Page) & HTMLext & """>�����</a> ����� ������� ������ ���������� " & Format(Page) & "  ��������"
            Print #hFile, "</BODY>"
            Print #hFile, "</HTML>"
        Case 1: '�������� ������ � ��������
            If Indent > LastIndent Then
                '����������� ������
                Do While LastIndent < Indent
                    Print #hFile, "<DL>"
                    LastIndent = LastIndent + 1
                Loop
            Else
                '��������� ������
                Do While LastIndent > Indent
                    Print #hFile, "</DL>"
                    LastIndent = LastIndent - 1
                Loop
            End If
            ''Print #hFile, "<P ID=" & Format(RecNum) & " onmouseover=""MainScroll(" & Format(RecNum) & ")""><FONT color=#000090 face=""Verdana, Arial, Helvetica, Geneva"" size=1><B>" & sMsg & "</B>:"
            Print #hFile, "<P ID=" & Format(RecNum) & " onmouseover=""MainScroll(" & Format(TRindex) & ")""><span class=""Date""><B>" & sMsg & "</B>:"
        Case 2: '�������� ��������� ���������
            ''Print #hFile, sMsg & "</FONT><br>"
            '��������� ������ ������ <br>
            ''Print #hFile, sMsg & "</FONT>"
            Print #hFile, sMsg & "</span>"
        Case 3: '��������� � �������� ������� (���������� "������" ��� ������ ����� ����)
''            If LastIndent > 0 Then
''                '�������� ������� �� ����
''                Do While LastIndent > 0
''                    Print #hFile, "</DL>"
''                    LastIndent = LastIndent - 1
''                Loop
''            End If
            '��������� ������ ������ �� 0, � 1.
            If LastIndent > 1 Then
                '�������� ������� �� ����
                Do While LastIndent > 1
                    Print #hFile, "</DL>"
                    LastIndent = LastIndent - 1
                Loop
            End If
            '��������� ��� <DL>, ������� ����� <HR>
            Print #hFile, "</DL>"
            '������� �����
            Print #hFile, "<HR><DL>"
        Case 4: '������� ��������� (������ ��������)
            Print #hFile, "</BODY></HTML>"
    End Select
    Close hFile
    WriteNavigator = True
End Function

'******************************************************
'* ������� ��������� �������� � �����, ��������,      *
'* ���� ������� ���� �����������.                     *
'******************************************************
Public Function EraseHTML(ByVal Page As Long) As Boolean
    EraseHTML = False
    '������� �������� � �����������
    If Dir(App.Path & DataPath & PagesPrefix & Format(Page) & HTMLext) <> strEmpty Then
        Kill (App.Path & DataPath & PagesPrefix & Format(Page) & HTMLext)
        EraseHTML = True
    End If
    '������� ��������-���������
    If Dir(App.Path & DataPath & NavPrefix & Format(Page) & HTMLext) <> strEmpty Then
        Kill (App.Path & DataPath & NavPrefix & Format(Page) & HTMLext)
        EraseHTML = True
    End If
    '������� ��������� ������
    If Dir(App.Path & DataPath & IndexPrefix & Format(Page) & HTMLext) <> strEmpty Then
        Kill (App.Path & DataPath & IndexPrefix & Format(Page) & HTMLext)
        EraseHTML = True
    End If
End Function

'******************************************************
'* ��������� ���� ������ ������� ��� �������� ������� *
'* � ���������� ���������� � HTML-����                *
'*      ������������ ��������                         *
'*          blValue = 0 �������� �� ��������          *
'*          blValue = 1 ��������� �������� �������    *
'*      PageLoaded - ����� ���������� ��������        *
'* opt. UpdateFile - ����� �� ��������� �� �������    *
'******************************************************
Public Function UpdateTOC(ByVal PageLoaded As Long, Optional ByVal UpdateFile As Boolean = True) As Boolean
Dim LastRecNum As Long
Dim i As Long
Dim blValue As Long '���� Boolean, �� ���� ����� �� ����� ��������, �������� �������� :(
Dim RetValue As Double
Dim hFile As Long
Dim hFile2 As Long
    If UpdateFile Then
        '����������� ����������� VB: ��� If ������� ��� AND
        If PageLoaded > 0 Then
            '�������� ��������� ���������
            hFile = FreeFile
            '��������� ����
            Open (App.Path & DataPath & PagesFile) For Random As hFile Len = Len(blValue)
            LastRecNum = LOF(hFile) \ Len(PageLoaded)
            If PageLoaded > LastRecNum + 1 Then
                '����������� ������, �������������� �����������
                For i = LastRecNum + 1 To PageLoaded - 1
                    blValue = 0
                    Put hFile, i, blValue
                Next i
            End If
            '�������� ����������
            blValue = 1
            Put hFile, PageLoaded, blValue
            Debug.Print "UpdateTOC -> LOG: Adding record number " & PageLoaded
            Close hFile
        End If
    End If
    '�������� ������ HTML
    '��������� HTML
    hFile = FreeFile
    Open (App.Path & DataPath & TOCFile) For Output As hFile
    '��������� ������ �������
    hFile2 = FreeFile
    Open (App.Path & DataPath & PagesFile) For Random As hFile2 Len = Len(blValue)
    LastRecNum = LOF(hFile2) \ Len(blValue)
    
    Print #hFile, "<HTML><HEAD><TITLE>����������</TITLE>"
    Print #hFile, "<SCRIPT language=JavaScript>"
    Print #hFile, "var match = 1;"
    Print #hFile, "var pChange = true;"
    Print #hFile, "var pattern;"
    Print #hFile, "var Mode;"
    Print #hFile, "var Params = 0;"
    Print #hFile, "function doOK() {"
    Print #hFile, "    switch(Mode) {"
    Print #hFile, "        case 1: {"
    Print #hFile, "            var target= window.parent.frames;"
    Print #hFile, "            var targ = target(1).frames;"
    Print #hFile, "            break;"
    Print #hFile, "        }"
    Print #hFile, "        case 2: {"
    Print #hFile, "            var targ = window.parent.frames;"
    Print #hFile, "        }"
    Print #hFile, "    }"
    Print #hFile, "    Params = 0;"
    Print #hFile, "    if (document.all.item(""whole"").checked) {"
    Print #hFile, "        Params += 2;"
    Print #hFile, "    }"
    Print #hFile, "    if (document.all.item(""match"").checked) {"
    Print #hFile, "        Params += 4;"
    Print #hFile, "    }"
    Print #hFile, "    var rng = targ(1).document.body.createTextRange();"
    Print #hFile, "    if (pChange) {"
    Print #hFile, "        pChange = false;"
    Print #hFile, "        pattern = MySearch.value;"
    Print #hFile, "    }"
    Print #hFile, "    else {"
    Print #hFile, "        var i = 1;"
    Print #hFile, "        while (match > i) {"
    Print #hFile, "            rng.findText(pattern, 9999999, Params);"
    Print #hFile, "            rng.collapse(false);"
    Print #hFile, "            i++;"
    Print #hFile, "        }"
    Print #hFile, "    }"
    Print #hFile, "    if (rng.findText(pattern, 9999999, Params)==true) {"
    Print #hFile, "        rng.select();"
    Print #hFile, "        rng.collapse(false);"
    Print #hFile, "        rng.scrollIntoView();"
    Print #hFile, "        match++;"
    Print #hFile, "    }"
    Print #hFile, "    else{"
    Print #hFile, "        alert(""����� ��������. ����� ���� ���������� ����������: "" + eval(match-1));"
    Print #hFile, "    }"
    Print #hFile, "}"
    Print #hFile, "function PatternChanged() {"
    Print #hFile, "    match = 1;"
    Print #hFile, "    pChange = true;"
    Print #hFile, "}"
    Print #hFile, "function HideIt() {"
    Print #hFile, "    document.all.SearchDiv.style.visibility = ""Hidden"";"
    Print #hFile, "}"
    Print #hFile, "function ShowIt() {"
    Print #hFile, "    Mode = 1;"
    Print #hFile, "    document.all.SearchDiv.style.visibility = ""Visible"";"
    Print #hFile, "}"
    Print #hFile, "function ShowItA() {"
    Print #hFile, "    Mode = 2;"
    Print #hFile, "    document.all.SearchDiv.style.visibility = ""Visible"";"
    Print #hFile, "}"
    
    
    Print #hFile, "</SCRIPT></HEAD>"
    Print #hFile, "<BODY alink=#000088 vlink=#008800 link=#008800 bgcolor=#F8F8F8>"
    Print #hFile, "<center><h2><u>EGCE</u></h2><h3>����������</h3><br>"
    Print #hFile, "<a href=""" & StartupFile & """ target=""MainFrame"" onclick=""HideIt();return true"">��������� ��������</a><br>"
    Print #hFile, "<a href=""" & FAQFile & """ target=""MainFrame"" onclick=""HideIt();return true"">��� � ����</a><br>"
    'TO DO �������� ������ �� �������
    For i = 1 To LastRecNum
        Get hFile2, i, blValue
        If blValue = 1 Then
            '��� �������� ��� ��������
            '������ ������ ����������� ID ��� ���� ����� �� ����� ���� ������� �������� TAB...
            Print #hFile, "<a ID=Anchor" & Format(i) & " href=""" & "pIndex" & Format(i) & HTMLext & """ target=""MainFrame"" onMouseOver=""window.status=''; return true""  onMouseOut=""window.status=''; return false"" onclick=""ShowIt();return true"">" & Format(i) & "</a>"
        Else
            Print #hFile, Chr(32) & Format(i) & Chr(32)
        End If
        '�� 10 ������ � ������
        If i \ 10 = i / 10 Then
            Print #hFile, "<br>"
        End If
    Next i
    Print #hFile, "<br><br>������:<br>"
    Print #hFile, "<a href=""" & ANamesPage & """ target=""MainFrame"" onclick=""ShowItA();return true"">�� ������</a> | "
    Print #hFile, "<a href=""" & ATimePage & """ target=""MainFrame"" onclick=""ShowItA();return true"">�� ""�����""</a><br>"
    Print #hFile, "<a href=""" & ANumsPage & """ target=""MainFrame"" onclick=""ShowItA();return true"">�� ����� ���������</a><br>"
    Print #hFile, "<a href=""" & ASizePage & """ target=""MainFrame"" onclick=""ShowItA();return true"">�� ������� ���������</a>"
    Print #hFile, "<br><br><a href="""" target=""MainFrame"" onclick=""HideIt();return true"">�������� �����</a>"
    Print #hFile, "<br><br>"
    Print #hFile, "<DIV ID=SearchDiv STYLE=""visibility: Hidden"">"
    Print #hFile, "<B>����� �� ��������: </B><INPUT ID=MySearch TYPE=text onfocus=""PatternChanged()"">"
    Print #hFile, "<BUTTON onclick=""doOK()"">OK</BUTTON><br>"
    Print #hFile, "<INPUT ID=match TYPE=CHECKBOX VALUE=4 UNCHECKED>�������"
    Print #hFile, "<INPUT ID=whole TYPE=CHECKBOX VALUE=2 UNCHECKED>�����"
    Print #hFile, "</DIV><br>"
    Print #hFile, "<font size=1>"
    Print #hFile, "������������� ����������: 1024x768<br>"
    Print #hFile, "������������� �������: IE 4.01 ��� ����<br>"
    'Add version information
    Print #hFile, "Elite Games Conference Extractor (EGCE)<br>v. " & App.Major & "." & App.Minor & "." & App.Revision & "<br>"
    Print #hFile, "EGCE &copy Copyright 2001 <b>Rade</b><br>"
    Print #hFile, "Elite Games &copy 1999-2001 ������ ���������� a.k.a. Ranger"
    Print #hFile, "</font></center>"
    Print #hFile, "</BODY></HTML>"
    Close hFile2
    Close hFile
    
    ''TO DO ������������ ����� ������ � �������� ���������� � �������� �� � TOC
    '(������� ��� � ������ �������)
    
    '��������� ������ �, ���� ����, ��������!! �������� ������ �� Close hFile
    ''TO DO �������� �� ��������!!!
    Call UpdateIndex(True)
    '������� HTML
    Call ShowHTML
End Function

'********************************************
'* �������� ����� �� ����� �������          *
'* ������� �������������� ��� ���� �����    *
'* ��� ����� ����� ������������� ��������   *
'* ����� ����� �� ������� �� ������� ������ *
'********************************************
Sub LoadResStrings(frm As Form)
Dim ctl As Control
Dim obj As Object
Dim sCtlType As String
Dim nVal As Integer
    On Error Resume Next

    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    
'��������!!!!!!!!!!!!!!!!!!!!
Exit Sub
'��������!!!!!!!!!!!!!!!!!!!!

    For Each ctl In frm.Controls
        'Set ctl.Font = fnt
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then
            ctl.Caption = LoadResString(CInt(ctl.Tag))
        ElseIf sCtlType = "Menu" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                If Not obj.Tag = strEmpty Then
                    'Tag <> 0  =>��������� �������� �� ����� ��������'
                    obj.Caption = LoadResString(CInt(obj.Tag))
                End If
                ''obj.Caption = LoadResString(CInt(obj.Tag))
                ''obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                If Not (obj.ToolTipText = 0) Then
                    obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
                End If
            Next
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                obj.Text = LoadResString(CInt(obj.Tag))
            Next
            For Each obj In ctl.Panels
                If Not (CInt(obj.Tag)) = 0 Then
                    'Tag <> 0  =>��������� �������� �� ����� ��������'
                    obj.Text = LoadResString(CInt(obj.Tag))
                End If
            Next
        ElseIf sCtlType = "CommonDialog" Then
            nVal = 0
        ElseIf sCtlType = "Line" Then
            nVal = 0
        ElseIf sCtlType = "Timer" Then
            nVal = 0
        ElseIf sCtlType = "ImageList" Then
            nVal = 0
        End If
    Next

End Sub

'**************************************************
'* ������ ������ �� ������, ������� ������        *
'* ���������� ������ �� ��������                  *
'**************************************************
Public Function ClearGarbage(ByVal strString As String) As String
Dim tmpString As String
    tmpString = Replace(strString, "&nbsp;", Chr(32), 1, -1, vbBinaryCompare)
    tmpString = Replace(tmpString, Chr(10), strEmpty, 1, -1, vbBinaryCompare)
    tmpString = Replace(tmpString, Chr(13), strEmpty, 1, -1, vbBinaryCompare)
    ClearGarbage = tmpString
End Function

'***********************************************
'* ����� ������� ������ ��� ������ ����������� *
'* ForceOverwrite = TRUE ��������� ��          *
'*      ������������� ��������������           *
'*      ����������                             *
'***********************************************
Sub UpdateIndex(ByVal ForceOverwrite As Boolean)
    '��� ������ �������� ��������� ��������
    If Dir(App.Path & DataPath & StartupFile) <> "" Then
        If ForceOverwrite Then
            Kill (App.Path & DataPath & StartupFile)
        End If
    End If
    hFile = FreeFile
    Open (App.Path & DataPath & StartupFile) For Output As hFile
    Print #hFile, "<HTML><HEAD><TITLE>��������� �������� EliteGames Conference Extractor</TITLE>"
    Print #hFile, "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; CHARSET=Windows-1251"">"
    Print #hFile, "<META http-equiv=Pragma content=no-cache>"
    Print #hFile, "<META name=GENERATOR content=""EGCE""></HEAD>"
    Print #hFile, "<BODY alink=#000088 vlink=#008800 link=#008800><CENTER>"
    Print #hFile, "<h3>����� ������ ����������� Elite Games</H3>"
    Print #hFile, "<h4>�������� ������ ����� ��� ������������ ��������������� ��������<br>"
    Print #hFile, "��� ���� �� ������������� ��� ��������� ����� ��������� ����������<br>� ���������:</H4>"
    Print #hFile, "<a href=""http://www.elite-russia.net/"">Elite Games</A><br>"
    Print #hFile, "<a href=""http://book.by.ru/cgi-bin///book.cgi?book=Elitegames"">������ ����������� Elite Games</A><br>"
    Print #hFile, "<a href=""http://x-dron.narod.ru/"">���� X-Dron'�</A>"
    Print #hFile, "</CENTER></BODY></HTML>"
    Close hFile
    
    '������ �������� ��������
    If Dir(App.Path & DataPath & IndexFile) <> "" Then
        If ForceOverwrite Then
            Kill (App.Path & DataPath & IndexFile)
        End If
    End If
    hFile = FreeFile
    Open (App.Path & DataPath & IndexFile) For Output As hFile
    Print #hFile, "<HTML><HEAD><TITLE>EliteGames Conference Extractor database</TITLE></HEAD>"
    Print #hFile, "<FRAMESET Cols="" 250, 1 *"" BORDERCOLOR=""royalblue"" frameborder=""no"" border=0 framespacing=""0"" marginwidth=0 marginheight=0>"
    Print #hFile, "<frame name=LeftFrame src=""pagesTOC.html"" style='mso-linked-frame:auto'>"
    Print #hFile, "<frame name=MainFrame src=""Startup.html"" style='mso-linked-frame:auto'>"
    Print #hFile, "<noframes><BODY>"
    Print #hFile, "��� ������� �� ������������ ������. �� ������� ������ ����������.<BR><BR>"
    Print #hFile, "<a href=""pagesTOC.html"">����������</A>"
    Print #hFile, "</BODY></HTML>"
    Close hFile
End Sub

'********************************************
'* ���������� ������ � ����������, ���� ��  *
'* ����� �������� � ����� ��������          *
'********************************************
Public Function ShowMsgBox(ByVal resMsg As Long, ByVal MsgFlags As Long, ByVal resCaption As Long, Optional ByVal msgboxHelp As Long = -6) As VbMsgBoxResult
Dim strHelp As String
Dim strCaption As String
    strHelp = LoadResString(resMsg)
    strCaption = LoadResString(resCaption)
    If msgboxHelp = -6 Then
        '������ ����������-��������� ������ �� ������ - ������ �� �����!
        '�������� ��������� - ���� ��� ����� ���������!
        ShowMsgBox = MsgBox(strHelp, MsgFlags, strCaption)
    Else
        '�� ��� �, ������� ������ :))
        'TO DO ��� ��� � ������ ����� �������? ��������!!!
        ShowMsgBox = MsgBox(strHelp, MsgFlags, strCaption, , msgboxHelp)
    End If
End Function

'**********************************************
'* ������� ������� � ������ � � ����� ������  *
'**********************************************
Public Function RemoveSpaces(ByVal strString As String) As String
Dim tmpString As String
    tmpString = strString
    Do While (Left(tmpString, 1) = Chr(32))
        tmpString = Right(tmpString, Len(tmpString) - 1)
    Loop
    Do While (Right(tmpString, 1) = Chr(32))
        tmpString = Left(tmpString, Len(tmpString) - 1)
    Loop
    RemoveSpaces = tmpString
End Function

'**********************************************
'*      ���������� ������� ������ CSS         *
'**********************************************
Public Sub WriteCSS()
Dim hFile As Long

    If Dir(App.Path & DataPath & CSSFile) <> strEmpty Then
        Kill (App.Path & DataPath & CSSFile)
    End If
    hFile = FreeFile
    Open (App.Path & DataPath & CSSFile) For Output As hFile

    Print #hFile, "BODY"
    Print #hFile, "{"
        Print #hFile, "BORDER-BOTTOM: medium none;"
        Print #hFile, "BORDER-LEFT: medium none;"
        Print #hFile, "BORDER-RIGHT: medium none;"
        Print #hFile, "BORDER-TOP: medium none;"
        Print #hFile, "COLOR: black;"
        Print #hFile, "FONT-FAMILY: Verdana, Arial;"
        Print #hFile, "FONT-SIZE: x-small;"
        Print #hFile, "FONT-WEIGHT: normal;"
        Print #hFile, "TEXT-ALIGN: justify;"
        Print #hFile, "TEXT-DECORATION: none;"
        Print #hFile, "WIDTH: 90%"
    Print #hFile, "}"
    Print #hFile, "A"
    Print #hFile, "{"
        Print #hFile, "COLOR: #006600;"
        Print #hFile, "FONT-FAMILY: Verdana, Arial;"
        Print #hFile, "TEXT-DECORATION: none"
    Print #hFile, "}"
    Print #hFile, "A: active"
    Print #hFile, "{"
        Print #hFile, "COLOR: #006600;"
        Print #hFile, "FONT-FAMILY: Verdana, Arial;"
        Print #hFile, "TEXT-DECORATION: none"
    Print #hFile, "}"
    Print #hFile, "A: link"
    Print #hFile, "{"
        Print #hFile, "COLOR: #006600;"
        Print #hFile, "FONT-FAMILY: Verdana, Arial;"
        Print #hFile, "TEXT-DECORATION: none"
    Print #hFile, "}"
    Print #hFile, "A: hover"
    Print #hFile, "{"
        Print #hFile, "COLOR: #00AA00;"
        Print #hFile, "FONT-FAMILY: Verdana, Arial;"
        Print #hFile, "TEXT-DECORATION: underline"
    Print #hFile, "}"
    Print #hFile, "A: visited"
    Print #hFile, "{"
        Print #hFile, "COLOR: #006600;"
        Print #hFile, "FONT-FAMILY: Verdana, Arial;"
        Print #hFile, "TEXT-DECORATION: none"
    Print #hFile, "}"
    Print #hFile, "TD"
    Print #hFile, "{"
        Print #hFile, "COLOR: black;"
        Print #hFile, "FONT-SIZE: x-small;"
        Print #hFile, "FONT-WEIGHT: normal;"
        Print #hFile, "TEXT-DECORATION: none"
    Print #hFile, "}"
    Print #hFile, ".Info"
    Print #hFile, "{"
        Print #hFile, "COLOR: white;"
        Print #hFile, "FONT-FAMILY: Verdana, Arial;"
        Print #hFile, "FONT-SIZE: xx-small;"
        Print #hFile, "FONT-WEIGHT: bold;"
        Print #hFile, "TEXT-ALIGN: right"
    Print #hFile, "}"
    Print #hFile, ".Nick"
    Print #hFile, "{"
        Print #hFile, "COLOR: #000090;"
        Print #hFile, "FONT-FAMILY: Verdana, Arial;"
        Print #hFile, "FONT-SIZE: x-small;"
        Print #hFile, "FONT-WEIGHT: normal"
    Print #hFile, "}"
    Print #hFile, ".Date"
    Print #hFile, "{"
        Print #hFile, "COLOR: #000090;"
        Print #hFile, "FONT-FAMILY: Verdana, Arial;"
        Print #hFile, "FONT-SIZE: xx-small;"
        Print #hFile, "FONT-WEIGHT: bold"
    Print #hFile, "}"
    
    Close hFile
End Sub

'**********************************************
'*    ���������� ������ ���������� �������    *
'*    ����������                              *
'**********************************************
Public Sub WritePane()
Dim hFile As Long

    If Dir(App.Path & DataPath & PaneFile) <> strEmpty Then
        Kill (App.Path & DataPath & PaneFile)
    End If
    hFile = FreeFile
    Open (App.Path & DataPath & PaneFile) For Output As hFile

Print #hFile, "var TimerID;            //Timer handler"
Print #hFile, "var Processing = false; //Processing flag"
Print #hFile, "var LastState = 1;      //1 - Pane was lifted up, 0 - Pane was downed"
Print #hFile, "var CurrSize = 60;      //Current frame size"
Print #hFile, "var Step;               //Step for frame border move"
Print #hFile, "var Direction = 1;      //Where to move frame border"

Print #hFile, "function MovePane(FinalSize, Mode)"
Print #hFile, "{"
Print #hFile, "    var tgt = window.parent.document.all(4);        //4 - Frameset tag"
Print #hFile, "    switch (Mode)"
Print #hFile, "    {"
Print #hFile, "        case 0: //Down"
Print #hFile, "        {"
Print #hFile, "            if ((Step < 8) && (CurrSize < FinalSize - 50))  //Increase Step up to 8 pixel per timer tick"
Print #hFile, "            {"
Print #hFile, "                //Step <<= 1;                               //Multiply Step by 2"
Print #hFile, "                ++Step;"
Print #hFile, "            }"
Print #hFile, "            if ((Step > 1) && (CurrSize > FinalSize - 150)) //Desrease Step down to 1 pixel per timer tick"
Print #hFile, "            {"
Print #hFile, "                //Step >>= 1;                               //Division Step by 2"
Print #hFile, "                --Step;"
Print #hFile, "            }"
Print #hFile, "            if ((Step == 1) && (CurrSize >= FinalSize)) //End of moving - reset variables"
Print #hFile, "            {"
Print #hFile, "                Step = 1;                                   //Reset Step Value"
Print #hFile, "                window.clearInterval(TimerID);              //Stop timer"
Print #hFile, "                Processing = false;"
Print #hFile, "                LastState = 0;"
Print #hFile, "            }"
Print #hFile, "            break;"
Print #hFile, "        }"
Print #hFile, "        case 1: //Up"
Print #hFile, "        {"
Print #hFile, "            if ((Step < 8) && (CurrSize > FrameSize + 50))  //Increase Step up to 8 pixel per timer tick"
Print #hFile, "            {"
Print #hFile, "                //Step <<= 1;                               //Multiply Step by 2"
Print #hFile, "                ++Step;"
Print #hFile, "            }"
Print #hFile, "            if ((Step > 1) && (CurrSize < FinalSize + 150)) //Desrease Step down to 1 pixel per timer tick"
Print #hFile, "            {"
Print #hFile, "                //Step >>= 1;                               //Division Step by 2"
Print #hFile, "                --Step;"
Print #hFile, "            }"
Print #hFile, "            if ((Step == 1) && (CurrSize <= FinalSize)) //End of moving - reset variables"
Print #hFile, "            {"
Print #hFile, "                Step = 1;                                   //Reset Step Value"
Print #hFile, "                window.clearInterval(TimerID);              //Stop timer"
Print #hFile, "                Processing = false;"
Print #hFile, "                LastState = 1;"
Print #hFile, "            }"
Print #hFile, "        }"
Print #hFile, "    }"
Print #hFile, "    CurrSize += Step * Direction;"
Print #hFile, "    tgt.rows = CurrSize + "",1*"";"
Print #hFile, "}"

Print #hFile, "function DoIt()"
Print #hFile, "{"
Print #hFile, "    var tgt = window.parent.document.all(4);        //4 - Frameset tag"
Print #hFile, "    if (!(Processing))"
Print #hFile, "    {"
Print #hFile, "        var strTmp = tgt.rows;                      //Get frame size and properties text"
Print #hFile, "        var pos = strTmp.indexOf("","");            //Search for "","" position"
Print #hFile, "        CurrSize = strTmp.substr(0, pos);           //Get frame size only"
Print #hFile, "        CurrSize = parseInt(CurrSize);"
Print #hFile, "        if (CurrSize < 250)"
Print #hFile, "        {"
Print #hFile, "            LastState = 1;  //Lifted up"
Print #hFile, "        }"
Print #hFile, "        else"
Print #hFile, "        {"
Print #hFile, "            LastState = 0;  //Downed"
Print #hFile, "        }"
'Print #hFile, "        alert(CurrSize);"
Print #hFile, "        Processing = true;"
Print #hFile, "        Step = 0;"
Print #hFile, "        switch (LastState)"
Print #hFile, "        {"
Print #hFile, "            case 0: //Downed"
Print #hFile, "            {"
Print #hFile, "                Direction = -1;"
Print #hFile, "                TimerID=window.setInterval(""MovePane(60, 1)"", 5);"
Print #hFile, "                break;"
Print #hFile, "            }"
Print #hFile, "            case 1: //Lifted up"
Print #hFile, "            {"
Print #hFile, "                Direction = 1;"
Print #hFile, "                TimerID=window.setInterval(""MovePane(500, 0)"", 5);"
Print #hFile, "            }"
Print #hFile, "        }"
Print #hFile, "    }"
Print #hFile, "}"
    
    Close hFile
End Sub

