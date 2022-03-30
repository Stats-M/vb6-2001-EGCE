Attribute VB_Name = "Module1"
Option Explicit

'Установите значение FullLogging = 1 для включения записи e-mail авторов в HTML-файл
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

Public Const AuthorsFile = "authors.db"     'Имя базы данных по авторам
Public Const MsgFile = "msg.db"             'Имя базы данных ссылок на сообщения
Public Const HistoryFile = "history.db"     'Имя журнала
Public Const PagesFile = "pages.db"         'Имя базы данных по страницам
Public Const IndexFile = "index.html"       'Главный индекс
Public Const TOCFile = "pagesTOC.html"      'Основное оглавление
Public Const StartupFile = "Startup.html"   'Стартовая страница
Public Const FAQFile = "FAQ.html"           'ЧаВо
Public Const PagesPrefix = "page"           'Префикс имени страницы сообщений
Public Const NavPrefix = "pTOC"             'Префикс имени страницы-навигатора
Public Const IndexPrefix = "pIndex"         'Префикс имени индекса страницы сообщений
Public Const HTMLext = ".html"              'Расширение имен страниц
Public Const ANamesPage = "authors1.html"   'Авторы по алфавиту
Public Const ANumsPage = "authors2.html"    'Авторы по кол-ву сообщений
Public Const ASizePage = "authors3.html"    'Авторы по размеру сообщений
Public Const ATimePage = "authors4.html"    'Авторы по времени появления в конфе
Public Const HelpFile = "\EGCEhelp.chm"     'Имя файла справки
Public Const CSSFile = "EGCE.css"           'Таблица стилей HTML
Public Const PaneFile = "pane.js"           'Таблица стилей HTML
Public Const DataPath = "\EGCEdata\"        'Путь к файлам данных
Public Const strEmpty = ""
Public Const ColorLight = "ffffff"          'Константы цвета для HTML-страницы
Public Const ColorDark = "ddddff"
Public Const LastPageAvailable% = 80        'Номер последней страницы старой конфы
Public Const iMaxSize% = 255                'Максимальный размер буфера
Public Const strSepURLDir = "/"             'Разделитель URL-адресов
Public Const strSepDir = "\"                'Разделитель директорий
Public Const strHHelpEXEname = "hh.exe"     'Имя программы просмотра справки *.CHM
Public Const strExplorer = "explorer.exe"
Public Const strFullLogging = "full"

''HKEY_LOCAL_MACHINE\Software\CLASSES\chm.file\shell\open\command

Public PauseTime As Variant, Start As Variant    'Для вычисления таймаутов
Public fBrw As frmBrowser           'Собственно окно программы
Public Result As VbMsgBoxResult     'Для окошек сообщений
Public bResult As Boolean           'Для результатов функций
Public WorkOffline As Boolean       'Работаем в офф-лайн по-умолчанию
Public hMsg As tMessage             'Записи для качалок
Public MsgMap As cMsgMap            'Класс для работы с базой данных сообщений
Public Authors As cAuthors          'Класс для работы с базой данных авторов
Public hFile As Long                'Указатель для работы с файлами
Public CurrPage As Long             'Номер закачиваемой страницы
Public RecNum As Long               'Номер записи в БД ссылок (для обратной связи)
Public BaseHREF As String           'Базовый адрес конференции
Public BaseHREFprefix As String     'Префикс базового адреса конференции
Public TRindex As Long              'Индекс текущего тега TR (для корректной навигации)
Public bFullLogging As Boolean

'Внешние DLL-функции
Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'********************************************
'* Получение пути установки Windows через   *
'* Win API                                  *
'* Полученный путь содержит закрывающий     *
'* разделитель директорий \                 *
'********************************************
Function GetWindowsDir() As String
Dim strBuf As String
Dim iZeroPos As Integer

    'Заполняем буфер пробелами
    strBuf = Space(iMaxSize)
    If GetWindowsDirectory(strBuf, iMaxSize) > 0 Then
        'Ищем терминатор строки
        iZeroPos = InStr(strBuf, Chr$(0))
        'Если терминатор есть, то удаляем его
        If iZeroPos > 0 Then
            strBuf = Left$(strBuf, iZeroPos - 1)
        End If
        'Если на конце строки нет разделителя директорий, добавляем его
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
'* Запуск справочной системы Windows (формат справки *.CHM) *
'* Поиск файла hh.exe через реестр производиться НЕ будет,  *
'* положимся на то, что этот файл в большинстве случаев     *
'* лежит в папке Windows                                    *
'************************************************************
Public Sub ShowCHMHelp()
Dim RetValue As Double
    'Получить путь к папке Windows через DLL call
    RetValue = Shell(GetWindowsDir & strHHelpEXEname & Chr(32) & App.Path & HelpFile, vbMaximizedFocus)
End Sub

'************************************************************
'* Показывает HTML страницу с оглавлением всех закачанных   *
'* адресов.                                                 *
'* Загрузка страницы осуществляется путем передачи ее имени *
'* Проводнику в качестве параметра                          *
'************************************************************
Public Sub ShowHTML()
Dim RetValue As Double
    'Показать страницу
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
    'Пусть хоть пару секунд на заставку посмотрят... :))
    
    'Пока смотрят заставку, грузим главное окошко
    Set fBrw = New frmBrowser
    'Инициализация классов (приехало из frmBrowser_Load)
    Set MsgMap = New cMsgMap
    Set Authors = New cAuthors
    
    'Включение записи лога программы
    Call StartLogging
    
    'Инициализация некоторых переменных
    RecNum = 0
    CurrPage = 1
    TRindex = 0
    bFullLogging = GetCommandLine   'Посмотреть что там с параметрами командной строки
    
    If Dir(App.Path & "\EGCEdata\", vbDirectory) = strEmpty Then
        MkDir App.Path & "\EGCEdata"
    End If
    'Если таблицы стилей еще нет, создаем ее
    If Dir(App.Path & DataPath & CSSFile) = strEmpty Then
        WriteCSS
    End If
    'Если скрипта управления панелью еще нет, создим его
    If Dir(App.Path & DataPath & PaneFile) = strEmpty Then
        WritePane
    End If
    'Указываем имя файла справки
    App.HelpFile = App.Path & HelpFile
    
    WorkOffline = True  'Без необходимости никуда не лезем
    'TO DO пока так, потом посмотрим...
    BaseHREFprefix = "http://book.by.ru/cgi-bin//"
    
    Load fBrw
    Unload frmSplash

    fBrw.Show
End Sub

'*****************************************************
'*      Записываем откопанные сведения в HTML        *
'*  Mode = 0 записать начало страницы                *
'*  Mode = 1 записать заголовок сообщения            *
'*           (при старте новой ветки)                *
'*  Mode = 2 записать автора сообщения               *
'*  Mode = 3 записать дату сообщения                 *
'*  Mode = 4 записать очередное сообщение            *
'*           (вызывать только после Mode=8)          *
'*  Mode = 5 новая тема                              *
'*  Mode = 6 конец темы                              *
'*  Mode = 7 конец страницы                          *
'*  Mode = 8 записать заголовок в тексте ответа      *
'*           (перед Mode=4)                          *
'*  Mode = 9 запись ссылки на e-mail автора          *
'*  Mode =10 запись автора (после вызова 9)          *
'*      Page    - номер страницы                     *
'*      Color   - цвет фона текста сообщения         *
'*      sMsg    - записываемый параметр              *
'*      RecNum  - номер записи в БД ссылок           *
'*****************************************************
Public Function WriteHTML(ByVal Mode As Long, ByVal Page As Long, ByVal COLOR As String, ByVal sMsg As String, ByVal RecNum As Long) As Boolean
    WriteHTML = False
    hFile = FreeFile
    Open (App.Path & DataPath & PagesPrefix & Format(Page) & HTMLext) For Append As hFile
    Select Case Mode
        Case 0: 'записать начало страницы
            Print #hFile, "<HTML><HEAD><TITLE>EGCE - Страница " & Format(Page) & " - Конференция Elite Games на WWW.BOOK.BY.RU</TITLE>"
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
        Case 1: 'записать заголовок сообщения (при старте новой ветки)
            ''Print #hFile, "<TR><TD vAlign=center width=""20%"" bgColor=#000090><FONT color=#FFFFFF face=""Verdana, Arial, Helvetica, Geneva"" size=1><B>Автор, дата</B></FONT></TD>"
            ''Print #hFile, "<TD vAlign=center bgColor=#000090><FONT color=#FFFFFF face=""Verdana, Arial, Helvetica, Geneva"" size=1><B>Нитка " & Str(hMsg.TopicNum) & ": " & sMsg & "</B></FONT></TD></TR>"
            Print #hFile, "<TR><TD vAlign=center width=""20%"" bgColor=#000090><span class=""Info"">Автор, дата</span></TD>"
            Print #hFile, "<TD vAlign=center bgColor=#000090><span class=""Info"">Нитка " & Str(hMsg.TopicNum) & ": " & sMsg & "</span></TD></TR>"
            TRindex = TRindex + 1
        Case 2: 'записать автора сообщения
            ''Print #hFile, "<TR ID=" & Format(RecNum) & " bgColor=#" & COLOR & " onmouseover=""NavigatorScroll(" & Format(RecNum) & ")""><TD vAlign=top width=""20%""><FONT color=#000090 face=""Verdana, Arial, Helvetica, Geneva"" size=2><center><B>" & sMsg & "</B></center></FONT><BR>"
            Print #hFile, "<TR ID=" & Format(RecNum) & " bgColor=#" & COLOR & " onmouseover=""NavigatorScroll(" & Format(RecNum) & ")"">"
            Print #hFile, "<TD vAlign=top><center><span class=""Nick""><B>" & sMsg & "</B></span><BR><BR>"
            TRindex = TRindex + 1
        Case 3: 'записать дату сообщения
            ''Print #hFile, "<FONT color=#000090 face=""Verdana, Arial, Helvetica, Geneva"" size=1><center><B>Дата: " & sMsg & "</B></center></FONT></TD>"
            Print #hFile, "<span class=""Date""><B>Дата: " & sMsg & "</B></span></center></TD>"
        Case 4: 'записать очередное сообщение (вызывать только после Mode=8)
            ''Print #hFile, sMsg & "</FONT></TD></TR>"
            Print #hFile, sMsg & "</span></TD></TR>"
        Case 5: 'новая тема
            Print #hFile, "<TABLE style='text-align:justify' border=0 cellPadding=10 cellSpacing=1 width=""95%""><TBODY>"
        Case 6: 'конец темы + РАЗДЕЛИТЕЛЬ
            Print #hFile, "</TBODY></TABLE>"
        Case 7: 'конец страницы
            Print #hFile, "<hr width=85%>"
            Print #hFile, "<font size=1>"
            Print #hFile, "This page was generated by:<br>"
            'Add version information
            Print #hFile, "Elite Games Conference Extractor (EGCE)<br>v. " & App.Major & "." & App.Minor & "." & App.Revision & " "
            Print #hFile, "EGCE &copy Copyright 2001 <b>Rade</b><br>"
            Print #hFile, "Elite Games &copy 1999-2001 Сергей Петровичев a.k.a. Ranger"
            Print #hFile, "</font>"
            Print #hFile, "</CENTER></BODY></HTML>"
        Case 8: 'записать заголовок в тексте ответа (перед Mode=4)
            ''Print #hFile, "<TD vAlign=top><FONT color=#000090 face=""Verdana, Arial, Helvetica, Geneva"" size=2><B>" & sMsg & "</B><br>"
            Print #hFile, "<TD vAlign=top><span class=""Nick""><B>" & sMsg & "</B><br>"
        Case 9: 'запись ссылки на e-mail автора
            ''Print #hFile, "<TR ID=" & Format(RecNum) & " bgColor=#" & COLOR & " onmouseover=""NavigatorScroll(" & Format(RecNum) & ")""><TD vAlign=top width=""20%""><FONT color=#000090 face=""Verdana, Arial, Helvetica, Geneva"" size=2><center><B><A href=""mailto:" & sMsg & """>"
            Print #hFile, "<TR ID=" & Format(RecNum) & " bgColor=#" & COLOR & " onmouseover=""NavigatorScroll(" & Format(RecNum) & ")"">"
            Print #hFile, "<TD vAlign=top><center><span class=""Nick""><B><A href=""mailto:" & sMsg & "?Subject=Elite Games Old Conference"">"
            TRindex = TRindex + 1
        Case 10: 'запись автора (после вызова 9)
            ''Print #hFile, sMsg & "</A></B></center></FONT><BR>"
            Print #hFile, sMsg & "</A></B></span><BR><BR>"
    End Select
    Close hFile
    WriteHTML = True
End Function

'*****************************************************
'*      Записывает навигатор и локальный индекс      *
'*  Mode = 0 записать навигатор, индекс + сброс      *
'*  Mode = 1 Записать автора с отступом              *
'*  Mode = 2 Записать заголовок сообщения            *
'*  Mode = 3 Вернуться к нулевому отступу при старте *
'*           новой темы                              *
'*  Mode = 4 Закрыть навигатор (запись окончена)     *
'*      Page    - номер страницы                     *
'*      Color   - цвет фона текста сообщения         *
'*      RecNum  - номер записи в БД ссылок           *
'*      Indent  - величина отступа заголовка         *
'*      sMsg    - записываемый параметр              *
'*****************************************************
Public Function WriteNavigator(ByVal Mode As Long, ByVal Page As Long, ByVal RecNum As Long, ByVal Indent As Long, ByVal sMsg As String) As Boolean
Dim hFile As Long
Static LastIndent As Long   'Отступ предыдущего сообщения. Static запрещает
                            'системе уничтожать переменную после завершения
                            'работы функции (блокировка памяти)
    WriteNavigator = False
    hFile = FreeFile
    Open (App.Path & DataPath & NavPrefix & Format(Page) & HTMLext) For Append As hFile
    Select Case Mode
        Case 0: 'Старт навигатора, локального индекса + сброс состояния (обнуление LastIndent)
'            LastIndent = 0
            'Начальный отступ теперь равен 1, а не 0.
            LastIndent = 1
            'Сначала навигатор
            Print #hFile, "<HTML><HEAD><TITLE>Навигатор EGCE - страница " & Format(Page) & "</TITLE>"
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
            '<br> нужен был чтобы при старте только 1 заголовок был виден в окне навигатора
            Close hFile
            'Теперь разберемся с локальным индексом
            hFile = FreeFile
            Open (App.Path & DataPath & IndexPrefix & Format(Page) & HTMLext) For Append As hFile
            Print #hFile, "<HTML><HEAD><TITLE>Индекс EGCE - страница " & Format(Page) & "</TITLE>"
            Print #hFile, "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; CHARSET=Windows-1251"">"
            Print #hFile, "</HEAD>"
            ''Print #hFile, "<frameset bordercolor=green rows=""150,1*"">"
            ''Так лучше? И эстетичнее...
            Print #hFile, "<frameset bordercolor=green rows=""60,1*"">"
            Print #hFile, " <frame name=UpperFrame src=""" & NavPrefix & Format(Page) & HTMLext & """ style='mso-linked-frame:auto'>"
            Print #hFile, " <frame name=LowerFrame src=""" & PagesPrefix & Format(Page) & HTMLext & """ style='mso-linked-frame:auto'>"
            Print #hFile, "<noframes>"
            Print #hFile, "<BODY>"
            Print #hFile, "Ваш браузер не поддерживает фреймы. Навигатор недоступен. Щелкните"
            Print #hFile, "<a href=""" & PagesPrefix & Format(Page) & HTMLext & """>здесь</a> чтобы увидеть только содержание " & Format(Page) & "  страницы"
            Print #hFile, "</BODY>"
            Print #hFile, "</HTML>"
        Case 1: 'Записать автора с отступом
            If Indent > LastIndent Then
                'Увеличиваем отступ
                Do While LastIndent < Indent
                    Print #hFile, "<DL>"
                    LastIndent = LastIndent + 1
                Loop
            Else
                'Уменьшаем отступ
                Do While LastIndent > Indent
                    Print #hFile, "</DL>"
                    LastIndent = LastIndent - 1
                Loop
            End If
            ''Print #hFile, "<P ID=" & Format(RecNum) & " onmouseover=""MainScroll(" & Format(RecNum) & ")""><FONT color=#000090 face=""Verdana, Arial, Helvetica, Geneva"" size=1><B>" & sMsg & "</B>:"
            Print #hFile, "<P ID=" & Format(RecNum) & " onmouseover=""MainScroll(" & Format(TRindex) & ")""><span class=""Date""><B>" & sMsg & "</B>:"
        Case 2: 'Записать заголовок сообщения
            ''Print #hFile, sMsg & "</FONT><br>"
            'Попробуем убрать лишний <br>
            ''Print #hFile, sMsg & "</FONT>"
            Print #hFile, sMsg & "</span>"
        Case 3: 'Вернуться к нулевому отступу (подчистить "хвосты" при старте новой темы)
''            If LastIndent > 0 Then
''                'Отменяем отступы до нуля
''                Do While LastIndent > 0
''                    Print #hFile, "</DL>"
''                    LastIndent = LastIndent - 1
''                Loop
''            End If
            'Начальный отступ теперь не 0, а 1.
            If LastIndent > 1 Then
                'Отменяем отступы до нуля
                Do While LastIndent > 1
                    Print #hFile, "</DL>"
                    LastIndent = LastIndent - 1
                Loop
            End If
            'Закрываем тег <DL>, стоящий после <HR>
            Print #hFile, "</DL>"
            'Отбивка нитки
            Print #hFile, "<HR><DL>"
        Case 4: 'Закрыть навигатор (запись окончена)
            Print #hFile, "</BODY></HTML>"
    End Select
    Close hFile
    WriteNavigator = True
End Function

'******************************************************
'* Стирает указанную страницу с диска, например,      *
'* если закачка была остановлена.                     *
'******************************************************
Public Function EraseHTML(ByVal Page As Long) As Boolean
    EraseHTML = False
    'Стереть страницу с сообщениями
    If Dir(App.Path & DataPath & PagesPrefix & Format(Page) & HTMLext) <> strEmpty Then
        Kill (App.Path & DataPath & PagesPrefix & Format(Page) & HTMLext)
        EraseHTML = True
    End If
    'Стереть страницу-навигатор
    If Dir(App.Path & DataPath & NavPrefix & Format(Page) & HTMLext) <> strEmpty Then
        Kill (App.Path & DataPath & NavPrefix & Format(Page) & HTMLext)
        EraseHTML = True
    End If
    'Стереть локальный индекс
    If Dir(App.Path & DataPath & IndexPrefix & Format(Page) & HTMLext) <> strEmpty Then
        Kill (App.Path & DataPath & IndexPrefix & Format(Page) & HTMLext)
        EraseHTML = True
    End If
End Function

'******************************************************
'* Обновляем базу данных страниц при успешной закачке *
'* и записываем оглавление в HTML-файл                *
'*      Записываемый параметр                         *
'*          blValue = 0 страница НЕ закачана          *
'*          blValue = 1 стараница закачана успешно    *
'*      PageLoaded - номер закачанной страницы        *
'* opt. UpdateFile - нужно ли обновлять БД страниц    *
'******************************************************
Public Function UpdateTOC(ByVal PageLoaded As Long, Optional ByVal UpdateFile As Boolean = True) As Boolean
Dim LastRecNum As Long
Dim i As Long
Dim blValue As Long 'Было Boolean, но файл никак не хотел читаться, пришлось заменить :(
Dim RetValue As Double
Dim hFile As Long
Dim hFile2 As Long
    If UpdateFile Then
        'Особенность компилятора VB: два If быстрее чем AND
        If PageLoaded > 0 Then
            'Получаем свободный указатель
            hFile = FreeFile
            'Открываем файл
            Open (App.Path & DataPath & PagesFile) For Random As hFile Len = Len(blValue)
            LastRecNum = LOF(hFile) \ Len(PageLoaded)
            If PageLoaded > LastRecNum + 1 Then
                'Подготовить записи, предшествующие необходимой
                For i = LastRecNum + 1 To PageLoaded - 1
                    blValue = 0
                    Put hFile, i, blValue
                Next i
            End If
            'Записать информацию
            blValue = 1
            Put hFile, PageLoaded, blValue
            Debug.Print "UpdateTOC -> LOG: Adding record number " & PageLoaded
            Close hFile
        End If
    End If
    'Начинаем запись HTML
    'Открываем HTML
    hFile = FreeFile
    Open (App.Path & DataPath & TOCFile) For Output As hFile
    'Открываем индекс страниц
    hFile2 = FreeFile
    Open (App.Path & DataPath & PagesFile) For Random As hFile2 Len = Len(blValue)
    LastRecNum = LOF(hFile2) \ Len(blValue)
    
    Print #hFile, "<HTML><HEAD><TITLE>Оглавление</TITLE>"
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
    Print #hFile, "        alert(""Поиск завершен. Всего было обнаружено совпадений: "" + eval(match-1));"
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
    Print #hFile, "<center><h2><u>EGCE</u></h2><h3>оглавление</h3><br>"
    Print #hFile, "<a href=""" & StartupFile & """ target=""MainFrame"" onclick=""HideIt();return true"">Начальная страница</a><br>"
    Print #hFile, "<a href=""" & FAQFile & """ target=""MainFrame"" onclick=""HideIt();return true"">ТИн и ЧаВо</a><br>"
    'TO DO Вставить ссылки на авторов
    For i = 1 To LastRecNum
        Get hFile2, i, blValue
        If blValue = 1 Then
            'Эта страница уже закачана
            'Каждой ссылке присваиваем ID для того чтобы ее можно было выбрать клавишей TAB...
            Print #hFile, "<a ID=Anchor" & Format(i) & " href=""" & "pIndex" & Format(i) & HTMLext & """ target=""MainFrame"" onMouseOver=""window.status=''; return true""  onMouseOut=""window.status=''; return false"" onclick=""ShowIt();return true"">" & Format(i) & "</a>"
        Else
            Print #hFile, Chr(32) & Format(i) & Chr(32)
        End If
        'По 10 ссылок в строке
        If i \ 10 = i / 10 Then
            Print #hFile, "<br>"
        End If
    Next i
    Print #hFile, "<br><br>Авторы:<br>"
    Print #hFile, "<a href=""" & ANamesPage & """ target=""MainFrame"" onclick=""ShowItA();return true"">по именам</a> | "
    Print #hFile, "<a href=""" & ATimePage & """ target=""MainFrame"" onclick=""ShowItA();return true"">по ""стажу""</a><br>"
    Print #hFile, "<a href=""" & ANumsPage & """ target=""MainFrame"" onclick=""ShowItA();return true"">по числу сообщений</a><br>"
    Print #hFile, "<a href=""" & ASizePage & """ target=""MainFrame"" onclick=""ShowItA();return true"">по размеру сообщений</a>"
    Print #hFile, "<br><br><a href="""" target=""MainFrame"" onclick=""HideIt();return true"">Показать файлы</a>"
    Print #hFile, "<br><br>"
    Print #hFile, "<DIV ID=SearchDiv STYLE=""visibility: Hidden"">"
    Print #hFile, "<B>Найти на странице: </B><INPUT ID=MySearch TYPE=text onfocus=""PatternChanged()"">"
    Print #hFile, "<BUTTON onclick=""doOK()"">OK</BUTTON><br>"
    Print #hFile, "<INPUT ID=match TYPE=CHECKBOX VALUE=4 UNCHECKED>Регистр"
    Print #hFile, "<INPUT ID=whole TYPE=CHECKBOX VALUE=2 UNCHECKED>Слово"
    Print #hFile, "</DIV><br>"
    Print #hFile, "<font size=1>"
    Print #hFile, "Рекомендуемое разрешение: 1024x768<br>"
    Print #hFile, "Рекомендуемый браузер: IE 4.01 или выше<br>"
    'Add version information
    Print #hFile, "Elite Games Conference Extractor (EGCE)<br>v. " & App.Major & "." & App.Minor & "." & App.Revision & "<br>"
    Print #hFile, "EGCE &copy Copyright 2001 <b>Rade</b><br>"
    Print #hFile, "Elite Games &copy 1999-2001 Сергей Петровичев a.k.a. Ranger"
    Print #hFile, "</font></center>"
    Print #hFile, "</BODY></HTML>"
    Close hFile2
    Close hFile
    
    ''TO DO организовать поиск файдов с похожими названиями и добавить их в TOC
    '(сделать это в начале функции)
    
    'Проверить индекс и, если надо, обновить!! Вставить ссылки до Close hFile
    ''TO DO параметр из настроек!!!
    Call UpdateIndex(True)
    'Покажем HTML
    Call ShowHTML
End Function

'********************************************
'* Загрузка строк из файла ресурса          *
'* Малость модифицировано под наши нужды    *
'* Все более менее нестандартное вырезано   *
'* нафиг чтобы не глючить по каждому поводу *
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
    
'ЗАГЛУШКА!!!!!!!!!!!!!!!!!!!!
Exit Sub
'ЗАГЛУШКА!!!!!!!!!!!!!!!!!!!!

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
                    'Tag <> 0  =>загружаем значения из файла ресурсов'
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
                    'Tag <> 0  =>загружаем значения из файла ресурсов'
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
'* Чистит строку от мусора, который мешает        *
'* корректной работе со строками                  *
'**************************************************
Public Function ClearGarbage(ByVal strString As String) As String
Dim tmpString As String
    tmpString = Replace(strString, "&nbsp;", Chr(32), 1, -1, vbBinaryCompare)
    tmpString = Replace(tmpString, Chr(10), strEmpty, 1, -1, vbBinaryCompare)
    tmpString = Replace(tmpString, Chr(13), strEmpty, 1, -1, vbBinaryCompare)
    ClearGarbage = tmpString
End Function

'***********************************************
'* Пишет главный индекс для показа результатов *
'* ForceOverwrite = TRUE указывает на          *
'*      необходимость принудительной           *
'*      перезаписи                             *
'***********************************************
Sub UpdateIndex(ByVal ForceOverwrite As Boolean)
    'Для начала создадим стартовую страницу
    If Dir(App.Path & DataPath & StartupFile) <> "" Then
        If ForceOverwrite Then
            Kill (App.Path & DataPath & StartupFile)
        End If
    End If
    hFile = FreeFile
    Open (App.Path & DataPath & StartupFile) For Output As hFile
    Print #hFile, "<HTML><HEAD><TITLE>Стартовая страница EliteGames Conference Extractor</TITLE>"
    Print #hFile, "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; CHARSET=Windows-1251"">"
    Print #hFile, "<META http-equiv=Pragma content=no-cache>"
    Print #hFile, "<META name=GENERATOR content=""EGCE""></HEAD>"
    Print #hFile, "<BODY alink=#000088 vlink=#008800 link=#008800><CENTER>"
    Print #hFile, "<h3>Архив старой конференции Elite Games</H3>"
    Print #hFile, "<h4>Выберите ссылку слева для демонстрации соответствующей страницы<br>"
    Print #hFile, "или одну из нижеследующих для получения более подробной информации<br>в Интернете:</H4>"
    Print #hFile, "<a href=""http://www.elite-russia.net/"">Elite Games</A><br>"
    Print #hFile, "<a href=""http://book.by.ru/cgi-bin///book.cgi?book=Elitegames"">Старая конференция Elite Games</A><br>"
    Print #hFile, "<a href=""http://x-dron.narod.ru/"">Сайт X-Dron'а</A>"
    Print #hFile, "</CENTER></BODY></HTML>"
    Close hFile
    
    'Теперь займемся индексом
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
    Print #hFile, "Ваш браузер не поддерживает фреймы. Вы увидите только оглавление.<BR><BR>"
    Print #hFile, "<a href=""pagesTOC.html"">Оглавление</A>"
    Print #hFile, "</BODY></HTML>"
    Close hFile
End Sub

'********************************************
'* Показывает окошки с сообщеними, если их  *
'* текст хранится в файле ресурсов          *
'********************************************
Public Function ShowMsgBox(ByVal resMsg As Long, ByVal MsgFlags As Long, ByVal resCaption As Long, Optional ByVal msgboxHelp As Long = -6) As VbMsgBoxResult
Dim strHelp As String
Dim strCaption As String
    strHelp = LoadResString(resMsg)
    strCaption = LoadResString(resCaption)
    If msgboxHelp = -6 Then
        'Индекс контекстно-зависимой помощи не указан - помощи не будет!
        'Спасение утопающих - дело рук самих утопающих!
        ShowMsgBox = MsgBox(strHelp, MsgFlags, strCaption)
    Else
        'Ну что ж, поможем убогим :))
        'TO DO Что там с именем файла справки? Уточнить!!!
        ShowMsgBox = MsgBox(strHelp, MsgFlags, strCaption, , msgboxHelp)
    End If
End Function

'**********************************************
'* Стирает пробелы в начале и в конце строки  *
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
'*      Записывает таблицу стилей CSS         *
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
'*    Записывает скрипт управления панелью    *
'*    навигатора                              *
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

