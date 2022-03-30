VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAuthors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Этот класс возьмет на себя обслуживание базы данных по авторам."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
''Option Compare Binary
Option Compare Text

'set this to 0 to disable debug code in this class
#Const DebugMode = 1

#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
#End If

Private Const TMPfile = "tmpAut.db"

'Локальные переменные используются ТОЛЬКО для интефейса класса
'Работа с файлом производится через переменную типа tAuthor
Private mvarLastRecNum As Long  'local copy      Read-Only
Private mvarNickname As String  'local copy
Private mvarTotalSize As Long   'local copy
Private mvarTotalNums As Long   'local copy
Private mvaremail As String     'local copy
Private mvaremailLen As Long 'local copy
Private mvarNicknameLen As Long 'local copy
'Add-on
Private mvarFirstDate As String 'local copy
Private mvarFirstDateLen As Long 'local copy


Private hFile As Long           'Указатель на файл, локальная переменная
Private hAuthor As tAuthor      'Переменная для работы с файлом
Private hHistory As tHistory    'Переменная для работы с журналом

'local variable(s) to hold property value(s)
Public Property Let FirstDateLen(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FirstDateLen = 5
    mvarFirstDateLen = vData
End Property

Public Property Get FirstDateLen() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FirstDateLen
    FirstDateLen = mvarFirstDateLen
End Property

Public Property Let FirstDate(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FirstDate = 5
    mvarFirstDate = vData
End Property

Public Property Get FirstDate() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FirstDate
    FirstDate = mvarFirstDate
End Property

Public Property Let NicknameLen(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NicknameLen = 5
    mvarNicknameLen = vData
End Property

Public Property Get NicknameLen() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NicknameLen
    NicknameLen = mvarNicknameLen
End Property

Public Property Let emailLen(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.emailLen = 5
    mvaremailLen = vData
End Property

Public Property Get emailLen() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.emailLen
    emailLen = mvaremailLen
End Property

Public Property Let email(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.email = 5
    mvaremail = vData
End Property

Public Property Get email() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.email
    email = mvaremail
End Property

Public Property Let TotalNums(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalNums = 5
    mvarTotalNums = vData
End Property

Public Property Get TotalNums() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalNums
    TotalNums = mvarTotalNums
End Property

Public Property Let TotalSize(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalSize = 5
    mvarTotalSize = vData
End Property

Public Property Get TotalSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalSize
    TotalSize = mvarTotalSize
End Property

Public Property Let Nickname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Nickname = 5
    mvarNickname = vData
End Property

Public Property Get Nickname() As String
Attribute Nickname.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Nickname
    Nickname = mvarNickname
End Property

Public Property Get LastRecNum() As Long
Dim hFile As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastRecNum
    hFile = FreeFile
    mvarLastRecNum = 0
    If CheckFile(ucAuthorsDB) Then
        Open (App.Path & DataPath & AuthorsFile) For Random Shared As hFile Len = Len(hAuthor)
        mvarLastRecNum = LOF(hFile) \ Len(hAuthor)
        Close hFile
    End If
    LastRecNum = mvarLastRecNum
End Property

'**********************************************************
'*       Поиск параметра tokenname в базе данных          *
'*  Mode=0 - поиск по имени                               *
'*  Mode=1 - поиск по e-mail (возможны несколько адресов) *
'*  На выходе получаем номер найденной записи в БД,       *
'*  либо 0, если такая запись не существует               *
'*  При поиске по eMail, ID задает номер автора в БД.     *
'*  В этом случае 0 возвращается, если адрес в БД  для ID *
'*  отсутствует или отличается от указанного.             *
'**********************************************************
Public Function SearchDB(ByVal Mode As enSearchMode, ByVal tokenname As String, ByVal tokenlen As Long, Optional ByVal ID As Long = -6) As Long
Dim i As Long
Dim hFile As Long
Dim UpperBound As Long
Dim hAuthor As tAuthor
    On Error GoTo SearchDBErr

'Запись в лог программы
LogEvent ("[Authors.SearcheDB] routine started")
    SearchDB = 0
    If CheckFile(ucAuthorsDB) Then
        UpperBound = LastRecNum
        hFile = FreeFile
        Open (App.Path & DataPath & AuthorsFile) For Random As hFile Len = Len(hAuthor)
        'Проверим один раз - сэкономим много времени!
        Select Case Mode
            Case ucSearchByName: 'Поиск по имени
                For i = 1 To UpperBound
                    Get hFile, i, hAuthor
                    If hAuthor.NicknameLen = tokenlen Then
                        'Сократим время сравнения. Будем сравнивать ники только в
                        'том случае, если число символов в них одинаково.
                        If Left(hAuthor.Nickname, hAuthor.NicknameLen) = Left(tokenname, tokenlen) Then
                            SearchDB = i
                            i = UpperBound
                        End If
                    End If
                Next i
'Запись в лог программы
LogEvent ("[Authors.SearcheDB] Поиск по имени, результат " + Str(SearchDB))
            Case ucSearchByMail: 'Поиск по e-mail
                If ID > 0 Then
                    'Задан ID автора, что сильно облегчит проверку
                    If ID <= UpperBound Then
                        'ID должен быть не больше числа зарегестрированных авторов
                        Get hFile, ID, hAuthor
                        If InStr(1, hAuthor.email, Left(tokenname, tokenlen), vbTextCompare) <> 0 Then
                            'Есть такой адрес!
                            SearchDB = ID
                        End If
                    End If
                Else
                    'Придется шарить по всей БД
                    For i = 1 To UpperBound
                        Get hFile, i, hAuthor
                        If InStr(1, hAuthor.email, Left(tokenname, tokenlen), vbTextCompare) <> 0 Then
                            'Есть такой адрес!
                            SearchDB = i
                            i = UpperBound
                        End If
                    Next i
                End If
'Запись в лог программы
LogEvent ("[Authors.SearcheDB] Поиск по eMail, результат " + Str(SearchDB))
        End Select
        Close hFile
    End If
    Exit Function

SearchDBErr:
    Call RaiseError(MyUnhandledError, "cAuthors::SearchDB Method")
End Function

'LOGIC APPROVED. INDEPENDENT VARIABLES.
'************************************************************
'* На основе БД по авторам записывает временный файл,       *
'* отсортированный по определенному ключу                   *
'************************************************************
Public Function SortDB(ByVal Mode As enAuthorFiles) As Boolean
Dim hFile1 As Long
Dim i As Long, j As Long
Dim BestPos As Long
Dim TMPrec As tAuthor
Dim strSample As String     'Для сортировки строк
Dim lValue As Long          'Для сортировки числовых значений
    On Error GoTo SortDBErr
    
    'В случае некорректных значений Mode, SortDB возвратит False
    SortDB = True
    If CheckFile(ucAuthorsDB) Then
        'Копируем БД авторов во временный файл для последующей обработки
''{FIXED} TO DO BUG file already open
        FileCopy (App.Path & DataPath & AuthorsFile), (App.Path & DataPath & TMPfile)
        'Открываем временный файл
        hFile1 = FreeFile
        Open (App.Path & DataPath & TMPfile) For Random As hFile1 Len = Len(hAuthor)
        Select Case Mode
            Case ucAuthorsNames:
                    strSample = strEmpty
                    'Сортируем по убыванию (так проще)
                    For i = 1 To LastRecNum 'Файлы пока идентичны, так что можно и так сделать
                        BestPos = i         'Если не изменится, значит, переставлять не нужно
                        Seek hFile1, i      'Указываем место, с какого надо читать файл
                        For j = i To LastRecNum
                            'Читаем очередную запись
                            Get hFile1, , TMPrec
                            'Сравниваем ники: если текущий ник ближе к концу алфавита...
                            If strSample < Left(TMPrec.Nickname, TMPrec.NicknameLen) Then
                                strSample = Left(TMPrec.Nickname, TMPrec.NicknameLen)
                                BestPos = Seek(hFile1) - 1 'Запомнить положение лучшей записи
                            End If
                        Next j
                        If Not (BestPos = i) Then
                            Get hFile1, i, hAuthor          'Прочитать старую запись
                            Get hFile1, BestPos, TMPrec     'Прочитать новую запись
                            Seek hFile1, i
                            Put hFile1, , TMPrec            'Записываем новую запись
                            Seek hFile1, BestPos
                            Put hFile1, , hAuthor           'Записываем старую запись
                        End If
                        strSample = strEmpty  'Обнуляем, чтобы корректно продолжать сравнения
                    Next i
            Case ucAuthorsNums:
                    lValue = 0
                    'Сортируем по убыванию (так проще)
                    For i = 1 To LastRecNum 'Файлы пока идентичны, так что можно и так сделать
                        BestPos = i         'Если не изменится, значит, переставлять не нужно
                        Seek hFile1, i      'Указываем место, с какого надо читать файл
                        For j = i To LastRecNum
                            'Читаем очередную запись
                            Get hFile1, , TMPrec
                            'Сравниваем число постингов
                            If lValue < TMPrec.TotalNums Then
                                lValue = TMPrec.TotalNums
                                BestPos = Seek(hFile1) - 1 'Запомнить положение лучшей записи
                            End If
                        Next j
                        If Not (BestPos = i) Then
                            Get hFile1, i, hAuthor          'Прочитать старую запись
                            Get hFile1, BestPos, TMPrec     'Прочитать новую запись
                            Seek hFile1, i
                            Put hFile1, , TMPrec            'Записываем новую запись
                            Seek hFile1, BestPos
                            Put hFile1, , hAuthor           'Записываем старую запись
                        End If
                        lValue = 0      'Обнуляем, чтобы корректно продолжать сравнения
                    Next i
            Case ucAuthorsSize:
                    lValue = 0
                    'Сортируем по убыванию (так проще)
                    For i = 1 To LastRecNum 'Файлы пока идентичны, так что можно и так сделать
                        BestPos = i         'Если не изменится, значит, переставлять не нужно
                        Seek hFile1, i      'Указываем место, с какого надо читать файл
                        For j = i To LastRecNum
                            'Читаем очередную запись
                            Get hFile1, , TMPrec
                            'Сравниваем размер постингов
                            If lValue < TMPrec.TotalSize Then
                                lValue = TMPrec.TotalSize
                                BestPos = Seek(hFile1) - 1 'Запомнить положение лучшей записи
                            End If
                        Next j
                        If Not (BestPos = i) Then
                            Get hFile1, i, hAuthor          'Прочитать старую запись
                            Get hFile1, BestPos, TMPrec     'Прочитать новую запись
                            Seek hFile1, i
                            Put hFile1, , TMPrec            'Записываем новую запись
                            Seek hFile1, BestPos
                            Put hFile1, , hAuthor           'Записываем старую запись
                        End If
                        lValue = 0      'Обнуляем, чтобы корректно продолжать сравнения
                    Next i
            Case ucAuthorsTime:
                'Сортировка по времени первого сообщения
                strSample = strEmpty
                'Сортируем по убыванию (так проще)
                'Авторы-ветераны конфы будут расположены в конце списка
                For i = 1 To LastRecNum 'Файлы пока идентичны, так что можно и так сделать
                    BestPos = i         'Если не изменится, значит, переставлять не нужно
                    Seek hFile1, i      'Указываем место, с какого надо читать файл
                    For j = i To LastRecNum
                        'Читаем очередную запись
                        Get hFile1, , TMPrec
                        'Сравниваем даты
                        If strSample < Left(TMPrec.FirstDate, TMPrec.FirstDateLen) Then
                            strSample = Left(TMPrec.FirstDate, TMPrec.FirstDateLen)
                            BestPos = Seek(hFile1) - 1 'Запомнить положение лучшей записи
                        End If
                    Next j
                    If Not (BestPos = i) Then
                        Get hFile1, i, hAuthor          'Прочитать старую запись
                        Get hFile1, BestPos, TMPrec     'Прочитать новую запись
                        Seek hFile1, i
                        Put hFile1, , TMPrec            'Записываем новую запись
                        Seek hFile1, BestPos
                        Put hFile1, , hAuthor           'Записываем старую запись
                    End If
                    strSample = strEmpty  'Обнуляем, чтобы корректно продолжать сравнения
                Next i
            Case Else:
                'Задан некорректный параметр
                SortDB = False
        End Select
        'Закрываем файл
        Close hFile1
    Else
        'БД по авторам не обнаружена
        SortDB = False
    End If
    'Временный файл записан. Переводом его в HTML займется WriteAHTMLs
    Exit Function
    
SortDBErr:
    Call RaiseError(MyUnhandledError, "cAuthors::SortDB Method")
End Function

'LOGIC APPROVED. INDEPENDENT VARIABLES.
'************************************************************
'*     Записывает HTML-файлы оглавления БД авторов          *
'************************************************************
Public Function WriteAHTMLs() As Boolean
Dim i As Long, j As Long
Dim hFile1 As Long
Dim hFile2 As Long
Dim lTMPval As Long
    On Error GoTo WriteAHTMLsErr
''TO DO ВНИМАНИЕ!!! Записывать только ПОСЛЕДНИЙ e-mail из нескольких возможных!!!
    'Применим метод "от противного": если будет хоть одна ошибка, то тогда WriteAHTMLs = False
    WriteAHTMLs = True
    For i = 1 To 4
        If SortDB(i) Then
            'Открываем временный файл
            hFile2 = FreeFile
            Open (App.Path & DataPath & TMPfile) For Random As hFile2 Len = Len(hAuthor)
            Select Case i
                Case 1: 'Сортировка по именам авторов
                    If CheckFile(ucAuthorsNames) Then
                        Kill App.Path & DataPath & ANamesPage
                    End If
                    'Открываем HTML-файл
                    hFile1 = FreeFile
                    Open (App.Path & DataPath & ANamesPage) For Output As hFile1
                    Print #hFile1, "<HTML><HEAD><TITLE>EGCE - Авторы в алфавитном порядке</TITLE>"
                Case 2: 'Сортировка по количеству сообщений
                    If CheckFile(ucAuthorsNums) Then
                        Kill App.Path & DataPath & ANumsPage
                    End If
                    'Открываем HTML-файл
                    hFile1 = FreeFile
                    Open (App.Path & DataPath & ANumsPage) For Output As hFile1
                    Print #hFile1, "<HTML><HEAD><TITLE>EGCE - Авторы по количеству сообщений</TITLE>"
                Case 3: 'Сортировка по размеру сообщений
                    If CheckFile(ucAuthorsSize) Then
                        Kill App.Path & DataPath & ASizePage
                    End If
                    'Открываем HTML-файл
                    hFile1 = FreeFile
                    Open (App.Path & DataPath & ASizePage) For Output As hFile1
                    Print #hFile1, "<HTML><HEAD><TITLE>EGCE - Авторы по суммарному размеру сообщений</TITLE>"
                Case 4: 'Сортировка по времени первого сообщения
                    If CheckFile(ucAuthorsTime) Then
                        Kill App.Path & DataPath & ATimePage
                    End If
                    'Открываем HTML-файл
                    hFile1 = FreeFile
                    Open (App.Path & DataPath & ATimePage) For Output As hFile1
                    Print #hFile1, "<HTML><HEAD><TITLE>EGCE - Авторы по времени появления первого сообщения</TITLE>"
            End Select
            Print #hFile1, "<LINK HREF=""" & CSSFile & """ REL=stylesheet TYPE=""text/css"">"
            Print #hFile1, "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; CHARSET=Windows-1251"">"
            Print #hFile1, "<META http-equiv=Pragma content=no-cache>"
            Print #hFile1, "<META name=GENERATOR content=""EGCE"">"
            ''Print #hFile1, "<STYLE type=text/css>"
            ''Print #hFile1, "BODY{"
            ''Print #hFile1, "a           {text-decoration:none; color:#006600}"
            ''Print #hFile1, "a:link      {text-decoration:none; color:#006600}"
            ''Print #hFile1, "a:active    {text-decoration:none; color:#00AA00}"
            ''Print #hFile1, "a:hover     {text-decoration:underline; color:#00AA00}"
            ''Print #hFile1, "a:visited   {text-decoration:none; color:#0000AA}"
            ''Print #hFile1, "}</STYLE></HEAD>"
            Print #hFile1, "</HEAD><BODY><CENTER>"
            Print #hFile1, "<TABLE style='text-align:justify' border=0 cellPadding=5 cellSpacing=1 width=""85%""><TBODY>"
            ''Print #hFile1, "<TR><TD vAlign=center width=""40%"" bgColor=#000090><FONT color=#FFFFFF face=""Verdana, Arial, Helvetica, Geneva"" size=1><B>Автор</B></FONT></TD>"
            ''Print #hFile1, "<TD vAlign=center width=""13%"" bgColor=#000090><FONT color=#FFFFFF face=""Verdana, Arial, Helvetica, Geneva"" size=1><B>Сообщений</B></FONT></TD>"
            ''Print #hFile1, "<TD vAlign=center width=""15%"" bgColor=#000090><FONT color=#FFFFFF face=""Verdana, Arial, Helvetica, Geneva"" size=1><B>Размер, байт</B></FONT></TD>"
            ''Print #hFile1, "<TD vAlign=center width=""32%"" bgColor=#000090><FONT color=#FFFFFF face=""Verdana, Arial, Helvetica, Geneva"" size=1><B>Дата регистрации (ГГГГ.ММ.ДД)</B></FONT></TD></TR>"
            Print #hFile1, "<TR><TD vAlign=center width=""40%"" bgColor=#000090><span class=""Info"">Автор</span></TD>"
            Print #hFile1, "<TD vAlign=center width=""13%"" bgColor=#000090><span class=""Info"">Сообщений</span></TD>"
            Print #hFile1, "<TD vAlign=center width=""15%"" bgColor=#000090><span class=""Info"">Размер, байт</span></TD>"
            Print #hFile1, "<TD vAlign=center width=""32%"" bgColor=#000090><span class=""Info"">Дата регистрации (ГГГГ.ММ.ДД)</span></TD></TR>"
            'Запись собственно информации
            'Список во временном файле записан в обратном порядке! Учитываем это.
            lTMPval = LastRecNum
            Select Case i
                Case 1:
                    For j = lTMPval To 1 Step -1
                        Get hFile2, j, hAuthor
                        'TO DO Включить запись ссылки на e-Mail автора
                        Print #hFile1, "<TR><TD><B>" & Left(hAuthor.Nickname, hAuthor.NicknameLen) & "</B></TD>"
                        Print #hFile1, "<TD align=right>" & Format(hAuthor.TotalNums) & "</TD>"
                        Print #hFile1, "<TD align=right>" & Format(hAuthor.TotalSize) & "</TD>"
                        Print #hFile1, "<TD align=right>" & Left(hAuthor.FirstDate, hAuthor.FirstDateLen) & "</TD></TR>"
                    Next j
                Case 2:
                    For j = 1 To lTMPval
                        Get hFile2, j, hAuthor
                        'TO DO Включить запись ссылки на e-Mail автора
                        Print #hFile1, "<TR><TD><B>" & Left(hAuthor.Nickname, hAuthor.NicknameLen) & "</B></TD>"
                        Print #hFile1, "<TD align=right>" & Format(hAuthor.TotalNums) & "</TD>"
                        Print #hFile1, "<TD align=right>" & Format(hAuthor.TotalSize) & "</TD>"
                        Print #hFile1, "<TD align=right>" & Left(hAuthor.FirstDate, hAuthor.FirstDateLen) & "</TD></TR>"
                    Next j
                Case 3:
                    For j = 1 To lTMPval
                        Get hFile2, j, hAuthor
                        'TO DO Включить запись ссылки на e-Mail автора
                        Print #hFile1, "<TR><TD><B>" & Left(hAuthor.Nickname, hAuthor.NicknameLen) & "</B></TD>"
                        Print #hFile1, "<TD align=right>" & Format(hAuthor.TotalNums) & "</TD>"
                        Print #hFile1, "<TD align=right>" & Format(hAuthor.TotalSize) & "</TD>"
                        Print #hFile1, "<TD align=right>" & Left(hAuthor.FirstDate, hAuthor.FirstDateLen) & "</TD></TR>"
                    Next j
                Case 4:
                    For j = lTMPval To 1 Step -1
                        Get hFile2, j, hAuthor
                        'TO DO Включить запись ссылки на e-Mail автора
                        Print #hFile1, "<TR><TD><B>" & Left(hAuthor.Nickname, hAuthor.NicknameLen) & "</B></TD>"
                        Print #hFile1, "<TD align=right>" & Format(hAuthor.TotalNums) & "</TD>"
                        Print #hFile1, "<TD align=right>" & Format(hAuthor.TotalSize) & "</TD>"
                        Print #hFile1, "<TD align=right>" & Left(hAuthor.FirstDate, hAuthor.FirstDateLen) & "</TD></TR>"
                    Next j
            End Select
            Print #hFile1, "</TBODY></TABLE><hr width=85%>"
            Print #hFile1, "<font size=1>"
            Print #hFile1, "This page was generated by:<br>"
            'Add version information
            Print #hFile1, "Elite Games Conference Extractor (EGCE) v. " & App.Major & "." & App.Minor & "." & App.Revision & "<br>"
            Print #hFile1, "EGCE &copy Copyright 2001 <b>Rade</b><br>"
            Print #hFile1, "Elite Games &copy 1999-2001 Сергей Петровичев a.k.a. Ranger"
            Print #hFile1, "</font>"
            Print #hFile1, "</CENTER></BODY></HTML>"
            'Закрываем временный файл
            Close hFile2
            'Закрываем текущий HTML-файл с сведениями об авторах
            Close hFile1
            'Стираем временный файл
            Kill App.Path & DataPath & TMPfile
        Else
            WriteAHTMLs = False
        End If
    Next i
    Exit Function
    
WriteAHTMLsErr:
    Call RaiseError(MyUnhandledError, "cAuthors::WriteAHTMLs Method")
End Function

'**********************************************************
'* Проверяет наличие записи в БД авторов. Если запись     *
'* есть, то обновляет ее, если нет, то создает ее         *
'*      PrimaryMode - запись ника в БД, если его там нет  *
'*      HistoryMode - обновление журнала и подсчет        *
'*                    статистики по авторам               *
'**********************************************************
Public Function UpdateDB(ByVal Mode As enFileMode) As Boolean
Dim SearchRes As Long
Dim HistoryLen As Long       'Количество записей в журнале
Dim hHFile As Long  'Указатель на файл журнала
Dim hFile As Long
Dim i As Long
Dim hAuthor As tAuthor
    On Error GoTo UpdateDBErr

    UpdateDB = False
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] entering routine")
    Select Case Mode
        Case ucPrimaryMode: 'Проверка существования ников в главной БД
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] режим ucPrimaryMode")
        ''TO DO исправить: оставить только поиск и запись новых, убрать подсчет длин
        ''и количеств сообщений, это будет сделано в HistoryMode
            SearchRes = SearchDB(ucSearchByName, Nickname, NicknameLen)
            If Not (SearchRes = 0) Then
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Автор " + Left(Nickname, NicknameLen) + " уже зарегестрирован под номером " + Str(SearchRes))
                'Запись уже существует, обновляем ее!
                hFile = FreeFile
                Open (App.Path & DataPath & AuthorsFile) For Random As hFile Len = Len(hAuthor)
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Файл авторов открыт")
                Get hFile, SearchRes, hAuthor
                                            'DEBUG routine
                                            If Left(hAuthor.Nickname, hAuthor.NicknameLen) = Nickname Then
                                                Debug.Print "UpdateDB method -> Nickname confirmed"
                                            Else
                                                Debug.Print "UpdateBD method -> ATTENTION!!! Nickname check failed!"
                                            End If
                                            'End of the DEBUG routine
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] eMail автора " + hAuthor.email)
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Было  постингов\размер " + Str(hAuthor.TotalNums) + " \ " + Str(hAuthor.TotalSize))
                'Заполняем поля записи
                With hAuthor
                    .TotalSize = .TotalSize + TotalSize
                    .TotalNums = .TotalNums + TotalNums
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Стало постингов\размер " + Str(.TotalNums) + " \ " + Str(.TotalSize))
                                                Debug.Print "UpdateDB method -> "; Nickname; " TotSize="; .TotalSize; "TotNum="; .TotalNums
                    If emailLen > 0 Then
                    'e-mail указан - проверяем его наличие в БД
                        If InStr(1, hAuthor.email, Left(email, emailLen), vbTextCompare) = 0 Then
                        ''If SearchDB(ucSearchByMail, email, emailLen, SearchRes) = 0 Then
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Поиск по eMail " + Left(email, emailLen) + " результатов не дал, добавляем адрес")
                        ''Если поиск по e-mail =0, то добавляем адрес, иначе адрес уже есть
                            If .emailLen = 0 Then
                                .email = Left(email, emailLen)
                                .emailLen = emailLen
                            Else
                                .email = Left(.email, .emailLen) & ";" & Left(email, emailLen)
                                .emailLen = .emailLen + 1 + emailLen
                            End If
                        End If
                    Else
                        'Очистим поля (если надо)
                        If Not .emailLen > 0 Then
                            .emailLen = 0
                            .email = strEmpty
                        End If
                    End If
                    'Обновим поле eMail для cAuthor, чтобы программа могла сразу же
                    'обновить свои переменные для отражения текущей картины в HTML
                    'Для этого проведем обратное присваивание
                    emailLen = .emailLen
                    email = Left(.email, emailLen)
                    'Проверяем, а не является ли эта запись более ранней?
                    'Формат времени YYYY.MM.DD (HH:MM:SS)
                    If Not (FirstDateLen <= 0) Then
                        If Left(.FirstDate, .FirstDateLen) > Left(FirstDate, FirstDateLen) Then
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Обновление даты регистрации")
                            'Откопали что-то постарее - обновим поля
                            .FirstDate = Left(FirstDate, FirstDateLen)
                            .FirstDateLen = FirstDateLen
                        End If
                    End If
                End With
''{FIXED} TO DO BUG Bad record number
                Put hFile, SearchRes, hAuthor
                'Не будем экономить копейки, закроем файл сразу же после записи
                Close hFile
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Файл авторов закрыт")
                UpdateDB = True
            Else
                'Новая запись!
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Новый автор - " + Left(Nickname, NicknameLen))
                                                Debug.Print "UpdateDB method -> new Nickname"
                SearchRes = LastRecNum + 1
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Регистрация нового автора с ID " + Str(SearchRes))
                hFile = FreeFile
                Open (App.Path & DataPath & AuthorsFile) For Random As hFile Len = Len(hAuthor)
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Файл авторов открыт")
                With hAuthor
                    .NicknameLen = NicknameLen
                    .Nickname = Left(Nickname, NicknameLen)
                    .TotalSize = TotalSize
                    .TotalNums = TotalNums
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Постингов\размер " + Str(.TotalNums) + " \ " + Str(.TotalSize))
                    If emailLen = 0 Then
                        .email = strEmpty
                        .emailLen = 0
                    Else
                        .emailLen = emailLen
                        .email = Left(email, emailLen)
                    End If
                    If Not (FirstDateLen <= 0) Then
                        .FirstDateLen = FirstDateLen
                        .FirstDate = Left(FirstDate, FirstDateLen)
                    End If
                End With
''TO DO BUG bad record length? С какой стати?
                Put hFile, SearchRes, hAuthor
                'Не будем экономить копейки, закроем файл сразу же после записи
                Close hFile
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Файл авторов закрыт")
                UpdateDB = True
            End If
        Case ucHistoryMode: 'Сбор статистики и обновление главной БД
            Dim lElements() As Long     'Указатель на записи в файле авторов
                        
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Режим ucHistoryMode")
            '(ID,0) - Число сообщений
            '(ID,1) - Размер сообщений
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Всего авторов в БД " + Str(LastRecNum))
            ReDim lElements(LastRecNum, 1)
'Запись в лог программы
LogEvent ("[Authors.UpdateDB] Рестарт статистикии БД авторов")
            'Стираем в БД авторов статистику
            RestartDB (ucClearMathOnly)
            'Открываем журнал
            hHFile = FreeFile
            Open (App.Path & DataPath & HistoryFile) For Random As hHFile Len = Len(hHistory)
''''''''''            'Открываем БД авторов
''''''''''            hFile = FreeFile
''''''''''            Open (App.Path & DataPath & AuthorsFile) For Random As hFile Len = Len(hAuthor)
            'Определяем количество записей в журнале
            HistoryLen = LOF(hHFile) \ Len(hHistory)
            If HistoryLen > 0 Then   'Журнал существует
                For i = 1 To HistoryLen
                    'Читаем очередную запись в журнале
                    Get hHFile, i, hHistory
                    If hHistory.AuthorID > 0 Then
                        lElements(hHistory.AuthorID, 0) = lElements(hHistory.AuthorID, 0) + 1
                        lElements(hHistory.AuthorID, 1) = lElements(hHistory.AuthorID, 1) + hHistory.MsgSize
''''''''''                        'Читаем соответствующую ей запись в БД авторов
''''''''''                        Get hFile, hHistory.AuthorID, hAuthor
''''''''''                        hAuthor.TotalNums = hAuthor.TotalNums + 1
''''''''''                        hAuthor.TotalSize = hAuthor.TotalSize + hHistory.MsgSize
''''''''''                        Put hFile, hHistory.AuthorID, hAuthor
                    End If
                Next i
            End If
            Close hHFile
            hFile = FreeFile
            Open (App.Path & DataPath & AuthorsFile) For Random As hFile Len = Len(hAuthor)
            For i = 1 To LastRecNum
                        Get hFile, i, hAuthor
                        hAuthor.TotalNums = lElements(i, 0)
                        hAuthor.TotalSize = lElements(i, 1)
                        Put hFile, i, hAuthor
            Next i
            Close hFile
            ReDim lElements(0, 0) 'Возвращаем память системе
            UpdateDB = True
    End Select
    Exit Function

UpdateDBErr:
    Call RaiseError(MyUnhandledError, "cAuthors::UpdateDB Method")
End Function

'************************************************************
'*         Проверить существование файла                    *
'*  ucAuthorsDB:    БД по авторам                           *
'*  ucAuthorsNames: HTML-страница с авторами в алфавитном   *
'*                  порядке                                 *
'*  ucAuthorsNums:  HTML-страница с авторами по количеству  *
'*                  сообщений                               *
'*  ucAuthorsSize:  HTML-страница с авторами по размеру     *
'*                  сообщений                               *
'*  ucHistoryDB:    журнал закачанных сообщений             *
'************************************************************
Private Function CheckFile(ByVal Mode As enAuthorFiles) As Boolean
    On Error GoTo CheckFileErr
    
    CheckFile = False
    Select Case Mode
        Case ucAuthorsDB:
            CheckFile = IIf(Dir(App.Path & DataPath & AuthorsFile, vbNormal) <> "", True, False)
            ''If (Dir(App.Path & DataPath & AuthorsFile, vbNormal) <> "") Then
            ''    'OK, файл существует
            ''    CheckFile = True
            ''End If
        Case ucAuthorsNames:
            CheckFile = IIf(Dir(App.Path & DataPath & ANamesPage, vbNormal) <> "", True, False)
        Case ucAuthorsNums:
            CheckFile = IIf(Dir(App.Path & DataPath & ANumsPage, vbNormal) <> "", True, False)
        Case ucAuthorsSize:
            CheckFile = IIf(Dir(App.Path & DataPath & ASizePage, vbNormal) <> "", True, False)
        Case ucHistoryDB:
            CheckFile = IIf(Dir(App.Path & DataPath & HistoryFile, vbNormal) <> "", True, False)
    End Select
    Exit Function
    
CheckFileErr:
    CheckFile = False
    Call RaiseError(MyUnhandledError, "cAuthors::CheckFile Method")
End Function

'***********************************************************
'* Рестарт БД: стирает всю информацию, кроме имен и e-mail *
'* Функция обслуживания БД, в основном, нужна при отладке, *
'* но никто не мешает включить ее в поздних версиях        *
'* Возвращает FALSE, если БД нет или она пуста             *
'***********************************************************
Public Function RestartDB(ByVal Mode As enRestartADB) As Boolean
Dim hFileAtmp As Long
Dim hFileA As Long
Dim UpperBound As Long
Dim i As Long
    On Error GoTo RestartDBErr

    RestartDB = False
    Select Case Mode
        Case ucClearAll:
            frmBrowser.Label14.Caption = "рестарт базы данных авторов - полное удаление информации"
            frmBrowser.Label14.Refresh
            If CheckFile(ucAuthorsDB) Then
                Kill (App.Path & DataPath & AuthorsFile)
            End If
        Case ucClearMathOnly:
            frmBrowser.Label14.Caption = "рестарт базы данных авторов - удаление статистической информации"
            frmBrowser.Label14.Refresh
            If CheckFile(ucAuthorsDB) Then
                'Определяем количество записей БД
                UpperBound = LastRecNum
                hFileA = FreeFile
                Open (App.Path & DataPath & AuthorsFile) For Random As hFileA Len = Len(hAuthor)
                'Если временный файл существует, стереть его
                If (Dir(App.Path & DataPath & TMPfile, vbNormal) <> "") Then
                    'файл существует, стираем
                    Kill (App.Path & DataPath & TMPfile)
                End If
                'Временный файл
                hFileAtmp = FreeFile
                Open (App.Path & DataPath & TMPfile) For Random As hFileAtmp Len = Len(hAuthor)
                For i = 1 To UpperBound
                    'Читаем запись
                    Get hFileA, , hAuthor
                    'Очищаем лишнее
                    hAuthor.TotalSize = 0
                    hAuthor.TotalNums = 0
                    'Пишем во временный файл
                    Put hFileAtmp, i, hAuthor
                Next i
                'Закрываем файлы (именно в этой последовательности, иначе ОШИБКА!!!)
                Close hFileA
                Close hFileAtmp
                'Стираем старый
''{FIXED}TO DO BUG FILE ALREADY OPEN
                Kill (App.Path & DataPath & AuthorsFile)
                'Переименовываем
''{FIXED} FILE NOT FOUND
                Name (App.Path & DataPath & TMPfile) As (App.Path & DataPath & AuthorsFile)
                RestartDB = True
            End If
    End Select
    frmBrowser.Label14.Caption = strEmpty
    frmBrowser.Label14.Refresh
    Exit Function

RestartDBErr:
    Call RaiseError(MyUnhandledError, "cAuthors::RestartDB Method")
End Function

Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
    With hAuthor
        .Nickname = ""
        .NicknameLen = 0
        .TotalSize = 0
        .TotalNums = 0
        .email = ""
        .emailLen = 0
    End With
End Sub

Private Sub Class_Terminate()
    'the class is being destroyed
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    Public Property Get ClassDebugID()
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

