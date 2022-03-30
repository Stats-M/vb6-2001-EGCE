VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMsgMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Задача этого класса - исключить пожирание большого количества памяти, заменив операции с памятью дисковыми операциями. Обеспечить прозрачность этих операций (в пределах возможного, разумеется)"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

'set this to 0 to disable debug code in this class
#Const DebugMode = 1

#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
#End If

Private Const TMPfile = "History.tmp"   'Временный файл для вспомогательных целей

'Локальные переменные используются ТОЛЬКО для интефейса класса
'Работа с файлом производится через переменную типа tMessage
Private mvarTopicNum As Long 'local copy
Private mvarmsgURL As String 'local copy
Private mvarAuthor As String 'local copy
Private mvaremail As String 'local copy
Private mvarmsgHead As String 'local copy
Private mvarmsgSize As Long 'local copy
Private mvarmsgDate As String 'local copy
Private mvarAuthorLen As Long 'local copy
Private mvaremailLen As Long 'local copy
Private mvarmsgDateLen As Long 'local copy
Private mvarmsgHeadLen As Long 'local copy
Private mvarmsgURLLen As Long 'local copy
Private mvarLastRecNum As Long 'local copy      Read-Only !!!
Private mvarCurrRec As Long 'local copy
Private mvarIndent As Long 'local copy

Private hFile As Long           'Указатель на файл, локальная переменная
Private hMessage As tMessage    'Переменная для работы с файлом
Private hHistory As tHistory    'Переменная для работы с журналом

Public Property Let Indent(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Indent = 5
    mvarIndent = vData
End Property

Public Property Get Indent() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Indent
    Indent = mvarIndent
End Property

Public Property Let msgURLLen(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.msgURLLen = 5
    mvarmsgURLLen = vData
End Property

Public Property Get msgURLLen() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.msgURLLen
    msgURLLen = mvarmsgURLLen
End Property

Public Property Let msgHeadLen(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.msgHeadLen = 5
    mvarmsgHeadLen = vData
End Property

Public Property Get msgHeadLen() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.msgHeadLen
    msgHeadLen = mvarmsgHeadLen
End Property

Public Property Let msgDateLen(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.msgDateLen = 5
    mvarmsgDateLen = vData
End Property

Public Property Get msgDateLen() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.msgDateLen
    msgDateLen = mvarmsgDateLen
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

Public Property Let AuthorLen(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AuthorLen = 5
    mvarAuthorLen = vData
End Property

Public Property Get AuthorLen() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AuthorLen
    AuthorLen = mvarAuthorLen
End Property

Public Property Let CurrRec(ByVal vData As Long)
    On Error GoTo CurrRecLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.msgSize = 5
    mvarCurrRec = vData
    Exit Property
CurrRecLetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::CurrRec Property Let")
End Property

Public Property Get CurrRec() As Long
    On Error GoTo CurrRecErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrRec
    CurrRec = mvarCurrRec
    Exit Property
CurrRecErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::CurrRec Property Get")
End Property

Public Property Get LastRecNum() As Long
Dim hFile As Long   'Указатель на файл
    On Error GoTo LastRecNumErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastRecNum
    hFile = FreeFile
    mvarLastRecNum = 0
    If CheckFile(ucMsgURLsMode) Then
        Open (App.Path & DataPath & MsgFile) For Random Shared As hFile Len = Len(hMessage)
        mvarLastRecNum = LOF(hFile) \ Len(hMessage)
        Close hFile
    End If
    LastRecNum = mvarLastRecNum
    Exit Property
LastRecNumErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::LastRecNum Property Get")
End Property

Public Property Let msgDate(ByVal vData As String)
    On Error GoTo msgDateLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.msgDate = 5
    mvarmsgDate = vData
    Exit Property
msgDateLetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::msgDate Property Let")
End Property

Public Property Get msgDate() As String
    On Error GoTo msgDateGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.msgDate
    msgDate = mvarmsgDate
    Exit Property
msgDateGetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::msgDate Property Get")
End Property

Public Property Let MsgSize(ByVal vData As Long)
    On Error GoTo msgSizeLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.msgSize = 5
    mvarmsgSize = vData
    Exit Property
msgSizeLetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::msgSize Property Let")
End Property

Public Property Get MsgSize() As Long
    On Error GoTo msgSizeGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.msgSize
    MsgSize = mvarmsgSize
    Exit Property
msgSizeGetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::msgSize Property Get")
End Property

Public Property Let msgHead(ByVal vData As String)
    On Error GoTo msgHeadLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.msgHead = 5
    mvarmsgHead = vData
    Exit Property
msgHeadLetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::msgHead Property Let")
End Property

Public Property Get msgHead() As String
    On Error GoTo msgHeadGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.msgHead
    msgHead = mvarmsgHead
    Exit Property
msgHeadGetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::msgHead Property Get")
End Property

Public Property Let email(ByVal vData As String)
    On Error GoTo emailLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.email = 5
    mvaremail = vData
    Exit Property
emailLetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::email Property Let")
End Property

Public Property Get email() As String
    On Error GoTo emailGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.email
    email = mvaremail
    Exit Property
emailGetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::email Property Get")
End Property

Public Property Let Author(ByVal vData As String)
    On Error GoTo AuthorLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Author = 5
    mvarAuthor = vData
    Exit Property
AuthorLetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::Author Property Let")
End Property

Public Property Get Author() As String
    On Error GoTo AuthorGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Author
    Author = mvarAuthor
    Exit Property
AuthorGetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::Author Property Get")
End Property

Public Property Let msgURL(ByVal vData As String)
    On Error GoTo msgURLLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.msgURL = 5
    mvarmsgURL = vData
    Exit Property
msgURLLetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::msgURL Property Let")
End Property

Public Property Get msgURL() As String
Attribute msgURL.VB_UserMemId = 0
    On Error GoTo msgURLGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.msgURL
    msgURL = mvarmsgURL
    Exit Property
msgURLGetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::msgURL Property Get")
End Property

Public Property Let TopicNum(ByVal vData As Long)
    On Error GoTo TopicNumLetErr
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TopicNum = 5
    mvarTopicNum = vData
    Exit Property
TopicNumLetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::TopicNum Property Let")
End Property

Public Property Get TopicNum() As Long
    On Error GoTo TopicNumGetErr
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TopicNum
    TopicNum = mvarTopicNum
    Exit Property
TopicNumGetErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::TopicNum Property Get")
End Property

'**********************************************************
'*  Сохраняет очередную запись в базе данных по ссылкам   *
'*  Возвращает номер записи при успешном завершениии,     *
'*  в противном случае возвращает ноль                    *
'**********************************************************
Public Function Save() As Long
Dim lValue As Long
    On Error GoTo SaveErr
    
    Save = 0
    ''If CheckFile Then
        lValue = LastRecNum + 1
        With hMessage
            .TopicNum = TopicNum
            .msgURL = msgURL
            .msgURLLen = msgURLLen
            .Author = Author
            .AuthorLen = AuthorLen
'ВНИМАНИЕ!!! В cMsgMap e-mail ТОЛЬКО ОДИН либо никакой!!! Это используется
'в ParserLevel2. Сбор информации об e-mail производится уже в cAuthors, куда
'передаются сведения, в том числе и отсюда
            .email = email
            .emailLen = emailLen
'msgHead используется для отлова ошибок индексации сервера book.ru
'Заголовок приходится резать дважды и сравнивать
            .msgHead = msgHead
            .msgHeadLen = msgHeadLen
            .MsgSize = MsgSize
            .msgDate = msgDate
            .msgDateLen = msgDateLen
            .Indent = Indent
        End With
        hFile = FreeFile
        Open (App.Path & DataPath & MsgFile) For Random As hFile Len = Len(hMessage)
        '{Fix} TO DO BUG BadRecLen
        Put hFile, lValue, hMessage
        Close hFile
        Save = lValue
        Debug.Print "cMsgMap:Save Method  Save successfull! RecID = "; Str(lValue)
    ''End If
    Exit Function

SaveErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::Save Method")
End Function

'**********************************************************
'*  Обновляет базу данных                                 *
'*                                                        *
'*      Mode = ucMsgURLsMode                              *
'*  Обновляет запись в базе данных по ссылкам, дописывая  *
'*  дату и размер сообщения (вызов из ParserLevel2)       *
'*      Mode = ucHistoryMode                              *
'*  Обновляет журнал.                                     *
'*  Вызывается после окончания обработки всех ссылок      *
'*  страницы или после нажатия кнопки СТОП.               *
'**********************************************************
Public Function UpdateDB(ByVal Mode As enFileMode, Optional ByVal RecNum As Long = -6) As Boolean
Dim i As Long
Dim hFile As Long
    On Error GoTo UpdateDBErr
    
    'Запись в лог программы
    LogEvent ("[MsgMap.UpdateDB] entering routine")
    UpdateDB = False
    Select Case Mode
        Case ucMsgURLsMode:
            'Запись в лог программы
            LogEvent ("[MsgMap.UpdateDB] режим ucMsgURLsMode")
            If RecNum <= LastRecNum And Not (RecNum = -6) Then
                'Запись в лог программы
                LogEvent ("[MsgMap.UpdateDB] Начало работы с записью номер " + Str(RecNum))
                If CheckFile(ucMsgURLsMode) Then
                    With hMessage
                        .TopicNum = TopicNum
                        '.msgURL = msgURL
                        If msgURLLen > 0 Then
                            .msgURL = Left(msgURL, msgURLLen)
                        Else
                            .msgURL = strEmpty
                        End If
                        .msgURLLen = msgURLLen
                        '.Author = Author
                        If AuthorLen > 0 Then
                            .Author = Left(Author, AuthorLen)
                        Else
                            .Author = strEmpty
                        End If
                        .AuthorLen = AuthorLen
            'ВНИМАНИЕ!!! В cMsgMap e-mail ТОЛЬКО ОДИН либо никакой!!! Это используется
            'в ParserLevel2. Сбор информации об e-mail производится уже в cAuthors, куда
            'передаются сведения, в том числе и отсюда
                        '.email = email
                        If emailLen > 0 Then
                            .email = Left(email, emailLen)
                        Else
                            .email = strEmpty
                        End If
                        .emailLen = emailLen
            'msgHead используется для отлова ошибок индексации сервера book.ru
            'Заголовок приходится резать дважды и сравнивать
                        '.msgHead = msgHead
                        If msgHeadLen > 0 Then
                            .msgHead = msgHead
                        Else
                            .msgHead = strEmpty
                        End If
                        .msgHeadLen = msgHeadLen
                        .MsgSize = MsgSize
                        '.msgDate = msgDate
                        If msgDateLen > 0 Then
                            .msgDate = msgDate
                        Else
                            .msgDate = strEmpty
                        End If
                        .msgDateLen = msgDateLen
                        .Indent = Indent
                    End With
                    hFile = FreeFile
                    Open (App.Path & DataPath & MsgFile) For Random As hFile Len = Len(hMessage)
                    Put hFile, RecNum, hMessage
                    'Запись в лог программы
                    LogEvent ("[MsgMap.UpdateDB] Произведено обновление записи номер " + Str(RecNum))
                    Close hFile
                    UpdateDB = True
                End If
            End If
        Case ucHistoryMode, ucUseCurrRec:
            Dim j As Long, k As Long    'Дополнительные счетчики
            Dim hHFile As Long  'Указатель на файл журнала
            Dim hTMPFile As Long    'Указатель на временный файл
            Dim lElements() As Long     'Указатель на записи в файле
            Dim BestPos As Long     'Текущий кандидат на перемещение
            Dim tmpstr As String
            Dim lUpperBound As Long     'Макс. значение счетчика
            
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] режим ucHistoryMode or ucUseCurrRec")
            'Если закачка была остановлена до ее окончания, то используем
            'не LastRecNum, а CurrRec-1 !!!
            If Mode = ucHistoryMode Then
                lUpperBound = LastRecNum
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] режим ucHistoryMode, количество записей = " + Str(lUpperBound))
            Else
                If CurrRec > 1 Then
                    lUpperBound = CurrRec - 1
                Else
                    lUpperBound = -1 'Нечего загружать! Отмена...
                End If
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] режим ucUseCurrRec, количество записей = " + Str(lUpperBound))
            End If
            'Используем тот факт, что дата есть во всех сообщениях. Поэтому, если
            'дата присутствует, то ссылка загружена. Если msgDataLen = 0, то эта
            'ссылка не загружалась.
            If CheckFile(ucMsgURLsMode) And lUpperBound > 0 Then
                'Открываем БД ссылок
                hFile = FreeFile
                Open (App.Path & DataPath & MsgFile) For Random As hFile Len = Len(hMessage)
                'Выделяем память
                ReDim lElements(lUpperBound, 1) 'Нулевые элементы будут использованы в
                                                'качестве временных переменных
                frmBrowser.Label14.Caption = "Чтение информации о новых поступлениях..."
                frmBrowser.Label14.Refresh
                For i = 1 To lUpperBound
                    Get hFile, i, hMessage   'Читаем запись hMessage
                    If hMessage.msgDateLen > 0 Then
                        'Читаем номер ссылки...
                        tmpstr = strEmpty
                        For j = 1 To hMessage.msgURLLen
                            If Mid(hMessage.msgURL, j, 1) Like "[0-9]" Then
                                tmpstr = tmpstr & Mid(hMessage.msgURL, j, 1)
                            End If
                        Next j
''TO DO Здесь! И через 2 строки заменил 0 на i !!!
                        lElements(i, 0) = CLng(Val(tmpstr))
'LEGACY lElements(0, 0) = CLng(Val(Right(Left(hMessage.msgURL, hMessage.msgURLLen), 9)))
                        '...и запоминаем порядковый номер записи в БД
                        lElements(i, 1) = i
                    End If
                Next i
'-----------------------------------------------------------------------------------
'Альтернатива: тратить каждый раз при добавлении записи в журнал время на сортировку
'              или добавлять записи в конец журнала + указатели на соседние, что вызовет
'              небольшое замедление работы при последующем чтении БД
'Выбор:        сортировка при добавлении записи
'Причина:      журнал может иметь до 2000 записей и, затратив время на сортировку,
'              получим выигрыш на последующих актах записи/чтения за счет простоты.
'Алгоритм:    -чтение заголовков загруженных записей в память
'             -сортировка
'             -обновление журнала
'-----------------------------------------------------------------------------------
                'Сортировка по убыванию (метод пузырька)
'                lElements(0, 0) - временная переменная - номер в URL
'                lElements(0, 1) - временная переменная - номер этой ссылки в журнале
''''                lElements(0, 0) = lElements(1, 0)   'Принимаем для начала первый элемент
''''                lElements(0, 1) = lElements(1, 1)   'и запоминаем его номер в БД ссылок
                frmBrowser.Label14.Caption = "Сортировка новых поступлений..."
                frmBrowser.Label14.Refresh
                For i = 1 To lUpperBound
                    BestPos = i
                    For j = i To lUpperBound
                        If lElements(j, 0) > lElements(BestPos, 0) Then  'Если эл-ты равны, то не трогаем их
                            BestPos = j
                        End If
                    Next j
                    lElements(0, 0) = lElements(BestPos, 0)
                    lElements(0, 1) = lElements(BestPos, 1)
                    lElements(BestPos, 0) = lElements(i, 0)
                    lElements(BestPos, 1) = lElements(i, 1)
                    lElements(i, 0) = lElements(0, 0)
                    lElements(i, 1) = lElements(0, 1)
                Next i
                'Запоминаем общее число записей в нулевом элементе lElements(0,0)
                'и порядковый номер текущей записи в lElements(0,1)
                lElements(0, 0) = lUpperBound
                lElements(0, 1) = 1
                'Попробуем отсечь ошибки закачки (URL=0)
                For i = lUpperBound To 1 Step -1
                    If lElements(i, 0) = 0 Then
                        'Уменьшаем счетчик закачанных страниц
                        lElements(0, 0) = lElements(0, 0) - 1
                    Else
                        'Больше ошибок нет
                        i = 1
                    End If
                Next i
                'Сортировка завершена. Обновление журнала
                'Открываем журнал
                hHFile = FreeFile
                Open (App.Path & DataPath & HistoryFile) For Random As hHFile Len = Len(hHistory)
                'Определяем количество записей в журнале
                j = LOF(hHFile) \ Len(hHistory)
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Файл журнала открыт, количество записей = " + Str(j))
                If j > 0 Then   'Журнал уже существует, нужно добавлять записи
                    Dim hTMPrec As tHistory
                    'Открываем временный файл
                    hTMPFile = FreeFile
                    Open (App.Path & DataPath & TMPfile) For Random As hTMPFile Len = Len(hHistory)
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Временный файл открыт")
                    'Стартуем с конца списка, для этого сделаем присвоение:
                    lElements(0, 1) = lElements(0, 0)   '=lUpperBound
                    'Хреначим подряд записи из журнала, попутно проверяя, не нужно ли вставить
                    'текущую (запись с индексом в массиве lElements(0,1)) в журнал.
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Входим в основной цикл обновления журнала")
                    For i = 1 To j
                        frmBrowser.Label14.Caption = "Обработано " & Str(i) & "записей из " & Str(j)
                        frmBrowser.Label14.Refresh
                        'Не забываем, что запись ведется по ВОЗРАСТАНИЮ URL
                        Get hHFile, i, hTMPrec     'Читаем очередную запись из журнала
                            'Остались еще не добавленные записи?
                            If lElements(0, 1) > 0 Then 'Проверка нужна, чтобы обходить большой кусок текста
                                                        'при отсутствии текущих записей
                                'Да, остались. Нужно проверить, а не текущую ли запись нужно вставить и есть
                                'ли еще что вставлять, учитывая, что проверка >0 на предыдущей строчке
                                'вне цикла недоступна => нужна еще одна такая же.
                                While (hTMPrec.URLvalue > lElements(lElements(0, 1), 0)) And (lElements(0, 1) > 0)
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] URL в журнале: " + Str(hTMPrec.URLvalue) + " URL текущий: " + Str(lElements(lElements(0, 1), 0)))
                                    'URL текущей записи меньше! Нужно вставлять.
                                    Get hFile, lElements(lElements(0, 1), 1), hMessage
                                    'Получаем номер автора в БД авторов
                                    hHistory.AuthorID = Authors.SearchDB(ucSearchByName, hMessage.Author, hMessage.AuthorLen)
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] ID текущего автора =" + Str(hHistory.AuthorID))
                                    If hHistory.AuthorID > 0 Then
                                        'Запоминаем 9 цифр: это номер ссылки текущей записи
                                        hHistory.URLvalue = lElements(lElements(0, 1), 0)
                                        'Заполняем размер сообщения и номер ветки
                                        hHistory.MsgSize = hMessage.MsgSize
                                        hHistory.TopicNum = hMessage.TopicNum
                                        hHistory.Reserve = 0   'Резервное поле
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Запись hHistory.URLvalue=" + Str(hHistory.URLvalue))
                                        'Запись идет ТОЛЬКО если автор зарегестрирован в БД!!!
                                        Put hTMPFile, , hHistory
                                    End If
                                    'Зарегестрирован автор в БД или нет - эту запись уже обработали.
                                    'Во второй раз ее обрабатывать нельзя в любом случае :)
                                    'Продвигаемся на шаг к началу списка
                                    lElements(0, 1) = lElements(0, 1) - 1
                                Wend
                            End If
                            'Необязательный блок пройден. Запись hTMPrec исполняется в любом случае.
                        'Рассмотрим случай когда URL равны
                        If lElements(0, 1) > 0 Then
                            If lElements(lElements(0, 1), 0) = hTMPrec.URLvalue Then
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Равенство URL, пропуск записи текущего элемента" + Str(hTMPrec.URLvalue) + "=" + Str(lElements(lElements(0, 1), 0)))
                                'В случае равенства просто пропустим текущую запись
                                'и поднимемся на шаг к началу списка без обработки
                                lElements(0, 1) = lElements(0, 1) - 1
                            End If
                        End If
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Запись hTMPrec.URLvalue=" + Str(hTMPrec.URLvalue))
                        Put hTMPFile, , hTMPrec
                        'Даем процессорное время других приложениям
                        DoEvents
                    Next i
                    'Журнал прочитан до конца, но список текущих записей может еще содержать
                    'необработанные элементы. Все аналогично предыдущему циклу!
                    While lElements(0, 1) > 0
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Журнал обработан, но остались еще текущие записи в количестве" + Str(lElements(0, 1)))
                        Get hFile, lElements(lElements(0, 1), 1), hMessage
                        'Получаем номер автора в БД авторов
                        hHistory.AuthorID = Authors.SearchDB(ucSearchByName, hMessage.Author, hMessage.AuthorLen)
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] ID текущего автора =" + Str(hHistory.AuthorID))
                        If hHistory.AuthorID > 0 Then
                            'Запоминаем 9 цифр: это номер ссылки текущей записи
                            hHistory.URLvalue = lElements(lElements(0, 1), 0)
                            'Заполняем размер сообщения и номер ветки
                            hHistory.MsgSize = hMessage.MsgSize
                            hHistory.TopicNum = hMessage.TopicNum
                            hHistory.Reserve = 0   'Резервное поле
                            'Запись идет ТОЛЬКО если автор зарегестрирован в БД!!!
                            Put hTMPFile, , hHistory
                        End If
                        'Продвигаемся на шаг к началу списка
                        lElements(0, 1) = lElements(0, 1) - 1
                    Wend
                    Close hTMPFile
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Временный файл закрыт")
                    Close hHFile
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Файл журнала закрыт")
                    'Стираем старый файл журнала
                    Kill (App.Path & DataPath & HistoryFile)
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Файл журнала уничтожен")
                    'Переименовываем временный файл
                    Name (App.Path & DataPath & TMPfile) As (App.Path & DataPath & HistoryFile)
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Временный файл переименован")
                            
                    
''''''''''                    For i = 1 To j
''''''''''                        Get hHFile, i, hTMPrec     'Читаем очередную запись из журнала
''''''''''''TO DO BUG subscript out of range
''''''''''''lElemets(0, 1)=10 при 9 успешно закачанных записей. Должно быть 9 или 8
''''''''''                        While (hTMPrec.URLvalue > lElements(lElements(0, 1), 0)) And (lElements(0, 1) <= lElements(0, 0))
''''''''''                            'Пока есть еще незаписанные элементы с номером ссылки,
''''''''''                            'меньшей, чем в текущей записи журнала, записываем элементы
''''''''''                            'НО только в том случае, если есть чего писать :), т.е.
''''''''''                            'не все элементы еще перебраны
''''''''''                            If lElements(0, 1) <= lElements(0, 0) Then
''''''''''                                Get hFile, lElements(lElements(0, 1), 1), hMessage
''''''''''                                'Получаем номер автора в БД авторов
''''''''''                                hHistory.AuthorID = Authors.SearchDB(ucSearchByName, Left(hMessage.Author, hMessage.AuthorLen))
''''''''''                                If hHistory.AuthorID > 0 Then
''''''''''                                    'Ищем 9 цифр: это номер ссылки
''''''''''                        'TO DO Это ведь уже сделано: см текущий элемент lElements
''''''''''                                    ''''''tmpstr = strEmpty
''''''''''                                    ''''''For k = 1 To hMessage.msgURLLen
''''''''''                                    ''''''    If Mid(hMessage.msgURL, k, 1) Like "[0-9]" Then
''''''''''                                    ''''''        tmpstr = tmpstr & Mid(hMessage.msgURL, k, 1)
''''''''''                                    ''''''    End If
''''''''''                                    ''''''Next k
''''''''''                                    ''''''hHistory.URLvalue = CLng(Val(tmpstr))
''''''''''                                    hHistory.URLvalue = lElements(lElements(0, 1), 0)
'''''''''''LEGACY hHistory.URLvalue = CLng(Val(Right(Left(hMessage.msgURL, hMessage.msgURLLen), 9)))
''''''''''                                    'Заполняем размер сообщения и номер ветки
''''''''''                                    hHistory.MsgSize = hMessage.MsgSize
''''''''''                                    hHistory.TopicNum = hMessage.TopicNum
''''''''''                                    hHistory.Reserve = 0   'Резервное поле
''''''''''                                    'Запись идет ТОЛЬКО если автор зарегестрирован в БД!!!
''''''''''                                    Put hTMPFile, , hHistory
''''''''''                                End If
''''''''''                                'Сдвигаем указатель необработанных ссылок к следующему элементу
''''''''''                                    If lElements(0, 1) < lElements(0, 0) Then
''''''''''                                        lElements(0, 1) = lElements(0, 1) + 1
''''''''''                                    End If
''''''''''                            End If
''''''''''                        Wend
''''''''''                        'Все новые ссылки обработаны, записываем старую
''''''''''                        Put hTMPFile, , hTMPrec
''''''''''                    Next i
''''''''''                    Close hTMPFile
''''''''''                    Close hHFile
''''''''''                    'Стираем старый файл журнала
''''''''''                    Kill (App.Path & DataPath & HistoryFile)
''''''''''                    'Переименовываем временный файл
''''''''''                    Name (App.Path & DataPath & TMPfile) As (App.Path & DataPath & HistoryFile)
                Else
                    'Журнала еще нет - пишем все подряд в обратном порядке, поскольку при сортировке
                    'самые маленькие числа (ранние постинги) оказались в конце списка
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Начало записи в журнал в обратном порядке")
                    For i = lElements(0, 0) To 1 Step -1
                        'Читаем из БД ссылок запись с номером lElements(i,1)
'TO DO здесь постоянно Bad Record Number при lEl(0,0)=8 (or 111 - num of pages downloaded), i=1, lEl(i,1)=0
'Пробую вылечить наверху над большим комментарием
                        
                        Get hFile, lElements(i, 1), hMessage
                        ''Get hFile, i, hMessage
                        'Получаем номер автора в БД авторов
                        hHistory.AuthorID = Authors.SearchDB(ucSearchByName, hMessage.Author, hMessage.AuthorLen)
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] ID текущего автора =" + Str(hHistory.AuthorID))
                        If hHistory.AuthorID > 0 Then
                            'Ищем 9 цифр: это номер ссылки
                    'TO DO Это ведь уже сделано: см текущий элемент lElements
                            ''''''tmpstr = strEmpty
                            ''''''For k = 1 To hMessage.msgURLLen
                            ''''''    If Mid(hMessage.msgURL, k, 1) Like "[0-9]" Then
                            ''''''        tmpstr = tmpstr & Mid(hMessage.msgURL, k, 1)
                            ''''''    End If
                            ''''''Next k
                            ''''''hHistory.URLvalue = CLng(Val(tmpstr))
                            hHistory.URLvalue = lElements(i, 0)
'LEGACY hHistory.URLvalue = CLng(Val(Right(Left(hMessage.msgURL, hMessage.msgURLLen), 9)))
                            'Заполняем размер сообщения и номер ветки
                            hHistory.MsgSize = hMessage.MsgSize
                            hHistory.TopicNum = hMessage.TopicNum
                            hHistory.Reserve = 0   'Резервное поле
                            'Запись идет ТОЛЬКО если автор зарегестрирован в БД!!!
                            Put hHFile, , hHistory
                        End If
                    Next i
                    Close hHFile
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Файл журнала закрыт")
                End If
                ReDim lElements(0, 0) 'Возвращаем память системе
                Close hFile
'Запись в лог программы
LogEvent ("[MsgMap.UpdateDB] Файл ссылок закрыт")
                UpdateDB = True
            Else
                Result = MsgBox("База данных ссылок отсутствует. БД авторов не обновлена", vbOKOnly, "БД не найдена")
            End If
    End Select
    Exit Function

UpdateDBErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::UpdateDB Method")
End Function

'**********************************************************
'*  Считывает очередную запись из базы данных по ссылкам  *
'*  GetNext = номер записи если есть что читать           *
'*  GetNext = 0 если записи закончились                   *
'**********************************************************
Public Function GetNext() As Long
Dim hFile As Long
    On Error GoTo GetNextErr
        
    CurrRec = CurrRec + 1
    GetNext = 0
    If Not (CurrRec > LastRecNum) Then
        'Проверка существования прошла автоматически! (LastRecNum не ноль)
        hFile = FreeFile
        Open (App.Path & DataPath & MsgFile) For Random As hFile Len = Len(hMessage)
        Get hFile, CurrRec, hMessage
        Close hFile
        GetNext = CurrRec
        With hMessage
            TopicNum = .TopicNum
            msgURLLen = .msgURLLen
            'msgURL = .msgURL
            If msgURLLen > 0 Then
                msgURL = Left(.msgURL, msgURLLen)
            Else
                msgURL = strEmpty
            End If
            AuthorLen = .AuthorLen
            'Author = .Author
            If AuthorLen > 0 Then
                Author = Left(.Author, AuthorLen)
            Else
                Author = strEmpty
            End If
            emailLen = .emailLen
            'email = .email
            If emailLen > 0 Then
                email = Left(.email, emailLen)
            Else
                email = strEmpty
            End If
            msgHeadLen = .msgHeadLen
            'msgHead = .msgHead
            If msgHeadLen > 0 Then
                msgHead = Left(.msgHead, msgHeadLen)
            Else
                msgHead = strEmpty
            End If
            MsgSize = .MsgSize
            msgDateLen = .msgDateLen
            'msgDate = .msgDate
            If msgDateLen > 0 Then
                msgDate = Left(.msgDate, msgDateLen)
            Else
                msgDate = strEmpty
            End If
            Indent = .Indent
        End With
    End If
    Exit Function

GetNextErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::GetNext Method")
End Function

'**********************************************************
'*  Считывает запись с номером ID из базы данных          *
'*  GetMsg = номер записи если есть что читать            *
'*  GetMsg = 0 в любом другом случае                      *
'**********************************************************
Public Function GetMsg(ByVal ID As Long) As Long
Dim hFile As Long
    On Error GoTo GetMsgErr
        
    GetMsg = 0
    If Not (ID > LastRecNum) Then
        'Особенность компилятора VB: в OR ВСЕГДА проверяются оба условия, так что разбиваем на две части
        If Not (ID < 1) Then
            'Проверка существования прошла автоматически! (LastRecNum не ноль)
            hFile = FreeFile
            Open (App.Path & DataPath & MsgFile) For Random Shared As hFile Len = Len(hMessage)
            Get hFile, ID, hMessage
            Close hFile
            GetMsg = ID
            With hMessage
                TopicNum = .TopicNum
                msgURLLen = .msgURLLen
                'msgURL = .msgURL
                If msgURLLen > 0 Then
                    msgURL = Left(.msgURL, msgURLLen)
                Else
                    msgURL = strEmpty
                End If
                AuthorLen = .AuthorLen
                'Author = .Author
                If AuthorLen > 0 Then
                    Author = Left(.Author, AuthorLen)
                Else
                    Author = strEmpty
                End If
                emailLen = .emailLen
                'email = .email
                If emailLen > 0 Then
                    email = Left(.email, emailLen)
                Else
                    email = strEmpty
                End If
                msgHeadLen = .msgHeadLen
                'msgHead = .msgHead
                If msgHeadLen > 0 Then
                    msgHead = Left(.msgHead, msgHeadLen)
                Else
                    msgHead = strEmpty
                End If
                MsgSize = .MsgSize
                msgDateLen = .msgDateLen
                'msgDate = .msgDate
                If msgDateLen > 0 Then
                    msgDate = Left(.msgDate, msgDateLen)
                Else
                    msgDate = strEmpty
                End If
                Indent = .Indent
            End With
        End If
    End If
    Exit Function

GetMsgErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::GetMsg Method")
End Function

'**********************************************************
'*         Проверить существование базы данных            *
'*  Mode=0 - работа с базой данных ссылок                 *
'*  Mode=1 - работа с журналом                            *
'**********************************************************
Private Function CheckFile(ByVal Mode As enFileMode) As Boolean
Dim hFile As Long
    On Error GoTo CheckFileErr
    
    ''CheckFile = False
    Select Case Mode
        Case ucMsgURLsMode:
            CheckFile = IIf(Dir(App.Path & DataPath & MsgFile, vbNormal) <> "", True, False)
            ''If (Dir(App.Path & DataPath & MsgFile, vbNormal) <> "") Then
            ''    'OK, файл существует
            ''    CheckFile = True
            ''End If
        Case ucHistoryMode:
            CheckFile = IIf(Dir(App.Path & DataPath & HistoryFile, vbNormal) <> "", True, False)
    End Select
    Exit Function
    
CheckFileErr:
    CheckFile = False
    Call RaiseError(MyUnhandledError, "cMsgMap::CheckFile Method")
End Function

'**********************************************************
'*      Уничтожает базу данных после ее обработки         *
'**********************************************************
Public Sub Clear()
    On Error GoTo ClearErr
    
    If CheckFile(ucMsgURLsMode) Then
        Kill (App.Path & DataPath & MsgFile)
    End If
    'Сброс указателя на текущую запись на случай закачки следующей страницы
    CurrRec = 0
    Exit Sub
    
ClearErr:
    Call RaiseError(MyUnhandledError, "cMsgMap::Clear Method")
End Sub

Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
    CurrRec = 0
    With hMessage
        .TopicNum = 0
        .msgURL = ""
        .msgURLLen = 0
        .Author = ""
        .AuthorLen = 0
        .email = ""
        .emailLen = 0
        .msgHead = ""
        .msgHeadLen = 0
        .MsgSize = 0
        .msgDate = ""
        .msgDateLen = 0
''        .Indent = 0
        'Начальный отступ теперь не 0, а 1.
        .Indent = 1
    End With
    Clear
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

