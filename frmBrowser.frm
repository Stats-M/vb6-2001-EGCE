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
      Caption         =   "Повтор"
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
      Caption         =   "Стоп"
      Height          =   350
      Left            =   9345
      TabIndex        =   10
      Top             =   630
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Начать"
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
      Caption         =   "Ошибок сервера "
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
      Caption         =   "Ошибок закачки"
      Height          =   225
      Left            =   210
      TabIndex        =   17
      Top             =   1155
      Width           =   2010
   End
   Begin VB.Label Label12 
      Caption         =   "Общая информация:"
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
      Caption         =   "Качалка N1 - статус:"
      Height          =   225
      Left            =   3255
      TabIndex        =   5
      Top             =   210
      Width           =   1590
   End
   Begin VB.Label Label5 
      Caption         =   "Из них обработано"
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
      Caption         =   "Найдено ссылок"
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
      Caption         =   "Закачивается страница"
      Height          =   225
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1905
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Файл"
      Begin VB.Menu mnuWriteAuthors 
         Caption         =   "Записать авторов"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Выход"
      End
   End
   Begin VB.Menu mnuDownload 
      Caption         =   "Закачка"
      Begin VB.Menu mnuConnect 
         Caption         =   "Подключиться"
      End
      Begin VB.Menu mnuHyp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Начать"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Остановить"
      End
      Begin VB.Menu mnuView 
         Caption         =   "Просмотр HTML"
      End
      Begin VB.Menu mnuHyp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Настройки..."
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "Избранное"
      Begin VB.Menu mnuProse 
         Caption         =   "Проза"
         Begin VB.Menu mnuF002 
            Caption         =   "Прыжок"
         End
      End
      Begin VB.Menu mnuOfftop 
         Caption         =   "Оффтопики"
         Begin VB.Menu mnuF001 
            Caption         =   "Завтрак холостяка"
         End
      End
      Begin VB.Menu mnuF003 
         Caption         =   "Создание клуба"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Справка"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Вызов справки"
      End
      Begin VB.Menu mnuHyp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе"
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

Private Const MaxRetryAttempts = 3  'Максимальное число повторных попыток закачки страницы

'Надеюсь, у всех компьютеры не меньше P-II? :) Кто еще не, я не виноват!
'Специально для P-II и круче будем пользоваться только 32-х битными переменными.
Private StartPos As Long        'Начальная позиция для вырезания подстрок
Private EndPos As Long          'Конечная позиция для вырезания подстрок
Private TopicNum As Long        'Номер ветки
Private email As String         'e-mail автора (отключено в публичной версии)
Private Nickname As String      'Ник автора
Private msgHREF As String       'URL собственно сообщения
Private msgHead As String       'Заголовок сообщения
Private Indent As Long          'Отступ сообщения
Private tokenname As String     'Разыскиваемый параметр
Private lResult As Long         'Дополнительная переменная Long
''LEGACY Private FirstTime(2) As Boolean 'Для правильного функционирования Inet
Private strMess As String       'Для общения закачек с остальной программой
Private vtData As Variant       'Рекомендовано Microsoft, для информации от качалки
Private bDone As Boolean        'Флаги работы качалок
Private bCancel As Boolean      'TRUE, если была нажата кнопка Cancel
Private GettingURLs As Boolean  'Указатель для таймера, что закачивается главная страница
Private Loading As Boolean      'При загрузке нужно обойти некоторые функции
Private lRetry As Long           'Повторные запросы
Private HTMLDoc As HTMLDocument

'Переменные для работы Feeder - цвета фона и т.д.
Private COLOR As String         'Цвет фона для HTML-страницы
Private bColor As Boolean       'Выбирает один из цветов фона
Private TopicChanged As Boolean 'Отслеживает изменение номера нити обсуждения
Private LastTopicNum As Long    'Номер темы последнего сообщения (начинать новую таблицу или нет в HTML)

Private Sub Test()
Dim HistoryLen As Long       'Количество записей в журнале
Dim hHFile As Long  'Указатель на файл журнала
Dim hFile As Long   'Указатель на БД авторов
Dim hAuthor As tAuthor
Dim hHistory As tHistory
Dim i As Long

    hHFile = FreeFile
    Open (App.Path & DataPath & HistoryFile) For Random As hHFile Len = Len(hHistory)
    hFile = FreeFile
    Open (App.Path & DataPath & AuthorsFile) For Random As hFile Len = Len(hAuthor)
    'Определяем количество записей в журнале
    HistoryLen = LOF(hHFile) \ Len(hHistory)
    Label10.Caption = Str(HistoryLen)
''    Text1.Text = strEmpty
''    If HistoryLen > 0 Then   'Журнал существует
''        For i = 1 To HistoryLen
''            'Читаем очередную запись в журнале
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
'*      Режем и кромсаем загруженную страницу        *
'* на отдельные нитки и передаем их дальше           *
'*****************************************************
Private Function Splitter(ByRef lpString As String) As Boolean
Dim tmpstr As String
    Splitter = False
    'Отрубаем таймер, дабы не мешал!
    Timer1.Enabled = False
    Label14.Caption = "Идет анализ страницы-оглавления"
    Label14.Refresh
    'Убедимся, что нам подсунули то, что надо и в тексте есть <dl>
    StartPos = InStr(1, lpString, "<DL>", vbTextCompare)
    If StartPos = 0 Then
        'Нас обманули!!!
        'TO DO smth
        Label4.Caption = "Ошибка закачки оглавления"
        Label4.Refresh
        Label14.Caption = "Ошибка закачки оглавления"
        Label14.Refresh
        If InStr(1, lpString, "The requested URL could not be retrieved", vbTextCompare) > 0 Then
            Label14.Caption = "Запрашиваемая страница не может быть получена"
            Label14.Refresh
        End If
        Exit Function
    End If
    'Автоопределение базового адреса! Иначе после каждого изменения на сервере нужно будет перебирать всю программу.
    'Начало зоны поиска (ищем слово "Добавить")
    lResult = InStr(1, lpString, "Добавить", vbTextCompare)
    'Конец зоны поиска (ищем слово "Читать")
    EndPos = InStr(lResult, lpString, "Читать", vbTextCompare)
    'Копируем регион
    BaseHREF = Mid(lpString, lResult, EndPos - lResult)
    'Ближе к телу :) Начало ссылки:
    ''lResult = InStr(1, BaseHREF, "href=`", vbTextCompare)book.cgi?book=
    lResult = InStr(1, BaseHREF, "book.cgi?book=", vbTextCompare)
    'Ищем знак равенства в ссылке (book=Elitegames...) и
    'запоминаем СЛЕДУЮЩУЮ позицию (знак равенства должен остаться в ссылке)
    EndPos = InStr(lResult, BaseHREF, "=", vbTextCompare) + 1
    'Обрезаем и получаем наш базовый адрес
'TO DO добавить анализ ссылки: нельзя же просто book.cgi отсылать без указания сервера!!!
    BaseHREF = Replace(Mid(BaseHREF, lResult, EndPos - lResult), "`", strEmpty, 1, -1, vbBinaryCompare)
    
    'Продолжаем обработку страницы: отсекаем все лишнее
    lpString = Right(lpString, Len(lpString) - StartPos)
    Label2.Caption = Str(CurrPage)
    'Очистка от мусора уже была в NavigateComplete2
    
    'Ищем номер нитки
    StartPos = InStr(1, lpString, "<p>", vbTextCompare) + Len("<p>") 'Сразу проскакиваем тег
    EndPos = InStr(StartPos, lpString, ".", vbTextCompare)
    'Не будем рисковать: две строки вне цикла погоды не сделают, а проверка
    'в его начале гораздо надежнее, чем в конце
    Do While EndPos - StartPos < 8
        'Пока не закончатся нитки (интересно, хватит 5-значного номера для номеров?)
        '8 = <p> + XXXXX
        If EndPos = 0 Then
            '<p> был, но это не номер нитки... Может такое быть?
            'TO DO smth
            Exit Do
        End If
        TopicNum = CLng(Val(Mid(lpString, StartPos, EndPos - StartPos)))
        'Находим конец нитки
        EndPos = InStr(EndPos, lpString, "<p>", vbTextCompare)
        If EndPos = 0 Then
            'Очередной <p> не найден. Маловероятно, этот тег всегда есть в конце страницы этой конференции
            'TO DO smth
            Exit Do
        End If
        'Отрезаем нитку с лидирующим тегом <p>...
        tmpstr = Mid(lpString, StartPos - 3, EndPos)
        '... и удаляем ее из основной строки: нефиг таскать такие объемы туда-сюда по функциям
        'Дополнительный символ нужен для гарантии безошибочной работы в дальнейшем
        lpString = Right(lpString, Len(lpString) - EndPos + 1)
        'Передаем нитку дальше на растерзание
        If ParserLevel1(TopicNum, tmpstr) Then
            'TO DO
        End If
        'Ищем номер следующей нитки
        StartPos = InStr(1, lpString, "<p>", vbTextCompare) + Len("<p>") 'Сразу проскакиваем тег
        EndPos = InStr(StartPos, lpString, ".", vbTextCompare)
    Loop
    'Включаем таймер уже в InitFeeder!
    Label14.Caption = "Страница-оглавление проанализирована"
    Label14.Refresh
    Splitter = True
End Function

'******************************************************
'*      Разматываем нитку на отдельные топики         *
'* На выходе имеем ссылки на каждое сообщение нитки   *
'* и записываем их в MsgMap (там сами разберутся)     *
'*      TopicNum - номер темы (для записи в БД ссылок)*
'******************************************************
Private Function ParserLevel1(ByVal TopicNum As Long, ByVal lpString As String) As Boolean
Dim StartPos As Long
Dim EndPos As Long
    '''EndPos = InStr(1, lpString, "A href=`http://book.by.ru/cgi-bin/book.cgi?book=", vbTextCompare)
    ''EndPos = InStr(1, lpString, "a href=`book.cgi?book=", vbTextCompare)
    'EndPos = InStr(1, lpString, "book.cgi?book=", vbTextCompare)
    EndPos = InStr(1, lpString, BaseHREF, vbTextCompare)
    Debug.Print "ВХод в ParserLevel1"
''    Indent = 0 'Первое сообщение ветки, отступ = 0
    'Дело в том, что при Indent = 0 первый тег <DL> пишется только начиная с
    'первого ответа на вопрос, что препятствует появлению отступа для самого
    'первого ответа. Для устранения этого недостатка начнем отступ с 1.
    Indent = 1 'Первое сообщение ветки, отступ = 1
    Do While EndPos <> 0
        'Определяем e-mail автора
        email = strEmpty
        'Ищем ближайшую ссылку на e-mail
        StartPos = InStr(1, lpString, "<a href=`mailto:", vbTextCompare)
        'Запомним - потом пригодится, т.к. до адреса еще будем e-mail и ником заниматься
        '''lResult = EndPos + Len("A href=`http://book.by.ru/cgi-bin/book.cgi?book=")
        ''lResult = EndPos + Len("a href=`book.cgi?book=")
        'lResult = EndPos + Len("book.cgi?book=")
        lResult = EndPos + Len(BaseHREF)
        If EndPos > StartPos Then
            'Очередность появления ссылок обычная (сначала e-mail, затем соообщение)
            If Not (StartPos = 0) Then
                'Есть e-mail автора
                StartPos = StartPos + Len("<A href=`mailto:")
                EndPos = InStr(StartPos, lpString, "`>", vbTextCompare)
                email = Mid(lpString, StartPos, EndPos - StartPos)
            End If  'в противном случае e-mail автора не указан
        End If  'в противном случае e-mail найден, но он относится уже к следующему
                'сообщению -> игнорируем его
        'Определяем имя автора
        Nickname = strEmpty
        StartPos = InStr(1, lpString, "<B><I>", vbTextCompare) + Len("<B><I>")
        EndPos = InStr(StartPos, lpString, "</I></B>", vbTextCompare)
        If Not (StartPos >= EndPos) Then
            'Имя автора указано (мало ли, человек забыл/не захотел указать имя 8-) )
            Nickname = Mid(lpString, StartPos, EndPos - StartPos)
            Nickname = RemoveSpaces(Nickname)
        End If
        'Определяем URL сообщения
        msgHREF = strEmpty
        EndPos = InStr(lResult, lpString, "`>", vbTextCompare)
        'Режем адрес, очищая его от &amp; , оставляя только &
        msgHREF = Replace(Mid(lpString, lResult, EndPos - lResult), "amp;", strEmpty, 1, -1, vbBinaryCompare)
        'Ответственное место! Вырезаем тему для последующего сравнения.
        'Только так можно отловить глюки индексации сервера Book.ru
        StartPos = EndPos + 2   'Проезжаем "`>"
        EndPos = InStr(StartPos, lpString, "</A>", vbTextCompare)
        msgHead = Mid(lpString, StartPos, EndPos - StartPos)
'Внимание! Теперь надо вырезать (если есть) комбинацию "(-)",
'которую сервер автоматом шлепает к заголовкам, ссылающимся на пустую тему.
'Также вырежем "(+)"
'Используем переменные lResult и StartPos как временно не используемые
        msgHead = RemoveSpaces(msgHead)
        StartPos = InStr(1, msgHead, "(-)", vbTextCompare)
        Do While StartPos > 0
        ''If StartPos > 0 Then
            '(-) найден. Проверим (грубо), находится ли он в хвосте...
            If StartPos > Len(msgHead) - 5 Then
                '(-) в конце строки - просто отрежем хвост
                msgHead = Left(msgHead, Len(msgHead) - 3)
            Else
                '(-) в середине. Вырежем только эти 3 символа
                msgHead = Left(msgHead, StartPos - 1) + Right(msgHead, Len(msgHead) - StartPos - 2)
            End If
        ''End If
            'Убираем пробелы в начале и конце строки
            msgHead = RemoveSpaces(msgHead)
            StartPos = InStr(1, msgHead, "(-)", vbTextCompare)
        Loop
        StartPos = InStr(1, msgHead, "(+)", vbTextCompare)
        Do While StartPos > 0
        ''If StartPos > 0 Then
            '(+) найден. Проверим (грубо), находится ли он в хвосте...
            If StartPos > Len(msgHead) - 5 Then
                '(+) в конце строки - просто отрежем хвост
                msgHead = Left(msgHead, Len(msgHead) - 3)
            Else
                '(+) в середине. Вырежем только эти 3 символа
                msgHead = Left(msgHead, StartPos - 1) + Right(msgHead, Len(msgHead) - StartPos - 2)
            End If
        ''End If
            'Убираем пробелы в начале и конце строки
            msgHead = RemoveSpaces(msgHead)
            StartPos = InStr(1, msgHead, "(+)", vbTextCompare)
        Loop
'Запись в лог программы
LogEvent ("[Parser Level 1] Заголовок темы: " + msgHead)
'{Fixed} TO DO Что за дерьмо? msgHead = RemoveSpaces(msgHead)
'Функция теперь возвращает результат, все должно быть ОК
        'Дату и текст сообщения будем получать уже в PareserLevel2
        'отрезаем обработанную часть
        lpString = Right(lpString, Len(lpString) - EndPos)
        'Заполняем запись
'Запись в лог программы
LogEvent ("[Parser Level 1] Начало сбора информации для автора " + Nickname)
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
        'Сохраняем в БД и получаем номер записи
        Label4.Caption = Str(MsgMap.Save)
        Label4.Refresh
'Запись в лог программы
LogEvent ("[Parser Level 1] Тема сохранена в файле ссылок под номером " + Label4.Caption)
        'Ищем ближайшую ссылку на сообщение (инициализация цикла)+поиск <dl> и </dl>
        '''EndPos = InStr(1, lpString, "A href=`http://book.by.ru/cgi-bin/book.cgi?book=", vbTextCompare)
        ''EndPos = InStr(1, lpString, "a href=`book.cgi?book=", vbTextCompare)
        'EndPos = InStr(1, lpString, "book.cgi?book=", vbTextCompare)
        EndPos = InStr(1, lpString, BaseHREF, vbTextCompare)
        'Определим отступ следующего сообщения. ВНИМАНИЕ! EndPos НУЖНО СОХРАНИТЬ!!!
        'Режем кусок от начала и до EndPos, если он не ноль.
        'В куске шарим <dl> или </dl>
        If Not (EndPos = 0) Then
            'Будем использовать переменную email в качестве временной
            'Сейчас адрес автора уже не нужен, а ее значение обнулится в начале цикла
            email = Left(lpString, EndPos)
            'Рассмотрим этот кусок текста внимательнее
            'Одновременно <dl> и </dl> встречаться не могут, поэтому идем "в лоб"
            lResult = InStr(1, email, "</dl>", vbTextCompare)
            Do While lResult > 0
                'Пока есть </dl>, вырезаем их по одному и уменьшаем счетчик
                email = Right(email, Len(email) - lResult - Len("</dl>"))
                Indent = Indent - 1
                lResult = InStr(1, email, "</dl>", vbTextCompare)
            Loop
            lResult = InStr(1, email, "<dl>", vbTextCompare)
            Do While lResult > 0
                'Пока есть <dl>, вырезаем их по одному и накручиваем счетчик
                email = Right(email, Len(email) - lResult - Len("<dl>"))
                Indent = Indent + 1
                lResult = InStr(1, email, "<dl>", vbTextCompare)
            Loop
            'Проверку вынесли за цикл чтобы постоянно не проверять
''            If Indent < 0 Then
''                Indent = 0
''            End If
            'Изменение начального отступа. Вместо 0 теперь будет 1.
            If Indent < 1 Then
                Indent = 1
            End If
        End If
    Loop
End Function

'******************************************************
'*      Подготавливает переменные для Feeder и        *
'*              стартует HTML-страницу                *
'* Также загружает качалку стартовым адресом          *
'******************************************************
Private Function InitFeeder() As Boolean
    Label14.Caption = "Подготовка к закачке ссылок"
    Label14.Refresh
    lRetry = 0 'Сколько попыток закачки произведено
    'Начало работы - формально соответствует новому номеру нити обсуждения
    TopicChanged = True
    bDone = False
    strMess = strEmpty
    'Начало страницы будет на белом фоне
    bColor = True
    COLOR = ColorLight
    'Создаем HTML-файл, предварительно стерев (если есть) существующий
    EraseHTML (CurrPage)
    bResult = Module1.WriteHTML(0, CurrPage, COLOR, strEmpty, RecNum)  'Шапка
    'Следующая строка нужна для нейтрализации закрытия темы в Feeder
    bResult = Module1.WriteHTML(5, CurrPage, COLOR, strEmpty, RecNum)  'Старт таблицы
    bResult = Module1.WriteNavigator(0, CurrPage, RecNum, Indent, strEmpty) 'Старт навигатора и локального индекса
    If MsgMap.LastRecNum > 0 Then
        'Запомним номер записи в БД ссылок. Позже (в ParserLevel2) эту запись
        'нужно будет обновить, тогда это значение и пригодится
        RecNum = MsgMap.GetNext
        If Not (RecNum = 0) Then
            'ОК, запись прочитана, запускаем качалку
            'Запоминаем все известные параметры из MsgMap пока их не перезаписали...
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
            'Запомним номер темы самого первого сообщения
            LastTopicNum = hMsg.TopicNum
            'Загружаем качалки работой без ожидания ее выполнения (это будет далее)
            InitFeeder = True
            Call StartInet(Left(hMsg.msgURL, hMsg.msgURLLen))
        End If
    Else
        InitFeeder = False
    End If
    Label14.Caption = "Идет закачка сообщений по ссылкам"
    Label14.Refresh
    'Включаем таймер
    Timer1.Enabled = True
End Function

'*****************************************************
'* Скармливает массив ссылок качалке                 *
'* Окончание работы означает завершение              *
'* обработки очередной страницы                      *
'*      Feeder = 0 успешное завершение               *
'*      Feeder <>0 номер записи, вызвавшей ошибку    *
'*****************************************************
Private Function Feeder() As Long
Dim tmpString As String
    'Отключаем таймер на время работы
    Timer1.Enabled = False
    'Берем строку от качалки и скармливаем ее ParserLev2, получаем текст сообщения
    tmpString = ParserLevel2(strMess)
    If lRetry > 0 Then
        If lRetry > MaxRetryAttempts Then
            'Страницу закачать не удалось после MaxRetryAttempts попыток
            Label16.Caption = Str(Val(Label16.Caption) + 1)
            Label14.Caption = "Ошибка закачки. Пропуск текущей ссылки"
            Label14.Refresh
            tmpString = "<i>Информация <b>EliteGames conference extractor</b></i>. Невозможно загрузить страницу. Количество попыток загрузки: " & Format(MaxRetryAttempts)
            'Сброс счетчика повторов
            lRetry = 0
        Else
            'Инициализация переменных
            bDone = False
            strMess = strEmpty
            'Повторный запрос страницы
            Label14.Caption = "Повторный запрос страницы по текущей ссылке"
            Label14.Refresh
            Call StartInet(Left(hMsg.msgURL, hMsg.msgURLLen))
            Timer1.Enabled = True
            Exit Function
        End If
    End If
    If LastTopicNum <> hMsg.TopicNum Then
        'Текущее сообщение относится уже к следующей теме
        TopicChanged = True
    End If
    'Запомним номер текущей темы
    LastTopicNum = hMsg.TopicNum
    'Если тема изменилась, то...
    If TopicChanged Then
        'Начало темы принудительно на светлом фоне
        COLOR = ColorLight
        'Нужно отразить факт смены цвета, подняв флаг
        bColor = True
        'закрываем тему в HTML и открываем новую
        bResult = WriteHTML(6, CurrPage, COLOR, strEmpty, RecNum)  'Закрываем таблицу
        bResult = WriteHTML(5, CurrPage, COLOR, strEmpty, RecNum)  'Старт таблицы и темы
        bResult = WriteHTML(1, CurrPage, COLOR, Left(hMsg.msgHead, hMsg.msgHeadLen), RecNum)
        'Отменяем отступ в навигаторе: новая тема
        bResult = WriteNavigator(3, CurrPage, RecNum, 0, strEmpty)
        'Опускаем флаг новой темы, иначе каждая запись будет трактоваться как новая
        TopicChanged = False
    End If
    'Продолжение нити обсуждения
    'Цвет в любом случае правильный
    If Not (hMsg.emailLen = 0) And (FullLogging Or bFullLogging) Then
        'Пишем всю известную информацию
        bResult = WriteHTML(9, CurrPage, COLOR, Left(hMsg.email, hMsg.emailLen), RecNum)
        bResult = WriteHTML(10, CurrPage, COLOR, Left(hMsg.Author, hMsg.AuthorLen), RecNum)
    Else
        'Запись без указания e-mail
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
    'Пишем навигатор
    bResult = WriteNavigator(1, CurrPage, RecNum, hMsg.Indent, Left(hMsg.Author, hMsg.AuthorLen))
    bResult = WriteNavigator(2, CurrPage, RecNum, 0, Left(hMsg.msgHead, hMsg.msgHeadLen))
    'Меняем цвет фона для следующего сообщения
    bColor = Not bColor
    COLOR = IIf(bColor, ColorLight, ColorDark)
    ''Select Case bColor
    ''    Case True:
    ''        Color = ColorLight
    ''    Case False:
    ''        Color = ColorDark
    ''End Select
    'Пытаемся прочитать следующую запись
    RecNum = MsgMap.GetNext
    If RecNum = 0 Then
        Call StopFeeder(ucDLcomplete)
        'Выходим без запуска таймера: работа окончена
        Exit Function
    Else
        'Прочитали запись, запускаем качалку!
        'Запоминаем все известные параметры из MsgMap пока их не перезаписали...
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
        'Инициализация переменных
        bDone = False
        strMess = strEmpty
        'Загружаем качалку работой без ожидания ее выполнения
        Call StartInet(Left(hMsg.msgURL, hMsg.msgURLLen))
    End If
    'Продолжаем работу, если не нажата кнопка СТОП
    If Not bCancel Then
        Timer1.Enabled = True
'Это ведь реализовано в mnuStop? Зачем еще?
''    Else
''        Call StopFeeder
    End If
End Function

'*****************************************************
'* Производит необходимые действия после завершения  *
'* (останова) закачки)                               *
'*      Mode=ucDLcomplete - закачка завершена        *
'*      Mode=ucDLstopped  - закачка остановлена      *
'*****************************************************
Private Sub StopFeeder(ByVal Mode As enDLresult)
    'Общие команды (для всех режимов)
    'Закрываем страницу
    bResult = Module1.WriteHTML(6, CurrPage, COLOR, strEmpty, RecNum)
    bResult = Module1.WriteHTML(7, CurrPage, COLOR, strEmpty, RecNum)
    'Закрываем навигатор
    bResult = WriteNavigator(4, CurrPage, RecNum, 0, strEmpty)
    Select Case Mode
        Case ucDLcomplete:
            'Обновить журнал: успешная закачка
            Label14.Caption = "Идет обновление журнала закачки..."
            Label14.Refresh
            bResult = MsgMap.UpdateDB(ucHistoryMode)
            'Обновить БД авторов
            ''TO DO BUG А надо ли это делать КАЖДЫЙ раз? И так ведь справляемся...
                    ''Label14.Caption = "Идет обновление статистической информации..."
                    ''Label14.Refresh
                    ''bResult = Authors.UpdateDB(ucHistoryMode)
            'Обновление индекса только при полной закачке текущей страницы
            Label14.Caption = "Запись главного индекса..."
            Label14.Refresh
            bResult = UpdateTOC(CurrPage)
        Case ucDLstopped:
            'Обновить журнал: закачка остановлена пользователем
            bResult = MsgMap.UpdateDB(ucUseCurrRec)
    End Select
    'Обновление статистики по авторам
    Label14.Caption = "Запись статистики в HTML..."
    Label14.Refresh
    bResult = Authors.WriteAHTMLs
    'Стереть БД
    MsgMap.Clear
    Label14.Caption = "Ожидание команды."
    Label14.Refresh
    Label8.Caption = "Готово."
    Label8.Refresh
    mnuExit.Enabled = True
    mnuStop.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = False
End Sub

'*****************************************************
'* Принимаем текст страницы с сообщением от качалки  *
'* и производим ее вскрытие                          *
'*      msgNum - номер записи в БД ссылок            *
'*****************************************************
'lpString передаем ByRef (ByVal не работает здесь)
Private Function ParserLevel2(ByRef lpString As String) As String
Dim StartPos As Long
Dim EndPos As Long
Dim tmpstr As String
Dim i As Long
Dim tmpLen As Long      'Для сравнения длин заголовков при проверке индексации

''ОТЛИЧИЯ МЕССАГ ОТ ЗАГОЛОВКОВ ОГЛАВЛЕНИЙ
''1)В мессагах ник закрывается двоеточием
''2)Если мессага пустая (только заголовок), то в оглавлении автоматом пишется (-)
''Оглавление - Химик Увы, там доставка от 2 дисков :( (-)
''Мессага    - Химик: Увы, там доставка от 2 дисков :(
    'Запись в лог программы
    LogEvent ("[Parser Level 2] enter routine")
    'Для начала перепрыгнем ПРИБЛИЗИТЕЛЬНО в нужную нам область
    Label14.Caption = "Анализ закачанной страницы"
    Label14.Refresh
    StartPos = InStr(1, lpString, "Добавить", vbTextCompare)
    If StartPos = 0 Then
        'Неполная закачка - повторный запрос
        lRetry = lRetry + 1
        Exit Function
    End If
    'Отрезаем все лишнее, чтобы не возиться
    '+1 чтобы буква "Д" осталась :) Это не нужно, но для эстетики :)
    lpString = Right(lpString, Len(lpString) - StartPos + 1)
    'Проезжаем почти до начала сообщения
    ''{FIXED} BUG Invalid call, add 1 for continue
    StartPos = InStr(1, lpString, "</center>", vbTextCompare) + Len("</center>")
    If StartPos = 0 Then
        'Неполная закачка - повторный запрос
        lRetry = lRetry + 1
        Exit Function
    End If
    'Теперь ищем хвост сообщения
    EndPos = InStr(StartPos, lpString, ">Ответить</A>", vbTextCompare)
    If EndPos = 0 Then
        'Неполная закачка - повторный запрос
        lRetry = lRetry + 1
        Exit Function
    End If
    'Отсекаем все лишнее (оставим закрывающую скобку (+1))
    lpString = Mid(lpString, StartPos, EndPos - StartPos + 1)
    'OK, теперь в зависимости от наличия e-mail автора вырезаем заголовок
    If hMsg.emailLen = 0 Then
        StartPos = InStr(1, lpString, "</I></B>:", vbTextCompare) + Len("</I></B>:")
    Else
        StartPos = InStr(1, lpString, "</I></B></A>:", vbTextCompare) + Len("</I></B></A>:")
    End If
    ''EndPos = InStr(StartPos, lpString, "<P>", vbTextCompare)
    'Это изменилось после того как сайт очнулся от спячки
    EndPos = InStr(StartPos, lpString, "</FONT> <P>", vbTextCompare)
    'Странно, но к этому моменту иногда доходит пустая строка. Еще одна проверка.
    If EndPos = 0 Then
        'Неполная закачка - повторный запрос
        lRetry = lRetry + 1
        Exit Function
    End If
    ''TO DO BUG lpString="" впосле окончания цикла чтения БД адресов, зациклился
    msgHead = Mid(lpString, StartPos, EndPos - StartPos)
    'Придется стереть (-), которые написаны авторами вручную. Иначе будет путаница
    'Используем переменные lResult и StartPos как временно не имеющие значения
    'Е-мое! Та же хрень и с (+)
    msgHead = RemoveSpaces(msgHead)
    StartPos = InStr(1, msgHead, "(-)", vbTextCompare)
    Do While StartPos > 0
    ''If StartPos > 0 Then
        '(-) найден. Проверим (грубо), находится ли он в хвосте...
        If StartPos > Len(msgHead) - 5 Then
            '(-) в конце строки - просто отрежем хвост
            msgHead = Left(msgHead, Len(msgHead) - 3)
        Else
            '(-) в середине. Вырежем только эти 3 символа
            msgHead = Left(msgHead, StartPos - 1) + Right(msgHead, Len(msgHead) - StartPos - 2)
        End If
    ''End If
        'Стираем пробелы
        msgHead = RemoveSpaces(msgHead)
        StartPos = InStr(1, msgHead, "(-)", vbTextCompare)
    Loop
    StartPos = InStr(1, msgHead, "(+)", vbTextCompare)
    Do While StartPos > 0
    ''If StartPos > 0 Then
        '(+) найден. Проверим (грубо), находится ли он в хвосте...
        If StartPos > Len(msgHead) - 5 Then
            '(+) в конце строки - просто отрежем хвост
            msgHead = Left(msgHead, Len(msgHead) - 3)
        Else
            '(+) в середине. Вырежем только эти 3 символа
            msgHead = Left(msgHead, StartPos - 1) + Right(msgHead, Len(msgHead) - StartPos - 2)
        End If
    ''End If
        'Стираем пробелы
        msgHead = RemoveSpaces(msgHead)
        StartPos = InStr(1, msgHead, "(+)", vbTextCompare)
    Loop
    'Запись в лог программы
    LogEvent ("[Parser Level 2] заголовок темы: " + msgHead)
    'Передвигаемся в начало сообщения
    StartPos = EndPos + Len("</FONT> <P>")
    'Проблема: в сообщении может быть множество тегов <P>, но мы должны
    'искать дату в конце сообщения.
    'Вероятность ложного срабатывания чрезвычайно мала, но, к сожалению, не равна 0:
    'мало ли кому придет в голову скопировать дату в сообщение...
    'Нужно искать образец: <P><I>27 Января 2001, 01:04:49</I> <FONT size=2><B>
    'Найти тег, закрывающий дату и то, что за ним
    lResult = InStr(StartPos, lpString, "</I> <FONT size=2><B>", vbTextCompare)
    'Это эмпирическая величина, 30 символов назад открывающий дату тег <I> еще
    'не должен был начаться, поэтому сейчас мы ЖЕЛЕЗНО отрежем именно то, что надо
    lResult = lResult - 30
    'Получаем точное расположение конца сообщения
    EndPos = InStr(lResult, lpString, "<I>", vbTextCompare)
    'Текст самого сообщения получим если подтвердится идентичность заголовков
    'иначе незачем время терять
    'Самое главное - проверить совпадение заголовков: все ли ОК с индексами book.ru?
    ''TO DO Временно отключим проверку If Left(hMsg.msgHead, hMsg.msgHeadLen) = msgHead Then
    ''Проверка перенесена чуть ниже
    tmpLen = Len(msgHead)
    If Not (hMsg.msgHeadLen = tmpLen) Then
        'Запись в лог программы
        LogEvent ("[Parser Level 2] Длина заголовка темы отлична от полученной в Parser Level 1")
        If hMsg.msgHeadLen > tmpLen Then
            'Заголовок короче, чем полученный в ParserLevel1
            'Дополним пробелами в конце
            msgHead = msgHead & String(hMsg.msgHeadLen - tmpLen, Chr(32))
            'Запись в лог программы
            LogEvent ("[Parser Level 2] Заголовок темы короче, чем полученный в Parser Level 1")
        'Else
            'Заголовок длиннее, чем полученный в ParserLevel1
            'Уменьшим величину tmpLen, присвоив ей значение из hMsg
            '(после End If, т.к. это нужно сделать в обоих случаях)
        End If
        tmpLen = hMsg.msgHeadLen
    End If
    If Left(hMsg.msgHead, hMsg.msgHeadLen) = Left(msgHead, tmpLen) Then
    'Да, это одно и то же сообщение
        'Получаем текст самого сообщения
        tmpstr = Mid(lpString, StartPos, EndPos - StartPos)
        'Часто сообщение завершается тегом параграфа <P>. Уберем, чтобы не плодить
        'пустых строк на итоговой странице
        Do While UCase(Right(tmpstr, 3)) = "<P>"
            tmpstr = Left(tmpstr, Len(tmpstr) - 3)
        Loop
        ParserLevel2 = tmpstr
        hMsg.MsgSize = Len(tmpstr)
        'Еще дата осталась...
        StartPos = EndPos + Len("<I>")
        EndPos = InStr(StartPos, lpString, "</i>", vbTextCompare)
        tmpstr = Mid(lpString, StartPos, EndPos - StartPos)
        hMsg.msgDate = tmpstr
        hMsg.msgDateLen = Len(tmpstr)
        'Сохранить собранные сведения в журнале
        Label14.Caption = "Сохранение полученной информации"
        Label14.Refresh
                'Забыл! Перекачать инфу в класс! Исправляюсь...
                'Да! Заодно же надо сохранить новые сведения (длина сообщения, дата, ее длина)
'Запись в лог программы
LogEvent ("[Parser Level 2] Сбор полученной информации для автора " + Left(hMsg.Author, hMsg.AuthorLen))
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
'Запись в лог программы
LogEvent ("[Parser Level 2] Сохранение полученной информации - вызов MsgMap.UpdateDB(ucMsgURLsMode, " + Str(RecNum) + ")")
        bResult = MsgMap.UpdateDB(ucMsgURLsMode, RecNum)
    Else
        'Запись в лог программы
        LogEvent ("[Parser Level 2] Заголовок темы не соответствует заголовку, полученному в Parser Level 1. Запись номер " + Str(RecNum))
        'Заголовки различаются! Ошибка...
        hMsg.MsgSize = 0
        hMsg.msgDate = strEmpty
        hMsg.msgDateLen = -1
        ParserLevel2 = "<i>Информация <b>EliteGames conference extractor</b></i>. Ошибка индексации на сервере www.book.by.ru Сообщение перезаписано другим и не может быть прочитано"
        Label17.Caption = Str(Val(Label17.Caption) + 1)
    End If
    With Authors
        'Запись в лог программы
        LogEvent ("[Parser Level 2] Начало передачи информации об авторах в cAuthors")
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
        'Остальные поля не несут полезной информации, обнуляем (? ЧЕГО???)
        .TotalNums = 1
        .TotalSize = hMsg.MsgSize
        If Not (hMsg.msgDateLen <= 0) Then
            'перевести дату в цифровой формат
            'Формат времени YYYY.MM.DD (HH:MM:SS)
            
            'Отсекаем месяц
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
                Case "Января":      tmpstr = "01"
                Case "Февраля":     tmpstr = "02"
                Case "Марта":       tmpstr = "03"
                Case "Апреля":      tmpstr = "04"
                Case "Мая":         tmpstr = "05"
                Case "Июня":        tmpstr = "06"
                Case "Июля":        tmpstr = "07"
                Case "Августа":     tmpstr = "08"
                Case "Сентября":    tmpstr = "09"
                Case "Октября":     tmpstr = "10"
                Case "Ноября":      tmpstr = "11"
                Case "Декабря":     tmpstr = "12"
                'На случай опечатки или еще чего
                Case Else:          tmpstr = "13"
            End Select
            'Отрезаем число, получаем конструкцию .MM.DD
            .FirstDate = "." & tmpstr & "." & Left(hMsg.msgDate, 2)
            'Вырезаем год
            For i = 3 To tmpLen
                If Mid(hMsg.msgDate, i, 1) Like "[0-9]" Then
                    tmpstr = Mid(hMsg.msgDate, i, 4)
                    i = tmpLen
                End If
            Next i
            'Получаем YYYY.MM.DD_
            .FirstDate = tmpstr & .FirstDate & Chr(32)
            'Вырезаем время
            tmpLen = InStr(1, hMsg.msgDate, "(", vbTextCompare)
            tmpstr = Mid(hMsg.msgDate, tmpLen, 10)
            .FirstDate = .FirstDate & tmpstr
            .FirstDateLen = 21
        Else
'Запись в лог программы
LogEvent ("[Parser Level 2] У текущего автора дата <=0")
            .FirstDate = strEmpty
            .FirstDateLen = 0
        End If
    End With
'Запись в лог программы
LogEvent ("[Parser Level 2] Сохранение информации об авторе - вызов Authors.UpdateDB(ucPrimaryMode)")
    If Authors.UpdateDB(ucPrimaryMode) Then
        'Увеличим значение счетчика обработанных записей
        Label6.Caption = Str(Val(Label6.Caption) + 1)
        'Учтем и старые eMail-адреса автора, если они были обнаружены в UpdateDB:
        hMsg.emailLen = Authors.emailLen
        hMsg.email = Left(Authors.email, Authors.emailLen)
    Else
'Запись в лог программы
LogEvent ("[Parser Level 2] Ошибка вызова Authors.UpdateDB(ucPrimaryMode)")
        'TO DO Обновить Label(Thread) +1 обработано + ОШИБКА!!!
        Label6.Caption = Str(Val(Label6.Caption) + 1)
        Label6.Caption = Str(Val(Label16.Caption) + 1)
    End If
    'Ссылка успешно обработана: повторная закачка не требуется (lRetry = 0)
    lRetry = 0
    'ВСЕ!!!
End Function

Private Sub Command1_Click()
    Call mnuStart_Click
End Sub

'*****************************************************
'* Нажата кнопка СТОП - останавливаем закачки        *
'*****************************************************
Private Sub Command2_Click()
    Call mnuStop_Click
End Sub

Private Sub Command3_Click()
    'Поднимаем флаг "Закачка страницы произведена", пусть дальше
    'разбираются, корректно или нет
    bDone = True
''{NET ERROR}    Call Inet1_DocumentComplete(Nothing, "dfg")
End Sub

Private Sub Form_Load()
    Loading = True  'Загрузка - не обрабатывать информацию
    strMess = strEmpty
    Inet1.Offline = WorkOffline
    WebBrowser1.Offline = WorkOffline
    'Опрос качалок будем производить раз в 2 секунды
    'TO DO в будущем - из настроек
    Timer1.Interval = 1500
    Timer1.Enabled = False
    Command2.Enabled = False
    mnuStop.Enabled = False
    mnuStart.Enabled = False
    'Нельзя давать повторный запрос в данный момент
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
    'TO DO Перенести следующую строку в главный цикл
    Set MsgMap = Nothing
    Set Authors = Nothing
    Set HTMLDoc = Nothing
    Inet1.Stop
    Set fBrw = Nothing
End Sub

Private Sub Inet1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If Not Loading Then
        'Ловим завершающее событие DocumentComplete
        If (pDisp Is Inet1.Object) Then  'Is проверяет идентичность объектов
''{NET ERROR}        If (pDisp Is Inet1.Object) Or (pDisp Is Nothing) Then  'Is проверяет идентичность объектов
            'Читаем HTML-источник документа
            Set HTMLDoc = Inet1.Document
            If HTMLDoc Is Nothing Then
            ''TO DO переделать позже (вывод сообщения)
            ' Not an HTLM document
                Exit Sub
            End If
            vtData = HTMLDoc.body.innerHTML
            'Убираем все кавычки, а то замучаемся с ними
                ''TO DO переделать позже
                Label8.Caption = "Чистим информацию от мусора"
                Label8.Refresh
            strMess = Replace(CStr(vtData), """", "`", 1, -1, vbBinaryCompare)
                ''TO DO переделать позже
                Label8.Caption = "Чистим последствия чистки"
                Label8.Refresh
            strMess = ClearGarbage(strMess)
            'Нельзя давать повторный запрос в данный момент
            Command3.Enabled = False
            Command3.Refresh
            bDone = True    'Страница получена
        End If
    End If
End Sub

Private Sub Inet1_DownloadBegin()
    ''TO DO переделать позже
    Label8.Caption = "Получен ответ..."
    Label8.Refresh
End Sub

Private Sub Inet1_DownloadComplete()
'TO DO заглушка!!! Если DocumentComplete будет работать, то эту процедуру надо будет стереть
Exit Sub
    'Читаем HTML-источник документа
    Set HTMLDoc = Inet1.Document
    If HTMLDoc Is Nothing Then
    ''TO DO переделать позже (вывод сообщения)
    ' Not an HTLM document
        Exit Sub
    End If
    strMess = Replace(HTMLDoc.body.innerHTML, """", "`", 1, -1, vbBinaryCompare)
''''''    vtData = HTMLDoc.body.innerHTML
    'Убираем все кавычки, а то замучаемся с ними
        ''TO DO переделать позже
        Label8.Caption = "Чистим информацию от мусора"
''''''    strMess = Replace(CStr(vtData), """", "`", 1, -1, vbBinaryCompare)
        ''TO DO переделать позже
        Label8.Caption = "Чистим последствия чистки"
    strMess = ClearGarbage(strMess)
    'Нельзя давать повторный запрос в данный момент
    Command3.Enabled = False
    Command3.Refresh
    bDone = True    'Страница получена
End Sub

Private Sub mnuAbout_Click()
    Load frmAbout
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuConnect_Click()
'Выбрана команда Подключиться

Dim LastRecNum As Long
Dim i As Long
Dim blValue As Long 'Было Boolean, но файл никак не хотел читаться, пришлось заменить :(
Dim hFile As Long
Dim NextPage As Long    'Номер следующей закачиваемой страницы
    
    'Открываем индекс страниц
    hFile = FreeFile
    Open (App.Path & DataPath & PagesFile) For Random As hFile Len = Len(blValue)
    LastRecNum = LOF(hFile) \ Len(blValue)
    NextPage = 0
    If LastRecNum > 0 Then
        'Уже что-то закачано, проверим, все ли там подряд закачивалось?
        For i = 1 To LastRecNum
            Get hFile, i, blValue
            If blValue = 0 Then
                'Нет, есть дырки, закачивать будем одну из дырок (первую попавшуюся)
                NextPage = i
                i = LastRecNum
            End If
        Next i
        If NextPage = 0 Then
            'В диапазоне (1, LastRecNum) все страницы закачаны
            If Not (LastRecNum >= 80) Then
                'Если не все 80 страниц еще закачаны, то продолжаем качать страницы
                'подряд с того места, где закончили в прошлый раз (LastRecNum)
                NextPage = LastRecNum + 1
            End If
        End If
    End If
    Close hFile
    Text2.Text = Format(NextPage)
    
    'Попытаемся загрузить страницу форума (Win автоматически выполнит запрос на соединение)
''DEBUG mnuStart.Enabled = True
    WorkOffline = False
    WebBrowser1.Offline = WorkOffline
    WebBrowser1.Navigate ("http://book.by.ru/cgi-bin/book.cgi?book=Elitegames")
End Sub

Private Sub mnuExit_Click()
    'Остановка записи лога программы
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
'* Определяет номер запрашиваемой страницы           *
'*****************************************************
Private Function GetPageNum() As Long
    'TO DO написать норамальное определение номера страницы (с начала списка!!!)
    GetPageNum = CLng(Val(Text2.Text))
    If GetPageNum > LastPageAvailable Then
        'TO DO Warning message
        GetPageNum = LastPageAvailable
    ElseIf GetPageNum < 1 Then
        GetPageNum = 1
    End If
End Function

'*****************************************************
'* Выбран пункт меню СТАРТ                           *
'* Инициирует загрузку и разделение страницы         *
'*****************************************************
Private Sub mnuStart_Click()
Dim strURL As String
    Loading = False 'Начать отслеживание загрузки страниц
    Inet1.Offline = WorkOffline
    'Получаем номер страницы для закачки
    CurrPage = GetPageNum
    ''TO DO Если номер задан неправильно, то выйти!!!
    'Запрещаем команды СТАРТ в меню и на кнопке
    Command1.Enabled = False
    mnuStart.Enabled = False
    'Начинаем отслеживать нажатие кнопки Cancel
    bCancel = False
    Command2.Enabled = True
    mnuStop.Enabled = True
    'Нельзя выходить при работающей закачке
    mnuExit.Enabled = False
    'Сброс информации об ошибках
    Label16.Caption = "0"
    Label17.Caption = "0"
    Label2.Caption = "0"
    Label4.Caption = "0"
    Label6.Caption = "0"
    'Прочитать страницу ссылок
    ''TO DO здесь задается адрес (учитывать автоочередь)
    strURL = "/book.cgi?book=Elitegames&p=" & Format(CurrPage) & "&ac=1"
    ''strURL = "C:\tmp\test8.htm"
    'Кнопка повтора пока не нужна.
    ''Command3.Enabled = False
    ''Command3.Refresh
    bDone = False
    GettingURLs = True
    StartInet (strURL)
    'Перебиваем сообщение от StartInet
    Label8.Caption = "Запрос оглавления ушел"
    Label8.Refresh
    'Пуск таймера!
    Timer1.Enabled = True
    Debug.Print "Запрос оглавления ушел!"
End Sub

'*****************************************************
'* Запускает качалку и посылает запрос на  msgURL    *
'* Ожидание результата и его обработка - в           *
'* Inet_DocumentComplete и Feeder соответственно     *
'*****************************************************
Private Sub StartInet(ByVal msgURL As String)
    'Даем возможность повторного запроса
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
    Debug.Print "Запрос ушел, " & BaseHREFprefix & BaseHREF & msgURL
    Label8.Caption = "Запрос ушел"
    Label8.Refresh
End Sub

'*****************************************************
'* Выбран пункт меню СТОП - останавливаем закачки    *
'*****************************************************
Private Sub mnuStop_Click()
    'Флаг нажатия кнопки СТОП - на всякий случай
    bCancel = True
    'Разрешаем кнопку ПУСК
    Command1.Enabled = True
    'Разрешаем пункты меню НАЧАТЬ и ВЫХОД
    mnuStart.Enabled = True
    mnuExit.Enabled = True
    'Остановить качалки при нажатии Cancel (Stop)
    Inet1.Stop
    Label8.Caption = "Закачка остановлена"
    Label8.Refresh
    Call StopFeeder(ucDLstopped)
    'TO DO разобраться тут с контекстом помощи
    Result = MsgBox("Закачка текущей страницы отменена. Тем не менее," + vbCrLf + "Вы можете просмотреть всю информацию, которую" + vbCrLf + "удалось закачать до момента остановки." + vbCrLf + "Воспользуйтесь разделом ""Остановка закачки"" справочной" + vbCrLf + "системы для детальных инструкций.", vbOKOnly + vbExclamation, "Закачка страницы " & Format(CurrPage) & " остановлена")
End Sub

'*****************************************************
'* Загружает HTML файл с результатами закачки        *
'*****************************************************
Private Sub mnuView_Click()
    Call ShowHTML
End Sub

'*****************************************************
'*   Обновляет HTML-файлы со сведениями об авторах   *
'*****************************************************
Private Sub mnuWriteAuthors_Click()
    'Обновление статистики по авторам
    Label14.Caption = "Запись статистики в HTML..."
    Label14.Refresh
    bResult = Authors.WriteAHTMLs
    Label14.Caption = "Ожидание команды."
    Label14.Refresh
    Label8.Caption = "Готово."
    Label8.Refresh
End Sub

Private Sub Timer1_Timer()
    If GettingURLs Then
    'Таймер: сейчас идет получение оглавления
        If bDone Then
            'На всякий случай попробуем стереть БД (мало ли что осталось от предыдущего раза)
            MsgMap.Clear
            ''Call Inet1_DocumentComplete(Nothing, "dfg")
            If Splitter(strMess) Then
                GettingURLs = False
                bResult = InitFeeder
            End If
        End If
    Else
    'Сейчас качается текст сообщений
        If bDone Then
            lResult = Feeder
        End If
    End If
End Sub

Private Sub WebBrowser1_DownloadBegin()
    'Если браузер начал получать данные, то мы в сети - можно разрешить закачку
    Command1.Enabled = True
    mnuStart.Enabled = True
End Sub
