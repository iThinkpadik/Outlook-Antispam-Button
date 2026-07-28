' --- НАСТРОЙКИ (при необходимости измените здесь) ---
' Адрес электронной почты, куда будут отправляться жалобы на спам
Public Const SPAM_REPORT_EMAIL As String = "Введите сюда адрес почты" '<---Поставить свою почту'
' --- КОНЕЦ НАСТРОЕК ---

Sub ForwardAndDeleteSpam()
    '
    ' Отправляет выделенное письмо как вложение с возможностью добавить комментарий.
    ' Предлагает выбор: удалить письмо, переместить в "Отправленные" или в "Архив".
    '
    On Error GoTo ErrorHandler

    Dim objItem As Object
    Dim objMsg As Outlook.MailItem
    Dim userComment As String
    Dim actionChoice As VbMsgBoxResult
    Dim targetFolder As Outlook.Folder
    Dim currentDate As String
    
    ' 1. Получаем выделенное письмо
    Set objItem = GetCurrentItem()
    
    ' Проверяем, что письмо действительно выделено
    If objItem Is Nothing Then
        MsgBox "Пожалуйста, выделите письмо, которое хотите отправить как спам.", vbExclamation, "Ничего не выбрано"
        Exit Sub
    End If

    ' 2. Запрашиваем комментарий у пользователя
    userComment = InputBox("Введите ваш комментарий к этому письму (например, почему вы считаете его спамом):", "Комментарий к жалобе на спам", "Подозрительное письмо")

    ' Если пользователь нажал "Отмена" - отменяем всю операцию
    If StrPtr(userComment) = 0 Then
        MsgBox "Операция отменена.", vbInformation, "Отмена"
        Exit Sub
    End If

    ' 3. Спрашиваем, что делать с исходным письмом
    actionChoice = MsgBox("Что сделать с исходным письмом после отправки?" & vbCrLf & vbCrLf & "Нажмите 'Да' - удалить письмо" & vbCrLf & "Нажмите 'Нет' - переместить в папку 'Отправленные'" & vbCrLf & "Нажмите 'Отмена' - переместить в папку 'Архив'" & vbCrLf & vbCrLf & "(Если папка 'Архив' не найдена, письмо будет перемещено в 'Отправленные')", vbYesNoCancel + vbQuestion, "Выберите действие")

    ' Если нажали Отмена - выходим
    If actionChoice = vbCancel Then
        MsgBox "Операция отменена.", vbInformation, "Отмена"
        Exit Sub
    End If

    ' 4. Формируем текущую дату и время
    currentDate = Format(Now, "dd.mm.yyyy HH:MM:SS")
    
    ' 5. Создаем новое письмо
    Set objMsg = Application.CreateItem(olMailItem)
    
    With objMsg
        ' Вкладываем исходное письмо
        .Attachments.Add objItem, olEmbeddeditem

        ' Тема письма (как у оригинала)
        .Subject = "REPORT SPAM: " & objItem.Subject

        ' Адрес получателя
        .To = SPAM_REPORT_EMAIL

        ' --- Формируем тело письма с комментарием и датой ---
        Dim bodyText As String
        bodyText = "Жалоба на подозрительное письмо." & vbCrLf & vbCrLf
        bodyText = bodyText & "Дата и время отправки: " & currentDate & vbCrLf & vbCrLf
        bodyText = bodyText & "Комментарий отправителя:" & vbCrLf
        bodyText = bodyText & "---" & vbCrLf
        bodyText = bodyText & userComment & vbCrLf
        bodyText = bodyText & "---" & vbCrLf & vbCrLf
        bodyText = bodyText & "Исходное письмо прикреплено к этому сообщению как вложение."

        .Body = bodyText

        ' Отправляем письмо
        .Send
    End With

    ' 6. Обрабатываем исходное письмо в зависимости от выбора пользователя
    Select Case actionChoice
        Case vbYes ' Удалить
            objItem.Delete
            MsgBox "Письмо отправлено как спам и удалено из папки 'Входящие'.", vbInformation, "Готово!"

        Case vbNo ' Переместить в "Отправленные"
            Set targetFolder = GetFolderByPath("Отправленные")
            If Not targetFolder Is Nothing Then
                objItem.Move targetFolder
                MsgBox "Письмо отправлено как спам и перемещено в папку 'Отправленные'.", vbInformation, "Готово!"
            Else
                ' Если папка не найдена, просто удаляем
                objItem.Delete
                MsgBox "Папка 'Отправленные' не найдена. Письмо удалено.", vbExclamation, "Предупреждение"
            End If

        Case Else ' На всякий случай (если что-то пошло не так)
            objItem.Delete
            MsgBox "Письмо отправлено как спам и удалено.", vbInformation, "Готово!"
    End Select
    
    ' 7. Очищаем память
    Set objItem = Nothing
    Set objMsg = Nothing
    Set targetFolder = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox "Произошла ошибка: " & Err.Description & vbCrLf & "Код ошибки: " & Err.Number, vbCritical, "Ошибка"
    
    ' Очищаем память при ошибке
    Set objItem = Nothing
    Set objMsg = Nothing
    Set targetFolder = Nothing
End Sub

Function GetCurrentItem() As Object
    ' Возвращает текущее выделенное письмо (или открытое в отдельном окне)
    On Error Resume Next

    Select Case TypeName(Application.ActiveWindow)
        Case "Explorer"
            ' Письмо выделено в списке (но не открыто)
            If Application.ActiveExplorer.Selection.Count > 0 Then
                Set GetCurrentItem = Application.ActiveExplorer.Selection.Item(1)
            End If
        Case "Inspector"
            ' Письмо открыто в отдельном окне
            Set GetCurrentItem = Application.ActiveInspector.CurrentItem
        Case Else
            ' Ничего не делаем
    End Select
    
    Set objApp = Nothing
End Function

Function GetFolderByPath(folderName As String) As Outlook.Folder
    ' Функция для поиска папки по имени в корне почтового ящика
    On Error Resume Next

    Dim objNamespace As Outlook.NameSpace
    Dim objFolder As Outlook.Folder
    
    Set objNamespace = Application.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox).Parent
    
    ' Ищем папку с указанным именем
    For Each objFolder In objFolder.Folders
        If objFolder.Name = folderName Then
            Set GetFolderByPath = objFolder
            Exit Function
        End If
    Next
    
    ' Если папка не найдена, возвращаем Nothing
    Set GetFolderByPath = Nothing
    Set objNamespace = Nothing
    Set objFolder = Nothing
End Function
