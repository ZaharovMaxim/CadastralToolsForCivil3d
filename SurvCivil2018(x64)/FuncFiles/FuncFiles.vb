Imports System.Reflection
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Public Class FuncFiles
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вывод диалогового окна выбор каталога
    Public Shared Function FuncGetAbsoluteFolderPathName(ByVal strTitle As String, Optional lngRegim As Long = 1) As String
        ' Определяем переменные
        Dim objShell As Object
        Dim objFolder As Object
        Dim objFolderItem As Object
        ' Создаём переменную Shell
        objShell = CreateObject("Shell.Application")
        ' Выводим диалоговое окно выбора папки с нужными параметрами
        objFolder = objShell.BrowseForFolder(0, strTitle, lngRegim)
        ' Если объект создан не удачно (НЕ выбрали какую-то папку), то
        ' возвращаем пустую строку в качестве результата работы функции...
        If objFolder Is Nothing Then
            FuncGetAbsoluteFolderPathName = ""
            Exit Function
        End If
        ' Получаем объект, у которого "можно спросить" его path
        objFolderItem = objFolder.Self
        ' Получаем значение Path
        FuncGetAbsoluteFolderPathName = objFolderItem.Path
        ' Удаляем все использованные объекты
        objFolderItem = Nothing
        objFolder = Nothing
        objShell = Nothing
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вывод всех каталогов для поска файлов поддержки
    Public Shared Function FuncGetSupportFolderPath(ByRef strPathFolders As String()) As Boolean
        FuncGetSupportFolderPath = False
        Try
            Dim AcadApp As Object = Application.AcadApplication
            Dim preferences As Object = AcadApp.GetType().InvokeMember("Preferences", BindingFlags.GetProperty, Nothing, AcadApp, Nothing)
            Dim files As Object = preferences.GetType().InvokeMember("Files", BindingFlags.GetProperty, Nothing, preferences, Nothing)
            Dim supportPath As String = files.GetType().InvokeMember("SupportPath", BindingFlags.GetProperty, Nothing, files, Nothing)
            If IsNothing(supportPath) = False Then
                If supportPath.Trim.Length > 0 Then
                    strPathFolders = supportPath.Split(";")
                    FuncGetSupportFolderPath = True
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'пути поиска файлов шаблона
    Public Shared Function FuncFindTemplateFiles(ByVal nameFileXML As String, ByVal nameFileDwg As String, ByRef putchFileXML As String, ByRef putchFileDWG As String) As Boolean
        FuncFindTemplateFiles = False
        Dim boolFindFileXML As Boolean = False
        Dim boolFindFileDwg As Boolean = False
        Try
            Dim AcadApp As Object = Application.AcadApplication
            Dim preferences As Object = AcadApp.GetType().InvokeMember("Preferences", BindingFlags.GetProperty, Nothing, AcadApp, Nothing)
            Dim files As Object = preferences.GetType().InvokeMember("Files", BindingFlags.GetProperty, Nothing, preferences, Nothing)
            Dim supportPath As String = files.GetType().InvokeMember("SupportPath", BindingFlags.GetProperty, Nothing, files, Nothing)
            Dim strPathFolders As String() = Nothing
            If IsNothing(supportPath) = False Then
                If supportPath.Trim.Length > 0 Then
                    strPathFolders = supportPath.Split(";")
                    If strPathFolders.Length > 0 Then
                        For i As Integer = 0 To strPathFolders.Length - 1
                            Dim tempDir As String = strPathFolders(i)
                            For Each userfiles As String In IO.Directory.GetFiles(tempDir)
                                Try
                                    Dim File1 As String = IO.Path.GetFileName(userfiles)
                                    If nameFileXML Like File1 Then
                                        putchFileXML = userfiles
                                        boolFindFileXML = True
                                    ElseIf nameFileDwg Like File1 Then
                                        putchFileDWG = userfiles
                                        boolFindFileDwg = True
                                    End If
                                Catch ex As ArgumentException
                                End Try
                                If boolFindFileDwg = True And boolFindFileXML = True Then
                                    Exit For
                                End If
                            Next
                            If boolFindFileDwg = True And boolFindFileXML = True Then
                                Exit For
                            End If
                        Next i
                    End If
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If boolFindFileDwg = True And boolFindFileXML = True Then
                Return True
            Else
                Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
                Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
                'отделяем имя файла от папки
                Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
                'определяем имя родительской папки для папки где находится сборка
                Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
                If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
                    Dim userDir As String = parentPatch2 & "\FileResources\Sample\"
                    If IO.File.Exists(userDir & nameFileXML) Then
                        putchFileXML = userDir & nameFileXML
                        boolFindFileXML = True
                    End If
                    If IO.File.Exists(userDir & nameFileDwg) Then
                        putchFileDWG = userDir & nameFileDwg
                        boolFindFileDwg = True
                    End If
                End If
            End If
        Catch ex As System.Exception
        End Try
        If boolFindFileDwg = True And boolFindFileXML = True Then
            Return True
        End If
    End Function

End Class
