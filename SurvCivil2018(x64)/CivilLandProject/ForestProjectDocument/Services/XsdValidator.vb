Imports System.Xml
Imports System.Xml.Schema
Imports System.Collections.Generic

''' <summary>
''' Валидация XML по XSD-схеме.
''' </summary>
Public Class XsdValidator

    ''' <summary>
    ''' Проверяет XML-файл по XSD-схеме.
    ''' </summary>
    ''' <param name="xmlPath">Путь к XML-файлу</param>
    ''' <param name="xsdPaths">Пути к XSD-файлам (основная схема + все импортируемые)</param>
    ''' <returns>Список ошибок (пустой, если всё ок)</returns>
    Public Shared Function Validate(xmlPath As String, xsdPaths As IEnumerable(Of String)) As List(Of String)
        Dim errors As New List(Of String)()

        Dim settings As New XmlReaderSettings()
        settings.ValidationType = ValidationType.Schema

        ' Подключаем все XSD
        For Each xsdPath In xsdPaths
            If IO.File.Exists(xsdPath) Then
                settings.Schemas.Add(Nothing, xsdPath)
            End If
        Next

        ' Обработчик ошибок валидации
        AddHandler settings.ValidationEventHandler,
            Sub(sender, e)
                errors.Add($"[{e.Severity}] {e.Message} (строка {e.Exception.LineNumber}, позиция {e.Exception.LinePosition})")
            End Sub

        Using reader As XmlReader = XmlReader.Create(xmlPath, settings)
            While reader.Read()
            End While
        End Using

        Return errors
    End Function

End Class
