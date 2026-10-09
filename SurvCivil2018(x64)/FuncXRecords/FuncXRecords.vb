Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.Xml
Imports CadastralTools.StructureElement
Imports CadastralTools
Imports Autodesk.AutoCAD.Runtime
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json
Public Class FuncXRecords
    'чтение данных XRecords из примитива
    Public Shared Function getXRecords(
        ByVal entity As Autodesk.AutoCAD.DatabaseServices.Entity,
        ByRef dataXRecord As StructureElement,
        Optional nameTableXRecords As StructureElement.tableXRecords = StructureElement.tableXRecords.PROJECT_STRUCTURES,
        Optional db As Database = Nothing) As Boolean

        If entity Is Nothing Then Return False
        dataXRecord = Nothing

        ' Берём БД: если не передали — из самой сущности (работает и для side database)
        Dim database As Database = If(db, entity.Database)
        If database Is Nothing Then Return False

        Try
            Using tr As Transaction = database.TransactionManager.StartTransaction()

                Dim extId As ObjectId = entity.ExtensionDictionary
                If extId.IsNull Then Return False

                Dim userDict As DBDictionary = TryCast(tr.GetObject(extId, OpenMode.ForRead), DBDictionary)
                If userDict Is Nothing OrElse userDict.Count = 0 Then Return False

                Dim name As String = [Enum].GetName(GetType(StructureElement.tableXRecords), nameTableXRecords)
                If String.IsNullOrEmpty(name) Then Return False

                ' GetAt бросает eKeyNotFound, если ключа нет — проверяем через Contains
                If Not userDict.Contains(name) Then Return False

                Dim idKeyDict As ObjectId = userDict.GetAt(name)
                If idKeyDict.IsNull Then Return False

                Dim rec As Xrecord = TryCast(tr.GetObject(idKeyDict, OpenMode.ForRead), Xrecord)
                If rec Is Nothing Then Return False

                Using rb As ResultBuffer = rec.Data
                    If rb Is Nothing Then Return False

                    Dim tVal As TypedValue() = rb.AsArray
                    If tVal Is Nothing OrElse tVal.Length = 0 Then Return False

                    dataXRecord = New StructureElement

                    For j As Integer = 0 To tVal.Length - 1
                        Dim valStr As String = tVal(j).ToString
                        valStr = valStr.Replace("(1000,", "")
                        If valStr.Length = 0 Then Continue For
                        valStr = valStr.Remove(valStr.Length - 1)

                        Dim valArrayStr As String() = valStr.Split("=")
                        If valArrayStr.Length < 2 Then Continue For

                        Dim nameField As String = valArrayStr(0).Trim()
                        Dim lenZn As Integer = valArrayStr(0).Trim.Length
                        Dim valueZn As String = Mid(valStr, lenZn + 2)

                        Select Case nameField
                            Case "Label"
                                dataXRecord.Label = valueZn
                            Case "ClassObject"
                                Try
                                    dataXRecord.NameClassObject = CType([Enum].Parse(GetType(CadastralTools.StructureElement.classObject), valueZn), CadastralTools.StructureElement.classObject)
                                Catch ex As TypeLoadException
                                    dataXRecord.NameClassObject = classObject.OtherElement
                                Catch ex As ArgumentException
                                    dataXRecord.NameClassObject = classObject.OtherElement
                                End Try

                            Case "TypeGeometry"
                                Try
                                    dataXRecord.TypeGeometry = CType([Enum].Parse(GetType(CadastralTools.StructureElement.typeGeometryObject), valueZn), CadastralTools.StructureElement.typeGeometryObject)
                                Catch ex As TypeLoadException
                                    dataXRecord.TypeGeometry = typeGeometryObject.none
                                Catch ex As ArgumentException
                                    dataXRecord.TypeGeometry = typeGeometryObject.none
                                End Try

                            Case "NameTypeObject"
                                Try
                                    dataXRecord.NameTypeObject = CType([Enum].Parse(GetType(CadastralTools.StructureElement.nameObject), valueZn), CadastralTools.StructureElement.nameObject)
                                Catch ex As TypeLoadException
                                    dataXRecord.NameTypeObject = nameObject.OtherObject
                                Catch ex As ArgumentException
                                    dataXRecord.NameTypeObject = nameObject.OtherObject
                                End Try

                            Case "TableObject"
                                dataXRecord.TableObject = valueZn

                            Case "Description"
                                dataXRecord.Description = valueZn

                            Case "tablePS"
                                dataXRecord.TablePS = valueZn

                            Case "KeyParameter"
                                dataXRecord.KeyParameter = valueZn

                            Case "IdStructure"
                                dataXRecord.IdStructure = valueZn

                            Case "IdElement"
                                dataXRecord.IdElement = valueZn

                            Case "Note"
                                dataXRecord.Note = valueZn

                            Case "IdObject"
                                dataXRecord.IdObject = entity.Handle.ToString
                        End Select
                    Next j
                End Using

                ' Не Commit, т.к. читали только ForRead — транзакция закроется через Using
                Return True
            End Using
        Catch ex As System.Exception
            System.Diagnostics.Trace.WriteLine("[getXRecords] " & ex.ToString())
            dataXRecord = Nothing
            Return False
        End Try
    End Function
    'запись данных в примитив
    Public Shared Function setXRecords(
        ByRef entityObject As Entity,
        ByRef dataRecords As StructureElement,
        Optional ByVal tableXRecords As CadastralTools.StructureElement.tableXRecords = CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES,
        Optional db As Database = Nothing) As Boolean

        If entityObject Is Nothing OrElse dataRecords Is Nothing Then Return False

        Dim database As Database = If(db, entityObject.Database)
        If database Is Nothing Then Return False

        Dim nameTableXRecord As String = tableXRecords.ToString()
        If String.IsNullOrEmpty(nameTableXRecord) Then Return False

        '--------------------------------------------------------------------------------
        ' Подготовка JSON: убираем ListCoord, если это валидный JSON
        '--------------------------------------------------------------------------------
        Dim jsonString As String = dataRecords.KeyParameter
        Try
            If Not String.IsNullOrEmpty(jsonString) AndAlso CadastralTools.JsonValidator.IsValidJson(jsonString) Then
                Dim jsonObject As JObject = JObject.Parse(jsonString)
                If jsonObject("ListCoord") IsNot Nothing Then
                    jsonObject.Remove("ListCoord")
                    jsonString = jsonObject.ToString()
                End If
            End If
        Catch ex As JsonException
            System.Diagnostics.Trace.WriteLine("[writeXRecords] JSON: " & ex.Message)
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine("[writeXRecords] JSON: " & ex.ToString())
        End Try

        '--------------------------------------------------------------------------------
        ' Готовим пары ключ/значение
        '--------------------------------------------------------------------------------
        Dim pairs As New List(Of KeyValuePair(Of String, String)) From {
        New KeyValuePair(Of String, String)("Label", If(dataRecords.Label, "")),
        New KeyValuePair(Of String, String)("ClassObject", If(dataRecords.NameClassObject.ToString(), "")),
        New KeyValuePair(Of String, String)("TypeGeometry", If(dataRecords.TypeGeometry.ToString(), "")),
        New KeyValuePair(Of String, String)("NameTypeObject", If(dataRecords.NameTypeObject.ToString(), "")),
        New KeyValuePair(Of String, String)("TableObject", If(dataRecords.TableObject, "")),
        New KeyValuePair(Of String, String)("Description", If(dataRecords.Description, "")),
        New KeyValuePair(Of String, String)("KeyParameter", If(jsonString, "")),
        New KeyValuePair(Of String, String)("tablePS", If(dataRecords.TablePS, "")),
        New KeyValuePair(Of String, String)("IdStructure", If(dataRecords.IdStructure, "")),
        New KeyValuePair(Of String, String)("IdElement", If(dataRecords.IdElement, "")),
        New KeyValuePair(Of String, String)("Note", If(dataRecords.Note, "")),
        New KeyValuePair(Of String, String)("IdObject", entityObject.Handle.ToString())
    }

        Try
            Using tr As Transaction = database.TransactionManager.StartTransaction()

                ' Открываем сущность на запись
                Dim entity As Entity = TryCast(tr.GetObject(entityObject.ObjectId, OpenMode.ForWrite), Entity)
                If entity Is Nothing Then Return False

                '--------------------------------------------------------------------------------
                ' Создаём/получаем ExtensionDictionary
                '--------------------------------------------------------------------------------
                Dim extId As ObjectId = entity.ExtensionDictionary
                Dim userDict As DBDictionary = Nothing

                If extId.IsNull Then
                    entity.CreateExtensionDictionary()
                    extId = entity.ExtensionDictionary
                    If extId.IsNull Then Return False
                End If

                userDict = TryCast(tr.GetObject(extId, OpenMode.ForWrite), DBDictionary)
                If userDict Is Nothing Then Return False

                '--------------------------------------------------------------------------------
                ' Формируем ResultBuffer
                '--------------------------------------------------------------------------------
                Dim rb As ResultBuffer = New ResultBuffer()
                For Each kvp As KeyValuePair(Of String, String) In pairs
                    If String.IsNullOrEmpty(kvp.Key) Then Continue For
                    Dim val As String = If(kvp.Value, "")
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, kvp.Key.Trim() & "=" & val))
                Next

                '--------------------------------------------------------------------------------
                ' Пишем Xrecord: обновляем существующий или создаём новый
                '--------------------------------------------------------------------------------
                If userDict.Contains(nameTableXRecord) Then
                    Dim xRecId As ObjectId = userDict.GetAt(nameTableXRecord)
                    If Not xRecId.IsNull Then
                        Dim xRec As Xrecord = TryCast(tr.GetObject(xRecId, OpenMode.ForWrite), Xrecord)
                        If xRec IsNot Nothing Then
                            ' Просто перезаписываем данные — без Erase/New/SetAt
                            Dim oldData As ResultBuffer = xRec.Data
                            xRec.Data = rb
                            If oldData IsNot Nothing Then oldData.Dispose()
                            tr.Commit()
                            Return True
                        End If
                    End If
                End If

                ' Нет существующего — создаём новый
                Dim newRec As New Xrecord()
                newRec.Data = rb
                userDict.SetAt(nameTableXRecord, newRec)
                tr.AddNewlyCreatedDBObject(newRec, True)
                tr.Commit()
                Return True
            End Using
        Catch ex As System.Exception
            System.Diagnostics.Trace.WriteLine("[writeXRecords] " & ex.ToString())
            Return False
        End Try
    End Function





    'чтение данных XRecords из примитива
    Public Shared Function readXRecords(ByVal acEnt As Autodesk.AutoCAD.DatabaseServices.Entity, Optional tr As Transaction = Nothing) As Dictionary(Of String, CadastralTools.StructureElement)
        Dim listDataEntity As Dictionary(Of String, CadastralTools.StructureElement) = New Dictionary(Of String, CadastralTools.StructureElement)
        If IsNothing(acEnt) = True Then Return Nothing
        If IsNothing(tr) = True Then
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            tr = doc.TransactionManager.StartTransaction()
        End If
        If IsNothing(tr) = False Then
            Using tr
                Dim extId As ObjectId = acEnt.ExtensionDictionary
                If extId.IsNull = False Then
                    Dim userDict As DBDictionary = CType(tr.GetObject(extId, OpenMode.ForRead), DBDictionary)
                    If userDict.Count > 0 Then
                        Dim enumValues As Array = [Enum].GetValues(GetType(tableXRecords))
                        For Each value As tableXRecords In enumValues
                            ' Получаем имя элемента
                            Dim name As String = [Enum].GetName(GetType(tableXRecords), value)
                            Dim IdkeyDict As ObjectId = userDict.GetAt(name)
                            Dim dataXRecord As StructureElement = New StructureElement()
                            If IsNothing(IdkeyDict) = False Then
                                If IdkeyDict.IsNull = False Then
                                    Dim rec As Xrecord = IdkeyDict.GetObject(OpenMode.ForRead)
                                    Dim rb As ResultBuffer = rec.Data
                                    Dim tVal As TypedValue() = rb.AsArray
                                    If tVal.Count > 0 Then
                                        Dim countArray As Integer = 0
                                        For j As Integer = 0 To tVal.Count - 1
                                            Dim valStr As String = tVal(j).ToString
                                            valStr = valStr.Replace("(1000,", "")
                                            valStr = valStr.Remove(valStr.Length - 1)
                                            Dim valArrayStr As String() = valStr.ToString.Split("=")
                                            If IsArray(valArrayStr) = True Then
                                                Dim nameField As String = valArrayStr(0)
                                                Dim lenZn As Integer = valArrayStr(0).Trim.Length
                                                Dim valueZn As String = Mid(valStr, lenZn + 2)
                                                Select Case nameField
                                                    Case "Label"
                                                        dataXRecord.Label = valueZn
                                                    Case "ClassObject"
                                                        Try
                                                            dataXRecord.NameClassObject = CType([Enum].Parse(GetType(CadastralTools.StructureElement.classObject), valueZn), CadastralTools.StructureElement.classObject)
                                                        Catch ex As TypeLoadException
                                                            dataXRecord.NameClassObject = classObject.OtherElement
                                                        Catch ex As ArgumentException
                                                            dataXRecord.NameClassObject = classObject.OtherElement
                                                        End Try
                                                    Case "TypeGeometry"
                                                        Try
                                                            dataXRecord.TypeGeometry = CType([Enum].Parse(GetType(CadastralTools.StructureElement.typeGeometryObject), valueZn), CadastralTools.StructureElement.typeGeometryObject)
                                                        Catch ex As TypeLoadException
                                                            dataXRecord.TypeGeometry = typeGeometryObject.none
                                                        Catch ex As ArgumentException
                                                            dataXRecord.TypeGeometry = typeGeometryObject.none
                                                        End Try
                                                        dataXRecord.TypeGeometry = CType([Enum].Parse(GetType(CadastralTools.StructureElement.typeGeometryObject), valueZn), CadastralTools.StructureElement.typeGeometryObject)
                                                    Case "NameTypeObject"
                                                        dataXRecord.NameTypeObject = CType([Enum].Parse(GetType(CadastralTools.StructureElement.nameObject), valueZn), CadastralTools.StructureElement.nameObject)
                                                    Case "Description"
                                                        dataXRecord.Description = valueZn
                                                    Case "tablePS"
                                                        dataXRecord.TablePS = valueZn
                                                    Case "KeyParameter"
                                                        dataXRecord.KeyParameter = valueZn
                                                    Case "IdStructure"
                                                        dataXRecord.IdStructure = valueZn
                                                    Case "IdElement"
                                                        dataXRecord.IdElement = valueZn
                                                    Case "Note"
                                                        dataXRecord.Note = valueZn
                                                    Case "IdObject"
                                                        dataXRecord.IdObject = acEnt.Handle.ToString
                                                End Select
                                            End If
                                        Next j
                                        If listDataEntity.ContainsKey(name) = False Then
                                            listDataEntity.Add(name, dataXRecord)
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        Return listDataEntity
                    End If
                End If
            End Using
        End If
        Return Nothing
    End Function
    'запись данных XRecords в примитив
    Public Shared Function writeXRecords(ByRef entityObject As Entity, ByRef dataRecords As StructureElement, Optional ByVal tableXRecords As CadastralTools.StructureElement.tableXRecords = tableXRecords.PROJECT_STRUCTURES) As Boolean
        writeXRecords = False
        Dim nameTableXRecord As String = tableXRecords.ToString
        If IsNothing(entityObject) = False Then
            If IsNothing(entityObject) = True Then Return False
            If TypeOf entityObject Is Entity Then
                Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim db As Database = doc.Database
                Dim jsonString As String = dataRecords.KeyParameter
                Try
                    If CadastralTools.JsonValidator.IsValidJson(jsonString) = True Then
                        Dim jsonObject As JObject = JObject.Parse(jsonString)
                        ' Проверяем, существует ли поле ListCoord
                        If jsonObject("ListCoord") IsNot Nothing Then
                            ' Удаляем поле ListCoord
                            jsonObject.Remove("ListCoord")
                            jsonString = jsonObject.ToString
                        End If
                    End If
                Catch ex As JsonException
                Catch ex As Exception
                End Try
                Try
                    Dim tr As Transaction = doc.TransactionManager.StartTransaction()
                    Using tr
                        Dim entity As Entity = entityObject
                        Dim arrayData As String(,) = {}
                        ReDim arrayData(1, 12)
                        arrayData(0, 0) = "Label"
                        arrayData(1, 0) = dataRecords.Label

                        arrayData(0, 1) = "ClassObject"
                        arrayData(1, 1) = dataRecords.NameClassObject.ToString

                        arrayData(0, 2) = "TypeGeometry"
                        arrayData(1, 2) = dataRecords.TypeGeometry.ToString

                        arrayData(0, 3) = "NameTypeObject"
                        arrayData(1, 3) = dataRecords.NameTypeObject.ToString

                        arrayData(0, 4) = "TableObject"
                        arrayData(1, 4) = dataRecords.TableObject.ToString

                        arrayData(0, 5) = "Description"
                        arrayData(1, 5) = dataRecords.Description

                        arrayData(0, 6) = "KeyParameter"
                        arrayData(1, 6) = jsonString

                        arrayData(0, 7) = "tablePS"
                        arrayData(1, 7) = dataRecords.TablePS

                        arrayData(0, 8) = "IdStructure"
                        arrayData(1, 8) = dataRecords.IdStructure

                        arrayData(0, 9) = "IdElement"
                        arrayData(1, 9) = dataRecords.IdElement

                        arrayData(0, 10) = "Note"
                        arrayData(1, 10) = dataRecords.Note

                        arrayData(0, 11) = "IdObject"
                        arrayData(1, 11) = entity.Handle.ToString

                        Dim extId As ObjectId = entity.ExtensionDictionary
                        If extId.IsNull = True Then
                            entity.UpgradeOpen()
                            entity.CreateExtensionDictionary()
                            extId = entity.ExtensionDictionary
                            Dim userDict As DBDictionary = CType(tr.GetObject(extId, OpenMode.ForRead), DBDictionary)
                            If IsNothing(userDict) = False Then
                                userDict.UpgradeOpen()
                                Dim xRec As Xrecord = New Xrecord()
                                Dim rb As ResultBuffer = New ResultBuffer()
                                For i As Integer = 0 To arrayData.GetUpperBound(1)
                                    Dim key As String = arrayData(0, i)
                                    Dim val As String = arrayData(1, i)
                                    If IsNothing(key) = False Then
                                        If key.Trim.Length > 0 Then
                                            If IsNothing(val) = True Then val = ""
                                            Dim recStr As String = key & "=" & val
                                            rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, recStr))
                                        End If
                                    End If
                                Next i
                                xRec.Data = rb
                                userDict.SetAt(nameTableXRecord, xRec)
                                tr.AddNewlyCreatedDBObject(xRec, True)
                                tr.Commit()
                                Return True
                            End If
                        Else
                            'считываем данные
                            Dim userDict As DBDictionary = CType(tr.GetObject(extId, OpenMode.ForWrite), DBDictionary)
                            If IsNothing(userDict) = False Then
                                If userDict.Contains(nameTableXRecord) = True Then
                                    Dim xRecId As ObjectId = userDict.Item(nameTableXRecord)
                                    If IsNothing(xRecId) = False Then
                                        If xRecId.IsNull = False Then
                                            Dim xRec As Xrecord = tr.GetObject(xRecId, OpenMode.ForWrite)
                                            xRec.Erase()
                                            If IsArray(arrayData) Then
                                                xRec = New Xrecord()
                                                Dim rb As ResultBuffer = New ResultBuffer()
                                                For i As Integer = 0 To arrayData.GetUpperBound(1)
                                                    Dim key As String = arrayData(0, i)
                                                    Dim val As String = arrayData(1, i)
                                                    If IsNothing(key) = False Then
                                                        If key.Trim.Length > 0 Then
                                                            If IsNothing(val) = True Then val = ""
                                                            Dim recStr As String = key & "=" & val
                                                            rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, recStr))
                                                        End If
                                                    End If
                                                Next i
                                                xRec.Data = rb
                                                userDict.SetAt(nameTableXRecord, xRec)
                                                tr.AddNewlyCreatedDBObject(xRec, True)
                                                tr.Commit()
                                                Return True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        Return False
                        'tr.Commit()
                    End Using
                Catch ex As System.Exception
                End Try
            End If
        End If
    End Function

End Class
