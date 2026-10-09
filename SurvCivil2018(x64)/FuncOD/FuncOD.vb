Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Gis.Map.HostMapApplicationServices
Imports Autodesk.Gis.Map.ObjectData
Imports Autodesk.Gis.Map
Public Class FuncOD
    '==================================================================================================================
    'функция создает новую таблицу OD
    Public Shared Function FuncCreateODTable(ByVal NameTable As String, ByVal ArrayFields As String(,)) As Boolean
        FuncCreateODTable = False
        Try
            If IsArray(ArrayFields) = False Then Exit Function
            '======================================================================================
            'открываем функционал Map3d
            Dim mapApp As MapApplication = HostMapApplicationServices.Application
            Dim activeProject As Project.ProjectModel = mapApp.ActiveProject
            'забираем все таблицы
            Dim tableList As ObjectData.Tables = mapApp.ActiveProject.ODTables 'забираем таблицы
            If tableList.Equals(NameTable) = False Then 'смотрим наличие такой таблицы по имени
                Dim fieldDefs As ObjectData.FieldDefinitions = mapApp.ActiveProject.MapUtility.NewODFieldDefinitions()
                Dim PosFiled As Integer = 0
                For i As Integer = 0 To ArrayFields.GetUpperBound(1)
                    Dim NameField As String = ArrayFields(1, i)
                    Dim DeskField As String = ArrayFields(3, i)
                    Dim TypeField As Integer = ArrayFields(4, i)
                    Dim typeOD As Constants.DataType
                    If TypeField Like "0" Then
                        typeOD = Constants.DataType.Integer
                    ElseIf TypeField Like "1" Then
                        typeOD = Constants.DataType.Real
                    Else
                        typeOD = Constants.DataType.Character
                    End If
                    Dim DefaultFiled As String = ArrayFields(5, i)
                    Dim val As String = ArrayFields(2, i)
                    'создаем поле
                    Try
                        Dim def0 As ObjectData.FieldDefinition = fieldDefs.Add(NameField, DeskField, typeOD, PosFiled)
                        PosFiled += 1
                    Catch ex As MapException
                    End Try
                Next i
                tableList.Add(NameTable, fieldDefs, "", True)
                fieldDefs.Dispose()
            End If
        Catch ex As MapException
            FuncCreateODTable = False
        Catch ex As System.Exception
            FuncCreateODTable = False
        End Try
        FuncCreateODTable = True
    End Function
    '==================================================================================================================
    '-Функция записывает ObjectData в примитив (данные записываются впервые)
    Public Shared Function FuncAddOD(ByVal DataArray(,) As String, ByVal NameTableObj As String, ByVal objID As ObjectId) As Boolean
        FuncAddOD = False
        If IsArray(DataArray) = False Then Exit Function
        'ссылки на объекты map3d
        Try
            Dim mapApp As MapApplication = HostMapApplicationServices.Application
            Dim activeProject As Project.ProjectModel = mapApp.ActiveProject
            Dim tableList As ObjectData.Tables = activeProject.ODTables 'забираем таблицы
            If tableList.Equals(NameTableObj) = False Then 'смотрим наличие такой таблицы по имени
                'если таблица существует начинаем запись данных
                Dim tableOBJ As ObjectData.Table = mapApp.ActiveProject.ODTables.Item(NameTableObj)
                Dim odRecords As ObjectData.Records = tableOBJ.GetObjectTableRecords(Convert.ToInt32(0), objID, Constants.OpenMode.OpenForRead, False)
                Dim odRecord As ObjectData.Record = odRecords.Item(0)
                odRecord = Autodesk.Gis.Map.ObjectData.Record.Create()
                tableOBJ.InitRecord(odRecord)
                Dim mapVal As Autodesk.Gis.Map.Utilities.MapValue
                'читаем поля таблицы
                Dim recTable As Autodesk.Gis.Map.ObjectData.FieldDefinitions = tableOBJ.FieldDefinitions 'читаем поля
                For i As Integer = 0 To recTable.Count - 1
                    Dim Fl As FieldDefinition = recTable.Item(i)
                    Dim NameField As String = Fl.Name 'имя поля
                    mapVal = odRecord(i)
                    Dim boolFlag As Boolean = False
                    For k As Integer = 0 To DataArray.GetUpperBound(1)
                        Dim NameFl As String = DataArray(1, k) 'имя поля из массива
                        Dim NewVAl As String = DataArray(2, k) 'значение которое необходимо записать
                        If NameFl Like NameField Then 'если названия полей совпадают заносим значение
                            If Fl.Type = Constants.DataType.Integer Then 'если это целое число 
                                Dim IntRec As Integer = Val(NewVAl)
                                mapVal.Assign(IntRec)
                                boolFlag = True
                                Exit For
                            ElseIf Fl.Type = Constants.DataType.Real Then
                                Dim IntRec As Double = Val(NewVAl)
                                mapVal.Assign(IntRec)
                                boolFlag = True
                                Exit For
                            Else
                                mapVal.Assign(NewVAl)
                                boolFlag = True
                                Exit For
                            End If
                        End If
                    Next k
                    If boolFlag = False Then
                        If Fl.Type = Constants.DataType.Integer Then
                            Dim IntRec As Integer = 0
                            mapVal.Assign(IntRec)
                        ElseIf Fl.Type = Constants.DataType.Real Then
                            Dim IntRec As Double = 0
                            mapVal.Assign(IntRec)
                        Else
                            mapVal.Assign("")
                        End If
                    End If
                Next i
                tableOBJ.AddRecord(odRecord, objID)
                recTable.Dispose()
                odRecord.Dispose()
                odRecords.Dispose()
                tableOBJ.Dispose()
                FuncAddOD = True
            End If
        Catch ex As MapException
        End Try
    End Function
    '-=====================================================================================
    '-Функция перезаписыват ObjectData в примитив (обновляет)
    Public Shared Function FuncReplaseOD(ByVal DataArray(,) As String, ByVal NameTableObj As String, ByVal objID As ObjectId) As Boolean
        FuncReplaseOD = False
        If IsArray(DataArray) = False Then Exit Function
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim mapApp As MapApplication = HostMapApplicationServices.Application
        Dim activeProject As Project.ProjectModel = mapApp.ActiveProject
        Dim tableList As ObjectData.Tables = activeProject.ODTables
        Dim countTablelist As Integer = tableList.TablesCount
        Dim acTrans As Transaction = db.TransactionManager.StartTransaction()
        Using acTrans
            Try
                Dim tableOBJ As ObjectData.Table = mapApp.ActiveProject.ODTables.Item(NameTableObj)
                Dim odRecords As ObjectData.Records = tableOBJ.GetObjectTableRecords(Convert.ToInt32(0), objID, Constants.OpenMode.OpenForWrite, False)
                Dim odRecord As ObjectData.Record = odRecords.Item(0)
                odRecord = Autodesk.Gis.Map.ObjectData.Record.Create()
                tableOBJ.InitRecord(odRecord)
                'читаем поля таблицы
                Dim recTable As Autodesk.Gis.Map.ObjectData.FieldDefinitions = tableOBJ.FieldDefinitions
                For Each rec As ObjectData.Record In odRecords
                    For i As Integer = 0 To rec.Count() - 1
                        Dim mapVal As Autodesk.Gis.Map.Utilities.MapValue = rec(i)
                        Dim mapField As ObjectData.FieldDefinition = recTable.Item(i)
                        For k As Integer = 0 To DataArray.GetUpperBound(1)
                            Dim NameFl As String = DataArray(1, k)
                            Dim NewVAl As String = DataArray(2, k)
                            If mapField.Name Like NameFl Then
                                If mapVal.Type = Constants.DataType.Integer Then
                                    Dim IntRec As Integer = Val(NewVAl)
                                    mapVal.Assign(IntRec)
                                ElseIf mapVal.Type = Constants.DataType.Real Then
                                    Dim IntRec As Double = Val(NewVAl)
                                    mapVal.Assign(IntRec)
                                Else
                                    mapVal.Assign(NewVAl)
                                End If
                            End If
                        Next k
                    Next
                    odRecords.UpdateRecord(rec)
                Next
                odRecord.Dispose()
                odRecords.Dispose()
            Catch ex As MapException
                Exit Function
            Catch ex As System.Exception
                Exit Function
            End Try
            acTrans.Commit()
        End Using
        FuncReplaseOD = True
    End Function


    '===============================================================================================
    '-Функция возвращает имена полей таблицы Расширенных данных (имя поля)
    Public Shared Function FuncReadFieldOD(ByRef DataArray() As String, ByVal NameTableObj As String) As Boolean
        FuncReadFieldOD = False
        Dim CountArray As Integer = 0
        If IsNothing(DataArray) = False Then
            CountArray = DataArray.Length - 1
        End If
        Dim mapApp As MapApplication = HostMapApplicationServices.Application
        Dim activeProject As Project.ProjectModel = mapApp.ActiveProject
        Dim tableList As ObjectData.Tables = activeProject.ODTables
        Try
            Dim tableOBJ As ObjectData.Table = mapApp.ActiveProject.ODTables.Item(NameTableObj)
            Dim recTable As Autodesk.Gis.Map.ObjectData.FieldDefinitions = tableOBJ.FieldDefinitions
            For i As Integer = 0 To recTable.Count - 1
                Dim fieldDefs As ObjectData.FieldDefinition = recTable.Item(i)
                ReDim Preserve DataArray(CountArray)
                DataArray(CountArray) = fieldDefs.Name
                CountArray += 1
            Next i
        Catch ex As MapException
            Exit Function
        End Try
        FuncReadFieldOD = True
    End Function

    '===============================================================================================
    '-Функция возвращает имена полей таблицы Расширенных данных (имя поля,дескриптор)
    Public Shared Function FuncReadFieldOD2(ByRef DataArray(,) As String, ByVal NameTableObj As String) As Boolean
        FuncReadFieldOD2 = False
        Dim CountArray As Integer = 0
        If IsNothing(DataArray) = False Then
            CountArray = DataArray.Length - 1
        End If
        Dim mapApp As MapApplication = HostMapApplicationServices.Application
        Dim activeProject As Project.ProjectModel = mapApp.ActiveProject
        Dim tableList As ObjectData.Tables = activeProject.ODTables
        Try
            Dim tableOBJ As ObjectData.Table = mapApp.ActiveProject.ODTables.Item(NameTableObj)
            Dim recTable As Autodesk.Gis.Map.ObjectData.FieldDefinitions = tableOBJ.FieldDefinitions
            For i As Integer = 0 To recTable.Count - 1
                Dim fieldDefs As ObjectData.FieldDefinition = recTable.Item(i)
                ReDim Preserve DataArray(1, CountArray)
                DataArray(0, CountArray) = fieldDefs.Name
                DataArray(1, CountArray) = fieldDefs.Description
                CountArray += 1
            Next i
        Catch ex As MapException
            Exit Function
        End Try
        FuncReadFieldOD2 = True
    End Function
    '===============================================================================================
    '-Функция возвращает OD данные из примитива (значение)
    Public Shared Function FuncReadValOD(ByRef DataArray() As String, ByVal NameTable As String, ByVal objID As ObjectId) As Boolean
        FuncReadValOD = False
        Dim CountDataArray As Integer = 0
        Erase DataArray
        Dim mapApp As MapApplication = HostMapApplicationServices.Application
        Dim activeProject As Project.ProjectModel = mapApp.ActiveProject
        Dim tableList As ObjectData.Tables = activeProject.ODTables
        Dim tlb As ObjectData.Table = tableList.Item(NameTable)
        Try
            Dim recs As ObjectData.Records = tlb.GetObjectTableRecords(0, objID, Constants.OpenMode.OpenForRead, True)
            'считываем поля таблицы
            Dim fieldDefs As ObjectData.FieldDefinitions = tlb.FieldDefinitions
            For Each rec As ObjectData.Record In recs
                For j As Integer = 0 To rec.Count() - 1
                    Dim val As Autodesk.Gis.Map.Utilities.MapValue = rec(j)
                    ReDim Preserve DataArray(CountDataArray)
                    Select Case val.Type
                        Case Constants.DataType.Character
                            DataArray(CountDataArray) = val.StrValue
                        Case Constants.DataType.Integer
                            DataArray(CountDataArray) = val.Int32Value.ToString
                        Case Constants.DataType.Real
                            DataArray(CountDataArray) = val.DoubleValue.ToString
                        Case Else
                            DataArray(CountDataArray) = "undefined"
                    End Select
                    CountDataArray += 1
                Next
            Next
            recs.Dispose()
        Catch ex As MapException
            Exit Function
        End Try
        FuncReadValOD = True
    End Function
    '-=====================================================================================
    'функция возвращает имена таблиц которые были записаны в примитив
    Public Shared Function FuncFindTableOD(ByRef DataArray() As String, ByVal obj As Autodesk.AutoCAD.DatabaseServices.DBObject) As Boolean
        FuncFindTableOD = False
        Dim CountTAble As Integer = 0
        If IsNothing(DataArray) = False Then CountTAble = DataArray.Length
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim mapApp As MapApplication = HostMapApplicationServices.Application
        Dim activeProject As Project.ProjectModel = mapApp.ActiveProject
        Dim tableList As ObjectData.Tables = activeProject.ODTables
        Dim NameTables As System.Collections.Specialized.StringCollection = tableList.GetTableNames
        Dim countTablelist As Integer = tableList.TablesCount
        For i As Integer = 0 To tableList.TablesCount - 1
            Dim NameTable As String = NameTables.Item(i)
            Dim tlb As ObjectData.Table = tableList.Item(NameTable)
            Try
                Dim recs As ObjectData.Records = tlb.GetObjectTableRecords(0, obj.Id, Constants.OpenMode.OpenForRead, True)
                Dim countRec As Integer = recs.Count
                If countRec > 0 Then
                    ReDim Preserve DataArray(CountTAble)
                    DataArray(CountTAble) = tlb.Name
                    CountTAble += 1
                    FuncFindTableOD = True
                End If
                recs.Dispose()
            Catch ex As MapException
            End Try
        Next i
        tr.Commit()
    End Function
End Class
