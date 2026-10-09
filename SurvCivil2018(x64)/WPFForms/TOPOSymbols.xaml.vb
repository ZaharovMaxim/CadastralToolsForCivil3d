Imports System.Windows
Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil.DatabaseServices
Imports Label = Autodesk.Civil.DatabaseServices.Label
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Environment
Imports System.Text.RegularExpressions
Imports Autodesk.AutoCAD.LayerManager
Imports System.Collections.ObjectModel
Imports System.Windows.Input

Public Class TOPOSymbols
    Public ps As Boolean = False
    Public UserProperties As UserPropertiesTopoSymbols = Nothing
    Private components1 As System.ComponentModel.IContainer
    Public ArrayTopoCategory As String(,) = Nothing 'массив с категориями условных знаков (0-кооддината X таблицы для сортировки, 1-Наименование категории, 2-Handle таблицы с линиями, 3-Handle таблицы с Блоками, 4-Handle таблицы с COGO точками)
    Public ArrayParametrObj As String() = Nothing 'параметры объекта
    Public PutchImageObject As String = "d:\Civil3DToolsUtility\FileResources\ImageObject\" 'путь к изображениям объектов
    Public PutchDorImage As String = "d:\Civil3DToolsUtility\FileResources\RoadZnaks\" 'пути для миниатюр дорожных знаков
    Public PutchTemplateDwg As String = "d:\Civil3DToolsUtility\FileResources\Sample\" 'пути для файла шаблонов
    Public PutchFileTemplateDwg As String = ""
    Public Shared supportPatchFolder As String() = Nothing
    Public Shared custdata As ObservableCollection(Of Customer) = New ObservableCollection(Of Customer)
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'столбцы
    Public Class Customer
        Public Property Имя_свойства As String
        Public Property Значение As String
    End Class
    '========================================================================================
    'функция возвращает массив со всеми таблицами с категориями условных знаков из фаблона шаблона (коордY,Имя категории таблицы,Handle таблиц)
    Public Shared Function FuncListGlobalCategoryTBLDestDWG(ByVal PutshDraw As String, ByRef ArrayTableCategory(,) As String) As Boolean
        FuncListGlobalCategoryTBLDestDWG = False
        If IsArray(ArrayTableCategory) = True Then
            Erase ArrayTableCategory
        End If
        Dim CountArrayTableCategory As Integer = 0
        Dim sourceDb As Database = New Database(False, True)
        If File.Exists(PutshDraw) = False Then
            Return False
        End If
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
            For Each id As ObjectId In sourceBtr
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Table Then
                    Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = ent
                    If tbl.Rows.Count > 3 And tbl.Columns.Count > 2 Then
                        Dim row1 As Autodesk.AutoCAD.DatabaseServices.Row = tbl.Rows.Item(0)
                        Dim row2 As Autodesk.AutoCAD.DatabaseServices.Row = tbl.Rows.Item(1)
                        If row1.IsMerged = True Then
                            Dim NameGlobalCategory As String = tbl.Cells(0, 0).Value 'глобальная категория
                            Dim CategoryPosition As Integer = -1
                            If IsArray(ArrayTableCategory) = False Then
                                ReDim Preserve ArrayTableCategory(2, CountArrayTableCategory)
                                ArrayTableCategory(0, CountArrayTableCategory) = tbl.Position.Y
                                ArrayTableCategory(1, CountArrayTableCategory) = NameGlobalCategory
                                ArrayTableCategory(2, CountArrayTableCategory) = tbl.Handle.ToString
                                CountArrayTableCategory += 1
                                CategoryPosition = 0
                            Else
                                For i As Integer = 0 To ArrayTableCategory.GetUpperBound(1)
                                    Dim TempNameCategory As String = ArrayTableCategory(1, i)
                                    If TempNameCategory.Trim Like NameGlobalCategory.Trim Then
                                        Dim TempZn As String = ArrayTableCategory(2, i)
                                        ArrayTableCategory(2, i) = TempZn & " " & tbl.Handle.ToString
                                        CategoryPosition = i
                                        Exit For
                                    End If
                                Next i
                            End If
                            If CategoryPosition = -1 Then
                                ReDim Preserve ArrayTableCategory(2, CountArrayTableCategory)
                                ArrayTableCategory(0, CountArrayTableCategory) = tbl.Position.Y
                                ArrayTableCategory(1, CountArrayTableCategory) = NameGlobalCategory
                                ArrayTableCategory(2, CountArrayTableCategory) = tbl.Handle.ToString
                                CategoryPosition = CountArrayTableCategory
                                CountArrayTableCategory += 1
                            End If
                        End If
                    End If
                End If
            Next
        End Using
        If IsArray(ArrayTableCategory) = True Then
            FuncListGlobalCategoryTBLDestDWG = True
        End If
        sourceDb.CloseInput(True)
    End Function
    '========================================================================================
    'функция возвращает массив со всеми таблицами с категориями условных знаков из фаблона шаблона (коордX,Имя таблицы,Handle таблиц)
    Public Shared Function FuncListTOPOCategoryTBLDestDWG(ByVal PutshDraw As String, ByVal HandleTables As String, ByRef ArrayTableCategory(,) As String) As Boolean
        FuncListTOPOCategoryTBLDestDWG = False
        If IsArray(ArrayTableCategory) = True Then
            Erase ArrayTableCategory
        End If
        Dim CountArrayTableCategory As Integer = 0
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try
        If IsNothing(HandleTables) = True Then
            Return False
        End If
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
            Dim ArrayStrHandle As String() = HandleTables.Split
            For i As Integer = 0 To ArrayStrHandle.Length - 1
                Dim StrHgTbl As String = ArrayStrHandle(i).Trim
                If StrHgTbl.Length >= 0 Then
                    Dim TableID As ObjectId = Nothing
                    Try
                        Dim ln As Long = Convert.ToInt64(StrHgTbl, 16)
                        Dim hn As Handle = New Handle(ln)
                        TableID = sourceDb.GetObjectId(False, hn, 0)
                    Catch ex As ArgumentOutOfRangeException
                        Continue For
                    End Try
                    If IsNothing(TableID) = True Then
                        Continue For
                    End If
                    If TableID.IsNull = True Then
                        Continue For
                    End If
                    Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = TableID.GetObject(OpenMode.ForRead)
                    If tbl.Rows.Count > 3 And tbl.Columns.Count > 2 Then
                        Dim row2 As Autodesk.AutoCAD.DatabaseServices.Row = tbl.Rows.Item(1)
                        If row2.IsMerged = True Then
                            Dim NameCategory As String = tbl.Cells(1, 0).Value 'глобальная категория
                            Dim CategoryPosition As Integer = -1
                            If IsArray(ArrayTableCategory) = False Then
                                ReDim Preserve ArrayTableCategory(2, CountArrayTableCategory)
                                ArrayTableCategory(0, CountArrayTableCategory) = tbl.Position.X
                                ArrayTableCategory(1, CountArrayTableCategory) = NameCategory
                                ArrayTableCategory(2, CountArrayTableCategory) = tbl.Handle.ToString
                                CountArrayTableCategory += 1
                                CategoryPosition = 0
                            Else
                                For i1 As Integer = 0 To ArrayTableCategory.GetUpperBound(1)
                                    Dim TempNameCategory As String = ArrayTableCategory(1, i1)
                                    If TempNameCategory.Trim Like NameCategory.Trim Then
                                        Dim TempZn As String = ArrayTableCategory(2, i1)
                                        ArrayTableCategory(2, i1) = TempZn & " " & tbl.Handle.ToString
                                        CategoryPosition = i
                                        Exit For
                                    End If
                                Next i1
                            End If
                            If CategoryPosition = -1 Then
                                ReDim Preserve ArrayTableCategory(2, CountArrayTableCategory)
                                ArrayTableCategory(0, CountArrayTableCategory) = tbl.Position.X
                                ArrayTableCategory(1, CountArrayTableCategory) = NameCategory
                                ArrayTableCategory(2, CountArrayTableCategory) = tbl.Handle.ToString
                                CategoryPosition = CountArrayTableCategory
                                CountArrayTableCategory += 1
                            End If
                        End If
                    End If
                End If
            Next i
        End Using
        If IsArray(ArrayTableCategory) = True Then
            FuncListTOPOCategoryTBLDestDWG = True
        End If
        sourceDb.CloseInput(True)
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция определяет типы объектов которые входят в данную категорию (Полилиния, Блок, ТочкаCOGO и т.п.)
    Public Shared Function FuncListTypeObjectDestDWG(ByVal PutshDraw As String, ByVal StringTableHandle As String, ByRef ArrayTableTypeObject(,) As String) As Boolean
        FuncListTypeObjectDestDWG = False
        If IsArray(ArrayTableTypeObject) = True Then
            Erase ArrayTableTypeObject
        End If
        Dim CountArrayTableTypeObject As Integer = 0

        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try

        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
            Dim ArrayTypeObject As String() = StringTableHandle.Split
            If IsArray(ArrayTypeObject) = True Then
                For i As Integer = 0 To ArrayTypeObject.Length - 1
                    Dim StrHandle As String = ArrayTypeObject(i)
                    If IsNothing(StrHandle) = True Then
                        Continue For
                    End If
                    If StrHandle.Trim.Length = 0 Then
                        Continue For
                    End If
                    Dim TableID As ObjectId = Nothing
                    Try
                        Dim ln As Long = Convert.ToInt64(StrHandle, 16)
                        Dim hn As Handle = New Handle(ln)
                        TableID = sourceDb.GetObjectId(False, hn, 0)
                    Catch ex As ArgumentOutOfRangeException
                        Continue For
                    End Try
                    If IsNothing(TableID) = True Then
                        Continue For
                    End If
                    If TableID.IsNull = True Then
                        Continue For
                    End If
                    Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = TableID.GetObject(OpenMode.ForRead)
                    For j As Integer = 0 To tbl.Columns.Count - 1
                        Dim NameType As String = tbl.Cells(2, j).Value 'читаем категорию
                        If NameType.Trim Like "Тип объекта" Then
                            For k As Integer = 3 To tbl.Rows.Count - 1
                                Dim NameTypeObject As String = tbl.Cells(k, j).Value
                                If IsNothing(NameTypeObject) = False Then
                                    If NameTypeObject.Trim.Length > 0 Then
                                        If IsArray(ArrayTableTypeObject) = False Then
                                            ReDim Preserve ArrayTableTypeObject(1, CountArrayTableTypeObject)
                                            ArrayTableTypeObject(0, CountArrayTableTypeObject) = NameTypeObject
                                            ArrayTableTypeObject(1, CountArrayTableTypeObject) = tbl.Handle.ToString
                                            CountArrayTableTypeObject += 1
                                        Else
                                            Dim boolFind As Boolean = False
                                            For k2 As Integer = 0 To ArrayTableTypeObject.GetUpperBound(1)
                                                Dim TempTypeOblect As String = ArrayTableTypeObject(0, k2)
                                                If TempTypeOblect Like NameTypeObject Then
                                                    boolFind = True
                                                    Exit For
                                                End If
                                            Next k2
                                            If boolFind = False Then
                                                ReDim Preserve ArrayTableTypeObject(1, CountArrayTableTypeObject)
                                                ArrayTableTypeObject(0, CountArrayTableTypeObject) = NameTypeObject
                                                ArrayTableTypeObject(1, CountArrayTableTypeObject) = tbl.Handle.ToString
                                                CountArrayTableTypeObject += 1
                                            End If
                                        End If
                                    End If
                                End If
                            Next k
                        End If
                    Next j
                Next i
            End If
        End Using
        If IsArray(ArrayTableTypeObject) = True Then
            FuncListTypeObjectDestDWG = True
        End If
        sourceDb.CloseInput(True)
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает все строки таблицы определенного типа объекта
    Public Shared Function FuncListObjectDestDWG(ByVal PutshDraw As String, ByVal StringTableHandle As String, ByRef ArrayTableTypeObject(,) As String) As Boolean
        FuncListObjectDestDWG = False
        If IsArray(ArrayTableTypeObject) = True Then
            Erase ArrayTableTypeObject
        End If
        Dim CountArrayTableTypeObject As Integer = 0
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try

        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
            Dim StrHandle As String = StringTableHandle.Trim
            If IsNothing(StrHandle) = False Then
                If StrHandle.Length > 0 Then
                    Try
                        Dim ln As Long = Convert.ToInt64(StrHandle, 16)
                        Dim hn As Handle = New Handle(ln)
                        Dim TableID As ObjectId = sourceDb.GetObjectId(False, hn, 0)
                        If TableID.IsNull = False Then
                            Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = TableID.GetObject(OpenMode.ForRead)
                            Dim CountColumns As Integer = tbl.Columns.Count - 1
                            ReDim Preserve ArrayTableTypeObject(CountColumns, 0)
                            For j As Integer = 0 To tbl.Columns.Count - 1
                                Dim valStr As String = tbl.Cells(2, j).Value
                                If IsNothing(valStr) = False Then
                                    If valStr.Trim Like "Номер" Then
                                        valStr = "-999999"
                                    End If
                                    ArrayTableTypeObject(j, 0) = valStr
                                Else
                                    ArrayTableTypeObject(j, 0) = "-"
                                End If
                            Next j
                            CountArrayTableTypeObject += 1
                            For k As Integer = 3 To tbl.Rows.Count - 1
                                Dim ValStrTemp As String = tbl.Cells(k, 0).Value
                                If IsNothing(ValStrTemp) = True Then
                                    Continue For
                                ElseIf ValStrTemp.Trim.Length = 0 Then
                                    Continue For
                                Else
                                    If IsNumeric(ValStrTemp) = False Then
                                        ValStrTemp = "0"
                                    End If
                                End If
                                ReDim Preserve ArrayTableTypeObject(CountColumns, CountArrayTableTypeObject)
                                For k2 As Integer = 0 To tbl.Columns.Count - 1
                                    Dim ValStr As String = tbl.Cells(k, k2).Value
                                    If IsNothing(ValStr) = False Then
                                        ArrayTableTypeObject(k2, CountArrayTableTypeObject) = ValStr
                                    End If
                                Next k2
                                CountArrayTableTypeObject += 1
                            Next k
                        End If
                    Catch ex As Exception
                    End Try
                End If
            End If
        End Using
        FuncListObjectDestDWG = True
        sourceDb.CloseInput(True)
    End Function
    '===========================================================================================================================================================================================
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция создает файл xml на основе файла DWG с условными знаками(Имя категории, Имя таблицы, Тип объекта, Характеристики)
    Public Shared Function FuncCreateXMLAllListObject(ByVal putchDwg As String, ByRef putchFileXml As String) As Boolean
        FuncCreateXMLAllListObject = False
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(putchDwg, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'получаем категории
        Dim ArrayTableCategory As String(,) = Nothing
        Dim CountArrayTableCategory As Integer = 0
        'словари
        'типы линий
        Dim dictionaryTypeLine As Dictionary(Of String, String()) = New Dictionary(Of String, String())
        Dim arrayTypeLine As String() = {"ПоСлою", "ПоБлоку", "Continuous", "CENTER", "CENTER2", "DIVIDE", "DASHDOT", "DASHED", "DASHED2", "DOT", "HIDDEN", "HIDDEN2"}
        dictionaryTypeLine.Add("Стандартные", arrayTypeLine)
        'словари
        Dim dictionaryPS As Dictionary(Of String, String()) = New Dictionary(Of String, String())
        Dim ArrayPSType As String() = {"OBJECT", "IS_SYSTEM"}
        dictionaryPS.Add("Стандартные", ArrayPSType)
        'стили текста
        'Dim dictionaryTextStyle As Dictionary(Of String, String()) = New Dictionary(Of String, String())
        'Dim ArrayTextStyle As String() = {"ESKDU", "Legend", "Standard", "СПДС"}
        'dictionaryTextStyle.Add("Стандартные", ArrayTextStyle)
        'стили мультивыносок
        Dim dictionaryMLeaderStyle As Dictionary(Of String, String()) = New Dictionary(Of String, String())
        Dim ArrayMLeaderStyle As String() = {"ESKDU", "Standard"}
        dictionaryMLeaderStyle.Add("Стандартные", ArrayMLeaderStyle)

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
            For Each id As ObjectId In sourceBtr
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Table Then
                    Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = ent
                    If tbl.Rows.Count > 3 And tbl.Columns.Count > 2 Then
                        Dim row1 As Autodesk.AutoCAD.DatabaseServices.Row = tbl.Rows.Item(0)
                        Dim row2 As Autodesk.AutoCAD.DatabaseServices.Row = tbl.Rows.Item(1)
                        If row1.IsMerged = True Then
                            Dim NameGlobalCategory As String = tbl.Cells(0, 0).Value 'глобальная категория
                            If IsNothing(NameGlobalCategory) = False Then
                                Dim CategoryPosition As Integer = -1
                                If IsArray(ArrayTableCategory) = False Then
                                    ReDim Preserve ArrayTableCategory(2, CountArrayTableCategory)
                                    ArrayTableCategory(0, CountArrayTableCategory) = tbl.Position.Y
                                    ArrayTableCategory(1, CountArrayTableCategory) = NameGlobalCategory.Trim
                                    ArrayTableCategory(2, CountArrayTableCategory) = tbl.Handle.ToString
                                    CountArrayTableCategory += 1
                                    CategoryPosition = 0
                                Else
                                    For i As Integer = 0 To ArrayTableCategory.GetUpperBound(1)
                                        Dim TempNameCategory As String = ArrayTableCategory(1, i)
                                        If TempNameCategory.Trim Like NameGlobalCategory.Trim Then
                                            Dim TempZn As String = ArrayTableCategory(2, i)
                                            ArrayTableCategory(2, i) = TempZn & " " & tbl.Handle.ToString
                                            CategoryPosition = i
                                            Exit For
                                        End If
                                    Next i
                                End If
                                If CategoryPosition = -1 Then
                                    ReDim Preserve ArrayTableCategory(2, CountArrayTableCategory)
                                    ArrayTableCategory(0, CountArrayTableCategory) = tbl.Position.Y
                                    ArrayTableCategory(1, CountArrayTableCategory) = NameGlobalCategory.Trim
                                    ArrayTableCategory(2, CountArrayTableCategory) = tbl.Handle.ToString
                                    CategoryPosition = CountArrayTableCategory
                                    CountArrayTableCategory += 1
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            If IsArray(ArrayTableCategory) = True Then
                If ArrayTableCategory.GetUpperBound(1) > 2 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayTableCategory, 0)
                End If
            End If
        End Using
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'получем все таблицы с категориями
        Dim arrayUserObject As String(,) = Nothing
        Dim countArrayUserObject As Integer = 0
        If IsArray(ArrayTableCategory) = True Then
            Dim xws As XmlWriterSettings = New XmlWriterSettings()
            xws.Indent = True
            xws.NewLineOnAttributes = True
            Dim xw As XmlWriter = XmlWriter.Create(putchFileXml, xws)
            Dim sourceTr2 As Transaction = sourceDb.TransactionManager.StartTransaction()
            Using sourceTr2
                Dim sourseBlockTable As BlockTable = sourceTr2.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
                Dim sourceBtr As BlockTableRecord = sourceTr2.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForWrite)
                'Dim sourceBlkTblRec As BlockTableRecord = sourceTr2.GetObject(sourceBtr(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim acLineTypTbl As LinetypeTable = sourceTr2.GetObject(sourceDb.LinetypeTableId, OpenMode.ForRead)
                xw.WriteStartDocument()
                xw.WriteStartElement("Classifier_ADSK")
                xw.WriteAttributeString("NameClassifier", "Стройпроект")
                'создаем ветку с услвными знаками
                xw.WriteStartElement("LibrarySymbols")
                xw.WriteAttributeString("NameLibrarySymbols", "Условные знаки")
                xw.WriteAttributeString("Version", "2")
                For i As Integer = 0 To ArrayTableCategory.GetUpperBound(1)
                    Dim ArrayStrHandle As String() = ArrayTableCategory(2, i).Split()
                    Dim tableCategory As String = ArrayTableCategory(1, i).Trim
                    Dim ArrayTable As String(,) = Nothing
                    Dim CountArrayTable As Integer = 0
                    If IsArray(ArrayStrHandle) = True Then
                        xw.WriteStartElement("Category")
                        xw.WriteAttributeString("CategoryName", ArrayTableCategory(1, i))
                    Else
                        Continue For
                    End If
                    'получаем все имена таблиц
                    For j As Integer = 0 To ArrayStrHandle.Length - 1
                        Dim StrHgTbl As String = ArrayStrHandle(j).Trim
                        If StrHgTbl.Length >= 0 Then
                            Dim TableID As ObjectId = Nothing
                            Try
                                Dim ln As Long = Convert.ToInt64(StrHgTbl, 16)
                                Dim hn As Handle = New Handle(ln)
                                TableID = sourceDb.GetObjectId(False, hn, 0)
                            Catch ex As ArgumentOutOfRangeException
                                Continue For
                            End Try
                            If IsNothing(TableID) = True Then
                                Continue For
                            End If
                            If TableID.IsNull = True Then
                                Continue For
                            End If
                            Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = TableID.GetObject(OpenMode.ForRead)
                            If tbl.Rows.Count > 3 And tbl.Columns.Count > 2 Then
                                Dim row2 As Autodesk.AutoCAD.DatabaseServices.Row = tbl.Rows.Item(1)
                                If row2.IsMerged = True Then
                                    Dim NameTable As String = tbl.Cells(1, 0).Value 'имя таблицы
                                    If IsNothing(NameTable) = False Then
                                        Dim CategoryPosition As Integer = -1
                                        If IsArray(ArrayTable) = False Then
                                            ReDim Preserve ArrayTable(2, CountArrayTable)
                                            ArrayTable(0, CountArrayTable) = tbl.Position.X
                                            ArrayTable(1, CountArrayTable) = NameTable.Trim
                                            ArrayTable(2, CountArrayTable) = tbl.Handle.ToString
                                            CountArrayTable += 1
                                            CategoryPosition = 0
                                        Else
                                            For j1 As Integer = 0 To ArrayTable.GetUpperBound(1)
                                                Dim TempNameCategory As String = ArrayTable(1, j1)
                                                If TempNameCategory.Trim Like NameTable.Trim Then
                                                    Dim TempZn As String = ArrayTable(2, j1)
                                                    ArrayTable(2, j1) = TempZn & " " & tbl.Handle.ToString
                                                    CategoryPosition = j1
                                                    Exit For
                                                End If
                                            Next j1
                                        End If
                                        If CategoryPosition = -1 Then
                                            ReDim Preserve ArrayTable(2, CountArrayTable)
                                            ArrayTable(0, CountArrayTable) = tbl.Position.X
                                            ArrayTable(1, CountArrayTable) = NameTable.Trim
                                            ArrayTable(2, CountArrayTable) = tbl.Handle.ToString
                                            CategoryPosition = CountArrayTable
                                            CountArrayTable += 1
                                        End If
                                    End If

                                End If
                            End If
                        End If
                    Next j
                    ''==============================================================================================
                    ''читем строки таблицы для типов объекта
                    If IsArray(ArrayTable) = True Then
                        If ArrayTableCategory.GetUpperBound(1) > 2 Then
                            Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayTable, 0)
                        End If
                        xw.WriteStartElement("Tables")
                        For j As Integer = 0 To ArrayTable.GetUpperBound(1)
                            Dim ArrayTablesHandle As String() = ArrayTable(2, j).Split()
                            Dim ArrayTypeObject As String(,) = Nothing
                            Dim CountArrayTypeObject As Integer = 0
                            If IsArray(ArrayTablesHandle) = True Then
                                xw.WriteStartElement("Table")
                                xw.WriteAttributeString("TableName", ArrayTable(1, j))
                                If IsArray(ArrayTablesHandle) = True Then
                                    For k As Integer = 0 To ArrayTablesHandle.Length - 1
                                        Dim StrHandle As String = ArrayTablesHandle(k)
                                        If IsNothing(StrHandle) = True Then
                                            Continue For
                                        End If
                                        If StrHandle.Trim.Length = 0 Then
                                            Continue For
                                        End If
                                        Dim TableID As ObjectId = Nothing
                                        Try
                                            Dim ln As Long = Convert.ToInt64(StrHandle, 16)
                                            Dim hn As Handle = New Handle(ln)
                                            TableID = sourceDb.GetObjectId(False, hn, 0)
                                        Catch ex As ArgumentOutOfRangeException
                                            Continue For
                                        End Try
                                        If IsNothing(TableID) = True Then
                                            Continue For
                                        End If
                                        If TableID.IsNull = True Then
                                            Continue For
                                        End If
                                        Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = TableID.GetObject(OpenMode.ForRead)
                                        For k1 As Integer = 0 To tbl.Columns.Count - 1
                                            Dim NameType As String = tbl.Cells(2, k1).Value 'читаем тип объекта
                                            If IsNothing(NameType) = False Then
                                                If NameType.Trim Like "Type" Then
                                                    For k2 As Integer = 3 To tbl.Rows.Count - 1
                                                        Dim NameTypeObject As String = tbl.Cells(k2, k1).Value
                                                        If IsNothing(NameTypeObject) = False Then
                                                            If NameTypeObject.Trim.Length > 0 Then
                                                                If IsArray(ArrayTypeObject) = False Then
                                                                    ReDim Preserve ArrayTypeObject(1, CountArrayTypeObject)
                                                                    ArrayTypeObject(0, CountArrayTypeObject) = NameTypeObject
                                                                    ArrayTypeObject(1, CountArrayTypeObject) = tbl.Handle.ToString
                                                                    CountArrayTypeObject += 1
                                                                Else
                                                                    Dim boolFind As Boolean = False
                                                                    For k3 As Integer = 0 To ArrayTypeObject.GetUpperBound(1)
                                                                        Dim TempTypeOblect As String = ArrayTypeObject(0, k3)
                                                                        If TempTypeOblect Like NameTypeObject Then
                                                                            boolFind = True
                                                                            Exit For
                                                                        End If
                                                                    Next k3
                                                                    If boolFind = False Then
                                                                        ReDim Preserve ArrayTypeObject(1, CountArrayTypeObject)
                                                                        ArrayTypeObject(0, CountArrayTypeObject) = NameTypeObject
                                                                        ArrayTypeObject(1, CountArrayTypeObject) = tbl.Handle.ToString
                                                                        CountArrayTypeObject += 1
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Next k2
                                                End If
                                            End If
                                        Next k1
                                    Next k
                                End If
                                '=======================================================================================================
                                'читаем строки таблицы
                                If IsArray(ArrayTypeObject) = True Then
                                    For k As Integer = 0 To ArrayTypeObject.GetUpperBound(1)
                                        Dim ArrayhgTables As String() = ArrayTypeObject(1, k).Split()
                                        If IsArray(ArrayhgTables) = True Then
                                            For k1 As Integer = 0 To ArrayhgTables.Length - 1
                                                xw.WriteStartElement("TypeObject")
                                                xw.WriteAttributeString("Type", ArrayTypeObject(0, k).Trim)
                                                xw.WriteEndElement()
                                                Dim StrHandle As String = ArrayhgTables(k1)
                                                If IsNothing(StrHandle) = True Then
                                                    Continue For
                                                End If
                                                If StrHandle.Trim.Length = 0 Then
                                                    Continue For
                                                End If
                                                Dim TableID As ObjectId = Nothing
                                                Try
                                                    Dim ln As Long = Convert.ToInt64(StrHandle, 16)
                                                    Dim hn As Handle = New Handle(ln)
                                                    TableID = sourceDb.GetObjectId(False, hn, 0)
                                                Catch ex As ArgumentOutOfRangeException
                                                    Continue For
                                                End Try
                                                If IsNothing(TableID) = True Then
                                                    Continue For
                                                End If
                                                If TableID.IsNull = True Then
                                                    Continue For
                                                End If
                                                Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = TableID.GetObject(OpenMode.ForRead)

                                                For k2 As Integer = 3 To tbl.Rows.Count - 1
                                                    Dim ValStr As String = tbl.Cells(k2, 1).TextString
                                                    If IsNothing(ValStr) = False Then
                                                        If ValStr.Trim.Length > 0 Then
                                                            ValStr = ValStr.Trim()
                                                            xw.WriteStartElement("Rows")
                                                            ReDim Preserve arrayUserObject(3, countArrayUserObject)
                                                            arrayUserObject(0, countArrayUserObject) = tableCategory
                                                            For k3 As Integer = 0 To tbl.Columns.Count - 1
                                                                ValStr = tbl.Cells(k2, k3).TextString
                                                                If IsNothing(ValStr) = True Then
                                                                    ValStr = "0"
                                                                End If
                                                                Dim NameField As String = tbl.Cells(2, k3).TextString
                                                                If IsNothing(NameField) = False Then
                                                                    If Not (NameField.Trim Like "Type") Then
                                                                        xw.WriteAttributeString(NameField.Trim, ValStr)
                                                                    Else
                                                                        arrayUserObject(1, countArrayUserObject) = ValStr
                                                                    End If
                                                                    If NameField.Trim Like "Name" Then
                                                                        arrayUserObject(2, countArrayUserObject) = ValStr
                                                                    End If
                                                                    If NameField.Trim Like "PS" Then
                                                                        arrayUserObject(3, countArrayUserObject) = ValStr
                                                                    End If
                                                                End If
                                                            Next k3
                                                            countArrayUserObject += 1
                                                            xw.WriteEndElement()
                                                        End If
                                                    End If
                                                Next k2
                                            Next k1
                                        End If
                                    Next k
                                    xw.WriteEndElement()
                                End If
                            End If
                        Next j
                        xw.WriteEndElement()
                    End If
                    xw.WriteEndElement()
                Next i
                xw.WriteEndElement()
                '===========================================================================================================================
                'информация по слоям
                '===========================================================================================================================
                Dim lfc As LayerFilterCollection = sourceDb.LayerFilters.Root.NestedFilters
                xw.WriteStartElement("Layers")
                xw.WriteAttributeString("NameLayers", "Слои")
                Dim acLyrTbl As LayerTable = sourceTr2.GetObject(sourceDb.LayerTableId, OpenMode.ForRead)
                If lfc.Count > 0 Then
                    For k1 As Integer = 0 To lfc.Count - 1
                        Dim lfilter As LayerFilter = lfc.Item(k1)
                        Dim nameFilter As String = lfilter.Name
                        Dim fExp As String = lfilter.FilterExpression
                        If nameFilter Like "Все используемые слои" Or nameFilter Like "Несогласованные новые слои" Or nameFilter Like "Переопределения ВЭ" Then
                            Continue For
                        End If
                        xw.WriteStartElement("LayerGroup")
                        xw.WriteAttributeString("NameGroup", nameFilter)
                        xw.WriteAttributeString("FilterExpression", fExp)
                        For Each id As ObjectId In acLyrTbl
                            Dim acLyrTblRec As LayerTableRecord = sourceTr2.GetObject(id, OpenMode.ForRead)
                            If lfilter.Filter(acLyrTblRec) = True Then
                                xw.WriteStartElement("Layer")
                                xw.WriteAttributeString("Name", acLyrTblRec.Name)
                                xw.WriteAttributeString("Color", acLyrTblRec.Color.ColorIndex)
                                Dim idLinetype As ObjectId = acLyrTblRec.LinetypeObjectId
                                Dim recTypeLine As LinetypeTableRecord = sourceTr2.GetObject(idLinetype, OpenMode.ForRead)
                                xw.WriteAttributeString("LineType", recTypeLine.Name)

                                Dim strWight As String = acLyrTblRec.LineWeight.ToString
                                strWight = strWight.Replace("LineWeight", "")
                                If Val(strWight) > 0 Then
                                    xw.WriteAttributeString("LineWidth", Val(strWight))
                                Else
                                    xw.WriteAttributeString("LineWidth", acLyrTblRec.LineWeight.ToString)
                                End If
                                xw.WriteEndElement()
                            End If
                        Next
                        xw.WriteEndElement()
                    Next k1
                End If
                xw.WriteEndElement()
                '=====================================================================================================================================================
                'типы линий
                '=====================================================================================================================================================
                'делаем сортировку по категорииям
                For i As Integer = 0 To arrayUserObject.GetUpperBound(1)
                    Dim categoryName As String = arrayUserObject(0, i)
                    Dim typeObject As String = arrayUserObject(1, i)
                    Dim nameObject As String = arrayUserObject(2, i)
                    Dim tablePS As String = arrayUserObject(3, i)

                    If dictionaryPS.ContainsKey(categoryName) = True Then
                        Dim arraystr As String() = dictionaryPS.Item(categoryName)
                        If IsArray(arraystr) = True Then
                            Dim ind As Integer = Array.IndexOf(arraystr, tablePS)
                            If ind < 0 Then
                                Dim lenArray As Integer = arraystr.Length
                                ReDim Preserve arraystr(lenArray)
                                arraystr(lenArray) = tablePS
                                dictionaryPS.Item(categoryName) = arraystr
                            End If
                        End If
                    Else
                        Dim arrayTemp As String() = {tablePS}
                        dictionaryPS.Add(categoryName, arrayTemp)
                    End If

                    If typeObject Like "Блок" Then
                        Continue For
                    ElseIf typeObject Like "Мультилиния" Then
                        Continue For
                    ElseIf typeObject Like "ТочкаCOGO" Then
                        Continue For
                    ElseIf typeObject Like "Метка" Then
                        Continue For
                    ElseIf typeObject Like "Характерная линия" Then
                        Continue For
                    ElseIf typeObject Like "Таблица" Then
                        Continue For
                    ElseIf typeObject Like "Размер" Then
                        Continue For
                    ElseIf typeObject Like "Мультивыноска" Then
                        If dictionaryMLeaderStyle.ContainsKey(categoryName) = True Then
                            Dim arraystr As String() = dictionaryMLeaderStyle.Item(categoryName)
                            If IsArray(arraystr) = True Then
                                Dim ind As Integer = Array.IndexOf(arraystr, nameObject)
                                If ind < 0 Then
                                    Dim lenArray As Integer = arraystr.Length
                                    ReDim Preserve arraystr(lenArray)
                                    arraystr(lenArray) = nameObject
                                    dictionaryMLeaderStyle.Item(categoryName) = arraystr
                                End If
                            End If
                        Else
                            Dim arrayTemp As String() = {nameObject}
                            dictionaryMLeaderStyle.Add(categoryName, arrayTemp)
                        End If
                    Else
                        Dim boolLineType As Boolean = acLineTypTbl.Has(nameObject)
                        If boolLineType = True Then
                            If dictionaryTypeLine.ContainsKey(categoryName) = True Then
                                Dim arraystr As String() = dictionaryTypeLine.Item(categoryName)
                                If IsArray(arraystr) = True Then
                                    Dim ind As Integer = Array.IndexOf(arraystr, nameObject)
                                    If ind < 0 Then
                                        Dim lenArray As Integer = arraystr.Length
                                        ReDim Preserve arraystr(lenArray)
                                        arraystr(lenArray) = nameObject
                                        dictionaryTypeLine.Item(categoryName) = arraystr
                                    End If
                                End If
                            Else
                                Dim arrayTemp As String() = {nameObject}
                                dictionaryTypeLine.Add(categoryName, arrayTemp)
                            End If
                        End If
                    End If
                Next i
                xw.WriteStartElement("LinetypeRecords")
                xw.WriteAttributeString("NameLineTypeRecords", "Типы линий")
                If dictionaryTypeLine.Count > 0 Then
                    For i As Integer = 0 To dictionaryTypeLine.Count - 1
                        Dim userKeyPair As KeyValuePair(Of String, String()) = dictionaryTypeLine.ElementAt(i)
                        Dim nameGroupTypeLine As String = userKeyPair.Key
                        Dim arrayDictTypeLine As String() = userKeyPair.Value
                        If IsArray(arrayDictTypeLine) = True Then
                            xw.WriteStartElement("LineTypeGroup")
                            xw.WriteAttributeString("NameGroupLinetypeRecords", nameGroupTypeLine)
                            For j As Integer = 0 To arrayDictTypeLine.Length - 1
                                Dim nameTypeRecords As String = arrayDictTypeLine(j)
                                If acLineTypTbl.Has(nameTypeRecords) = True Then
                                    Dim idLineTypeRecords As ObjectId = acLineTypTbl.Item(nameTypeRecords)
                                    Dim recTypeLine As LinetypeTableRecord = sourceTr2.GetObject(idLineTypeRecords, OpenMode.ForRead)
                                    Dim nameTypeLine As String = recTypeLine.Name
                                    Dim deskTypeLine As String = recTypeLine.AsciiDescription
                                    Dim CodeTypeLine As String = "A,"
                                    Dim lenD As Integer = recTypeLine.NumDashes
                                    If lenD > 0 Then
                                        Try
                                            For k As Integer = 0 To lenD - 1
                                                Dim l As Double = 0
                                                Dim indexShape As Integer = 0
                                                Dim textVal As String = ""
                                                Dim nameShape As String = ""
                                                'находим значения
                                                l = recTypeLine.DashLengthAt(k) 'длина штриха
                                                indexShape = recTypeLine.ShapeNumberAt(k) 'индекс формы (0 - нет формы)
                                                Try
                                                    textVal = recTypeLine.TextAt(k) 'читаем текст
                                                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                                End Try
                                                If indexShape > 0 Then
                                                    Dim idTextStyleShape As ObjectId = recTypeLine.ShapeStyleAt(k)
                                                    Dim textStyleShape As TextStyleTableRecord = sourceTr2.GetObject(idTextStyleShape, OpenMode.ForRead)
                                                    If textStyleShape.IsShapeFile = True Then
                                                        Using acTrans As Transaction = sourceDb.TransactionManager.StartTransaction()
                                                            Dim Btr As BlockTableRecord = acTrans.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForWrite, False)
                                                            Dim pShape As Autodesk.AutoCAD.DatabaseServices.Shape = New Autodesk.AutoCAD.DatabaseServices.Shape(Point3d.Origin, 1, 0, 1)
                                                            pShape.ShapeNumber = indexShape
                                                            pShape.StyleId = idTextStyleShape
                                                            Btr.AppendEntity(pShape)
                                                            nameShape = pShape.Name
                                                        End Using
                                                        'имя формы не найдено, ищем имя формы в файле shp
                                                        If nameShape.Trim.Length = 0 Then
                                                            If IsArray(supportPatchFolder) = False Then
                                                                Dim boolFindFolderPatch As Boolean = FuncFiles.FuncGetSupportFolderPath(supportPatchFolder)
                                                            End If
                                                            If IsArray(supportPatchFolder) = True Then
                                                                Dim nameFileSHP As String = textStyleShape.FileName.Replace(".shx", ".shp")
                                                                Dim boolFindShp As Boolean = False
                                                                Dim regMatch As String = "\*"
                                                                For kd As Integer = 0 To supportPatchFolder.Length - 1
                                                                    Dim searFiles As String() = Directory.GetFiles(supportPatchFolder(kd), nameFileSHP)
                                                                    If IsArray(searFiles) = True Then
                                                                        If searFiles.Length > 0 Then
                                                                            'читаем файл с формой
                                                                            Dim ffileShp As String = searFiles(0)
                                                                            If IO.File.Exists(ffileShp) = True Then
                                                                                Dim sr As StreamReader = New StreamReader(ffileShp)
                                                                                While sr.Peek <> -1
                                                                                    Dim st As String = sr.ReadLine().Trim
                                                                                    Dim flagread As String = Left(st, 1)
                                                                                    If Regex.IsMatch(flagread, regMatch) Then 'chr(42)
                                                                                        Dim rez As String() = st.Split(",")
                                                                                        If rez.GetUpperBound(0) > 1 Then
                                                                                            Dim indForm As String = Mid(rez(0), 2)
                                                                                            If indForm Like indexShape Then
                                                                                                nameShape = rez(2)
                                                                                                boolFindShp = True
                                                                                                Exit While
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End While
                                                                                sr.Close()
                                                                                If boolFindShp = True Then
                                                                                    Exit For
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next kd
                                                            End If
                                                        End If
                                                    End If
                                                    CodeTypeLine = CodeTypeLine & l & ",[" & nameShape & "," & textStyleShape.FileName & ",S=" & Math.Round(recTypeLine.ShapeScaleAt(k), 4) & ",R=" & Math.Round(recTypeLine.ShapeRotationAt(k), 6) & ",X=" & Math.Round(recTypeLine.ShapeOffsetAt(k).X, 4) & ",Y=" & Math.Round(recTypeLine.ShapeOffsetAt(k).Y, 4) & "]"
                                                ElseIf textVal.Trim.Length > 0 Then
                                                    Dim id As ObjectId = recTypeLine.ShapeStyleAt(k)
                                                    Dim acShape As TextStyleTableRecord = sourceTr2.GetObject(id, OpenMode.ForRead)
                                                    Dim styleName As String = acShape.Name
                                                    If styleName Like "STANDARD" Or styleName Like "Standard" Then
                                                        If nameGroupTypeLine Like "Проект планировки и межевания территории" Or nameGroupTypeLine Like "Земельный кадастр" Then
                                                            styleName = "CADASTRAL_OBJECT"
                                                        ElseIf nameGroupTypeLine Like "Топография" Then
                                                            styleName = "Bm-431"
                                                        End If
                                                    End If
                                                    CodeTypeLine = CodeTypeLine & l & ",[" & Chr(34) & textVal & Chr(34) & "," & styleName & ",S=" & Math.Round(recTypeLine.ShapeScaleAt(k), 4) & ",R=" & Math.Round(recTypeLine.ShapeRotationAt(k), 6) & ",X=" & Math.Round(recTypeLine.ShapeOffsetAt(k).X, 4) & ",Y=" & Math.Round(recTypeLine.ShapeOffsetAt(k).Y, 4) & "]"
                                                Else
                                                    CodeTypeLine = CodeTypeLine & l
                                                End If
                                                If k < lenD - 1 Then
                                                    CodeTypeLine = CodeTypeLine & ","
                                                End If
                                            Next k
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        End Try
                                    End If
                                    xw.WriteStartElement("LinetypeRecord")
                                    'заносим первыми три типа линии ПоСлою, ПоБлоку и Непрерывный
                                    xw.WriteAttributeString("NameLineType", nameTypeLine)
                                    xw.WriteAttributeString("DeskLineType", deskTypeLine)
                                    xw.WriteAttributeString("CodeLineType", CodeTypeLine)
                                    xw.WriteEndElement()
                                End If
                            Next j
                            xw.WriteEndElement()
                        End If
                    Next i
                End If
                xw.WriteEndElement()
                '=======================================================================================================================================================
                'стили текста
                '=======================================================================================================================================================
                Dim StText As TextStyleTable = sourceTr2.GetObject(sourceDb.TextStyleTableId, OpenMode.ForRead)
                xw.WriteStartElement("TextStyles")
                xw.WriteAttributeString("NameTextStyles", "Стили текстов")
                For Each idTextStyle In StText
                    Dim userTextStyle As TextStyleTableRecord = sourceTr2.GetObject(idTextStyle, OpenMode.ForRead)
                    xw.WriteStartElement("TextStyle")
                    xw.WriteAttributeString("NameTextStyle", userTextStyle.Name)
                    xw.WriteAttributeString("NameFont", userTextStyle.FileName)
                    xw.WriteAttributeString("Italic", userTextStyle.Font.Italic)
                    xw.WriteAttributeString("Annotative", userTextStyle.Annotative.ToString)
                    xw.WriteAttributeString("TextStyleHeight", userTextStyle.TextSize.ToString)
                    xw.WriteAttributeString("TextStyleRotate", userTextStyle.ObliquingAngle.ToString)
                    xw.WriteAttributeString("TextStylePrior", userTextStyle.PriorSize.ToString)
                    xw.WriteAttributeString("TextStyleSHP", userTextStyle.IsShapeFile.ToString)
                    xw.WriteEndElement()
                Next
                xw.WriteEndElement()
                '=======================================================================================================================================================
                'стили мультивыносок
                '=======================================================================================================================================================
                Dim acMleaderStyleTbl As DBDictionary = sourceTr2.GetObject(sourceDb.MLeaderStyleDictionaryId, OpenMode.ForRead)
                xw.WriteStartElement("MLeaderStyles")
                xw.WriteAttributeString("NameMLeaderStyles", "Стили мультивыносок")
                If dictionaryMLeaderStyle.Count > 0 Then
                    For i As Integer = 0 To dictionaryMLeaderStyle.Count - 1
                        Dim userKeyPair As KeyValuePair(Of String, String()) = dictionaryMLeaderStyle.ElementAt(i)
                        Dim nameGroupTypeMleader As String = userKeyPair.Key
                        Dim arrayDictMLeader As String() = userKeyPair.Value
                        If IsArray(arrayDictMLeader) = True Then
                            xw.WriteStartElement("MLeaderStylesGroup")
                            xw.WriteAttributeString("NameMLeaderStylesGroup", nameGroupTypeMleader)
                            For j As Integer = 0 To arrayDictMLeader.Length - 1
                                Dim nameTypeRecords As String = arrayDictMLeader(j)
                                For Each mleader As DBDictionaryEntry In acMleaderStyleTbl
                                    Dim id As ObjectId = mleader.Value
                                    Dim userMLeader As MLeaderStyle = sourceTr2.GetObject(id, OpenMode.ForRead)
                                    If userMLeader.Name Like nameTypeRecords Then
                                        Dim drw As DrawMLeaderOrderType = userMLeader.DrawMLeaderOrderType
                                        xw.WriteStartElement("MLeaderStyle")
                                        xw.WriteAttributeString("NameMLeaderStyle", userMLeader.Name)
                                        xw.WriteAttributeString("ArrowSize", userMLeader.ArrowSize)
                                        Dim arrowSumbId As ObjectId = userMLeader.ArrowSymbolId
                                        If arrowSumbId.IsNull = False Then
                                            Dim userblk As BlockTableRecord = sourceTr2.GetObject(arrowSumbId, OpenMode.ForRead)
                                            xw.WriteAttributeString("NameArrow", userblk.Name)
                                        Else
                                            Dim idblk As ObjectId = userMLeader.BlockId
                                            If idblk.IsNull = False Then
                                                Dim userblk As BlockTableRecord = sourceTr2.GetObject(idblk, OpenMode.ForRead)
                                                xw.WriteAttributeString("NameArrow", userblk.Name)
                                            End If
                                        End If
                                        xw.WriteEndElement()
                                        Exit For
                                    End If
                                Next
                            Next j
                            xw.WriteEndElement()
                        End If
                    Next i
                End If
                xw.WriteEndElement()
                '=====================================================================================================================================================
                'таблицы PS
                '=====================================================================================================================================================
                xw.WriteStartElement("PropertyTables")
                xw.WriteAttributeString("NamePropertyTables", "Таблицы кодификатора")
                Dim dictPropSetDef = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(sourceDb)
                Dim strNameDictCollection As Specialized.StringCollection = dictPropSetDef.NamesInUse
                If strNameDictCollection.Count > 0 Then
                    For i As Integer = 0 To strNameDictCollection.Count - 1
                        Dim nameTablePS As String = strNameDictCollection.Item(i)
                        If dictPropSetDef.Has(nameTablePS, sourceTr2) Then
                            Dim idPsTable As ObjectId = dictPropSetDef.GetAt(nameTablePS)
                            If IsNothing(idPsTable) = False Then
                                If idPsTable.IsNull = False Then
                                    xw.WriteStartElement("PropertyTable")
                                    Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDefinition = sourceTr2.GetObject(idPsTable, OpenMode.ForRead, False, False)
                                    Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertyDefinitionCollection = pset.Definitions
                                    xw.WriteAttributeString("TableName", nameTablePS)
                                    xw.WriteAttributeString("TableDescription", pset.Description)
                                    Dim arrayFields As String(,) = Nothing
                                    Dim countArrayFields As Integer = 0
                                    For Each field As Autodesk.AEC.PropertyData.DatabaseServices.PropertyDefinition In psetDataColl
                                        Dim TempNameField As String = field.Name
                                        Dim data As String = field.DefaultData
                                        Dim TypeField As Autodesk.AEC.PropertyData.DataType = field.DataType
                                        ReDim Preserve arrayFields(6, countArrayFields)
                                        arrayFields(0, countArrayFields) = field.DisplayOrder
                                        arrayFields(1, countArrayFields) = field.Name
                                        arrayFields(2, countArrayFields) = field.Description
                                        arrayFields(3, countArrayFields) = ""
                                        arrayFields(4, countArrayFields) = TypeField
                                        arrayFields(5, countArrayFields) = field.DefaultData
                                        arrayFields(6, countArrayFields) = field.IsVisible
                                        countArrayFields += 1
                                    Next
                                    If IsArray(arrayFields) = True Then
                                        If arrayFields.GetUpperBound(1) > 0 Then
                                            Dim boolSort As Boolean = MathFunction.FuncSortStrArray(arrayFields, 0)
                                        End If
                                        For j1 As Integer = 0 To arrayFields.GetUpperBound(1)
                                            xw.WriteStartElement("Field")
                                            xw.WriteAttributeString("Name", arrayFields(1, j1))
                                            xw.WriteAttributeString("Description", arrayFields(2, j1))
                                            xw.WriteAttributeString("UnitType", arrayFields(3, j1))
                                            xw.WriteAttributeString("TypeField", arrayFields(4, j1))
                                            xw.WriteAttributeString("DefaultData", arrayFields(5, j1))
                                            xw.WriteAttributeString("DisplayOrder", arrayFields(0, j1))
                                            xw.WriteAttributeString("Visible", arrayFields(6, j1))
                                            xw.WriteEndElement()
                                        Next j1
                                    End If
                                    xw.WriteEndElement()
                                End If
                            End If
                        End If
                    Next i
                    xw.WriteEndElement()
                End If
                '=======================================================================================================================================
                'блоки
                '=======================================================================================================================================
                xw.WriteStartElement("BlocksDefinition")
                Dim BlockTable As BlockTable = sourceTr2.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
                Dim attDefClass = RXObject.GetClass(GetType(AttributeDefinition))
                For Each blockId As ObjectId In BlockTable
                    Dim BlockTableRecord As BlockTableRecord = sourceTr2.GetObject(blockId, OpenMode.ForRead)
                    xw.WriteStartElement("BlockReference")
                    xw.WriteAttributeString("NameBlockReference", "Определение блока")
                    Dim nameBlk As String = BlockTableRecord.Name
                    Dim desk As String = BlockTableRecord.Comments
                    xw.WriteAttributeString("NameBlock", nameBlk)
                    xw.WriteAttributeString("Description", desk)
                    If BlockTableRecord.HasAttributeDefinitions = True Then
                        For Each id As ObjectId In BlockTableRecord
                            If id.ObjectClass = attDefClass Then
                                Dim attDef As AttributeDefinition = CType(sourceTr2.GetObject(id, OpenMode.ForRead), AttributeDefinition)
                                Dim nameDef As String = attDef.Tag
                                Dim promptDef As String = attDef.Prompt
                                Dim valDef As String = attDef.TextString
                                Dim height As Double = attDef.Height
                                Dim visibleAttr As Boolean = attDef.Invisible
                                Dim x As Double = Math.Round(attDef.Position.X, 3)
                                Dim y As Double = Math.Round(attDef.Position.Y, 3)
                                Dim textStyle As String = attDef.TextStyleName
                                If attDef.HasFields = True Then
                                    Dim idField As ObjectId = attDef.GetField() 'смотрим старый атрибут
                                    Dim fd As Field = sourceTr2.GetObject(idField, OpenMode.ForRead) 'читаем его
                                    valDef = fd.GetFieldCode 'получаем код поля
                                End If
                                xw.WriteStartElement("AttributeDefinition")
                                xw.WriteAttributeString("TagAttribute", nameDef)
                                xw.WriteAttributeString("DefinitionAttribute", promptDef)
                                xw.WriteAttributeString("TextStringAttribute", valDef)
                                xw.WriteAttributeString("RotationAttribute", attDef.Rotation)
                                xw.WriteAttributeString("HeightAttribute", height)
                                xw.WriteAttributeString("TextStyleAttribute", textStyle)
                                xw.WriteAttributeString("XAttribute", x)
                                xw.WriteAttributeString("YAttribute", y)
                                xw.WriteAttributeString("InvisibleAttribute", visibleAttr)
                                xw.WriteAttributeString("ColorAttribute", attDef.ColorIndex)
                                xw.WriteAttributeString("TextStyle", attDef.TextStyleName)
                                xw.WriteAttributeString("HorizontalAlignment", attDef.HorizontalMode.ToString)
                                xw.WriteAttributeString("VerticalAlignment", attDef.VerticalMode.ToString)
                                Dim v As Integer = attDef.VerticalMode
                                xw.WriteEndElement()
                            End If
                        Next
                    End If
                    xw.WriteEndElement()
                Next
                xw.WriteEndElement()
                xw.WriteEndDocument()
                xw.Flush()
                xw.Close()
                xw.Dispose()
                xw = Nothing
            End Using
        End If
        sourceDb.CloseInput(True)
        FuncCreateXMLAllListObject = True
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает все категории условных знаков
    Public Shared Function FuncReadCategoryToXML(ByVal putchFileXml As String, ByRef arrayCategory As String()) As Boolean
        FuncReadCategoryToXML = False
        Dim CountArrayCategory As Integer = 0
        If IsArray(arrayCategory) = True Then
            CountArrayCategory = arrayCategory.Length
        End If
        'читаем xml и загружаем ветки с категориями
        Try
            Dim xDoc As XmlDocument = New XmlDocument()
            If IO.File.Exists(putchFileXml) = True Then
                Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
                While reader.Read()
                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "Category" Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        If nameN Like "CategoryName" Then
                                            Dim valName As String = reader.Value
                                            If IsNothing(valName) = False Then
                                                If valName.Trim.Length > 0 Then
                                                    ReDim Preserve arrayCategory(CountArrayCategory)
                                                    arrayCategory(CountArrayCategory) = valName
                                                    CountArrayCategory += 1
                                                    Exit While
                                                End If
                                            End If
                                        End If
                                    End While
                                End If
                            End If
                        Case XmlNodeType.EndElement
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "Category" Then
                                Exit Function
                            End If
                    End Select
                End While
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает все таблицы определенной категории
    Public Shared Function FuncReadTablesToXML(ByVal putchFileXml As String, ByVal category As String, ByRef arrayTables As String()) As Boolean
        FuncReadTablesToXML = False
        Dim CountArrayTables As Integer = 0
        If IsArray(arrayTables) = True Then
            CountArrayTables = arrayTables.Length
        End If
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolCategory As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If category Like reader.Value Then
                                        boolCategory = True
                                    Else
                                        boolCategory = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Table" Then
                            If boolCategory = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        If nameN.Trim Like "TableName" Then
                                            Dim valName As String = reader.Value
                                            If IsNothing(valName) = False Then
                                                If valName.Trim.Length > 0 Then
                                                    ReDim Preserve arrayTables(CountArrayTables)
                                                    arrayTables(CountArrayTables) = valName
                                                    CountArrayTables += 1
                                                    Exit While
                                                End If
                                            End If
                                        End If
                                    End While
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LibrarySymbols" Then
                            Exit Function
                        ElseIf NameBlock Like "Tables" Then
                            If IsArray(arrayTables) = True Then
                                Exit Function
                            End If
                        End If
                End Select
            End While
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает объекты определенного типа из выбранной таблицы
    Public Shared Function FuncReadTypeObjectToXML(ByVal putchFileXml As String, ByVal category As String, ByVal table As String, ByVal typeObject As String, ByRef arrayObject As String()) As Boolean
        Dim count As Integer = 0
        FuncReadTypeObjectToXML = False
        Dim CountArrayObject As Integer = 0
        If IsArray(arrayObject) = True Then
            CountArrayObject = arrayObject.Length
        End If
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolCategory As Boolean = False
            Dim boolTable As Boolean = False
            Dim boolType As Boolean = False
            Using reader As XmlReader = XmlReader.Create(putchFileXml,
    New XmlReaderSettings() With {
        .DtdProcessing = DtdProcessing.Ignore,
        .ValidationType = ValidationType.None
    })
                While reader.Read()
                    'Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            count += 1
                            If NameBlock Like "Category" Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        If category Like reader.Value Then
                                            boolCategory = True
                                        Else
                                            boolCategory = False
                                        End If
                                    End While
                                End If
                            ElseIf NameBlock Like "Table" And boolCategory = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        If table Like reader.Value Then
                                            boolTable = True
                                        Else
                                            boolTable = False
                                        End If
                                    End While
                                End If
                            ElseIf NameBlock Like "TypeObject" And boolTable = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        If typeObject Like reader.Value Then
                                            boolType = True
                                        Else
                                            boolType = False
                                        End If
                                    End While
                                End If
                            ElseIf NameBlock Like "Rows" And boolType = True Then
                                Dim dataProperties As String(,) = Nothing
                                Dim countDataProperties As Integer = 0
                                Dim boolDataFind As Boolean = False
                                If boolTable = True And boolCategory = True And boolType = True Then
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim nameN As String = reader.Name
                                            If nameN Like "Description" Then
                                                Dim valName As String = reader.Value
                                                If IsNothing(valName) = False Then
                                                    If valName.Trim.Length > 0 Then
                                                        ReDim Preserve arrayObject(CountArrayObject)
                                                        arrayObject(CountArrayObject) = valName
                                                        CountArrayObject += 1
                                                        Exit While
                                                    End If
                                                End If
                                            End If
                                        End While
                                    End If
                                End If
                            End If
                        Case XmlNodeType.EndElement
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "LibrarySymbols" Then
                                Exit Function
                            ElseIf NameBlock Like "Tables" Then
                                If IsArray(arrayObject) = True Then
                                    If arrayObject.Length > 0 Then
                                        Exit Function
                                    End If
                                End If
                            End If
                    End Select
                End While
            End Using
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает свойства объекта из файла xml (Имя категории, Имя таблицы, Тип объекта, Описание объекта, результат - массив с Характеристиками)
    Public Shared Function FuncReadPropertiesObjectToXML(ByVal putchFileXml As String, ByVal category As String, ByVal table As String, ByVal typeObject As String, ByVal deskObject As String, ByRef dataObjectProperties As String(,)) As Boolean
        FuncReadPropertiesObjectToXML = False
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolCategory As Boolean = False
            Dim boolTable As Boolean = False
            Dim boolType As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If category Like reader.Value Then
                                        boolCategory = True
                                    Else
                                        boolCategory = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Table" And boolCategory = True Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If table Like reader.Value Then
                                        boolTable = True
                                    Else
                                        boolTable = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "TypeObject" And boolTable = True Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If typeObject Like reader.Value Then
                                        boolType = True
                                    Else
                                        boolType = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Rows" And boolType = True Then
                            Dim dataProperties As String(,) = Nothing
                            Dim countDataProperties As Integer = 0
                            Dim boolDataFind As Boolean = False
                            If boolTable = True And boolCategory = True And boolType = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        ReDim Preserve dataProperties(1, countDataProperties)
                                        dataProperties(0, countDataProperties) = nameN
                                        dataProperties(1, countDataProperties) = NameObject
                                        countDataProperties += 1
                                        If nameN Like "Description" Then
                                            If Not (NameObject Like deskObject) Then
                                                boolDataFind = False
                                                Exit While
                                            Else
                                                boolDataFind = True
                                            End If
                                        End If
                                    End While
                                    If boolDataFind = True Then
                                        dataObjectProperties = dataProperties
                                        FuncReadPropertiesObjectToXML = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LibrarySymbols" Then
                            Exit Function
                        End If
                End Select
            End While
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает свойства объекта из файла xml (Описание объекта, результат - массив с Характеристиками)
    Public Shared Function FuncReadPropertiesAllObjectToXML(ByVal putchFileXml As String, ByVal deskObject As String, ByRef dataObjectProperties As String(,), Optional category As String = "", Optional table As String = "", Optional typeObject As String = "") As Boolean
        FuncReadPropertiesAllObjectToXML = False
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolCategory As Boolean = False
            Dim boolTable As Boolean = False
            Dim boolType As Boolean = False
            If category.Trim.Length = 0 Then
                boolCategory = True
            End If
            If table.Trim.Length = 0 Then
                boolTable = True
            End If
            If typeObject.Trim.Length = 0 Then
                boolType = True
            End If

            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If category.Trim.Length = 0 Then
                                Continue While
                            End If
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If category Like reader.Value Then
                                        boolCategory = True
                                    Else
                                        boolCategory = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Table" Then
                            If table.Trim.Length = 0 Then
                                Continue While
                            End If
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If table Like reader.Value Then
                                        boolTable = True
                                    Else
                                        boolTable = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "TypeObject" Then
                            If typeObject.Trim.Length = 0 Then
                                Continue While
                            End If
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If typeObject Like reader.Value Then
                                        boolType = True
                                    Else
                                        boolType = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Rows" And boolType = True Then
                            Dim dataProperties As String(,) = Nothing
                            Dim countDataProperties As Integer = 0
                            Dim boolDataFind As Boolean = False
                            If boolTable = True And boolCategory = True And boolType = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        ReDim Preserve dataProperties(1, countDataProperties)
                                        dataProperties(0, countDataProperties) = nameN
                                        dataProperties(1, countDataProperties) = NameObject
                                        countDataProperties += 1
                                        If nameN Like "Description" Then
                                            If Not (NameObject Like deskObject) Then
                                                boolDataFind = False
                                                Exit While
                                            Else
                                                boolDataFind = True
                                            End If
                                        End If
                                    End While
                                    If boolDataFind = True Then
                                        dataObjectProperties = dataProperties
                                        FuncReadPropertiesAllObjectToXML = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LibrarySymbols" Then
                            Exit Function
                        ElseIf NameBlock Like "Tables" Then
                            If IsArray(dataObjectProperties) = True Then
                                Exit Function
                            End If
                        End If
                End Select
            End While
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает информацию о типе линии из xml и создает тип линии
    Public Shared Function FuncCreateLineTypeToXML(ByVal putchFileXml As String, ByVal nameLineTypeRecords As String) As LinetypeTableRecord
        FuncCreateLineTypeToXML = Nothing
        Dim deskTypeLine As String = ""
        Dim patternTypeLine As String = ""
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolFindLineType As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LinetypeRecord" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If nameN Like "NameLineType" Then
                                        If nameLineTypeRecords Like reader.Value Then
                                            boolFindLineType = True
                                        Else
                                            boolFindLineType = False
                                            Exit While
                                        End If
                                    ElseIf nameN Like "DeskLineType" Then
                                        deskTypeLine = reader.Value
                                    ElseIf nameN Like "CodeLineType" Then
                                        patternTypeLine = reader.Value
                                    End If
                                End While
                                If boolFindLineType = True Then
                                    Dim recTypeLine As LinetypeTableRecord = AcadFunction.FuncCreateTypeLineByCode(nameLineTypeRecords, deskTypeLine, patternTypeLine)
                                    Return recTypeLine
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LinetypeRecords" Then
                            Exit Function
                        End If
                End Select
            End While
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает информацию о характеристиках слоя из xml и создает слой 
    Public Shared Function FuncCreateLayerToXML(ByVal putchFileXml As String, ByVal nameLayer As String) As ObjectId
        FuncCreateLayerToXML = Nothing
        Dim indColor As Integer = 7
        Dim nameLineType As String = "Continuous"
        Dim idLineTypeRec As ObjectId = Nothing
        Dim widthLine As Double = 20
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolFindLayer As Boolean = False
            Dim boolFindLayers As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Layer" And boolFindLayers = True Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If nameN Like "Name" Then
                                        If nameLayer Like reader.Value Then
                                            boolFindLayer = True
                                        Else
                                            boolFindLayer = False
                                            Exit While
                                        End If
                                    ElseIf nameN Like "Color" Then
                                        indColor = Val(reader.Value)
                                        If indColor < 1 Then
                                            indColor = 7
                                        ElseIf indColor > 255 Then
                                            indColor = 7
                                        End If
                                    ElseIf nameN Like "LineType" Then
                                        nameLineType = reader.Value
                                    ElseIf nameN Like "LineWidth" Then
                                        Dim tempWidthLine As String = reader.Value
                                        If tempWidthLine Like "ПоСлою" Then
                                            widthLine = -1
                                        ElseIf tempWidthLine Like "ПоБлоку" Then
                                            widthLine = -2
                                        ElseIf tempWidthLine Like "ByLineWeightDefault" Then
                                            widthLine = -3
                                        Else
                                            widthLine = Val(tempWidthLine)
                                        End If
                                    End If
                                End While
                                If boolFindLayer = True Then
                                    idLineTypeRec = AcadFunction.FuncGetLineTypeId(nameLineType)
                                    If IsNothing(idLineTypeRec) = True OrElse idLineTypeRec.IsNull = True Then
                                        Dim recTypeLine As LinetypeTableRecord = FuncCreateLineTypeToXML(putchFileXml, nameLineType)
                                        idLineTypeRec = recTypeLine.Id
                                    End If
                                    Dim idLayer As ObjectId = AcadFunction.FuncAddLayer(nameLayer, indColor, idLineTypeRec, widthLine)
                                    Return idLayer
                                End If
                            End If
                        ElseIf NameBlock Like "Layers" Then
                            boolFindLayers = True
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LinetypeRecords" Then
                            Exit Function
                        End If
                End Select
            End While
        End If
    End Function
    ' '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает информацию о характеристиках стиля текста из xml и создает стиль текста
    Public Shared Function FuncCreateTextStyleToXML(ByVal putchFileXml As String, ByVal nameTextStyle As String) As ObjectId
        FuncCreateTextStyleToXML = Nothing
        Dim boolItalic As Boolean = True
        Dim NameFont As String = "Arial"
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolFindTextStyle As Boolean = False
            Dim boolFindTextStyles As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "TextStyle" And boolFindTextStyles = True Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If nameN Like "NameTextStyle" Then
                                        If nameTextStyle Like reader.Value Then
                                            boolFindTextStyle = True
                                        Else
                                            boolFindTextStyle = False
                                            Exit While
                                        End If
                                    ElseIf nameN Like "NameFont" Then
                                        NameFont = reader.Value
                                    ElseIf nameN Like "Italic" Then
                                        If reader.Value Like "True" Then
                                            boolItalic = True
                                        Else
                                            boolItalic = False
                                        End If
                                    End If
                                End While
                                If boolFindTextStyle = True Then
                                    Dim idTS As ObjectId = AcadFunction.FuncCreateTextStyle(nameTextStyle, NameFont, boolItalic)
                                    Return idTS
                                End If
                            End If
                        ElseIf NameBlock Like "TextStyles" Then
                            boolFindTextStyles = True
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "TextStyles" Then
                            Exit Function
                        End If
                End Select
            End While
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает стоки таблицы PS из xml (Имя категории, Имя таблицы, Тип объекта, Описание объекта, результат - массив с Характеристиками)
    Public Shared Function FuncReadDataTablePSToXML(ByVal putchFileXml As String, ByVal namePSTable As String, ByRef arrayDictionaryPS As Dictionary(Of Integer, String(,))) As String
        FuncReadDataTablePSToXML = False
        Dim countRec As Integer = 1
        arrayDictionaryPS.Clear()
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolTablePS As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "PropertyTable" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If nameN Like "TableName" Then
                                        If namePSTable Like reader.Value Then
                                            boolTablePS = True
                                        Else
                                            boolTablePS = False
                                        End If
                                    ElseIf nameN Like "TableDescription" Then
                                        If boolTablePS = True Then
                                            FuncReadDataTablePSToXML = reader.Value
                                        End If
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Field" Then
                            If boolTablePS = True Then
                                If reader.HasAttributes = True Then
                                    Dim arrayFieldsPS As String(,) = Nothing
                                    Dim countFieldsPS As Integer = 0
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim val As String = reader.Value
                                        ReDim Preserve arrayFieldsPS(1, countFieldsPS)
                                        arrayFieldsPS(0, countFieldsPS) = nameN
                                        arrayFieldsPS(1, countFieldsPS) = val
                                        countFieldsPS += 1
                                    End While
                                    If IsArray(arrayFieldsPS) = True Then
                                        arrayDictionaryPS.Add(countRec, arrayFieldsPS)
                                        countRec += 1
                                    End If
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "PropertyTable" Then
                            If IsArray(arrayDictionaryPS) = True Then
                                Exit Function
                            End If
                        End If
                End Select
            End While
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает определенную таблицу из файла xml (Имя категории, Имя таблицы, Тип объекта, результат - массив с Характеристиками)
    Public Shared Function FuncReadDataTableToXml(ByVal putchFileXml As String, ByVal category As String, ByVal table As String, ByVal typeObject As String, ByRef dataObjectProperties As String(,)) As Integer
        FuncReadDataTableToXml = 0
        Dim countDataObjectProperties As Integer = 0
        If IsArray(dataObjectProperties) = True Then
            countDataObjectProperties = dataObjectProperties.GetUpperBound(1)
        End If
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolCategory As Boolean = False
            Dim boolTable As Boolean = False
            Dim boolType As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If category Like reader.Value Then
                                        boolCategory = True
                                    Else
                                        boolCategory = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Tables" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If table Like reader.Value Then
                                        boolTable = True
                                    Else
                                        boolTable = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "TypeObject" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If typeObject Like reader.Value Then
                                        boolType = True
                                    Else
                                        boolType = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Rows" Then
                            Dim boolDataFind As Boolean = False
                            If boolTable = True And boolCategory = True And boolType = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        ReDim Preserve dataObjectProperties(1, countDataObjectProperties)
                                        dataObjectProperties(0, countDataObjectProperties) = nameN
                                        dataObjectProperties(1, countDataObjectProperties) = NameObject
                                        countDataObjectProperties += 1
                                    End While
                                End If
                            End If
                        End If
                    Case XmlNodeType.Text
                        Dim a As String = ""
                    Case XmlNodeType.EndElement
                        Dim a As String = ""
                End Select
            End While
        End If
        FuncReadDataTableToXml = countDataObjectProperties
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'процедуры для работы с палитрой
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'изменение размера палитры
    Private Sub TopoZNPalette2_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Try
            Dim s1 As Double = Me.Height 'ширина формы (палитры)
            Dim s2 As Double = Me.Width
            If s1 > 150 Then
                TreeView1.Height = s1 - 60
                Dim pos As Thickness = TreeView1.Margin
                Dim posDataGrid As Thickness = New Thickness(pos.Left, pos.Top + TreeView1.Height, pos.Right, Me.Height - TreeView1.Height - 20)
            End If
            If s2 > 20 Then
                TreeView1.Width = s2 - 15
                DataGrid.Width = s2 - 15
            End If
        Catch ex As Exception
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'первый запуск палитры условных знаков
    Private Sub TopoZNPalette2_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim db As Database = acDoc.Database
        Try
            Dim strDWGName As String = acDoc.Name
            Dim objBase As Object = Application.GetSystemVariable("DWGTITLED")
            If System.Convert.ToInt16(objBase) > 0 Then
                Dim obj1 As Object = Application.GetSystemVariable("DBMOD")
                If Not (System.Convert.ToInt16(obj1) = 0) Then
                    acDoc.Database.SaveAs(strDWGName, True, DwgVersion.Current, acDoc.Database.SecurityParameters)
                End If
            End If
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. Путь к изображениям объектов
        If IO.Directory.Exists(parentPatch2 & "\FileResources\ImageObject\") Then
            PutchImageObject = parentPatch2 & "\FileResources\ImageObject\"
            PutchImageObject = PutchImageObject.Replace("\", "/")
        End If
        '2. Путь к дорожным знакам
        If IO.Directory.Exists(parentPatch2 & "\FileResources\RoadZnaks\") Then
            PutchDorImage = parentPatch2 & "\FileResources\RoadZnaks\"
            PutchDorImage = PutchDorImage.Replace("\", "/")
        End If
        '3. Путь к файлам шаблонам с условными знами
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTemplateDwg = parentPatch2 & "\FileResources\Sample\"
        End If
        '4.читаем шаблоны условных знаков
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '1. получем доступ к пользовательскому каталогу
        'временный каталог для автосохранения таблицы
        Dim TempDir As String = ""
        Dim NameFileTopo As FileStream = Nothing
        Try
            TempDir = GetFolderPath(SpecialFolder.ApplicationData)
            If IO.Directory.Exists(TempDir & "\Civil3DToolsUtility\TOPOTemplates") = False Then
                IO.Directory.CreateDirectory(TempDir & "\Civil3DToolsUtility\TOPOTemplates")
            End If
            If IO.Directory.Exists(TempDir & "\Civil3DToolsUtility\TOPOTemplates") = True Then
                TempDir = TempDir & "\Civil3DToolsUtility\TOPOTemplates"
            End If
        Catch ex As System.Security.SecurityException
        End Try
        '=============================================================================================
        'читаем файлы dwg
        If IO.Directory.Exists(PutchTemplateDwg) = True Then
            Try
                'создаем дерево в TreeView
                For Each fileDWG As String In Directory.GetFiles(PutchTemplateDwg)
                    'читаем расширение
                    Dim ext As String = Path.GetExtension(fileDWG)
                    If ext Like ".dwg" Then
                        Dim NameFileEx As String = Path.GetFileNameWithoutExtension(fileDWG)
                        'формируем текстовые файлы с шаблонами
                        Dim fileXml As String = TempDir & "\" & NameFileEx & ".xml"
                        If IO.File.Exists(fileXml) = False Then
                            Dim boolCreatexml As Boolean = FuncCreateXMLAllListObject(fileDWG, fileXml)
                            Dim destCopyFileDwg As String = TempDir & "\" & NameFileEx & ".dwg"
                            IO.File.Copy(fileDWG, destCopyFileDwg, True)
                        Else
                            Dim fileCreatedDateDwg As Date = File.GetLastWriteTime(fileDWG)
                            Dim fileCreatedDateXml As Date = File.GetLastWriteTime(fileXml)
                            If fileCreatedDateXml < fileCreatedDateDwg Then
                                Dim boolCreatexml As Boolean = FuncCreateXMLAllListObject(fileDWG, fileXml)
                                Dim destCopyFileDwg As String = TempDir & "\" & NameFileEx & ".dwg"
                                IO.File.Copy(fileDWG, destCopyFileDwg, True)
                            End If
                        End If
                        Dim ArrayDWG As String() = {"0", fileDWG, fileXml}
                        Dim node0 As TreeViewItem = New TreeViewItem
                        node0.Header = NameFileEx
                        node0.Tag = ArrayDWG
                        TreeView1.Items.Add(node0)
                        PutchFileTemplateDwg = fileDWG
                    End If
                Next
            Catch ex As IOException
            Catch ex As UnauthorizedAccessException
            Catch ex As ArgumentException
            End Try
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '5.загрузка дорожных знаков
        Dim ArrayDZN As String() = {"DZN", PutchDorImage}
        Dim node01 As TreeViewItem = New TreeViewItem
        node01.Header = "Дорожные знаки"
        node01.Tag = ArrayDZN
        TreeView1.Items.Add(node01)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '6. заполняем список с масштабами
        Dim ArrayAnno As String(,) = Nothing
        Dim boolAnnoM As Boolean = AcadFunction.FuncReadNameAndUnitAnnotativeScale(ArrayAnno)
        Dim anno As AnnotationScale = db.Cannoscale
        Dim AnnoArrayCB As String() = Nothing
        Dim CountAnnoArray As Integer = 0
        ReDim Preserve AnnoArrayCB(CountAnnoArray)
        AnnoArrayCB(CountAnnoArray) = anno.Name
        CountAnnoArray += 1
        If IsArray(ArrayAnno) = True Then
            For i As Integer = 0 To ArrayAnno.GetUpperBound(1)
                ReDim Preserve AnnoArrayCB(CountAnnoArray)
                AnnoArrayCB(CountAnnoArray) = ArrayAnno(0, i)
                CountAnnoArray += 1
            Next i
        End If
        If IsArray(AnnoArrayCB) = True Then
            comboBox.ItemsSource = AnnoArrayCB
            comboBox.Text = anno.Name
        End If
        ps = True
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ВЫБОР УЗЛА ДЕРЕВА
    Private Sub TreeView1_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles TreeView1.SelectedItemChanged
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Try
            Dim TreeViewControl As System.Windows.Controls.TreeView = e.Source
            Dim ObjTreeView As Object = TreeViewControl.SelectedValue
            TreeViewControl.Focus()
            If TypeOf ObjTreeView Is DockPanel Then
                'выбран конкретный условный знак, идет загрузка его на кнопку
                Dim dPan As DockPanel = ObjTreeView
                Erase ArrayParametrObj
                ArrayParametrObj = dPan.Tag
                Button1.Tag = dPan.Tag
                'смотрим элементы
                Dim CountObj As Integer = dPan.Children.Count
                Dim imgObj As System.Windows.Controls.Image = Nothing
                Dim lbObj As System.Windows.Controls.Label = Nothing
                If CountObj = 1 Then
                    lbObj = dPan.Children.Item(0)
                    Dim TextBlock1 As TextBlock = New TextBlock
                    TextBlock1.Text = lbObj.Content
                    TextBlock1.FontSize = 16
                    TextBlock1.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
                    TextBlock1.VerticalAlignment = System.Windows.VerticalAlignment.Center
                    Button1.Content = TextBlock1
                Else
                    imgObj = dPan.Children.Item(0)
                    'ищем заново путь к изображению
                    If IsNothing(imgObj.Source) = False Then
                        Dim PutchImage1 As String = imgObj.Source.ToString
                        Dim uriImage As Uri = New Uri(PutchImage1) 'в качестве миниатюры берем первый найденный файл
                        Dim imgBitMap As BitmapImage = New BitmapImage(uriImage)
                        Dim imageButton As System.Windows.Controls.Image = New System.Windows.Controls.Image
                        imageButton.Source = imgBitMap
                        imageButton.Stretch = Stretch.Uniform
                        Dim DocPanel1 As DockPanel = New DockPanel
                        DocPanel1.Children.Add(imageButton)
                        imgObj.VerticalAlignment = System.Windows.VerticalAlignment.Center
                        imgObj.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
                        Button1.Content = DocPanel1
                        lbObj = dPan.Children.Item(1)
                    Else
                        Button1.Content = "Нет изображения!!!"
                    End If
                End If

                If IsArray(ArrayParametrObj) = True Then
                    If ArrayParametrObj.Length > 6 Then
                        Dim TagObj0 As String = ArrayParametrObj(0)
                        Dim TTP As ToolTip = New ToolTip
                        'читаем поля NameTable 7
                        Dim TagXML As String = ArrayParametrObj(5)
                        Dim TagPS As String = ArrayParametrObj(7)
                        Dim dictPS As Dictionary(Of Integer, String(,)) = New Dictionary(Of Integer, String(,))
                        Dim findPropPS As String = FuncReadDataTablePSToXML(TagXML, TagPS, dictPS)
                        If dictPS.Count > 0 Then
                            custdata.Clear()
                            For i As Integer = 0 To dictPS.Count - 1
                                If dictPS.ContainsKey(i) = True Then
                                    Dim arrayValue As String(,) = dictPS.Item(i)
                                    Dim Acustim As Customer = New Customer()
                                    For j As Integer = 0 To arrayValue.GetUpperBound(1)
                                        Dim nameField As String = arrayValue(0, j)
                                        Dim valStr As String = arrayValue(1, j)
                                        If nameField Like "Name" Then
                                            Acustim.Имя_свойства = valStr
                                        ElseIf nameField Like "DefaultData" Then
                                            Acustim.Значение = valStr
                                        End If
                                    Next j
                                    custdata.Add(Acustim)
                                    DataGrid.HeadersVisibility = False
                                End If
                            Next i
                            DataGrid.ItemsSource = custdata
                        End If
                        If TagObj0 Like "Полилиния" Then
                            TTP.Content = "Выберите необходимые объекты в пространстве модели: линии, полилинии, дуги, окружности и т.п."
                            Button1.ToolTip = TTP
                        ElseIf TagObj0 Like "Блок" Then
                            TTP.Content = "Укажите в пространстве модели точку вставки блока"
                            Button1.ToolTip = TTP
                        ElseIf TagObj0 Like "Дорожный знак" Then
                            Button1.ToolTip = "Укажите блок для присоединения атрибута содержащего в себе индекс дорожного знака"
                        ElseIf TagObj0 Like "Характерная линия" Then
                            Button1.ToolTip = "Укажите полилинию в пространстве модели"
                        ElseIf TagObj0 Like "ТочкаCOGO" Then
                            Button1.Content = "Укажите точку COGO для изменения ее исходного описания"
                        End If
                    End If
                End If
                '================================================================================================
                'смотрим трассы и поверхности
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    'трасса
                    Dim nameAign As String = UserProperties.ComboAlign.Text
                    Dim arrayAlign As String() = Nothing
                    Dim countArrayAlign As Integer = 1
                    Dim alignColl As ObjectIdCollection = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument().GetAlignmentIds
                    ReDim arrayAlign(alignColl.Count)
                    arrayAlign(0) = " "
                    For Each id As ObjectId In alignColl
                        Dim align As Alignment = acTrans.GetObject(id, OpenMode.ForRead)
                        arrayAlign(countArrayAlign) = align.Name
                        countArrayAlign += 1
                    Next
                    UserProperties.ComboAlign.ItemsSource = arrayAlign
                    If arrayAlign.Contains(nameAign) Then
                        UserProperties.ComboAlign.Text = nameAign
                    End If
                    'поверхность
                    Dim nameSurf As String = UserProperties.ComboSurface.Text
                    Dim arraySurface As String() = Nothing
                    Dim countArraySurface As Integer = 1
                    Dim surfaceColl As ObjectIdCollection = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument().GetSurfaceIds
                    ReDim arraySurface(surfaceColl.Count)
                    arraySurface(0) = " "
                    For Each id As ObjectId In surfaceColl
                        Dim surf As TinSurface = acTrans.GetObject(id, OpenMode.ForRead)
                        arraySurface(countArraySurface) = surf.Name
                        countArraySurface += 1
                    Next
                    UserProperties.ComboSurface.ItemsSource = arraySurface
                    If arraySurface.Contains(nameSurf) Then
                        UserProperties.ComboSurface.Text = nameSurf
                    End If
                End Using
            Else
                'вспомогательный элемент, начинаем его разворачивать далее
                'пытаемся развернут элемент
                Dim ISelect As TreeViewItem = ObjTreeView 'ссылка на выбранный элемент
                Dim TypeSelect As String() = ISelect.Tag 'считываем скрытые данные
                Dim TypeItem As String = TypeSelect(0) 'тип выбранной ветки
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбран самый первый родительский элемент
                'элементы в ветку не загружены и выбран родитель (имя файла шаблона)
                If ISelect.Items.Count = 0 And TypeItem Like "0" Then
                    Dim Node1 As TreeViewItem = New TreeViewItem
                    'читаем xml и загружаем ветки с категориями
                    Dim xDoc As XmlDocument = New XmlDocument()
                    Dim PutchDestXml As String = TypeSelect(2) 'категория ветки
                    If IO.File.Exists(PutchDestXml) = True Then
                        Dim reader As XmlTextReader = New XmlTextReader(PutchDestXml)
                        While reader.Read()
                            Select Case reader.NodeType
                                Case XmlNodeType.Element
                                    Dim NameBlock As String = reader.Name 'читаем ветку
                                    If NameBlock Like "Category" Then
                                        If reader.HasAttributes = True Then
                                            While (reader.MoveToNextAttribute())
                                                Dim nameN As String = reader.Name
                                                Dim NameCategory As String = reader.Value
                                                Node1 = New TreeViewItem
                                                Node1.Header = NameCategory
                                                Dim ArrayInf As String() = {"GCategory", NameCategory, PutchDestXml, TypeSelect(1)}
                                                Node1.Tag = ArrayInf
                                                PutchFileTemplateDwg = PutchDestXml
                                                ISelect.Items.Add(Node1)
                                                ISelect.ExpandSubtree()
                                            End While
                                        End If
                                    End If
                                Case XmlNodeType.EndElement
                            End Select
                        End While
                    End If
                    'второй уровень ИМЯ КАТЕГОРИИ (ТОПОГРАФМЯ, ПРОЕКТ ПЛАНИРОВКИ и т.п.)
                ElseIf ISelect.Items.Count = 0 And TypeItem Like "GCategory" Then
                    Dim NameCategory As String = TypeSelect(1) 'имя категории
                    Dim PutchDestXml As String = TypeSelect(2) 'путь к файлу xml
                    Dim PutchDestDwg As String = TypeSelect(3) 'путь к файлу dwg
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    Dim Node1 As TreeViewItem = New TreeViewItem
                    'читаем xml и загружаем ветки с категориями
                    Dim xDoc As XmlDocument = New XmlDocument()
                    If IO.File.Exists(PutchDestXml) = True Then
                        Dim reader As XmlTextReader = New XmlTextReader(PutchDestXml)
                        Dim boolCategory As Boolean = False
                        While reader.Read()
                            Select Case reader.NodeType
                                Case XmlNodeType.Element
                                    Dim NameBlock As String = reader.Name 'читаем ветку
                                    If NameBlock Like "Category" Then
                                        If reader.HasAttributes = True Then
                                            While (reader.MoveToNextAttribute())
                                                Dim nameN As String = reader.Name
                                                Dim TempNameCategory As String = reader.Value
                                                If TempNameCategory Like NameCategory Then
                                                    boolCategory = True
                                                Else
                                                    boolCategory = False
                                                End If
                                            End While
                                        End If
                                    ElseIf NameBlock Like "Table" Then
                                        If reader.HasAttributes = True Then
                                            While (reader.MoveToNextAttribute())
                                                Dim nameN As String = reader.Name
                                                Dim NameTables As String = reader.Value
                                                If boolCategory = True Then
                                                    Node1 = New TreeViewItem
                                                    Node1.Header = NameTables
                                                    Dim ArrayInf As String() = {"Category", NameCategory, NameTables, PutchDestXml, PutchDestDwg}
                                                    Node1.Tag = ArrayInf
                                                    PutchFileTemplateDwg = PutchDestXml
                                                    ISelect.Items.Add(Node1)
                                                    ISelect.ExpandSubtree()
                                                End If
                                            End While
                                        End If
                                    End If
                            End Select
                        End While
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'третий уровень (Геодезические пункты, растительность и т.п.)
                ElseIf ISelect.Items.Count = 0 And TypeItem Like "Category" Then
                    Dim NameCategory As String = TypeSelect(1) 'категория
                    Dim NameTable As String = TypeSelect(2) 'таблица
                    Dim PutchDestXml As String = TypeSelect(3) 'путь к файлу xml
                    Dim PutchDestDwg As String = TypeSelect(4) 'путь к файлу dwg
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    Dim Node1 As TreeViewItem = New TreeViewItem
                    'читаем xml и загружаем ветки с категориями
                    Dim xDoc As XmlDocument = New XmlDocument()
                    If IO.File.Exists(PutchDestXml) = True Then
                        Dim boolCategory As Boolean = False
                        Dim boolTable As Boolean = False

                        Dim reader As XmlTextReader = New XmlTextReader(PutchDestXml)
                        While reader.Read()
                            Select Case reader.NodeType
                                Case XmlNodeType.Element
                                    Dim NameBlock As String = reader.Name 'читаем ветку
                                    If NameBlock Like "Category" Then
                                        If reader.HasAttributes = True Then
                                            While (reader.MoveToNextAttribute())
                                                Dim nameN As String = reader.Name
                                                If NameCategory Like reader.Value Then
                                                    boolCategory = True
                                                Else
                                                    boolCategory = False
                                                End If
                                            End While
                                        End If
                                    ElseIf NameBlock Like "Table" Then
                                        If reader.HasAttributes = True Then
                                            While (reader.MoveToNextAttribute())
                                                Dim nameN As String = reader.Name
                                                If NameTable Like reader.Value Then
                                                    boolTable = True
                                                Else
                                                    boolTable = False
                                                End If
                                            End While
                                        End If
                                    ElseIf NameBlock Like "TypeObject" Then
                                        If boolTable = True And boolCategory = True Then
                                            If reader.HasAttributes = True Then
                                                While (reader.MoveToNextAttribute())
                                                    Dim nameN As String = reader.Name
                                                    Dim TypeObject As String = reader.Value
                                                    Node1 = New TreeViewItem
                                                    Node1.Header = TypeObject
                                                    Dim ArrayInf As String() = {"TypeObject", NameCategory, NameTable, TypeObject, PutchDestXml, PutchDestDwg}
                                                    Node1.Tag = ArrayInf
                                                    PutchFileTemplateDwg = PutchDestXml
                                                    ISelect.Items.Add(Node1)
                                                    ISelect.ExpandSubtree()
                                                End While
                                            End If
                                        End If
                                    End If
                                Case XmlNodeType.EndElement
                            End Select
                        End While
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'выбрана ветка с типом объекта
                ElseIf ISelect.Items.Count = 0 And TypeItem Like "TypeObject" Then
                    Dim NameCategory As String = TypeSelect(1) 'категория
                    Dim NameTable As String = TypeSelect(2) 'таблица
                    Dim TypeObject As String = TypeSelect(3) 'тип объекта
                    Dim PutchDestXml As String = TypeSelect(4) 'путь к файлу xml
                    Dim PutchDestDwg As String = TypeSelect(5) 'путь к файлу dwg
                    Dim tablePS As String = "OBJECT"
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    Dim Node1 As TreeViewItem = New TreeViewItem
                    'читаем xml и загружаем ветки с категориями
                    Dim xDoc As XmlDocument = New XmlDocument()
                    If IO.File.Exists(PutchDestXml) = True Then
                        Dim boolCategory As Boolean = False
                        Dim boolTable As Boolean = False
                        Dim boolType As Boolean = False
                        Dim reader As XmlTextReader = New XmlTextReader(PutchDestXml)
                        While reader.Read()
                            Select Case reader.NodeType
                                Case XmlNodeType.Element
                                    Dim NameBlock As String = reader.Name 'читаем ветку
                                    If NameBlock Like "Category" Then
                                        If reader.HasAttributes = True Then
                                            While (reader.MoveToNextAttribute())
                                                Dim nameN As String = reader.Name
                                                If NameCategory Like reader.Value Then
                                                    boolCategory = True
                                                Else
                                                    boolCategory = False
                                                End If
                                            End While
                                        End If
                                    ElseIf NameBlock Like "Table" Then
                                        If reader.HasAttributes = True Then
                                            While (reader.MoveToNextAttribute())
                                                Dim nameN As String = reader.Name
                                                If NameTable Like reader.Value Then
                                                    boolTable = True
                                                Else
                                                    boolTable = False
                                                End If
                                            End While
                                        End If
                                    ElseIf NameBlock Like "TypeObject" Then
                                        If reader.HasAttributes = True Then
                                            While (reader.MoveToNextAttribute())
                                                Dim nameN As String = reader.Name
                                                If TypeObject Like reader.Value Then
                                                    boolType = True
                                                Else
                                                    boolType = False
                                                End If
                                            End While
                                        End If
                                    ElseIf NameBlock Like "Rows" Then
                                        Dim DestObject As String = ""
                                        If boolTable = True And boolCategory = True And boolType = True Then
                                            If reader.HasAttributes = True Then
                                                While (reader.MoveToNextAttribute())
                                                    Dim nameN As String = reader.Name
                                                    Dim NameObject As String = reader.Value
                                                    If nameN Like "Description" Then
                                                        DestObject = NameObject
                                                    ElseIf nameN Like "PS" Then
                                                        tablePS = NameObject
                                                    ElseIf nameN Like "Putch" Then
                                                        Dim iImage As System.Windows.Controls.Image = New System.Windows.Controls.Image()
                                                        iImage.Width = 60
                                                        iImage.Height = 30
                                                        iImage.Stretch = Stretch.Uniform
                                                        If Directory.Exists(PutchImageObject) = True Then
                                                            Try
                                                                'ищем файл с миниатюрой
                                                                Dim StrImagePutch As String() = Directory.GetFiles(PutchImageObject, NameObject & ".*")
                                                                If IsArray(StrImagePutch) = True Then
                                                                    If StrImagePutch.Length > 0 Then
                                                                        Dim uriImage As Uri = New Uri(StrImagePutch(0)) 'в качестве миниатюры берем первый найденный файл
                                                                        Dim imgBitMap As BitmapImage = New BitmapImage(uriImage)
                                                                        iImage.Source = imgBitMap
                                                                    End If
                                                                End If
                                                            Catch ex As FileNotFoundException
                                                            Catch ex As ArgumentException
                                                            Catch ex As IOException
                                                            Catch ex As IndexOutOfRangeException
                                                            End Try
                                                        End If
                                                        Dim lbStr As System.Windows.Controls.Label = New System.Windows.Controls.Label
                                                        lbStr.Content = DestObject
                                                        Dim docP As DockPanel = New DockPanel
                                                        DockPanel.SetDock(iImage, Dock.Left)
                                                        DockPanel.SetDock(lbStr, Dock.Bottom)
                                                        docP.Children.Add(iImage)
                                                        docP.Children.Add(lbStr)

                                                        Dim strTag As String() = {"Object", NameCategory, NameTable, TypeObject, DestObject, PutchDestXml, PutchDestDwg, tablePS}
                                                        docP.Tag = strTag
                                                        ISelect.Items.Add(docP)
                                                    End If
                                                    ISelect.ExpandSubtree()
                                                End While
                                            End If
                                        End If
                                    End If
                                Case XmlNodeType.EndElement
                            End Select
                        End While
                    End If
                ElseIf ISelect.Items.Count = 0 And TypeItem Like "DZN" Then
                    'считываем блоки
                    Dim PutchImage As String = TypeSelect(1) 'категория ветки
                    If Directory.Exists(PutchImage) = True Then
                        'читаем директорию 
                        'считываем подкаталоги
                        For Each fDir As String In Directory.GetDirectories(PutchImage)
                            Dim pos As Integer = fDir.LastIndexOf("\")
                            If pos = -1 Then
                                pos = fDir.LastIndexOf("/")
                            End If
                            If pos > -1 Then
                                Dim NameDir2 As String = Mid(fDir, pos + 2)
                                Dim Node1 As TreeViewItem = New TreeViewItem
                                Node1.Header = NameDir2
                                Dim ArrayInf As String() = {TypeItem, fDir}
                                Node1.Tag = ArrayInf
                                ISelect.Items.Add(Node1)
                                ISelect.ExpandSubtree()
                            End If
                        Next
                        'считываем файлы
                        For Each ffile As String In Directory.GetFiles(PutchImage)
                            'смотрим если расширение файла не dwg то пропускаем
                            Dim r As String = Path.GetExtension(ffile)
                            If r Like ".png" Then
                                'получаем путь к файлу
                                Dim pos As Integer = ffile.LastIndexOf("\")
                                If pos = -1 Then
                                    pos = ffile.LastIndexOf("/")
                                End If
                                If pos > -1 Then
                                    Dim FullNamefile As String = Left(ffile, pos)
                                    'получаем имя файла без расширения
                                    Dim NameFile = Path.GetFileNameWithoutExtension(ffile)
                                    'отделяем код знака
                                    Dim PosZn1 As Integer = NameFile.LastIndexOf("(")
                                    Dim PosZn2 As Integer = NameFile.LastIndexOf(")")
                                    Dim IndZn As String = Mid(NameFile, PosZn1 + 2, PosZn2 - PosZn1 - 1) 'код знака
                                    Dim NameGOST As String = Mid(NameFile, 4, PosZn1) 'Имя знака
                                    'достаем миниатюру
                                    Dim iImage As System.Windows.Controls.Image = New System.Windows.Controls.Image()
                                    iImage.Width = 30
                                    iImage.Height = 30
                                    iImage.Stretch = Stretch.Uniform
                                    Try
                                        Dim PutchStr As String = ffile
                                        'ищем файл с миниатюрой
                                        Dim StrImagePutch As String() = Directory.GetFiles(FullNamefile, NameFile & ".*")
                                        If StrImagePutch.Length > 0 Then
                                            Dim uriImage As Uri = New Uri(StrImagePutch(0)) 'в качестве миниатюры берем первый найденный файл
                                            Dim imgBitMap As BitmapImage = New BitmapImage(uriImage)
                                            iImage.Source = imgBitMap
                                        End If
                                    Catch ex As FileNotFoundException
                                    Catch ex As ArgumentException
                                    Catch ex As IOException
                                    End Try
                                    Dim lbStr As System.Windows.Controls.Label = New System.Windows.Controls.Label
                                    lbStr.Content = NameFile
                                    Dim docP As DockPanel = New DockPanel
                                    docP.VerticalAlignment = System.Windows.VerticalAlignment.Bottom
                                    docP.Children.Add(iImage)
                                    docP.Children.Add(lbStr)
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'формируем массив для тега (тип, индекс знака, имя знака, путь к файлу, полное имя файла)
                                    Dim strTag As String() = {"DORZN", IndZn, NameGOST, PutchImage, "", "", ""}
                                    docP.Tag = strTag
                                    ISelect.Items.Add(docP)
                                End If
                            End If
                            TreeView1.HorizontalAlignment = HorizontalAlignment.Left
                            TreeView1.VerticalAlignment = VerticalAlignment.Top
                        Next
                    End If
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'НАЖАТА КНОПКА НАРИСОВАТЬ
    Private Sub Button1_Click(sender As Object, e As RoutedEventArgs) Handles Button1.Click
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed1 As Editor = acDoc.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            Dim docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
            Using docLock
                ArrayParametrObj = Button1.Tag
                Try
                    Dim ObjTreeView As Object = TreeView1.SelectedValue
                    If TypeOf ObjTreeView Is DockPanel Then
                        Dim dPan As DockPanel = ObjTreeView
                        Dim arrayData1 As String() = dPan.Tag
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'или топография или дорожный знак
                        If arrayData1(0).Trim Like "DORZN" Then
                            Dim KodDorZn As String = arrayData1(1) 'код знака
                            Dim NameDorZn As String = arrayData1(2) 'имя знака
                            Dim fileDwg As String = arrayData1(3)
                            Dim DeskNameDorZn As String = arrayData1(4) 'описание файла
                            'ищем все настройки пример - знак дорожный на одной стойке
                            Dim ArrayPropZn As String(,) = Nothing
                            Dim NameGroupDorZn As String = ""
                            Dim NameStyleLabel As String = ""
                            Dim NameKod As String = ""
                            Dim NameUDPClass As String = ""
                            Dim NameTablePs As String = ""
                            Dim LocationPointZn As Point3d = Nothing 'координаты дорожного знака
                            Dim NameODField As String(,) = Nothing 'массив для записи расширенных данных
                            Dim IDTablePS As ObjectId = Nothing
                            Try
                                Dim TempDir As String = GetFolderPath(SpecialFolder.ApplicationData)
                                If IO.Directory.Exists(TempDir & "\Civil3DToolsUtility\TOPOTemplates") = True Then
                                    TempDir = TempDir & "\Civil3DToolsUtility\TOPOTemplates"
                                Else
                                    MsgBox("Директория" & TempDir & "\Civil3DToolsUtility\TOPOTemplates" & ", не найдена.")
                                    Exit Sub
                                End If
                                For Each fileXML As String In Directory.GetFiles(TempDir)
                                    Dim ext As String = Path.GetExtension(fileXML)
                                    If ext Like ".xml" Then
                                        Dim CountArrayLineProperties = TOPOSymbols.FuncReadDataTableToXml(fileXML, "Земельный кадастр", "ЕГРН-Кадастровое деление", "Полилиния", ArrayPropZn)
                                        Dim boolFindProp1 As Boolean = FuncReadPropertiesObjectToXML(fileXML, "Топография", "Объекты обустройства автомобильных дорог", "ТочкаCOGO", "Знак дорожный на одной стойке", ArrayPropZn)
                                        If IsArray(ArrayPropZn) = True Then
                                            Exit For
                                        End If
                                    End If
                                Next
                                If IsArray(ArrayPropZn) = False Then
                                    MsgBox("Не удалось найти в файле-шаблоне, описание знака <Знак дорожный на одной стойке>. Макро прервано!!!")
                                    Exit Sub
                                End If
                                For i As Integer = 0 To ArrayPropZn.GetUpperBound(1)
                                    Dim NameField As String = ArrayPropZn(0, i)
                                    If NameField Like "PointGroup" Then
                                        NameGroupDorZn = ArrayPropZn(1, i)
                                    ElseIf NameField Like "LabelStyle" Then
                                        NameStyleLabel = ArrayPropZn(1, i)
                                    ElseIf NameField Like "Name" Then
                                        NameKod = ArrayPropZn(1, i)
                                    ElseIf NameField.Trim Like "PS" Then
                                        NameTablePs = ArrayPropZn(1, i)
                                    End If
                                Next i
                                '=========================================================================================================
                                'находим классификацию
                                Dim udpClassification As UDPClassification = Nothing
                                If CivilDoc.PointUDPClassifications.Contains(NameTablePs) = True Then
                                    udpClassification = CivilDoc.PointUDPClassifications.Item(NameTablePs)
                                Else
                                    'копируем ключ-описатель
                                    Dim boolCreatePG As Boolean = AcadFunction.FuncCopyDeskKeyDestDraw(PutchFileTemplateDwg, NameKod, True)
                                End If
                                If CivilDoc.PointUDPClassifications.Contains(NameTablePs) = True Then
                                    udpClassification = CivilDoc.PointUDPClassifications.Item(NameTablePs)
                                Else
                                    MsgBox("Классификация " & NameTablePs & " не найдена. Макро прервано!!!")
                                End If
                                NameUDPClass = udpClassification.Name
                            Catch ex As System.AccessViolationException
                                MsgBox("Память повреждена! Не удалось выполнить копирование стилей!")
                                Exit Sub
                            Catch ex1 As System.Exception
                                MsgBox("Память повреждена! Не удалось выполнить копирование стилей!")
                                Exit Sub
                            End Try
                            Try
                                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                                    Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)
                                    Dim SelectObj As PromptEntityResult = acDoc.Editor.GetEntity("^C^C " & "Выберите блок или точку COGO: ")
                                    If Not (SelectObj.Status = PromptStatus.OK) Then
                                        Exit Sub 'если нажали ESC то выходим из программы
                                    End If
                                    Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(SelectObj.ObjectId, OpenMode.ForWrite)
                                    If Not IsDBNull(obj) Then
                                        'читаем расширенные данные
                                        Dim boolTablePs As Boolean = FuncPropertySet.FuncFindPSetTable(obj, NameTablePs)
                                        If boolTablePs = True Then
                                            '============================================================================================
                                            'обрабатываем точку COGO
                                            If TypeOf obj Is Autodesk.Civil.DatabaseServices.CogoPoint Then
                                                Dim EntPoint As Autodesk.Civil.DatabaseServices.CogoPoint = obj
                                                '=======================================================================================
                                                'ищем классификацию
                                                Dim UDp1 As Autodesk.Civil.DatabaseServices.UDPString = Nothing 'код знака1
                                                Dim UDP2 As Autodesk.Civil.DatabaseServices.UDPString = Nothing 'код знака2
                                                Dim UDP3 As Autodesk.Civil.DatabaseServices.UDPString = Nothing 'код знака3
                                                Dim UDP4 As Autodesk.Civil.DatabaseServices.UDPString = Nothing 'код знака4
                                                For Each IdUdp As Autodesk.Civil.DatabaseServices.UDP In CivilApplication.ActiveDocument.PointUDPs
                                                    If IsNothing(IdUdp) = False Then
                                                        Dim NameUdp As String = IdUdp.Name
                                                        If IdUdp.Name Like "GOST_NUMBER1" Then
                                                            If TypeOf IdUdp Is Autodesk.Civil.DatabaseServices.UDPString Then 'строка
                                                                UDp1 = TryCast(IdUdp, Autodesk.Civil.DatabaseServices.UDPString)
                                                            End If
                                                        ElseIf IdUdp.Name Like "GOST_NUMBER2" Then
                                                            If TypeOf IdUdp Is Autodesk.Civil.DatabaseServices.UDPString Then 'строка
                                                                UDP2 = TryCast(IdUdp, Autodesk.Civil.DatabaseServices.UDPString)
                                                            End If
                                                        ElseIf IdUdp.Name Like "GOST_NUMBER3" Then
                                                            If TypeOf IdUdp Is Autodesk.Civil.DatabaseServices.UDPString Then 'строка
                                                                UDP3 = TryCast(IdUdp, Autodesk.Civil.DatabaseServices.UDPString)
                                                            End If
                                                        ElseIf IdUdp.Name Like "GOST_NUMBER4" Then
                                                            If TypeOf IdUdp Is Autodesk.Civil.DatabaseServices.UDPString Then 'строка
                                                                UDP4 = TryCast(IdUdp, Autodesk.Civil.DatabaseServices.UDPString)
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                                '====================================================================================================
                                                'тело
                                                If IsNothing(UDp1) = False And IsNothing(UDP2) = False And IsNothing(UDP3) = False And IsNothing(UDP4) = False Then
                                                    'читаем пользовательские свойства
                                                    '1.читаем первое свойство
                                                    Dim result1 As String = ""
                                                    Dim UDpRead1 As Autodesk.Civil.DatabaseServices.UDPString = TryCast(UDp1, Autodesk.Civil.DatabaseServices.UDPString)
                                                    Try
                                                        result1 = EntPoint.GetUDPValue(UDpRead1).ToString
                                                    Catch ex As System.InvalidOperationException
                                                    End Try
                                                    '2.читаем второе свойство
                                                    Dim result2 As String = ""
                                                    Dim UDpRead2 As Autodesk.Civil.DatabaseServices.UDPString = TryCast(UDP2, Autodesk.Civil.DatabaseServices.UDPString)
                                                    Try
                                                        result2 = EntPoint.GetUDPValue(UDpRead2).ToString
                                                    Catch ex As System.InvalidOperationException
                                                    End Try
                                                    '3.читаем третье свойство
                                                    Dim result3 As String = ""
                                                    Dim UDpRead3 As Autodesk.Civil.DatabaseServices.UDPString = TryCast(UDP3, Autodesk.Civil.DatabaseServices.UDPString)
                                                    Try
                                                        result3 = EntPoint.GetUDPValue(UDpRead3).ToString
                                                    Catch ex As System.InvalidOperationException
                                                    End Try
                                                    '4.Читаем четвертое свойство
                                                    Dim result4 As String = ""
                                                    Dim UDpRead4 As Autodesk.Civil.DatabaseServices.UDPString = TryCast(UDP4, Autodesk.Civil.DatabaseServices.UDPString)
                                                    Try
                                                        result4 = EntPoint.GetUDPValue(UDpRead4).ToString
                                                    Catch ex As System.InvalidOperationException
                                                    End Try
                                                    'смотрим 1 свойство
                                                    If result1 Like "" Then
                                                        EntPoint.SetUDPValue(UDp1, KodDorZn)
                                                        ReDim Preserve NameODField(2, 1)
                                                        NameODField(0, 0) = NameUDPClass
                                                        NameODField(1, 0) = "GOST_NUMBER1"
                                                        NameODField(2, 0) = KodDorZn
                                                        NameODField(0, 1) = NameUDPClass
                                                        NameODField(1, 1) = "TEXT1"
                                                        NameODField(2, 1) = DeskNameDorZn
                                                        'смотрим 2 свойство
                                                    ElseIf result2 Like "" Then
                                                        EntPoint.SetUDPValue(UDP2, KodDorZn)
                                                        ReDim Preserve NameODField(2, 1)
                                                        NameODField(0, 0) = NameUDPClass
                                                        NameODField(1, 0) = "GOST_NUMBER2"
                                                        NameODField(2, 0) = KodDorZn
                                                        NameODField(0, 1) = NameUDPClass
                                                        NameODField(1, 1) = "TEXT2"
                                                        NameODField(2, 1) = DeskNameDorZn
                                                        'смотрим 3 свойство
                                                    ElseIf result3 Like "" Then
                                                        EntPoint.SetUDPValue(UDP3, KodDorZn)
                                                        ReDim Preserve NameODField(2, 1)
                                                        NameODField(0, 0) = NameUDPClass
                                                        NameODField(1, 0) = "GOST_NUMBER3"
                                                        NameODField(2, 0) = KodDorZn
                                                        NameODField(0, 1) = NameUDPClass
                                                        NameODField(1, 1) = "TEXT3"
                                                        NameODField(2, 1) = DeskNameDorZn
                                                        'смотрим 4 свойство
                                                    ElseIf result4 Like "" Then
                                                        EntPoint.SetUDPValue(UDP4, KodDorZn)
                                                        ReDim Preserve NameODField(2, 1)
                                                        NameODField(0, 0) = NameUDPClass
                                                        NameODField(1, 0) = "GOST_NUMBER4"
                                                        NameODField(2, 0) = KodDorZn
                                                        NameODField(0, 1) = NameUDPClass
                                                        NameODField(1, 1) = "TEXT4"
                                                        NameODField(2, 1) = DeskNameDorZn
                                                    Else
                                                        EntPoint.SetUDPValue(UDp1, KodDorZn)
                                                        ReDim Preserve NameODField(2, 1)
                                                        NameODField(0, 0) = NameUDPClass
                                                        NameODField(1, 0) = "GOST_NUMBER1"
                                                        NameODField(2, 0) = KodDorZn
                                                        NameODField(0, 1) = NameUDPClass
                                                        NameODField(1, 1) = "TEXT1"
                                                        NameODField(2, 1) = DeskNameDorZn
                                                    End If
                                                    Dim ArrayTables As String() = Nothing
                                                    Dim boolfindTables As Boolean = FuncPropertySet.FuncFindPSetTable(EntPoint, NameTablePs)
                                                    If boolfindTables = True Then
                                                        Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(EntPoint, NameODField)
                                                    End If
                                                End If
                                                '================================================================================================
                                                'если это блок
                                            ElseIf TypeOf obj Is BlockReference Then
                                                Dim BlkObj As BlockReference = obj 'доступ к блоку
                                                Dim NameBlk As String = BlkObj.Name 'полное описание объекта
                                                Dim ArrayArrtib As String(,) = Nothing
                                                Dim FindBoolArrtibBlk As Boolean = AcadFunction.FuncReadAllBlockAttr(BlkObj, ArrayArrtib)
                                                Dim NameFieldWrite As String = "GOST_NUMBER1"
                                                Dim boolFindField As Boolean = False
                                                Dim ArrayDataBlk As String(,) = Nothing
                                                If IsArray(ArrayArrtib) = True Then
                                                    For i As Integer = 0 To ArrayArrtib.GetUpperBound(1)
                                                        Dim NameField As String = ArrayArrtib(0, i).Trim
                                                        Dim ValZn As String = ArrayArrtib(1, i).Trim
                                                        If IsNothing(NameField) = False Then
                                                            If NameField.IndexOf("GOST_NUMBER1") > -1 Then
                                                                If ValZn.Trim.Length = 0 Or ValZn.Trim = "0" Or ValZn.Trim = "-" Then
                                                                    boolFindField = True
                                                                    '==========================================================================================
                                                                    'формируем массив для записи
                                                                    ReDim Preserve ArrayDataBlk(1, 0)
                                                                    ArrayDataBlk(0, 0) = NameField
                                                                    ArrayDataBlk(1, 0) = KodDorZn
                                                                    '==========================================================================================
                                                                    'формируем массив для расширенных данных
                                                                    ReDim Preserve NameODField(2, 1)
                                                                    NameODField(0, 0) = NameTablePs
                                                                    NameODField(1, 0) = "GOST_NUMBER1"
                                                                    NameODField(2, 0) = KodDorZn
                                                                    NameODField(0, 1) = NameTablePs
                                                                    NameODField(1, 1) = "TEXT1"
                                                                    NameODField(2, 1) = DeskNameDorZn
                                                                    boolFindField = True
                                                                    Exit For
                                                                End If
                                                            ElseIf NameField.IndexOf("GOST_NUMBER2") = 0 Then
                                                                If ValZn.Trim.Length = 0 Or ValZn.Trim = "0" Or ValZn.Trim = "-" Then
                                                                    boolFindField = True
                                                                    '==========================================================================================
                                                                    'формируем массив для записи
                                                                    ReDim Preserve ArrayDataBlk(1, 0)
                                                                    ArrayDataBlk(0, 0) = NameField
                                                                    ArrayDataBlk(1, 0) = KodDorZn
                                                                    '==========================================================================================
                                                                    'формируем массив для расширенных данных
                                                                    ReDim Preserve NameODField(2, 1)
                                                                    NameODField(0, 0) = NameTablePs
                                                                    NameODField(1, 0) = "GOST_NUMBER2"
                                                                    NameODField(2, 0) = KodDorZn
                                                                    NameODField(0, 1) = NameTablePs
                                                                    NameODField(1, 1) = "TEXT2"
                                                                    NameODField(2, 1) = DeskNameDorZn
                                                                    boolFindField = True
                                                                    Exit For
                                                                End If
                                                            ElseIf NameField.IndexOf("GOST_NUMBER3") = 0 Then
                                                                If ValZn.Trim.Length = 0 Or ValZn.Trim = "0" Or ValZn.Trim = "-" Then
                                                                    boolFindField = True
                                                                    '==========================================================================================
                                                                    'формируем массив для записи
                                                                    ReDim Preserve ArrayDataBlk(1, 0)
                                                                    ArrayDataBlk(0, 0) = NameField
                                                                    ArrayDataBlk(1, 0) = KodDorZn
                                                                    '==========================================================================================
                                                                    'формируем массив для расширенных данных
                                                                    ReDim Preserve NameODField(2, 1)
                                                                    NameODField(0, 0) = NameTablePs
                                                                    NameODField(1, 0) = "GOST_NUMBER3"
                                                                    NameODField(2, 0) = KodDorZn
                                                                    NameODField(0, 1) = NameTablePs
                                                                    NameODField(1, 1) = "TEXT3"
                                                                    NameODField(2, 1) = DeskNameDorZn
                                                                    boolFindField = True
                                                                    Exit For
                                                                End If
                                                            ElseIf NameField.IndexOf("GOST_NUMBER4") = 0 Then
                                                                If ValZn.Trim.Length = 0 Or ValZn.Trim = "0" Or ValZn.Trim = "-" Then
                                                                    boolFindField = True
                                                                    '==========================================================================================
                                                                    'формируем массив для записи
                                                                    ReDim Preserve ArrayDataBlk(1, 0)
                                                                    ArrayDataBlk(0, 0) = NameField
                                                                    ArrayDataBlk(1, 0) = KodDorZn
                                                                    '==========================================================================================
                                                                    'формируем массив для расширенных данных
                                                                    ReDim Preserve NameODField(2, 1)
                                                                    NameODField(0, 0) = NameTablePs
                                                                    NameODField(1, 0) = "GOST_NUMBER4"
                                                                    NameODField(2, 0) = KodDorZn
                                                                    NameODField(0, 1) = NameTablePs
                                                                    NameODField(1, 1) = "TEXT4"
                                                                    NameODField(2, 1) = DeskNameDorZn
                                                                    boolFindField = True
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    Next i
                                                End If
                                                If boolFindField = True Then
                                                    Dim BoolArrtibBlk As Boolean = AcadFunction.FuncWriteAllBlockAttr(BlkObj, ArrayDataBlk)
                                                    Dim boolData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(BlkObj, NameODField)
                                                End If
                                            Else
                                                MsgBox("Не верный тип объекта. Выбрать можно только COGO точку или Блок. Макро прервано!!!")
                                                Exit Sub
                                            End If
                                        Else
                                            'вставляем вспомогательный блок
                                            Dim NameBlkDesk As String = "Знак дорожный расположенный на опоре"
                                            Dim NameBlk As String = ""
                                            Dim ArrayPropBlk As String(,) = Nothing
                                            Dim boolPropBlkUPK As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchFileTemplateDwg, NameBlkDesk, ArrayPropBlk, "Блок")
                                            If IsArray(ArrayPropBlk) = True Then
                                                For i As Integer = 0 To ArrayPropBlk.GetUpperBound(1)
                                                    Dim NameField As String = ArrayPropBlk(0, i)
                                                    If NameField.Trim Like "Имя объекта" Then
                                                        NameBlk = ArrayPropBlk(1, i)
                                                        If IsNothing(NameBlk) = False Then
                                                            If NameBlk.Trim.Length > 0 Then
                                                                Dim boolInsRefBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchFileTemplateDwg, NameBlk)
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                Next i
                                            End If
                                            If bt.Has(NameBlk) = True Then
                                                'вставляем блок
                                                If TypeOf obj Is Autodesk.Civil.DatabaseServices.CogoPoint Then
                                                    Dim EntPoint As Autodesk.Civil.DatabaseServices.CogoPoint = obj
                                                    LocationPointZn = EntPoint.Location
                                                ElseIf TypeOf obj Is BlockReference Then
                                                    Dim BlkObj As BlockReference = obj 'доступ к блоку
                                                    LocationPointZn = BlkObj.Position
                                                Else
                                                    MsgBox("Не верный тип объекта. Выбрать можно только COGO точку или Блок. Макро прервано!!!")
                                                    Exit Sub
                                                End If
                                                Dim blkZn As BlockReference = AcadFunction.FuncInsBlkByArrayProperties(LocationPointZn, ArrayPropBlk, 0)
                                                If IsNothing(blkZn) = False Then
                                                    If blkZn.Id.IsNull = False Then
                                                        Dim arrayDataBlk As String(,) = Nothing
                                                        ReDim Preserve arrayDataBlk(1, 0)
                                                        arrayDataBlk(0, 0) = "GOST_NUMBER1"
                                                        arrayDataBlk(1, 0) = KodDorZn
                                                        Dim BoolwriteArrtibBlk As Boolean = AcadFunction.FuncWriteAllBlockAttr(blkZn, arrayDataBlk)
                                                    End If
                                                End If
                                                Dim boolInsTable As Boolean = FuncPropertySet.FuncAddDataSetProperties(blkZn, IDTablePS)
                                                If boolInsTable = True Then
                                                    Dim boolData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(blkZn, NameODField)
                                                End If
                                            End If
                                        End If
line6:
                                    End If
                                    acTrans.Commit()
                                End Using
                            Catch ex As System.AccessViolationException
                            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                            Catch ex As System.Exception
                            End Try
                        Else
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'рисование условных знаков
                            Dim ArrayODRec As String(,) = Nothing
                            Dim categoryZn As String = arrayData1(1)
                            Dim tableZn As String = arrayData1(2)
                            Dim typeZn As String = arrayData1(3)
                            Dim deskZn As String = arrayData1(4)
                            Dim putchFileXML As String = arrayData1(5)
                            Dim putchFileDWG As String = arrayData1(6)
                            Dim acSurf As TinSurface = Nothing
                            Dim acAlign As Alignment = Nothing
                            Dim offset As Double = 0
                            'достаем масштаб файла-шаблона
                            Dim MDest As Double = AcadFunction.FuncReadCurrentAnnoDestDraw(putchFileDWG)
                            'определяем масштаб целевого чертежа
                            Dim NameAnno As String = comboBox.Text
                            Dim annoScale As AnnotationScale = AcadFunction.FuncReadAnnotativeScaleDWGByName(NameAnno)
                            Dim M As Double = 1
                            If IsNothing(annoScale) = False Then
                                Dim MDraw As Double = (annoScale.PaperUnits / annoScale.Scale) * 1000
                                M = MDraw / MDest
                            End If
                            Dim ArrayParametrZn As String(,) = Nothing
                            Dim boolFindDataXml As Boolean = FuncReadPropertiesObjectToXML(putchFileXML, categoryZn, tableZn, typeZn, deskZn, ArrayParametrZn)
                            '==========================================================================================================================
                            'описание
                            Dim DeskObject As String = ""
                            'имя объекта
                            Dim NameObject As String = ""
                            Dim IDNameObject As ObjectId = Nothing
                            'слой
                            Dim NameLayer As String = "0"
                            Dim IDNameLayer As ObjectId = Nothing
                            Dim NameLayerIsh As String = ""
                            'цвет
                            Dim IndColor As Integer = 256
                            'вес линии
                            Dim LineWidth As Integer = 20
                            'масштаб объекта
                            Dim Scale As Double = 1
                            'масштаб линии
                            Dim LineScale As Double = 1
                            'тип линии для блоков и т.п.
                            Dim LineTypeName As String = "ПоСлою"
                            Dim IDLineTypeName As ObjectId = Nothing
                            'глобальная ширина
                            Dim GScale As Double = 0
                            'PS
                            Dim PSTable As String = ""
                            Dim idPsTable As ObjectId = Nothing
                            'группа точек
                            Dim PointGroup As String = ""
                            Dim idPointGroup As ObjectId = Nothing
                            'стиль маркера
                            Dim styleMarker As String = ""
                            Dim idStyleMarker As ObjectId = Nothing
                            'площадка
                            Dim nameSite As String = ""
                            Dim IdSite As ObjectId = Nothing
                            'высота ячейки таблицы
                            Dim rowheight As Double = 4
                            'высота текста
                            Dim textheight As Double = 2.5
                            'цвет текста
                            Dim textColor As Integer = 256
                            'стиль текста
                            Dim TextStyle As String = ""
                            Dim idTextStyle As ObjectId = Nothing
                            'высота стрелки для мультивыноски
                            Dim arrowSize As Double = 2
                            'примечание
                            Dim Note As String = ""
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'копирование объектов их внешнего файла
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                                If IsArray(ArrayParametrZn) = True Then
                                    For i As Integer = 0 To ArrayParametrZn.GetUpperBound(1)
                                        Dim dataField As String = ArrayParametrZn(0, i)
                                        Dim val As String = ArrayParametrZn(1, i)
                                        If IsNothing(val) = False Then
                                            If val.Trim.Length > 0 Then
                                                If dataField Like "Name" Then
                                                    If typeZn Like "Полилиния" Or typeZn Like "Полигон" Then
                                                        'копируем тип линии
                                                        'FuncCreateLineTypeToXML
                                                        Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(acCurDb.LinetypeTableId, OpenMode.ForRead)
                                                        If acLineTypTbl.Has(val) = False Then
                                                            Dim recTypeline As LinetypeTableRecord = FuncCreateLineTypeToXML(putchFileXML, val)
                                                            'Dim boolTypeline As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(putchFileDWG, val)
                                                        Else
                                                            If UserProperties.CheboxStyle.IsChecked = True Then
                                                                Dim boolTypeline As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(putchFileDWG, val, True)
                                                            Else
                                                                Dim boolTypeline As Boolean = True
                                                            End If
                                                        End If
                                                    ElseIf typeZn Like "Блок" Then
                                                        'копируем определение блока
                                                        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                                                        If acBlkTbl.Has(val) = False Then
                                                            Dim boolinsBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(putchFileDWG, val)
                                                        Else
                                                            If UserProperties.CheboxStyle.IsChecked = True Then
                                                                Dim boolinsBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(putchFileDWG, val, True)
                                                            Else
                                                                Dim boolinsBlk As Boolean = True
                                                            End If
                                                        End If
                                                    ElseIf typeZn Like "Мультилиния" Then
                                                        'копируем стиль мультилинии
                                                        Dim mlDict As DBDictionary = acTrans.GetObject(acCurDb.MLStyleDictionaryId, OpenMode.ForRead)
                                                        If mlDict.Contains(val) = False Then
                                                            Dim boolTypeMline As Boolean = AcadFunction.FuncCreateTypeMLineDestDraw(putchFileDWG, val)
                                                        Else
                                                            If UserProperties.CheboxStyle.IsChecked = True Then
                                                                Dim boolTypeMline As Boolean = AcadFunction.FuncCreateTypeMLineDestDraw(putchFileDWG, val, True)
                                                            Else
                                                                Dim boolTypeMline As Boolean = True
                                                            End If
                                                        End If
                                                    ElseIf typeZn Like "Мультивыноска" Then
                                                        'копируем стиль мультивыноски
                                                        Dim IDMLeaderStyles As ObjectId = AcadFunction.FuncReturnIDMLeaderStylebyName(val)
                                                        If IsNothing(IDMLeaderStyles) = True Then
                                                            Dim boolStyleMLeader As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(putchFileDWG, val)
                                                        ElseIf IDMLeaderStyles.IsNull = True Then
                                                            Dim boolStyleMLeader As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(putchFileDWG, val)
                                                        Else
                                                            If UserProperties.CheboxStyle.IsChecked = True Then
                                                                Dim boolStyleMLeader As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(putchFileDWG, val, True)
                                                            Else
                                                                Dim boolStyleMLeader As Boolean = True
                                                            End If
                                                        End If
                                                    ElseIf typeZn Like "Таблица" Then
                                                        'копируем стиль таблицы
                                                        Dim acTableStyleTbl As DBDictionary = acTrans.GetObject(acCurDb.TableStyleDictionaryId, OpenMode.ForRead)
                                                        Dim findTable As Boolean = False
                                                        For Each id As DBDictionaryEntry In acTableStyleTbl
                                                            If id.Key Like val Then
                                                                findTable = True
                                                                Exit For
                                                            End If
                                                        Next
                                                        If findTable = False Then
                                                            Dim boolCopyTable As Boolean = AcadFunction.FuncCopyTablePSDestDraw(putchFileDWG, val)
                                                        Else
                                                            If UserProperties.CheboxStyle.IsChecked = True Then
                                                                Dim boolCopyTable As Boolean = AcadFunction.FuncCopyTablePSDestDraw(putchFileDWG, val)
                                                            Else
                                                                Dim boolCopyTable As Boolean = True
                                                            End If
                                                        End If
                                                    ElseIf typeZn Like "Размер" Then
                                                        'копируем стиль таблицы
                                                        Dim acDimStyleTbl As DimStyleTable = acTrans.GetObject(acCurDb.DimStyleTableId, OpenMode.ForRead)
                                                        If acDimStyleTbl.Has(val) = False Then
                                                            Dim boolinsBlk As Boolean = AcadFunction.FuncCreateDimStyleDestDraw(putchFileDWG, val)
                                                        Else
                                                            If UserProperties.CheboxStyle.IsChecked = True Then
                                                                Dim boolinsBlk As Boolean = AcadFunction.FuncCreateDimStyleDestDraw(putchFileDWG, val, True)
                                                            Else
                                                                Dim boolinsBlk As Boolean = True
                                                            End If
                                                        End If
                                                    ElseIf typeZn Like "Характерная линия" Then
                                                        'копируем стиль характерной линии
                                                        Dim idFeatureLineStyle As ObjectId = Nothing
                                                        Try
                                                            idFeatureLineStyle = CivilDoc.Styles.FeatureLineStyles.Item(val)
                                                        Catch ex As System.ArgumentException
                                                        End Try
                                                        If IsNothing(idFeatureLineStyle) = False Then
                                                            If idFeatureLineStyle.IsNull = True Then
                                                                Dim boolCreateFeatyreLineStyle As Boolean = AcadFunction.FuncCopyFeatureLineStyleDestDraw(putchFileDWG, val)
                                                            Else
                                                                If UserProperties.CheboxStyle.IsChecked = True Then
                                                                    Dim boolCreateFeatyreLineStyle As Boolean = AcadFunction.FuncCopyFeatureLineStyleDestDraw(putchFileDWG, val, True)
                                                                End If
                                                            End If
                                                        Else
                                                            If UserProperties.CheboxFeatureLine.IsChecked = True Then
                                                                Dim boolCreateFeatyreLineStyle As Boolean = AcadFunction.FuncCopyFeatureLineStyleDestDraw(putchFileDWG, val, True)
                                                            End If
                                                        End If
                                                    ElseIf typeZn Like "ТочкаCOGO" Then
                                                        'копируем ключ-описатель
                                                        Dim pdKeySetIds As ObjectIdCollection = Autodesk.Civil.DatabaseServices.PointDescriptionKeySetCollection.GetPointDescriptionKeySets(acCurDb).SearchOrder
                                                        Dim boolKey As Boolean = False
                                                        If pdKeySetIds.Count > 0 Then
                                                            For Each idKeySet As ObjectId In pdKeySetIds
                                                                Dim DeskKeySet As Autodesk.Civil.DatabaseServices.PointDescriptionKeySet = acTrans.GetObject(idKeySet, OpenMode.ForRead)
                                                                Dim idDesk As ObjectId = Nothing
                                                                Try
                                                                    idDesk = DeskKeySet.Item(val)
                                                                Catch ex As System.ArgumentException
                                                                End Try
                                                                If IsNothing(idDesk) = False Then
                                                                    If idDesk.IsNull = False Then
                                                                        boolKey = True
                                                                        Exit For
                                                                    Else
                                                                        boolKey = False
                                                                    End If
                                                                Else
                                                                    boolKey = False
                                                                End If
                                                            Next
                                                        End If
                                                        If boolKey = False Then
                                                            Dim boolCreatePG As Boolean = AcadFunction.FuncCopyDeskKeyDestDraw(putchFileDWG, val)
                                                        Else
                                                            If UserProperties.CheboxStyle.IsChecked = True Then
                                                                Dim boolCreatePG As Boolean = AcadFunction.FuncCopyDeskKeyDestDraw(putchFileDWG, val, True)
                                                            End If
                                                        End If
                                                    ElseIf typeZn Like "Label" Then

                                                    ElseIf typeZn Like "NoteLineLabel" Then
                                                        'копируем стиль метки линии
                                                        Dim idGeneralLineLabel As ObjectId = Nothing
                                                        Try
                                                            idGeneralLineLabel = CivilDoc.Styles.LabelStyles.GeneralLineLabelStyles.Item(val)
                                                        Catch ex As System.ArgumentException
                                                        End Try
                                                        If IsNothing(idGeneralLineLabel) = False Then
                                                            If idGeneralLineLabel.IsNull = True Then
                                                                Dim boolCreateLabelLineStyle As Boolean = AcadFunction.FuncCopyLabelLineStyleDestDraw(putchFileDWG, val)
                                                            Else
                                                                If UserProperties.CheboxStyle.IsChecked = True Then
                                                                    Dim boolCreateLabelLineStyle As Boolean = AcadFunction.FuncCopyLabelLineStyleDestDraw(putchFileDWG, val, True)
                                                                End If
                                                            End If
                                                        Else
                                                            If UserProperties.CheboxStyle.IsChecked = True Then
                                                                Dim boolCreateLabelLineStyle As Boolean = AcadFunction.FuncCopyLabelLineStyleDestDraw(putchFileDWG, val, True)
                                                            End If
                                                        End If
                                                    End If
                                                    '======================================================================================================================
                                                    'копируем слой
                                                ElseIf dataField Like "Layer" Then
                                                    Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                                                    If acLyrTbl.Has(val) = False Then
                                                        If Not (Mid(val, 1, 1) Like "_") Then
                                                            Dim idLayer As ObjectId = FuncCreateLayerToXML(putchFileXML, val)
                                                            'Dim boolLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(putchFileDWG, val)
                                                        End If
                                                    End If
                                                    '======================================================================================================================
                                                    'копируем тип линии для блоков и т.п.
                                                ElseIf dataField Like "LineType" Then
                                                    If val Like "ПоСлою" = True Then
                                                        LineTypeName = "ByLayer"
                                                    ElseIf val Like "ПоБлоку" Then
                                                        LineTypeName = "ByBlock"
                                                    Else
                                                        Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(acCurDb.LinetypeTableId, OpenMode.ForRead)
                                                        If acLineTypTbl.Has(val) = False Then
                                                            If UserProperties.CheboxStyle.IsChecked = True Then
                                                                Dim boolNameTypeLine As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(putchFileDWG, val, True)
                                                                If boolNameTypeLine = False Then
                                                                    LineTypeName = "ByLayer"
                                                                End If
                                                            Else
                                                                'Dim boolNameTypeLine As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(putchFileDWG, val)
                                                                Dim recTypeline As LinetypeTableRecord = FuncCreateLineTypeToXML(putchFileXML, val)
                                                                If IsNothing(recTypeline) = True Then
                                                                    LineTypeName = "ByLayer"
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    '=============================================================================================================
                                                    'копируем стиль текста
                                                ElseIf dataField Like "TextStyle" Then
                                                    idTextStyle = AcadFunction.FuncReturnIDTextStylebyName(val)
                                                    If IsNothing(idTextStyle) = True Then
                                                        Dim boolCopyStyleText As Boolean = AcadFunction.FuncCreateTextStyleDestDraw(putchFileDWG, val)
                                                    ElseIf idTextStyle.IsNull = True Then
                                                        Dim boolCopyStyleText As Boolean = AcadFunction.FuncCreateTextStyleDestDraw(putchFileDWG, val)
                                                    End If
                                                    '======================================================================================================================
                                                    'копируем таблицу PS
                                                ElseIf dataField Like "PS" Then
                                                    Dim dictPropSetDef As Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(acCurDb)
                                                    If dictPropSetDef.Has(val, acTrans) = False Then
                                                        Dim idPs As ObjectId = FuncPropertySet.FuncCreateTablePSByXML(putchFileXML, val)
                                                        'Dim boolInsPs As Boolean = AcadFunction.FuncCopyTablePSDestDraw(putchFileDWG, val)
                                                    End If
                                                    '======================================================================================================================
                                                    'копируем группу точек
                                                ElseIf dataField Like "PointGroup" Then
                                                    Dim idGroupPoint As ObjectId = Nothing
                                                    Try
                                                        idGroupPoint = CivilDoc.PointGroups.Item(val)
                                                    Catch ex As System.ArgumentException
                                                    End Try
                                                    If IsNothing(idGroupPoint) = False Then
                                                        If idGroupPoint.IsNull = True Then
                                                            Dim boolCreatePG As Boolean = AcadFunction.FuncCopyGroupPointsDestDraw(putchFileDWG, val)
                                                        End If
                                                    Else
                                                        Dim boolCreatePG As Boolean = AcadFunction.FuncCopyGroupPointsDestDraw(putchFileDWG, val)
                                                    End If
                                                    '======================================================================================================================
                                                    'копируем площадку
                                                ElseIf dataField Like "Site" Then
                                                    Dim idcolsite As ObjectIdCollection = CivilDoc.GetSiteIds
                                                    If idcolsite.Count > 0 Then
                                                        For j As Integer = 0 To idcolsite.Count - 1
                                                            Dim SiteObj As Autodesk.Civil.DatabaseServices.Site = acTrans.GetObject(idcolsite.Item(j), OpenMode.ForRead)
                                                            If SiteObj.Name Like val Then
                                                                IdSite = SiteObj.Id
                                                                Exit For
                                                            End If
                                                        Next j
                                                    End If
                                                    If IsNothing(IdSite) = True Then
                                                        If IdSite.IsNull = True Then
                                                            Try
                                                                IdSite = Autodesk.Civil.DatabaseServices.Site.Create(CivilDoc, val)
                                                            Catch ex As System.ArgumentException
                                                            End Try
                                                        End If
                                                    End If
                                                End If
                                                '======================================================================================================================
                                                'копируем стиль маркера
                                            ElseIf dataField Like "MarkerStyle" Then
                                                Dim idMarkerStyle As ObjectId = Nothing
                                                Try
                                                    idMarkerStyle = CivilDoc.Styles.MarkerStyles.Item(val)
                                                Catch ex As System.ArgumentException
                                                End Try
                                                If IsNothing(idMarkerStyle) = False Then
                                                    If idMarkerStyle.IsNull = True Then
                                                        Dim boolCreateLabelLineStyle As Boolean = AcadFunction.FuncCopyLabelLineStyleDestDraw(putchFileDWG, val)
                                                    Else
                                                        If UserProperties.CheboxStyle.IsChecked = True Then
                                                            Dim boolCreateLabelLineStyle As Boolean = AcadFunction.FuncCopyLabelLineStyleDestDraw(putchFileDWG, val, True)
                                                        End If
                                                    End If
                                                Else
                                                    If UserProperties.CheboxStyle.IsChecked = True Then
                                                        Dim boolCreateLabelLineStyle As Boolean = AcadFunction.FuncCopyLabelLineStyleDestDraw(putchFileDWG, val, True)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next i
                                End If
                                acTrans.Commit()
                            End Using
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'назначаем характеристики объекту
                            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                                If IsArray(ArrayParametrZn) = True Then
                                    For i As Integer = 0 To ArrayParametrZn.GetUpperBound(1)
                                        Dim dataField As String = ArrayParametrZn(0, i)
                                        Dim val As String = ArrayParametrZn(1, i)
                                        If IsNothing(val) = False Then
                                            If val.Trim.Length > 0 Then
                                                If dataField Like "Name" Then
                                                    If typeZn Like "Полилиния" Or typeZn Like "Полигон" Then
                                                        'копируем тип линии
                                                        Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(acCurDb.LinetypeTableId, OpenMode.ForRead)
                                                        If acLineTypTbl.Has(val) = True Then
                                                            IDNameObject = acLineTypTbl.Item(val)
                                                            NameObject = val
                                                            Dim recTypeLine As LinetypeTableRecord = acTrans.GetObject(IDNameObject, OpenMode.ForRead)
                                                            'Dim str As String = AcadFunction.FuncReturnTypeLineCode(recTypeLine)
                                                            'MsgBox(str)
                                                        Else
                                                            MsgBox("Не удалось скопировать стиль линии из шаблона!!! Макро прервано.")
                                                            Exit Sub
                                                        End If
                                                    ElseIf typeZn Like "Блок" Then
                                                        'копируем определение блока
                                                        If acBlkTbl.Has(val) = True Then
                                                            IDNameObject = acBlkTbl.Item(val)
                                                            NameObject = val
                                                        Else
                                                            MsgBox("Не удалось скопировать определение блока из шаблона!!! Макро прервано.")
                                                            Exit Sub
                                                        End If
                                                    ElseIf typeZn Like "Мультилиния" Then
                                                        'копируем стиль мультилинии
                                                        Dim mlDict As DBDictionary = acTrans.GetObject(acCurDb.MLStyleDictionaryId, OpenMode.ForRead)
                                                        If mlDict.Contains(val) = True Then
                                                            IDNameObject = mlDict.Item(val)
                                                            NameObject = val
                                                        Else
                                                            MsgBox("Не удалось скопировать стиль мультилинии из шаблона!!! Макро прервано.")
                                                            Exit Sub
                                                        End If
                                                    ElseIf typeZn Like "Мультивыноска" Then
                                                        'копируем стиль мультивыноски
                                                        IDNameObject = AcadFunction.FuncReturnIDMLeaderStylebyName(val)
                                                        If IsNothing(IDNameObject) = False Then
                                                            If IDNameObject.IsNull = False Then
                                                                NameObject = val
                                                            Else
                                                                MsgBox("Не удалось скопировать стиль мультивыноски из шаблона!!! Макро прервано.")
                                                                Exit Sub
                                                            End If
                                                        End If
                                                    ElseIf typeZn Like "Таблица" Then
                                                        'копируем стиль таблицы
                                                        Dim acTableStyleTbl As DBDictionary = acTrans.GetObject(acCurDb.TableStyleDictionaryId, OpenMode.ForRead)
                                                        Dim findTable As Boolean = False
                                                        For Each id As DBDictionaryEntry In acTableStyleTbl
                                                            If id.Key Like val Then
                                                                findTable = True
                                                                IDNameObject = acTableStyleTbl.GetAt(val)
                                                                NameObject = val
                                                                Exit For
                                                            End If
                                                        Next
                                                        If findTable = False Then
                                                            MsgBox("Не удалось скопировать стиль Таблицы из шаблона!!! Макро прервано.")
                                                            Exit Sub
                                                        End If
                                                    ElseIf typeZn Like "Размер" Then
                                                        'копируем размерный стиль
                                                        Dim acDimStyleTbl As DimStyleTable = acTrans.GetObject(acCurDb.DimStyleTableId, OpenMode.ForRead)
                                                        If acDimStyleTbl.Has(val) = True Then
                                                            IDNameObject = acDimStyleTbl.Item(val)
                                                            NameObject = val
                                                        Else
                                                            MsgBox("Не удалось скопировать Размерный стиль из шаблона!!! Макро прервано.")
                                                            Exit Sub
                                                        End If
                                                    ElseIf typeZn Like "Характерная линия" Then
                                                        'копируем стиль характерной линии
                                                        Try
                                                            IDNameObject = CivilDoc.Styles.FeatureLineStyles.Item(val)
                                                            NameObject = val
                                                        Catch ex As System.ArgumentException
                                                            MsgBox("Не удалось скопировать стиль Характерной линии из шаблона!!! Макро прервано.")
                                                            Exit Sub
                                                        End Try
                                                    ElseIf typeZn Like "ТочкаCOGO" Then
                                                        'копируем ключ-описатель
                                                        Dim pdKeySetIds As ObjectIdCollection = Autodesk.Civil.DatabaseServices.PointDescriptionKeySetCollection.GetPointDescriptionKeySets(acCurDb).SearchOrder
                                                        Dim boolKey As Boolean = False
                                                        If pdKeySetIds.Count > 0 Then
                                                            For Each idKeySet As ObjectId In pdKeySetIds
                                                                Dim DeskKeySet As Autodesk.Civil.DatabaseServices.PointDescriptionKeySet = acTrans.GetObject(idKeySet, OpenMode.ForRead)
                                                                Try
                                                                    IDNameObject = DeskKeySet.Item(val)
                                                                    NameObject = val
                                                                    Exit For
                                                                Catch ex As System.ArgumentException
                                                                End Try
                                                            Next
                                                            If IsNothing(IDNameObject) = True Then
                                                                MsgBox("Не удалось скопировать ключ-описатель из шаблона!!! Макро прервано.")
                                                                Exit Sub
                                                            End If
                                                            If IDNameObject.IsNull = True Then
                                                                MsgBox("Не удалось скопировать ключ-описатель из шаблона!!! Макро прервано.")
                                                                Exit Sub
                                                            End If
                                                        End If
                                                    ElseIf typeZn Like "NoteLineLabel" Then
                                                        'копируем стиль метки линии
                                                        Try
                                                            IDNameObject = CivilDoc.Styles.LabelStyles.GeneralLineLabelStyles.Item(val)
                                                            NameObject = val
                                                        Catch ex As System.ArgumentException
                                                            MsgBox("Не удалось скопировать стиль Примечания из шаблона!!! Макро прервано.")
                                                            Exit Sub
                                                        End Try
                                                        '======================================================================================================================
                                                        'копируем слой
                                                    End If
                                                ElseIf dataField Like "Description" Then
                                                    DeskObject = val
                                                ElseIf dataField Like "Layer" Then
                                                    Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                                                    If acLyrTbl.Has(val) = True Then
                                                        IDNameLayer = acLyrTbl.Item(val)
                                                        NameLayer = val
                                                    Else
                                                        NameLayer = "0"
                                                        IDNameLayer = acLyrTbl.Item("0")
                                                        NameLayerIsh = val
                                                    End If
                                                    '=============================================================================================================
                                                    '7.копируем стиль текста
                                                ElseIf dataField Like "TextStyle" Then
                                                    idTextStyle = AcadFunction.FuncReturnIDTextStylebyName(val)
                                                    If IsNothing(idTextStyle) = False Then
                                                        If idTextStyle.IsNull = True Then
                                                            TextStyle = val
                                                        Else
                                                            idTextStyle = acCurDb.Textstyle
                                                        End If
                                                    End If
                                                    '======================================================================================================================
                                                    'копируем таблицу PS
                                                ElseIf dataField Like "PS" Then
                                                    Dim dictPropSetDef As Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(acCurDb)
                                                    If dictPropSetDef.Has(val, acTrans) = True Then
                                                        idPsTable = dictPropSetDef.GetAt(val)
                                                        PSTable = val
                                                        Dim boolReadFields As Boolean = FuncPropertySet.FuncReadAllValueFieldSetTables(PSTable, ArrayODRec)
                                                    Else
                                                        PSTable = ""
                                                    End If
                                                    '======================================================================================================================
                                                    'копируем группу точек
                                                ElseIf dataField Like "PointGroup" Then
                                                    Dim idGroupPoint As ObjectId = Nothing
                                                    Try
                                                        idPointGroup = CivilDoc.PointGroups.Item(val)
                                                        PointGroup = val
                                                    Catch ex As System.ArgumentException
                                                        PointGroup = ""
                                                    End Try
                                                    '======================================================================================================================
                                                    'копируем площадку
                                                ElseIf dataField Like "Site" Then
                                                    Dim idcolsite As ObjectIdCollection = CivilDoc.GetSiteIds
                                                    If idcolsite.Count > 0 Then
                                                        For j As Integer = 0 To idcolsite.Count - 1
                                                            Dim SiteObj As Autodesk.Civil.DatabaseServices.Site = acTrans.GetObject(idcolsite.Item(j), OpenMode.ForRead)
                                                            If SiteObj.Name Like val Then
                                                                IdSite = SiteObj.Id
                                                                nameSite = val
                                                                Exit For
                                                            End If
                                                        Next j
                                                    Else
                                                        nameSite = ""
                                                    End If
                                                    '======================================================================================================================
                                                    'копируем стиль маркера
                                                ElseIf dataField Like "MarkerStyle" Then
                                                    Dim idMarkerStyle As ObjectId = Nothing
                                                    Try
                                                        idStyleMarker = CivilDoc.Styles.MarkerStyles.Item(val)
                                                        styleMarker = val
                                                    Catch ex As System.ArgumentException
                                                    End Try
                                                    '========================================================================================================
                                                    'масштаб 
                                                ElseIf dataField Like "LineScale" Then
                                                    LineScale = CDbl(val)
                                                    If LineScale Then LineScale = Math.Abs(LineScale)
                                                    If LineScale = 0 Then LineScale = 1
                                                    Dim PLTSCAle As Integer = Application.GetSystemVariable("PSLTSCALE") 'маштаб автоматически подстраивается под видовой экран
                                                    Dim LTSCAle As Double = Application.GetSystemVariable("LTSCALE") 'коэффициент глобального масштаба
                                                    LineScale = LineScale / LTSCAle 'делаем поправку за глобальный масштаб
                                                    If PLTSCAle = 0 Then 'если опция одинакового показа линии для любого масштаба отключена
                                                        Dim ScaleLocal As Double = M
                                                        LineScale = LineScale * ScaleLocal
                                                    End If
                                                    '==========================================================================================================
                                                    '4. цвет линии
                                                ElseIf dataField Like "Color" Then
                                                    If val Like "ПоСлою" Then
                                                        IndColor = 256
                                                    ElseIf val Like "ПоБлоку" Then
                                                        IndColor = 0
                                                    ElseIf val Like "ByLayer" Then
                                                        IndColor = 256
                                                    ElseIf val Like "ByBlock" Then
                                                        IndColor = 0
                                                    Else
                                                        If IsNumeric(val) = True Then
                                                            Dim indexColor As Integer = CInt(val)
                                                            If indexColor < 1 Then
                                                                IndColor = 0
                                                            ElseIf indexColor > 255 Then
                                                                IndColor = 256
                                                            Else
                                                                IndColor = indexColor
                                                            End If
                                                        End If
                                                    End If
                                                ElseIf dataField Like "LineWidth" Then
                                                    '=========================================================================================================
                                                    '5. толщина линии
                                                    If val Like "ПоСлою" Then
                                                        LineWidth = LineWeight.ByLayer
                                                    ElseIf val Like "ПоБлоку" Then
                                                        LineWidth = LineWeight.ByBlock
                                                    ElseIf val Like "ByLayer" Then
                                                        LineWidth = LineWeight.ByLayer
                                                    ElseIf val Like "ByBlock" Then
                                                        LineWidth = LineWeight.ByBlock
                                                    Else
                                                        Try
                                                            LineWidth = CInt(val)
                                                        Catch ex As ArgumentException
                                                            LineWidth = 20
                                                        End Try
                                                    End If
                                                ElseIf dataField Like "Note" Then
                                                    Note = val
                                                End If
                                            End If
                                        End If
                                    Next i
                                End If
                                '=========================================================================================================
                                'ищем поверхность
                                Dim collSurf As ObjectIdCollection = CivilDoc.GetSurfaceIds
                                acSurf = Nothing
                                If collSurf.Count > 0 Then
                                    For Each idSurf As ObjectId In collSurf
                                        Dim acSurfTemp As TinSurface = acTrans.GetObject(idSurf, OpenMode.ForRead)
                                        If acSurfTemp.Name Like UserProperties.ComboSurface.Text Then
                                            acSurf = acSurfTemp
                                            Exit For
                                        End If
                                    Next
                                End If
                                '============================================================================================================
                                'ищем трассу
                                Dim collAlign As ObjectIdCollection = CivilDoc.GetAlignmentIds
                                acAlign = Nothing
                                If collAlign.Count > 0 Then
                                    For Each idAlign As ObjectId In collAlign
                                        Dim acAlignTemp As Alignment = acTrans.GetObject(idAlign, OpenMode.ForRead)
                                        If acAlignTemp.Name Like UserProperties.ComboAlign.Text Then
                                            acAlign = acAlignTemp
                                            Exit For
                                        End If
                                    Next
                                End If
                                acTrans.Commit()
                            End Using
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'начинаем выбирать полилинии
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            If typeZn Like "Полилиния" Or typeZn Like "Полигон" Then
                                Do
                                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                                        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                                        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                        '===========================================================================================================================
                                        Dim acSSPrompt As PromptEntityResult = ed1.GetEntity("^C^C " & vbLf & "Выберите полилинию: ")
                                        If Not (acSSPrompt.Status = PromptStatus.OK) Then
                                            Exit Do
                                        End If
                                        Dim acObjIdColl As ObjectIdCollection = New ObjectIdCollection()
                                        Try
                                            Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForWrite)
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'считываем поля из PS Таблицы и заносим данные
                                            If IsArray(ArrayODRec) = True Then
                                                For i As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                    Dim TempField As String = ArrayODRec(1, i)
                                                    If TempField.Trim Like "LOCALID" Then
                                                        Dim rng As Random = New Random
                                                        ArrayODRec(2, i) = rng.Next
                                                    ElseIf TempField.Trim Like "LABEL" Then
                                                        ArrayODRec(2, i) = tableZn
                                                    ElseIf TempField.Trim Like "PURPOSE" Then
                                                        ArrayODRec(2, i) = deskZn
                                                    ElseIf TempField.Trim Like "NAME" Then
                                                        ArrayODRec(2, i) = deskZn
                                                    End If
                                                Next i
                                                'считываем прежние характеристики
                                                Dim ArrayPSTableObj As String(,) = Nothing
                                                Dim boolArrayTablePS As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(ent, ArrayPSTableObj)
                                                If IsArray(ArrayPSTableObj) = True Then
                                                    For i As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                        Dim TempField1 As String = ArrayODRec(1, i)
                                                        For j As Integer = 0 To ArrayPSTableObj.GetUpperBound(1)
                                                            Dim TempField2 As String = ArrayPSTableObj(1, j)
                                                            If TempField1 Like TempField2 Then
                                                                If Not (TempField1 Like "NAME" Or TempField1 Like "PURPOSE") Then
                                                                    If ArrayPSTableObj(2, j).Trim.Length > 0 Then
                                                                        ArrayODRec(2, i) = ArrayPSTableObj(2, j)
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    Next
                                                End If
                                                'записываем данные с поля NOTE
                                                Dim ArrayNote As String() = Note.Split(";")
                                                If IsArray(ArrayNote) = True Then
                                                    For i As Integer = 0 To ArrayNote.Length - 1
                                                        Dim str As String = ArrayNote(i)
                                                        Dim pos As Integer = str.IndexOf("=")
                                                        If pos > 0 Then
                                                            Dim NameField As String = Mid(str, 1, pos)
                                                            Dim valStr As String = Mid(str, pos + 2)
                                                            For j As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                                Dim TempField1 As String = ArrayODRec(1, j)
                                                                If TempField1.Trim Like NameField.Trim Then
                                                                    ArrayODRec(2, j) = valStr
                                                                End If
                                                            Next
                                                            'если есть вмещение в знаке
                                                            If NameField.Trim Like "2Offset" Then
                                                                offset = Val(valStr.Trim)
                                                                If IsArray(ArrayODRec) = True Then
                                                                    For i1 As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                                        Dim TempField As String = ArrayODRec(1, i1)
                                                                        If TempField.Trim Like "KeyParameter" Then
                                                                            If ArrayODRec(2, i1).Trim.Length = 0 Then
                                                                                ArrayODRec(2, i1) = offset
                                                                            End If
                                                                        End If
                                                                    Next i1
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                                '============================================================================================
                                                'записываем даные из даблицы
                                                For i As Integer = 0 To DataGrid.Items.Count - 1
                                                    If TypeOf DataGrid.Items.Item(i) Is Customer Then
                                                        Dim cust As Customer = DataGrid.Items.Item(i)
                                                        Dim nameF As String = cust.Имя_свойства
                                                        Dim valN As String = cust.Значение
                                                        If IsNothing(valN) = False Then
                                                            If valN.Trim.Length > 0 Then
                                                                For j As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                                    If ArrayODRec(1, j) Like nameF Then
                                                                        ArrayODRec(2, j) = valN
                                                                        Exit For
                                                                    End If
                                                                Next j
                                                            End If
                                                        End If
                                                    End If
                                                Next i
                                                'удаляем лишние таблицы
                                                Dim ArrayTables As String() = Nothing
                                                Dim boolFindTables As Boolean = FuncPropertySet.FuncReadAllDataSetTables(ent, ArrayTables)
                                                If IsArray(ArrayTables) = True Then
                                                    Dim boolDelPS As Boolean = FuncPropertySet.FuncDeleteDataSetTables1(ent, ArrayTables)
                                                End If
                                            End If
                                            '=======================================================================================================================================================
                                            'оставить полилинию
                                            If UserProperties.CheboxFeatureLine.IsChecked = False Then
                                                If TypeOf ent Is Line Or TypeOf ent Is Polyline Or TypeOf ent Is Polyline2d Or TypeOf ent Is Spline Or TypeOf ent Is Arc Or TypeOf ent Is Circle Then
                                                    'присваиваем тип линии
                                                    ent.Linetype = NameObject
                                                    'присваиваем слой
                                                    ent.Layer = NameLayer
                                                    'присваиваем цвет
                                                    ent.ColorIndex = IndColor
                                                    'присваиваем масштаб
                                                    ent.LinetypeScale = LineScale
                                                    'присваиваем вес линии
                                                    ent.LineWeight = LineWidth
                                                    '======================================================================================================================================
                                                    'далее все индивидуально
                                                    If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                                        Dim Poly As Autodesk.AutoCAD.DatabaseServices.Polyline = ent
                                                        If GScale >= 0 Then
                                                            Poly.ConstantWidth = GScale
                                                        End If
                                                        'смещение линии
                                                        If offset > 0 Then
                                                            Dim CollDb As DBObjectCollection = Poly.GetOffsetCurves(offset)
                                                            For Each acEnt As Autodesk.AutoCAD.DatabaseServices.Entity In CollDb
                                                                acBlkTblRec.AppendEntity(acEnt)
                                                                acTrans.AddNewlyCreatedDBObject(acEnt, True)
                                                                acEnt.Linetype = "Continuous"
                                                                acEnt.LineWeight = LineWeight.LineWeight020
                                                                acEnt.LinetypeScale = 1
                                                                acEnt.ColorIndex = IndColor
                                                                acEnt.Layer = NameLayer
                                                            Next
                                                            Dim CollDb2 As DBObjectCollection = Poly.GetOffsetCurves(-1 * offset)
                                                            For Each acEnt As Autodesk.AutoCAD.DatabaseServices.Entity In CollDb2
                                                                acBlkTblRec.AppendEntity(acEnt)
                                                                acTrans.AddNewlyCreatedDBObject(acEnt, True)
                                                                acEnt.Linetype = "Continuous"
                                                                acEnt.LineWeight = LineWeight.LineWeight020
                                                                acEnt.LinetypeScale = 1
                                                                acEnt.ColorIndex = IndColor
                                                                acEnt.Layer = NameLayer
                                                            Next
                                                        End If
                                                    ElseIf TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
                                                        Dim Poly As Autodesk.AutoCAD.DatabaseServices.Polyline2d = ent
                                                        If GScale >= 0 Then
                                                            Poly.ConstantWidth = GScale
                                                        End If
                                                        'смещение линии
                                                        If offset > 0 Then
                                                            Dim CollDb As DBObjectCollection = Poly.GetOffsetCurves(offset)
                                                            For Each acEnt As Autodesk.AutoCAD.DatabaseServices.Entity In CollDb
                                                                acBlkTblRec.AppendEntity(acEnt)
                                                                acTrans.AddNewlyCreatedDBObject(acEnt, True)
                                                                acEnt.Linetype = "Continuous"
                                                                acEnt.LineWeight = LineWeight.LineWeight020
                                                                acEnt.LinetypeScale = 1
                                                                acEnt.ColorIndex = IndColor
                                                                acEnt.Layer = NameLayer
                                                            Next
                                                            Dim CollDb2 As DBObjectCollection = Poly.GetOffsetCurves(-1 * offset)
                                                            For Each acEnt As Autodesk.AutoCAD.DatabaseServices.Entity In CollDb2
                                                                acBlkTblRec.AppendEntity(acEnt)
                                                                acTrans.AddNewlyCreatedDBObject(acEnt, True)
                                                                acEnt.Linetype = "Continuous"
                                                                acEnt.LineWeight = LineWeight.LineWeight020
                                                                acEnt.LinetypeScale = 1
                                                                acEnt.ColorIndex = IndColor
                                                                acEnt.Layer = NameLayer
                                                            Next
                                                        End If
                                                    End If
                                                    acDoc.Editor.Regen()
                                                    'записываем новые  данные
                                                    If IsNothing(idPsTable) = False Then
                                                        If idPsTable.IsNull = False Then
                                                            Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(ent, idPsTable)
                                                            Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(ent, ArrayODRec)
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                '==============================================================================================================
                                                'создаем характерную линию
                                                '===============================================================================================
                                                'создаем стиль для характерной линии
                                                If CivilDoc.Styles.FeatureLineStyles.Contains(NameObject) = False Then
                                                    Dim PropArrayFeatureLine As String(,) = Nothing
                                                    ReDim PropArrayFeatureLine(1, 5)
                                                    PropArrayFeatureLine(0, 0) = "Описание объекта"
                                                    PropArrayFeatureLine(1, 0) = deskZn
                                                    PropArrayFeatureLine(0, 1) = "Имя объекта"
                                                    PropArrayFeatureLine(1, 1) = NameObject
                                                    PropArrayFeatureLine(0, 2) = "Наименование слоя"
                                                    PropArrayFeatureLine(1, 2) = NameLayer
                                                    PropArrayFeatureLine(0, 3) = "Цвет"
                                                    PropArrayFeatureLine(1, 3) = IndColor
                                                    PropArrayFeatureLine(0, 4) = "Вес линии"
                                                    PropArrayFeatureLine(1, 4) = LineWidth
                                                    PropArrayFeatureLine(0, 5) = "Масштаб"
                                                    PropArrayFeatureLine(1, 5) = LineScale
                                                    Dim StyleFeatureLine As Styles.FeatureLineStyle = FuncFeatureLine.FuncCreateFeatureLineStyleByArrayProperties(PropArrayFeatureLine)
                                                    If IsNothing(StyleFeatureLine) = False Then
                                                        If StyleFeatureLine.Id.IsNull = False Then
                                                            IDNameObject = StyleFeatureLine.Id
                                                        Else
                                                            MsgBox("Не удалось создать стиль Характерной линии!!! Макро превано.")
                                                            Exit Do
                                                        End If
                                                    Else
                                                        MsgBox("Не удалось создать стиль Характерной линии!!! Макро превано.")
                                                        Exit Do
                                                    End If
                                                Else
                                                    IDNameObject = CivilDoc.Styles.FeatureLineStyles.Item(NameObject)
                                                End If
                                                '======================================================================================================
                                                'создаем площадку
                                                Dim idSitesCollection As ObjectIdCollection = CivilDoc.GetSiteIds()
                                                If idSitesCollection.Count > 0 Then
                                                    For Each idSiteTemp As ObjectId In idSitesCollection
                                                        Dim acSite As Site = acTrans.GetObject(idSiteTemp, OpenMode.ForRead)
                                                        If acSite.Name.Trim Like NameLayer Then
                                                            IdSite = acSite.Id
                                                        End If
                                                    Next
                                                End If
                                                If IsNothing(IdSite) = True Then
                                                    Try
                                                        IdSite = Autodesk.Civil.DatabaseServices.Site.Create(CivilDoc, NameLayer)
                                                    Catch ex As System.ArgumentException
                                                    End Try
                                                End If
                                                If IdSite.IsNull = True Then
                                                    Try
                                                        IdSite = Autodesk.Civil.DatabaseServices.Site.Create(CivilDoc, NameLayer)
                                                    Catch ex As System.ArgumentException
                                                    End Try
                                                End If
                                                '=============================================================================================================
                                                'если это уже характерная линия
                                                If TypeOf ent Is FeatureLine Then
                                                    Dim featLine As FeatureLine = ent
                                                    featLine.StyleId = IDNameObject
                                                    featLine.Layer = NameLayer
                                                    '=============================================================================================================
                                                    'переносим на площадку
                                                    If IsNothing(IdSite) = False Then
                                                        If IdSite.IsNull = False Then
                                                            Try
                                                                FeatureLine.MoveToSite(ent.Id, IdSite)
                                                            Catch ex As System.ArgumentException
                                                            End Try
                                                        End If
                                                    End If
                                                    acDoc.Editor.Regen()
                                                    'записываем новые  данные
                                                    If IsNothing(idPsTable) = False Then
                                                        If idPsTable.IsNull = False Then
                                                            Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(ent, idPsTable)
                                                            Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(ent, ArrayODRec)
                                                        End If
                                                    End If
                                                Else
                                                    'преобразовываем полилинию
                                                    If TypeOf ent Is Line Or TypeOf ent Is Polyline Or TypeOf ent Is Polyline2d Or TypeOf ent Is Arc Then
                                                        Dim IDFeatureLine As ObjectId = Nothing
                                                        Dim FeatObject As FeatureLine = Nothing
                                                        If TypeOf ent Is Polyline2d Then
                                                            Dim Pline2d As Polyline2d = ent
                                                            Dim PolyTypeUser As Integer = Pline2d.PolyType
                                                            If PolyTypeUser > 0 Then
                                                                Pline2d.ConvertToPolyType(Poly2dType.SimplePoly)
                                                            End If
                                                        End If
                                                        Try
                                                            If IsNothing(IdSite) = False Then
                                                                If IdSite.IsNull = False Then
                                                                    IDFeatureLine = Autodesk.Civil.DatabaseServices.FeatureLine.Create("", ent.Id, IdSite)
                                                                Else
                                                                    IDFeatureLine = Autodesk.Civil.DatabaseServices.FeatureLine.Create("", ent.Id)
                                                                End If
                                                            Else
                                                                IDFeatureLine = Autodesk.Civil.DatabaseServices.FeatureLine.Create("", ent.Id)
                                                            End If
                                                        Catch ex As ArgumentException
                                                            MsgBox("Невозможно преобразовать объект в Характерную линию!!!")
                                                            Continue Do
                                                        End Try
                                                        '=================================================================================================
                                                        'заносим в нужный слой
                                                        If IsNothing(IDFeatureLine) = False Then
                                                            If IDFeatureLine.IsNull = False Then
                                                                FeatObject = acTrans.GetObject(IDFeatureLine, OpenMode.ForWrite)
                                                                FeatObject.Layer = NameLayer
                                                                If IsNothing(IDNameObject) = False Then
                                                                    If IDNameObject.IsNull = False Then
                                                                        FeatObject.StyleId = IDNameObject
                                                                    End If
                                                                End If
                                                                FeatObject.Description = deskZn
                                                                '===================================================================================================
                                                                'делаем интерполяцию высот
                                                                '============================================================================================================
                                                                'ищем cogo точки вдоль характерной линии
                                                                Dim ArrayPoint As Integer() = Nothing
                                                                'создаем габаритный контейнер полилинии для фильтрации точек
                                                                Dim MinCoord As Point3d = FeatObject.GeometricExtents.MinPoint
                                                                Dim MaxCoord As Point3d = FeatObject.GeometricExtents.MaxPoint
                                                                Dim Count1 As Integer = 0
                                                                Dim CountPoint As Integer = CivilDoc.CogoPoints.Count
                                                                If CountPoint > 0 Then
                                                                    For Each PointId As ObjectId In CivilDoc.CogoPoints
                                                                        Dim PointCOGO1 As CogoPoint = PointId.GetObject(OpenMode.ForRead)
                                                                        If PointCOGO1.Northing >= MinCoord.Y - 10 And PointCOGO1.Easting >= MinCoord.X - 10 Then
                                                                            If PointCOGO1.Northing <= MaxCoord.Y + 10 And PointCOGO1.Easting <= MaxCoord.X + 10 Then
                                                                                ReDim Preserve ArrayPoint(Count1)
                                                                                ArrayPoint(Count1) = PointCOGO1.PointNumber
                                                                                Count1 += 1
                                                                            End If
                                                                        End If
                                                                    Next
                                                                End If
                                                                '==========================================================================================================
                                                                'забираем узлы характерной линии в массив
                                                                Dim PointColl As Point3dCollection = New Point3dCollection
                                                                Dim ftype As Autodesk.Civil.FeatureLinePointType = Autodesk.Civil.FeatureLinePointType.PIPoint
                                                                PointColl = FeatObject.GetPoints(ftype)
                                                                Dim ArrayPointFeatureLine As Double(,) = Nothing
                                                                Dim CountArrayPointFeatureLine As Integer = 0
                                                                Dim boolFindCogo As Boolean = False
                                                                If PointColl.Count > 1 Then
                                                                    If IsArray(ArrayPoint) = True Then
                                                                        For Each pt3d As Point3d In PointColl
                                                                            Dim PointCogoFind As CogoPoint = FuncSurvey.FuncFindCogoPointByCoordinates(pt3d, ArrayPoint, 0.05)
                                                                            ReDim Preserve ArrayPointFeatureLine(2, CountArrayPointFeatureLine)
                                                                            ArrayPointFeatureLine(0, CountArrayPointFeatureLine) = pt3d.X
                                                                            ArrayPointFeatureLine(1, CountArrayPointFeatureLine) = pt3d.Y
                                                                            If IsNothing(PointCogoFind) = False Then
                                                                                If PointCogoFind.Id.IsNull = False Then
                                                                                    If PointCogoFind.IsPlanar = False Then
                                                                                        ArrayPointFeatureLine(2, CountArrayPointFeatureLine) = PointCogoFind.Elevation
                                                                                        boolFindCogo = True
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                            CountArrayPointFeatureLine += 1
                                                                        Next
                                                                    End If
                                                                End If
                                                                '================================================================================================
                                                                'назначаем высоты первой и последней точкам
                                                                If IsArray(ArrayPointFeatureLine) = True And boolFindCogo = True Then
                                                                    For i As Integer = 0 To ArrayPointFeatureLine.GetUpperBound(1)
                                                                        Dim TempH As Double = ArrayPointFeatureLine(2, i)
                                                                        If TempH <> 0 Then
                                                                            ArrayPointFeatureLine(2, 0) = TempH
                                                                            Exit For
                                                                        End If
                                                                    Next i
                                                                    For i As Integer = ArrayPointFeatureLine.GetUpperBound(1) To 0 Step -1
                                                                        Dim TempH As Double = ArrayPointFeatureLine(2, i)
                                                                        If TempH <> 0 Then
                                                                            ArrayPointFeatureLine(2, ArrayPointFeatureLine.GetUpperBound(1)) = TempH
                                                                            Exit For
                                                                        End If
                                                                    Next i
                                                                    '=====================================================================================
                                                                    'интерполируем высоты на остальные точки
                                                                    Dim startElevPoint As Double = 0
                                                                    Dim endElevPoint As Double = 0
                                                                    Dim dist As Double = 0
                                                                    For i As Integer = 1 To ArrayPointFeatureLine.GetUpperBound(1) - 1
                                                                        Dim fPoint As Double = ArrayPointFeatureLine(2, i)
                                                                        If fPoint = 0 Then
                                                                            'смотрим предыдущий пикет
                                                                            startElevPoint = ArrayPointFeatureLine(2, i - 1)
                                                                            Dim pt1 As Point2d = New Point2d(ArrayPointFeatureLine(0, i - 1), ArrayPointFeatureLine(1, i - 1))
                                                                            Dim pt2 As Point2d = New Point2d(ArrayPointFeatureLine(0, i), ArrayPointFeatureLine(1, i))
                                                                            dist = pt1.GetVectorTo(pt2).Length
                                                                            For j As Integer = i + 1 To ArrayPointFeatureLine.GetUpperBound(1)
                                                                                Dim ftempPoint As Double = ArrayPointFeatureLine(2, j)
                                                                                If ftempPoint = 0 Then
                                                                                    Dim pt3 As Point2d = New Point2d(ArrayPointFeatureLine(0, j - 1), ArrayPointFeatureLine(1, j - 1))
                                                                                    Dim pt4 As Point2d = New Point2d(ArrayPointFeatureLine(0, j), ArrayPointFeatureLine(1, j))
                                                                                    dist = dist + pt3.GetVectorTo(pt4).Length
                                                                                Else
                                                                                    endElevPoint = ArrayPointFeatureLine(2, j)
                                                                                    Dim temppt1 As Point2d = New Point2d(ArrayPointFeatureLine(0, j - 1), ArrayPointFeatureLine(1, j - 1))
                                                                                    Dim temppt2 As Point2d = New Point2d(ArrayPointFeatureLine(0, j), ArrayPointFeatureLine(1, j))
                                                                                    dist = dist + temppt1.GetVectorTo(temppt2).Length
                                                                                    Exit For
                                                                                End If
                                                                            Next j
                                                                            '===========================================
                                                                            'делаем расчет
                                                                            Dim deltaH As Double = endElevPoint - startElevPoint
                                                                            Dim ddist As Double = pt1.GetVectorTo(pt2).Length / dist
                                                                            Dim rez1 As Double = ddist * deltaH
                                                                            ArrayPointFeatureLine(2, i) = startElevPoint + rez1
                                                                        End If
                                                                    Next i
                                                                    '==============================================================================================================
                                                                    'переносим высоты на характерную линию
                                                                    For i As Integer = 0 To ArrayPointFeatureLine.GetUpperBound(1)
                                                                        Dim H As Double = ArrayPointFeatureLine(2, i)
                                                                        Try
                                                                            FeatObject.SetPointElevation(i, H)
                                                                        Catch ex As System.ArgumentOutOfRangeException
                                                                        End Try
                                                                    Next
                                                                    If ent.IsWriteEnabled = True Then
                                                                        ent.Erase()
                                                                    End If
                                                                Else
                                                                    '================================================================================================================
                                                                    'назначаем высоты из поверхности
                                                                    If IsNothing(acSurf) = False Then
                                                                        If acSurf.Id.IsNull = False Then
                                                                            Try
                                                                                FeatObject.AssignElevationsFromSurface(acSurf.Id, False)
                                                                            Catch ex As System.ArgumentException
                                                                            End Try
                                                                        End If
                                                                    End If
                                                                End If
                                                                '==============================================================================================================
                                                                'вставляем расширенные данные
                                                                If IsNothing(idPsTable) = False Then
                                                                    If idPsTable.IsNull = False Then
                                                                        Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(FeatObject, idPsTable)
                                                                        Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(FeatObject, ArrayODRec)
                                                                    End If
                                                                End If
                                                                '====================================================================================================================
                                                                'добавляем как структурную линию
                                                                If UserProperties.CheboxAddSurface.IsChecked = True Then
                                                                    If IsNothing(acSurf) = False Then
                                                                        If acSurf.Id.IsNull = False Then
                                                                            Dim idObjColl As ObjectIdCollection = New ObjectIdCollection
                                                                            idObjColl.Add(FeatObject.Id)
                                                                            Dim opS As SurfaceOperationAddBreakline = acSurf.BreaklinesDefinition.AddStandardBreaklines(idObjColl, 0.2, 50, 1, 0.01745)
                                                                            opS.Description = deskZn
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Catch ex As System.Exception
                                        End Try
                                        acTrans.Commit()
                                    End Using
                                Loop
                                '==============================================================================================================
                                'выбран тип линии - полилиния
                            ElseIf typeZn Like "Мультилиния" Then
                                Do
                                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                                        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                                        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                        Dim acSSPrompt As PromptEntityResult = ed1.GetEntity("^C^C " & vbLf & "Выберите мультилинию или полилинию для преобразования: ")
                                        If Not (acSSPrompt.Status = PromptStatus.OK) Then
                                            acTrans.Commit()
                                            Exit Do
                                        End If
                                        Try
                                            Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForWrite)
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'считываем поля из PS Таблицы и заносим данные
                                            If IsArray(ArrayODRec) = True Then
                                                For i As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                    Dim TempField As String = ArrayODRec(1, i)
                                                    If TempField.Trim Like "LOCALID" Then
                                                        Dim rng As Random = New Random
                                                        ArrayODRec(2, i) = rng.Next
                                                    ElseIf TempField.Trim Like "LABEL" Then
                                                        ArrayODRec(2, i) = tableZn
                                                    ElseIf TempField.Trim Like "PURPOSE" Then
                                                        ArrayODRec(2, i) = deskZn
                                                    ElseIf TempField.Trim Like "NAME" Then
                                                        ArrayODRec(2, i) = deskZn
                                                    End If
                                                Next i
                                                'считываем прежние характеристики
                                                Dim ArrayPSTableObj As String(,) = Nothing
                                                Dim boolArrayTablePS As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(ent, ArrayPSTableObj)
                                                If IsArray(ArrayPSTableObj) = True Then
                                                    For i As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                        Dim TempField1 As String = ArrayODRec(1, i)
                                                        For j As Integer = 0 To ArrayPSTableObj.GetUpperBound(1)
                                                            Dim TempField2 As String = ArrayPSTableObj(1, j)
                                                            If TempField1 Like TempField2 Then
                                                                If Not (TempField1 Like "NAME" Or TempField1 Like "PURPOSE") Then
                                                                    If ArrayPSTableObj(2, j).Trim.Length > 0 Then
                                                                        ArrayODRec(2, i) = ArrayPSTableObj(2, j)
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    Next
                                                End If
                                                'записываем данные с поля NOTE
                                                Dim ArrayNote As String() = Note.Split(";")
                                                If IsArray(ArrayNote) = True Then
                                                    For i As Integer = 0 To ArrayNote.Length - 1
                                                        Dim str As String = ArrayNote(i)
                                                        Dim pos As Integer = str.IndexOf("=")
                                                        If pos > 0 Then
                                                            Dim NameField As String = Mid(str, 1, pos)
                                                            Dim valStr As String = Mid(str, pos + 2)
                                                            For j As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                                Dim TempField1 As String = ArrayODRec(1, j)
                                                                If TempField1.Trim Like NameField.Trim Then
                                                                    ArrayODRec(2, j) = valStr
                                                                End If
                                                            Next
                                                        End If
                                                    Next
                                                End If
                                                'удаляем лишние таблицы
                                                Dim ArrayTables As String() = Nothing
                                                Dim boolFindTables As Boolean = FuncPropertySet.FuncReadAllDataSetTables(ent, ArrayTables)
                                                If IsArray(ArrayTables) = True Then
                                                    Dim boolDelPS As Boolean = FuncPropertySet.FuncDeleteDataSetTables1(ent, ArrayTables)
                                                End If
                                            End If

                                            Dim ArrayVertex As Double(,) = Nothing
                                            Dim CountArrayVertex As Integer = 0
                                            Dim acMline As Mline = Nothing
                                            If TypeOf ent Is Line Or TypeOf ent Is Polyline Or TypeOf ent Is Polyline2d Then
                                                Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(acCurDb.LinetypeTableId, OpenMode.ForRead)
                                                If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                                    Dim Poly As Autodesk.AutoCAD.DatabaseServices.Polyline = ent
                                                    For i As Integer = 0 To Poly.NumberOfVertices - 1
                                                        ReDim Preserve ArrayVertex(1, CountArrayVertex)
                                                        ArrayVertex(0, CountArrayVertex) = Poly.GetPoint2dAt(i).X
                                                        ArrayVertex(1, CountArrayVertex) = Poly.GetPoint2dAt(i).Y
                                                        CountArrayVertex += 1
                                                    Next i
                                                    If Poly.Closed = True Then
                                                        ReDim Preserve ArrayVertex(1, CountArrayVertex)
                                                        ArrayVertex(0, CountArrayVertex) = Poly.StartPoint.X
                                                        ArrayVertex(1, CountArrayVertex) = Poly.StartPoint.Y
                                                    End If
                                                ElseIf TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
                                                    Dim Poly As Autodesk.AutoCAD.DatabaseServices.Polyline2d = ent
                                                    Try
                                                        For Each vid As ObjectId In Poly
                                                            Dim v2d As Autodesk.AutoCAD.DatabaseServices.Vertex2d = acTrans.GetObject(vid, OpenMode.ForRead)
                                                            ReDim Preserve ArrayVertex(1, CountArrayVertex)
                                                            ArrayVertex(0, CountArrayVertex) = v2d.Position.X
                                                            ArrayVertex(1, CountArrayVertex) = v2d.Position.Y
                                                            CountArrayVertex += 1
                                                        Next
                                                    Catch ex As Exception
                                                    End Try
                                                    If Poly.Closed = True Then
                                                        ReDim Preserve ArrayVertex(1, CountArrayVertex)
                                                        ArrayVertex(0, CountArrayVertex) = Poly.StartPoint.X
                                                        ArrayVertex(1, CountArrayVertex) = Poly.StartPoint.Y
                                                    End If
                                                ElseIf TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                                    Dim Poly As Autodesk.AutoCAD.DatabaseServices.Line = ent
                                                    ReDim Preserve ArrayVertex(1, CountArrayVertex)
                                                    ArrayVertex(0, CountArrayVertex) = Poly.StartPoint.X
                                                    ArrayVertex(1, CountArrayVertex) = Poly.StartPoint.Y
                                                    CountArrayVertex += 1
                                                    ReDim Preserve ArrayVertex(1, CountArrayVertex)
                                                    ArrayVertex(0, CountArrayVertex) = Poly.EndPoint.X
                                                    ArrayVertex(1, CountArrayVertex) = Poly.EndPoint.Y
                                                End If
                                                '===============================================================
                                                Dim IDMlLine As ObjectId = AcadFunction.FuncDrawMLine(ArrayVertex, NameLayer, IndColor, LineWidth, LineScale, NameObject)
                                                acMline = acTrans.GetObject(IDMlLine, OpenMode.ForWrite)
                                            ElseIf TypeOf ent Is Mline Then
                                                acMline = ent
                                                If acMline.IsWriteEnabled = False Then
                                                    acMline.UpgradeOpen()
                                                End If
                                                '==================================================
                                                'цвет
                                                acMline.ColorIndex = IndColor
                                                '==================================================
                                                'слой
                                                acMline.Layer = NameLayer
                                                '=================================================
                                                'масштаб
                                                acMline.Scale = LineScale
                                                'стиль мультилинии
                                                acMline.Style = IDNameObject
                                                '================================================
                                                'вес линиии
                                                acMline.LineWeight = LineWidth
                                            End If
                                            acDoc.Editor.Regen()
                                            'записываем новые  данные
                                            If IsNothing(idPsTable) = False Then
                                                If idPsTable.IsNull = False Then
                                                    Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(acMline, idPsTable)
                                                    Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acMline, ArrayODRec)
                                                End If
                                            End If
                                            ent.Erase()
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        End Try
                                        acTrans.Commit()
                                    End Using
                                Loop
                                '=========================================================================
                                'вставка блока
                            ElseIf typeZn Like "Блок" Then
                                Do
                                    Try
                                        Dim acSSPrompt As PromptPointResult = ed1.GetPoint("^C^C " & vbLf & "Укажите точку вставки блока: ")
                                        If Not (acSSPrompt.Status = PromptStatus.OK) Then
                                            Exit Do
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'считываем поля из PS Таблицы и заносим данные
                                        If IsArray(ArrayODRec) = True Then
                                            For i As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                Dim TempField As String = ArrayODRec(1, i)
                                                If TempField.Trim Like "LOCALID" Then
                                                    Dim rng As Random = New Random
                                                    ArrayODRec(2, i) = rng.Next
                                                ElseIf TempField.Trim Like "LABEL" Then
                                                    ArrayODRec(2, i) = tableZn
                                                ElseIf TempField.Trim Like "PURPOSE" Then
                                                    ArrayODRec(2, i) = deskZn
                                                ElseIf TempField.Trim Like "NAME" Then
                                                    ArrayODRec(2, i) = deskZn
                                                End If
                                            Next i
                                            'записываем данные с поля NOTE
                                            Dim ArrayNote As String() = Note.Split(";")
                                            If IsArray(ArrayNote) = True Then
                                                For i As Integer = 0 To ArrayNote.Length - 1
                                                    Dim str As String = ArrayNote(i)
                                                    Dim pos As Integer = str.IndexOf("=")
                                                    If pos > 0 Then
                                                        Dim NameField As String = Mid(str, 1, pos)
                                                        Dim valStr As String = Mid(str, pos + 2)
                                                        For j As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                            Dim TempField1 As String = ArrayODRec(1, j)
                                                            If TempField1.Trim Like NameField.Trim Then
                                                                ArrayODRec(2, j) = valStr
                                                            End If
                                                        Next
                                                    End If
                                                Next
                                            End If
                                        End If
                                        'вставляем блок
                                        Dim acBlkObj As BlockReference = AcadFunction.FuncInsertDrawingBlk(acSSPrompt.Value, NameObject, NameLayer, IndColor, Scale * M, LineWidth, 0, 0)
                                        If IsNothing(acBlkObj) = False Then
                                            If acBlkObj.Id.IsNull = False Then
                                                If acBlkObj.IsWriteEnabled = False Then
                                                    acBlkObj.UpgradeOpen()
                                                End If
                                            End If
                                        End If
                                        'записываем новые  данные
                                        If IsNothing(idPsTable) = False Then
                                            If idPsTable.IsNull = False Then
                                                Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(acBlkObj, idPsTable)
                                                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acBlkObj, ArrayODRec)
                                            End If
                                        End If
                                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                    End Try
                                Loop
                                '=====================================================================================================================
                                'точка COGO
                            ElseIf typeZn Like "ТочкаCOGO" Then
                                Do
                                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                                        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                                        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                        Try
                                            Dim acSSPrompt As PromptEntityResult = ed1.GetEntity("^C^C " & vbLf & "Выберите точку COGO: ")
                                            If Not (acSSPrompt.Status = PromptStatus.OK) Then
                                                Exit Do
                                            End If
                                            Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForRead, False) 'открываем объект на чтение
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'считываем поля из PS Таблицы и заносим данные
                                            If IsArray(ArrayODRec) = True Then
                                                For i As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                    Dim TempField As String = ArrayODRec(1, i)
                                                    If TempField.Trim Like "LOCALID" Then
                                                        Dim rng As Random = New Random
                                                        ArrayODRec(2, i) = rng.Next
                                                    ElseIf TempField.Trim Like "LABEL" Then
                                                        ArrayODRec(2, i) = tableZn
                                                    ElseIf TempField.Trim Like "PURPOSE" Then
                                                        ArrayODRec(2, i) = deskZn
                                                    ElseIf TempField.Trim Like "NAME" Then
                                                        ArrayODRec(2, i) = deskZn
                                                    End If
                                                Next i
                                                'считываем прежние характеристики
                                                Dim ArrayPSTableObj As String(,) = Nothing
                                                Dim boolArrayTablePS As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(ent, ArrayPSTableObj)
                                                If IsArray(ArrayPSTableObj) = True Then
                                                    For i As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                        Dim TempField1 As String = ArrayODRec(1, i)
                                                        For j As Integer = 0 To ArrayPSTableObj.GetUpperBound(1)
                                                            Dim TempField2 As String = ArrayPSTableObj(1, j)
                                                            If TempField1 Like TempField2 Then
                                                                If Not (TempField1 Like "NAME" Or TempField1 Like "PURPOSE") Then
                                                                    If ArrayPSTableObj(2, j).Trim.Length > 0 Then
                                                                        ArrayODRec(2, i) = ArrayPSTableObj(2, j)
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    Next
                                                End If
                                                'записываем данные с поля NOTE
                                                Dim ArrayNote As String() = Note.Split(";")
                                                If IsArray(ArrayNote) = True Then
                                                    For i As Integer = 0 To ArrayNote.Length - 1
                                                        Dim str As String = ArrayNote(i)
                                                        Dim pos As Integer = str.IndexOf("=")
                                                        If pos > 0 Then
                                                            Dim NameField As String = Mid(str, 1, pos)
                                                            Dim valStr As String = Mid(str, pos + 2)
                                                            For j As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                                Dim TempField1 As String = ArrayODRec(1, j)
                                                                If TempField1.Trim Like NameField.Trim Then
                                                                    ArrayODRec(2, j) = valStr
                                                                End If
                                                            Next
                                                        End If
                                                    Next
                                                End If
                                                'удаляем лишние таблицы
                                                Dim ArrayTables As String() = Nothing
                                                Dim boolFindTables As Boolean = FuncPropertySet.FuncReadAllDataSetTables(ent, ArrayTables)
                                                If IsArray(ArrayTables) = True Then
                                                    Dim boolDelPS As Boolean = FuncPropertySet.FuncDeleteDataSetTables1(ent, ArrayTables)
                                                End If
                                            End If
                                            If TypeOf ent Is Autodesk.Civil.DatabaseServices.CogoPoint Then
                                                If ent.IsWriteEnabled = False Then
                                                    ent.UpgradeOpen()
                                                End If
                                                Dim Blk As Autodesk.Civil.DatabaseServices.CogoPoint = ent
                                                Blk.RawDescription = NameObject
                                                Blk.ColorIndex = IndColor
                                                'Blk.Linetype = NameTypeLine
                                                Blk.LineWeight = LineWidth
                                                Blk.ScaleXY = Scale
                                                Blk.ApplyDescriptionKeys()

                                                'записываем новые  данные
                                                If IsNothing(idPsTable) = False Then
                                                    If idPsTable.IsNull = False Then
                                                        Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(ent, idPsTable)
                                                        Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(ent, ArrayODRec)
                                                    End If
                                                End If
                                            End If
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        End Try
                                        acTrans.Commit()
                                    End Using
                                Loop
                                '=======================================================================================================================
                                'Характерная линия                            
                            ElseIf typeZn Like "Характерная линия" Then
                                Do
                                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                                        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                                        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                        Try
                                            '===================================================================================================================
                                            Dim acSSPrompt As PromptEntityResult = ed1.GetEntity("^C^C " & vbLf & "Выберите объект: ")
                                            If acSSPrompt.Status = PromptStatus.Cancel Then
                                                Exit Do
                                            End If
                                            Dim obj As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForWrite, True) 'открываем объект на чтение
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'считываем поля из PS Таблицы и заносим данные
                                            If IsArray(ArrayODRec) = True Then
                                                For i As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                    Dim TempField As String = ArrayODRec(1, i)
                                                    If TempField.Trim Like "LOCALID" Then
                                                        Dim rng As Random = New Random
                                                        ArrayODRec(2, i) = rng.Next
                                                    ElseIf TempField.Trim Like "LABEL" Then
                                                        ArrayODRec(2, i) = tableZn
                                                    ElseIf TempField.Trim Like "PURPOSE" Then
                                                        ArrayODRec(2, i) = deskZn
                                                    ElseIf TempField.Trim Like "NAME" Then
                                                        ArrayODRec(2, i) = deskZn
                                                    End If
                                                Next i
                                                'считываем прежние характеристики
                                                Dim ArrayPSTableObj As String(,) = Nothing
                                                Dim boolArrayTablePS As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(obj, ArrayPSTableObj)
                                                If IsArray(ArrayPSTableObj) = True Then
                                                    For i As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                        Dim TempField1 As String = ArrayODRec(1, i)
                                                        For j As Integer = 0 To ArrayPSTableObj.GetUpperBound(1)
                                                            Dim TempField2 As String = ArrayPSTableObj(1, j)
                                                            If TempField1 Like TempField2 Then
                                                                If Not (TempField1 Like "NAME" Or TempField1 Like "PURPOSE") Then
                                                                    If ArrayPSTableObj(2, j).Trim.Length > 0 Then
                                                                        ArrayODRec(2, i) = ArrayPSTableObj(2, j)
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    Next
                                                End If
                                                'записываем данные с поля NOTE
                                                Dim ArrayNote As String() = Note.Split(";")
                                                If IsArray(ArrayNote) = True Then
                                                    For i As Integer = 0 To ArrayNote.Length - 1
                                                        Dim str As String = ArrayNote(i)
                                                        Dim pos As Integer = str.IndexOf("=")
                                                        If pos > 0 Then
                                                            Dim NameField As String = Mid(str, 1, pos)
                                                            Dim valStr As String = Mid(str, pos + 2)
                                                            For j As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                                Dim TempField1 As String = ArrayODRec(1, j)
                                                                If TempField1.Trim Like NameField.Trim Then
                                                                    ArrayODRec(2, j) = valStr
                                                                End If
                                                            Next
                                                        End If
                                                    Next
                                                End If
                                                'удаляем лишние таблицы
                                                Dim ArrayTables As String() = Nothing
                                                Dim boolFindTables As Boolean = FuncPropertySet.FuncReadAllDataSetTables(obj, ArrayTables)
                                                If IsArray(ArrayTables) = True Then
                                                    Dim boolDelPS As Boolean = FuncPropertySet.FuncDeleteDataSetTables1(obj, ArrayTables)
                                                End If
                                            End If
                                            '==============================================================================================================================
                                            'выбрана харакрерная линия
                                            Dim FeatLine As Autodesk.Civil.DatabaseServices.FeatureLine = Nothing
                                            If TypeOf obj Is Autodesk.Civil.DatabaseServices.FeatureLine Then
                                                FeatLine = obj
                                                '=======================================================================================================
                                                'меняем стиль линии
                                                FeatLine.StyleName = NameObject
                                                Dim idTempStyle As ObjectId = Nothing
                                                'Слой меняем
                                                FeatLine.Layer = NameLayer
                                                '=======================================================================================================
                                                'новое описание
                                                FeatLine.Description = deskZn
                                                '========================================================================================================
                                                'переносим на площадку
                                                If IsNothing(IdSite) = False Then
                                                    If IdSite.IsNull = False Then
                                                        Try
                                                            FeatureLine.MoveToSite(FeatLine.Id, IdSite)
                                                        Catch ex As System.ArgumentException
                                                        End Try
                                                    End If
                                                End If
                                                '===========================================================================================================================================================
                                                'Объект автокада
                                            ElseIf TypeOf obj Is Autodesk.AutoCAD.DatabaseServices.Polyline Or TypeOf obj Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Or TypeOf obj Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Or TypeOf obj Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                                Try
                                                    Dim idFeat As ObjectId = Nothing
                                                    If IsNothing(IdSite) = False Then
                                                        If IdSite.IsNull = False Then
                                                            idFeat = Autodesk.Civil.DatabaseServices.FeatureLine.Create("", obj.Id, IdSite)
                                                        Else
                                                            idFeat = Autodesk.Civil.DatabaseServices.FeatureLine.Create("", obj.Id)
                                                        End If
                                                    Else
                                                        idFeat = Autodesk.Civil.DatabaseServices.FeatureLine.Create("", obj.Id)
                                                    End If
                                                    FeatLine = acTrans.GetObject(idFeat, OpenMode.ForWrite)
                                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                    'удаляем лишние отметки
                                                    Dim pt As Point3dCollection = FeatLine.GetPoints(Autodesk.Civil.FeatureLinePointType.ElevationPoint)
                                                    If pt.Count > 0 Then
                                                        For Each pttemp As Point3d In pt
                                                            FeatLine.DeleteElevationPoint(pttemp)
                                                        Next
                                                    End If
                                                    '================================================================================================================
                                                    'вводим описание объекта
                                                    FeatLine.Description = deskZn
                                                    'меняем слой
                                                    FeatLine.Layer = NameLayer
                                                    '===============================================================================================================
                                                    'применяем стиль Характерной линии
                                                    Dim idTempStyle As ObjectId = Nothing
                                                    Try
                                                        idTempStyle = CivilDoc.Styles.FeatureLineStyles.Item(NameObject)
                                                    Catch ex As System.ArgumentException
                                                    End Try
                                                    If IsNothing(idTempStyle) = False Then
                                                        If idTempStyle.IsNull = False Then
                                                            FeatLine.StyleName = NameObject
                                                        End If
                                                    End If
                                                    '====================================================================================================================
                                                    'назначаем отметки характерной линии
                                                    If IsNothing(acSurf) = False Then
                                                        Try
                                                            FeatLine.AssignElevationsFromSurface(acSurf.Id, False)
                                                        Catch ex As System.ArgumentException
                                                        End Try
                                                    End If
                                                    '=======================================================================================================================
                                                    'стираем старый объект
                                                    obj.Erase()
                                                Catch ex As System.ArgumentException
                                                End Try
                                            End If
                                            'записываем новые  данные
                                            If IsNothing(idPsTable) = False Then
                                                If idPsTable.IsNull = False Then
                                                    Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(FeatLine, idPsTable)
                                                    Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(FeatLine, ArrayODRec)
                                                End If
                                            End If
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        End Try
                                        acTrans.Commit()
                                    End Using
                                Loop
                                '=======================================================================================================================
                                'метка на линии
                            ElseIf typeZn Like "NoteLineLabel" Then
                                '=======================================================================================
                                'выбираем объекты
                                Do
                                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                                        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                                        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                        Try
                                            Dim acSSPrompt As PromptEntityResult = ed1.GetEntity("^C^C " & vbLf & "Выберите объект линейного типа: ")
                                            If acSSPrompt.Status = PromptStatus.Cancel Then
                                                acTrans.Commit()
                                                Exit Do
                                            End If
                                            Dim obj As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForWrite, True) 'открываем объект на чтение
                                            Dim IdLabelLineStyle As ObjectId = Nothing
                                            Dim IdLabelCurvStyle As ObjectId = Nothing

                                            If CivilDoc.Styles.LabelStyles.GeneralLineLabelStyles.Contains(NameObject) = True Then
                                                IdLabelLineStyle = CivilDoc.Styles.LabelStyles.GeneralLineLabelStyles.Item(NameObject)
                                            Else
                                                IdLabelLineStyle = CivilDoc.Styles.LabelStyles.GeneralLineLabelStyles.Item(0)
                                            End If
                                            If CivilDoc.Styles.LabelStyles.GeneralCurveLabelStyles.Contains(NameObject) = True Then
                                                IdLabelCurvStyle = CivilDoc.Styles.LabelStyles.GeneralCurveLabelStyles.Item(NameObject)
                                            Else
                                                IdLabelCurvStyle = CivilDoc.Styles.LabelStyles.GeneralCurveLabelStyles.Item(NameObject)
                                            End If

                                            Dim IDlb As ObjectId = GeneralSegmentLabel.Create(obj.Id, 0, IdLabelLineStyle, IdLabelCurvStyle)
                                            'Dim surfaceElevnLblId As ObjectId = SurfaceElevationLabel.Create(surfaceId, pnt2d)
                                            'Dim label As SurfaceElevationLabel = surfaceElevnLblId.GetObject(OpenMode.ForRead)

                                            Dim lb As GeneralSegmentLabel = acTrans.GetObject(IDlb, OpenMode.ForWrite)
                                            lb.LabelLocation = acSSPrompt.PickedPoint
                                            Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                                            If acLyrTbl.Has(NameLayer) = True Then
                                                lb.Layer = NameLayer
                                            End If
                                        Catch ex As System.Exception
                                        End Try
                                    End Using
                                Loop
                                '===========================================================================================================================
                                'вставка мультивыноски
                            ElseIf typeZn Like "Мультивыноска" Then
                                Dim ContentMLeader = "<LABEL>"
                                Dim numerator As String = "Text1"
                                Dim denumerator As String = ""
                                'получаем ID таблицы PS
                                If PSTable.Trim.Length > 0 Then
                                    Dim boolReadFields As Boolean = FuncPropertySet.FuncReadAllValueFieldSetTables(PSTable, ArrayODRec)
                                End If
                                '===============================================================================================================
                                'формируем строку текста для мультивыноски из поля примечание
                                If IsNothing(Note) = False Then
                                    If Note.Trim.Length > 0 Then
                                        Dim ArrayNote As String() = Note.Split(";")
                                        If IsArray(ArrayNote) = True Then
                                            For i As Integer = 0 To ArrayNote.Length - 1
                                                Dim str As String = ArrayNote(i)
                                                Dim pos As Integer = str.IndexOf("=")
                                                If pos > 0 Then
                                                    Dim NameField As String = Mid(str, 1, pos)
                                                    Dim valStr As String = Mid(str, pos + 2)
                                                    If i < ArrayNote.Length - 1 Then
                                                        For j As Integer = i + 1 To ArrayNote.Length - 1
                                                            Dim TempStr As String = ArrayNote(j)
                                                            If TempStr.IndexOf("=") = -1 Then
                                                                valStr = valStr & ";" & TempStr
                                                                i = j
                                                            End If
                                                        Next
                                                    End If
                                                    If NameField.Trim Like "numerator" Then
                                                        numerator = valStr
                                                    ElseIf NameField.Trim Like "denominator" Then
                                                        denumerator = valStr
                                                    End If
                                                End If
                                            Next i
                                        End If
                                    End If
                                End If

                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'тело программы
                                '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                'Выбираем ось трассы
                                '================================================================================================
                                'формируем надпись
                                Dim TypeParentObject As String = "" 'имя поверхности или трассы
                                Dim PosSt As Integer = -1
                                PosSt = numerator.IndexOf("Station")
                                If PosSt = -1 Then
                                    PosSt = denumerator.IndexOf("Station")
                                End If

                                Dim PosSurfElev As Integer = -1
                                PosSurfElev = numerator.IndexOf("Elevation")
                                If PosSurfElev = -1 Then
                                    PosSurfElev = denumerator.IndexOf("Elevation")
                                End If

                                Dim PosSurfSlope As Integer = -1
                                PosSurfSlope = numerator.IndexOf("Slope")
                                If PosSurfSlope = -1 Then
                                    PosSurfSlope = denumerator.IndexOf("Slope")
                                End If

                                '=======================================================================================
                                'Выбор точки или объекта
                                Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                                pKeyOpts = New PromptKeywordOptions("")
                                pKeyOpts.Message = vbLf & "Способ создания? "
                                pKeyOpts.Keywords.Add("Объект")
                                pKeyOpts.Keywords.Add("Точка")
                                pKeyOpts.Keywords.Default = "Объект"
                                Dim pKeyRes As PromptResult = ed1.GetKeywords(pKeyOpts)
                                If pKeyRes.Status = PromptStatus.OK Then
                                    Do
                                        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                                            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                                            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                            Try
                                                Dim ArrayData As String(,) = Nothing
                                                Dim PicPoint As Point3d = New Point3d(-1, -1, -1)
                                                Dim obj As Autodesk.AutoCAD.DatabaseServices.Entity = Nothing
                                                '=======================================================================================================
                                                'выбор объекта
                                                If pKeyRes.StringResult Like "Объект" Then
                                                    Dim acSSPrompt As PromptEntityResult = ed1.GetEntity("^C^C " & vbLf & "Выберите объект: ")
                                                    If acSSPrompt.Status = PromptStatus.Cancel Then
                                                        acTrans.Commit()
                                                        Exit Do
                                                    End If
                                                    obj = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForRead)
                                                    'считываем расширенные данные
                                                    Dim boolFindData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(obj, ArrayData)
                                                    'формируем набор характеристик для записи в мультивыноску
                                                    Dim boolLocalId As Boolean = False
                                                    If IsArray(ArrayData) = True And IsArray(ArrayODRec) = True Then
                                                        For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                                            Dim TempParentFields As String = ArrayData(1, i)
                                                            For j As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                                Dim TempFieldsMLeader As String = ArrayODRec(1, j)
                                                                If TempParentFields Like TempFieldsMLeader Then
                                                                    ArrayODRec(2, j) = ArrayData(2, i)
                                                                ElseIf TempParentFields Like "LOCALID" And TempFieldsMLeader Like "IDRNG" Then
                                                                    ArrayODRec(2, j) = ArrayData(2, i)
                                                                    boolLocalId = True
                                                                End If
                                                            Next j
                                                        Next i
                                                    End If
                                                    PicPoint = acSSPrompt.PickedPoint
                                                    obj = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForRead) 'открываем объект на чтение
                                                    If NameLayer Like "0" Then
                                                        NameLayerIsh = obj.Layer & NameLayerIsh
                                                        Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                                                        If acLyrTbl.Has(NameLayerIsh) = True Then
                                                            IDNameLayer = acLyrTbl.Item(NameLayerIsh)
                                                            NameLayer = NameLayerIsh
                                                        Else
                                                            Dim boolLayer As ObjectId = AcadFunction.FuncAddLayer(NameLayerIsh)
                                                            If acLyrTbl.Has(NameLayerIsh) = True Then
                                                                IDNameLayer = acLyrTbl.Item(NameLayerIsh)
                                                                NameLayer = NameLayerIsh
                                                            End If
                                                        End If
                                                    End If
                                                    '============================================================================================
                                                    'выбор точки вставки
                                                Else
                                                    Dim acSSPrompt As PromptPointResult = ed1.GetPoint("^C^C " & vbLf & "Выберите точку: ")
                                                    PicPoint = acSSPrompt.Value
                                                    If acSSPrompt.Status = PromptStatus.Cancel Then
                                                        acTrans.Commit()
                                                        Exit Do
                                                    End If
                                                End If

                                                If numerator.IndexOf("Station") > -1 Then
                                                    Dim st As Double = -1
                                                    Dim off As Double = -1
                                                    Try
                                                        acAlign.StationOffset(PicPoint.X, PicPoint.Y, st, off)
                                                        numerator = numerator.Replace("<Station>", acAlign.GetStationStringWithEquations(st))
                                                    Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                        Continue Do
                                                    Catch ex As ArgumentException
                                                        Continue Do
                                                    End Try
                                                End If
                                                If denumerator.IndexOf("Station") > -1 Then
                                                    Dim st As Double = -1
                                                    Dim off As Double = -1
                                                    Try
                                                        acAlign.StationOffset(PicPoint.X, PicPoint.Y, st, off)
                                                        denumerator = denumerator.Replace("<Station>", acAlign.GetStationStringWithEquations(st))
                                                    Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                        Continue Do
                                                    Catch ex As ArgumentException
                                                        Continue Do
                                                    End Try
                                                End If

                                                If numerator.IndexOf("Elevation") > -1 Then
                                                    Try
                                                        Dim elev As Double = acSurf.FindElevationAtXY(PicPoint.X, PicPoint.Y)
                                                        numerator = numerator.Replace("<Elevation>", FuncFormatZN.FuncFormatCoordinate(elev, 2))
                                                    Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                        Continue Do
                                                    Catch ex As ArgumentException
                                                        Continue Do
                                                    End Try
                                                End If
                                                If denumerator.IndexOf("Elevation") > -1 Then
                                                    Try
                                                        Dim elev As Double = acSurf.FindElevationAtXY(PicPoint.X, PicPoint.Y)
                                                        denumerator = denumerator.Replace("<Elevation>", FuncFormatZN.FuncFormatCoordinate(elev, 2))
                                                    Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                        Continue Do
                                                    Catch ex As ArgumentException
                                                        Continue Do
                                                    End Try
                                                End If

                                                If numerator.IndexOf("Slope") > -1 Then
                                                    Try
                                                        Dim slope As Double = acSurf.FindSlopeAtXY(PicPoint.X, PicPoint.Y) * 1000
                                                        numerator = numerator.Replace("<Slope>", FuncFormatZN.FuncFormatCoordinate(slope, 2))
                                                    Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                        Continue Do
                                                    Catch ex As ArgumentException
                                                        Continue Do
                                                    End Try
                                                End If
                                                If denumerator.IndexOf("Slope") > -1 Then
                                                    Try
                                                        Dim slope As Double = acSurf.FindSlopeAtXY(PicPoint.X, PicPoint.Y) * 1000
                                                        denumerator = denumerator.Replace("<Slope>", FuncFormatZN.FuncFormatCoordinate(slope, 2))
                                                    Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                        Continue Do
                                                    Catch ex As ArgumentException
                                                        Continue Do
                                                    End Try
                                                End If

                                                If numerator.IndexOf("X") > -1 Then
                                                    numerator = numerator.Replace("<X>", FuncFormatZN.FuncFormatCoordinate(PicPoint.Y, 2))
                                                End If
                                                If numerator.IndexOf("Y") > -1 Then
                                                    numerator = numerator.Replace("<Y>", FuncFormatZN.FuncFormatCoordinate(PicPoint.X, 2))
                                                End If

                                                If denumerator.IndexOf("X") > -1 Then
                                                    denumerator = denumerator.Replace("<X>", FuncFormatZN.FuncFormatCoordinate(PicPoint.Y, 2))
                                                End If

                                                If denumerator.IndexOf("Y") > -1 Then
                                                    denumerator = denumerator.Replace("<Y>", FuncFormatZN.FuncFormatCoordinate(PicPoint.X, 2))
                                                End If

                                                If numerator.IndexOf("Area") > -1 Then
                                                    Dim ArrayIDOstrov As ObjectId() = Nothing
                                                    Dim CountArrayIDOstrov As Integer = 0
                                                    Dim ValLocalID As String = FuncPropertySet.FuncReadValueSetProperties(obj, "LOCALID")
                                                    If ValLocalID.Trim.Length > 0 Then
                                                        'ищем внутренние островки
                                                        Dim ArrayOstrov As String(,) = Nothing
                                                        Dim boolFindOstrov As Boolean = FuncCadastr.FuncFindOstrov(ValLocalID, ArrayOstrov)
                                                        If IsArray(ArrayOstrov) = True Then
                                                            For i As Integer = 0 To ArrayOstrov.GetUpperBound(1)
                                                                Dim hg As String = ArrayOstrov(1, i)
                                                                Dim IDOstrov As ObjectId = AcadFunction.FuncReturnObjectID(hg)
                                                                If IsNothing(IDOstrov) = False Then
                                                                    If IDOstrov.IsNull = False Then
                                                                        ReDim Preserve ArrayIDOstrov(CountArrayIDOstrov)
                                                                        ArrayIDOstrov(CountArrayIDOstrov) = IDOstrov
                                                                        CountArrayIDOstrov += 1
                                                                    End If
                                                                End If
                                                            Next i
                                                        End If
                                                        Dim AreaStr As String = FuncCadastr.FuncCreateFieldSQRTPolyline(obj.Id, ArrayIDOstrov)
                                                        numerator = numerator.Replace("<Area>", AreaStr)
                                                    Else
                                                        Dim StrId As Long = obj.Id.OldIdPtr
                                                        Dim AreaStr As String = "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Area \f " & """" & "%lu2%pr2" & """" & ">%"
                                                        numerator = numerator.Replace("<Area>", AreaStr)
                                                    End If
                                                End If
                                                If denumerator.IndexOf("Area") > -1 Then
                                                    Dim ArrayIDOstrov As ObjectId() = Nothing
                                                    Dim CountArrayIDOstrov As Integer = 0
                                                    Dim ValLocalID As String = FuncPropertySet.FuncReadValueSetProperties(obj, "LOCALID")
                                                    If ValLocalID.Trim.Length > 0 Then
                                                        'ищем внутренние островки
                                                        Dim ArrayOstrov As String(,) = Nothing
                                                        Dim boolFindOstrov As Boolean = FuncCadastr.FuncFindOstrov(ValLocalID, ArrayOstrov)
                                                        If IsArray(ArrayOstrov) = True Then
                                                            For i As Integer = 0 To ArrayOstrov.GetUpperBound(1)
                                                                Dim hg As String = ArrayOstrov(1, i)
                                                                Dim IDOstrov As ObjectId = AcadFunction.FuncReturnObjectID(hg)
                                                                If IsNothing(IDOstrov) = False Then
                                                                    If IDOstrov.IsNull = False Then
                                                                        ReDim Preserve ArrayIDOstrov(CountArrayIDOstrov)
                                                                        ArrayIDOstrov(CountArrayIDOstrov) = IDOstrov
                                                                        CountArrayIDOstrov += 1
                                                                    End If
                                                                End If
                                                            Next i
                                                        End If
                                                        Dim AreaStr As String = FuncCadastr.FuncCreateFieldSQRTPolyline(obj.Id, ArrayIDOstrov)
                                                        denumerator = denumerator.Replace("<Area>", AreaStr)
                                                    Else
                                                        Dim StrId As Long = obj.Id.OldIdPtr
                                                        Dim AreaStr As String = "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Area \f " & """" & "%lu2%pr2" & """" & ">%"
                                                        denumerator = denumerator.Replace("<Area>", AreaStr)
                                                    End If
                                                End If

                                                If numerator.IndexOf("Lenght") > -1 Then
                                                    Dim StrId As Long = obj.Id.OldIdPtr
                                                    Dim AreaStr As String = "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Lenght \f " & """" & "%lu2%pr2" & """" & ">%"
                                                    numerator = numerator.Replace("<Lenght>", AreaStr)
                                                End If
                                                If denumerator.IndexOf("Lenght") > -1 Then
                                                    Dim StrId As Long = obj.Id.OldIdPtr
                                                    Dim AreaStr As String = "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Lenght \f " & Chr(34) & "%lu2%pr2" & Chr(34) & ">%"
                                                    denumerator = denumerator.Replace("<Lenght>", AreaStr)
                                                End If

                                                'формируем надпись
                                                If IsArray(ArrayData) = True Then
                                                    For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                                        Dim TempParentFields As String = ArrayData(1, i).Trim
                                                        Dim Pos As Integer = numerator.IndexOf(TempParentFields)
                                                        If Pos > -1 Then
                                                            If IsNothing(ArrayData(2, i)) = False Then
                                                                numerator = numerator.Replace("<" & TempParentFields & ">", ArrayData(2, i).Trim)
                                                            End If
                                                        End If
                                                        Dim Pos1 As Integer = denumerator.IndexOf(TempParentFields)
                                                        If Pos1 > -1 Then
                                                            If IsNothing(ArrayData(2, i)) = False Then
                                                                denumerator = denumerator.Replace("<" & TempParentFields & ">", ArrayData(2, i).Trim)
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                                ContentMLeader = numerator
                                                If denumerator.Length > 0 Then
                                                    ContentMLeader = numerator & vbLf & denumerator
                                                End If
                                                '=========================================================================================================
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                'вставляем объект
                                                'определяем точку вставки полки для мультивыноски
                                                If IsNothing(NameLayer) = True Then
                                                    NameLayer = obj.Layer
                                                ElseIf NameLayer.Trim.Length = 0 Then
                                                    NameLayer = obj.Layer
                                                End If
                                                Dim IDMLeader As ObjectId = AcadFunction.FuncDrawMleader(PicPoint.Convert2d(New Plane), ContentMLeader, 5, 5, NameLayer, textheight, IDNameObject, idTextStyle,, True)
                                                If IsNothing(IDMLeader) = False Then
                                                    If IDMLeader.IsNull = False Then
                                                        Dim MLeader1 As MLeader = IDMLeader.GetObject(OpenMode.ForWrite)
                                                        MLeader1.ArrowSize = M * arrowSize
                                                        MLeader1.LineWeight = LineWidth
                                                        If IsNothing(idPsTable) = False Then
                                                            If idPsTable.IsNull = False Then
                                                                If IsArray(ArrayODRec) = True Then
                                                                    For i As Integer = 0 To ArrayODRec.GetUpperBound(1)
                                                                        Dim nameF As String = ArrayODRec(1, i)
                                                                        If nameF Like "LABEL" Then
                                                                            ArrayODRec(2, i) = deskZn
                                                                        ElseIf nameF Like "PARENT_NAME" Then
                                                                            ArrayODRec(2, i) = TypeParentObject
                                                                        End If
                                                                    Next
                                                                End If
                                                                Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(MLeader1, idPsTable)
                                                                Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(MLeader1, ArrayODRec)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Catch ex As System.Exception
                                            End Try
                                            acTrans.Commit()
                                        End Using
                                    Loop
                                End If
                            End If
                            'сьрасывам ссылки на поверхность
                            If IsNothing(acSurf) = False Then
                                acSurf.Dispose()
                                acSurf = Nothing
                            End If
                            If IsNothing(acAlign) = False Then
                                acAlign.Dispose()
                                acAlign = Nothing
                            End If
                        End If
                    End If
                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                End Try
            End Using
            docLock.Dispose()
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор элемента
    Private Sub DataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles DataGrid.SelectionChanged
        Try
            Dim userSelectRow As System.Windows.Controls.DataGrid = e.Source
            Dim userData As Object = userSelectRow.SelectedValue
            If TypeOf userData Is Customer Then
                Dim userCustomerData As Customer = userData
                If userCustomerData.Имя_свойства Like "Category" Then
                    Dim Kategory As String() = {"Земли сельскохозяйственного назначения", "Земли населенных пунктов", "Земли промышленности", "Земли особо охраняемых территорий и объектов", "Земли лесного фонда", "Земли водного фонда", "Земли запаса", "Категория не установлена"}
                    Dim formSectZn As SelectTemplateFiles = New SelectTemplateFiles()
                    formSectZn.ListBox1.DataSource = Kategory
                    formSectZn.ShowDialog()
                    If formSectZn.boolShow = True Then
                        Dim newData As String = formSectZn.ListBox1.SelectedValue
                        If IsNothing(newData) = False Then
                            userCustomerData.Значение = newData
                        End If
                    End If
                    e.Source = userCustomerData
                    DataGrid.Items.Refresh()
                ElseIf userCustomerData.Имя_свойства Like "Utilization" Then
                    Dim ArrayUtilization As String() = {"1 Сельскохозяйственное использование", "1.1	Растениеводство", "1.2 Выращивание зерновых и иных сельскохозяйственных культур", "1.3 Овощеводство", "1.4	Выращивание тонизирующих, лекарственных, цветочных культур",
                "1.5 Садоводство", "1.6 Выращивание льна и конопли", "1.7	Животноводство", "1.8 Скотоводство", "1.9	Звероводство", "1.1 Птицеводство", "1.11 Свиноводство", "1.12 Пчеловодство", "1.13 Рыбоводство", "1.14 Научное обеспечение сельского хозяйства",
                "1.15 Хранение и переработка сельскохозяйственной продукции", "1.16 Ведение личного подсобного хозяйства на полевых участках", "1.17 Питомники", "1.18 Обеспечение сельскохозяйственного производства", "1.19 Сенокошение", "1.2 Выпас сельскохозяйственных животных",
                "2 Жилая застройка", "2.1 Для индивидуального жилищного строительства", "2.1.1 Малоэтажная многоквартирная жилая застройка", "2.2 Для ведения личного подсобного хозяйства (приусадебный земельный участок)", "2.3 Блокированная жилая застройка", "2.4	Передвижное жилье",
                "2.5 Среднеэтажная жилая застройка", "2.6 Многоэтажная жилая застройка (высотная застройка)", "2.7 Обслуживание жилой застройки", "2.7.1	Хранение автотранспорта", "3 Общественное использование объектов капитального строительства", "3.1 Коммунальное обслуживание",
                "3.1.1 Предоставление коммунальных услуг", "3.1.2 Административные здания организаций, обеспечивающих предоставление коммунальных услуг", "3.2 Социальное обслуживание", "3.2.1 Дома социального обслуживания", "3.2.2 Оказание социальной помощи населению",
                "3.2.3 Оказание услуг связи", "3.2.4	Общежития", "3.3	Бытовое обслуживание", "3.4 Здравоохранение", "3.4.1 Амбулаторно-поликлиническое обслуживание", "3.4.2	Стационарное медицинское обслуживание", "3.4.3 Медицинские организации особого назначения", "3.5 Образование и просвещение",
                "3.5.1 Дошкольное, начальное и среднее общее образование", "3.5.2 Среднее и высшее профессиональное образование", "3.6 Культурное развитие", "3.6.1 Объекты культурно-досуговой деятельности", "3.6.2 Парки культуры и отдыха", "3.6.3 Цирки и зверинцы",
            "3.7 Религиозное использование", "3.7.1 Осуществление религиозных обрядов", "3.7.2 Религиозное управление и образование", "3.8 Общественное управление", "3.8.1 Государственное управление", "3.8.2 Представительская деятельность", "3.9 Обеспечение научной деятельности",
            "3.9.1 Обеспечение деятельности в области гидрометеорологии и смежных с ней областях", "3.9.2 Проведение научных исследований", "3.9.3 Проведение научных испытаний", "3.10 Ветеринарное обслуживание", "3.10.1	Амбулаторное ветеринарное обслуживание",
            "3.10.2 Приюты для животных", "4 Предпринимательство", "4.1 Деловое управление", "4.2 Объекты торговли (торговые центры, торгово-развлекательные центры (комплексы)", "4.3 Рынки", "4.4 Магазины", "4.5 Банковская и страховая деятельность", "4.6 Общественное питание",
            "4.7 Гостиничное обслуживание", "4.8	Развлечение", "4.8.1 Развлекательные мероприятия", "4.8.2	Проведение азартных игр", "4.8.3 Проведение азартных игр в игорных зонах", "4.9 Служебные гаражи", "4.9.1 Объекты дорожного сервиса", "4.9.1.1 Заправка транспортных средств",
            "4.9.1.2 Обеспечение дорожного отдыха", "4.9.1.3 Автомобильные мойки", "4.9.1.4 Ремонт автомобилей", "4.1 Выставочно-ярмарочная деятельность", "5 Отдых (рекреация)", "5.1 Спорт", "5.1.1 Обеспечение спортивно-зрелищных мероприятий", "5.1.2	Обеспечение занятий спортом в помещениях",
            "5.1.3 Площадки для занятий спортом", "5.1.4 Оборудованные площадки для занятий спортом", "5.1.5 Водный спорт", "5.1.6 Авиационный спорт", "5.1.7 Спортивные базы", "5.2 Природно-познавательный туризм", "5.2.1 Туристическое обслуживание", "5.3 Охота и рыбалка", "5.4 Причалы для маломерных судов",
    "5.5 Поля для гольфа или конных прогулок", "6 Производственная деятельность", "6.1 Недропользование", "6.2	Тяжелая промышленность", "6.2.1 Автомобилестроительная промышленность", "6.3 Легкая промышленность", "6.3.1 Фармацевтическая промышленность", "6.4 Пищевая промышленность", "6.5 Нефтехимическая промышленность",
    "6.6 Строительная промышленность", "6.7 Энергетика", "6.7.1 Атомная энергетика", "6.8 Связь", "6.9 Склад", "6.9.1 Складские площадки", "6.1 Обеспечение космической деятельности", "6.11 Целлюлозно-бумажная промышленность", "6.12 Научно-производственная деятельность", "7 Транспорт", "7.1 Железнодорожный транспорт",
    "7.1.1 Железнодорожные пути", "7.1.2 Обслуживание железнодорожных перевозок", "7.2 Автомобильный транспорт", "7.2.1 Размещение автомобильных дорог", "7.2.2 Обслуживание перевозок пассажиров", "7.2.3 Стоянки транспорта общего пользования", "7.3 Водный транспорт", "7.4 Воздушный транспорт", "7.5	Трубопроводный транспорт",
    "7.6 Внеуличный транспорт", "8 Обеспечение обороны и безопасности", "8.1 Обеспечение вооруженных сил", "8.2 Охрана Государственной границы Российской Федерации", "8.3 Обеспечение внутреннего правопорядка", "8.4 Обеспечение деятельности по исполнению наказаний", "9 Деятельность по особой охране и изучению природы",
    "9.1 Охрана природных территорий", "9.2 Курортная деятельность", "9.2.1 Санаторная деятельность", "9.3 Историко-культурная деятельность", "10 Использование лесов", "10.1 Заготовка древесины", "10.2 Лесные плантации", "10.3 Заготовка лесных ресурсов", "10.4 Резервные леса", "11 Водные объекты",
    "11.1 Общее пользование водными объектами", "11.2 Специальное пользование водными объектами", "11.3 Гидротехнические сооружения", "12 Земельные участки (территории) общего пользования", "12.0.1 Улично-дорожная сеть", "12.0.2 Благоустройство территории", "12.1	Ритуальная деятельность",
    "12.2 Специальная деятельность", "12.3 Запас", "13 Земельные участки общего назначения", "13.1 Ведение огородничества", "13.2 Ведение садоводства"}
                    Dim formSectZn As SelectTemplateFiles = New SelectTemplateFiles()
                    formSectZn.ListBox1.DataSource = ArrayUtilization
                    formSectZn.ShowDialog()
                    If formSectZn.boolShow = True Then
                        Dim newData As String = formSectZn.ListBox1.SelectedValue
                        If IsNothing(newData) = False Then
                            userCustomerData.Значение = newData
                        End If
                    End If
                    e.Source = userCustomerData
                    DataGrid.Items.Refresh()
                ElseIf userCustomerData.Имя_свойства Like "Right" Then
                    Dim ArrayVidPrava As String() = {"", "Частная собственность", "Государственная собственность", "Муниципальная собственность", "Собственность РФ", "Собственность субъекта РФ", "Неразграниченаая государственная собственность"}
                    Dim formSectZn As SelectTemplateFiles = New SelectTemplateFiles()
                    formSectZn.ListBox1.DataSource = ArrayVidPrava
                    formSectZn.ShowDialog()
                    If formSectZn.boolShow = True Then
                        Dim newData As String = formSectZn.ListBox1.SelectedValue
                        If IsNothing(newData) = False Then
                            userCustomerData.Значение = newData
                        End If
                    End If
                    e.Source = userCustomerData
                    DataGrid.Items.Refresh()
                ElseIf userCustomerData.Имя_свойства Like "Method_create" Then
                    Dim VidCadastralWork As String() = {"Образование земельного участка из земель, находящихся в государственной или муниципальной собственности",
                                                "Образование земельного участка путем раздела земельного участка",
                                                "Образование земельного участка путем раздела с сохранением земельного участка, раздел которого осуществлен, в измененных границах",
                                                "Образование земельного участка из состава земельного участка, представляющего собой единое землепользование",
                                                "Образование земельного участка путем объединения земельных участков",
                                                "Образование земельного участка при создании искусственного земельного участка",
                                                "Образование земельного участка путем перераспределения земельных участков",
                                                "Образование земельного участка путем перераспределения земельных участков или земель, находящихся в государственной или муниципальной собственности",
                                                "Уточнение местоположения границ и (или) площади земельного участка",
                                                "Уточнение описания местоположения границ и (или) площади земельного участка с целью исправления ошибки в описании местоположения его границ и (или) площади",
                                                "Образование части земельного участка",
                                                "Уточнение части земельного участка",
                                                "Установление сервитута",
                                                "Установление публичного сервитута"}
                    Dim formSectZn As SelectTemplateFiles = New SelectTemplateFiles()
                    formSectZn.ListBox1.DataSource = VidCadastralWork
                    formSectZn.ShowDialog()
                    If formSectZn.boolShow = True Then
                        Dim newData As String = formSectZn.ListBox1.SelectedValue
                        If IsNothing(newData) = False Then
                            userCustomerData.Значение = newData
                        End If
                    End If
                    e.Source = userCustomerData
                    DataGrid.Items.Refresh()
                ElseIf userCustomerData.Имя_свойства Like "Confiscation" Then
                    Dim ArrayObosnovanie As String() = {"Полное изъятие земельного участка для государственных или муниципальных нужд (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.56.3)",
                                                "Частичное изъятие земельного участка для государственных или муниципальных нужд (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.56.3)",
                                                "Резервирование земельного участка для государственных или муниципальных нужд (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.70.1)",
                                                "Установление публичного сервитута для прохода или проезда через земельный участок, в том числе в целях обеспечения свободного доступа граждан к водному объекту общего пользования и его береговой полосе (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.23)",
                                                "Установление публичного сервитута для размещения на земельном участке межевых знаков, геодезических пунктов государственных геодезических сетей, гравиметрических пунктов, нивелирных пунктов и подъездов к ним (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.23)",
                                                "Установление публичного сервитута для проведения дренажных и мелиоративных работ на земельном участке (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.23)",
                                                "Установление публичного сервитута для забора (изъятия) водных ресурсов из водных объектов и водопоя (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.23)",
                                                "Установление публичного сервитута для строительства, реконструкции, эксплуатации, капитального ремонта объектов электросетевого хозяйства, тепловых сетей, водопроводных сетей, сетей водоотведения, линий и сооружений связи, линейных объектов системы газоснабжения, нефтепроводов и нефтепродуктопроводов, их неотъемлемых технологических частей (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.39.37)",
                                                "Установление публичного сервитута для складирования строительных и иных материалов, размещение временных или вспомогательных сооружений (включая ограждения, бытовки, навесы) и (или) строительной техники, которые необходимы для обеспечения строительства, реконструкции, ремонта объектов транспортной инфраструктуры федерального, регионального или местного значения, на срок указанных строительства, реконструкции, ремонта (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.39.37)",
                                                "Установление публичного сервитута для устройства пересечений автомобильных дорог или железнодорожных путей с железнодорожными путями общего пользования на земельных участках, находящихся в государственной собственности, в границах полос отвода железных дорог, а также устройство пересечений автомобильных дорог или железнодорожных путей с автомобильными дорогами или примыканий автомобильных дорог к другим автомобильным дорогам на земельных участках, находящихся в государственной или муниципальной собственности, в границах полосы отвода автомобильной дороги (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.39.37)",
                                                "Установление публичного сервитута для размещения автомобильных дорог и железнодорожных путей в туннелях (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.23)",
                                                "Установление публичного сервитута для реконструкции, капитального ремонта участков (частей) инженерных сооружений, являющихся линейными объектами (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.23)",
                                                "Установление сервитута для обеспечения прохода и проезда через соседний земельный участок, строительства, реконструкции и (или) эксплуатации линейных объектов, не препятствующих использованию земельного участка в соответствии с разрешенным использованием, а также других нужд собственника недвижимого имущества, которые не могут быть обеспечены без установления сервитута (Гражданский кодекс Российской Федерации от 30.11.1994 N 51-ФЗ, Ст.274)",
                                                "Нет"}
                    Dim formSectZn As SelectTemplateFiles = New SelectTemplateFiles()
                    formSectZn.ListBox1.DataSource = ArrayObosnovanie
                    formSectZn.ShowDialog()
                    If formSectZn.boolShow = True Then
                        Dim newData As String = formSectZn.ListBox1.SelectedValue
                        If IsNothing(newData) = False Then
                            userCustomerData.Значение = newData
                        End If
                    End If
                    e.Source = userCustomerData
                    DataGrid.Items.Refresh()
                ElseIf userCustomerData.Имя_свойства Like "CommonLands" Then
                    Dim ArrayVidPrava As String() = {"Да", "Нет"}
                    Dim formSectZn As SelectTemplateFiles = New SelectTemplateFiles()
                    formSectZn.ListBox1.DataSource = ArrayVidPrava
                    formSectZn.ShowDialog()
                    If formSectZn.boolShow = True Then
                        Dim newData As String = formSectZn.ListBox1.SelectedValue
                        If IsNothing(newData) = False Then
                            userCustomerData.Значение = newData
                        End If
                    End If
                    e.Source = userCustomerData
                    DataGrid.Items.Refresh()
                End If
            End If
            DataGrid.SelectedIndex = 0
        Catch ex As System.Exception
        End Try
    End Sub
End Class

