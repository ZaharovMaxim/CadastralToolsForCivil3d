Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.ApplicationServices
Public Class ExportCadastralObject
    '=============================================================================================
    'создать таблицу OD из PSSet
    <CommandMethod("ExportPSTableToODTable")>
    Public Sub ExportPSTableToODTable()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim docMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        'Dim doc As Document = docMgr.Open("D:\\1.dwg", False)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'сохраняем документ
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
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем объекты
            Dim acSSPrompt As PromptSelectionResult
            Dim acSSet1 As SelectionSet = Nothing
            acSSPrompt = ed.GetSelection() 'выбираем контура
            If acSSPrompt.Status = PromptStatus.OK Then
                acSSet1 = acSSPrompt.Value
            Else
                locker.Dispose()
                Exit Sub
            End If
            Dim count As Integer = 0
            For Each objId1 As ObjectId In acSSet1.GetObjectIds()
                Try
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(objId1, OpenMode.ForRead)
                    'читаем таблицы PS Set
                    Dim ArrayPsTable As String() = Nothing
                    Dim CountReadPsTable As Integer = FuncPropertySet.FuncReadAllDataSetTables(ent, ArrayPsTable)
                    If IsArray(ArrayPsTable) = True Then
                        For i As Integer = 0 To ArrayPsTable.Length - 1
                            Dim NamePsTable As String = ArrayPsTable(i)
                            'читаем поля таблицы и его значения
                            Dim ArrayData As String(,) = Nothing
                            Dim boolArrayData As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(ent, ArrayData, NamePsTable)
                            'читаем характеристики полей PSTable
                            Dim ArrayFields As String(,) = Nothing
                            Dim boolArrayFields As Boolean = FuncPropertySet.FuncReadAllValueFieldSetTables(NamePsTable, ArrayFields)
                            'массив с характеристиками полей заполнен
                            If IsNothing(ArrayFields) = False Then
                                'создаем новую таблицу OD
                                Dim boolCreateTable As Boolean = FuncOD.FuncCreateODTable(NamePsTable, ArrayFields)
                                'проверяем может примитиву уже приспоена такая таблица
                                Dim arrayODTable As String() = Nothing
                                Dim boolArrayODTable As Boolean = FuncOD.FuncFindTableOD(arrayODTable, ent)
                                If IsArray(arrayODTable) = True Then
                                    For j As Integer = 0 To arrayODTable.Length - 1
                                        Dim NameOd As String = arrayODTable(j)
                                        If NameOd Like NamePsTable Then
                                            'обновляем данные
                                            If ent.IsWriteEnabled = False Then
                                                ent.UpgradeOpen()
                                            End If
                                            Dim boolOdReplase As Boolean = FuncOD.FuncReplaseOD(ArrayData, NamePsTable, ent.Id)
                                            Continue For
                                        End If
                                    Next j
                                End If
                                'записываем данные в нее
                                If ent.IsWriteEnabled = False Then
                                    ent.UpgradeOpen()
                                End If
                                Dim boolWriteData As Boolean = FuncOD.FuncAddOD(ArrayData, NamePsTable, ent.Id)
                            End If
                        Next i
                    End If
                Catch ex As System.Exception
                End Try
                count += 1
            Next
            tr.Commit()
        End Using
        locker.Dispose()
    End Sub
    '=============================================================================================
    'выбор объектов по значению поля набора данных
    <CommandMethod("SelectObjectToValuePSFields")>
    Public Sub SelectObjectToValuePSFields()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'сохраняем документ
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
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & "Введите имя поля таблицы PS: ")
            pStrOpts.AllowSpaces = True
            Dim pStrResFields As PromptResult = acDoc.Editor.GetString(pStrOpts)
            Dim pStrOpts1 As PromptStringOptions = New PromptStringOptions(vbLf & "Введите текстовую строку для поиска: ")
            pStrOpts1.AllowSpaces = True
            Dim pStrResValue As PromptResult = acDoc.Editor.GetString(pStrOpts1)
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.SelectAll()
            Dim arrayID As ObjectId() = Nothing
            Dim countarrayId As Integer = 0
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        If Not IsDBNull(acEnt) Then
                            Dim TempStr As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, pStrResFields.StringResult)
                            If TempStr.Trim.Length > 0 Then
                                If TempStr Like pStrResValue.StringResult Then
                                    ReDim Preserve arrayID(countarrayId)
                                    arrayID(countarrayId) = acEnt.Id
                                    countarrayId += 1
                                End If
                            End If
                        End If
                    End If
                Next
                If IsArray(arrayID) = True Then
                    ed.SetImpliedSelection(arrayID)
                End If
            End If
            tr.Commit()
        End Using
        locker.Dispose()
    End Sub
    '=============================================================================================
    'выбор объектов по значению поля набора данных (выбор происходит если поле не пустое)
    <CommandMethod("SelectObjectToNotNullValuePS")>
    Public Sub SelectObjectToNotNullValuePS()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'сохраняем документ
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
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & "Введите имя поля таблицы PS: ")
            pStrOpts.AllowSpaces = True
            Dim pStrResFields As PromptResult = acDoc.Editor.GetString(pStrOpts)
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.SelectAll()
            Dim arrayID As ObjectId() = Nothing
            Dim countarrayId As Integer = 0
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        If Not IsDBNull(acEnt) Then
                            Dim TempStr As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, pStrResFields.StringResult)
                            If TempStr.Trim.Length > 0 Then
                                ReDim Preserve arrayID(countarrayId)
                                arrayID(countarrayId) = acEnt.Id
                                countarrayId += 1
                            End If
                        End If
                    End If
                Next
                If IsArray(arrayID) = True Then
                    ed.SetImpliedSelection(arrayID)
                End If
            End If
            tr.Commit()
        End Using
        locker.Dispose()
    End Sub
    '=============================================================================================
    'выбор объектов по значению поля набора данных  и подсчет площади
    <CommandMethod("SelectObjectToValuePSFields_CalcSqere")>
    Public Sub SelectObjectToValuePSFields_CalcSqere()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'сохраняем документ
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
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & "Введите имя поля таблицы PS: ")
            pStrOpts.AllowSpaces = True
            Dim pStrResFields As PromptResult = acDoc.Editor.GetString(pStrOpts)
            Dim pStrOpts1 As PromptStringOptions = New PromptStringOptions(vbLf & "Введите текстовую строку для поиска: ")
            pStrOpts1.AllowSpaces = True
            Dim pStrResValue As PromptResult = acDoc.Editor.GetString(pStrOpts1)
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.SelectAll()
            Dim arrayID As ObjectId() = Nothing
            Dim countObject As Integer = 0
            Dim AreaObject As Double = 0
            Dim LenghtObject As Double = 0
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        If Not IsDBNull(acEnt) Then
                            Dim TempStr As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, pStrResFields.StringResult)
                            If TempStr.Trim.Length > 0 Then
                                If TempStr Like pStrResValue.StringResult Then
                                    countObject += 1
                                    If TypeOf acEnt Is Polyline Then
                                        Dim acPoly As Polyline = acEnt
                                        AreaObject += acPoly.Area
                                        LenghtObject += acPoly.Length
                                    ElseIf TypeOf acEnt Is Polyline2d Then
                                        Dim acpoly2d As Polyline2d = acEnt
                                        AreaObject += acpoly2d.Area
                                        LenghtObject += acpoly2d.Length
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                MsgBox("Количество найденных объектов - " & countObject & vbLf & "Общая площадь - " & AreaObject & vbLf & "Общая длина - " & LenghtObject)
            End If
            tr.Commit()
        End Using
        locker.Dispose()
    End Sub
End Class
