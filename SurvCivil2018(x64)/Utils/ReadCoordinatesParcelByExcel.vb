Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Public Class ReadCoordinatesParcelByExcel
    '=============================================================================================
    'отрисовка полигона на основе соординат точек из таблицы Excel (1 - номер точки,2-X,3-Y,4-desk)    '
    <CommandMethod("ReadCoordParcelByExcel")>
    Public Sub ReadCoordParcelByExcel()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
        Dim PutchTables As String = "d:\AutodeskCivilTools\FileResources\Sample\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTables = parentPatch2 & "\FileResources\Sample\"
        End If
        Dim ArrayTemplates As String() = Nothing
        Dim CountArrayTemplates As Integer = 0
        Dim PutchTemplateFile As String = ""
        'читаем шаблоны условных знаков
        If IO.Directory.Exists(PutchTables) = True Then
            Try
                For Each fileDWG As String In Directory.GetFiles(PutchTables)
                    'читаем расширение
                    Dim ext As String = Path.GetExtension(fileDWG)
                    If ext Like ".dwg" Then
                        Dim tr1 As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
                        Using tr1
                            Dim acBlkTbl As BlockTable = tr1.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                            If acBlkTbl.Has("PROJECT_POINT") = False Then
                                Dim boolCopyBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(fileDWG, "PROJECT_POINT")
                                If boolCopyBlk = True Then
                                    tr1.Commit()
                                    Exit For
                                End If
                            End If
                        End Using
                    End If
                Next
            Catch ex As IOException
            Catch ex As UnauthorizedAccessException
            Catch ex As ArgumentException
            End Try
        Else
            MsgBox("Файл шаблона не найден. Макро превано!!!")
            Exit Sub
        End If
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Try
            Using tr
                Dim acBlkTbl As BlockTable
                acBlkTbl = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                      OpenMode.ForWrite)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Выбираем файл
                Dim ArrayDataWrite As String(,) = Nothing
                Dim CountArrayDataWrite As Integer = 0
                Dim opFile As New System.Windows.Forms.OpenFileDialog()
                opFile.InitialDirectory = "c:\"
                opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
                opFile.RestoreDirectory = True
                opFile.ShowDialog()
                Dim fFile1 As String = opFile.FileName
                If IsNothing(fFile1) = True Then Exit Sub
                If fFile1.Trim.Length = 0 Then Exit Sub
                Dim xlApp As Excel.Application = Nothing
                Dim xlBook As Excel.Workbook = Nothing
                Dim xlSheets As Excel.Worksheet = Nothing
                xlApp = CreateObject("Excel.Application")
                xlBook = xlApp.Workbooks.Open(fFile1)
                Dim xlsheet As Excel.Worksheet
                xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
                Dim ArrayData As String(,) = Nothing
                Dim CountArrayData As Integer = 0
                Dim ErrCount As Integer = 0
                'запрос на обработку всей выбранной полилинии или только ее части
                Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                pIntOpts.Message = vbCrLf & "Поменять местами X,Y? "
                pIntOpts.Keywords.Add("Да")
                pIntOpts.Keywords.Add("Нет")
                pIntOpts.Keywords.Default = "Нет"
                pIntOpts.AllowNone = True
                Dim pIntRes As PromptResult = acDoc.Editor.GetKeywords(pIntOpts)

                Dim boolInsBlk As Boolean = False
                If acBlkTbl.Has("PROJECT_POINT") = True Then
                    boolInsBlk = True
                End If
                For indRows As Integer = 1 To xlsheet.Rows.Count - 1
                    Dim flag As Boolean = False
                    Dim NumPoint As String = xlsheet.Cells(indRows, 1).value
                    If IsNothing(NumPoint) = True Then
                        NumPoint = ""
                    End If
                    Dim x1 As String = xlsheet.Cells(indRows, 2).value
                    If IsNothing(x1) = True Then
                        x1 = ""
                    End If
                    x1 = x1.Replace(",", ".")
                    Dim y1 As String = xlsheet.Cells(indRows, 3).value
                    If IsNothing(y1) = True Then
                        y1 = ""
                    End If
                    y1 = y1.Replace(",", ".")
                    Dim desk As String = xlsheet.Cells(indRows, 4).value
                    If IsNothing(desk) = True Then
                        desk = ""
                    End If
                    If x1.Trim.Length > 0 And y1.Trim.Length > 0 Then
                        ReDim Preserve ArrayData(3, CountArrayData)
                        ArrayData(0, CountArrayData) = NumPoint
                        If pIntRes.StringResult Like "Нет" Then
                            ArrayData(1, CountArrayData) = x1
                            ArrayData(2, CountArrayData) = y1
                        Else
                            ArrayData(1, CountArrayData) = y1
                            ArrayData(2, CountArrayData) = x1
                        End If
                        ArrayData(3, CountArrayData) = desk
                        CountArrayData += 1
                        ErrCount = 0
                    Else
                        If ErrCount > 20 Then
                            Exit For
                        Else
                            If IsArray(ArrayData) = True Then
                                If ArrayData.GetUpperBound(1) > 0 Then
                                    Dim acPoly As Polyline = New Polyline()
                                    acPoly.SetDatabaseDefaults()
                                    For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                        Try
                                            Dim x As Double = Val(ArrayData(2, i))
                                            Dim y As Double = Val(ArrayData(1, i))
                                            acPoly.AddVertexAt(i, New Point2d(x, y), 0, 0, 0)
                                            Dim NamePoint As String = ArrayData(0, i)
                                            If boolInsBlk = True Then
                                                If NamePoint.Trim.Length > 0 Then
                                                    Dim blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(New Point3d(x, y, 0), "PROJECT_POINT")
                                                    Dim ArrayAttrib As String(,) = Nothing
                                                    ReDim ArrayAttrib(1, 0)
                                                    ArrayAttrib(0, 0) = "NUM_GEOPOINT"
                                                    ArrayAttrib(1, 0) = NamePoint
                                                    Dim boolInsAttarib As Boolean = AcadFunction.FuncWriteAllBlockAttr(blk, ArrayAttrib)
                                                End If
                                            End If
                                        Catch ex As System.Exception
                                        End Try
                                    Next i
                                    acBlkTblRec.AppendEntity(acPoly)
                                    tr.AddNewlyCreatedDBObject(acPoly, True)
                                    Erase ArrayData
                                    CountArrayData = 0
                                End If
                            End If
                            ErrCount += 1
                        End If
                    End If
                Next indRows
                tr.Commit()
                xlBook.Close()
                xlsheet = Nothing
                xlBook = Nothing
                xlApp.Quit()
                xlApp = Nothing
            End Using
        Catch ex As Exception
        End Try
    End Sub

    '=============================================================================================
    'отрисовка полигона на основе соординат точек из таблицы Excel (Рослесхоз координаты в 4 колонки)    '
    <CommandMethod("ReadCoordForestByExcel")>
    Public Sub ReadCoordForestByExcel()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
        Dim PutchTables As String = "d:\AutodeskCivilTools\FileResources\Sample\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTables = parentPatch2 & "\FileResources\Sample\"
        End If
        Dim ArrayTemplates As String() = Nothing
        Dim CountArrayTemplates As Integer = 0
        Dim PutchTemplateFile As String = ""
        Dim psTable As String = "CADASTRAL_POINT"
        Dim idPsTable As ObjectId = Nothing
        Dim arrayOdRec As String(,) = Nothing
        Dim countRows As Integer = 1
        'читаем шаблоны условных знаков
        If IO.Directory.Exists(PutchTables) = True Then
            Try
                For Each fileDWG As String In Directory.GetFiles(PutchTables)
                    'читаем расширение
                    Dim ext As String = Path.GetExtension(fileDWG)
                    If ext Like ".dwg" Then
                        Dim tr1 As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
                        Using tr1
                            Dim acBlkTbl As BlockTable = tr1.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                            If acBlkTbl.Has("PROJECT_POINT") = False Then
                                Dim boolCopyBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(fileDWG, "PROJECT_POINT")
                                If boolCopyBlk = True Then
                                    PutchTemplateFile = fileDWG
                                    tr1.Commit()
                                    Exit For
                                End If
                            End If
                        End Using
                    End If
                Next
            Catch ex As IOException
            Catch ex As UnauthorizedAccessException
            Catch ex As ArgumentException
            End Try
        Else
            MsgBox("Файл шаблона не найден. Макро превано!!!")
            Exit Sub
        End If
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Try
            Using tr
                Dim acBlkTbl As BlockTable
                acBlkTbl = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                      OpenMode.ForWrite)
                Dim dictPropSetDef As Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(acCurDb)
                If dictPropSetDef.Has(psTable, tr) = False Then
                    Dim boolInsPs As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateFile, psTable)
                    If boolInsPs = True Then
                        idPsTable = dictPropSetDef.GetAt(psTable)
                        Dim boolReadFields As Boolean = FuncPropertySet.FuncReadAllValueFieldSetTables(psTable, ArrayODRec)
                    End If
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Выбираем файл
                Dim ArrayDataWrite As String(,) = Nothing
                Dim CountArrayDataWrite As Integer = 0
                Dim opFile As New System.Windows.Forms.OpenFileDialog()
                opFile.InitialDirectory = "c:\"
                opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
                opFile.RestoreDirectory = True
                opFile.ShowDialog()
                Dim fFile1 As String = opFile.FileName
                If IsNothing(fFile1) = True Then Exit Sub
                If fFile1.Trim.Length = 0 Then Exit Sub
                Dim xlApp As Excel.Application = Nothing
                Dim xlBook As Excel.Workbook = Nothing
                Dim xlSheets As Excel.Worksheet = Nothing
                xlApp = CreateObject("Excel.Application")
                xlBook = xlApp.Workbooks.Open(fFile1)
                Dim xlsheet As Excel.Worksheet
                xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
                Dim ArrayData As String(,) = Nothing
                Dim CountArrayData As Integer = 0
                Dim ErrCount As Integer = 0

                Dim boolInsBlk As Boolean = False
                If acBlkTbl.Has("PROJECT_POINT") = True Then
                    boolInsBlk = True
                End If
                For indRows As Integer = 1 To xlsheet.Rows.Count - 1
                    Dim flag As Boolean = False
                    countRows = indRows
                    Dim NumPoint As String = xlsheet.Cells(indRows, 1).value
                    If IsNothing(NumPoint) = True Then
                        NumPoint = ""
                    End If
                    If IsNumeric(NumPoint) = True Then
                        For j As Integer = 0 To 4
                            Try
                                Dim strNumCount As String = xlsheet.Cells(indRows, 1 + j * 4).value
                                If IsNothing(strNumCount) = True Then Continue For
                                If IsNumeric(strNumCount) = False Then Continue For
                                Dim numCount As Integer = Val(strNumCount)

                                Dim strnumPt As String = xlsheet.Cells(indRows, 2 + j * 4).value
                                If IsNothing(strnumPt) = True Then Continue For
                                If IsNumeric(strnumPt) = False Then Continue For
                                Dim numPT As Integer = Val(strnumPt)

                                Dim strX As String = xlsheet.Cells(indRows, 3 + j * 4).value
                                If IsNothing(strX) = True Then Continue For
                                If IsNumeric(strX) = False Then Continue For
                                Dim X As Double = Val(strX) / 100

                                Dim strY As String = xlsheet.Cells(indRows, 4 + j * 4).value
                                If IsNothing(strY) = True Then Continue For
                                If IsNumeric(strY) = False Then Continue For
                                Dim Y As Double = Val(strY) / 100


                                Dim blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(New Point3d(Y, X, 0), "PROJECT_POINT")
                                Dim ArrayAttrib As String(,) = Nothing
                                ReDim ArrayAttrib(1, 0)
                                ArrayAttrib(0, 0) = "NUM_GEOPOINT"
                                ArrayAttrib(1, 0) = numPt
                                Dim boolInsAttarib As Boolean = AcadFunction.FuncWriteAllBlockAttr(blk, ArrayAttrib)
                                'записываем новые  данные
                                If IsNothing(idPsTable) = False Then
                                    If idPsTable.IsNull = False Then
                                        If IsArray(arrayOdRec) = True Then
                                            For i As Integer = 0 To arrayOdRec.GetUpperBound(1)
                                                If arrayOdRec(1, i) Like "Num_geopoint" Then
                                                    arrayOdRec(2, i) = numPt
                                                ElseIf arrayOdRec(1, i) Like "NOTE" Then
                                                    arrayOdRec(2, i) = numCount
                                                End If
                                            Next i
                                        End If
                                        Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(blk, idPsTable)
                                        Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(blk, arrayOdRec)
                                    End If
                                End If
                            Catch ex As System.Exception

                            End Try

                        Next j
                        ErrCount = 0
                    Else
                        ErrCount += 1
                        If ErrCount > 20 Then Exit For
                    End If
                Next indRows
                tr.Commit()
                xlBook.Close()
                xlsheet = Nothing
                xlBook = Nothing
                xlApp.Quit()
                xlApp = Nothing
                MsgBox("Макро завершено. Строка-" & countRows)
            End Using
        Catch ex As Exception
        End Try
    End Sub


    <CommandMethod("ReadC")>
    Public Sub ReadC()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim xlApp As Excel.Application = Nothing
        Dim xlBook As Excel.Workbook = Nothing
        Dim xlSheets As Excel.Worksheet = Nothing
        'xlApp = CreateObject("Excel.Application")
        Try
            xlApp = Marshal.GetActiveObject("Excel.Application")
            Dim icount As Integer = xlApp.Workbooks.Count
            If icount > 0 Then
                For i As Integer = 0 To icount - 1
                    xlBook = xlApp.Workbooks.Item(i)
                    If xlBook.Name Like "Спецификация2" Then
                        xlSheets = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист

                    End If


                Next
            End If
        Catch ex As Exception
        End Try
    End Sub


    '=============================================================================================
    'запись правильных значений PropertyData в солиды на основе данных Excel
    <CommandMethod("Read111")>
    Public Sub Read111()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
        Dim acTrans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Try
            Using acTrans
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim acSSPr As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите поверхность: ")
                If Not (acSSPr.Status = PromptStatus.OK) Then
                    Exit Sub
                End If
                Dim entSurface As Autodesk.Civil.DatabaseServices.TinSurface = acTrans.GetObject(acSSPr.ObjectId, OpenMode.ForWrite)
                Dim colTinSurface As Autodesk.Civil.DatabaseServices.TinSurfaceVertexCollection = entSurface.Vertices

                Dim acSSPr2 As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите полилинию: ")
                If Not (acSSPr2.Status = PromptStatus.OK) Then
                    Exit Sub
                End If
                Dim entPline As Autodesk.AutoCAD.DatabaseServices.Polyline = acTrans.GetObject(acSSPr2.ObjectId, OpenMode.ForWrite)
                Dim boundMin As Point3d = entPline.GeometricExtents.MinPoint
                Dim boundMax As Point3d = entPline.GeometricExtents.MaxPoint

                boundMin = New Point3d(boundMin.X - 50, boundMin.Y - 50, 0)
                boundMax = New Point3d(boundMax.X + 50, boundMax.Y + 50, 0)

                Dim ArrayVertexTinSurface As Point3d() = Nothing
                Dim count As Integer = 0
                For Each TinVertex As TinSurfaceVertex In colTinSurface
                    If TinVertex.Location.X > boundMin.X And TinVertex.Location.Y > boundMin.Y Then
                        If TinVertex.Location.X < boundMax.X And TinVertex.Location.Y < boundMax.Y Then
                            ReDim Preserve ArrayVertexTinSurface(count)
                            ArrayVertexTinSurface(count) = TinVertex.Location
                            count += 1
                        End If
                    End If
                Next

                For i As Integer = 0 To entPline.NumberOfVertices - 1
                    Dim ptPline As Point2d = entPline.GetPoint2dAt(i)
                    For j As Integer = 0 To ArrayVertexTinSurface.Length - 1
                        Dim tempPoint As Point2d = ArrayVertexTinSurface(j).Convert2d(New Plane)
                        If ptPline.GetVectorTo(tempPoint).Length < 0.1 Then
                            entPline.SetPointAt(i, tempPoint)
                            Exit For
                        End If
                    Next j
                Next i



                For i As Integer = 0 To ArrayVertexTinSurface.Length - 1
                    Dim pt3d As Point3d = ArrayVertexTinSurface(i)
                    Dim tempPt As Point3d = entPline.GetClosestPointTo(pt3d, False)
                    If pt3d.Convert2d(New Plane).GetVectorTo(tempPt.Convert2d(New Plane)).Length < 0.8 Then
                        Dim boolInsertVertex As Boolean = AcadFunction.FuncInsertVertexToPolyline(entPline, pt3d)
                    End If
                Next

                'For i As Integer = 0 To entPline.NumberOfVertices - 1
                '    Dim ptPline As Point2d = entPline.GetPoint2dAt(i)
                '    Dim boolDelVertex As Boolean = False
                '    Dim NumbVertex As Integer = 0
                '    For j As Integer = 0 To ArrayVertexTinSurface.Length - 1
                '        Dim tempPoint As Point2d = ArrayVertexTinSurface(j).Convert2d(New Plane)
                '        If ptPline.GetVectorTo(tempPoint).Length = 0 Then
                '            boolDelVertex = False
                '            NumbVertex = i
                '            Exit For
                '        Else
                '            boolDelVertex = True
                '        End If
                '    Next
                '    If boolDelVertex = True Then
                '        entPline.RemoveVertexAt(i)
                '    End If
                'Next
                acTrans.Commit()
            End Using
        Catch ex As Exception
        End Try
    End Sub

    <CommandMethod("iv")>
    Public Sub addplyver()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim chk As Integer = 1
        Dim ets As PromptEntityOptions = New PromptEntityOptions(vbLf & "Select entity to Insert vertex")
        Dim ets1 As PromptEntityResult = ed.GetEntity(ets)

        If ets1.Status = PromptStatus.OK Then
            Dim tr As Transaction = acCurDb.TransactionManager.StartTransaction()

            Using tr
                Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(ets1.ObjectId, OpenMode.ForWrite)
                Dim lws As Polyline = TryCast(obj, Polyline)

                If lws IsNot Nothing Then
                    Dim pnt1 As Point3d = lws.GetClosestPointTo(ets1.PickedPoint, False)
                    Dim tpt As PromptPointOptions = New PromptPointOptions(vbLf & "Pick New vertex location")
                    Dim i As Integer = 0

                    While (i <= (lws.NumberOfVertices - 1))

                        If (lws.GetDistAtPoint(pnt1) < (lws.GetDistAtPoint(lws.GetPoint3dAt(i)))) AndAlso chk <> 0 Then
                            Dim tpt1 As PromptPointResult = ed.GetPoint(tpt)
                            Dim pnt As Point3d = tpt1.Value
                            Dim pnt2 As Point2d = New Point2d(pnt.X, pnt.Y)
                            Dim px As Double = pnt2.X
                            Dim py As Double = pnt2.Y
                            ed.WriteMessage(px.ToString() & "," & (py.ToString()))
                            lws.AddVertexAt(i, pnt2, 0, 0, 0)
                            chk = 0
                            Exit While
                        End If

                        i += 1
                    End While
                End If

                tr.Commit()
            End Using
        End If
    End Sub
    <CommandMethod("iv2")>
    Public Sub addplyver2()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim chk As Integer = 1
        Dim ets As PromptEntityOptions = New PromptEntityOptions(vbLf & "Select entity1:")
        Dim ets1 As PromptEntityResult = ed.GetEntity(ets)


        Dim et As PromptEntityOptions = New PromptEntityOptions(vbLf & "Select Solid3d:")
        Dim et1 As PromptEntityResult = ed.GetEntity(et)



        If ets1.Status = PromptStatus.OK Then
            Dim tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Using tr
                Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(ets1.ObjectId, OpenMode.ForRead)
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(et1.ObjectId, OpenMode.ForRead)
                Dim lws As Circle = TryCast(obj, Circle)

                If lws IsNot Nothing Then
                    Dim pt As Point3d = MathFunction.FuncFindIntersectPointBySolid(ent, lws.Center, New Point3d(0, 0, 0))
                End If

                tr.Commit()
            End Using
        End If
    End Sub

    <CommandMethod("iv3")>
    Public Sub iv3()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim tr As Transaction = acCurDb.TransactionManager.StartTransaction()
        Dim lt As LinetypeTable = CType(tr.GetObject(acCurDb.LinetypeTableId, OpenMode.ForWrite), LinetypeTable)
        Dim acTextStyleTbl As TextStyleTable = tr.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead)
        For Each id As ObjectId In acTextStyleTbl
            Dim acTextStyleRec As TextStyleTableRecord = tr.GetObject(id, OpenMode.ForRead)
            Dim NameStyle As String = acTextStyleRec.Name.Trim
        Next

        Using tr
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        Dim linet As String = acEnt.Linetype
                        Dim idTypeLine As ObjectId = lt.Item(linet)
                        Dim recl As LinetypeTableRecord = idTypeLine.GetObject(OpenMode.ForRead)
                        Dim str As String = AcadFunction.FuncReturnTypeLineCode(recl)
                    End If
                Next
            End If
        End Using
        Dim linetype1 As String = "A,4.0,-2.0"
        Dim linetype2 As String = "A,12.5,-4.0,[""В0П"",STANDARD,S=2.5,U=0.0,X=-2.5,Y=-1.25],-5.0"
        Dim linetype3 As String = "A,3.0,[476-3,GUGK_2005.shx,S=6,R=0,X=0,Y=0.3],3,-4.0,6.0,-4.0"
        Dim createlinetype As LinetypeTableRecord = AcadFunction.FuncCreateTypeLineByCode("new", "desk", linetype3)

        MsgBox("")
    End Sub


End Class
