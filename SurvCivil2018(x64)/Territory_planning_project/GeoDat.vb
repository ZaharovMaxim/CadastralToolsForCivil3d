Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.IO
Imports System
Imports Autodesk.Civil.ApplicationServices
Imports Microsoft.Office.Interop
Public Class GeoDat
    <CommandMethod("GeoDAT")>
    Public Sub GEODAT()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа
        Dim IDArrayPLine As String(,) = Nothing 'массив для хранения глобальных ID участков и их Hangle (0-Handle , 1-ID)
        Dim CountPline As Integer = 0 'счетчик участков
        Dim ArrayParcel As String(,) = Nothing 'массив для хранения контуров на данном участке
        Dim counter1(,) As String = Nothing 'массив для контуров участка (1 в случае если участок простой)
        Dim Koordinate As String(,) = Nothing 'массив для координат участка
        Dim ArrayOstrov As String() = Nothing ' массив для хранения островков на выбранном контуре
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'создаем класс с новой формой
        Dim GeoDat1 As FormGeoDat = New FormGeoDat()
        GeoDat1.ShowDialog()
        If GeoDat1.BoolForm = 0 Then Exit Sub
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ссылка на чертеж .dwg
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        'открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Try
            Using tr
                '==============================================================================
                'создаем набор точек со всего чертежа
                Dim GeoPointCoolection As ObjectIdCollection = New ObjectIdCollection
                Dim sourceBtr As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead, False)
                For Each id As ObjectId In sourceBtr
                    If id.IsNull = False Then
                        Dim entObject As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                        Dim idLr As ObjectId = entObject.LayerId
                        Dim lyrTemp As LayerTableRecord = idLr.GetObject(OpenMode.ForRead)
                        If lyrTemp.IsFrozen = True Then
                            Continue For
                        End If
                        If lyrTemp.IsOff = True Then
                            Continue For
                        End If
                        'блок
                        If TypeOf entObject Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                            Dim blkPt As BlockReference = entObject
                            Dim nameUserBlk As String = AcadFunction.FuncReadNameBlk(blkPt)
                            If nameUserBlk Like "PROJECT_POINT" Then
                                GeoPointCoolection.Add(id)
                            ElseIf nameUserBlk Like "TEMP_POINT" Then
                                GeoPointCoolection.Add(id)
                            End If
                        End If
                    End If
                Next
                '================================================================================
                'выбираем либо полилинии либо участки либо подучастки
                Dim resEntity As PromptEntityResult = Nothing
                Dim optEntity As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите полилинию, участок или подучасток: ")
                optEntity.AllowNone = True
                resEntity = ed.GetEntity(optEntity)
                If resEntity.Status = PromptStatus.OK Then
                    Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(resEntity.ObjectId, OpenMode.ForRead)
                    If TypeOf obj Is Polyline Then
                        'если полилиния не преобразована в участок
                        ReDim Preserve counter1(1, CountPline)
                        counter1(0, CountPline) = obj.Handle.ToString
                        CountPline = 1
                    End If
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'записывает координаты участка в массив
                'тело программы
                Dim countCoord As Integer = 0 'счетчик координат для массима Koordinate
                Dim IDKounter As String = "" 'геоидентификатор контура
                Dim ext As TypedValue() = Nothing 'массив для расширенных данных
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity 'переменная для полилиний
                Dim Pline As Polyline
                Dim Point1(2) As String 'массив для характеристики точек(имя, Mt, описание)
                Dim VertCoord As Point2d = New Point2d(0, 0)
                Dim FindPoint As ObjectId = Nothing 'id для блока с атрибутом
                'перебираем все контура входящие в участок
                For j As Integer = 0 To counter1.GetUpperBound(1)
                    ed.WriteMessage("Обрабатывается участок " & j + 1)
                    'получаем доступ к полилинии для извлечения координат
                    Dim HandlePLine As String = counter1(0, j)
                    Dim ln As Long = Convert.ToInt64(HandlePLine, 16)
                    Dim hn As Handle = New Handle(ln)
                    Dim id As ObjectId = db.GetObjectId(False, hn, 0)
                    ent = tr.GetObject(id, OpenMode.ForWrite)
                    Pline = ent
                    '==================================================================
                    'определяем количество вершин в контуре
                    Dim vn As Integer = Pline.NumberOfVertices
                    ReDim Preserve Koordinate(5, countCoord)
                    Koordinate(1, countCoord) = "" 'пустая строка в таблице
                    countCoord += 1
                    'последовательно перебираем все вершины контура для записи в массив координат поворотных точек
                    For i1 As Integer = 0 To vn - 1
                        'запоминаем координаты вершины полилинии
                        VertCoord = Pline.GetPoint2dAt(i1)
                        'записываем характеристики в массив
                        ReDim Preserve Koordinate(5, countCoord)
                        Koordinate(0, countCoord) = IDKounter 'Handle полилинии на которой точка
                        Koordinate(2, countCoord) = VertCoord.X 'кооордината Х
                        Koordinate(3, countCoord) = VertCoord.Y 'координата Y
                        'находим номер для первой точки
                        Dim findPoint1 As Boolean = False
                        If GeoPointCoolection.Count > 0 Then
                            Dim FindPointID As ObjectId = FuncCadastr.FuncFindGeoPoint(GeoPointCoolection, VertCoord, 0.05)
                            If IsNothing(FindPointID) = False Then
                                If FindPointID.IsNull = False Then
                                    Dim entBlk As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(FindPointID, OpenMode.ForRead)
                                    If TypeOf entBlk Is BlockReference Then
                                        Dim blk As BlockReference = entBlk
                                        Dim ArrayAttrib As String() = Nothing
                                        Dim FindPointGeo As Boolean = FuncCadastr.FuncReturnBlkArtibuteGeoPoint(blk, ArrayAttrib)
                                        If IsArray(ArrayAttrib) = True Then
                                            Koordinate(1, countCoord) = ArrayAttrib(0) 'номер точки 1, 2 ...
                                            findPoint1 = True
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        If findPoint1 = False Then
                            Koordinate(1, countCoord) = i1 + 1 'номер точки 1, 2 ...
                        End If
                        countCoord += 1
                    Next
                    'возвращаемся на первую вершину и записываем ее в массив
                    VertCoord = Pline.GetPoint2dAt(0)
                    'записываем характеристики в массив
                    ReDim Preserve Koordinate(5, countCoord)
                    Koordinate(0, countCoord) = IDKounter 'Handle полилинии на которой точка
                    Koordinate(2, countCoord) = VertCoord.X 'кооордината Х
                        Koordinate(3, countCoord) = VertCoord.Y 'координата Y
                        'находим номер для первой точки
                        Dim findPoint2 As Boolean = False
                        If GeoPointCoolection.Count > 0 Then
                            Dim FindPointID As ObjectId = FuncCadastr.FuncFindGeoPoint(GeoPointCoolection, VertCoord, 0.05)
                            If IsNothing(FindPointID) = False Then
                                If FindPointID.IsNull = False Then
                                    Dim entBlk As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(FindPointID, OpenMode.ForRead)
                                    If TypeOf entBlk Is BlockReference Then
                                        Dim blk As BlockReference = entBlk
                                        Dim ArrayAttrib As String() = Nothing
                                        Dim FindPointGeo As Boolean = FuncCadastr.FuncReturnBlkArtibuteGeoPoint(blk, ArrayAttrib)
                                        If IsArray(ArrayAttrib) = True Then
                                            Koordinate(1, countCoord) = ArrayAttrib(0) 'номер точки 1, 2 ...
                                            findPoint2 = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If findPoint2 = False Then
                            Koordinate(1, countCoord) = 1 'номер точки 1, 2 ...
                        End If

                Next j
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim AngleDist As String(,) = Nothing 'массив для конечных данных
                Dim countArray As Integer = 0 'счетчик
                For count As Integer = 0 To Koordinate.GetUpperBound(1) - 1
                    Dim Rumb As String = ""
                    Dim NumPt1 As String = Koordinate(1, count)
                    Dim NumPt2 As String = Koordinate(1, count + 1)
                    If NumPt1 Like "" Then
                        ReDim Preserve AngleDist(2, countArray)
                        AngleDist(0, countArray) = ""
                        AngleDist(1, countArray) = ""
                        AngleDist(2, countArray) = ""
                        countArray += 1
                    ElseIf NumPt2 Like "" Then

                    Else
                        Dim Ptx1 As Double = Koordinate(2, count)
                        Dim Ptx2 As Double = Koordinate(2, count + 1)
                        Dim Pty1 As Double = Koordinate(3, count)
                        Dim Pty2 As Double = Koordinate(3, count + 1)
                        Dim Pt1 As Point2d = New Point2d(Ptx1, Pty1)
                        Dim Pt2 As Point2d = New Point2d(Ptx2, Pty2)
                        Dim vk As Vector2d = Pt1.GetVectorTo(Pt2)
                        Dim dist As Double = vk.Length 'находим длину линии
                        If GeoDat1.RadioButton8.Checked = True Then
                            dist = dist / 1000
                        End If
                        Dim RoundznDist As Integer = GeoDat1.NumericUpDown1.Value
                        dist = Math.Round(dist, RoundznDist)
                        'форматирование углов
                        Dim RoundznAngle As Integer = GeoDat1.NumericUpDown2.Value
                        Dim ddmmssZn As Integer = 1
                        Dim AngleFormat As String = "0.0000"
                        If GeoDat1.RadioButton4.Checked = True Then
                            ddmmssZn = 0
                        ElseIf GeoDat1.RadioButton5.Checked = True Then
                            ddmmssZn = 1
                        Else
                            ddmmssZn = 2
                        End If
                        'находим угол
                        If GeoDat1.RadioButton1.Checked = True Then
                            AngleFormat = FuncFormatZN.FuncAngleDirectFormat(Pt1, Pt2, ddmmssZn, RoundznAngle)
                        ElseIf GeoDat1.RadioButton2.Checked = True Then
                            AngleFormat = FuncFormatZN.FuncAngleRumbFormat(Pt1, Pt2, ddmmssZn, RoundznAngle)
                        End If
                        ReDim Preserve AngleDist(2, countArray)
                        AngleDist(0, countArray) = NumPt1 & "-" & NumPt2
                        AngleDist(1, countArray) = AngleFormat
                        AngleDist(2, countArray) = dist
                        countArray += 1
                    End If
                Next
                '========================================================================================================
                'вставка таблицы
                If GeoDat1.RadioButton9.Checked = True Then
                    Dim insPoint As Point3d = New Point3d(0, 0, 0)
                    Dim ppr As PromptPointResult = ed.GetPoint(vbLf & "Укажите точку вставки таблицы: ")
                    If ppr.Status = PromptStatus.OK Then
                        insPoint = ppr.Value
                    End If
                    'заводим нужную таблицу
                    Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForWrite)
                    Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim tb As Autodesk.AutoCAD.DatabaseServices.Table = New Autodesk.AutoCAD.DatabaseServices.Table()
                    tb.TableStyle = db.Tablestyle
                    tb.SetSize(3, 3)
                    tb.SetRowHeight(4)
                    tb.Rows(1).TextHeight = 2
                    tb.SetColumnWidth(20)
                    tb.Position = insPoint
                    'заголовок
                    'заголовок таблицы
                    tb.Cells(1, 0).Value = "Номер"
                    tb.Cells(1, 1).Value = "Румбы"
                    tb.Cells(1, 2).Value = "Расстояние, м"
                    '========================================================================
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    For count As Integer = 1 To AngleDist.GetUpperBound(1)
                        tb.InsertRows(count + 1, 1, 1)
                        tb.Rows(count + 1).TextHeight = 2
                        'номер точки
                        tb.Cells(count + 1, 0).Value = AngleDist(0, count)
                        'румб
                        tb.Cells(count + 1, 1).Value = AngleDist(1, count)
                        'длина линии
                        tb.Cells(count + 1, 2).Value = AngleDist(2, count)
                    Next count
                    ms.UpgradeOpen()
                    ms.AppendEntity(tb)
                    tr.AddNewlyCreatedDBObject(tb, True)
                ElseIf GeoDat1.RadioButton10.Checked = True Then
                    'выбираем либо полилинии либо участки либо подучастки
                    Dim resEntityTb As PromptEntityResult = Nothing
                    Dim optEntityTb As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите таблицу: ")
                    optEntityTb.AllowNone = True
                    resEntityTb = ed.GetEntity(optEntityTb)
                    If resEntityTb.Status = PromptStatus.OK Then
                        Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(resEntityTb.ObjectId, OpenMode.ForWrite)
                        If TypeOf obj Is Autodesk.AutoCAD.DatabaseServices.Table Then
                            Dim tb As Autodesk.AutoCAD.DatabaseServices.Table = obj
                            Dim countRows As Integer = tb.Rows.Count
                            Dim countCol As Integer = tb.Columns.Count
                            If countCol < 3 Then
                                MsgBox("Выберите таблицу с количеством столбцов более 2. Макро прервано!")
                                Exit Sub
                            Else
                                For count As Integer = 1 To AngleDist.GetUpperBound(1)
                                    tb.InsertRows(countRows + count - 1, 1, 1)
                                    tb.Rows(countRows + count - 1).TextHeight = 1
                                    'номер точки
                                    tb.Cells(countRows + count - 1, 0).Value = AngleDist(0, count)
                                    'румб
                                    tb.Cells(countRows + count - 1, 1).Value = AngleDist(1, count)
                                    'длина линии
                                    tb.Cells(countRows + count - 1, 2).Value = AngleDist(2, count)
                                Next count
                            End If
                        End If
                    End If
                ElseIf GeoDat1.RadioButton11.Checked = True Then
                    'запускаем Excel
                    Dim xlApp As Excel.Application = Nothing
                    Dim xlBook As Excel.Workbook = Nothing
                    Dim xlSheets As Excel.Worksheet = Nothing
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    Try
                        xlApp = CreateObject("Excel.Application")
                        xlBook = xlApp.Workbooks.Add
                        xlApp.Visible = True
                        Dim file1 As String = xlApp.GetSaveAsFilename("Таблица координат", "Excel Files (*.xls), .xls")
                        If file1 = "False" Then 'если ошибка то выходим из программы
                            xlBook.Close()
                            xlApp.Quit()
                            xlSheets = Nothing
                            xlBook = Nothing
                            xlApp = Nothing
                            Exit Sub
                        End If
                        Dim xlsheet As Excel.Worksheet
                        xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
                        xlsheet.Activate()
                        Dim cell1 As Excel.Range 'первая ячейка в строке
                        xlsheet.Cells(1, 1).value = "Номер точки" 'обозначение точки
                        If GeoDat1.RadioButton1.Checked = True Then
                            xlsheet.Cells(1, 2).value = "Дирекционный угол"  'угол
                        Else
                            xlsheet.Cells(1, 2).value = "Румб"  'угол
                        End If
                        xlsheet.Cells(1, 3).value = "Расстояние, м"
                        For i As Integer = 1 To AngleDist.GetUpperBound(1)
                            xlsheet.Cells(i + 1, 1).value = AngleDist(0, i)  'обозначение точки
                            xlsheet.Cells(i + 1, 2).value = AngleDist(1, i)  'угол
                            xlsheet.Cells(i + 1, 3).value = AngleDist(2, i)  'обозначение точки
                        Next i
                    Catch
                        xlSheets = Nothing
                        xlBook = Nothing
                        xlApp = Nothing
                    End Try
                    xlSheets = Nothing
                    xlBook = Nothing
                    xlApp = Nothing
                ElseIf GeoDat1.RadioButton12.Checked = True Then
                    'запускаем Excel
                    Dim wdApp As Word.Application = Nothing
                    Dim wdDoc As Word.Document = Nothing
                    Try
                        wdApp = CreateObject("Word.Application")
                        wdDoc = wdApp.Documents.Add()
                        wdApp.Visible = True
                        Dim tableLocation As Word.Range = wdDoc.Range(Start:=0, End:=0)
                        Dim newTable As Word.Table = wdDoc.Tables.Add(tableLocation, 2, 3)
                        With newTable.Cell(1, 1).Range
                            .Text = "Номер"
                            .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                        End With
                        With newTable.Cell(1, 2).Range
                            If GeoDat1.RadioButton1.Checked = True Then
                                .Text = "Дирекционный угол"
                            Else
                                .Text = "Румб"
                            End If
                            .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                        End With
                        With newTable.Cell(1, 3).Range
                            .Text = "Расстояние, м"
                            .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                        End With
                        For i As Integer = 1 To AngleDist.GetUpperBound(1)
                            newTable.Rows.Add()
                            With newTable.Cell(1 + i, 1).Range
                                .Text = AngleDist(0, i)
                                .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            End With
                            With newTable.Cell(1 + i, 2).Range
                                .Text = AngleDist(1, i)
                                .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            End With
                            With newTable.Cell(1 + i, 3).Range
                                .Text = AngleDist(2, i)
                                .ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            End With
                        Next i
                    Catch
                        wdDoc = Nothing
                        wdApp = Nothing
                    End Try
                    wdDoc = Nothing
                    wdApp = Nothing
                End If
                tr.Commit()
            End Using
        Catch

        End Try
    End Sub
End Class
