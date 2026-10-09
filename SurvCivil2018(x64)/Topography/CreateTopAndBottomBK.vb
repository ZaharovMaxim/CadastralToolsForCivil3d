Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.IO
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Public Class CreateTopAndBottomBK
    '=============================================================================================
    'формирует отображение верха и низа бордюрного камня
    <CommandMethod("CreateTopAndBottomBK2")>
    Public Sub CreateTopAndBottomBK2()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim FormTopAndBottomBK As TopAndBottomBK = New TopAndBottomBK
        FormTopAndBottomBK.ShowDialog()
        If FormTopAndBottomBK.boolShow = False Then
            Exit Sub
        End If
        Dim PutchTemplate As String = FormTopAndBottomBK.ComboBox1.Tag & FormTopAndBottomBK.ComboBox1.Text
        If File.Exists(PutchTemplate) = False Then
            MsgBox("Файл шаблона не найден!!! Макро прервано.")
            Exit Sub
        End If
        '==========================================================================================================================================
        'начальные установки
        '==========================================================================================================================================
        'проверяем классификацию точек
        '==========================================================================================================================================
        Dim ArrayPropBK As String(,) = Nothing 'отметка высот бордюрного камня
        Dim DeskCOGOBK As String = "Отметка высот бордюрного камня"
        Dim NameTable As String = "Отметки высот"
        Dim NameGroupPointBK As String = ""
        Dim NameTableODBK As String = ""
        Dim boolPropCogoPoint As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplate, DeskCOGOBK, ArrayPropBK, "ТочкаCOGO", NameTable)
        If IsArray(ArrayPropBK) = True Then
            For k As Integer = 0 To ArrayPropBK.GetUpperBound(1)
                Dim NameField As String = ArrayPropBK(0, k)
                If NameField.Trim Like "Полевой код" Then
                    DeskCOGOBK = ArrayPropBK(1, k)
                ElseIf NameField.Trim Like "Группа точек" Then
                    NameGroupPointBK = ArrayPropBK(1, k)
                ElseIf NameField.Trim Like "Наименование таблицы данных (OD)" Then
                    NameTableODBK = ArrayPropBK(1, k)
                End If
            Next k
        Else
            MsgBox("Характеристики блока не найдены. Макро прервано!!!")
            Exit Sub
        End If
        '======================================================================================================
        Dim ArrayPropVBK As String(,) = Nothing 'отметка высот бордюрного камня
        Dim DeskCOGOVBK As String = "Отметка высот верха бордюрного камня"
        Dim NameGroupPointVBK As String = ""
        Dim NameTableODVBK As String = ""
        Dim boolPropCogoPointVBK As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplate, DeskCOGOVBK, ArrayPropVBK, "ТочкаCOGO", NameTable)
        If IsArray(ArrayPropVBK) = True Then
            For k As Integer = 0 To ArrayPropVBK.GetUpperBound(1)
                Dim NameField As String = ArrayPropVBK(0, k)
                If NameField.Trim Like "Полевой код" Then
                    DeskCOGOVBK = ArrayPropVBK(1, k)
                ElseIf NameField.Trim Like "Группа точек" Then
                    NameGroupPointVBK = ArrayPropVBK(1, k)
                ElseIf NameField.Trim Like "Наименование таблицы данных (OD)" Then
                    NameTableODVBK = ArrayPropVBK(1, k)
                End If
            Next k
        Else
            MsgBox("Характеристики блока не найдены. Макро прервано!!!")
            Exit Sub
        End If
        Using acTrans As Transaction = acDoc.TransactionManager.StartTransaction()
            '========================================================================================================
            'копируем группу точек для точек бордюрного камня
            If CivilDoc.PointGroups.Contains(NameGroupPointBK) = False Then
                Dim boolCopyGroupPoint As Boolean = AcadFunction.FuncCopyGroupPointsDestDraw(PutchTemplate, NameGroupPointBK)
                If Not (NameGroupPointBK Like NameGroupPointVBK) Then
                    If CivilDoc.PointGroups.Contains(NameGroupPointVBK) = False Then
                        Dim boolCopyGroupPoint1 As Boolean = AcadFunction.FuncCopyGroupPointsDestDraw(PutchTemplate, NameGroupPointVBK)
                    End If
                End If
            End If
            Dim IDPointGroupBK As ObjectId = CivilDoc.PointGroups.Item(NameGroupPointBK)
            Dim PointGroupBK As PointGroup = Nothing
            Dim UDP As UDPDouble = Nothing
            If IsNothing(IDPointGroupBK) = False Then
                If IDPointGroupBK.IsNull = False Then
                    PointGroupBK = acTrans.GetObject(IDPointGroupBK, OpenMode.ForWrite)
                    Dim UDPClass As String = PointGroupBK.UDPClassificationName
                    For Each IdUdp As UDP In CivilDoc.PointUDPs
                        If IsNothing(IdUdp) = False Then
                            Try
                                Dim nameClass As String = IdUdp.ClassificationName
                                If nameClass Like UDPClass Then
                                    If TypeOf IdUdp Is UDPDouble Then
                                        UDP = IdUdp
                                        Exit For
                                    End If
                                End If
                            Catch ex As System.NullReferenceException
                            End Try
                        End If
                    Next
                End If
            End If

            Dim IDPointGroupVBK As ObjectId = CivilDoc.PointGroups.Item(NameGroupPointVBK)
            Dim PointGroupVBK As PointGroup = Nothing
            If IsNothing(IDPointGroupVBK) = False Then
                If IDPointGroupVBK.IsNull = False Then
                    PointGroupVBK = acTrans.GetObject(IDPointGroupVBK, OpenMode.ForWrite)
                End If
            End If

            '========================================================================================================
            'копируем стили
            Dim boolCopyDeskKeyBK As Boolean = False
            Dim boolCopyDeskKeyVBK As Boolean = False
            Dim pdKeySetIds As ObjectIdCollection = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(acCurDb).SearchOrder
            If pdKeySetIds.Count > 0 Then
                For Each idKeySet As ObjectId In pdKeySetIds
                    Dim DeskKeySet As PointDescriptionKeySet = acTrans.GetObject(idKeySet, OpenMode.ForRead)
                    Dim pdKeyIds As ObjectIdCollection = DeskKeySet.GetPointDescriptionKeyIds
                    For Each idDeskKey As ObjectId In pdKeyIds
                        If idDeskKey.IsNull = False Then
                            Dim DeskKey As PointDescriptionKey = acTrans.GetObject(idDeskKey, OpenMode.ForRead)
                            If DeskKey.Code Like DeskCOGOBK Then
                                boolCopyDeskKeyBK = True
                            ElseIf DeskKey.Code Like DeskCOGOVBK Then
                                boolCopyDeskKeyVBK = True
                            End If
                        End If
                    Next
                Next
            End If
            If boolCopyDeskKeyBK = False Then
                boolCopyDeskKeyBK = AcadFunction.FuncCopyDeskKeyDestDraw(PutchTemplate, DeskCOGOBK)
            End If
            If boolCopyDeskKeyVBK = False Then
                boolCopyDeskKeyVBK = AcadFunction.FuncCopyDeskKeyDestDraw(PutchTemplate, DeskCOGOVBK)
            End If
            '========================================================================================================
            'копируем таблицу PSSet
            Dim dictPropSetDef = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(acCurDb)
            If dictPropSetDef.Has(NameTableODVBK, acTrans) = False Then
                Dim boolCopyTablePS As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplate, NameTableODBK)
            End If
            If Not (NameGroupPointVBK Like NameGroupPointBK) Then
                If dictPropSetDef.Has(NameTableODVBK, acTrans) = False Then
                    Dim boolCopyTablePS1 As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplate, NameTableODVBK)
                End If
            End If
            Dim IDTablePS As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(NameTableODBK)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'значение по умолчанию для тех точек верха которых небыли найдены
            Dim BKV0 As Double = FormTopAndBottomBK.NumericUpDown1.Value
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ручной режим
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If FormTopAndBottomBK.RadioButton2.Checked Then
                Dim resBlk As PromptEntityResult = Nothing
                Dim optBlk As PromptEntityOptions = New PromptEntityOptions(vbLf & "Укажите точку COGO низа бордюрного камня: ")
                optBlk.AllowNone = True
                Dim optBlk1 As PromptEntityOptions = New PromptEntityOptions(vbLf & "Укажите точку COGO верха бордюрного камня: ")
                optBlk1.AllowNone = True
                Do
                    Dim tr1 As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
                    Using tr1
                        Try
                            Dim CoGoPoint1 As CogoPoint = Nothing
                            Dim BKV As Double = 0
                            Dim RotateMarker As Double = 0
                            resBlk = ed.GetEntity(optBlk)
                            Dim obj1 As Autodesk.AutoCAD.DatabaseServices.DBObject = Nothing
                            If resBlk.Status = PromptStatus.OK Then
                                obj1 = tr1.GetObject(resBlk.ObjectId, OpenMode.ForWrite)
                                If TypeOf obj1 Is CogoPoint Then
                                    CoGoPoint1 = obj1
                                    Dim RDesk As String = CoGoPoint1.RawDescription 'читаем ее описание
                                End If
                            Else
                                Exit Do
                            End If
                            'выбираем вторую точку верха бордюра
                            Dim CoGoPoint2 As CogoPoint = Nothing
                            resBlk = ed.GetEntity(optBlk1)
                            Dim obj2 As Autodesk.AutoCAD.DatabaseServices.DBObject = Nothing
                            If resBlk.Status = PromptStatus.OK Then
                                obj2 = tr1.GetObject(resBlk.ObjectId, OpenMode.ForWrite)
                                If TypeOf obj2 Is CogoPoint Then
                                    CoGoPoint2 = obj2
                                End If
                            End If
                            'вычисляем отметку верха
                            Dim ElevPoint1 As Double = -1
                            If CoGoPoint1.IsPlanar = False Then
                                ElevPoint1 = CoGoPoint1.Elevation
                            Else
                                MsgBox("Выбранная точка не имеет высотной отметки!!! Макро прервано.")
                                GoTo line6
                            End If
                            Dim ElevPoint2 As Double = -1
                            If CoGoPoint2.IsPlanar = False Then
                                ElevPoint2 = CoGoPoint2.Elevation
                            Else
                                MsgBox("Выбранная точка не имеет высотной отметки!!! Макро прервано.")
                                GoTo line6
                            End If
                            If ElevPoint1 <= ElevPoint2 Then
                                CoGoPoint2.RawDescription = DeskCOGOVBK
                                Dim h As Double = Val(ElevPoint2 - ElevPoint1)
                                Dim HPoint As Double = ElevPoint1 + h
                                CoGoPoint1.RawDescription = DeskCOGOBK
                                CoGoPoint1.ApplyDescriptionKeys() 'применяем ключи описатели
                                PointGroupBK.Update() 'обновляем группу
                                CoGoPoint1.SetUDPValue(UDP, h) 'присваиваем ключ
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'формируем массив 
                                Dim ArrayWriteOD2 As String(,) = Nothing
                                ReDim ArrayWriteOD2(2, 1)
                                ArrayWriteOD2(0, 0) = NameTableODBK
                                ArrayWriteOD2(1, 0) = "LABEL"
                                ArrayWriteOD2(2, 0) = CoGoPoint1.FullDescription
                                ArrayWriteOD2(0, 1) = NameTableODBK
                                ArrayWriteOD2(1, 1) = UDP.Name
                                ArrayWriteOD2(2, 1) = Math.Round(HPoint, 2)
                                'записываем OD данные
                                Dim boolBK As Boolean = FuncPropertySet.FuncFindPSetTable(CoGoPoint1, "BORDER_POINT")
                                If boolBK = True Then
                                    Dim boolBRWrite As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(CoGoPoint1, ArrayWriteOD2)
                                Else
                                    Dim boolAddTable As Boolean = FuncPropertySet.FuncAddDataSetProperties(CoGoPoint1, IDTablePS)
                                    Dim boolAddData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(CoGoPoint1, ArrayWriteOD2)
                                End If
                            End If
line6:
                        Catch ex As Exception
                            ed.WriteMessage("Ошибка. Макро прервано!!!" & vbLf)
                        End Try
                        tr1.Commit()
                    End Using
                Loop
                '=============================================================================================================================
                'автоматическое заполнение
                '============================================================================================================================
            Else
                'выбираем польлинию для определения угла поворота метки
                Dim ArrayPointXY As Double(,) = Nothing 'массив для координат точек
                Dim SelectPline As PromptEntityResult = ed.GetEntity("Выберите полилинию определяющую низ бортового камня:  ")
                If SelectPline.Status = PromptStatus.OK Then
                    Dim acEnt1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(SelectPline.ObjectId, OpenMode.ForWrite)
                    If TypeOf acEnt1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                        Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt1
                        ReDim ArrayPointXY(1, Pline.NumberOfVertices)
                        For i As Integer = 0 To Pline.NumberOfVertices - 1
                            Dim PointXY As Point2d = Pline.GetPoint2dAt(i)
                            ArrayPointXY(0, i) = PointXY.X
                            ArrayPointXY(1, i) = PointXY.Y
                        Next i
                    ElseIf TypeOf acEnt1 Is Polyline2d Then
                        Dim Pline As Polyline2d = acEnt1
                        Dim CountVert As Integer = 0
                        For Each VertID As ObjectId In Pline
                            Dim v2d As Vertex2d = acTrans.GetObject(VertID, OpenMode.ForRead)
                            ReDim Preserve ArrayPointXY(1, CountVert)
                            ArrayPointXY(0, CountVert) = v2d.Position.X
                            ArrayPointXY(1, CountVert) = v2d.Position.Y
                        Next
                    End If
                End If
                'выбор полилинии
                If SelectPline.Status = PromptStatus.OK Then
                    Dim ArrayPoint As Integer() = Nothing
                    Dim acEnt1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(SelectPline.ObjectId, OpenMode.ForWrite)
                    If TypeOf acEnt1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                        Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt1
                        'создаем габаритный контейнер полилинии для фильтрации точек
                        Dim MinCoord As Point3d = Pline.GeometricExtents.MinPoint
                        Dim MaxCoord As Point3d = Pline.GeometricExtents.MaxPoint
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
                        'находим максимальное расстояние между точками
                        Dim resPoint As PromptPointResult = Nothing
                        Dim optPoint As PromptPointOptions = New PromptPointOptions(vbLf & "Укажите максимальное расстояние до COGO точки (первая точка): ")
                        resPoint = ed.GetPoint(optPoint)
                        Dim Point1 As Point2d = resPoint.Value.Convert2d(New Plane)

                        Dim optPoint1 As PromptPointOptions = New PromptPointOptions(vbLf & "Укажите максимальное расстояние до COGO точки (вторая точка): ")
                        resPoint = ed.GetPoint(optPoint1)
                        Dim Point2 As Point2d = resPoint.Value.Convert2d(New Plane)

                        Dim Lenght As Double = Point1.GetVectorTo(Point2).Length
                        Lenght = Math.Round(Lenght, 2)
                        'начинаем перебор вершин полилинии
                        For i As Integer = 0 To ArrayPointXY.GetUpperBound(1)
                            Try
                                Dim h As Double = 0
                                Dim NumPoint As Integer = 0
                                Dim RotateMarker As Double = 0
                                Dim PointCogoBKN As CogoPoint = Nothing
                                Dim pointCogoBKV As CogoPoint = Nothing
                                Dim Point As Point3d = New Point3d(ArrayPointXY(0, i), ArrayPointXY(1, i), 0)
                                'ищем точку в узле полилинии 
                                If IsArray(ArrayPoint) = True Then
                                    PointCogoBKN = FuncSurvey.FuncFindCogoPointByCoordinates(Point, ArrayPoint, 0.05)
                                    'шщем ближайшую точку к исходной
                                    If Not (PointCogoBKN Is Nothing) Then
                                        pointCogoBKV = FuncSurvey.FuncFindCogoPointByDistance(PointCogoBKN.PointNumber, Lenght)
                                    Else
                                        Continue For
                                    End If
                                    If Not (pointCogoBKV Is Nothing) Then
                                        Dim ElevPoint1 As Double = -1
                                        Dim ElevPoint2 As Double = -2
                                        If PointCogoBKN.IsPlanar = False Then
                                            ElevPoint1 = PointCogoBKN.Elevation
                                        End If
                                        If pointCogoBKV.IsPlanar = False Then
                                            ElevPoint2 = pointCogoBKV.Elevation
                                        End If
                                        If ElevPoint1 <> -1 And ElevPoint2 <> -2 Then
                                            If ElevPoint1 < ElevPoint2 Then
                                                h = Math.Round((ElevPoint2 - ElevPoint1), 2)
                                                If Not (ArrayPointXY Is Nothing) And FormTopAndBottomBK.CheckBox1.Checked = True Then
                                                    RotateMarker = FuncSurvey.FuncCalcRotateMarkerCoGoPoint(ArrayPointXY, PointCogoBKN.Easting, PointCogoBKN.Northing)
                                                    PointCogoBKN.LabelRotation = RotateMarker
                                                End If
                                                PointCogoBKN.RawDescription = DeskCOGOBK
                                                pointCogoBKV.RawDescription = DeskCOGOVBK
                                                PointCogoBKN.ApplyDescriptionKeys()
                                                pointCogoBKV.ApplyDescriptionKeys()
                                                PointGroupBK.Update()
                                                NumPoint = PointCogoBKN.PointNumber
                                                PointCogoBKN.SetUDPValue(UDP, h)
                                                '==========================================================================
                                                'записываем расширенные данные
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                'формируем массив 
                                                Dim ArrayWriteOD2 As String(,) = Nothing
                                                ReDim ArrayWriteOD2(2, 1)
                                                ArrayWriteOD2(0, 0) = NameTableODBK
                                                ArrayWriteOD2(1, 0) = "LABEL"
                                                ArrayWriteOD2(2, 0) = PointCogoBKN.FullDescription
                                                ArrayWriteOD2(0, 1) = NameTableODBK
                                                ArrayWriteOD2(1, 1) = UDP.Name
                                                ArrayWriteOD2(2, 1) = Math.Round(ElevPoint2, 2)
                                                'записываем OD данные
                                                Dim boolBK As Boolean = FuncPropertySet.FuncFindPSetTable(PointCogoBKN, NameTableODBK)
                                                If boolBK = True Then
                                                    Dim boolBRWrite As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(PointCogoBKN, ArrayWriteOD2)
                                                Else
                                                    Dim boolAddTable As Boolean = FuncPropertySet.FuncAddDataSetProperties(PointCogoBKN, IDTablePS)
                                                    Dim boolAddData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(PointCogoBKN, ArrayWriteOD2)
                                                End If
                                            End If
                                        End If
                                    Else
                                        Dim ElevPoint1 As Double = -1
                                        If PointCogoBKN.IsPlanar = False Then
                                            ElevPoint1 = PointCogoBKN.Elevation
                                            h = Math.Round(BKV0, 2)
                                            If Not (ArrayPointXY Is Nothing) And FormTopAndBottomBK.CheckBox1.Checked = True Then
                                                RotateMarker = FuncSurvey.FuncCalcRotateMarkerCoGoPoint(ArrayPointXY, PointCogoBKN.Easting, PointCogoBKN.Northing)
                                                PointCogoBKN.LabelRotation = RotateMarker
                                            End If
                                            PointCogoBKN.RawDescription = DeskCOGOBK
                                            PointCogoBKN.ApplyDescriptionKeys()
                                            PointGroupBK.Update()
                                            NumPoint = PointCogoBKN.PointNumber
                                            PointCogoBKN.SetUDPValue(UDP, h)
                                            '==========================================================================
                                            'записываем расширенные данные
                                            Dim ArrayWriteOD2 As String(,) = Nothing
                                            ReDim ArrayWriteOD2(2, 1)
                                            ArrayWriteOD2(0, 0) = NameTableODBK
                                            ArrayWriteOD2(1, 0) = "LABEL"
                                            ArrayWriteOD2(2, 0) = PointCogoBKN.FullDescription
                                            ArrayWriteOD2(0, 1) = NameTableODBK
                                            ArrayWriteOD2(1, 1) = UDP.Name
                                            ArrayWriteOD2(2, 1) = Math.Round(ElevPoint1 + h, 2)
                                            'записываем OD данные
                                            Dim boolBK As Boolean = FuncPropertySet.FuncFindPSetTable(PointCogoBKN, NameTableODBK)
                                            If boolBK = True Then
                                                Dim boolBRWrite As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(PointCogoBKN, ArrayWriteOD2)
                                            Else
                                                Dim boolAddTable As Boolean = FuncPropertySet.FuncAddDataSetProperties(PointCogoBKN, IDTablePS)
                                                Dim boolAddData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(PointCogoBKN, ArrayWriteOD2)
                                            End If
                                        End If
                                    End If
                                End If

line7:
                            Catch
                            End Try
                        Next i
                    Else
                        MsgBox("Выберите полилинию которая проходит по низу бордюрного камня!!!")
                        GoTo line1
                    End If
                Else
                    MsgBox("Выберите полилинию которая проходит по низу бордюрного камня!!!")
                    GoTo line1
                End If

            End If
line1:
            acTrans.Commit()
        End Using
    End Sub

    '=============================================================================================
    'формирует отображение верха и низа бордюрного камня
    <CommandMethod("CreateTopAndBottomBK")>
    Public Sub CreateTopAndBottomBK()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
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
        Dim FormTopAndBottomBK As FormTopAndBottomBK = New FormTopAndBottomBK
        FormTopAndBottomBK.ShowDialog()
        If FormTopAndBottomBK.boolShow = False Then
            Exit Sub
        End If
        Dim PutchTemplate As String = FormTopAndBottomBK.ComboBox1.Tag & FormTopAndBottomBK.ComboBox1.Text
        If File.Exists(PutchTemplate) = False Then
            MsgBox("Файл шаблона не найден!!! Макро прервано.")
            Exit Sub
        End If
        '==========================================================================================================================================
        'начальные установки
        '==========================================================================================================================================
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Using acTrans As Transaction = acDoc.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim NamePointGroupPiket As String = FormTopAndBottomBK.ComboBox2.Text
                Dim PointPKGroup As PointGroup = Nothing
                For i As Integer = 0 To CivilDoc.PointGroups.Count - 1
                    Dim PTGroup As PointGroup = acTrans.GetObject(CivilDoc.PointGroups.Item(i), OpenMode.ForRead)
                    If PTGroup.Name Like NamePointGroupPiket Then
                        PointPKGroup = PTGroup
                        Exit For
                    End If
                Next i
                '=====================================================================================================================================
                'копируем стиль низа и верха бордюрного камня и Характерной линии
                Dim DeskCOGOBK As String = FormTopAndBottomBK.TextBox1.Text
                Dim DeskCOGOVBK As String = FormTopAndBottomBK.TextBox2.Text
                Dim DeskFeatureLine As String = FormTopAndBottomBK.TextBox3.Text
                '======================================================================================================================================
                'полевой код для верха и низа бордюрного камня и имя стиля Характерной линии и площадки
                Dim PKODNBK As String = ""
                Dim PKODVBK As String = ""
                Dim NameGroupPoinNBKt As String = ""
                Dim NameGroupPoinVBKt As String = ""
                Dim NameStyleFeatLine As String = ""
                Dim NameLayerFeatLine As String = "0"
                Dim NameSite As String = ""
                Dim IDSiteUser As ObjectId = Nothing
                Dim UDP As UDPDouble = Nothing
                Dim NameUDPUser As String = FormTopAndBottomBK.TextBox4.Text.Trim 'пользовательская классификация
                '======================================================================================================================================
                'ищем характеристики для низа бордюрного камня
                Dim ArrayPropBK As String(,) = Nothing 'Массив с характеристиками низа бордюрного камня
                Dim boolPropCogoPoint As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplate, DeskCOGOBK, ArrayPropBK, "ТочкаCOGO", "")
                If IsArray(ArrayPropBK) = True Then
                    For k As Integer = 0 To ArrayPropBK.GetUpperBound(1)
                        Dim NameField As String = ArrayPropBK(0, k)
                        If NameField.Trim Like "Name" Then
                            PKODNBK = ArrayPropBK(1, k)
                        ElseIf NameField.Trim Like "PointGroup" Then
                            NameGroupPoinNBKt = ArrayPropBK(1, k)
                        End If
                    Next k
                End If
                '======================================================================================================
                'ищем характеристики для верха бордюрного камня
                Dim ArrayPropVBK As String(,) = Nothing 'отметка высот бордюрного камня
                Dim boolPropCogoPointVBK As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplate, DeskCOGOVBK, ArrayPropVBK, "ТочкаCOGO", "")
                If IsArray(ArrayPropVBK) = True Then
                    For k As Integer = 0 To ArrayPropVBK.GetUpperBound(1)
                        Dim NameField As String = ArrayPropVBK(0, k)
                        If NameField.Trim Like "Name" Then
                            PKODVBK = ArrayPropVBK(1, k)
                        ElseIf NameField.Trim Like "PointGroup" Then
                            NameGroupPoinVBKt = ArrayPropVBK(1, k)
                        End If
                    Next k
                End If
                '======================================================================================================
                'ищем характеристики для характерной линии
                Dim ArrayPropFeatLine As String(,) = Nothing 'отметка высот бордюрного камня
                Dim boolPropFeatLine As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplate, DeskFeatureLine, ArrayPropFeatLine, "Характерная линия", "")
                If IsArray(ArrayPropFeatLine) = True Then
                    For k As Integer = 0 To ArrayPropFeatLine.GetUpperBound(1)
                        Dim NameField As String = ArrayPropFeatLine(0, k)
                        If NameField.Trim Like "Name" Then
                            NameStyleFeatLine = ArrayPropFeatLine(1, k)
                        ElseIf NameField.Trim Like "Layer" Then
                            NameLayerFeatLine = ArrayPropFeatLine(1, k)
                        ElseIf NameField.Trim Like "Site" Then
                            NameSite = ArrayPropFeatLine(1, k)
                        End If
                    Next k
                End If
                '======================================================================================================================================
                'проверяем наличие найденных элементов на целевом чертеже
                Dim boolCopyDeskKeyBK As Boolean = False
                Dim boolCopyDeskKeyVBK As Boolean = False
                If FormTopAndBottomBK.CheckBox5.Checked = True Then
                    '=================================================================================================================================
                    'копируем стили точек
                    If PKODNBK.Trim.Length > 0 Then
                        Dim pdKeySetIds As ObjectIdCollection = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(acCurDb).SearchOrder
                        If pdKeySetIds.Count > 0 Then
                            For Each idKeySet As ObjectId In pdKeySetIds
                                Dim DeskKeySet As PointDescriptionKeySet = acTrans.GetObject(idKeySet, OpenMode.ForRead)
                                Dim pdKeyIds As ObjectIdCollection = DeskKeySet.GetPointDescriptionKeyIds
                                For Each idDeskKey As ObjectId In pdKeyIds
                                    If idDeskKey.IsNull = False Then
                                        Dim DeskKey As PointDescriptionKey = acTrans.GetObject(idDeskKey, OpenMode.ForRead)
                                        If DeskKey.Code Like PKODNBK Then
                                            boolCopyDeskKeyBK = True
                                        ElseIf DeskKey.Code Like PKODVBK Then
                                            boolCopyDeskKeyVBK = True
                                        End If
                                    End If
                                Next
                            Next
                        End If
                    End If
                    If boolCopyDeskKeyBK = False Then
                        boolCopyDeskKeyBK = AcadFunction.FuncCopyDeskKeyDestDraw(PutchTemplate, PKODNBK)
                    End If
                    If boolCopyDeskKeyVBK = False Then
                        boolCopyDeskKeyVBK = AcadFunction.FuncCopyDeskKeyDestDraw(PutchTemplate, PKODVBK)
                    End If
                    '========================================================================================================================================
                    'копируем группу точек
                    Dim boolGroupPointBKN As Boolean = False
                    Dim boolGroupPointBKV As Boolean = False
                    Dim PTGroupBKN As PointGroup = Nothing
                    Dim PTGroupBKV As PointGroup = Nothing
                    For i As Integer = 0 To CivilDoc.PointGroups.Count - 1
                        Dim PTGroup As PointGroup = acTrans.GetObject(CivilDoc.PointGroups.Item(i), OpenMode.ForRead)
                        If PTGroup.Name Like NameGroupPoinNBKt Then
                            PTGroupBKN = PTGroup
                            boolGroupPointBKN = True
                        End If
                        If PTGroup.Name Like NameGroupPoinNBKt Then
                            PTGroupBKV = PTGroup
                            boolGroupPointBKV = True
                        End If
                    Next i
                    If boolGroupPointBKN = False Then
                        If NameGroupPoinNBKt.Trim.Length > 0 Then
                            boolGroupPointBKN = AcadFunction.FuncCopyGroupPointsDestDraw(PutchTemplate, NameGroupPoinNBKt)
                        End If
                    End If
                    If boolGroupPointBKV = False Then
                        If NameGroupPoinVBKt.Trim.Length > 0 Then
                            boolGroupPointBKV = AcadFunction.FuncCopyGroupPointsDestDraw(PutchTemplate, NameGroupPoinVBKt)
                        End If
                    End If
                    '===========================================================================================================
                    'пользовательская классификация
                    Dim BoolUDP As Boolean = False
                    For Each IdUdp As UDP In CivilDoc.PointUDPs
                        If IsNothing(IdUdp) = False Then
                            Try
                                If TypeOf IdUdp Is UDPDouble Then
                                    If IdUdp.Name Like NameUDPUser Then
                                        UDP = IdUdp
                                        BoolUDP = True
                                        Exit For
                                    End If
                                End If
                            Catch ex As System.NullReferenceException
                            End Try
                        End If
                    Next
                End If
                '=======================================================================================================================
                'копируем стили характерной линии
                Dim boolCopyStyleFeatLine As Boolean = False
                If FormTopAndBottomBK.CheckBox1.Checked = True Then
                    'копируем стиль линии
                    Dim idFeatStyle As ObjectId = Nothing
                    Try
                        idFeatStyle = CivilDoc.Styles.FeatureLineStyles.Item(NameStyleFeatLine)
                    Catch ex As System.ArgumentException
                    End Try
                    If IsNothing(idFeatStyle) = True Then
                        boolCopyStyleFeatLine = AcadFunction.FuncCopyFeatureLineStyleDestDraw(PutchTemplate, NameStyleFeatLine)
                    End If
                    If idFeatStyle.IsNull = True Then
                        boolCopyStyleFeatLine = AcadFunction.FuncCopyFeatureLineStyleDestDraw(PutchTemplate, NameStyleFeatLine)
                    Else
                        boolCopyStyleFeatLine = True
                    End If
                    'копируем слой
                    Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                    If acLyrTbl.Has(NameLayerFeatLine) = False Then
                        Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplate, NameLayerFeatLine)
                        If boolCopyLayer = False Then
                            NameLayerFeatLine = "0"
                        End If
                    End If
                    'создаем площадку
                    Dim boolSite As Boolean = False
                    For Each idSite As ObjectId In CivilDoc.GetSiteIds
                        Dim UserSite As Site = acTrans.GetObject(idSite, OpenMode.ForRead)
                        If UserSite.Name Like NameSite Then
                            IDSiteUser = UserSite.Id
                            boolSite = True
                            Exit For
                        End If
                    Next
                    If boolSite = False Then
                        Try
                            IDSiteUser = Site.Create(CivilDoc, NameSite)
                            boolSite = True
                        Catch ex As System.ArgumentException
                        End Try
                    End If
                End If

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'значение по умолчанию для тех точек верха которых небыли найдены
                Dim BKV0 As Double = FormTopAndBottomBK.NumericUpDown1.Value
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ручной режим
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If FormTopAndBottomBK.RadioButton1.Checked Then
                    Dim resBlk As PromptEntityResult = Nothing
                    Dim optBlk As PromptEntityOptions = New PromptEntityOptions(vbLf & "Укажите точку COGO низа бордюрного камня: ")
                    optBlk.AllowNone = True
                    Dim optBlk1 As PromptEntityOptions = New PromptEntityOptions(vbLf & "Укажите точку COGO верха бордюрного камня: ")
                    optBlk1.AllowNone = True
                    Do
                        Dim tr1 As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
                        Using tr1
                            Try
                                Dim CoGoPoint1 As CogoPoint = Nothing
                                resBlk = ed.GetEntity(optBlk)
                                Dim obj1 As Autodesk.AutoCAD.DatabaseServices.DBObject = Nothing
                                If resBlk.Status = PromptStatus.OK Then
                                    obj1 = tr1.GetObject(resBlk.ObjectId, OpenMode.ForWrite)
                                    If TypeOf obj1 Is CogoPoint Then
                                        CoGoPoint1 = obj1
                                    End If
                                Else
                                    Exit Do
                                End If
                                'отметка точки низа БК
                                Dim ElevPoint1 As Double = -1
                                If CoGoPoint1.IsPlanar = False Then
                                    ElevPoint1 = CoGoPoint1.Elevation
                                Else
                                    MsgBox("Выбранная точка не имеет высотной отметки!!!")
                                    Exit Do
                                End If
                                'выбираем вторую точку верха бордюра
                                If FormTopAndBottomBK.CheckBox3.Checked = False Then
                                    Dim CoGoPoint2 As CogoPoint = Nothing
                                    resBlk = ed.GetEntity(optBlk1)
                                    Dim obj2 As Autodesk.AutoCAD.DatabaseServices.DBObject = Nothing
                                    If resBlk.Status = PromptStatus.OK Then
                                        obj2 = tr1.GetObject(resBlk.ObjectId, OpenMode.ForWrite)
                                        If TypeOf obj2 Is CogoPoint Then
                                            CoGoPoint2 = obj2
                                        End If
                                    End If

                                    Dim ElevPoint2 As Double = -1
                                    If CoGoPoint2.IsPlanar = False Then
                                        ElevPoint2 = CoGoPoint2.Elevation
                                    Else
                                        MsgBox("Выбранная точка не имеет высотной отметки!!! Макро прервано.")
                                        Exit Do
                                    End If
                                    If ElevPoint1 <= ElevPoint2 Then
                                        CoGoPoint2.RawDescription = PKODVBK
                                        Dim h As Double = Val(ElevPoint2 - ElevPoint1)
                                        Dim HPoint As Double = ElevPoint1 + h
                                        CoGoPoint1.RawDescription = PKODNBK
                                        CoGoPoint1.SetUDPValue(UDP, h) 'присваиваем ключ
                                        CoGoPoint1.ApplyDescriptionKeys() 'применяем ключи описатели
                                    End If
                                Else
                                    CoGoPoint1.RawDescription = PKODNBK
                                    CoGoPoint1.ApplyDescriptionKeys() 'применяем ключи описатели
                                    CoGoPoint1.SetUDPValue(UDP, FormTopAndBottomBK.NumericUpDown3.Value) 'присваиваем ключ
                                End If
                            Catch ex As Exception
                                ed.WriteMessage("Ошибка. Макро прервано!!!" & vbLf)
                            End Try
                            tr1.Commit()
                        End Using
                    Loop
                    '=============================================================================================================================
                    'автоматическое заполнение
                    '============================================================================================================================
                Else
                    'выбираем польлинию для определения угла поворота метки
                    'массив 0=X,1=Y,2=Hниза,3=Hверха,4-Номер точки низа,5=Номер точки верха,6=Расстояние до точек
                    Dim ArrayPointXY As Double(,) = Nothing 'массив для координат точек выбранной полилинии
                    Dim acPline As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
                    Dim acPline2d As Autodesk.AutoCAD.DatabaseServices.Polyline2d = Nothing
                    Dim acPline3d As Autodesk.AutoCAD.DatabaseServices.Polyline2d = Nothing
                    Dim acFeatLine As Autodesk.Civil.DatabaseServices.FeatureLine = Nothing
                    Dim boolPline As Boolean = False
                    Dim boolPline2d As Boolean = False
                    Dim boolFeatLine As Boolean = False
                    Dim boolClose As Boolean = False
                    Dim Ext As Extents3d = Nothing
                    Dim SelectPline As PromptEntityResult = ed.GetEntity("Выберите полилинию определяющую низ бортового камня:  ")
                    If SelectPline.Status = PromptStatus.OK Then
                        Dim acEnt1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(SelectPline.ObjectId, OpenMode.ForRead)
                        If TypeOf acEnt1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt1
                            Dim PtTemp As Point2d = New Point2d(-1, -1)
                            Dim CountVert As Integer = 0
                            If Pline.Closed = True Then
                                boolClose = True
                            End If
                            For i As Integer = 0 To Pline.NumberOfVertices - 1
                                Dim PointXY As Point2d = Pline.GetPoint2dAt(i)
                                If i > 1 And PointXY.GetDistanceTo(Pline.GetPoint2dAt(0)) = 0 Then
                                    boolClose = True
                                    Exit For
                                End If
                                If PointXY.GetDistanceTo(PtTemp) > 0.04 Then
                                    ReDim Preserve ArrayPointXY(6, CountVert)
                                    ArrayPointXY(0, CountVert) = PointXY.X
                                    ArrayPointXY(1, CountVert) = PointXY.Y
                                    PtTemp = PointXY
                                    CountVert += 1
                                End If
                            Next i
                            boolPline = True
                            Ext = Pline.GeometricExtents
                        ElseIf TypeOf acEnt1 Is Polyline2d Then
                            acPline2d = acEnt1
                            Dim PtTemp As Point2d = New Point2d(-1, -1)
                            Dim StartPt As Point2d = New Point2d(-2, -2)
                            Dim CountVert As Integer = 0
                            If acPline2d.Closed = True Then
                                boolClose = True
                            End If
                            For Each VertID As ObjectId In acPline2d
                                Dim v2d As Vertex2d = acTrans.GetObject(VertID, OpenMode.ForRead)
                                Dim Pt As Point2d = New Point2d(v2d.Position.X, v2d.Position.Y)
                                If CountVert = 0 Then
                                    StartPt = Pt
                                End If
                                If CountVert > 1 And Pt.GetDistanceTo(StartPt) = 0 Then
                                    boolClose = True
                                    Exit For
                                End If
                                If Pt.GetDistanceTo(PtTemp) > 0.04 Then
                                    ReDim Preserve ArrayPointXY(6, CountVert)
                                    ArrayPointXY(0, CountVert) = Pt.X
                                    ArrayPointXY(1, CountVert) = Pt.Y
                                    PtTemp = Pt
                                    CountVert += 1
                                End If
                            Next
                            boolPline2d = True
                            Ext = acPline2d.GeometricExtents
                        ElseIf TypeOf acEnt1 Is FeatureLine Then
                            acFeatLine = acEnt1
                            If acFeatLine.Closed = True Then
                                boolClose = True
                            End If
                            Dim PointColl As Point3dCollection = New Point3dCollection
                            Dim ftype As Autodesk.Civil.FeatureLinePointType = Autodesk.Civil.FeatureLinePointType.PIPoint
                            PointColl = acFeatLine.GetPoints(ftype)
                            If PointColl.Count > 1 Then
                                Dim CountVert As Integer = 0
                                Dim PtTemp As Point3d = New Point3d(-1, -1, -1)
                                For Each pt3d As Point3d In PointColl
                                    If CountVert > 1 And pt3d.Convert2d(New Plane).GetDistanceTo(PointColl.Item(0).Convert2d(New Plane)) = 0 Then
                                        boolClose = True
                                        Exit For
                                    End If
                                    If pt3d.Convert2d(New Plane).GetDistanceTo(PtTemp.Convert2d(New Plane)) > 0.04 Then
                                        ReDim Preserve ArrayPointXY(6, CountVert)
                                        ArrayPointXY(0, CountVert) = pt3d.X
                                        ArrayPointXY(1, CountVert) = pt3d.Y
                                        ArrayPointXY(2, CountVert) = pt3d.Z
                                        PtTemp = pt3d
                                        CountVert += 1
                                    End If
                                Next
                            End If
                            boolFeatLine = True
                            Ext = acFeatLine.GeometricExtents
                        End If
                        If IsArray(ArrayPointXY) = False Then
                            MsgBox("Не удалось записать в массив координаты точек выбранного примитива! Макро прервано.")
                            Exit Sub
                        End If
                        '============================================================================================================
                        'определяем сторону смещения
                        Dim PicPoint As Point3d = Nothing
                        If FormTopAndBottomBK.CheckBox6.Checked = False Then
                            Dim opt As PromptPointOptions = New PromptPointOptions("Укажите сторону смещения:")
                            Dim pVAlPt As PromptPointResult = ed.GetPoint(opt)
                            PicPoint = pVAlPt.Value
                        End If
                        '============================================================================================================
                        'ищем cogo точки вдоль полилинии
                        Dim ArrayPoint As Integer() = Nothing
                        'создаем габаритный контейнер полилинии для фильтрации точек
                        Dim MinCoord As Point3d = Ext.MinPoint
                        Dim MaxCoord As Point3d = Ext.MaxPoint
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
                        '=========================================================================================================
                        'максимальное расстояние до верха БК
                        Dim Lenght As Double = FormTopAndBottomBK.NumericUpDown2.Value
                        Lenght = Math.Round(Lenght, 2)
                        'начинаем перебор вершин полилинии
                        For i As Integer = 0 To ArrayPointXY.GetUpperBound(1)
                            Try
                                Dim h As Double = 0
                                Dim RotateMarker As Double = 0
                                Dim PointCogoBKN As CogoPoint = Nothing
                                Dim pointCogoBKV As CogoPoint = Nothing
                                Dim ElevPointBKN As Double = -1
                                Dim ElevPointBKV As Double = -1
                                Dim Point As Point3d = New Point3d(ArrayPointXY(0, i), ArrayPointXY(1, i), 0)
                                'ищем точку в узле полилинии 
                                If IsArray(ArrayPoint) = True Then
                                    PointCogoBKN = FuncSurvey.FuncFindCogoPointByCoordinates(Point, ArrayPoint, 0.05)
                                    If IsNothing(PointCogoBKN) = True Then
                                        If IsNothing(ArrayPointXY(2, i)) = True Then
                                            ArrayPointXY(2, i) = 0
                                        End If
                                    ElseIf PointCogoBKN.Id.IsNull = True Then
                                        If IsNothing(ArrayPointXY(2, i)) = True Then
                                            ArrayPointXY(2, i) = 0
                                        End If
                                    Else
                                        If PointCogoBKN.IsPlanar = False Then
                                            ArrayPointXY(2, i) = PointCogoBKN.Elevation
                                            ElevPointBKN = PointCogoBKN.Elevation
                                            ArrayPointXY(4, i) = PointCogoBKN.PointNumber
                                        Else
                                            ArrayPointXY(2, i) = 0
                                        End If
                                    End If
                                    'шщем ближайшую точку к исходной
                                    If FormTopAndBottomBK.CheckBox3.Checked = False Then
                                        If Not (PointCogoBKN Is Nothing) Then
                                            If PointCogoBKN.Id.IsNull = False Then
                                                pointCogoBKV = FuncSurvey.FuncFindCogoPointByDistance(PointCogoBKN.PointNumber, Lenght)
                                                If IsNothing(pointCogoBKV) = True Then
                                                    ArrayPointXY(3, i) = 0
                                                ElseIf pointCogoBKV.Id.IsNull = True Then
                                                    ArrayPointXY(3, i) = 0
                                                Else
                                                    If pointCogoBKV.IsPlanar = False Then
                                                        ArrayPointXY(3, i) = pointCogoBKV.Elevation
                                                        ElevPointBKV = pointCogoBKV.Elevation
                                                        ArrayPointXY(5, i) = pointCogoBKV.PointNumber
                                                        Dim distPtNV As Double = PointCogoBKN.Location.Convert2d(New Plane).GetDistanceTo(pointCogoBKV.Location.Convert2d(New Plane))
                                                        ArrayPointXY(6, i) = distPtNV
                                                    Else
                                                        ArrayPointXY(3, i) = 0
                                                    End If
                                                End If
                                            Else
                                                ArrayPointXY(3, i) = 0
                                                Continue For
                                            End If
                                        Else
                                            ArrayPointXY(3, i) = 0
                                            Continue For
                                        End If
                                    Else
                                        If ElevPointBKN <> -1 Then
                                            ArrayPointXY(3, i) = PointCogoBKN.Elevation + FormTopAndBottomBK.NumericUpDown3.Value
                                        Else
                                            ArrayPointXY(3, i) = 0
                                        End If
                                    End If
                                End If
                            Catch
                            End Try
                        Next i
                        '==========================================================================================================
                        'находим высоту начала БК (по низу)
                        Dim H0 As Double = ArrayPointXY(2, 0)
                        If H0 = 0 Then
                            For i As Integer = 0 To ArrayPointXY.GetUpperBound(1)
                                Dim TempH As Double = ArrayPointXY(2, i)
                                If TempH <> 0 Then
                                    ArrayPointXY(2, 0) = TempH
                                    Exit For
                                End If
                            Next i
                        End If
                        '=========================================================================================================
                        'находим высоту конца БК (по низу)
                        H0 = ArrayPointXY(2, ArrayPointXY.GetUpperBound(1))
                        If H0 = 0 Then
                            For i As Integer = ArrayPointXY.GetUpperBound(1) To 0 Step -1
                                Dim TempH As Double = ArrayPointXY(2, i)
                                If TempH <> 0 Then
                                    ArrayPointXY(2, ArrayPointXY.GetUpperBound(1)) = TempH
                                    Exit For
                                End If
                            Next i
                        End If
                        '==========================================================================================================
                        'находим высоту начала верха БК
                        H0 = ArrayPointXY(3, 0)
                        If H0 = 0 Then
                            If FormTopAndBottomBK.CheckBox3.Checked = True Then
                                ArrayPointXY(3, 0) = ArrayPointXY(2, 0) + FormTopAndBottomBK.NumericUpDown3.Value
                            Else
                                For i As Integer = 0 To ArrayPointXY.GetUpperBound(1)
                                    Dim TempH As Double = ArrayPointXY(3, i)
                                    If TempH <> 0 Then
                                        ArrayPointXY(3, 0) = TempH
                                        Exit For
                                    End If
                                Next i
                            End If
                        End If
                        H0 = ArrayPointXY(3, 0)
                        If H0 = 0 Then
                            ArrayPointXY(3, 0) = ArrayPointXY(2, 0)
                        End If
                        '=============================================================================================================
                        'находим высоту конца БК
                        H0 = ArrayPointXY(3, ArrayPointXY.GetUpperBound(1))
                        If H0 = 0 Then
                            If FormTopAndBottomBK.CheckBox3.Checked = True Then
                                ArrayPointXY(3, ArrayPointXY.GetUpperBound(1)) = ArrayPointXY(2, ArrayPointXY.GetUpperBound(1)) + FormTopAndBottomBK.NumericUpDown3.Value
                            Else
                                ArrayPointXY(3, ArrayPointXY.GetUpperBound(1)) = ArrayPointXY(2, ArrayPointXY.GetUpperBound(1))
                            End If
                        End If
                        H0 = ArrayPointXY(3, ArrayPointXY.GetUpperBound(1))
                        If H0 = 0 Then
                            For i As Integer = ArrayPointXY.GetUpperBound(1) To 0 Step -1
                                Dim TempH As Double = ArrayPointXY(3, ArrayPointXY.GetUpperBound(1))
                                If TempH <> 0 Then
                                    ArrayPointXY(3, ArrayPointXY.GetUpperBound(1)) = TempH
                                    Exit For
                                End If
                            Next i
                        End If
                        '===============================================================================================================
                        'интерполируем остальной массив по верху бордюрного камня
                        '===============================================================================================================
                        Dim startElevPoint As Double = 0
                        Dim endElevPoint As Double = 0
                        Dim dist As Double = 0
                        For i As Integer = 1 To ArrayPointXY.GetUpperBound(1) - 1
                            Dim fPoint As Double = ArrayPointXY(3, i)
                            If fPoint = 0 Then
                                'смотрим предыдущий пикет
                                startElevPoint = ArrayPointXY(3, i - 1)
                                Dim pt1 As Point2d = New Point2d(ArrayPointXY(0, i - 1), ArrayPointXY(1, i - 1))
                                Dim pt2 As Point2d = New Point2d(ArrayPointXY(0, i), ArrayPointXY(1, i))
                                dist = pt1.GetVectorTo(pt2).Length
                                For j As Integer = i + 1 To ArrayPointXY.GetUpperBound(1)
                                    Dim ftempPoint As Double = ArrayPointXY(3, j)
                                    If ftempPoint = 0 Then
                                        Dim pt3 As Point2d = New Point2d(ArrayPointXY(0, j - 1), ArrayPointXY(1, j - 1))
                                        Dim pt4 As Point2d = New Point2d(ArrayPointXY(0, j), ArrayPointXY(1, j))
                                        dist = dist + pt3.GetVectorTo(pt4).Length
                                    Else
                                        endElevPoint = ArrayPointXY(3, j)
                                        Dim temppt1 As Point2d = New Point2d(ArrayPointXY(0, j - 1), ArrayPointXY(1, j - 1))
                                        Dim temppt2 As Point2d = New Point2d(ArrayPointXY(0, j), ArrayPointXY(1, j))
                                        dist = dist + temppt1.GetVectorTo(temppt2).Length
                                        Exit For
                                    End If
                                Next j
                                '===========================================
                                'делаем расчет
                                Dim deltaH As Double = endElevPoint - startElevPoint
                                Dim ddist As Double = pt1.GetVectorTo(pt2).Length / dist
                                Dim rez1 As Double = ddist * deltaH
                                ArrayPointXY(3, i) = startElevPoint + rez1
                            End If
                        Next i
                        '==================================================================================================
                        'находим ширину бордюрного камня
                        Dim DistBK As Double = 0
                        Dim CountBK As Integer = 0
                        If FormTopAndBottomBK.CheckBox2.Checked = True Then
                            For i As Integer = 0 To ArrayPointXY.GetUpperBound(1)
                                Dim tempL As Double = ArrayPointXY(6, i)
                                If tempL > 0.04 Then
                                    DistBK = DistBK + tempL
                                    CountBK += 1
                                End If
                            Next
                        End If
                        If CountBK > 0 And DistBK > 0 Then
                            DistBK = Math.Round(DistBK / CountBK, 2)
                        Else
                            DistBK = FormTopAndBottomBK.NumericUpDown1.Value
                        End If
                        '===========================================================================================================
                        'заносим коды в точки COGO
                        For i As Integer = 0 To ArrayPointXY.GetUpperBound(1)
                            Dim NumBKN As Integer = ArrayPointXY(4, i)
                            Dim NumBKV As Integer = ArrayPointXY(5, i)
                            If NumBKN = 0 Then Continue For
                            Dim IDpointCogoBKN As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumBKN)
                            Dim pointCogoBKN As CogoPoint = Nothing
                            If IDpointCogoBKN.IsNull = False Then
                                pointCogoBKN = acTrans.GetObject(IDpointCogoBKN, OpenMode.ForWrite)
                                Dim h As Double = Math.Round((ArrayPointXY(3, i) - ArrayPointXY(2, i)), 2)
                                If FormTopAndBottomBK.CheckBox7.Checked = True Then
                                    Dim RotateMarker As Double = FuncSurvey.FuncCalcRotateMarkerCoGoPoint(ArrayPointXY, pointCogoBKN.Easting, pointCogoBKN.Northing)
                                    pointCogoBKN.LabelRotation = RotateMarker
                                End If
                                If IsNothing(UDP) = False Then
                                    If h >= 0 Then
                                        pointCogoBKN.SetUDPValue(UDP, h)
                                    End If
                                End If
                                pointCogoBKN.RawDescription = PKODNBK
                                pointCogoBKN.ApplyDescriptionKeys()
                            End If
                            If NumBKV > 0 Then
                                Dim IDpointCogoBKV As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumBKV)
                                Dim pointCogoBKV As CogoPoint = Nothing
                                If IDpointCogoBKV.IsNull = False Then
                                    pointCogoBKV = acTrans.GetObject(IDpointCogoBKV, OpenMode.ForWrite)
                                    pointCogoBKV.RawDescription = PKODVBK
                                    pointCogoBKV.ApplyDescriptionKeys()
                                End If
                            End If
                        Next i
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'строим характерную линию верха бордюрного камня
                        '===========================================================================================================
                        Dim pts As Point3dCollection = New Point3dCollection
                        Dim acPoly As Polyline = New Polyline()
                        acPoly.SetDatabaseDefaults()
                        For i As Integer = 0 To ArrayPointXY.GetUpperBound(1)
                            acPoly.AddVertexAt(i, New Point2d(ArrayPointXY(0, i), ArrayPointXY(1, i)), 0, 0, 0)
                        Next i
                        acBlkTblRec.AppendEntity(acPoly)
                        acTrans.AddNewlyCreatedDBObject(acPoly, True)
                        '=============================================================================================================
                        Dim OffsetPLine As Integer = 1
                        If FormTopAndBottomBK.CheckBox6.Checked = False Then
                            Dim pDir As Point3d = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("VIEWDIR")
                            Dim boolR As Boolean = MathFunction.FuncIsRightDirection(acPoly, PicPoint, pDir.GetAsVector)
                            If boolR = False Then
                                OffsetPLine = -1
                            End If
                        Else
                            Dim NumR As Integer = 0
                            Dim NumL As Integer = 0
                            Dim pDir As Point3d = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("VIEWDIR")
                            For i As Integer = 0 To ArrayPointXY.GetUpperBound(1)
                                Dim NumBKV As Integer = ArrayPointXY(5, i)
                                If NumBKV > 0 Then
                                    Dim IDpointCogoBKV As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumBKV)
                                    Dim pointCogoBKV As CogoPoint = Nothing
                                    If IDpointCogoBKV.IsNull = False Then
                                        pointCogoBKV = acTrans.GetObject(IDpointCogoBKV, OpenMode.ForWrite)
                                        Dim boolR As Boolean = MathFunction.FuncIsRightDirection(acPoly, pointCogoBKV.Location, pDir.GetAsVector)
                                        If boolR = False Then
                                            NumL += 1
                                        Else
                                            NumR += 1
                                        End If
                                    End If
                                End If
                            Next
                            If NumL > NumR Then
                                OffsetPLine = -1
                            End If
                        End If
                        '==========================================================================================================
                        'делаем смещение
                        If boolClose = False Then
                            Dim dbcoll As DBObjectCollection = acPoly.GetOffsetCurves(0.01 * OffsetPLine)
                            If dbcoll.Count > 0 Then
                                Dim LenghtPline As Double = 0
                                For Each odj As Autodesk.AutoCAD.DatabaseServices.DBObject In dbcoll
                                    If TypeOf odj Is Polyline Then
                                        Dim Poly1 As Polyline = odj
                                        If Poly1.Length > LenghtPline Then
                                            pts = New Point3dCollection
                                            If Poly1.NumberOfVertices - 1 = ArrayPointXY.GetUpperBound(1) Then
                                                For i As Integer = 0 To Poly1.NumberOfVertices - 1
                                                    pts.Add(New Point3d(Poly1.GetPoint2dAt(i).X, Poly1.GetPoint2dAt(i).Y, ArrayPointXY(3, i)))
                                                Next
                                                LenghtPline = Poly1.Length
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                            '==========================================================================================================
                            'делаем 2 смещение
                            Dim dbcoll2 As DBObjectCollection = acPoly.GetOffsetCurves(DistBK * OffsetPLine)
                            If dbcoll2.Count > 0 Then
                                Dim LenghtPline As Double = 0
                                For Each odj As Autodesk.AutoCAD.DatabaseServices.DBObject In dbcoll2
                                    If TypeOf odj Is Polyline Then
                                        Dim Poly1 As Polyline = odj
                                        If Poly1.Length > LenghtPline Then
                                            pts = New Point3dCollection
                                            If Poly1.NumberOfVertices - 1 = ArrayPointXY.GetUpperBound(1) Then
                                                For i As Integer = Poly1.NumberOfVertices - 1 To 0 Step -1
                                                    pts.Add(New Point3d(Poly1.GetPoint2dAt(i).X, Poly1.GetPoint2dAt(i).Y, ArrayPointXY(3, i)))
                                                Next
                                                LenghtPline = Poly1.Length
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                            acPoly.Erase()
                            If pts.Count > 1 Then
                                Using poly3d As Polyline3d = New Polyline3d(Poly3dType.SimplePoly, pts, True)
                                    acBlkTblRec.AppendEntity(poly3d)
                                    acTrans.AddNewlyCreatedDBObject(poly3d, True)
                                    If FormTopAndBottomBK.CheckBox1.Checked = True Then
                                        Dim idFeat As ObjectId = Autodesk.Civil.DatabaseServices.FeatureLine.Create("", poly3d.Id)
                                        Dim FeatLine As Autodesk.Civil.DatabaseServices.FeatureLine = acTrans.GetObject(idFeat, OpenMode.ForWrite)
                                        FeatLine.StyleName = NameStyleFeatLine
                                        FeatLine.Layer = NameLayerFeatLine
                                        Try
                                            FeatureLine.MoveToSite(idFeat, IDSiteUser)
                                        Catch ex As System.ArgumentException
                                        End Try
                                    End If
                                    If FormTopAndBottomBK.CheckBox4.Checked = False Then
                                        poly3d.Erase()
                                    End If
                                End Using
                            End If
                        Else
                            Dim dbcoll As DBObjectCollection = acPoly.GetOffsetCurves(0.01 * OffsetPLine)
                            For Each odj As Autodesk.AutoCAD.DatabaseServices.DBObject In dbcoll
                                Dim Poly1 As Polyline = odj
                                If Poly1.NumberOfVertices - 1 = ArrayPointXY.GetUpperBound(1) Then
                                    For i As Integer = 0 To Poly1.NumberOfVertices - 1
                                        pts.Add(New Point3d(Poly1.GetPoint2dAt(i).X, Poly1.GetPoint2dAt(i).Y, ArrayPointXY(3, i)))
                                    Next
                                End If
                                Using poly3d As Polyline3d = New Polyline3d(Poly3dType.SimplePoly, pts, True)
                                    acBlkTblRec.AppendEntity(poly3d)
                                    acTrans.AddNewlyCreatedDBObject(poly3d, True)
                                    If FormTopAndBottomBK.CheckBox1.Checked = True Then
                                        Dim idFeat As ObjectId = Autodesk.Civil.DatabaseServices.FeatureLine.Create("", poly3d.Id)
                                        Dim FeatLine As Autodesk.Civil.DatabaseServices.FeatureLine = acTrans.GetObject(idFeat, OpenMode.ForWrite)
                                        FeatLine.StyleName = NameStyleFeatLine
                                        FeatLine.Layer = NameLayerFeatLine
                                        Try
                                            FeatureLine.MoveToSite(idFeat, IDSiteUser)
                                        Catch ex As System.ArgumentException
                                        End Try
                                    End If
                                    If FormTopAndBottomBK.CheckBox4.Checked = False Then
                                        poly3d.Erase()
                                    End If
                                End Using
                            Next
                            '==========================================================================================================
                            'делаем 2 смещение
                            Dim dbcoll2 As DBObjectCollection = acPoly.GetOffsetCurves(DistBK * OffsetPLine)
                            pts = New Point3dCollection
                            For Each odj As Autodesk.AutoCAD.DatabaseServices.DBObject In dbcoll2
                                Dim Poly1 As Polyline = odj
                                If Poly1.NumberOfVertices - 1 = ArrayPointXY.GetUpperBound(1) Then
                                    For i As Integer = 0 To Poly1.NumberOfVertices - 1
                                        pts.Add(New Point3d(Poly1.GetPoint2dAt(i).X, Poly1.GetPoint2dAt(i).Y, ArrayPointXY(3, i)))
                                    Next
                                End If
                                Using poly3d As Polyline3d = New Polyline3d(Poly3dType.SimplePoly, pts, True)
                                    acBlkTblRec.AppendEntity(poly3d)
                                    acTrans.AddNewlyCreatedDBObject(poly3d, True)
                                    If FormTopAndBottomBK.CheckBox1.Checked = True Then
                                        Dim idFeat As ObjectId = Autodesk.Civil.DatabaseServices.FeatureLine.Create("", poly3d.Id)
                                        Dim FeatLine As Autodesk.Civil.DatabaseServices.FeatureLine = acTrans.GetObject(idFeat, OpenMode.ForWrite)
                                        FeatLine.StyleName = NameStyleFeatLine
                                        FeatLine.Layer = NameLayerFeatLine
                                        Try
                                            FeatureLine.MoveToSite(idFeat, IDSiteUser)
                                        Catch ex As System.ArgumentException
                                        End Try
                                    End If
                                    If FormTopAndBottomBK.CheckBox4.Checked = False Then
                                        poly3d.Erase()
                                    End If
                                End Using
                            Next
                            acPoly.Erase()
                        End If
                    End If
                End If
line1:
                acTrans.Commit()
            End Using
        End Using
    End Sub
End Class
