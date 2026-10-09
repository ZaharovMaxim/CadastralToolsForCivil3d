Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.Globalization
'раставляет отметки по кюветам
Public Class СalcMarksAndSlopesOnCuvettes
    <CommandMethod("СalcMarksAndSlopesOnCuvettes")>
    Public Sub СalcMarksAndSlopesOnCuvettes()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
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
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim FormСalcMarksAndSlopesOnCuvettes As СalcMarksSlopesCuvets = New СalcMarksSlopesCuvets
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'запускаем форму
        FormСalcMarksAndSlopesOnCuvettes.ShowDialog()
        If FormСalcMarksAndSlopesOnCuvettes.boolIns = False Then
            tr.Commit()
            locker.Dispose()
            Exit Sub
        End If

        '=======================================================================================
        'начальные установки
        Dim RoundZn As Integer = FormСalcMarksAndSlopesOnCuvettes.NumericUpDown4.Value 'округление
        Dim OffsetLabel As Integer = FormСalcMarksAndSlopesOnCuvettes.NumericUpDown2.Value 'смещение по X
        Dim AngleLabel As Integer = FormСalcMarksAndSlopesOnCuvettes.NumericUpDown1.Value 'угол отклонения
        Dim NameAnnoScale As String = FormСalcMarksAndSlopesOnCuvettes.ComboBox7.Text 'масштаб аннотаций
        Dim PutshTemplate As String = FormСalcMarksAndSlopesOnCuvettes.PatchFileTemplate
        If IO.File.Exists(PutshTemplate) = False Then
            MsgBox("Шаблон не найден. Макро прервано!!!")
            Exit Sub
        End If
        'смотрим мультивыноску
        Dim NameMLider As String = "Отметка дна кювета"
        Dim ArrayPropertiesMleaderCoord As String(,) = Nothing
        Dim IDMLeaderStyle As ObjectId = Nothing
        Dim NameLayerMLeader As String = "0"
        'копируем стиль из шаблона
        'читаем характеристики стиля мультивыноски
        Dim mlstyles As DBDictionary = tr.GetObject(acCurDb.MLeaderStyleDictionaryId, OpenMode.ForWrite)
        Dim acLyrTbl As LayerTable = tr.GetObject(acCurDb.LayerTableId, OpenMode.ForWrite)
        Dim boolReadStyleML As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutshTemplate, NameMLider, ArrayPropertiesMleaderCoord, "Мультивыноска", "Трасса автомобильной дороги")
        If IsArray(ArrayPropertiesMleaderCoord) = True Then
            For i As Integer = 0 To ArrayPropertiesMleaderCoord.GetUpperBound(1)
                If ArrayPropertiesMleaderCoord(0, i) Like "Name" Then
                    Dim NameMleaderStyleCoord As String = ArrayPropertiesMleaderCoord(1, i)
                    If IsNothing(NameMleaderStyleCoord) = False Then
                        Dim boolCopyMleaderStyle As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(PutshTemplate, NameMleaderStyleCoord)
                        For Each item As DBDictionaryEntry In mlstyles
                            If NameMleaderStyleCoord Like item.Key Then
                                IDMLeaderStyle = item.Value
                                Exit For
                            End If
                        Next
                    End If
                ElseIf ArrayPropertiesMleaderCoord(0, i) Like "Layer" Then
                    Dim NameLayerTemp As String = ArrayPropertiesMleaderCoord(1, i)
                    If IsNothing(NameLayerTemp) = False Then
                        Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutshTemplate, NameLayerTemp)
                        If acLyrTbl.Has(NameLayerTemp) = True Then
                            NameLayerMLeader = NameLayerTemp
                            Exit For
                        End If
                    End If
                End If
            Next i
        End If
        'ВСТАВЛЯЕМ НEОБХОДИМЫЕ БЛОКИ
        '=======================================================================
        ''вставляем определение блока в базу
        Dim NameBlk As String = "Уклон-Расстояние"
        Dim ArrayPropBlk As String(,) = Nothing
        Dim boolPropBlkUPK As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutshTemplate, NameBlk, ArrayPropBlk, "Блок", "Трасса автомобильной дороги")
        If IsArray(ArrayPropBlk) = True Then
            For i As Integer = 0 To ArrayPropBlk.GetUpperBound(1)
                Dim NameField As String = ArrayPropBlk(0, i)
                If NameField.Trim Like "Name" Then
                    NameBlk = ArrayPropBlk(1, i)
                    If IsNothing(NameBlk) = False Then
                        If NameBlk.Trim.Length > 0 Then
                            Dim boolInsRefBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutshTemplate, NameBlk)
                            Exit For
                        End If
                    End If
                End If
            Next i
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'выставляем масштаб
        Dim NameScale As String = acCurDb.Cannoscale.Name
        Dim boolizmCannotScale As Boolean = False
        If Not (NameScale Like FormСalcMarksAndSlopesOnCuvettes.ComboBox7.Text) Then
            Dim CannotDest As AnnotationScale = AcadFunction.FuncReadAnnotativeScaleDWGByName(FormСalcMarksAndSlopesOnCuvettes.ComboBox7.Text)
            If IsNothing(CannotDest) = False Then
                acCurDb.Cannoscale = CannotDest
                boolizmCannotScale = True
            End If
        End If

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверяем коридор
            Dim NameCoridor As String = FormСalcMarksAndSlopesOnCuvettes.ComboBox1.Text
            Dim CollectionCoridor As CorridorCollection = CivilDoc.CorridorCollection
            If CollectionCoridor.Count < 1 Then
                locker.Dispose()
                MsgBox("В чертеже нет проектных корридоров. Макро прервано!!!")
                Exit Sub
            End If
            'получаем нужный корридор
            Dim Coridor1 As Corridor = Nothing
            If IsNothing(NameCoridor) = False Then
                For Each id As ObjectId In CollectionCoridor
                    Coridor1 = tr.GetObject(id, OpenMode.ForRead) 'получаем доступ к корридору
                    If Coridor1.Name Like NameCoridor Then
                        Exit For
                    End If
                Next
            End If
            If IsNothing(Coridor1) = True Then
                MsgBox("Нет выбранного корридора!!!")
                locker.Dispose()
                Exit Sub
            End If
            '=============================================================================================
            Dim SurfUser As TinSurface = Nothing
            Dim NameSurf As String = FormСalcMarksAndSlopesOnCuvettes.ComboBox5.Text
            If IsNothing(NameSurf) = False Then
                If NameSurf.Trim.Length > 0 Then
                    For Each idSurf As ObjectId In CivilDoc.GetSurfaceIds
                        Dim SurfTemp As TinSurface = tr.GetObject(idSurf, OpenMode.ForRead)
                        If SurfTemp.Name Like NameSurf Then
                            SurfUser = SurfTemp
                            Exit For
                        End If
                    Next
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'смотрим коды линий
            Dim Kod1 As FeatureLineCodeInfo = Nothing
            Dim Kod2 As FeatureLineCodeInfo = Nothing
            Dim NameKod1 As String = FormСalcMarksAndSlopesOnCuvettes.ComboBox2.Text
            Dim NameKod2 As String = FormСalcMarksAndSlopesOnCuvettes.ComboBox3.Text
            If Coridor1.FeatureLineCodeInfos.Count > 0 Then
                For Each featureLine As FeatureLineCodeInfo In Coridor1.FeatureLineCodeInfos
                    Dim CodeNAme As String = featureLine.CodeName
                    If CodeNAme Like NameKod1 Then
                        Kod1 = featureLine
                    End If
                    If CodeNAme Like NameKod2 Then
                        Kod2 = featureLine
                    End If
                    If IsNothing(Kod1) = False And IsNothing(Kod2) = False Then Exit For
                Next
            End If
            If IsNothing(Kod1) = True Then
                If FormСalcMarksAndSlopesOnCuvettes.RadioButton4.Checked = True Or FormСalcMarksAndSlopesOnCuvettes.RadioButton5.Checked = True Then
                    MsgBox("Нет выбранных кодов характерных линий!!!")
                End If
            End If
            If IsNothing(Kod2) = True Then
                If FormСalcMarksAndSlopesOnCuvettes.RadioButton5.Checked = True Then
                    MsgBox("Нет выбранных кодов характерных линий!!!")
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '1. выбираем профиль
            Dim NameProfilActiv As String = FormСalcMarksAndSlopesOnCuvettes.ComboBox6.Text
            Dim ObjProfil1 As Autodesk.Civil.DatabaseServices.Profile = Nothing
            Try
                For Each id As ObjectId In acBlkTblRec
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    If TypeOf ent Is Autodesk.Civil.DatabaseServices.Profile Then
                        Dim Profile1 As Autodesk.Civil.DatabaseServices.Profile = ent
                        If Profile1.Name Like NameProfilActiv Then
                            ObjProfil1 = Profile1
                            Exit For
                        End If
                    End If
                Next
            Catch
            End Try
            If IsNothing(ObjProfil1) = True Then
                Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите линию профиля: ")
                Dim resBlk As PromptEntityResult = m_Editor.GetEntity(opt)
                Dim pvId As ObjectId = resBlk.ObjectId
                Dim obj As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(pvId, OpenMode.ForRead)
                If TypeOf obj Is Autodesk.Civil.DatabaseServices.Profile Then
                    ObjProfil1 = obj
                End If
            End If
            If IsNothing(ObjProfil1) = True Then
                locker.Dispose()
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ищем трассу по которой был создан профиль
            Dim IDAlign As ObjectId = ObjProfil1.AlignmentId
            'получаем доступ к трассе
            Dim Align As Alignment = tr.GetObject(IDAlign, OpenMode.ForRead)
            'получаем полилинию из трассы
            Dim IDPline As ObjectId = Align.GetPolyline
            Dim Pline As Polyline = New Polyline
            If IDPline.IsNull = False Then
                Pline = tr.GetObject(IDPline, OpenMode.ForRead)
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'массив для хранения точек характерной линии
            Dim ArrayPointFeatureLine1 As Double(,) = Nothing
            Dim ArrayPointFeatureLine2 As Double(,) = Nothing
            Dim CountArrayPointFeatureLine1 As Integer = 0
            Dim CountArrayPointFeatureLine2 As Integer = 0
            Dim ArrayFrofPKTemp As String(,) = Nothing 'массив для хранения пикетов переломов профиля и характерной линии
            'ищем характерные линии корридора
            If IsNothing(Kod1) = False Or IsNothing(Kod2) = False Then
                Dim Fline1 As CorridorFeatureLine = Nothing
                Dim Fline2 As CorridorFeatureLine = Nothing
                Dim PTPointCol1 As FeatureLinePointCollection = Nothing
                Dim PTPointCol2 As FeatureLinePointCollection = Nothing
                'смотрим характерные линии
                Dim blines As BaselineCollection = Coridor1.Baselines
                If Coridor1.Baselines.Count = 0 Then
                    MsgBox("В чертеже нет базовых линий. Макро прервано!!!")
                    Exit Sub
                End If
                For Each bLine As Baseline In blines
                    Dim MbaseFL As BaselineFeatureLines = bLine.MainBaselineFeatureLines
                    Dim FLMap As FeatureLineCollectionMap = MbaseFL.FeatureLineCollectionMap
                    For j As Integer = 0 To FLMap.Count - 1
                        Dim oFeatureLinesCol As FeatureLineCollection = FLMap.Item(j)
                        For k As Integer = 0 To oFeatureLinesCol.Count - 1
                            Dim featureline As CorridorFeatureLine = oFeatureLinesCol.Item(k)
                            Dim NameFl As String = featureline.CodeName
                            If NameFl Like Kod1.CodeName Then
                                Fline1 = featureline
                                Dim idFeatureLineCollection As ObjectIdCollection = Fline1.ExportAsPolyline3dCollection()
                                If idFeatureLineCollection.Count > 0 Then
                                    'заносим начало и конец характерных линий в массив
                                    For Each idPline3d As ObjectId In idFeatureLineCollection
                                        Dim Pline3d As Polyline3d = tr.GetObject(idPline3d, OpenMode.ForRead)
                                        'находим координаты точек начала и конца характерной линии
                                        Dim stPoint As Point3d = Pline3d.StartPoint
                                        Dim EndPoint As Point3d = Pline3d.EndPoint
                                        'находим их пикетажное положение начала полилинии
                                        Dim PKPointStart As Double = 0
                                        Dim OffsetPointStart As Double = 0
                                        Try
                                            Align.StationOffset(stPoint.X, stPoint.Y, PKPointStart, OffsetPointStart)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        End Try
                                        'находим их пикетажное положение конца полилинии
                                        Dim PKPointEnd As Double = 0
                                        Dim OffsetPointEnd As Double = 0
                                        Try
                                            Align.StationOffset(EndPoint.X, EndPoint.Y, PKPointEnd, OffsetPointEnd)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        End Try
                                        'заносим в массив начало и конец характерной линии
                                        If IsArray(ArrayFrofPKTemp) = False Then
                                            'массив вида (Пикет, начало или конец)
                                            ReDim Preserve ArrayFrofPKTemp(1, 1)
                                            ArrayFrofPKTemp(0, 0) = PKPointStart 'пикет
                                            ArrayFrofPKTemp(1, 0) = 0
                                            ArrayFrofPKTemp(0, 1) = PKPointEnd
                                            ArrayFrofPKTemp(1, 1) = 1
                                        Else
                                            Dim Count As Integer = ArrayFrofPKTemp.GetUpperBound(1) + 2
                                            ReDim Preserve ArrayFrofPKTemp(1, Count)
                                            ArrayFrofPKTemp(0, Count - 1) = PKPointStart 'пикет
                                            ArrayFrofPKTemp(1, Count - 1) = 0
                                            ArrayFrofPKTemp(0, Count) = PKPointEnd
                                            ArrayFrofPKTemp(1, Count) = 1
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'заносим пикеты характерных точек профиля в диапазоне начала и конца характерной линии
                                        Dim FindVertPointFrofile As Boolean = FuncAlignment.FuncReadVerticalGeometryToStatinName(ObjProfil1, ArrayFrofPKTemp, PKPointStart, PKPointEnd)
                                        'заносим все точки перелома в массив
                                        If Pline3d.Length > 0 Then
                                            For Each VertID As ObjectId In Pline3d
                                                Dim vert3d As PolylineVertex3d = tr.GetObject(VertID, OpenMode.ForRead)
                                                ReDim Preserve ArrayPointFeatureLine1(2, CountArrayPointFeatureLine1)
                                                'находим их пикетажное положение начала полилинии
                                                Dim PKPoint As Double = 0
                                                Dim OffsetPoint As Double = 0
                                                Try
                                                    Align.StationOffset(vert3d.Position.X, vert3d.Position.Y, PKPoint, OffsetPoint)
                                                    ArrayPointFeatureLine1(0, CountArrayPointFeatureLine1) = PKPoint
                                                    ArrayPointFeatureLine1(1, CountArrayPointFeatureLine1) = OffsetPoint
                                                    ArrayPointFeatureLine1(2, CountArrayPointFeatureLine1) = vert3d.Position.Z
                                                    CountArrayPointFeatureLine1 += 1
                                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                End Try
                                            Next
                                        End If
                                    Next
                                End If
                            End If
                            If NameFl Like Kod2.CodeName And FormСalcMarksAndSlopesOnCuvettes.RadioButton5.Checked = True Then
                                Fline2 = featureline
                                Dim idFeatureLineCollection As ObjectIdCollection = Fline2.ExportAsPolyline3dCollection()
                                If idFeatureLineCollection.Count > 0 Then
                                    For Each idPline3d As ObjectId In idFeatureLineCollection
                                        Dim Pline3d As Polyline3d = tr.GetObject(idPline3d, OpenMode.ForRead)
                                        For Each VertID As ObjectId In Pline3d
                                            Dim vert3d As PolylineVertex3d = tr.GetObject(VertID, OpenMode.ForRead)
                                            ReDim Preserve ArrayPointFeatureLine2(2, CountArrayPointFeatureLine2)
                                            'находим их пикетажное положение начала полилинии
                                            Dim PKPoint As Double = 0
                                            Dim OffsetPoint As Double = 0
                                            Try
                                                Align.StationOffset(vert3d.Position.X, vert3d.Position.Y, PKPoint, OffsetPoint)
                                                ArrayPointFeatureLine2(0, CountArrayPointFeatureLine2) = PKPoint
                                                ArrayPointFeatureLine2(1, CountArrayPointFeatureLine2) = OffsetPoint
                                                ArrayPointFeatureLine2(2, CountArrayPointFeatureLine2) = vert3d.Position.Z
                                                CountArrayPointFeatureLine2 += 1
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        Next
                                    Next
                                End If
                            End If
                        Next
                    Next
                Next
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'загружаем пикеты с шагом
            'если есть шаг
            Dim CountArrayProfPK As Integer = ArrayFrofPKTemp.Length
            If FormСalcMarksAndSlopesOnCuvettes.CheckBox3.Checked = True Then
                'определяем старт
                If IsArray(ArrayFrofPKTemp) = True Then
                    CountArrayProfPK = ArrayFrofPKTemp.Length
                    'получаем начальное значение цикла
                    Dim ChagRound As Double = FormСalcMarksAndSlopesOnCuvettes.NumericUpDown3.Value
                    Dim IntPKStart As Double = ObjProfil1.StartingStation + ChagRound
                    IntPKStart = Math.Ceiling(ObjProfil1.StartingStation / ChagRound) 'делим на шаг и округляем вверх
                    IntPKStart = IntPKStart * ChagRound
                    If IntPKStart = ObjProfil1.StartingStation Then IntPKStart = IntPKStart + ChagRound
                    'вычисляем конечное значение цикла
                    Dim IntPKEnd As Double = ObjProfil1.EndingStation 'конечное значение цикла
                    IntPKEnd = Math.Floor(ObjProfil1.EndingStation / ChagRound)
                    IntPKEnd = IntPKEnd * ChagRound
                    If IntPKEnd = ObjProfil1.EndingStation Then IntPKEnd = IntPKEnd - ChagRound
                    'записываем в массив начальное значение профиля
                    ReDim Preserve ArrayFrofPKTemp(1, CountArrayProfPK)
                    ArrayFrofPKTemp(0, CountArrayProfPK) = ObjProfil1.StartingStation
                    CountArrayProfPK += 1
                    For i As Double = IntPKStart To IntPKEnd Step ChagRound
                        ReDim Preserve ArrayFrofPKTemp(1, CountArrayProfPK)
                        ArrayFrofPKTemp(0, CountArrayProfPK) = i
                        CountArrayProfPK += 1
                    Next i
                    'записываем последнее значение
                    ReDim Preserve ArrayFrofPKTemp(1, CountArrayProfPK)
                    ArrayFrofPKTemp(0, CountArrayProfPK) = ObjProfil1.EndingStation
                End If
            End If
            'сортируем массив
            'сортируем массив по первому элементу
            Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayFrofPKTemp, 0)
            'делаем новый массив, удаляем повторяющиеся элементы
            Dim ArrayFrofPK As Double(,) = Nothing
            Dim CountArrayProfPKNew As Integer = 0
            If IsArray(ArrayFrofPKTemp) = True Then
                For i As Integer = 0 To ArrayFrofPKTemp.GetUpperBound(1) - 1
                    Dim PK1 As Double = Val(ArrayFrofPKTemp(0, i))
                    Dim PK2 As Double = Val(ArrayFrofPKTemp(0, i + 1))
                    If PK1 <> PK2 Then
                        ReDim Preserve ArrayFrofPK(1, CountArrayProfPKNew)
                        ArrayFrofPK(0, CountArrayProfPKNew) = PK1
                        CountArrayProfPKNew += 1
                    End If
                Next
                Dim PKEnd As Double = ArrayFrofPKTemp(0, ArrayFrofPKTemp.GetUpperBound(1))
                ReDim Preserve ArrayFrofPK(1, CountArrayProfPKNew)
                ArrayFrofPK(0, CountArrayProfPKNew) = PKEnd
            End If
            'начинаем перебирать массив
            Dim ArrayCoord As Double(,) = Nothing 'массив точек для вставки блока Уклон-Расстояние
            Dim CountArrayCoord As Integer = 0
            If IsArray(ArrayFrofPK) = True Then
                For i As Integer = 0 To ArrayFrofPK.GetUpperBound(1)
                    Dim PK As Double = ArrayFrofPK(0, i)
                    'ищем высоту
                    Dim Elev As Double = 0
                    Try
                        Elev = ObjProfil1.ElevationAt(PK)
                    Catch ex As ArgumentException
                    End Try
                    Dim PKIns As Double = 0 'пикет встаки точки
                    Dim PKINSOFFSET As Double = 0
                    Dim PKPoint1 As Double = 0
                    Dim PKOffset1 As Double = 0
                    Dim PKHeight1 As Double = 0
                    If IsArray(ArrayPointFeatureLine1) = True Then
                        For i1 As Integer = 0 To ArrayPointFeatureLine1.GetUpperBound(1) - 1
                            '1. первая точка
                            Dim Fpt1Stat As Double = ArrayPointFeatureLine1(0, i1) 'забираем пикет из массива
                            Dim FPt1Off As Double = ArrayPointFeatureLine1(1, i1) 'забираем смещение из массива
                            Dim FPt1Height As Double = ArrayPointFeatureLine1(2, i1) 'забираем высоту точки из массива
                            '2. вторая точка
                            Dim Fpt2Stat As Double = ArrayPointFeatureLine1(0, i1 + 1) 'забираем пикет из массива
                            Dim FPt2Off As Double = ArrayPointFeatureLine1(1, i1 + 1) 'забираем смещение из массива
                            Dim FPt2Height As Double = ArrayPointFeatureLine1(2, i1) 'забираем высоту точки из массива
                            If PK > Fpt1Stat And PK < Fpt2Stat Then
                                PKPoint1 = Fpt1Stat + (PK - Fpt1Stat)
                                PKOffset1 = (FPt1Off + FPt2Off) / 2
                                PKHeight1 = (FPt1Height + FPt2Height) / 2
                                Exit For
                            ElseIf PK < Fpt1Stat And PK > Fpt2Stat Then
                                PKPoint1 = Fpt2Stat + (PK - Fpt2Stat)
                                PKOffset1 = (FPt1Off + FPt2Off) / 2
                                PKHeight1 = (FPt1Height + FPt2Height) / 2
                                Exit For
                            ElseIf Fpt1Stat = PK Then
                                PKPoint1 = Fpt1Stat
                                PKOffset1 = FPt1Off
                                PKHeight1 = FPt1Height
                                Exit For
                            ElseIf Fpt2Stat = PK Then
                                PKPoint1 = Fpt2Stat
                                PKOffset1 = FPt2Off
                                PKHeight1 = FPt2Height
                                Exit For
                            End If
                        Next i1
                    End If
                    '================================================================
                    'ищем точку на 1 характерной линии
                    'для этого перебираем массив из точек
                    Dim PKPoint2 As Double = 0
                    Dim PKOffset2 As Double = 0
                    Dim PKHeight2 As Double = 0
                    If IsArray(ArrayPointFeatureLine2) = True Then
                        For i1 As Integer = 0 To ArrayPointFeatureLine2.GetUpperBound(1) - 1
                            Dim Fpt1Stat As Double = ArrayPointFeatureLine2(0, i1) 'забираем пикет из массива
                            Dim FPt1Off As Double = ArrayPointFeatureLine2(1, i1) 'забираем смещение из массива
                            Dim FPt1Height As Double = ArrayPointFeatureLine2(2, i1) 'забираем высоту точки из массива
                            '2. вторая точка
                            Dim Fpt2Stat As Double = ArrayPointFeatureLine2(0, i1 + 1) 'забираем пикет из массива
                            Dim FPt2Off As Double = ArrayPointFeatureLine2(1, i1 + 1) 'забираем смещение из массива
                            Dim FPt2Height As Double = ArrayPointFeatureLine2(2, i1) 'забираем высоту точки из массива
                            If PK > Fpt1Stat And PK < Fpt2Stat Then
                                PKPoint2 = Fpt1Stat + (PK - Fpt1Stat)
                                PKOffset2 = (FPt1Off + FPt2Off) / 2
                                PKHeight2 = (FPt1Height + FPt2Height) / 2
                                Exit For
                            ElseIf PK < Fpt1Stat And PK > Fpt2Stat Then
                                PKPoint2 = Fpt2Stat + (PK - Fpt2Stat)
                                PKOffset2 = (FPt1Off + FPt2Off) / 2
                                PKHeight2 = (FPt1Height + FPt2Height) / 2
                                Exit For
                            ElseIf Fpt1Stat = PK Then
                                PKPoint2 = Fpt1Stat
                                PKOffset2 = FPt1Off
                                PKHeight2 = FPt1Height
                                Exit For
                            ElseIf Fpt2Stat = PK Then
                                PKPoint2 = Fpt2Stat
                                PKOffset2 = FPt2Off
                                PKHeight2 = FPt2Height
                                Exit For
                            End If
                        Next i1
                    End If

                    '==========================================================================================
                    'усредняем значения для смещения
                    If PKOffset1 <> 0 And PKOffset2 <> 0 Then
                        PKINSOFFSET = (PKOffset1 + PKOffset2) / 2
                    ElseIf PKOffset1 = 0 And PKOffset2 <> 0 Then
                        PKINSOFFSET = PKOffset2
                    ElseIf PKOffset1 <> 0 And PKOffset2 = 0 Then
                        PKINSOFFSET = PKOffset1
                    Else
                        PKINSOFFSET = 0
                    End If
                    'усредняем значения для пикета
                    If PKPoint1 <> 0 And PKPoint2 <> 0 Then
                        PKIns = (PKPoint1 + PKPoint2) / 2
                    ElseIf PKPoint1 = 0 And PKPoint2 <> 0 Then
                        PKIns = PKPoint2
                    ElseIf PKPoint1 <> 0 And PKPoint2 = 0 Then
                        PKIns = PKPoint1
                    Else
                        PKIns = PK
                    End If

                    '================================================================================================
                    'смотрим смещение
                    If PKINSOFFSET = 0 Then
                        If FormСalcMarksAndSlopesOnCuvettes.RadioButton3.Checked = False Then
                            GoTo lineEnd
                        End If
                    End If
                    '================================================================================================
                    'вычисляем координаты
                    Try
                        'находим координаты этой точки на кювете
                        Dim xkuv As Double = -1
                        Dim Ykuv As Double = -1
                        Try
                            Align.PointLocation(PKIns, PKINSOFFSET, xkuv, Ykuv)
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                        End Try
                        'находим координаты этой точки на трассе
                        Dim xalign As Double = -1
                        Dim Yalign As Double = -1
                        Try
                            Align.PointLocation(PKIns, 0, xalign, Yalign)
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                        End Try
                        'форматируем высоту
                        Elev = Math.Round(Elev, RoundZn)
                        ReDim Preserve ArrayCoord(3, CountArrayCoord)
                        ArrayCoord(0, CountArrayCoord) = xkuv
                        ArrayCoord(1, CountArrayCoord) = Ykuv
                        ArrayCoord(2, CountArrayCoord) = Elev
                        ArrayCoord(3, CountArrayCoord) = PKIns
                        CountArrayCoord += 1
                        If FormСalcMarksAndSlopesOnCuvettes.RadioButton2.Checked = True Then
                            If xkuv <> -1 And Ykuv <> -1 Then
                                Dim MLeader1 As MLeader = New MLeader
                                MLeader1.SetDatabaseDefaults()
                                MLeader1.MLeaderStyle = IDMLeaderStyle
                                MLeader1.Layer = NameLayerMLeader
                                Dim leaderIndex As Integer = MLeader1.AddLeader()
                                Dim leaderLineIndex As Integer = MLeader1.AddLeaderLine(leaderIndex)
                                MLeader1.AddFirstVertex(leaderLineIndex, New Point3d(xkuv, Ykuv, 0))
                                If xalign <> -1 And Yalign <> -1 Then
                                    Dim Pt1 As Point2d = New Point2d(xkuv, Ykuv)
                                    Dim Pt2 As Point2d = New Point2d(xalign, Yalign)
                                    Dim DirN As Double = Pt2.GetVectorTo(Pt1).Angle
                                    DirN = DirN + AngleLabel
                                    Dim XPolka As Double = xkuv + OffsetLabel * Math.Cos(DirN)
                                    Dim YPolka As Double = Ykuv + OffsetLabel * Math.Sin(DirN)
                                    MLeader1.AddLastVertex(leaderLineIndex, New Point3d(XPolka, YPolka, 0))
                                Else
                                    MLeader1.AddLastVertex(leaderLineIndex, New Point3d(xkuv + 15, Ykuv + 15, 0))
                                End If
                                MLeader1.ContentType = ContentType.MTextContent
                                Dim acMText As MText = New MText()
                                acMText.SetDatabaseDefaults()
                                If RoundZn = 0 Then
                                    acMText.SetContentsRtf(Elev.ToString)
                                ElseIf RoundZn = 1 Then
                                    acMText.SetContentsRtf(String.Format(CultureInfo.InvariantCulture, "{0:0.0}", Elev))
                                ElseIf RoundZn = 2 Then
                                    acMText.SetContentsRtf(String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Elev))
                                ElseIf RoundZn = 3 Then
                                    acMText.SetContentsRtf(String.Format(CultureInfo.InvariantCulture, "{0:0.000}", Elev))
                                ElseIf RoundZn = 4 Then
                                    acMText.SetContentsRtf(String.Format(CultureInfo.InvariantCulture, "{0:0.0000}", Elev))
                                End If
                                MLeader1.MText = acMText
                                acBlkTblRec.AppendEntity(MLeader1)
                                tr.AddNewlyCreatedDBObject(MLeader1, True)
                            End If
                        ElseIf FormСalcMarksAndSlopesOnCuvettes.RadioButton1.Checked = True Then
                            If IsNothing(SurfUser) = False Then
                                If SurfUser.Id.IsNull = False Then
                                    Dim PtCoGo As Point3d = New Point3d(xkuv, Ykuv, Elev)
                                    Dim surfaceElevnLblId As ObjectId = SurfaceElevationLabel.Create(SurfUser.Id, PtCoGo.Convert2d(New Plane))
                                    Dim label As SurfaceElevationLabel = surfaceElevnLblId.GetObject(OpenMode.ForWrite)
                                    Dim nameSt2 As String = label.StyleName
                                    If xalign <> -1 And Yalign <> -1 Then
                                        Dim Pt1 As Point2d = New Point2d(xkuv, Ykuv)
                                        Dim Pt2 As Point2d = New Point2d(xalign, Yalign)
                                        Dim DirN As Double = Pt2.GetVectorTo(Pt1).Angle
                                        DirN = DirN + AngleLabel
                                        Dim XPolka As Double = xkuv + OffsetLabel * Math.Cos(DirN)
                                        Dim YPolka As Double = Ykuv + OffsetLabel * Math.Sin(DirN)
                                        Dim NewLBLoc As Point3d = New Point3d(XPolka, YPolka, 0)
                                        label.LabelLocation = NewLBLoc
                                    Else
                                        Dim NewLBLoc As Point3d = New Point3d(xkuv + 15, Ykuv + 15, 0)
                                        label.LabelLocation = NewLBLoc
                                    End If
                                End If
                            End If

                        End If
                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                    End Try
lineEnd:
                Next i
            End If
            '=========================================================================================================
            'расставляем уклоны
            If IsArray(ArrayCoord) = True Then
                If ArrayCoord.GetUpperBound(1) > 1 Then
                    'вставляем определение блока
                    For m1 As Integer = 0 To ArrayCoord.GetUpperBound(1) - 1
                        Dim pt1 As Point2d = New Point2d(ArrayCoord(0, m1), ArrayCoord(1, m1))
                        Dim pt2 As Point2d = New Point2d(ArrayCoord(0, m1 + 1), ArrayCoord(1, m1 + 1))
                        'находим пикеты этих двух точек
                        Dim PKStart1 As Double = ArrayCoord(3, m1)
                        Dim PKEnd1 As Double = ArrayCoord(3, m1 + 1)
                        Dim Dist As Double = Math.Abs(PKEnd1 - PKStart1)
                        'вычисляем координаты точки вставки
                        Dim Xs As Double = (pt1.X + pt2.X) / 2
                        Dim Ys As Double = (pt1.Y + pt2.Y) / 2
                        Dim PtInsBlk As Point3d = New Point3d(Xs, Ys, 0)
                        Dim Vect As Vector2d = pt1.GetVectorTo(pt2)
                        Dim angle As Double = Vect.Angle
                        Dim flag As Integer = 0
                        Select Case angle
                            Case Math.PI / 2 To Math.PI
                                angle = Math.PI + angle
                                flag = 1
                            Case Math.PI To (3 * Math.PI) / 2
                                angle = Math.PI + angle
                                flag = 1
                            Case Else
                                angle = angle
                        End Select
                        'вставляем блок в базу
                        If acBlkTbl.Has(NameBlk) = True Then
                            Dim acBlk As BlockReference = AcadFunction.FuncInsBlkByArrayProperties(PtInsBlk, ArrayPropBlk, angle)
                            'вычисляем уклон
                            Dim Z1 As Double = ArrayCoord(2, m1)
                            Dim Z2 As Double = ArrayCoord(2, m1 + 1)
                            Dim IR As Double = (Z2 - Z1) / Dist
                            Dim i As Double = Math.Round(IR * 1000, 1)
                            i = Math.Abs(i)
                            Dim istr As String = String.Format(CultureInfo.InvariantCulture, "{0:0.0}", i)
                            Dim ArrayZn As String(,) = Nothing
                            ReDim ArrayZn(1, 1)
                            ArrayZn(0, 0) = "УКЛОН"
                            ArrayZn(1, 0) = istr
                            ArrayZn(0, 1) = "РАССТОЯНИЕ"
                            'форматируем расстояние
                            If RoundZn = 0 Then
                                Dist = Math.Round(Dist, 0)
                                ArrayZn(1, 1) = Dist
                            ElseIf RoundZn = 1 Then
                                Dist = Math.Round(Dist, 1)
                                ArrayZn(1, 1) = String.Format(CultureInfo.InvariantCulture, "{0:0.0}", Dist)
                            ElseIf RoundZn = 2 Then
                                Dist = Math.Round(Dist, 2)
                                ArrayZn(1, 1) = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Dist)
                            ElseIf RoundZn = 3 Then
                                Dist = Math.Round(Dist, 3)
                                ArrayZn(1, 1) = String.Format(CultureInfo.InvariantCulture, "{0:0.000}", Dist)
                            ElseIf RoundZn = 4 Then
                                Dist = Math.Round(Dist, 4)
                                ArrayZn(1, 1) = String.Format(CultureInfo.InvariantCulture, "{0:0.0000}", Dist)
                            End If
                            'получаем доступ к блоку
                            If IsNothing(acBlk) = False Then
                                If acBlk.Id.IsNull = False Then
                                    If acBlk.IsWriteEnabled = False Then
                                        acBlk.UpgradeOpen()
                                    End If
                                    'записываем атрибуты
                                    Dim boolWriteAttrib As Boolean = AcadFunction.FuncWriteAllBlockAttr(acBlk, ArrayZn)
                                    'смотрим отраженное состояние
                                    If acBlk.IsDynamicBlock = True Then
                                        Dim blockDyn As BlockTableRecord = tr.GetObject(acBlk.DynamicBlockTableRecord, OpenMode.ForWrite)
                                        Dim col As DynamicBlockReferencePropertyCollection = acBlk.DynamicBlockReferencePropertyCollection
                                        For Each dp As DynamicBlockReferenceProperty In col
                                            Dim values As Object() = dp.GetAllowedValues()
                                            Dim Name As String = dp.PropertyName
                                            If Name Like "Направление стрелки" Then
                                                If flag = 1 Then
                                                    dp.Value = values(1)
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            End If

                        End If
                    Next m1
                End If
lineend1:
            End If
            If boolizmCannotScale = True Then
                Dim CannotDest As AnnotationScale = AcadFunction.FuncReadAnnotativeScaleDWGByName(NameScale)
                If IsNothing(CannotDest) = False Then
                    acCurDb.Cannoscale = CannotDest
                End If
            End If
            tr.Commit()
            locker.Dispose()
        End Using
    End Sub
End Class
