Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
'=============================================================================================
'работа с коридорами (ведомость по кодам) ведомость планировки
Public Class ListPlanningStatement
    <CommandMethod("ListPlanStatement")>
    Public Sub ListPlanStatement()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim Db As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
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

        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'вызываем форму
                'запуск диалогового окна
                Dim FormPlanStatement As PlanStatement = New PlanStatement
                FormPlanStatement.ShowDialog()
                If FormPlanStatement.boolCoridor = False Then
                    Exit Sub
                End If
                Dim RoundZn As Integer = FormPlanStatement.NumericUpDown2.Value 'точность вычислений
                Dim DopuskAssembly As Double = FormPlanStatement.NumericUpDown1.Value 'допуск для поиска конструкций
                '1. Получаем нужный коридор
                Dim NameCoridor As String = FormPlanStatement.ComboBox1.Text.Trim
                If NameCoridor.Trim.Length = 0 Then
                    locker.Dispose()
                    Exit Sub
                End If
                Dim IdCoridor As ObjectId = CivilDoc.CorridorCollection.Item(NameCoridor)
                Dim ObjCoridor As Corridor = tr.GetObject(IdCoridor, OpenMode.ForRead) 'получаем доступ к коридору
                '2. Получаем базовую линию
                Dim NameBaseLine As String = FormPlanStatement.ComboBox2.Text.Trim
                If NameBaseLine Like "" Then
                    locker.Dispose()
                    Exit Sub
                End If
                Dim ObjBaseLine As Baseline = ObjCoridor.Baselines.Item(NameBaseLine)
                '3. Получаем начальный и конечный пикеты
                Dim PKStart As Double = FormPlanStatement.PKStart
                Dim PKEnd As Double = FormPlanStatement.PKEnd
                Dim ErrorLine As Double = Math.Abs(PKEnd - PKStart)
                '4. Получаем трассу
                Dim AlignID As ObjectId = ObjBaseLine.AlignmentId
                If IsNothing(AlignID) = True Then
                    locker.Dispose()
                    Exit Sub
                End If
                Dim ObjAlign As Alignment = tr.GetObject(AlignID, OpenMode.ForRead) 'получаем доступ к трассе
                Dim sqwPiketCollection As StationEquationCollection = ObjAlign.StationEquations
                '5. Достаем коды
                Dim Kod1 As String = FormPlanStatement.ComboBox3.Text
                Dim Kod2 As String = FormPlanStatement.ComboBox4.Text
                If Kod1 Like Kod2 And Not (Kod1 Like "Выбрать все") And Not (Kod2 Like "Выбрать все") Then
                    MsgBox("Необходимо выбрать разные коды!!! Макро прервано.")
                    locker.Dispose()
                    Exit Sub
                End If
                '7. Имя шаблона
                Dim NameFileTemplate As String = FormPlanStatement.ComboBox7.Tag & FormPlanStatement.ComboBox7.Text
                '8. переменные
                Dim PointAssembl As String(,) = Nothing 'массив с элементами (0-пикет, 1-смещение, 2-высота, 
                Dim CountPointAssembl As Integer = 0 'счетчик элементов
                Dim CountAssembl As Integer = 0 'счетчик конструкций
                Dim PKAssemblError As Integer = 0 'начальный пикет в случае ошибок
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ПОИСК ПО КОНСТРУКЦИЯМ (выводит все конструкции)
                'по 2 кодам выбранным пользователем
                If FormPlanStatement.RadioButton1.Checked = True And Not (FormPlanStatement.ComboBox3.Text Like "Выбрать все") Then
                    'получаем коллекцию регионов входящих в базовую линию
                    Dim RegionColl As BaselineRegionCollection = ObjBaseLine.BaselineRegions
                    For Each BaseMyRegion As BaselineRegion In RegionColl
                        'если регион находится в диапазоне пикетов пользователя продолжаем
                        If BaseMyRegion.StartStation >= PKStart And BaseMyRegion.EndStation <= PKEnd Then
                            'если регион рабочий
                            If BaseMyRegion.NeedsProcessing = True Then
                                Dim AppAssemblColl As AppliedAssemblyCollection = BaseMyRegion.AppliedAssemblies
                                If AppAssemblColl.Count > 0 Then
                                    For Each appAssembly As AppliedAssembly In AppAssemblColl
                                        'забираем 1 код
                                        Dim PKAssemly As Double = 0 'пикет по умолчанию
                                        Dim PointCalc1 As CalculatedPointCollection = appAssembly.GetPointsByCode(Kod1)
                                        'вычисляем координаты первой точки
                                        Dim StationPt1 As Double = 0
                                        Dim OffsetPt1 As Double = 0
                                        Dim ElevationPt1 As Double = 0
                                        Dim PtAssembly1 As Point3d = Nothing
                                        Dim PtAbsAssembly1 As Point3d = Nothing
                                        Dim appSubColl As AppliedSubassemblyCollection = appAssembly.GetAppliedSubassemblies
                                        If appSubColl.Count > 0 Then
                                            For i1 As Integer = 0 To appSubColl.Count - 1
                                                Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i1)
                                                Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                                Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                                If pointCol.Count > 0 Then
                                                    For i2 As Integer = 0 To pointCol.Count - 1
                                                        Dim ptAssembl As CalculatedPoint = pointCol.Item(i2)
                                                        If ptAssembl.CorridorCodes.Count > 0 Then
                                                            For i3 As Integer = 0 To ptAssembl.CorridorCodes.Count - 1
                                                                Dim strCode As String = ptAssembl.CorridorCodes.Item(i3)
                                                                If strCode Like Kod1 Then
                                                                    StationPt1 = ptAssembl.StationOffsetElevationToBaseline.X
                                                                    OffsetPt1 = ptAssembl.StationOffsetElevationToBaseline.Y
                                                                    ElevationPt1 = ptAssembl.StationOffsetElevationToBaseline.Z
                                                                    PtAssembly1 = New Point3d(StationPt1, OffsetPt1, ElevationPt1)
                                                                    PtAbsAssembly1 = blineSubAssembl.StationOffsetElevationToXYZ(PtAssembly1)
                                                                    PKAssemly = StationPt1
                                                                    GoTo LineFind1
                                                                End If
                                                            Next i3
                                                        End If
                                                    Next i2
                                                End If
                                            Next i1
                                        End If
LineFind1:
                                        '================================================================================================
                                        'читаем второй код
                                        '================================================================================================
                                        Dim StationPt2 As Double = 0
                                        Dim OffsetPt2 As Double = 0
                                        Dim ElevationPt2 As Double = 0
                                        Dim PtAssembly2 As Point3d = Nothing
                                        Dim PtAbsAssembly2 As Point3d = Nothing
                                        'вычисляем координаты второй точки
                                        Dim PointCalc2 As CalculatedPointCollection = appAssembly.GetPointsByCode(Kod2)
                                        If appSubColl.Count > 0 Then
                                            For i1 As Integer = 0 To appSubColl.Count - 1
                                                Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i1)
                                                Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                                Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                                If pointCol.Count > 0 Then
                                                    For i2 As Integer = 0 To pointCol.Count - 1
                                                        Dim ptAssembl As CalculatedPoint = pointCol.Item(i2)
                                                        If ptAssembl.CorridorCodes.Count > 0 Then
                                                            For i3 As Integer = 0 To ptAssembl.CorridorCodes.Count - 1
                                                                Dim strCode As String = ptAssembl.CorridorCodes.Item(i3)
                                                                If strCode Like Kod2 Then
                                                                    StationPt2 = ptAssembl.StationOffsetElevationToBaseline.X
                                                                    OffsetPt2 = ptAssembl.StationOffsetElevationToBaseline.Y
                                                                    ElevationPt2 = ptAssembl.StationOffsetElevationToBaseline.Z
                                                                    PtAssembly2 = New Point3d(StationPt2, OffsetPt2, ElevationPt2)
                                                                    PtAbsAssembly2 = blineSubAssembl.StationOffsetElevationToXYZ(PtAssembly2)
                                                                    GoTo LineFind2
                                                                End If
                                                            Next i3
                                                        End If
                                                    Next i2
                                                End If
                                            Next i1
                                        End If
LineFind2:
                                        If PKAssemly < StationPt2 Then
                                            PKAssemly = StationPt2
                                        End If
                                        If StationPt1 >= PKStart Then
                                            'заносим значения в массив
                                            ReDim Preserve PointAssembl(9, CountPointAssembl)
                                            PointAssembl(0, CountPointAssembl) = PKAssemly
                                            PointAssembl(1, CountPointAssembl) = PtAbsAssembly1.X
                                            PointAssembl(2, CountPointAssembl) = PtAbsAssembly1.Y
                                            PointAssembl(3, CountPointAssembl) = PtAbsAssembly1.Z
                                            PointAssembl(4, CountPointAssembl) = PtAbsAssembly2.X
                                            PointAssembl(5, CountPointAssembl) = PtAbsAssembly2.Y
                                            PointAssembl(6, CountPointAssembl) = PtAbsAssembly2.Z
                                            'расстояние между точками (горизонтальное расстояние)
                                            Dim Pt1 As Point2d = New Point2d(PtAbsAssembly1.X, PtAbsAssembly1.Y)
                                            Dim Pt2 As Point2d = New Point2d(PtAbsAssembly2.X, PtAbsAssembly2.Y)
                                            'горизонтальное расстояние
                                            Dim vect2d As Vector2d = Pt1.GetVectorTo(Pt2)
                                            PointAssembl(7, CountPointAssembl) = vect2d.Length
                                            'вертикальное расстояние
                                            Dim vect3d As Vector3d = PtAbsAssembly1.GetVectorTo(PtAbsAssembly2)
                                            PointAssembl(8, CountPointAssembl) = vect3d.Length
                                            'превышение
                                            PointAssembl(9, CountPointAssembl) = PtAbsAssembly1.Z - PtAbsAssembly2.Z
                                            CountPointAssembl += 1
                                        Else
                                            ReDim Preserve PointAssembl(9, CountPointAssembl)
                                            PointAssembl(0, CountPointAssembl) = PKAssemly
                                            PointAssembl(1, CountPointAssembl) = 0
                                            PointAssembl(2, CountPointAssembl) = 0
                                            PointAssembl(3, CountPointAssembl) = 0
                                            PointAssembl(4, CountPointAssembl) = 0
                                            PointAssembl(5, CountPointAssembl) = 0
                                            PointAssembl(6, CountPointAssembl) = 0
                                            PointAssembl(7, CountPointAssembl) = 0
                                            PointAssembl(8, CountPointAssembl) = 0
                                            PointAssembl(9, CountPointAssembl) = 0
                                            CountPointAssembl += 1
                                        End If
                                        CountAssembl += 1
                                    Next
                                End If
                            End If
                        End If
                    Next
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'поиск по секциям
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                ElseIf FormPlanStatement.RadioButton2.Checked = True And Not (FormPlanStatement.ComboBox3.Text Like "Выбрать все") Then
                    ' Получаем доступ к группам секций
                    Dim SampleLineGroupCollection As ObjectIdCollection = ObjAlign.GetSampleLineGroupIds
                    Dim IDSectionCollection As ObjectIdCollection = New ObjectIdCollection
                    If SampleLineGroupCollection.Count > 0 Then
                        For Each sampleLineId As ObjectId In SampleLineGroupCollection
                            Dim sampleGroupLine As SampleLineGroup = tr.GetObject(sampleLineId, OpenMode.ForRead)
                            If sampleGroupLine.Name Like FormPlanStatement.ComboBox5.Text.Trim Then
                                IDSectionCollection = sampleGroupLine.GetSampleLineIds
                                Exit For
                            End If
                        Next
                    End If
                    'начинаем перебор секций
                    Dim ArrayStationSampleLine As Double(,) = Nothing
                    Dim CountStationSampleLine As Integer = 0
                    If IDSectionCollection.Count = 0 Then
                        MsgBox("Коллекция линий сечений пуста. Макро прервано!!!")
                        Exit Sub
                    End If
                    If IDSectionCollection.Count > 0 Then
                        For Each idSample As ObjectId In IDSectionCollection
                            Dim SamplLine As SampleLine = tr.GetObject(idSample, OpenMode.ForRead)
                            Dim PK As Double = SamplLine.Station
                            If PK >= PKStart And PK <= PKEnd Then
                                ReDim Preserve ArrayStationSampleLine(1, CountStationSampleLine)
                                ArrayStationSampleLine(0, CountStationSampleLine) = SamplLine.Station
                                ArrayStationSampleLine(1, CountStationSampleLine) = SamplLine.Number
                                CountStationSampleLine += 1
                            End If
                        Next
                    End If
                    '==================================================================================================
                    'начинаем доставать конструкции
                    '==================================================================================================
                    Dim PKStat As Double = 0
                    Dim NumbStat As Integer = 0
                    Dim Dopusk As Double = FormPlanStatement.NumericUpDown1.Value 'допуск для поиска конструкций
                    If IsArray(ArrayStationSampleLine) = True Then
                        For i As Integer = 0 To ArrayStationSampleLine.GetUpperBound(1)
                            PKStat = ArrayStationSampleLine(0, i)
                            NumbStat = ArrayStationSampleLine(1, i)
                            Dim appAssembly As AppliedAssembly = Nothing
                            Try
                                'получаем доступ к конструкции
                                appAssembly = ObjBaseLine.GetAppliedAssemblyAtStation(PKStat)
                            Catch ex As System.ArgumentOutOfRangeException
                                'пытаемся найти ближайшую конструкцию
                                appAssembly = FuncAlignment.FuncFindBaseLineAssembly(ObjBaseLine, PKStat, ErrorLine, Dopusk)
                                If IsNothing(appAssembly) = True Then
                                    ReDim Preserve PointAssembl(7, CountPointAssembl)
                                    PointAssembl(0, CountPointAssembl) = PKStat
                                    PointAssembl(1, CountPointAssembl) = "Нет конструкции"
                                    CountPointAssembl += 1
                                    GoTo line1
                                End If
                            Catch ex As System.ArgumentException
                                'пытаемся найти ближайшую конструкцию
                                appAssembly = FuncAlignment.FuncFindBaseLineAssembly(ObjBaseLine, PKStat, ErrorLine, Dopusk)
                                If IsNothing(appAssembly) = True Then
                                    ReDim Preserve PointAssembl(7, CountPointAssembl)
                                    PointAssembl(0, CountPointAssembl) = PKStat
                                    PointAssembl(1, CountPointAssembl) = "Нет конструкции"
                                    CountPointAssembl += 1
                                    GoTo line1
                                End If
                            End Try
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            '=============================================================================================
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Dim PKAssemly As Double = 0 'пикет по умолчанию
                            Try
                                'забираем 1 код
                                Dim PointCalc1 As CalculatedPointCollection = appAssembly.GetPointsByCode(Kod1)
                                'вычисляем координаты первой точки
                                Dim StationPt1 As Double = 0
                                Dim OffsetPt1 As Double = 0
                                Dim ElevationPt1 As Double = 0
                                Dim PtAssembly1 As Point3d = Nothing
                                Dim PtAbsAssembly1 As Point3d = Nothing
                                Dim appSubColl As AppliedSubassemblyCollection = appAssembly.GetAppliedSubassemblies
                                If appSubColl.Count > 0 Then
                                    For i1 As Integer = 0 To appSubColl.Count - 1
                                        Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i1)
                                        Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                        Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                        If pointCol.Count > 0 Then
                                            For i2 As Integer = 0 To pointCol.Count - 1
                                                Dim ptAssembl As CalculatedPoint = pointCol.Item(i2)
                                                If ptAssembl.CorridorCodes.Count > 0 Then
                                                    For i3 As Integer = 0 To ptAssembl.CorridorCodes.Count - 1
                                                        Dim strCode As String = ptAssembl.CorridorCodes.Item(i3)
                                                        If strCode Like Kod1 Then
                                                            StationPt1 = ptAssembl.StationOffsetElevationToBaseline.X
                                                            OffsetPt1 = ptAssembl.StationOffsetElevationToBaseline.Y
                                                            ElevationPt1 = ptAssembl.StationOffsetElevationToBaseline.Z
                                                            PtAssembly1 = New Point3d(StationPt1, OffsetPt1, ElevationPt1)
                                                            PtAbsAssembly1 = blineSubAssembl.StationOffsetElevationToXYZ(PtAssembly1)
                                                            PKAssemly = StationPt1
                                                            GoTo LineFind3
                                                        End If
                                                    Next i3
                                                End If
                                            Next i2
                                        End If
                                    Next i1
                                End If
LineFind3:
                                '=================================================================================
                                'второй код
                                Dim StationPt2 As Double = 0
                                Dim OffsetPt2 As Double = 0
                                Dim ElevationPt2 As Double = 0
                                Dim PtAssembly2 As Point3d = Nothing
                                Dim PtAbsAssembly2 As Point3d = Nothing
                                'вычисляем координаты второй точки
                                Dim PointCalc2 As CalculatedPointCollection = appAssembly.GetPointsByCode(Kod2)
                                If appSubColl.Count > 0 Then
                                    For i1 As Integer = 0 To appSubColl.Count - 1
                                        Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i1)
                                        Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                        Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                        If pointCol.Count > 0 Then
                                            For i2 As Integer = 0 To pointCol.Count - 1
                                                Dim ptAssembl As CalculatedPoint = pointCol.Item(i2)
                                                If ptAssembl.CorridorCodes.Count > 0 Then
                                                    For i3 As Integer = 0 To ptAssembl.CorridorCodes.Count - 1
                                                        Dim strCode As String = ptAssembl.CorridorCodes.Item(i3)
                                                        If strCode Like Kod2 Then
                                                            StationPt2 = ptAssembl.StationOffsetElevationToBaseline.X
                                                            OffsetPt2 = ptAssembl.StationOffsetElevationToBaseline.Y
                                                            ElevationPt2 = ptAssembl.StationOffsetElevationToBaseline.Z
                                                            PtAssembly2 = New Point3d(StationPt2, OffsetPt2, ElevationPt2)
                                                            PtAbsAssembly2 = blineSubAssembl.StationOffsetElevationToXYZ(PtAssembly2)
                                                            GoTo LineFind4
                                                        End If
                                                    Next i3
                                                End If
                                            Next i2
                                        End If
                                    Next i1
                                End If
LineFind4:
                                If PKAssemly < StationPt2 Then
                                    PKAssemly = StationPt2
                                End If
                                If StationPt1 >= PKStart Then
                                    'заносим значения в массив
                                    ReDim Preserve PointAssembl(9, CountPointAssembl)
                                    PointAssembl(0, CountPointAssembl) = PKStat
                                    PointAssembl(1, CountPointAssembl) = PtAbsAssembly1.X
                                    PointAssembl(2, CountPointAssembl) = PtAbsAssembly1.Y
                                    PointAssembl(3, CountPointAssembl) = PtAbsAssembly1.Z
                                    PointAssembl(4, CountPointAssembl) = PtAbsAssembly2.X
                                    PointAssembl(5, CountPointAssembl) = PtAbsAssembly2.Y
                                    PointAssembl(6, CountPointAssembl) = PtAbsAssembly2.Z
                                    'расстояние между точками (горизонтальное расстояние)
                                    Dim Pt1 As Point2d = New Point2d(PtAbsAssembly1.X, PtAbsAssembly1.Y)
                                    Dim Pt2 As Point2d = New Point2d(PtAbsAssembly2.X, PtAbsAssembly2.Y)
                                    'горизонтальное расстояние
                                    Dim vect2d As Vector2d = Pt1.GetVectorTo(Pt2)
                                    PointAssembl(7, CountPointAssembl) = vect2d.Length
                                    'вертикальное расстояние
                                    Dim vect3d As Vector3d = PtAbsAssembly1.GetVectorTo(PtAbsAssembly2)
                                    PointAssembl(8, CountPointAssembl) = vect3d.Length
                                    'превышение
                                    PointAssembl(9, CountPointAssembl) = PtAbsAssembly1.Z - PtAbsAssembly2.Z
                                    CountPointAssembl += 1
                                Else
                                    ReDim Preserve PointAssembl(9, CountPointAssembl)
                                    PointAssembl(0, CountPointAssembl) = PKAssemly
                                    PointAssembl(1, CountPointAssembl) = 0
                                    PointAssembl(2, CountPointAssembl) = 0
                                    PointAssembl(3, CountPointAssembl) = 0
                                    PointAssembl(4, CountPointAssembl) = 0
                                    PointAssembl(5, CountPointAssembl) = 0
                                    PointAssembl(6, CountPointAssembl) = 0
                                    PointAssembl(7, CountPointAssembl) = 0
                                    PointAssembl(8, CountPointAssembl) = 0
                                    PointAssembl(9, CountPointAssembl) = 0
                                    CountPointAssembl += 1
                                End If
                                CountAssembl += 1
                            Catch ex As ArgumentOutOfRangeException
                            End Try
line1:
                        Next i
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'выбор всех кодов конструкций по секциям
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                ElseIf FormPlanStatement.RadioButton2.Checked = True And FormPlanStatement.ComboBox3.Text Like "Выбрать все" Then
                    ' Получаем доступ к группам секций
                    Dim SampleLineGroupCollection As ObjectIdCollection = ObjAlign.GetSampleLineGroupIds
                    Dim IDSectionCollection As ObjectIdCollection = New ObjectIdCollection
                    If SampleLineGroupCollection.Count > 0 Then
                        For Each sampleLineId As ObjectId In SampleLineGroupCollection
                            Dim sampleGroupLine As SampleLineGroup = tr.GetObject(sampleLineId, OpenMode.ForRead)
                            If sampleGroupLine.Name Like FormPlanStatement.ComboBox5.Text.Trim Then
                                IDSectionCollection = sampleGroupLine.GetSampleLineIds
                                Exit For
                            End If
                        Next
                        'начинаем перебор секций
                        Dim ArrayStationSampleLine As Double(,) = Nothing
                        Dim CountStationSampleLine As Integer = 0
                        If IDSectionCollection.Count > 0 Then
                            For Each idSample As ObjectId In IDSectionCollection
                                Dim SamplLine As SampleLine = tr.GetObject(idSample, OpenMode.ForRead)
                                Dim PK As Double = SamplLine.Station
                                If PK >= PKStart And PK <= PKEnd Then
                                    ReDim Preserve ArrayStationSampleLine(1, CountStationSampleLine)
                                    ArrayStationSampleLine(0, CountStationSampleLine) = SamplLine.Station
                                    ArrayStationSampleLine(1, CountStationSampleLine) = SamplLine.Number
                                    CountStationSampleLine += 1
                                End If
                            Next
                        End If
                        '=============================================================================================================
                        'запускаем Excel
                        Dim xlApp As Excel.Application = Nothing
                        Dim xlBook As Excel.Workbook = Nothing
                        Dim xlSheets As Excel.Worksheet = Nothing
                        xlApp = CreateObject("Excel.Application")
                        If IO.File.Exists(NameFileTemplate) = True Then
                            xlBook = xlApp.Workbooks.Open(NameFileTemplate)
                        Else
                            Dim avFiles As String = xlApp.GetOpenFilename("Excel files(*.xls*),*.xls*", 1, "Excel файлы", , True)
                            xlBook = xlApp.Workbooks.Open(avFiles)
                        End If
                        Dim file1 As String = xlApp.GetSaveAsFilename("Ведомость планировки_", "Excel Files (*.xlsx), .xlsx")
                        If file1 = "False" Then 'если ошибка то выходим из программы
                            xlBook.Close()
                            xlApp.Quit()
                            xlSheets = Nothing
                            xlBook = Nothing
                            xlApp = Nothing
                            Exit Sub
                        End If
                        'если шаблон не найден находим его самостоянельно
                        If IsNothing(xlBook) = True Then
                            'выбор файла
                            Dim NameFiles As String() = {""}
                            Dim fileDialog1 As Autodesk.AutoCAD.Windows.OpenFileDialog = New Autodesk.AutoCAD.Windows.OpenFileDialog("Открыть файл Excel", "", "xlsx", "xlsx", OpenFileDialogFlags.AllowMultiple)
                            Dim FileRez As String = fileDialog1.ShowDialog()
                            NameFiles = fileDialog1.GetFilenames
                            If IsNothing(NameFiles) = True Then GoTo LineEnd1
                            If File.Exists(NameFiles(0)) = True Then
                                Dim PutchFile As String = Path.GetDirectoryName(NameFiles(0))
                                Dim NameFileUser As String = Path.GetFileName(NameFiles(0))
                                Dim posT As Integer = NameFileUser.LastIndexOf(".")
                                Dim NewNameUser As String = Mid(NameFileUser, 1, posT - 1) & "_copy.xlsx"
                                Dim newCopyFile As String = PutchFile & "\" & NewNameUser
                                IO.File.Copy(NameFiles(0), newCopyFile, True)
                                xlBook = xlApp.Workbooks.Open(newCopyFile)
                            Else
                                GoTo LineEnd1
                            End If
                        End If
                        xlApp.Visible = True
                        'находим нужный лист(Ведомость объемов работ)
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim xlsheet As Excel.Worksheet
                        If FormPlanStatement.CheckBox2.Checked = True Then
                            xlsheet = xlBook.Worksheets.Item("Лист1") 'делаем ссылку на этот лист
                            xlsheet.Activate()
                            'ищем строку кодов (флаг первая ячейка с текстом ПК+)
                            'начальные значения для поиска кодов
                            Dim Col As Integer = 2
                            Dim Row As Integer = -1
                            For i As Integer = 1 To 100
                                Dim StrPK As String = xlsheet.Cells(i, 1).value
                                If StrPK Like "ПК+" Then
                                    Row = i + 1
                                    Exit For
                                End If
                            Next i
                            '==========================================================================================================
                            'если ячейка не найдена
                            If Row = -1 Then
                                Dim IntCol As Integer = xlApp.InputBox("Введите номер столбца для поиска ячейки ПК+", "Поиск ячейки")
                                Col = IntCol + 1
                                Dim IntRow As Integer = xlApp.InputBox("Введите номер строки для поиска ячейки ПК+", "Поиск ячейки")
                                Row = IntRow
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Dim ArrayCode As String() = FormPlanStatement.ComboBox3.DataSource
                            'начинаем доставать конструкции
                            Dim PKStat As Double = 0
                            Dim NumbStat As Integer = 0
                            Dim Dopusk As Double = FormPlanStatement.NumericUpDown1.Value 'допуск для поиска конструкций
                            If IsArray(ArrayStationSampleLine) = True And IsArray(ArrayCode) = True Then
                                For i As Integer = 0 To ArrayStationSampleLine.GetUpperBound(1)
                                    'записываем пикет
                                    PKStat = ArrayStationSampleLine(0, i)
                                    If FormPlanStatement.ComboBox6.Text Like "0.00" Then
                                        'усли есть рубленные пикеты тогда вычленяем +
                                        If sqwPiketCollection.Count > 0 Then
                                            Try
                                                Dim strPk As String = ObjAlign.GetStationStringWithEquations(PKStart)
                                                Dim valstr As Double = strPk.Replace("+", "")
                                                xlsheet.Cells(i + 4, 1) = valstr
                                            Catch ex As System.ArgumentException
                                                xlsheet.Cells(i + 4, 1) = Math.Round(PKStat, RoundZn)
                                            End Try
                                        Else
                                            xlsheet.Cells(i + 4, 1) = Math.Round(PKStat, RoundZn)
                                        End If
                                    Else
                                        Try
                                            Dim PiketNPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(ObjAlign, PKStat, RoundZn)
                                            xlsheet.Cells(i + 4, 1) = PiketNPKStrFormat
                                        Catch ex As System.ArgumentException
                                            xlsheet.Cells(i + 4, 1) = Math.Round(PKStat, RoundZn)
                                        End Try
                                    End If
                                    'NumbStat = ArrayStationSampleLine(1, i)
                                    '1. получаем доступ к конструкции
                                    Dim appAssembly As AppliedAssembly = Nothing
                                    Try
                                        appAssembly = ObjBaseLine.GetAppliedAssemblyAtStation(PKStat)
                                    Catch ex As ArgumentException
                                        'пытаемся найти ближайшую конструкцию
                                        appAssembly = FuncAlignment.FuncFindBaseLineAssembly(ObjBaseLine, PKStat, ErrorLine, Dopusk)
                                    End Try
                                    If IsNothing(appAssembly) = False Then
                                        '2. забираем все точки из из конструкции
                                        Dim CountCol As Integer = 0
                                        Dim appSubColl As AppliedSubassemblyCollection = appAssembly.GetAppliedSubassemblies 'получаем список подконструкций
                                        '============================================================================================================
                                        'перебираем ячейки таблицы
                                        Do
                                            '===========================================================================================
                                            'считываем пары кодов из таблицы Excel
                                            Dim KodStr1 As String = xlsheet.Cells(Row - 1, Col + CountCol).value 'ссылка на первую ячейку
                                            Dim KodStr2 As String = xlsheet.Cells(Row, Col + CountCol).value 'ссылка на вторую ячейку
                                            Dim Typedist As String = xlsheet.Cells(Row + 1, Col + CountCol).value 'ссылка на третью ячейку
                                            '==========================================================================================
                                            'если хотя бы один код в таблице не найден то выходим из цикла
                                            If KodStr1 Like "" And KodStr2 Like "" Then
                                                Exit Do
                                            End If
                                            '========================================================================================
                                            'начальные значения для 1 точки
                                            Dim StationPt1 As Double = 0
                                            Dim OffsetPt1 As Double = 0
                                            Dim ElevationPt1 As Double = 0
                                            Dim PtAssembly1 As Point3d = Nothing
                                            Dim PtAbsAssembly1 As Point3d = Nothing
                                            Dim boolKod1 As Boolean = False
                                            If appSubColl.Count > 0 Then 'если коллекция существует
                                                For i1 As Integer = 0 To appSubColl.Count - 1
                                                    Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i1) 'получем подконструкцию
                                                    Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo 'получаем базовую линию подконструкции
                                                    Dim pointCol As CalculatedPointCollection = SubAssembl.Points 'получаем коллекцию точек подконструкции
                                                    'если коллекция существует
                                                    If pointCol.Count > 0 Then
                                                        For i2 As Integer = 0 To pointCol.Count - 1
                                                            Dim ptAssembl As CalculatedPoint = pointCol.Item(i2)
                                                            If ptAssembl.CorridorCodes.Count > 0 Then
                                                                For i3 As Integer = 0 To ptAssembl.CorridorCodes.Count - 1
                                                                    Dim strCode As String = ptAssembl.CorridorCodes.Item(i3)
                                                                    If strCode Like KodStr1 Then
                                                                        StationPt1 = ptAssembl.StationOffsetElevationToBaseline.X
                                                                        OffsetPt1 = ptAssembl.StationOffsetElevationToBaseline.Y
                                                                        ElevationPt1 = ptAssembl.StationOffsetElevationToBaseline.Z
                                                                        PtAssembly1 = New Point3d(StationPt1, OffsetPt1, ElevationPt1)
                                                                        PtAbsAssembly1 = blineSubAssembl.StationOffsetElevationToXYZ(PtAssembly1)
                                                                        boolKod1 = True
                                                                        GoTo LineAssembl1
                                                                    End If
                                                                Next i3
                                                            End If
                                                        Next i2
                                                    End If
                                                Next i1
                                            End If
LineAssembl1:
                                            '==========================================================================================
                                            'вторая точка
                                            'начальные значения
                                            Dim StationPt2 As Double = 0
                                            Dim OffsetPt2 As Double = 0
                                            Dim ElevationPt2 As Double = 0
                                            Dim PtAssembly2 As Point3d = Nothing
                                            Dim PtAbsAssembly2 As Point3d = Nothing
                                            Dim boolKod2 As Boolean = False
                                            'вычисляем координаты второй точки
                                            If appSubColl.Count > 0 Then
                                                For i1 As Integer = 0 To appSubColl.Count - 1
                                                    Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i1)
                                                    Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                                    Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                                    'если коллекция существует
                                                    If pointCol.Count > 0 Then
                                                        For i2 As Integer = 0 To pointCol.Count - 1
                                                            Dim ptAssembl As CalculatedPoint = pointCol.Item(i2)
                                                            If ptAssembl.CorridorCodes.Count > 0 Then
                                                                For i3 As Integer = 0 To ptAssembl.CorridorCodes.Count - 1
                                                                    Dim strCode As String = ptAssembl.CorridorCodes.Item(i3)
                                                                    If strCode Like KodStr2 Then
                                                                        StationPt2 = ptAssembl.StationOffsetElevationToBaseline.X
                                                                        OffsetPt2 = ptAssembl.StationOffsetElevationToBaseline.Y
                                                                        ElevationPt2 = ptAssembl.StationOffsetElevationToBaseline.Z
                                                                        PtAssembly2 = New Point3d(StationPt2, OffsetPt2, ElevationPt2)
                                                                        PtAbsAssembly2 = blineSubAssembl.StationOffsetElevationToXYZ(PtAssembly2)
                                                                        boolKod2 = True
                                                                        GoTo LineAssembl2
                                                                    End If
                                                                Next i3
                                                            End If
                                                        Next i2
                                                    End If
                                                Next i1
                                            End If
LineAssembl2:
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'делаем анализ полученных данных
                                            If boolKod1 = True And boolKod2 = True Then
                                                Dim X1 As Double = PtAbsAssembly1.X
                                                Dim Y1 As Double = PtAbsAssembly1.Y
                                                Dim Z1 As Double = PtAbsAssembly1.Z
                                                Dim X2 As Double = PtAbsAssembly2.X
                                                Dim Y2 As Double = PtAbsAssembly2.Y
                                                Dim Z2 As Double = PtAbsAssembly2.Z
                                                If Typedist Like "г" Or Typedist Like "Г" Then
                                                    Dim lg3d As Vector3d = PtAbsAssembly1.GetVectorTo(PtAbsAssembly2)
                                                    Dim v2d As Vector2d = lg3d.Convert2d(New Plane())
                                                    Dim Pt1 As Point2d = New Point2d(Y1, Z1)
                                                    Dim Pt2 As Point2d = New Point2d(Y2, Z2)
                                                    Dim vect1 As Vector2d = Pt1.GetVectorTo(Pt2)
                                                    Dim a As Double = v2d.Length
                                                    xlsheet.Cells(Row + 1 + i, Col + CountCol).value = Math.Round(a, RoundZn)
                                                ElseIf Typedist Like "в" Or Typedist Like "В" Then
                                                    Dim lg3d As Vector3d = PtAbsAssembly1.GetVectorTo(PtAbsAssembly2)
                                                    xlsheet.Cells(Row + 1 + i, Col + CountCol).value = Math.Round(lg3d.Length, RoundZn)
                                                ElseIf Typedist Like "п" Or Typedist Like "П" Then
                                                    Dim a As Double = Math.Abs(Z1 - Z2)
                                                    xlsheet.Cells(Row + 1 + i, Col + CountCol).value = Math.Round(a, RoundZn)
                                                Else
                                                    If FormPlanStatement.RadioButton5.Checked = True Then
                                                        'нужны превышения
                                                        Dim a As Double = Math.Abs(Z1 - Z2)
                                                        xlsheet.Cells(Row + 1 + i, Col + CountCol).value = Math.Round(a, RoundZn)
                                                    ElseIf FormPlanStatement.RadioButton4.Checked = True Then
                                                        Dim lg3d As Vector3d = PtAbsAssembly1.GetVectorTo(PtAbsAssembly2)
                                                        xlsheet.Cells(Row + 1 + i, Col + CountCol).value = Math.Round(lg3d.Length, RoundZn)
                                                    Else
                                                        Dim lg3d As Vector3d = PtAbsAssembly1.GetVectorTo(PtAbsAssembly2)
                                                        Dim v2d As Vector2d = lg3d.Convert2d(New Plane())
                                                        Dim Pt1 As Point2d = New Point2d(Y1, Z1)
                                                        Dim Pt2 As Point2d = New Point2d(Y2, Z2)
                                                        Dim vect1 As Vector2d = Pt1.GetVectorTo(Pt2)
                                                        Dim a As Double = v2d.Length
                                                        xlsheet.Cells(Row + 1 + i, Col + CountCol).value = Math.Round(a, RoundZn)
                                                    End If
                                                End If
                                            End If
                                            CountCol += 1
                                        Loop
                                    End If
                                Next i
                            End If
                            xlSheets = Nothing
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'отчет по материалам
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        If FormPlanStatement.CheckBox3.Checked = True Then
                            xlsheet = xlBook.Worksheets.Add()
                            xlsheet.Name = "Материалы"
                            xlsheet.Activate()

                            'формируем шапку таблицы
                            Dim cell1 As Excel.Range = Nothing 'первая ячейка в строке
                            Dim cell2 As Excel.Range = Nothing 'последняя ячейка в строке
                            Dim CellRng As Excel.Range = Nothing 'массив ячеек
                            'ПК
                            cell1 = xlsheet.Cells(2, 1)
                            cell2 = xlsheet.Cells(3, 1)
                            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                            CellRng.Merge(Type.Missing)
                            cell1.Value = "'ПК+"
                            'расстояние
                            cell1 = xlsheet.Cells(2, 2)
                            cell2 = xlsheet.Cells(3, 2)
                            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                            CellRng.Merge(Type.Missing)
                            cell1.Value = "'Расстояние, м"
                            'начинаем доставать конструкции
                            Dim PKStat0 As Double = 0
                            Dim NumbStat As Integer = 0
                            Dim Dopusk As Double = FormPlanStatement.NumericUpDown1.Value 'допуск для поиска конструкций
                            If IsArray(ArrayStationSampleLine) = True Then
                                Dim RowInd As Integer = 4
                                Dim countMaterial As Integer = 0
                                For i As Integer = 0 To ArrayStationSampleLine.GetUpperBound(1)
                                    'записываем пикет
                                    '============================================================================================
                                    Dim PKStat As Double = ArrayStationSampleLine(0, i) 'пикет конструкции
                                    If FormPlanStatement.ComboBox6.Text Like "0.00" Then
                                        'усли есть рубленные пикеты тогда вычленяем +
                                        If sqwPiketCollection.Count > 0 Then
                                            Try
                                                Dim strPk As String = ObjAlign.GetStationStringWithEquations(PKStart)
                                                Dim valstr As Double = strPk.Replace("+", "")
                                                xlsheet.Cells(i + RowInd, 1) = valstr
                                            Catch ex As System.ArgumentException
                                                xlsheet.Cells(i + RowInd, 1) = Math.Round(PKStat, RoundZn)
                                            End Try
                                        Else
                                            xlsheet.Cells(i + RowInd, 1) = Math.Round(PKStat, RoundZn)
                                        End If
                                    Else
                                        Try
                                            Dim PiketNPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(ObjAlign, PKStat, RoundZn)
                                            xlsheet.Cells(i + RowInd, 1) = PiketNPKStrFormat
                                        Catch ex As System.ArgumentException
                                            xlsheet.Cells(i + RowInd, 1) = Math.Round(PKStat, RoundZn)
                                        End Try
                                    End If
                                    '==============================================================================================
                                    'записываем расстояние
                                    If i = 0 Then
                                        xlsheet.Cells(i + RowInd, 2).value = 0
                                    Else
                                        Dim Dist As Double = Math.Abs(PKStat - PKStat0)
                                        PKStat0 = PKStat
                                        xlsheet.Cells(i + RowInd, 2).value = Math.Round(Dist, RoundZn)
                                    End If
                                    '==============================================================================================
                                    '==============================================================================================
                                    'начинаем разбор конструкции
                                    '==============================================================================================
                                    '==============================================================================================
                                    '1. получаем доступ к конструкции
                                    Dim appAssembly As AppliedAssembly = Nothing
                                    Try
                                        appAssembly = ObjBaseLine.GetAppliedAssemblyAtStation(PKStat)
                                    Catch ex As ArgumentException
                                        'пытаемся найти ближайшую конструкцию
                                        appAssembly = FuncAlignment.FuncFindBaseLineAssembly(ObjBaseLine, PKStat, ErrorLine, Dopusk)
                                    End Try
                                    '===============================================================================================
                                    'конструкция найдена продолжаем
                                    If IsNothing(appAssembly) = False Then
                                        'получаем из нее коллекцию материалов
                                        Dim ShColl As CalculatedShapeCollection = appAssembly.Shapes
                                        'если марериалы существуют тогда достаем их
                                        If ShColl.Count > 0 Then
                                            For j As Integer = 0 To ShColl.Count - 1
                                                Dim Sh As CalculatedShape = ShColl.Item(j) 'читаем материал
                                                Dim NameSh As String = Sh.CorridorCodes.Item(0) 'имя материала
                                                Dim AreaStr As Double = Sh.Area 'площадь
                                                '============================================================================================================
                                                'записываем в таблицу
                                                'перебираем ячейки таблицы
                                                Dim posCol As Integer = 3 'читаем столбцы с наименованием матерала с 3 позиции
                                                Do
                                                    Dim KodShape As String = xlsheet.Cells(3, posCol).value 'ссылка на ячейку с материалом
                                                    If KodShape Like NameSh Then 'если такой материал уже существует тогда просто записываем значение с соответсвующую строку
                                                        Dim val As Double = CDbl(xlsheet.Cells(i + RowInd, posCol).value)
                                                        xlsheet.Cells(i + RowInd, posCol).value = val + AreaStr
                                                        Exit Do
                                                    ElseIf KodShape Like "" Or IsNothing(KodShape) = True Then 'если материал новый тогда проименовываем следующий столбец
                                                        xlsheet.Cells(3, posCol).value = NameSh 'ссылка на ячейку с материалом
                                                        xlsheet.Cells(i + RowInd, posCol).value = AreaStr
                                                        countMaterial += 1
                                                        Exit Do
                                                    Else
                                                        posCol += 1
                                                    End If
                                                Loop
                                            Next j
                                        End If
                                    End If
                                Next i
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'форматируем таблицу
                                If countMaterial > 0 Then
                                    'Площадь
                                    cell1 = xlsheet.Cells(2, 3)
                                    cell2 = xlsheet.Cells(2, 2 + countMaterial)
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.Merge(Type.Missing)
                                    cell1.Value = "'Площадь, кв.м"
                                    'вычисляем объемы
                                    cell1 = xlsheet.Cells(2, 3 + countMaterial)
                                    cell2 = xlsheet.Cells(2, 2 + countMaterial * 2)
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.Merge(Type.Missing)
                                    cell1.Value = "'Объем, куб.м"
                                    For i As Integer = 0 To countMaterial - 1
                                        xlsheet.Cells(3, 3 + i + countMaterial).value = xlsheet.Cells(3, 3 + i).value 'ссылка на ячейку с материалом
                                        'вычисляем объем
                                        Dim StartPosRow As Integer = 4
                                        Dim startPosCol As Integer = 3 + countMaterial + i
                                        Do
                                            Dim Dist As String = xlsheet.Cells(StartPosRow, 1).value
                                            If IsNothing(Dist) = True Or Dist = "" Then
                                                Exit Do
                                            End If
                                            Dim posDist As Integer = 0 - (countMaterial + i + 1)
                                            Dim posSqr As Integer = 0 - countMaterial
                                            xlsheet.Cells(StartPosRow, startPosCol).FormulaR1C1 = "=RC[" & posDist & "]*RC[" & posSqr & "]"
                                            StartPosRow += 1
                                        Loop
                                    Next
                                End If
                            End If
                            xlSheets = Nothing
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'отчет по связям
                        If FormPlanStatement.CheckBox4.Checked = True Then
                            xlsheet = xlBook.Worksheets.Add()
                            xlsheet.Name = "Связи"
                            xlsheet.Activate()
                            'формируем шапку таблицы
                            Dim cell1 As Excel.Range = Nothing 'первая ячейка в строке
                            Dim cell2 As Excel.Range = Nothing 'последняя ячейка в строке
                            Dim CellRng As Excel.Range = Nothing 'массив ячеек
                            'ПК
                            cell1 = xlsheet.Cells(2, 1)
                            cell2 = xlsheet.Cells(3, 1)
                            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                            CellRng.Merge(Type.Missing)
                            cell1.Value = "'ПК+"
                            'расстояние
                            cell1 = xlsheet.Cells(2, 2)
                            cell2 = xlsheet.Cells(3, 2)
                            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                            CellRng.Merge(Type.Missing)
                            cell1.Value = "'Расстояние, м"
                            'начинаем доставать конструкции
                            Dim PKStat0 As Double = 0
                            Dim NumbStat As Integer = 0
                            Dim Dopusk As Double = FormPlanStatement.NumericUpDown1.Value 'допуск для поиска конструкций
                            If IsArray(ArrayStationSampleLine) = True Then
                                Dim RowInd As Integer = 4
                                Dim ColInd As Integer = 3
                                Dim countMaterial As Integer = 0
                                Dim ArrayCodeExcel As String(,) = Nothing
                                Dim countArrayCodeExcel As Integer = 0
                                For i As Integer = 0 To ArrayStationSampleLine.GetUpperBound(1)
                                    'записываем пикет
                                    '============================================================================================
                                    Dim PKStat As Double = ArrayStationSampleLine(0, i) 'пикет конструкции
                                    '============================================================================================
                                    'записываем его в таблицу
                                    If FormPlanStatement.ComboBox6.Text Like "0.00" Then
                                        'усли есть рубленные пикеты тогда вычленяем +
                                        If sqwPiketCollection.Count > 0 Then
                                            Try
                                                Dim strPk As String = ObjAlign.GetStationStringWithEquations(PKStart)
                                                Dim valstr As Double = strPk.Replace("+", "")
                                                xlsheet.Cells(i + RowInd, 1) = valstr
                                            Catch ex As System.ArgumentException
                                                xlsheet.Cells(i + RowInd, 1) = Math.Round(PKStat, RoundZn)
                                            End Try
                                        Else
                                            xlsheet.Cells(i + RowInd, 1) = Math.Round(PKStat, RoundZn)
                                        End If
                                    Else
                                        Try
                                            Dim PiketNPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(ObjAlign, PKStat, RoundZn)
                                            xlsheet.Cells(i + RowInd, 1) = PiketNPKStrFormat
                                        Catch ex As System.ArgumentException
                                            xlsheet.Cells(i + RowInd, 1) = Math.Round(PKStat, RoundZn)
                                        End Try
                                    End If
                                    '==============================================================================================
                                    'записываем расстояние
                                    If i = 0 Then
                                        xlsheet.Cells(i + RowInd, 2).value = 0
                                    Else
                                        Dim Dist As Double = Math.Abs(PKStat - PKStat0)
                                        PKStat0 = PKStat
                                        xlsheet.Cells(i + RowInd, 2).value = Math.Round(Dist, RoundZn)
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'разбор конструкции
                                    '1. получаем доступ к конструкции
                                    Dim appAssembly As AppliedAssembly = Nothing
                                    Try
                                        appAssembly = ObjBaseLine.GetAppliedAssemblyAtStation(PKStat)
                                    Catch ex As ArgumentException
                                        'пытаемся найти ближайшую конструкцию
                                        appAssembly = FuncAlignment.FuncFindBaseLineAssembly(ObjBaseLine, PKStat, ErrorLine, Dopusk)
                                    End Try
                                    '===============================================================================================
                                    'конструкция найдена продолжаем
                                    If IsNothing(appAssembly) = False Then
                                        'получаем из нее коллекцию материалов
                                        Dim LinkColl As CalculatedLinkCollection = appAssembly.Links
                                        If LinkColl.Count > 0 Then
                                            For Each LinkObj As CalculatedLink In LinkColl
                                                'получаем имя связи
                                                If LinkObj.CorridorCodes.Count > 0 Then
                                                    Dim NameLink As String = LinkObj.CorridorCodes.Item(0)
                                                    cell1 = xlsheet.Cells(2, ColInd)
                                                    cell2 = xlsheet.Cells(2, ColInd + 8)
                                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Merge(Type.Missing)
                                                    cell1.Value = NameLink
                                                End If
                                                'получаем список точек
                                                Dim appSubAssembl As AppliedSubassembly = LinkObj.SubassemblyBelongedTo
                                                Dim blineSubAssembl As BaseBaseline = appSubAssembl.BaselineHookedTo
                                                Dim ColPoint As CalculatedPointCollection = LinkObj.CalculatedPoints
                                                'если в коллекции 2 точки (рассматриваются только первые две точки
                                                If ColPoint.Count > 1 Then
                                                    Dim pt1 As CalculatedPoint = ColPoint.Item(0)
                                                    Dim pt2 As CalculatedPoint = ColPoint.Item(1)
                                                    If pt1.CorridorCodes.Count > 0 And pt2.CorridorCodes.Count > 0 Then
                                                        If countArrayCodeExcel = 0 Then
                                                            ReDim Preserve ArrayCodeExcel(1, countArrayCodeExcel)
                                                            ArrayCodeExcel(0, countArrayCodeExcel) = pt1.CorridorCodes.Item(0)
                                                            ArrayCodeExcel(1, countArrayCodeExcel) = pt2.CorridorCodes.Item(0)
                                                            countArrayCodeExcel += 1
                                                        Else
                                                            'ищем совпадения в массиве
                                                            Dim flag As Boolean = False
                                                            For i1 As Integer = 0 To ArrayCodeExcel.GetUpperBound(1)
                                                                If ArrayCodeExcel(0, i1) Like pt1.CorridorCodes.Item(0) And ArrayCodeExcel(1, i1) Like pt2.CorridorCodes.Item(0) Then
                                                                    ColInd = i1 * 9 + 3
                                                                    flag = True
                                                                    Exit For
                                                                End If
                                                            Next
                                                            'ветка если код не найден
                                                            If flag = False Then
                                                                ReDim Preserve ArrayCodeExcel(1, countArrayCodeExcel)
                                                                ArrayCodeExcel(0, countArrayCodeExcel) = pt1.CorridorCodes.Item(0)
                                                                ArrayCodeExcel(1, countArrayCodeExcel) = pt2.CorridorCodes.Item(0)
                                                                ColInd = countArrayCodeExcel * 9 + 3
                                                                countArrayCodeExcel += 1
                                                            End If
                                                        End If
                                                        Dim Pt3d1 As Point3d = pt1.StationOffsetElevationToSubassembly
                                                        Dim Pt3d2 As Point3d = pt2.StationOffsetElevationToSubassembly
                                                        'абсолютные координаты 1 точки
                                                        Dim PtAssembly1 As Point3d = New Point3d(Pt3d1.X, Pt3d1.Y, Pt3d1.Z)
                                                        Dim PtAbsAssembly1 As Point3d = blineSubAssembl.StationOffsetElevationToXYZ(PtAssembly1)
                                                        'абсолютные координаты 2 точки
                                                        Dim PtAssembly2 As Point3d = New Point3d(Pt3d2.X, Pt3d2.Y, Pt3d2.Z)
                                                        Dim PtAbsAssembly2 As Point3d = blineSubAssembl.StationOffsetElevationToXYZ(PtAssembly2)
                                                        'записываем в таблицу
                                                        xlsheet.Cells(3, ColInd).value = "2D, м"
                                                        xlsheet.Cells(3, ColInd + 1).value = "3D, м"
                                                        xlsheet.Cells(3, ColInd + 2).value = "i"
                                                        xlsheet.Cells(3, ColInd + 3).value = "X" & pt1.CorridorCodes.Item(0) & " ,м"
                                                        xlsheet.Cells(3, ColInd + 4).value = "Y" & pt1.CorridorCodes.Item(0) & " ,м"
                                                        xlsheet.Cells(3, ColInd + 5).value = "Z" & pt1.CorridorCodes.Item(0) & " ,м"
                                                        xlsheet.Cells(3, ColInd + 6).value = "X" & pt2.CorridorCodes.Item(0) & " ,м"
                                                        xlsheet.Cells(3, ColInd + 7).value = "Y" & pt2.CorridorCodes.Item(0) & " ,м"
                                                        xlsheet.Cells(3, ColInd + 8).value = "Z" & pt2.CorridorCodes.Item(0) & " ,м"
                                                        Dim Pt1v As Point2d = New Point2d(PtAbsAssembly1.Y, PtAbsAssembly1.Z)
                                                        Dim Pt2v As Point2d = New Point2d(PtAbsAssembly2.Y, PtAbsAssembly2.Z)
                                                        Dim vect1 As Vector2d = Pt1v.GetVectorTo(Pt2v)
                                                        Dim a As Double = vect1.Length
                                                        xlsheet.Cells(RowInd + i, ColInd).value = Math.Round(a, RoundZn)
                                                        Dim lg3d As Vector3d = PtAbsAssembly1.GetVectorTo(PtAbsAssembly2)
                                                        xlsheet.Cells(RowInd + i, ColInd + 1).value = Math.Round(lg3d.Length, RoundZn)
                                                        xlsheet.Cells(RowInd + i, ColInd + 2).value = "i"
                                                        xlsheet.Cells(RowInd + i, ColInd + 3).value = PtAbsAssembly1.X
                                                        xlsheet.Cells(RowInd + i, ColInd + 4).value = PtAbsAssembly1.Y
                                                        xlsheet.Cells(RowInd + i, ColInd + 5).value = PtAbsAssembly1.Z
                                                        xlsheet.Cells(RowInd + i, ColInd + 6).value = PtAbsAssembly2.X
                                                        xlsheet.Cells(RowInd + i, ColInd + 7).value = PtAbsAssembly2.Y
                                                        xlsheet.Cells(RowInd + i, ColInd + 8).value = PtAbsAssembly2.Z
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        locker.Dispose()
                        xlBook = Nothing
                        xlApp.Quit()
                        xlApp = Nothing
                    End If
LineEnd1:
                    Exit Sub
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'выбор всех кодов ПО КОНСТРУКЦИЯМ
                ElseIf FormPlanStatement.RadioButton1.Checked = True And FormPlanStatement.ComboBox3.Text Like "Выбрать все" Then
                    '============================================================================================================
                    'запускаем эксель
                    Dim xlApp As Excel.Application = Nothing
                    Dim xlBook As Excel.Workbook = Nothing
                    Dim xlSheets As Excel.Worksheet = Nothing
                    xlApp = CreateObject("Excel.Application")
                    If IO.File.Exists(NameFileTemplate) = True Then
                        xlBook = xlApp.Workbooks.Open(NameFileTemplate)
                    Else
                        Dim avFiles As String = xlApp.GetOpenFilename("Excel files(*.xls*),*.xls*", 1, "Excel файлы", , True)
                        xlBook = xlApp.Workbooks.Open(avFiles)
                    End If
                    Dim file1 As String = xlApp.GetSaveAsFilename("Ведомость планировки_", "Excel Files (*.xlsx), .xlsx")
                    If file1 = "False" Then 'если ошибка то выходим из программы
                        xlBook.Close()
                        xlApp.Quit()
                        xlSheets = Nothing
                        xlBook = Nothing
                        xlApp = Nothing
                        Exit Sub
                    End If
                    'если шаблон не найден находим его самостоянельно
                    If IsNothing(xlBook) = True Then
                        Dim NameFiles As String() = {""}
                        Dim fileDialog1 As Autodesk.AutoCAD.Windows.OpenFileDialog = New Autodesk.AutoCAD.Windows.OpenFileDialog("Открыть файл Excel", "", "xlsx", "xlsx", OpenFileDialogFlags.AllowMultiple)
                        Dim FileRez As String = fileDialog1.ShowDialog()
                        NameFiles = fileDialog1.GetFilenames
                        If IsNothing(NameFiles) = True Then GoTo LineEnd2
                        If File.Exists(NameFiles(0)) = True Then
                            xlBook = xlApp.Workbooks.Open(NameFiles(0))
                        Else
                            GoTo LineEnd1
                        End If
                    End If
                    xlApp.Visible = True
                    'находим нужный лист(Ведомость объемов работ)
                    Dim xlsheet As Excel.Worksheet
                    If FormPlanStatement.CheckBox2.Checked = True Then
                        xlsheet = xlBook.Worksheets.Item("Лист1") 'делаем ссылку на этот лист
                        xlsheet.Activate()
                        'ищем строку кодов (флаг первая ячейка с текстом ПК+)
                        'начальные значения для поиска кодов
                        Dim Col As Integer = 2
                        Dim Row As Integer = -1
                        For i As Integer = 1 To 100
                            Dim StrPK As String = xlsheet.Cells(i, 1).value
                            If StrPK Like "ПК+" Then
                                Row = i + 1
                                Exit For
                            End If
                        Next i
                        '==========================================================================================================
                        'если ячейка не найдена
                        If Row = -1 Then
                            Dim IntCol As Integer = xlApp.InputBox("Введите номер столбца для поиска ячейки ПК+", "Поиск ячейки")
                            Col = IntCol + 1
                            Dim IntRow As Integer = xlApp.InputBox("Введите номер строки для поиска ячейки ПК+", "Поиск ячейки")
                            Row = IntRow
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim ArrayCode As String() = FormPlanStatement.ComboBox3.DataSource
                        'начинаем доставать конструкции
                        Dim PKStat As Double = 0
                        Dim NumbStat As Integer = 0
                        Dim Dopusk As Double = FormPlanStatement.NumericUpDown1.Value 'допуск для поиска конструкций
                        'получаем доступ к конструкции
                        Dim AssemlyCount As Integer = 1 'порядковый номер конструкции
                        '==============================================================================================================
                        'получаем коллекцию регионов входящих в базовую линию
                        Dim RegionColl As BaselineRegionCollection = ObjBaseLine.BaselineRegions
                        For Each BaseMyRegion As BaselineRegion In RegionColl
                            'если регион находится в диапазоне пикетов пользователя продолжаем
                            If BaseMyRegion.StartStation >= PKStart And BaseMyRegion.EndStation <= PKEnd Then
                                'если регион рабочий
                                If BaseMyRegion.NeedsProcessing = True Then
                                    Dim AppAssemblColl As AppliedAssemblyCollection = BaseMyRegion.AppliedAssemblies
                                    If AppAssemblColl.Count > 0 Then
                                        For Each appAssembly As AppliedAssembly In AppAssemblColl
                                            Dim PTPointCol As CalculatedPointCollection = appAssembly.Points
                                            PKStart = PTPointCol.Item(0).StationOffsetElevationToBaseline.X
                                            If FormPlanStatement.ComboBox6.Text Like "0.00" Then
                                                'усли есть рубленные пикеты тогда вычленяем +
                                                If sqwPiketCollection.Count > 0 Then
                                                    Try
                                                        Dim strPk As String = ObjAlign.GetStationStringWithEquations(PKStart)
                                                        Dim valstr As Double = strPk.Replace("+", "")
                                                        xlsheet.Cells(AssemlyCount + 4, 1) = valstr
                                                    Catch ex As System.ArgumentException
                                                        xlsheet.Cells(AssemlyCount + 4, 1) = Math.Round(PKStart, RoundZn)
                                                    End Try
                                                Else
                                                    xlsheet.Cells(AssemlyCount + 4, 1) = Math.Round(PKStart, RoundZn)
                                                End If
                                            Else
                                                Try
                                                    Dim PiketNPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(ObjAlign, PKStart, RoundZn)
                                                    xlsheet.Cells(AssemlyCount + 4, 1) = PiketNPKStrFormat
                                                Catch ex As System.ArgumentException
                                                    xlsheet.Cells(AssemlyCount + 4, 1) = Math.Round(PKStart, RoundZn)
                                                End Try
                                            End If
                                            'порядковый номер столбца отчета
                                            Dim CountCol As Integer = 0
                                            Do
                                                Dim KodStr1 As String = xlsheet.Cells(Row - 1, Col + CountCol).value 'ссылка на первую ячейку
                                                Dim KodStr2 As String = xlsheet.Cells(Row, Col + CountCol).value 'ссылка на вторую ячейку
                                                Dim Typedist As String = xlsheet.Cells(Row + 1, Col + CountCol).value 'ссылка на третью ячейку
                                                If KodStr1 Like "" And KodStr2 Like "" Then
                                                    Exit Do
                                                End If
                                                'вычисляем координаты первой точки
                                                Dim StationPt1 As Double = 0
                                                Dim OffsetPt1 As Double = 0
                                                Dim ElevationPt1 As Double = 0
                                                Dim PtAssembly1 As Point3d = Nothing
                                                Dim PtAbsAssembly1 As Point3d = Nothing
                                                Dim boolKod1 As Boolean = False
                                                Dim appSubColl As AppliedSubassemblyCollection = appAssembly.GetAppliedSubassemblies
                                                If appSubColl.Count > 0 Then
                                                    For i1 As Integer = 0 To appSubColl.Count - 1
                                                        Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i1)
                                                        Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                                        Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                                        If pointCol.Count > 0 Then
                                                            For i2 As Integer = 0 To pointCol.Count - 1
                                                                Dim ptAssembl As CalculatedPoint = pointCol.Item(i2)
                                                                If ptAssembl.CorridorCodes.Count > 0 Then
                                                                    For i3 As Integer = 0 To ptAssembl.CorridorCodes.Count - 1
                                                                        Dim strCode As String = ptAssembl.CorridorCodes.Item(i3)
                                                                        If strCode Like KodStr1 Then
                                                                            StationPt1 = ptAssembl.StationOffsetElevationToBaseline.X
                                                                            OffsetPt1 = ptAssembl.StationOffsetElevationToBaseline.Y
                                                                            ElevationPt1 = ptAssembl.StationOffsetElevationToBaseline.Z
                                                                            PtAssembly1 = New Point3d(StationPt1, OffsetPt1, ElevationPt1)
                                                                            PtAbsAssembly1 = blineSubAssembl.StationOffsetElevationToXYZ(PtAssembly1)
                                                                            boolKod1 = True
                                                                            GoTo lineAssembl3
                                                                        End If
                                                                    Next i3
                                                                End If
                                                            Next i2
                                                        End If

                                                    Next i1
                                                End If
lineAssembl3:
                                                '==========================================================================================
                                                'вторая точка
                                                Dim StationPt2 As Double = 0
                                                Dim OffsetPt2 As Double = 0
                                                Dim ElevationPt2 As Double = 0
                                                Dim PtAssembly2 As Point3d = Nothing
                                                Dim PtAbsAssembly2 As Point3d = Nothing
                                                Dim boolKod2 As Boolean = False
                                                'вычисляем координаты второй точки
                                                If appSubColl.Count > 0 Then
                                                    For i1 As Integer = 0 To appSubColl.Count - 1
                                                        Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i1)
                                                        Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                                        Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                                        If pointCol.Count > 0 Then
                                                            For i2 As Integer = 0 To pointCol.Count - 1
                                                                Dim ptAssembl As CalculatedPoint = pointCol.Item(i2)
                                                                If ptAssembl.CorridorCodes.Count > 0 Then
                                                                    For i3 As Integer = 0 To ptAssembl.CorridorCodes.Count - 1
                                                                        Dim strCode As String = ptAssembl.CorridorCodes.Item(i3)
                                                                        If strCode Like KodStr2 Then
                                                                            StationPt2 = ptAssembl.StationOffsetElevationToBaseline.X
                                                                            OffsetPt2 = ptAssembl.StationOffsetElevationToBaseline.Y
                                                                            ElevationPt2 = ptAssembl.StationOffsetElevationToBaseline.Z
                                                                            PtAssembly2 = New Point3d(StationPt2, OffsetPt2, ElevationPt2)
                                                                            PtAbsAssembly2 = blineSubAssembl.StationOffsetElevationToXYZ(PtAssembly2)
                                                                            boolKod2 = True
                                                                            GoTo lineAssembl4
                                                                        End If
                                                                    Next i3
                                                                End If
                                                            Next i2
                                                        End If
                                                    Next i1
                                                End If
lineAssembl4:
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                'делаем анализ полученных данных
                                                If boolKod1 = True And boolKod2 = True Then
                                                    Dim X1 As Double = PtAbsAssembly1.X
                                                    Dim Y1 As Double = PtAbsAssembly1.Y
                                                    Dim Z1 As Double = PtAbsAssembly1.Z
                                                    Dim X2 As Double = PtAbsAssembly2.X
                                                    Dim Y2 As Double = PtAbsAssembly2.Y
                                                    Dim Z2 As Double = PtAbsAssembly2.Z
                                                    If Typedist Like "г" Or Typedist Like "Г" Then
                                                        Dim lg3d As Vector3d = PtAbsAssembly1.GetVectorTo(PtAbsAssembly2)
                                                        Dim v2d As Vector2d = lg3d.Convert2d(New Plane())
                                                        Dim Pt1 As Point2d = New Point2d(Y1, Z1)
                                                        Dim Pt2 As Point2d = New Point2d(Y2, Z2)
                                                        Dim vect1 As Vector2d = Pt1.GetVectorTo(Pt2)
                                                        Dim a As Double = v2d.Length
                                                        xlsheet.Cells(Row + 1 + AssemlyCount, Col + CountCol).value = Math.Round(a, RoundZn)
                                                    ElseIf Typedist Like "в" Or Typedist Like "В" Then
                                                        Dim lg3d As Vector3d = PtAbsAssembly1.GetVectorTo(PtAbsAssembly2)
                                                        xlsheet.Cells(Row + 1 + AssemlyCount, Col + CountCol).value = Math.Round(lg3d.Length, RoundZn)
                                                    ElseIf Typedist Like "п" Or Typedist Like "П" Then
                                                        Dim a As Double = Math.Abs(Z1 - Z2)
                                                        xlsheet.Cells(Row + 1 + AssemlyCount, Col + CountCol).value = Math.Round(a, RoundZn)
                                                    Else
                                                        If FormPlanStatement.RadioButton5.Checked = True Then
                                                            'нужны превышения
                                                            Dim a As Double = Math.Abs(Z1 - Z2)
                                                            xlsheet.Cells(Row + 1 + AssemlyCount, Col + CountCol).value = Math.Round(a, RoundZn)
                                                        ElseIf FormPlanStatement.RadioButton4.Checked = True Then
                                                            Dim lg3d As Vector3d = PtAbsAssembly1.GetVectorTo(PtAbsAssembly2)
                                                            xlsheet.Cells(Row + 1 + AssemlyCount, Col + CountCol).value = Math.Round(lg3d.Length, RoundZn)
                                                        Else
                                                            Dim lg3d As Vector3d = PtAbsAssembly1.GetVectorTo(PtAbsAssembly2)
                                                            Dim v2d As Vector2d = lg3d.Convert2d(New Plane())
                                                            Dim Pt1 As Point2d = New Point2d(Y1, Z1)
                                                            Dim Pt2 As Point2d = New Point2d(Y2, Z2)
                                                            Dim vect1 As Vector2d = Pt1.GetVectorTo(Pt2)
                                                            Dim a As Double = v2d.Length
                                                            xlsheet.Cells(Row + 1 + AssemlyCount, Col + CountCol).value = Math.Round(a, RoundZn)
                                                        End If
                                                    End If
                                                End If
                                                CountCol += 1
                                            Loop
                                            AssemlyCount += 1
                                        Next
                                    End If
                                End If
                            End If
                        Next
                        xlSheets = Nothing
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'отчет по материалам
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    If FormPlanStatement.CheckBox3.Checked = True Then
                        xlsheet = xlBook.Worksheets.Add()
                        xlsheet.Name = "Материалы"
                        xlsheet.Activate()
                        'формируем шапку таблицы
                        Dim cell1 As Excel.Range = Nothing 'первая ячейка в строке
                        Dim cell2 As Excel.Range = Nothing 'последняя ячейка в строке
                        Dim CellRng As Excel.Range = Nothing 'массив ячеек
                        'ПК
                        cell1 = xlsheet.Cells(2, 1)
                        cell2 = xlsheet.Cells(3, 1)
                        CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                        CellRng.Merge(Type.Missing)
                        cell1.Value = "'ПК+"
                        'расстояние
                        cell1 = xlsheet.Cells(2, 2)
                        cell2 = xlsheet.Cells(3, 2)
                        CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                        CellRng.Merge(Type.Missing)
                        cell1.Value = "'Расстояние, м"
                        'начинаем доставать конструкции
                        Dim PKStat0 As Double = 0
                        Dim NumbStat As Integer = 0
                        Dim Dopusk As Double = FormPlanStatement.NumericUpDown1.Value 'допуск для поиска конструкций
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'достаем активные регионы по базовой линии
                        Dim reg As BaselineRegionCollection = ObjBaseLine.BaselineRegions
                        Dim RowInd As Integer = 4
                        Dim countMaterial As Integer = 0
                        For Each BlineRegion As BaselineRegion In reg
                            Dim PKStartReg As Double = BlineRegion.StartStation 'пикет начала
                            Dim PKEndReg As Double = BlineRegion.EndStation 'пикет конца
                            'если регион находится в диапазоне
                            If PKStartReg >= PKStart And PKEndReg <= PKEnd Then
                                'если регион рабочий
                                If BlineRegion.NeedsProcessing = True Then
                                    'достаем конструкции
                                    Dim AppAssemblCollection As AppliedAssemblyCollection = BlineRegion.AppliedAssemblies
                                    'перебираем коллекцию 
                                    For Each appAssembly As AppliedAssembly In AppAssemblCollection
                                        'записываем пикет
                                        '============================================================================================
                                        Dim PKStat As Double = appAssembly.Points.Item(0).StationOffsetElevationToBaseline.X 'пикет конструкции
                                        If FormPlanStatement.ComboBox6.Text Like "0.00" Then
                                            'усли есть рубленные пикеты тогда вычленяем +
                                            If sqwPiketCollection.Count > 0 Then
                                                Try
                                                    Dim strPk As String = ObjAlign.GetStationStringWithEquations(PKStart)
                                                    Dim valstr As Double = strPk.Replace("+", "")
                                                    xlsheet.Cells(RowInd, 1) = valstr
                                                Catch ex As System.ArgumentException
                                                    xlsheet.Cells(RowInd, 1) = Math.Round(PKStat, RoundZn)
                                                End Try
                                            Else
                                                xlsheet.Cells(RowInd, 1) = Math.Round(PKStat, RoundZn)
                                            End If
                                        Else
                                            Try
                                                Dim PiketNPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(ObjAlign, PKStat, RoundZn)
                                                xlsheet.Cells(RowInd, 1) = PiketNPKStrFormat
                                            Catch ex As System.ArgumentException
                                                xlsheet.Cells(RowInd, 1) = Math.Round(PKStat, RoundZn)
                                            End Try
                                        End If
                                        '==============================================================================================
                                        'записываем расстояние
                                        If RowInd = 4 Then
                                            xlsheet.Cells(RowInd, 2).value = 0
                                        Else
                                            Dim Dist As Double = Math.Abs(PKStat - PKStat0)
                                            PKStat0 = PKStat
                                            xlsheet.Cells(RowInd, 2).value = Math.Round(Dist, RoundZn)
                                        End If
                                        '=========================================================================================================
                                        'получаем из нее коллекцию материалов
                                        Dim ShColl As CalculatedShapeCollection = appAssembly.Shapes
                                        'если марериалы существуют тогда достаем их
                                        If ShColl.Count > 0 Then
                                            For j As Integer = 0 To ShColl.Count - 1
                                                Dim Sh As CalculatedShape = ShColl.Item(j) 'читаем материал
                                                Dim NameSh As String = Sh.CorridorCodes.Item(0) 'имя материала
                                                Dim AreaStr As Double = Sh.Area 'площадь
                                                '============================================================================================================
                                                'записываем в таблицу
                                                'перебираем ячейки таблицы
                                                Dim posCol As Integer = 3 'читаем столбцы с наименованием матерала с 3 позиции
                                                Do
                                                    Dim KodShape As String = xlsheet.Cells(3, posCol).value 'ссылка на ячейку с материалом
                                                    If KodShape Like NameSh Then 'если такой материал уже существует тогда просто записываем значение с соответсвующую строку
                                                        Dim val As Double = CDbl(xlsheet.Cells(RowInd, posCol).value)
                                                        xlsheet.Cells(RowInd, posCol).value = val + AreaStr
                                                        Exit Do
                                                    ElseIf KodShape Like "" Or IsNothing(KodShape) = True Then 'если материал новый тогда проименовываем следующий столбец
                                                        xlsheet.Cells(3, posCol).value = NameSh 'ссылка на ячейку с материалом
                                                        xlsheet.Cells(RowInd, posCol).value = AreaStr
                                                        countMaterial += 1
                                                        Exit Do
                                                    Else
                                                        posCol += 1
                                                    End If
                                                Loop
                                            Next j
                                        End If
                                        RowInd += 1
                                    Next
                                End If
                            End If
                        Next
                        '==================================================================================================
                        'все значения записаны форматируем таблицу
                        If countMaterial > 0 Then
                            'Площадь
                            cell1 = xlsheet.Cells(2, 3)
                            cell2 = xlsheet.Cells(2, 2 + countMaterial)
                            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                            CellRng.Merge(Type.Missing)
                            cell1.Value = "'Площадь, кв.м"
                            'вычисляем объемы
                            cell1 = xlsheet.Cells(2, 3 + countMaterial)
                            cell2 = xlsheet.Cells(2, 2 + countMaterial * 2)
                            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                            CellRng.Merge(Type.Missing)
                            cell1.Value = "'Объем, куб.м"
                            For i As Integer = 0 To countMaterial - 1
                                xlsheet.Cells(3, 3 + i + countMaterial).value = xlsheet.Cells(3, 3 + i).value 'ссылка на ячейку с материалом
                                'вычисляем объем
                                Dim StartPosRow As Integer = 4
                                Dim startPosCol As Integer = 3 + countMaterial + i
                                Do
                                    Dim Dist As String = xlsheet.Cells(StartPosRow, 1).value
                                    If IsNothing(Dist) = True Or Dist = "" Then
                                        Exit Do
                                    End If
                                    Dim posDist As Integer = 0 - (countMaterial + i + 1)
                                    Dim posSqr As Integer = 0 - countMaterial
                                    xlsheet.Cells(StartPosRow, startPosCol).FormulaR1C1 = "=RC[" & posDist & "]*RC[" & posSqr & "]"
                                    StartPosRow += 1
                                Loop
                            Next
                        End If
                        xlsheet = Nothing
                    End If
LineEnd2:
                    locker.Dispose()
                    xlBook = Nothing
                    xlApp.Quit()
                    xlApp = Nothing
                    Exit Sub
                End If
                '=============================================================================================
                'отчет по двум кодам
                '=============================================================================================
                'формируем отчет
                If IsNothing(PointAssembl) = False Then
                    'запускаем эксель
                    Dim xlApp As Excel.Application = Nothing
                    Dim xlBook As Excel.Workbook = Nothing
                    Dim xlSheets As Excel.Worksheet = Nothing
                    Dim cell1 As Excel.Range 'первая ячейка в строке
                    Dim cell2 As Excel.Range 'последняя ячейка в строке
                    Dim CellRng As Excel.Range 'массив ячеек
                    'находим все директории с файлами
                    Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
                    Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
                    'отделяем имя файла от папки
                    Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
                    'определяем имя родительской папки для папки где находится сборка
                    Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
                    'находим папки для блоков, линий и подписей
                    '1. путь к блокам
                    Dim PutchTables As String = "d:\AutodeskCivilTools\TemplateTables\Roads\Ведомости планировки\" 'пути для шаблона таблицы
                    If IO.Directory.Exists(parentPatch2 & "\TemplateTables\Roads\Ведомости планировки\") = True Then
                        PutchTables = parentPatch2 & "\TemplateTables\Roads\Ведомости планировки\"
                    Else
                        If IO.Directory.Exists(PutchTables) = False Then
                            MsgBox("Не удалось найти папку" & parentPatch2 & "\TemplateTables\Roads\Ведомости планировки\")
                            Exit Sub
                        End If
                    End If
                    xlApp = CreateObject("Excel.Application")
                    xlBook = xlApp.Workbooks.Open(PutchTables & "Ведомость планировки.xlsx")
                    xlApp.Visible = True
                    'находим нужный лист(Ведомость объемов работ)
                    Dim xlsheet As Excel.Worksheet
                    xlsheet = xlBook.Worksheets.Item("Ведомость планировки") 'делаем ссылку на этот лист
                    xlsheet.Activate()
                    xlsheet.Cells(2, 1) = "Наименование трассы " & ObjAlign.Name
                    xlsheet.Cells(3, 1) = "Наименование коридора " & ObjCoridor.Name
                    xlsheet.Cells(4, 2) = "'" & Kod1
                    xlsheet.Cells(4, 5) = "'" & Kod2
                    Dim exlPos As Integer = 7
                    For i As Integer = 0 To PointAssembl.GetUpperBound(1)
                        cell1 = xlsheet.Cells(exlPos + i, 1) 'ссылка на первую ячейку
                        cell2 = xlsheet.Cells(exlPos + i, 10) 'ссылка на 6 ячейку в строке
                        CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                        CellRng.EntireRow.Insert() 'вставляем строку
                        Dim PK As Double = Val(PointAssembl(0, i))
                        If i = 0 Then
                            PK = MathFunction.FuncTrimDigitCall(PK, 3)
                        ElseIf i = PointAssembl.GetUpperBound(1) Then
                            PK = MathFunction.FuncTrimDigitFloo(PK, 3)
                        End If
                        '=======================================================================================================
                        'записываем пикет
                        If FormPlanStatement.ComboBox6.Text Like "0.00" Then
                            xlsheet.Cells(exlPos + i, 1) = Val(PK)
                        Else
                            Try
                                xlsheet.Cells(exlPos + i, 1) = ObjAlign.GetStationStringWithEquations(PK) 'форматированный пикет
                            Catch ex As System.ArgumentOutOfRangeException
                            End Try
                        End If
                        xlsheet.Cells(exlPos + i, 2) = Math.Round(Val(PointAssembl(2, i)), RoundZn) 'координата Х кода 1
                        xlsheet.Cells(exlPos + i, 3) = Math.Round(Val(PointAssembl(1, i)), RoundZn) 'координата Y кода 1
                        xlsheet.Cells(exlPos + i, 4) = Math.Round(Val(PointAssembl(3, i)), RoundZn) 'координата Z кода 1
                        xlsheet.Cells(exlPos + i, 5) = Math.Round(Val(PointAssembl(5, i)), RoundZn) 'координата Х кода 2
                        xlsheet.Cells(exlPos + i, 6) = Math.Round(Val(PointAssembl(4, i)), RoundZn) 'координата Y кода 2
                        xlsheet.Cells(exlPos + i, 7) = Math.Round(Val(PointAssembl(6, i)), RoundZn) 'координата Z кода 2
                        Dim LKod1 As Double = 0
                        If FormPlanStatement.RadioButton3.Checked = True Then
                            LKod1 = Math.Round(Val(PointAssembl(7, i)), RoundZn) 'длина линии по коду1
                        ElseIf FormPlanStatement.RadioButton4.Checked = True Then
                            LKod1 = Math.Round(Val(PointAssembl(8, i)), RoundZn) 'длина линии по коду1
                        Else
                            LKod1 = Math.Round(Val(PointAssembl(9, i)), RoundZn) 'длина линии по коду1
                        End If
                        If xlsheet.Cells(exlPos + i, 2).value = 0 And xlsheet.Cells(exlPos + i, 3).value = 0 Or xlsheet.Cells(exlPos + i, 5).value = 0 And xlsheet.Cells(exlPos + i, 6).value = 0 Then
                            xlsheet.Cells(exlPos + i, 8) = 0
                        Else
                            xlsheet.Cells(exlPos + i, 8) = LKod1
                        End If
                        If i < PointAssembl.GetUpperBound(1) Then
                            Dim LKod2 As Double = Math.Round(Val(PointAssembl(7, i + 1)), RoundZn) 'длина линии по коду2
                            Dim PK1 As Double = Val(PointAssembl(0, i))
                            Dim PK2 As Double = Val(PointAssembl(0, i + 1))
                            Dim l As Double = Math.Abs(PK2 - PK1)
                            xlsheet.Cells(exlPos + i, 9) = Math.Round(l, RoundZn) 'длина линии
                            Dim sq As Double = 0.5 * l * (LKod1 + LKod2) 'площадь
                            xlsheet.Cells(exlPos + i, 10) = Math.Round(sq, RoundZn) 'площадь
                        End If
                    Next i
                    xlSheets = Nothing
                    xlBook = Nothing
                    xlApp.Quit()
                    xlApp = Nothing
                End If
                locker.Dispose()
            Catch ex As Exception
                locker.Dispose()
            End Try
        End Using
    End Sub
End Class
