Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.Configuration
Imports Microsoft.Office.Interop
Imports Autodesk.AutoCAD.Colors
'ведомость рабочих отметок по верху покрытия
Public Class ListMarksOnTopCoating
    <CommandMethod("ListMarksOnTopCoating")>
    Public Sub ListMarksOnTopCoating()
        Dim WApp As Word.Application = Nothing
        Dim WDoc As Word.Document = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acCurDb As Database = acDoc.Database
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
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

        Dim DocColectionAlignment As ObjectIdCollection = CivilDoc.GetAlignmentIds
        Dim DocCollectionCorridor As CorridorCollection = CivilDoc.CorridorCollection
        Dim CoridorObj As Corridor = Nothing
        Dim BaseLine As Baseline = Nothing
        Dim AlignObj As Alignment = Nothing
        Dim ProfObj As Autodesk.Civil.DatabaseServices.Profile = Nothing
        Dim FeatLine1 As CorridorFeatureLine = Nothing
        Dim FeatLine2 As CorridorFeatureLine = Nothing
        If DocColectionAlignment.Count = 0 Then
            MsgBox("В данном чертеже нет ТРАСС!!! Макро прервано")
            Exit Sub
        End If
        If DocCollectionCorridor.Count = 0 Then
            MsgBox("В данном чертеже нет КОРИДОРОВ!!! Макро прервано")
            Exit Sub
        End If
        '=====================================================================================================================
        'останавливаем реактор автосохранения
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        'запуск ссылки на диалогового окна
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim FormMarksOnTopCoating As MarksOnTopCoating = New MarksOnTopCoating
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'устанавливаем системную переменную указаннаю пользователем
        Dim boolSistPK As Boolean = False
        Dim boolSistSlope As Boolean = False
        Dim boolSistDist As Boolean = False
        Dim boolSistCoord As Boolean = False
        'смотрим точность системной переменной до запуска программы и сохраняем ее
        Dim SistPeremPK As Integer = CivilDoc.Settings.DrawingSettings.AmbientSettings.Station.Precision.Value 'точность после запятой для пикетажа
        Dim SistPeremDist As Integer = CivilDoc.Settings.DrawingSettings.AmbientSettings.Distance.Precision.Value 'точность после запятой для расстояний
        Dim SistPeremSlope As Integer = CivilDoc.Settings.DrawingSettings.AmbientSettings.Slope.Precision.Value 'точность после запятой для уклонов
        Dim SistPeremElev As Integer = CivilDoc.Settings.DrawingSettings.AmbientSettings.Elevation.Precision.Value 'точность после запятой для отметок

        Using tr
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '1. ВЫБИРАЕМ КОРИДОР
            Dim resBlk As PromptEntityResult = Nothing
            Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите коридор: ")
            resBlk = m_Editor.GetEntity(opt)
            If Not (resBlk.Status = PromptStatus.None Or resBlk.Status = PromptStatus.Cancel Or resBlk.Status = PromptStatus.Error) Then 'нет выбора
                Dim pvId As ObjectId = resBlk.ObjectId
                Dim obj As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(pvId, OpenMode.ForRead)
                If TypeOf obj Is Corridor Then
                    CoridorObj = obj
                    FormMarksOnTopCoating.TextBox3.Text = CoridorObj.Name
                End If
            Else
                MsgBox("Ошибка выбора коридора. Макро прервано!!!")
                locker.Dispose()
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '1.1 вычисляем базовые линии для коридора
            Dim BLines As BaselineCollection = Nothing
            If IsNothing(CoridorObj) = False Then
                If CoridorObj.Id.IsNull = False Then
                    BLines = CoridorObj.Baselines
                    If BLines.Count > 0 Then
                        Dim ArrayBLines As String() = Nothing
                        Dim countArrayBLines As Integer = 0
                        For Each bLine1 As Baseline In BLines
                            Dim NAmeBline As String = bLine1.Name
                            ReDim Preserve ArrayBLines(countArrayBLines)
                            ArrayBLines(countArrayBLines) = NAmeBline
                            countArrayBLines += 1
                        Next
                        If IsArray(ArrayBLines) = True Then
                            FormMarksOnTopCoating.ComboBox1.DataSource = ArrayBLines
                        End If
                    Else
                        MsgBox("В выбранном корридоре нет базовых линий!!! Макро прервано.")
                        locker.Dispose()
                        Exit Sub
                    End If
                Else
                    MsgBox("Ошибка выбора коридора. Макро прервано!!!")
                    locker.Dispose()
                    Exit Sub
                End If
            End If
            '2. ВЫБИРАЕМ ТРАССУ (смотрим трассы для базовых линий)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'массив с трассами
            Dim ArrayAlignName As String() = Nothing
            Dim CountArrayAlignName As Integer = 0
            'массив с профилями
            Dim ArrayProfileName As String() = Nothing
            Dim CountArrayProfileName As Integer = 0
            'массив с группами секций
            Dim ArraySec As String() = Nothing
            Dim CountArraySec As Integer = 0
            If BLines.Count > 0 Then
                For i As Integer = 0 To BLines.Count - 1
                    Dim bline1 As Baseline = BLines.Item(i)
                    'вычисляем на нее трассы и профиля
                    '==========================================================================
                    'трассы
                    Dim AlignId As ObjectId = Nothing
                    Dim Align As Alignment = Nothing
                    AlignId = bline1.AlignmentId
                    If IsNothing(AlignId) = False Then
                        Align = tr.GetObject(AlignId, OpenMode.ForRead)
                        ReDim Preserve ArrayAlignName(CountArrayAlignName)
                        ArrayAlignName(CountArrayAlignName) = Align.Name
                        CountArrayAlignName += 1
                        If CountArrayProfileName = 0 Then
                            Dim StartPK As Double = Align.StartingStation
                            Dim EndPK As Double = Align.EndingStation
                            FormMarksOnTopCoating.TextBox1.Tag = StartPK
                            FormMarksOnTopCoating.TextBox2.Tag = EndPK
                            Try
                                FormMarksOnTopCoating.TextBox1.Text = Align.GetStationStringWithEquations(StartPK)
                            Catch ex As System.ArgumentException
                                FormMarksOnTopCoating.TextBox1.Text = StartPK
                            End Try
                            Try
                                FormMarksOnTopCoating.TextBox2.Text = Align.GetStationStringWithEquations(EndPK)
                            Catch ex As System.ArgumentException
                                FormMarksOnTopCoating.TextBox2.Text = EndPK
                            End Try
                        End If
                    End If
                    '===========================================================================
                    'профиля
                    Dim ProfId As ObjectId = Nothing
                    If IsNothing(ProfId) = False Then
                        ProfId = bline1.ProfileId
                        Dim Prof As Autodesk.Civil.DatabaseServices.Profile = tr.GetObject(ProfId, OpenMode.ForRead)
                        ReDim Preserve ArrayProfileName(CountArrayProfileName)
                        ArrayProfileName(CountArrayProfileName) = Prof.Name
                        CountArrayProfileName += 1
                    End If
                    '=================================================================================
                    'смотрим группы сечений
                    Dim SampleLineGroupCollection As ObjectIdCollection = Align.GetSampleLineGroupIds
                    If SampleLineGroupCollection.Count > 0 Then
                        For Each sampleLineId As ObjectId In SampleLineGroupCollection
                            If IsNothing(sampleLineId) = False Then
                                If sampleLineId.IsNull = False Then
                                    Dim sampleGroupLine As SampleLineGroup = tr.GetObject(sampleLineId, OpenMode.ForRead)
                                    ReDim Preserve ArraySec(CountArraySec)
                                    ArraySec(CountArraySec) = sampleGroupLine.Name
                                    CountArraySec += 1
                                End If
                            End If
                        Next
                    End If
                Next i
            End If
            If IsArray(ArrayAlignName) = False Then
                MsgBox("В текущем чертеже ТРАСС нет. Макро прервано!!!")
                locker.Dispose()
                Exit Sub
            End If
            If IsArray(ArrayProfileName) = False Then
                MsgBox("В текущем чертеже ПРОФИЛЕЙ нет. Макро прервано!!!")
                locker.Dispose()
                Exit Sub
            End If
            FormMarksOnTopCoating.ComboBox6.DataSource = ArrayAlignName
            FormMarksOnTopCoating.ComboBox6.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
            FormMarksOnTopCoating.ComboBox7.DataSource = ArrayProfileName
            FormMarksOnTopCoating.ComboBox7.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '5. ИЩЕМ КОДЫ ХАРАКТЕРНЫХ ЛИНИЙ
            Dim ArrayCode As String() = {""}
            Dim countCode As Integer = 1
            Dim KodCentre As String = ""
            Dim countRow As Integer = 0
            If CoridorObj.FeatureLineCodeInfos.Count > 0 Then
                For Each featureLine As FeatureLineCodeInfo In CoridorObj.FeatureLineCodeInfos
                    Dim CodeNAme As String = featureLine.CodeName
                    FormMarksOnTopCoating.DataGridView1.Rows.Insert(countRow)
                    FormMarksOnTopCoating.DataGridView1.Rows(countRow).Cells(2).Value = CodeNAme
                    countRow += 1
                    ReDim Preserve ArrayCode(countCode)
                    ArrayCode(countCode) = CodeNAme
                    countCode += 1
                Next
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'заполняем список
            FormMarksOnTopCoating.ComboBox2.DataSource = ArrayCode
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If IsArray(ArraySec) = True Then
                FormMarksOnTopCoating.ComboBox3.DataSource = ArraySec
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '===============================================================================================================
            'запускаем диалоговое окно
            FormMarksOnTopCoating.ShowDialog()
            If FormMarksOnTopCoating.boolProf = False Then Exit Sub
            If FormMarksOnTopCoating.boolProf = False Then
                locker.Dispose()
                MsgBox("Макро прервано пользователем!!!")
                Exit Sub
            End If
            Dim Shag As Integer = FormMarksOnTopCoating.NumericUpDown1.Value
            'блок с точностью данных
            '================================================================================================================
            Dim RoundZnDist As Integer = FormMarksOnTopCoating.NumericUpDown5.Value 'точность для расстояний
            Dim RoundZnSlope As Integer = FormMarksOnTopCoating.NumericUpDown4.Value 'точность уклонов
            Dim RoundZnCoord As Integer = FormMarksOnTopCoating.NumericUpDown2.Value 'точность для координат
            Dim RoundZnPK As Integer = FormMarksOnTopCoating.NumericUpDown6.Value 'точность для пикетов
            Dim RoundZnElev As Integer = FormMarksOnTopCoating.NumericUpDown3.Value 'точность для отметок
            If SistPeremPK <> RoundZnPK Then 'устанавливаем новое значение
                CivilDoc.Settings.DrawingSettings.AmbientSettings.Station.Precision.Value = RoundZnPK
                boolSistPK = True
            End If
            '=================================================================================================================
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим характерные линии коридора
            Dim ArrayNameCodeLeft As String(,) = Nothing
            Dim CountArrayNameCodeLeft As Integer = 0
            Dim ArrayNameCodeRight As String(,) = Nothing
            Dim CountArrayNameCodeRight As Integer = 0
            If FormMarksOnTopCoating.DataGridView1.RowCount > 2 Then
                For i As Integer = 0 To FormMarksOnTopCoating.DataGridView1.RowCount - 1
                    If Val(FormMarksOnTopCoating.DataGridView1.Rows(i).Cells(0).Value) = 1 Then
                        Dim CodePoint As String = FormMarksOnTopCoating.DataGridView1.Rows(i).Cells(3).Value
                        If Left(CodePoint, 1) = "T" Or Left(CodePoint, 1) = "Т" Then
                            ReDim Preserve ArrayNameCodeLeft(2, CountArrayNameCodeLeft)
                            ArrayNameCodeLeft(0, CountArrayNameCodeLeft) = Val(Mid(CodePoint, 2)) 'порядковый номер
                            ArrayNameCodeLeft(1, CountArrayNameCodeLeft) = Val(FormMarksOnTopCoating.DataGridView1.Rows(i).Cells(2).Value) 'флаг на вычисление уклонов
                            ArrayNameCodeLeft(2, CountArrayNameCodeLeft) = FormMarksOnTopCoating.DataGridView1.Rows(i).Cells(2).Value 'код точки
                            CountArrayNameCodeLeft += 1
                        End If
                    End If
                Next i
                'для правой стороны
                For i As Integer = 0 To FormMarksOnTopCoating.DataGridView1.RowCount - 1
                    If Val(FormMarksOnTopCoating.DataGridView1.Rows(i).Cells(1).Value) = 1 Then
                        Dim CodePoint As String = FormMarksOnTopCoating.DataGridView1.Rows(i).Cells(4).Value
                        If Left(CodePoint, 1) = "T" Or Left(CodePoint, 1) = "Т" Then
                            ReDim Preserve ArrayNameCodeRight(2, CountArrayNameCodeRight)
                            ArrayNameCodeRight(0, CountArrayNameCodeRight) = Val(Mid(CodePoint, 2)) 'порядковый номер
                            ArrayNameCodeRight(1, CountArrayNameCodeRight) = Val(FormMarksOnTopCoating.DataGridView1.Rows(i).Cells(2).Value) 'флаг на вычисление уклонов
                            ArrayNameCodeRight(2, CountArrayNameCodeRight) = FormMarksOnTopCoating.DataGridView1.Rows(i).Cells(2).Value 'код точки
                            CountArrayNameCodeRight += 1
                        End If
                    End If
                Next i
            End If
            'сортируем массивы
            If IsArray(ArrayNameCodeLeft) = True Then
                If ArrayNameCodeLeft.GetUpperBound(1) > 0 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayNameCodeLeft, 0)
                End If
            End If
            If IsArray(ArrayNameCodeRight) = True Then
                If ArrayNameCodeRight.GetUpperBound(1) > 0 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayNameCodeRight, 0)
                End If
            End If
            'находим базовую линию
            Dim bLine As Baseline = Nothing
            Dim bLineCollection As BaselineCollection = CoridorObj.Baselines
            For Each bLine In bLineCollection
                If bLine.Name Like FormMarksOnTopCoating.ComboBox1.Text Then
                    Exit For
                End If
            Next
            If IsNothing(bLine) = True Then
                MsgBox("Не удалось найти характерные линии для выбранного корридора. Макро прервано!!!")
                locker.Dispose()
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим трассу для выбранной базовой линии
            Dim AlignId1 As ObjectId = bLine.AlignmentId
            AlignObj = tr.GetObject(AlignId1, OpenMode.ForRead)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим линию профиля
            Dim ProfId1 As ObjectId = bLine.ProfileId
            ProfObj = tr.GetObject(ProfId1, OpenMode.ForRead)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'пикетаж
            Dim PKStartAlign As Double = FormMarksOnTopCoating.TextBox1.Tag
            Dim PKEndAlign As Double = FormMarksOnTopCoating.TextBox2.Tag
            '=============================================================================================================
            'регионы для базовой линии
            Dim reg As BaselineRegionCollection = bLine.BaselineRegions
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'в случае если выбрана опция группы сечений то нет смысла выводить графу примечание
            If FormMarksOnTopCoating.RadioButton5.Checked = True Then FormMarksOnTopCoating.CheckBox6.Checked = False
            '=============================================================================================================
            '============================================================
            'перебираем все элементы трассы для их сортировки в процессе возрастания
            Dim ArrayPKResult As Double(,) = Nothing 'результирующий массив (0-RAW пикет, 1-конец элемента прыдущего, 2-начало следующего элемента, 3-рубленный пикет)
            Dim CountArrayResult As Integer = 0
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'этот блок запускается в случае если выбрана опция ПРИМЕЧАНИЕ
            Dim ArrayEntity As Double(,) = Nothing
            If FormMarksOnTopCoating.CheckBox6.Checked = True Then
                'массив для элементов трассы
                For Each myAe As AlignmentEntity In AlignObj.Entities
                    Dim AlignEnt As AlignmentEntityType = myAe.EntityType
                    Select Case AlignEnt
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'прямой участок
                        Case AlignmentEntityType.Line 'прямой участок
                            Dim AlignLine As AlignmentLine = myAe
                            Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsLine(AlignObj, AlignLine, ArrayEntity, PKStartAlign, PKEndAlign)
     '===========================================================================================================================
'дуга одиночная
                        Case AlignmentEntityType.Arc
                            Dim AlignArc As AlignmentArc = myAe
                            Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsArc(AlignObj, AlignArc, ArrayEntity, PKStartAlign, PKEndAlign)
'========================================================================================================================================
'спираль одиночная
                        Case AlignmentEntityType.Spiral
                            If TypeOf myAe Is AlignmentSTS Then
                                Dim AlignSTS As AlignmentSTS = myAe 'составная спираль спираль дуга спираль
                                Dim AlignS1 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    AlignS1 = AlignSTS.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                Try
                                    AlignS1 = AlignSTS.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                                'если спираль существует
                                If IsNothing(AlignS1) = False Then
                                    Dim boolSpcChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, AlignS1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                '==================================================================================================================================
                            ElseIf TypeOf myAe Is AlignmentSCS Then
                                Dim AlignSCS As AlignmentSCS = myAe 'составная спираль спираль прямая спираль
                                Dim AlignS2 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    AlignS2 = AlignSCS.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                Try
                                    AlignS2 = AlignSCS.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                                'если спираль существует
                                If IsNothing(AlignS2) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, AlignS2, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                '=================================================================================================
                            ElseIf TypeOf myAe Is AlignmentSpiral Then
                                Dim AlignS3 As AlignmentSpiral = Nothing 'отдельная спираль
                                Try
                                    AlignS3 = myAe
                                Catch ex As System.InvalidOperationException
                                End Try
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSpiral(AlignObj, AlignS3, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'=================================================================================================
'спираль дуга спираль
                        Case AlignmentEntityType.SpiralCurveSpiral
                            Dim AlignSCS As AlignmentSCS = myAe
                            Dim Spin As AlignmentSubEntitySpiral = Nothing
                            Try
                                Spin = AlignSCS.SpiralIn
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Spin) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Spin, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignSCS.Arc
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim SpOut As AlignmentSubEntitySpiral = Nothing
                            Try
                                SpOut = AlignSCS.SpiralOut
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(SpOut) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, SpOut, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'=============================================================================================================================
'переходная кривая прямая переходная кривая
                        Case AlignmentEntityType.SpiralLineSpiral
                            Dim AlignSLS As AlignmentSTS = myAe
                            Dim Spin As AlignmentSubEntitySpiral = Nothing
                            Try
                                Spin = AlignSLS.SpiralIn
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Spin) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Spin, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Line1 As AlignmentSubEntityLine = Nothing
                            Try
                                Line1 = AlignSLS.Tangent
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Line1) = False Then
                                Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsSubLine(AlignObj, Line1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim SpOut As AlignmentSubEntitySpiral = Nothing
                            Try
                                SpOut = AlignSLS.SpiralOut
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(SpOut) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, SpOut, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'=========================================================================================================================
'переходная кривая прямая
                        Case AlignmentEntityType.SpiralLine
                            Dim AlignSL As AlignmentSTS = myAe
                            Dim Spin As AlignmentSubEntitySpiral = Nothing
                            Try
                                Spin = AlignSL.SpiralIn
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Spin) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Spin, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Line1 As AlignmentSubEntityLine = Nothing
                            Try
                                Line1 = AlignSL.Tangent
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Line1) = False Then
                                Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsSubLine(AlignObj, Line1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'======================================================================================================================
'прямая спираль 
                        Case AlignmentEntityType.LineSpiral
                            Dim AlignLS As AlignmentSTS = myAe
                            Dim Line1 As AlignmentSubEntityLine = Nothing
                            Try
                                Line1 = AlignLS.Tangent
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Line1) = False Then
                                Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsSubLine(AlignObj, Line1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim SpOut As AlignmentSubEntitySpiral = Nothing
                            Try
                                SpOut = AlignLS.SpiralOut
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(SpOut) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, SpOut, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'===================================================================================================================================
'переходная кривая дуга 
                        Case AlignmentEntityType.SpiralCurve
                            Dim AlignSC As AlignmentSCS = myAe
                            Dim Spin As AlignmentSubEntitySpiral = Nothing
                            Try
                                Spin = AlignSC.SpiralIn
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Spin) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Spin, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignSC.Arc
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'==========================================================================================================================================
                        Case AlignmentEntityType.CurveSpiral 'кривая переходная кривая
                            Dim AlignCS As AlignmentSCS = myAe
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignCS.Arc
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim SpOut As AlignmentSubEntitySpiral = AlignCS.SpiralOut
                            Try
                                SpOut = AlignCS.SpiralOut
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(SpOut) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, SpOut, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'==============================================================================================================================================
'переходная кривая переходная кривая дуга переходная кривая переходная кривая
                        Case AlignmentEntityType.SpiralSpiralCurveSpiralSpiral
                            Dim AlignSSCSS As AlignmentSSCSS = myAe
                            Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp1 = AlignSSCSS.Spiral1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp1) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp2 = AlignSSCSS.Spiral2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp2) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignSSCSS.Arc
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp3 = AlignSSCSS.Spiral3
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp3) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp3, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp4 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp4 = AlignSSCSS.Spiral4
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp4) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp4, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'====================================================================================================================================
'дуга дуга обратная дуга
                        Case AlignmentEntityType.CurveCurveReverseCurve
                            Dim AlignCCRC As AlignmentCCRC = myAe
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignCCRC.Arc1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc2 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc2 = AlignCCRC.Arc2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc2) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc3 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc3 = AlignCCRC.Arc3
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc3) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc3, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'==============================================================================================================================
'дуга обратная дуга
                        Case AlignmentEntityType.CurveReverseCurve
                            Dim AlignCRC As AlignmentCRC = myAe
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignCRC.Arc1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc2 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc2 = AlignCRC.Arc2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc2) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'==================================================================================================================================
'кривавая спираль спираль
                        Case AlignmentEntityType.CurveSpiralSpiral
                            Dim AlignCSS As AlignmentSSCSS = myAe
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignCSS.Arc
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp3 = AlignCSS.Spiral3
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp3) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp3, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp4 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp4 = AlignCSS.Spiral4
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp4) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp4, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'==========================================================================================================================
'спираль спираль
                        Case AlignmentEntityType.SpiralSpiral
                            Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                            Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                            If TypeOf myAe Is AlignmentSTS Then
                                Dim AlignCS1 As AlignmentSTS = myAe
                                Try
                                    Sp1 = AlignCS1.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                Try
                                    Sp2 = AlignCS1.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                            ElseIf TypeOf myAe Is AlignmentSCS Then
                                Dim AlignCS2 As AlignmentSCS = myAe
                                Try
                                    Sp1 = AlignCS2.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                Try
                                    Sp2 = AlignCS2.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                            End If
                            If IsNothing(Sp1) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            If IsNothing(Sp2) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'=============================================================================================================================
'спираль спираль дуга
                        Case AlignmentEntityType.SpiralSpiralCurve
                            Dim AlignSSC As AlignmentSSCSS = myAe
                            Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp1 = AlignSSC.Spiral1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp1) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp2 = AlignSSC.Spiral2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp2) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignSSC.Arc
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'================================================================================================================================
'спираль дуга спираль спираль дуга спираль
                        Case AlignmentEntityType.SpiralCurveSpiralSpiralCurveSpiral
                            Dim AlignSCSSCS As AlignmentSCSSCS = myAe
                            Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp1 = AlignSCSSCS.Spiral1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp1) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignSCSSCS.Arc1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp2 = AlignSCSSCS.Spiral2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp2) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp3 = AlignSCSSCS.Spiral3
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp3) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp3, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc2 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc2 = AlignSCSSCS.Arc2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc2) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp4 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp4 = AlignSCSSCS.Spiral4
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp4) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp4, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
'==================================================================================================================================================
'спираль дуга спираль дуга спираль
                        Case AlignmentEntityType.SpiralCurveSpiralCurveSpiral
                            Dim AlignSCSCS As AlignmentSCSCS = myAe
                            Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp1 = AlignSCSCS.Spiral1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp1) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignSCSCS.Arc1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp2 = AlignSCSCS.Spiral2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp2) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc2 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc2 = AlignSCSCS.Arc2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc2) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp3 = AlignSCSCS.Spiral3
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp3) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(AlignObj, Sp3, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
 '=================================================================================================================                            
                        Case AlignmentEntityType.MultipleSegments
                            Dim AlignCS As AlignmentMultipleSegments = myAe
                            'структура не понятна
'=================================================================================================================
                        Case AlignmentEntityType.CurveLineCurve
                            Dim AlignCLC As AlignmentCTC = myAe
                            Dim Arc1 As AlignmentSubEntityArc = AlignCLC.Arc1
                            Try
                                Arc1 = AlignCLC.Arc1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Line1 As AlignmentSubEntityLine = Nothing
                            Try
                                Line1 = AlignCLC.Tangent
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Line1) = False Then
                                Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsSubLine(AlignObj, Line1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc2 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc2 = AlignCLC.Arc1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc2) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(AlignObj, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                    End Select
                Next
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'второй этап
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'сортируем массив
                If IsArray(ArrayEntity) = True Then
                    Dim boolArray As Boolean = MathFunction.FuncSortDblArray(ArrayEntity, 0)
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim ArrayXYZ As Double(,) = Nothing
            Dim CountArrayXYZ As Integer = 0
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'запускаем цикл по элементам найденным в массиве
            If IsArray(ArrayEntity) = True Then
                For i As Integer = 0 To ArrayEntity.GetUpperBound(1)
                    '1. формируем массив с начальными данными
                    Dim ArrayObject As Double() = Nothing
                    ReDim ArrayObject(2)
                    'записываем характеристику начала элемента
                    ArrayObject(0) = ArrayEntity(0, i) 'начало элемента
                    'записываем характеристику конца элемента
                    ArrayObject(1) = ArrayEntity(1, i)
                    'тип элемента
                    Dim TypeObj As Integer = ArrayEntity(4, i)
                    If TypeObj = 0 Then
                        ArrayObject(2) = 0
                    ElseIf TypeObj = 1 Then
                        ArrayObject(2) = 1
                    ElseIf TypeObj = 2 Then
                        ArrayObject(2) = 2
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'если выбрана опция по пикетам тогда бъем элемент трассы в соответсвии с шагом
                    If FormMarksOnTopCoating.RadioButton1.Checked = True Then
                        '=====================================================================================================================
                        'получаем пикеты с заданным шагом по всем элементам
                        Dim TempArrayXYZ As Double(,) = Nothing
                        Dim boolFuncPK As Boolean = FuncAlignment.FuncSplitAlignObjectRublPK(AlignObj, ArrayObject, TempArrayXYZ, Shag)
                        '=====================================================================================================================
                        'записываем элементы в сводный массив
                        If IsArray(TempArrayXYZ) = True Then
                            For i3 As Integer = 0 To TempArrayXYZ.GetUpperBound(1)
                                ReDim Preserve ArrayXYZ(1, CountArrayXYZ)
                                ArrayXYZ(0, CountArrayXYZ) = TempArrayXYZ(0, i3)
                                ArrayXYZ(1, CountArrayXYZ) = TempArrayXYZ(1, i3)
                                CountArrayXYZ += 1
                            Next i3
                        End If
                        System.Windows.Forms.Application.DoEvents()
                    Else
                        If TypeObj = 0 Then
                            ReDim Preserve ArrayXYZ(1, CountArrayXYZ)
                            ArrayXYZ(0, CountArrayXYZ) = ArrayObject(0)
                            ArrayXYZ(1, CountArrayXYZ) = 1
                            CountArrayXYZ += 1
                            ReDim Preserve ArrayXYZ(1, CountArrayXYZ)
                            ArrayXYZ(0, CountArrayXYZ) = ArrayObject(1)
                            ArrayXYZ(1, CountArrayXYZ) = 2
                            CountArrayXYZ += 1
                        ElseIf TypeObj = 1 Then
                            ReDim Preserve ArrayXYZ(1, CountArrayXYZ)
                            ArrayXYZ(0, CountArrayXYZ) = ArrayObject(0)
                            ArrayXYZ(1, CountArrayXYZ) = 3
                            CountArrayXYZ += 1
                            ReDim Preserve ArrayXYZ(1, CountArrayXYZ)
                            ArrayXYZ(0, CountArrayXYZ) = ArrayObject(1)
                            ArrayXYZ(1, CountArrayXYZ) = 4
                            CountArrayXYZ += 1
                        ElseIf TypeObj = 2 Then
                            ReDim Preserve ArrayXYZ(1, CountArrayXYZ)
                            ArrayXYZ(0, CountArrayXYZ) = ArrayObject(0)
                            ArrayXYZ(1, CountArrayXYZ) = 5
                            CountArrayXYZ += 1
                            ReDim Preserve ArrayXYZ(1, CountArrayXYZ)
                            ArrayXYZ(0, CountArrayXYZ) = ArrayObject(1)
                            ArrayXYZ(1, CountArrayXYZ) = 6
                            CountArrayXYZ += 1
                        End If
                    End If
                Next i
            End If
            '============================================================================================================================
            'заталкиваем в сводный массив рубленные пикеты
            Dim stationEqw As StationEquationCollection = AlignObj.StationEquations
            If FormMarksOnTopCoating.CheckBox6.Checked = True Then
                If stationEqw.Count > 0 Then
                    For i As Integer = 0 To stationEqw.Count - 1
                        Dim PKRubl As Double = stationEqw.Item(i).RawStationBack
                        ReDim Preserve ArrayXYZ(1, CountArrayXYZ)
                        ArrayXYZ(0, CountArrayXYZ) = PKRubl
                        ArrayXYZ(1, CountArrayXYZ) = 7
                        CountArrayXYZ += 1
                    Next i
                End If
            End If
            '================================================================================================================================
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'сортируем массив
            If IsArray(ArrayXYZ) = True Then
                Dim boolArray As Boolean = MathFunction.FuncSortDblArray(ArrayXYZ, 0)
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '===================================================================================================================================
            'удаляем ненужное или дубликаты
            If IsArray(ArrayXYZ) = True Then
                For i As Integer = 0 To ArrayXYZ.GetUpperBound(1)
                    'проверяем пикет на вхождение региона
                    Dim PK As Double = ArrayXYZ(0, i)
                    Dim KodPK As Integer = ArrayXYZ(1, i)
                    'проверяем на вхождение в диапазон
                    Dim boolRegion As Boolean = False
                    For Each BlineRegion As BaselineRegion In reg
                        If PK >= BlineRegion.StartStation And PK <= BlineRegion.EndStation Then
                            If BlineRegion.NeedsProcessing = True Then
                                boolRegion = True
                            End If
                            Exit For
                        End If
                    Next
                    '====================================================
                    If boolRegion = True Then
                        If IsArray(ArrayPKResult) = True Then
                            Dim TempPK As Double = ArrayPKResult(0, CountArrayResult - 1)
                            If Math.Abs(TempPK - PK) < 0.001 Then
                                'пикеты совпадают тогда берем только код (начало прямой, дуги или спирали)
                                If KodPK = 1 Or KodPK = 3 Or KodPK = 5 Then
                                    ArrayPKResult(2, CountArrayResult - 1) = KodPK 'начало элемента
                                ElseIf KodPK = 2 Or KodPK = 4 Or KodPK = 6 Then
                                    ArrayPKResult(1, CountArrayResult - 1) = KodPK 'конец элемента
                                ElseIf KodPK = 7 Then
                                    ArrayPKResult(3, CountArrayResult - 1) = KodPK 'рубленный пикет
                                End If
                            Else
                                ReDim Preserve ArrayPKResult(3, CountArrayResult)
                                ArrayPKResult(0, CountArrayResult) = PK
                                If KodPK = 1 Or KodPK = 3 Or KodPK = 5 Then
                                    ArrayPKResult(2, CountArrayResult) = KodPK 'начало элемента
                                ElseIf KodPK = 2 Or KodPK = 4 Or KodPK = 6 Then
                                    ArrayPKResult(1, CountArrayResult) = KodPK 'конец элемента
                                ElseIf KodPK = 7 Then
                                    ArrayPKResult(3, CountArrayResult) = KodPK 'рубленный пикет
                                End If
                                CountArrayResult += 1
                            End If
                        Else
                            ReDim Preserve ArrayPKResult(3, CountArrayResult)
                            ArrayPKResult(0, CountArrayResult) = PK
                            If KodPK = 1 Or KodPK = 3 Or KodPK = 5 Then
                                ArrayPKResult(2, CountArrayResult) = KodPK 'начало элемента
                            ElseIf KodPK = 2 Or KodPK = 4 Or KodPK = 6 Then
                                ArrayPKResult(1, CountArrayResult) = KodPK 'конец элемента
                            ElseIf KodPK = 7 Then
                                ArrayPKResult(3, CountArrayResult) = KodPK 'рубленный пикет
                            End If
                            CountArrayResult += 1
                        End If
                    End If
                Next i
                Erase ArrayEntity
                Erase ArrayXYZ
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ОПЦИЯ ПО ГРУППАМ СЕЧЕНИЙ
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If FormMarksOnTopCoating.RadioButton5.Checked = True Then
                If IsArray(ArrayPKResult) = True Then Erase ArrayPKResult
                CountArrayResult = 0
                ' Получаем доступ к группам секций
                Dim SampleLineGroupCollection As ObjectIdCollection = AlignObj.GetSampleLineGroupIds
                If SampleLineGroupCollection.Count > 0 Then
                    For Each sampleLineId As ObjectId In SampleLineGroupCollection
                        Dim sampleGroupLine As SampleLineGroup = tr.GetObject(sampleLineId, OpenMode.ForRead)
                        If sampleGroupLine.Name Like FormMarksOnTopCoating.ComboBox3.Text Then
                            Dim IDSectionCollection As ObjectIdCollection = sampleGroupLine.GetSampleLineIds
                            For Each idSample As ObjectId In IDSectionCollection
                                Dim SamplLine As SampleLine = tr.GetObject(idSample, OpenMode.ForRead)
                                Dim PK As Double = SamplLine.Station
                                '======================================================================
                                If PK >= PKStartAlign And PK <= PKEndAlign Then 'если пикет входит в пользовательский диапазон
                                    '==========================================================================================
                                    'проверяем пикет на вхождение в регион
                                    Dim boolRegion As Boolean = False
                                    For Each BlineRegion As BaselineRegion In reg
                                        If PK >= BlineRegion.StartStation And PK <= BlineRegion.EndStation Then
                                            If BlineRegion.NeedsProcessing = True Then
                                                boolRegion = True
                                            End If
                                            Exit For
                                        End If
                                    Next
                                    If boolRegion = True Then
                                        ReDim Preserve ArrayPKResult(3, CountArrayResult)
                                        ArrayPKResult(0, CountArrayResult) = PK
                                        ArrayPKResult(1, CountArrayResult) = 0
                                        ArrayPKResult(2, CountArrayResult) = 0
                                        ArrayPKResult(3, CountArrayResult) = 0
                                        CountArrayResult += 1
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If

                '======================================================================
                'сортируем массив
                If IsArray(ArrayPKResult) = True Then
                    Dim boolArray As Boolean = MathFunction.FuncSortDblArray(ArrayPKResult, 0)
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'запускаем Excel и формируем шапку таблицы
            'находим все директории с файлами
            'запускаем эксель
            Dim xlApp As Excel.Application = Nothing
            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Worksheet = Nothing
            Dim cell1 As Excel.Range = Nothing 'первая ячейка в строке
            Dim cell2 As Excel.Range = Nothing 'последняя ячейка в строке
            Dim CellRng As Excel.Range = Nothing 'массив ячеек
            xlApp = CreateObject("Excel.Application")
            xlBook = xlApp.Workbooks.Add()
            xlApp.Visible = True
            'находим нужный лист(Ведомость объемов работ)
            Dim xlsheet As Excel.Worksheet = xlBook.Sheets.Item("Лист1")
            xlsheet.Activate()
            '==================================================================================
            'формируем шапку таблицы
            '===============================================================================
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'верхняя ячейка
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выводим столбец с пикетажем
            xlsheet.Cells(1, 2) = "'ПК…+…" 'номер пикета
            cell1 = xlsheet.Cells(1, 2) 'ссылка на первую ячейку
            cell2 = xlsheet.Cells(4, 2) 'ссылка на 6 ячейку в строке
            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
            CellRng.Merge(Type.Missing)
            Dim PosStartPKRow = 2 'стартовый столбец для ПК
            '====================================================================================
            'смотрим количество кодов для стороны лево и стороны право
            Dim LenArrayLeft As Integer = 0
            Dim LenArrayRight As Integer = 0
            Dim CountDistCells As Integer = 0
            Dim PosEndDistRow As Integer = 0
            If IsArray(ArrayNameCodeLeft) = True Then
                LenArrayLeft = ArrayNameCodeLeft.GetUpperBound(1) + 1 'количество ячеек лево
            End If
            If IsArray(ArrayNameCodeRight) = True Then
                LenArrayRight = ArrayNameCodeRight.GetUpperBound(1) + 1 'количество ячеек право
            End If
            CountDistCells = LenArrayLeft + LenArrayRight 'общее количество ячеек
            '===============================================================================
            'смотрим нужно ли выводить расстояния
            Dim PosStartDistRow As Integer = -1
            Dim PosStartOtmRow As Integer = -1
            Dim PosStartUklRow As Integer = -1
            Dim PosStartCoordRow As Integer = -1
            Dim PosStartNoteRow As Integer = -1
            Dim PosEnd As Integer = 2
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'расстояния
            If FormMarksOnTopCoating.CheckBox2.Checked = True And CountDistCells > 0 Then
                'стартовый столбец для расстояний
                PosStartDistRow = PosStartPKRow + 1 'начало расстояний
                PosEnd = PosStartDistRow + CountDistCells - 1 'начало отметок
                xlsheet.Cells(2, PosStartDistRow) = "'Расстояние от оси, м" 'номер пикета
                cell1 = xlsheet.Cells(1, PosStartDistRow) 'ссылка на первую ячейку
                cell2 = xlsheet.Cells(2, PosEnd) 'ссылка на 6 ячейку в строке
                CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                '=================================================================================
                'формируем лево и право
                'вторая ячейка
                If LenArrayLeft > 0 Then
                    xlsheet.Cells(3, PosStartDistRow) = "'Левая сторона" 'номер пикета
                    cell1 = xlsheet.Cells(3, PosStartDistRow) 'ссылка на первую ячейку
                    cell2 = xlsheet.Cells(3, PosStartPKRow + LenArrayLeft) 'ссылка на 6 ячейку в строке
                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                    CellRng.Merge(Type.Missing)
                End If
                If LenArrayRight > 0 Then
                    xlsheet.Cells(3, PosStartPKRow + LenArrayLeft + 1) = "'Правая сторона" 'номер пикета
                    cell1 = xlsheet.Cells(3, PosStartPKRow + LenArrayLeft + 1) 'ссылка на первую ячейку
                    cell2 = xlsheet.Cells(3, PosEnd) 'ссылка на 6 ячейку в строке
                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                    CellRng.Merge(Type.Missing)
                End If
                '=======================================================================================
                'формируем коды
                Dim ind As Integer = 0
                If LenArrayLeft > 0 Then
                    For i As Integer = ArrayNameCodeLeft.GetUpperBound(1) To 0 Step -1
                        xlsheet.Cells(4, PosStartDistRow + ind) = "'T" & ArrayNameCodeLeft(0, i) & vbLf & ArrayNameCodeLeft(2, i) 'номер пикета
                        ind += 1
                    Next i
                End If
                If LenArrayRight > 0 Then
                    For i As Integer = 0 To ArrayNameCodeRight.GetUpperBound(1)
                        xlsheet.Cells(4, PosStartDistRow + ind) = "'T" & ArrayNameCodeRight(0, i) & vbLf & ArrayNameCodeRight(2, i) 'номер пикета
                        ind += 1
                    Next i
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'отметки
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If FormMarksOnTopCoating.CheckBox3.Checked = True Then
                PosStartOtmRow = PosEnd + 1  'начало отметок
                PosEnd = PosStartOtmRow + CountDistCells  'конечная отметка
                '======================================================================================
                xlsheet.Cells(2, PosStartOtmRow) = "'Отметка, м" 'номер пикета
                cell1 = xlsheet.Cells(1, PosStartOtmRow) 'ссылка на первую ячейку
                cell2 = xlsheet.Cells(2, PosEnd) 'ссылка на 6 ячейку в строке
                CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                '=========================================================================================
                If LenArrayLeft > 0 Then
                    xlsheet.Cells(3, PosStartOtmRow) = "'Левая сторона" 'номер пикета
                    cell1 = xlsheet.Cells(3, PosStartOtmRow) 'ссылка на первую ячейку
                    cell2 = xlsheet.Cells(3, PosStartOtmRow + LenArrayLeft - 1) 'ссылка на 6 ячейку в строке
                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                    CellRng.Merge(Type.Missing)
                End If
                '=========================================================================================
                xlsheet.Cells(3, PosStartOtmRow + LenArrayLeft) = "'Ось" 'номер пикета
                cell1 = xlsheet.Cells(3, PosStartOtmRow + LenArrayLeft) 'ссылка на первую ячейку
                cell2 = xlsheet.Cells(4, PosStartOtmRow + LenArrayLeft) 'ссылка на 6 ячейку в строке
                CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                '=========================================================================================
                If LenArrayRight > 0 Then
                    xlsheet.Cells(3, PosStartOtmRow + LenArrayLeft + 1) = "'Правая сторона" 'номер пикета
                    cell1 = xlsheet.Cells(3, PosStartOtmRow + LenArrayLeft + 1) 'ссылка на первую ячейку
                    cell2 = xlsheet.Cells(3, PosEnd) 'ссылка на 6 ячейку в строке
                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                    CellRng.Merge(Type.Missing)
                End If
                '=======================================================================================
                'формируем коды
                Dim ind As Integer = 0
                If LenArrayLeft > 0 Then
                    For i As Integer = ArrayNameCodeLeft.GetUpperBound(1) To 0 Step -1
                        xlsheet.Cells(4, PosStartOtmRow + ind) = "'T" & ArrayNameCodeLeft(0, i) & vbLf & ArrayNameCodeLeft(2, i) 'номер пикета
                        ind += 1
                    Next i
                End If
                ind = ind + 1
                If LenArrayRight > 0 Then
                    For i As Integer = 0 To ArrayNameCodeRight.GetUpperBound(1)
                        xlsheet.Cells(4, PosStartOtmRow + ind) = "'T" & ArrayNameCodeRight(0, i) & vbLf & ArrayNameCodeRight(2, i) 'номер пикета
                        ind += 1
                    Next i
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'рисуем уклоны
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If FormMarksOnTopCoating.CheckBox4.Checked = True And CountDistCells > 0 Then
                PosStartUklRow = PosEnd + 1
                PosEnd = PosStartUklRow + CountDistCells - 1
                xlsheet.Cells(2, PosStartUklRow) = "'Уклон, ‰" 'номер пикета
                cell1 = xlsheet.Cells(1, PosStartUklRow) 'ссылка на первую ячейку
                cell2 = xlsheet.Cells(2, PosEnd) 'ссылка на 6 ячейку в строке
                CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                '=================================================================================
                'формируем лево и право
                'вторая ячейка
                If LenArrayLeft > 0 Then
                    xlsheet.Cells(3, PosStartUklRow) = "'Левая сторона" 'номер пикета
                    cell1 = xlsheet.Cells(3, PosStartUklRow) 'ссылка на первую ячейку
                    cell2 = xlsheet.Cells(3, PosStartUklRow + LenArrayLeft - 1) 'ссылка на 6 ячейку в строке
                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                    CellRng.Merge(Type.Missing)
                End If
                If LenArrayRight > 0 Then
                    xlsheet.Cells(3, PosStartUklRow + LenArrayLeft) = "'Правая сторона" 'номер пикета
                    cell1 = xlsheet.Cells(3, PosStartUklRow + LenArrayLeft) 'ссылка на первую ячейку
                    cell2 = xlsheet.Cells(3, PosEnd) 'ссылка на 6 ячейку в строке
                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                    CellRng.Merge(Type.Missing)
                End If
                '=======================================================================================
                'формируем коды
                Dim ind As Integer = 0
                If LenArrayLeft > 0 Then
                    For i As Integer = ArrayNameCodeLeft.GetUpperBound(1) To 0 Step -1
                        xlsheet.Cells(4, PosStartUklRow + ind) = "'T" & ArrayNameCodeLeft(0, i) & vbLf & ArrayNameCodeLeft(2, i) 'номер пикета
                        ind += 1
                    Next i
                End If
                If LenArrayRight > 0 Then
                    For i As Integer = 0 To ArrayNameCodeRight.GetUpperBound(1)
                        xlsheet.Cells(4, PosStartUklRow + ind) = "'T" & ArrayNameCodeRight(0, i) & vbLf & ArrayNameCodeRight(2, i) 'номер пикета
                        ind += 1
                    Next i
                End If
            End If

            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'рисуем координаты
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim PosEndCoordRow As Integer = 0
            If FormMarksOnTopCoating.CheckBox5.Checked = True Then
                PosStartCoordRow = PosEnd + 1  'начало отметок
                PosEnd = PosStartCoordRow + CountDistCells * 2 + 1 'конечная отметка
                '======================================================================================
                xlsheet.Cells(1, PosStartCoordRow) = "'Координаты, м." 'координаты
                cell1 = xlsheet.Cells(1, PosStartCoordRow) 'ссылка на первую ячейку
                cell2 = xlsheet.Cells(1, PosEnd) 'ссылка на 6 ячейку в строке
                CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                '=========================================================================================
                If LenArrayLeft > 0 Then
                    xlsheet.Cells(2, PosStartCoordRow) = "'Левая сторона" 'номер пикета
                    cell1 = xlsheet.Cells(2, PosStartCoordRow) 'ссылка на первую ячейку
                    cell2 = xlsheet.Cells(2, PosStartCoordRow + LenArrayLeft * 2 - 1) 'ссылка на 6 ячейку в строке
                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                    CellRng.Merge(Type.Missing)
                End If
                '=========================================================================================
                xlsheet.Cells(2, PosStartCoordRow + LenArrayLeft * 2) = "'Ось" 'номер пикета
                cell1 = xlsheet.Cells(2, PosStartCoordRow + LenArrayLeft * 2) 'ссылка на первую ячейку
                cell2 = xlsheet.Cells(3, PosStartCoordRow + LenArrayLeft * 2 + 1) 'ссылка на 6 ячейку в строке
                CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                xlsheet.Cells(4, PosStartCoordRow + LenArrayLeft * 2) = "'X" 'номер пикета
                xlsheet.Cells(4, PosStartCoordRow + LenArrayLeft * 2 + 1) = "'Y" 'номер пикета
                '=========================================================================================
                If LenArrayRight > 0 Then
                    xlsheet.Cells(2, PosStartCoordRow + LenArrayLeft * 2 + 2) = "'Правая сторона" 'номер пикета
                    cell1 = xlsheet.Cells(2, PosStartCoordRow + LenArrayLeft * 2 + 2) 'ссылка на первую ячейку
                    cell2 = xlsheet.Cells(2, PosEnd) 'ссылка на 6 ячейку в строке
                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                    CellRng.Merge(Type.Missing)
                End If
                '=======================================================================================
                'формируем коды
                Dim ind As Integer = 0
                If LenArrayLeft > 0 Then
                    For i As Integer = ArrayNameCodeLeft.GetUpperBound(1) To 0 Step -1
                        xlsheet.Cells(3, PosStartCoordRow + ind) = "'T" & ArrayNameCodeLeft(0, i) & vbLf & ArrayNameCodeLeft(2, i) 'номер пикета
                        xlsheet.Cells(4, PosStartCoordRow + ind) = "'X"
                        xlsheet.Cells(4, PosStartCoordRow + ind + 1) = "'Y"
                        cell1 = xlsheet.Cells(3, PosStartCoordRow + ind) 'ссылка на первую ячейку
                        cell2 = xlsheet.Cells(3, PosStartCoordRow + ind + 1) 'ссылка на 6 ячейку в строке
                        CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                        CellRng.Merge(Type.Missing)
                        ind += 2
                    Next i
                End If
                ind = ind + 2
                If LenArrayRight > 0 Then
                    For i As Integer = 0 To ArrayNameCodeRight.GetUpperBound(1)
                        xlsheet.Cells(3, PosStartCoordRow + ind) = "'T" & ArrayNameCodeRight(0, i) & vbLf & ArrayNameCodeRight(2, i) 'номер пикета
                        xlsheet.Cells(4, PosStartCoordRow + ind) = "'X"
                        xlsheet.Cells(4, PosStartCoordRow + ind + 1) = "'Y"
                        cell1 = xlsheet.Cells(3, PosStartCoordRow + ind) 'ссылка на первую ячейку
                        cell2 = xlsheet.Cells(3, PosStartCoordRow + ind + 1) 'ссылка на 6 ячейку в строке
                        CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                        CellRng.Merge(Type.Missing)
                        ind += 2
                    Next i
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'примечание
            If FormMarksOnTopCoating.CheckBox6.Checked = True Then
                PosStartNoteRow = PosEnd + 1
                xlsheet.Cells(1, PosStartNoteRow) = "'Примечание" 'номер пикета
                cell1 = xlsheet.Cells(1, PosStartNoteRow) 'ссылка на первую ячейку
                cell2 = xlsheet.Cells(4, PosStartNoteRow) 'ссылка на 6 ячейку в строке
                CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ЗАПОЛНЯЕМ ТАБЛИЦУ
            Dim CountRowsExcel As Integer = 5 'начальное значение счетчика
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ОТЧЕТ ПО ПИКЕТАМ
            If FormMarksOnTopCoating.RadioButton5.Checked = True Or FormMarksOnTopCoating.RadioButton1.Checked = True Then
                '=========================================================================================================
                If IsArray(ArrayPKResult) = False Then
                    GoTo Line1
                End If
                For i As Integer = 0 To ArrayPKResult.GetUpperBound(1)
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'формируем пикеты и примечания
                    Dim PK As Double = ArrayPKResult(0, i) 'пикет точки
                    Dim KodPk1 As Integer = ArrayPKResult(1, i) 'конец элемента
                    Dim KodPk2 As Integer = ArrayPKResult(2, i) 'начало элемента
                    Dim KodPk3 As Integer = ArrayPKResult(3, i) 'рубленный пикет
                    Dim Desk1 As String = ""
                    Dim Desk2 As String = ""
                    Dim Desk3 As String = ""
                    Dim Desk As String = ""
                    If KodPk2 = 1 Then
                        Desk1 = "Начало прямой"
                    End If
                    If KodPk1 = 2 Then
                        Desk2 = "Конец прямой"
                    End If
                    If KodPk2 = 3 Then
                        Desk1 = "Начало круговой кривой"
                    End If
                    If KodPk1 = 4 Then
                        Desk2 = "Конец круговой кривой"
                    End If
                    If KodPk2 = 5 Then
                        Desk1 = "Начало переходной кривой"
                    End If
                    If KodPk1 = 6 Then
                        Desk2 = "Конец переходной кривой"
                    End If
                    If KodPk3 = 7 Then
                        Desk3 = "Рубленный пикет"
                    End If
                    If Desk2.Length > 0 Then
                        Desk = Desk2
                    End If
                    If Desk1.Length > 0 Then
                        If Desk.Length > 0 Then
                            Desk = Desk & vbLf & Desk1
                        Else
                            Desk = Desk1
                        End If
                    End If
                    If Desk3.Length > 0 Then
                        If Desk.Length > 0 Then
                            Desk = Desk & vbLf & Desk3
                        Else
                            Desk = Desk3
                        End If
                    End If
                    'переводим в форматированный пикет
                    Dim PKFormat As String = ""
                    If stationEqw.Count > 0 Then
                        PKFormat = FuncAlignment.FuncFormatRublPiket(AlignObj, PK, RoundZnPK)
                    Else
                        Try
                            PKFormat = AlignObj.GetStationStringWithEquations(PK)
                        Catch ex As System.ArgumentException
                            PKFormat = Math.Round(PK, RoundZnPK)
                        End Try
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'записываем пикет
                    xlsheet.Cells(CountRowsExcel, PosStartPKRow).Value = PKFormat
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '============================================================================================
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'получаем конструкцию
                    Dim ObjAplAssembl As AppliedAssembly = Nothing
                    Try
                        ObjAplAssembl = bLine.GetAppliedAssemblyAtStation(PK)
                    Catch ex As System.ArgumentOutOfRangeException
                    Catch ex As System.ArgumentException
                    End Try
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'в случае неудачи ищем ближайшую конструкцию
                    If IsNothing(ObjAplAssembl) = True Then
                        Dim MaxDist As Double = 100000000
                        For Each oBaselineRegion As BaselineRegion In reg
                            If oBaselineRegion.NeedsProcessing = True Then
                                For Each oAppliedAssembly As AppliedAssembly In oBaselineRegion.AppliedAssemblies
                                    'определяем коды для точки для поиска пикетного положения конструкции
                                    For i1 As Integer = 0 To ArrayNameCodeLeft.GetUpperBound(1)
                                        Dim NameCod As String = ArrayNameCodeLeft(2, i1) 'код точки
                                        Dim PointCalc1 As CalculatedPointCollection = oAppliedAssembly.GetPointsByCode(NameCod) 'коды лево
                                        If PointCalc1.Count > 0 Then
                                            Dim PKConstr As Double = PointCalc1.Item(0).StationOffsetElevationToBaseline.X
                                            If Math.Abs((PKConstr - PK)) < MaxDist Then
                                                MaxDist = Math.Abs((PKConstr - PK))
                                                ObjAplAssembl = oAppliedAssembly
                                                If MaxDist < 0.2 Then
                                                    GoTo Line3
                                                End If
                                            End If
                                        End If
                                    Next i1
                                    If IsNothing(ObjAplAssembl) = True Then
                                        For i1 As Integer = 0 To ArrayNameCodeRight.GetUpperBound(1)
                                            Dim NameCod As String = ArrayNameCodeRight(2, i1) 'код точки
                                            Dim PointCalc1 As CalculatedPointCollection = oAppliedAssembly.GetPointsByCode(NameCod) 'коды право
                                            If PointCalc1.Count > 0 Then
                                                Dim PKConstr As Double = PointCalc1.Item(0).StationOffsetElevationToBaseline.X
                                                If Math.Abs((PKConstr - PK)) < MaxDist Then
                                                    MaxDist = Math.Abs((PKConstr - PK))
                                                    ObjAplAssembl = oAppliedAssembly
                                                    If MaxDist < 0.2 Then
                                                        GoTo Line3
                                                    End If
                                                End If
                                            End If
                                        Next i1
                                    End If
                                Next
                            End If
                        Next
                    End If
Line3:
                    '=====================================================================================================
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'начинаем разбор конструкции
                    If IsNothing(ObjAplAssembl) = False Then
                        Dim appSubColl As AppliedSubassemblyCollection = ObjAplAssembl.GetAppliedSubassemblies 'находим подконструкции
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'ищем код осевой точки (если конечно это указал пользователь)
                        Dim PtXYZCentre As Point3d = Nothing
                        Dim CodCentre As String = FormMarksOnTopCoating.ComboBox2.Text
                        Dim flagCentre As Boolean = False
                        If appSubColl.Count > 0 Then
                            For i1 As Integer = 0 To appSubColl.Count - 1
                                Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i1)
                                Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                'ищем этот код в данной подконструкции
                                For i2 As Integer = 0 To pointCol.Count - 1
                                    Dim ptAssembl As CalculatedPoint = pointCol.Item(i2)
                                    'если код осевой линии не задан принимаем за осевую точку верх конструкции
                                    If FormMarksOnTopCoating.CheckBox1.Checked = False Then
                                        Dim CPTXYZ2Temp As Point3d = ptAssembl.StationOffsetElevationToBaseline 'координаты относительно базовой линии (пикет и смещение)
                                        If CPTXYZ2Temp.Y = 0 And CPTXYZ2Temp.Z = 0 Then
                                            PtXYZCentre = blineSubAssembl.StationOffsetElevationToXYZ(CPTXYZ2Temp)
                                            GoTo line4
                                        End If
                                    Else 'если центр конструкции задан явно
                                        If ptAssembl.CorridorCodes.Count > 0 Then
                                            Dim strCode As String = ptAssembl.CorridorCodes.Item(0)
                                            'если коды совпадают
                                            If strCode Like CodCentre Then
                                                Dim CPTXYZ2Temp As Point3d = ptAssembl.StationOffsetElevationToBaseline 'координаты относительно базовой линии (пикет и смещение)
                                                PtXYZCentre = blineSubAssembl.StationOffsetElevationToXYZ(CPTXYZ2Temp)
                                                GoTo line4
                                            End If
                                        End If
                                    End If
                                Next i2
                            Next i1
                        End If

line4:
                        'ищем коды и вычисляем данные
                        'формируем коды
                        Dim ind As Integer = 0
                        If IsArray(ArrayNameCodeLeft) = True Then
                            Dim prevXYZCentre As Point3d = PtXYZCentre
                            Dim indCoordLeft As Integer = ArrayNameCodeLeft.GetUpperBound(1) + 1
                            For i1 As Integer = 0 To ArrayNameCodeLeft.GetUpperBound(1)
                                Dim NameCod As String = ArrayNameCodeLeft(2, i1) 'код точки
                                Dim PointCalc1 As CalculatedPointCollection = ObjAplAssembl.GetPointsByCode(NameCod) 'коды лево
                                'если коллекция кодов и коллекция подконструкций существует
                                If PointCalc1.Count > 0 And appSubColl.Count > 0 Then
                                    For i2 As Integer = 0 To appSubColl.Count - 1
                                        Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i2)
                                        Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                        Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                        'ищем этот код в данной подконструкции
                                        For i3 As Integer = 0 To pointCol.Count - 1
                                            Dim ptAssembl As CalculatedPoint = pointCol.Item(i3)
                                            If ptAssembl.CorridorCodes.Count > 0 Then
                                                Dim strCode As String = ptAssembl.CorridorCodes.Item(0)
                                                'если коды совпадают
                                                If strCode Like NameCod Then
                                                    Dim CPTXYZ2Temp As Point3d = ptAssembl.StationOffsetElevationToBaseline 'координаты относительно базовой линии (пикет и смещение)
                                                    Dim PtAbsAssembly1 As Point3d = blineSubAssembl.StationOffsetElevationToXYZ(CPTXYZ2Temp) 'переводим их в абсолютные значения
                                                    If CPTXYZ2Temp.Y < 0 Then
                                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                        'записываем расстояния
                                                        If FormMarksOnTopCoating.CheckBox2.Checked = True Then
                                                            If FormMarksOnTopCoating.RadioButton4.Checked = True Then
                                                                Dim v3d As Vector3d = PtXYZCentre.GetVectorTo(PtAbsAssembly1)
                                                                Dim v2d As Vector2d = v3d.Convert2d(New Plane())
                                                                Dim Dist As Double = v2d.Length
                                                                xlsheet.Cells(CountRowsExcel, PosStartDistRow + ArrayNameCodeLeft.GetUpperBound(1) - i1) = Math.Round(Dist, RoundZnDist, MidpointRounding.AwayFromZero)
                                                            Else
                                                                'записываем наклонное расстояние
                                                                Dim Dist As Double = PtXYZCentre.GetVectorTo(PtAbsAssembly1).Length
                                                                xlsheet.Cells(CountRowsExcel, PosStartDistRow + ArrayNameCodeLeft.GetUpperBound(1) - i1) = Math.Round(Dist, RoundZnDist, MidpointRounding.AwayFromZero)
                                                            End If
                                                        End If
                                                        'записываем уклоны
                                                        If FormMarksOnTopCoating.CheckBox4.Checked = True Then
                                                            Dim Deltah As Double = PtAbsAssembly1.Z - prevXYZCentre.Z
                                                            Dim pt1 As Point2d = New Point2d(prevXYZCentre.X, prevXYZCentre.Y)
                                                            Dim pt2 As Point2d = New Point2d(PtAbsAssembly1.X, PtAbsAssembly1.Y)
                                                            'Dim Dist As Double = Math.Abs(CPTXYZ2Temp.Y)
                                                            Dim DistGor As Double = pt1.GetVectorTo(pt2).Length
                                                            Dim Ukl As Double = (Deltah / DistGor) * 1000
                                                            xlsheet.Cells(CountRowsExcel, PosStartUklRow + ArrayNameCodeLeft.GetUpperBound(1) - i1) = Math.Round(Ukl, RoundZnSlope, MidpointRounding.AwayFromZero)
                                                        End If
                                                        'записываем отметки
                                                        If FormMarksOnTopCoating.CheckBox3.Checked = True Then
                                                            xlsheet.Cells(CountRowsExcel, PosStartOtmRow + ArrayNameCodeLeft.GetUpperBound(1) - i1) = Math.Round(PtAbsAssembly1.Z, RoundZnElev, MidpointRounding.AwayFromZero)
                                                        End If
                                                        'записываем координаты
                                                        If FormMarksOnTopCoating.CheckBox5.Checked = True Then
                                                            Dim indCoordLeft2 As Integer = PosStartCoordRow + indCoordLeft * 2 - 1
                                                            xlsheet.Cells(CountRowsExcel, indCoordLeft2 - 1) = Math.Round(PtAbsAssembly1.Y, RoundZnCoord, MidpointRounding.AwayFromZero)
                                                            xlsheet.Cells(CountRowsExcel, indCoordLeft2) = Math.Round(PtAbsAssembly1.X, RoundZnCoord, MidpointRounding.AwayFromZero)
                                                        End If
                                                        prevXYZCentre = PtAbsAssembly1
                                                        indCoordLeft -= 1
                                                        GoTo Line5
                                                    End If
                                                End If
                                            End If
                                        Next i3
                                    Next i2
                                End If
Line5:
                                ind += 1
                            Next i1
                        End If
                        '==============================================================================================
                        'записываем данные по оси
                        If FormMarksOnTopCoating.CheckBox3.Checked = True Then
                            xlsheet.Cells(CountRowsExcel, PosStartOtmRow + ind) = Math.Round(PtXYZCentre.Z, RoundZnElev, MidpointRounding.AwayFromZero)
                        End If
                        'записываем координаты по оси
                        If FormMarksOnTopCoating.CheckBox5.Checked = True Then
                            xlsheet.Cells(CountRowsExcel, PosStartCoordRow + ind * 2) = Math.Round(PtXYZCentre.Y, RoundZnCoord, MidpointRounding.AwayFromZero)
                            xlsheet.Cells(CountRowsExcel, PosStartCoordRow + ind * 2 + 1) = Math.Round(PtXYZCentre.X, RoundZnCoord, MidpointRounding.AwayFromZero)
                        End If
                        '==============================================================================================
                        'записываем правую сторону
                        If IsArray(ArrayNameCodeRight) = True Then
                            Dim prevXYZCentre1 As Point3d = PtXYZCentre
                            For i1 As Integer = 0 To ArrayNameCodeRight.GetUpperBound(1)
                                Dim NameCod As String = ArrayNameCodeRight(2, i1) 'код точки
                                Dim PointCalc1 As CalculatedPointCollection = ObjAplAssembl.GetPointsByCode(NameCod) 'коды лево
                                'если коллекция кодов и коллекция подконструкций существует
                                If PointCalc1.Count > 0 And appSubColl.Count > 0 Then
                                    For i2 As Integer = 0 To appSubColl.Count - 1
                                        Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i2)
                                        Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                        Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                        'ищем этот код в данной подконструкции
                                        For i3 As Integer = 0 To pointCol.Count - 1
                                            Dim ptAssembl As CalculatedPoint = pointCol.Item(i3)
                                            If ptAssembl.CorridorCodes.Count > 0 Then
                                                Dim strCode As String = ptAssembl.CorridorCodes.Item(0)
                                                'если коды совпадают
                                                If strCode Like NameCod Then
                                                    Dim CPTXYZ2Temp As Point3d = ptAssembl.StationOffsetElevationToBaseline 'координаты относительно базовой линии (пикет и смещение)
                                                    Dim PtAbsAssembly1 As Point3d = blineSubAssembl.StationOffsetElevationToXYZ(CPTXYZ2Temp) 'переводим их в абсолютные значения
                                                    If CPTXYZ2Temp.Y > 0 Then
                                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                        'записываем расстояния
                                                        If FormMarksOnTopCoating.CheckBox2.Checked = True Then
                                                            If FormMarksOnTopCoating.RadioButton4.Checked = True Then
                                                                Dim v3d As Vector3d = PtXYZCentre.GetVectorTo(PtAbsAssembly1)
                                                                Dim v2d As Vector2d = v3d.Convert2d(New Plane())
                                                                Dim Dist As Double = v2d.Length
                                                                xlsheet.Cells(CountRowsExcel, PosStartDistRow + ind) = Math.Round(Dist, RoundZnDist, MidpointRounding.AwayFromZero)
                                                            Else
                                                                'записываем наклонное расстояние
                                                                Dim Dist As Double = PtXYZCentre.GetVectorTo(PtAbsAssembly1).Length
                                                                xlsheet.Cells(CountRowsExcel, PosStartDistRow + ind) = Math.Round(Dist, RoundZnDist, MidpointRounding.AwayFromZero)
                                                            End If
                                                        End If
                                                        'записываем уклоны
                                                        If FormMarksOnTopCoating.CheckBox4.Checked = True Then
                                                            Dim Deltah As Double = PtAbsAssembly1.Z - prevXYZCentre1.Z
                                                            Dim pt1 As Point2d = New Point2d(prevXYZCentre1.X, prevXYZCentre1.Y)
                                                            Dim pt2 As Point2d = New Point2d(PtAbsAssembly1.X, PtAbsAssembly1.Y)
                                                            Dim DistGor As Double = pt1.GetVectorTo(pt2).Length
                                                            Dim Ukl As Double = (Deltah / DistGor) * 1000
                                                            xlsheet.Cells(CountRowsExcel, PosStartUklRow + ind) = Math.Round(Ukl, RoundZnSlope, MidpointRounding.AwayFromZero)
                                                        End If
                                                        'записываем отметки
                                                        If FormMarksOnTopCoating.CheckBox3.Checked = True Then
                                                            xlsheet.Cells(CountRowsExcel, PosStartOtmRow + ind + 1) = Math.Round(PtAbsAssembly1.Z, RoundZnElev, MidpointRounding.AwayFromZero)
                                                        End If
                                                        'записываем координаты
                                                        If FormMarksOnTopCoating.CheckBox5.Checked = True Then
                                                            xlsheet.Cells(CountRowsExcel, PosStartCoordRow + ind * 2 + 2) = Math.Round(PtAbsAssembly1.Y, RoundZnCoord, MidpointRounding.AwayFromZero)
                                                            xlsheet.Cells(CountRowsExcel, PosStartCoordRow + ind * 2 + 3) = Math.Round(PtAbsAssembly1.X, RoundZnCoord, MidpointRounding.AwayFromZero)
                                                        End If
                                                        prevXYZCentre1 = PtAbsAssembly1
                                                        GoTo line10
                                                    End If
                                                End If
                                            End If

                                        Next i3
                                    Next i2
                                End If
line10:
                                ind += 1
                            Next i1
                        End If
                    End If
                    '==================================================================================================
                    'записываем примечания
                    If FormMarksOnTopCoating.CheckBox6.Checked = True Then
                        xlsheet.Cells(CountRowsExcel, PosStartNoteRow) = Desk
                    End If
                    CountRowsExcel += 1
Line2:
                Next i
            Else
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ПО КОНСТРУКЦИЯМ
                For Each BlineRegion As BaselineRegion In reg
                    Dim PKStart As Double = BlineRegion.StartStation 'пикет начала
                    Dim PKEnd As Double = BlineRegion.EndStation 'пикет конца
                    'если регион находится в диапазоне
                    If PKStart >= PKStartAlign And PKEnd <= PKEndAlign Then
                        'если регион рабочий
                        If BlineRegion.NeedsProcessing = True Then
                            'список пикетов
                            Dim ArraySt As Double() = BlineRegion.SortedStations
                            For i As Integer = 0 To ArraySt.Length - 1
                                Dim PK As Double = ArraySt(i)
                                'переводим в форматированный пикет
                                Dim PKFormat As String = ""
                                If stationEqw.Count > 0 Then
                                    PKFormat = FuncAlignment.FuncFormatRublPiket(AlignObj, PK, RoundZnPK, MidpointRounding.AwayFromZero)
                                Else
                                    Try
                                        PKFormat = AlignObj.GetStationStringWithEquations(PK)
                                    Catch ex As System.ArgumentException
                                        PKFormat = Math.Round(PK, RoundZnPK, MidpointRounding.AwayFromZero)
                                    End Try
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'записываем пикет
                                xlsheet.Cells(CountRowsExcel, PosStartPKRow).Value = PKFormat
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'получаем конструкцию по пикету
                                Dim ObjAplAssembl As AppliedAssembly = Nothing
                                If IsArray(ArrayNameCodeRight) = True Or IsArray(ArrayNameCodeLeft) = True Then
                                    Try
                                        ObjAplAssembl = bLine.GetAppliedAssemblyAtStation(PK)
                                    Catch ex As System.ArgumentOutOfRangeException
                                    Catch ex As System.ArgumentException
                                    End Try
                                End If
                                If IsNothing(ObjAplAssembl) = False Then
                                    Dim appSubColl As AppliedSubassemblyCollection = ObjAplAssembl.GetAppliedSubassemblies 'находим подконструкции
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'ищем код осевой точки (если конечно это указал пользователь)
                                    Dim PtXYZCentre As Point3d = Nothing
                                    Dim CodCentre As String = FormMarksOnTopCoating.ComboBox2.Text
                                    Dim flagCentre As Boolean = False
                                    If appSubColl.Count > 0 Then
                                        For i1 As Integer = 0 To appSubColl.Count - 1
                                            Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i1)
                                            Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                            Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                            'ищем этот код в данной подконструкции
                                            For i2 As Integer = 0 To pointCol.Count - 1
                                                Dim ptAssembl As CalculatedPoint = pointCol.Item(i2)
                                                'если код осевой линии не задан принимаем за осевую точку верх конструкции
                                                If FormMarksOnTopCoating.CheckBox1.Checked = False Then
                                                    Dim CPTXYZ2Temp As Point3d = ptAssembl.StationOffsetElevationToBaseline 'координаты относительно базовой линии (пикет и смещение)
                                                    If CPTXYZ2Temp.Y = 0 And CPTXYZ2Temp.Z = 0 Then
                                                        PtXYZCentre = blineSubAssembl.StationOffsetElevationToXYZ(CPTXYZ2Temp)
                                                        GoTo Line6
                                                    End If
                                                Else 'если центр конструкции задан явно
                                                    If ptAssembl.CorridorCodes.Count > 0 Then
                                                        Dim strCode As String = ptAssembl.CorridorCodes.Item(0)
                                                        'если коды совпадают
                                                        If strCode Like CodCentre Then
                                                            Dim CPTXYZ2Temp As Point3d = ptAssembl.StationOffsetElevationToBaseline 'координаты относительно базовой линии (пикет и смещение)
                                                            PtXYZCentre = blineSubAssembl.StationOffsetElevationToXYZ(CPTXYZ2Temp)
                                                            GoTo Line6
                                                        End If
                                                    End If
                                                End If
                                            Next i2
                                        Next i1
                                    End If
Line6:
                                    'формируем кода ЛЕВО
                                    Dim ind As Integer = 0
                                    If IsArray(ArrayNameCodeLeft) = True Then
                                        Dim prevXYZCentre As Point3d = PtXYZCentre
                                        Dim indCoordLeft As Integer = ArrayNameCodeLeft.GetUpperBound(1) + 1
                                        For i1 As Integer = 0 To ArrayNameCodeLeft.GetUpperBound(1)
                                            Dim NameCod As String = ArrayNameCodeLeft(2, i1) 'код точки
                                            Dim PointCalc1 As CalculatedPointCollection = ObjAplAssembl.GetPointsByCode(NameCod) 'коды лево
                                            'если коллекция кодов и коллекция подконструкций существует
                                            If PointCalc1.Count > 0 And appSubColl.Count > 0 Then
                                                For i2 As Integer = 0 To appSubColl.Count - 1
                                                    Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i2)
                                                    Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                                    Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                                    'ищем этот код в данной подконструкции
                                                    For i3 As Integer = 0 To pointCol.Count - 1
                                                        Dim ptAssembl As CalculatedPoint = pointCol.Item(i3)
                                                        If ptAssembl.CorridorCodes.Count > 0 Then
                                                            Dim strCode As String = ptAssembl.CorridorCodes.Item(0)
                                                            'если коды совпадают
                                                            If strCode Like NameCod Then
                                                                Dim CPTXYZ2Temp As Point3d = ptAssembl.StationOffsetElevationToBaseline 'координаты относительно базовой линии (пикет и смещение)
                                                                Dim PtAbsAssembly1 As Point3d = blineSubAssembl.StationOffsetElevationToXYZ(CPTXYZ2Temp) 'переводим их в абсолютные значения
                                                                If CPTXYZ2Temp.Y < 0 Then
                                                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                                    'записываем расстояния
                                                                    If FormMarksOnTopCoating.CheckBox2.Checked = True Then
                                                                        If FormMarksOnTopCoating.RadioButton4.Checked = True Then
                                                                            Dim v3d As Vector3d = PtXYZCentre.GetVectorTo(PtAbsAssembly1)
                                                                            Dim v2d As Vector2d = v3d.Convert2d(New Plane())
                                                                            Dim Dist As Double = v2d.Length
                                                                            xlsheet.Cells(CountRowsExcel, PosStartDistRow + ArrayNameCodeLeft.GetUpperBound(1) - i1) = Math.Round(Dist, RoundZnDist, MidpointRounding.AwayFromZero)
                                                                        Else
                                                                            'записываем наклонное расстояние
                                                                            Dim Dist As Double = PtXYZCentre.GetVectorTo(PtAbsAssembly1).Length
                                                                            xlsheet.Cells(CountRowsExcel, PosStartDistRow + ArrayNameCodeLeft.GetUpperBound(1) - i1) = Math.Round(Dist, RoundZnDist, MidpointRounding.AwayFromZero)
                                                                        End If
                                                                    End If
                                                                    'записываем уклоны
                                                                    If FormMarksOnTopCoating.CheckBox4.Checked = True Then
                                                                        Dim Deltah As Double = PtAbsAssembly1.Z - prevXYZCentre.Z
                                                                        Dim pt1 As Point2d = New Point2d(prevXYZCentre.X, prevXYZCentre.Y)
                                                                        Dim pt2 As Point2d = New Point2d(PtAbsAssembly1.X, PtAbsAssembly1.Y)
                                                                        'Dim Dist As Double = Math.Abs(CPTXYZ2Temp.Y)
                                                                        Dim DistGor As Double = pt1.GetVectorTo(pt2).Length
                                                                        Dim Ukl As Double = (Deltah / DistGor) * 1000
                                                                        xlsheet.Cells(CountRowsExcel, PosStartUklRow + ArrayNameCodeLeft.GetUpperBound(1) - i1) = Math.Round(Ukl, RoundZnSlope, MidpointRounding.AwayFromZero)
                                                                    End If
                                                                    'записываем отметки
                                                                    If FormMarksOnTopCoating.CheckBox3.Checked = True Then
                                                                        xlsheet.Cells(CountRowsExcel, PosStartOtmRow + ArrayNameCodeLeft.GetUpperBound(1) - i1) = Math.Round(PtAbsAssembly1.Z, RoundZnElev, MidpointRounding.AwayFromZero)
                                                                    End If
                                                                    'записываем координаты
                                                                    If FormMarksOnTopCoating.CheckBox5.Checked = True Then
                                                                        Dim indCoordLeft2 As Integer = PosStartCoordRow + indCoordLeft * 2 - 1
                                                                        xlsheet.Cells(CountRowsExcel, indCoordLeft2 - 1) = Math.Round(PtAbsAssembly1.Y, RoundZnCoord, MidpointRounding.AwayFromZero)
                                                                        xlsheet.Cells(CountRowsExcel, indCoordLeft2) = Math.Round(PtAbsAssembly1.X, RoundZnCoord, MidpointRounding.AwayFromZero)
                                                                    End If
                                                                    prevXYZCentre = PtAbsAssembly1
                                                                    indCoordLeft -= 1
                                                                    GoTo Line7
                                                                End If
                                                            End If
                                                        End If
                                                    Next i3
                                                Next i2
                                            End If
Line7:
                                            ind += 1
                                        Next i1
                                    End If
                                    '==============================================================================================
                                    'записываем данные по оси
                                    If FormMarksOnTopCoating.CheckBox3.Checked = True Then
                                        xlsheet.Cells(CountRowsExcel, PosStartOtmRow + ind) = Math.Round(PtXYZCentre.Z, RoundZnElev, MidpointRounding.AwayFromZero)
                                    End If
                                    'записываем координаты по оси
                                    If FormMarksOnTopCoating.CheckBox5.Checked = True Then
                                        xlsheet.Cells(CountRowsExcel, PosStartCoordRow + ind * 2) = Math.Round(PtXYZCentre.Y, RoundZnCoord, MidpointRounding.AwayFromZero)
                                        xlsheet.Cells(CountRowsExcel, PosStartCoordRow + ind * 2 + 1) = Math.Round(PtXYZCentre.X, RoundZnCoord, MidpointRounding.AwayFromZero)
                                    End If
                                    '==============================================================================================
                                    'записываем правую сторону
                                    If IsArray(ArrayNameCodeRight) = True Then
                                        Dim prevXYZCentre1 As Point3d = PtXYZCentre
                                        For i1 As Integer = 0 To ArrayNameCodeRight.GetUpperBound(1)
                                            Dim NameCod As String = ArrayNameCodeRight(2, i1) 'код точки
                                            Dim PointCalc1 As CalculatedPointCollection = ObjAplAssembl.GetPointsByCode(NameCod) 'коды лево
                                            'если коллекция кодов и коллекция подконструкций существует
                                            If PointCalc1.Count > 0 And appSubColl.Count > 0 Then
                                                For i2 As Integer = 0 To appSubColl.Count - 1
                                                    Dim SubAssembl As AppliedSubassembly = appSubColl.Item(i2)
                                                    Dim blineSubAssembl As BaseBaseline = SubAssembl.BaselineHookedTo
                                                    Dim pointCol As CalculatedPointCollection = SubAssembl.Points
                                                    'ищем этот код в данной подконструкции
                                                    For i3 As Integer = 0 To pointCol.Count - 1
                                                        Dim ptAssembl As CalculatedPoint = pointCol.Item(i3)
                                                        If ptAssembl.CorridorCodes.Count > 0 Then
                                                            Dim strCode As String = ptAssembl.CorridorCodes.Item(0)
                                                            'если коды совпадают
                                                            If strCode Like NameCod Then
                                                                Dim CPTXYZ2Temp As Point3d = ptAssembl.StationOffsetElevationToBaseline 'координаты относительно базовой линии (пикет и смещение)
                                                                Dim PtAbsAssembly1 As Point3d = blineSubAssembl.StationOffsetElevationToXYZ(CPTXYZ2Temp) 'переводим их в абсолютные значения
                                                                If CPTXYZ2Temp.Y > 0 Then
                                                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                                    'записываем расстояния
                                                                    If FormMarksOnTopCoating.CheckBox2.Checked = True Then
                                                                        If FormMarksOnTopCoating.RadioButton4.Checked = True Then
                                                                            Dim v3d As Vector3d = PtXYZCentre.GetVectorTo(PtAbsAssembly1)
                                                                            Dim v2d As Vector2d = v3d.Convert2d(New Plane())
                                                                            Dim Dist As Double = v2d.Length
                                                                            xlsheet.Cells(CountRowsExcel, PosStartDistRow + ind) = Math.Round(Dist, RoundZnDist, MidpointRounding.AwayFromZero)
                                                                        Else
                                                                            'записываем наклонное расстояние
                                                                            Dim Dist As Double = PtXYZCentre.GetVectorTo(PtAbsAssembly1).Length
                                                                            xlsheet.Cells(CountRowsExcel, PosStartDistRow + ind) = Math.Round(Dist, RoundZnDist, MidpointRounding.AwayFromZero)
                                                                        End If
                                                                    End If
                                                                    'записываем уклоны
                                                                    If FormMarksOnTopCoating.CheckBox4.Checked = True Then
                                                                        Dim Deltah As Double = PtAbsAssembly1.Z - prevXYZCentre1.Z
                                                                        Dim pt1 As Point2d = New Point2d(prevXYZCentre1.X, prevXYZCentre1.Y)
                                                                        Dim pt2 As Point2d = New Point2d(PtAbsAssembly1.X, PtAbsAssembly1.Y)
                                                                        Dim DistGor As Double = pt1.GetVectorTo(pt2).Length
                                                                        Dim Ukl As Double = (Deltah / DistGor) * 1000
                                                                        xlsheet.Cells(CountRowsExcel, PosStartUklRow + ind) = Math.Round(Ukl, RoundZnSlope, MidpointRounding.AwayFromZero)
                                                                    End If
                                                                    'записываем отметки
                                                                    If FormMarksOnTopCoating.CheckBox3.Checked = True Then
                                                                        xlsheet.Cells(CountRowsExcel, PosStartOtmRow + ind + 1) = Math.Round(PtAbsAssembly1.Z, RoundZnElev, MidpointRounding.AwayFromZero)
                                                                    End If
                                                                    'записываем координаты
                                                                    If FormMarksOnTopCoating.CheckBox5.Checked = True Then
                                                                        xlsheet.Cells(CountRowsExcel, PosStartCoordRow + ind * 2 + 2) = Math.Round(PtAbsAssembly1.Y, RoundZnCoord, MidpointRounding.AwayFromZero)
                                                                        xlsheet.Cells(CountRowsExcel, PosStartCoordRow + ind * 2 + 3) = Math.Round(PtAbsAssembly1.X, RoundZnCoord, MidpointRounding.AwayFromZero)
                                                                    End If
                                                                    prevXYZCentre1 = PtAbsAssembly1
                                                                    GoTo Line9
                                                                End If
                                                            End If
                                                        End If
                                                    Next i3
                                                Next i2
                                            End If
Line9:
                                            ind += 1
                                        Next i1
                                        '==================================================================================================
                                        'записываем примечания
                                        Dim Desk As String = ""
                                        If IsArray(ArrayPKResult) = True Then
                                            Dim Desk1 As String = ""
                                            Dim Desk2 As String = ""
                                            Dim Desk3 As String = ""
                                            For k As Integer = 0 To ArrayPKResult.GetUpperBound(1)
                                                If PK = ArrayPKResult(0, k) Then
                                                    Dim KodPk1 As Integer = ArrayPKResult(1, k) 'конец элемента
                                                    Dim KodPk2 As Integer = ArrayPKResult(2, k) 'начало элемента
                                                    Dim KodPk3 As Integer = ArrayPKResult(3, k) 'рубленный пикет
                                                    If KodPk2 = 1 Then
                                                        Desk1 = "Начало прямой"
                                                    End If
                                                    If KodPk1 = 2 Then
                                                        Desk2 = "Конец прямой"
                                                    End If
                                                    If KodPk2 = 3 Then
                                                        Desk1 = "Начало круговой кривой"
                                                    End If
                                                    If KodPk1 = 4 Then
                                                        Desk2 = "Конец круговой кривой"
                                                    End If
                                                    If KodPk2 = 5 Then
                                                        Desk1 = "Начало переходной кривой"
                                                    End If
                                                    If KodPk1 = 6 Then
                                                        Desk2 = "Конец переходной кривой"
                                                    End If
                                                    If KodPk3 = 7 Then
                                                        Desk3 = "Рубленный пикет"
                                                    End If
                                                End If
                                            Next k
                                            If Desk2.Length > 0 Then
                                                Desk = Desk2
                                            End If
                                            If Desk1.Length > 0 Then
                                                If Desk.Length > 0 Then
                                                    Desk = Desk & vbLf & Desk1
                                                Else
                                                    Desk = Desk1
                                                End If
                                            End If
                                            If Desk3.Length > 0 Then
                                                If Desk.Length > 0 Then
                                                    Desk = Desk & vbLf & Desk3
                                                Else
                                                    Desk = Desk3
                                                End If
                                            End If
                                        End If
                                        If FormMarksOnTopCoating.CheckBox6.Checked = True Then
                                            xlsheet.Cells(CountRowsExcel, PosStartNoteRow) = Desk
                                        End If
                                    End If
                                End If
                                CountRowsExcel += 1
                            Next i
                        End If
                    End If
                Next
                xlApp = Nothing
            End If
Line1:
            If boolSistPK = True Then 'устанавливаем новое значение
                CivilDoc.Settings.DrawingSettings.AmbientSettings.Station.Precision.Value = SistPeremPK
            End If
            tr.Commit()
            locker.Dispose()
            MsgBox("Макро выполнено!")
        End Using
    End Sub
End Class
