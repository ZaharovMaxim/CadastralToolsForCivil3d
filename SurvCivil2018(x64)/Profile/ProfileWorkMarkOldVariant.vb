Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.Globalization
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Public Class ProfileWorkMarkOldVariant
    '=============================================================================================
    'ведомость рабочих отметок
    <CommandMethod("ProjProfileWorkMarks2")>
    Public Sub PROFWORKMARK2()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim objAlign As Autodesk.Civil.DatabaseServices.Alignment = Nothing 'объект типа трасса
        Dim ProfileWiev As ProfileView = Nothing 'вид профиля
        Dim ProfileRed As Profile = Nothing 'красный профиль
        Dim ProfileBlask As Profile = Nothing 'черный профиль
        Dim ProfileDop As Profile = Nothing 'дополнительный профиль
        '=====================================================================================================================
        'запуск диалогового окна
        Dim ProfDialog As ProfileWorkMark = New ProfileWorkMark
        ProfDialog.ShowDialog()
        If ProfDialog.boolProf = False Then Exit Sub
        '=====================================================================================================================
        'начальные установки
        'точность отметок
        Dim RoundZnH As Integer = 2
        'точность по Н
        If ProfDialog.ComboBox5.Text Like "0" Then
            RoundZnH = 0
        ElseIf ProfDialog.ComboBox5.Text Like "0.0" Then
            RoundZnH = 1
        ElseIf ProfDialog.ComboBox5.Text Like "0.00" Then
            RoundZnH = 2
        ElseIf ProfDialog.ComboBox5.Text Like "0.000" Then
            RoundZnH = 3
        ElseIf ProfDialog.ComboBox5.Text Like "0.0000" Then
            RoundZnH = 4
        Else
            RoundZnH = 8
        End If
        'отступ от оси
        Dim OtstupY As Double = ProfDialog.NumericUpDown2.Value
        If ProfDialog.CheckBox7.Checked = True Then
            OtstupY = OtstupY + 4.5
        End If
        'высота текста
        Dim HText As Double = ProfDialog.NumericUpDown3.Value
        'слой объекта
        Dim LayerText As String = ProfDialog.ComboBox6.Text
        Dim IDInsLayer As ObjectId = AcadFunction.FuncAddLayer(LayerText)
        '====================================================================================================================
        'тело программы
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'стиль текста
            Dim StuleText As String = ProfDialog.ComboBox4.Text
            Dim idStyleText As ObjectId = Nothing
            Dim acTextStyleTbl As TextStyleTable = tr.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead)
            For Each id As ObjectId In acTextStyleTbl
                Dim acTextStyleRec As TextStyleTableRecord = tr.GetObject(id, OpenMode.ForRead)
                If acTextStyleRec.Name Like StuleText Then
                    idStyleText = id
                    Exit For
                End If
            Next
            'запрос исходных данных
            'выбираем красный прфиль
            If ProfDialog.RadioButton6.Checked = True Or ProfDialog.RadioButton5.Checked = True Then
                Dim selectRedResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите красный профиль:")
                If selectRedResult.Status = PromptStatus.OK Then
                    Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectRedResult.ObjectId, OpenMode.ForRead)
                    ProfileRed = TryCast(dbObj, Profile)
                Else
                    MsgBox("Ошибка выбора данных. Макро прервано!!!")
                    Exit Sub
                End If
                'выбираем черный прфиль
                Dim selectBlackResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите черный профиль:")
                If selectBlackResult.Status = PromptStatus.OK Then
                    Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectBlackResult.ObjectId, OpenMode.ForRead)
                    ProfileBlask = TryCast(dbObj, Profile)
                Else
                    MsgBox("Ошибка выбора данных. Макро прервано!!!")
                    Exit Sub
                End If
                'выбираем вид профиля
                Dim selectWiewProfileResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите вид профиля:")
                If selectWiewProfileResult.Status = PromptStatus.OK Then
                    Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectWiewProfileResult.ObjectId, OpenMode.ForRead)
                    ProfileWiev = TryCast(dbObj, ProfileView)
                Else
                    MsgBox("Ошибка выбора данных. Макро прервано!!!")
                    Exit Sub
                End If
                ''выбираем дополнительный профиль
                If ProfDialog.CheckBox7.Checked = True Then
                    'выбираем красный прфиль
                    Dim selectRedResult1 As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите дополнительный профиль:")
                    If selectRedResult1.Status = PromptStatus.OK Then
                        Dim dbObj1 As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectRedResult1.ObjectId, OpenMode.ForRead)
                        ProfileDop = TryCast(dbObj1, Profile)
                    Else
                        MsgBox("Ошибка выбора данных. Макро прервано!!!")
                        Exit Sub
                    End If
                End If
                'проверяем что выбрано
                If IsNothing(ProfileBlask) = True Or IsNothing(ProfileRed) = True Or IsNothing(ProfileWiev) = True Then
                    MsgBox("Ошибка выбора данных. Макро прервано!!!")
                    Exit Sub
                End If
                If IsNothing(ProfileDop) = True And ProfDialog.CheckBox7.Checked = True Then
                    MsgBox("Ошибка выбора данных. Не правильно выбран дополнительный профиль. Макро прервано!!!")
                    Exit Sub
                End If
            End If
            '===============================================================================================================================
            'работа с меткой Civil 3d
            If ProfDialog.RadioButton6.Checked = True Then
                Dim ProfileLabel As ProfileDataBandLabelGroup = Nothing
                'выберите метку
                Dim selectLabelResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите метку с рабочими отметками:")
                If selectLabelResult.Status = PromptStatus.OK Then
                    Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectLabelResult.ObjectId, OpenMode.ForWrite)
                    ProfileLabel = TryCast(dbObj, ProfileDataBandLabelGroup)
                End If
                'если данных не хватает
                If IsNothing(ProfileLabel) = True Then
                    MsgBox("Ошибка выбора данных. Макро прервано!!!")
                    Exit Sub
                End If
                'смотрим количество элементов для метки
                Dim CountLabel As Integer = ProfileLabel.SubEntityCount
                'сбрасываем позицию метки
                ProfileLabel.ResetAllSubCommonLabelLocations()
                'ищем габаритный контейнер
                Dim BoundsLabel As Extents3d = ProfileLabel.Bounds
                Dim Dist As Double = BoundsLabel.MinPoint.Y - BoundsLabel.MaxPoint.Y
                Dist = Math.Abs(Dist / 2)
                If CountLabel > 0 Then
                    For i As Integer = 0 To CountLabel - 1
                        Dim lg As LabelGroupSubEntity = ProfileLabel.GetAt(i)
                        'ищем якорь метки (Координаты)
                        Dim PosLocation As Point3d = lg.AnchorInfo.Location
                        Dim PK As Double = -1
                        Dim h As Double = 0
                        ProfileWiev.FindStationAndElevationAtXY(PosLocation.X, PosLocation.Y, PK, h)
                        If PK <> -1 Then
                            Try
                                Dim HRed As Double = ProfileRed.ElevationAt(PK)
                                Dim HBlack As Double = ProfileBlask.ElevationAt(PK)
                                If HRed >= HBlack Then
                                    Dim x As Double = 0
                                    Dim y As Double = 0
                                    ProfileWiev.FindXYAtStationAndElevation(PK, HRed, x, y)
                                    Dim inspoint As Point3d = New Point3d(x, y + Dist + OtstupY, 0)
                                    lg.LabelLocation = inspoint
                                Else
                                    Dim x As Double = 0
                                    Dim y As Double = 0
                                    ProfileWiev.FindXYAtStationAndElevation(PK, HRed, x, y)
                                    Dim inspoint As Point3d = New Point3d(x, y - Dist - OtstupY, 0)
                                    lg.LabelLocation = inspoint
                                End If
                            Catch ex As ArgumentException
                            End Try
                        End If
                    Next i
                End If

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'текст автокад
            ElseIf ProfDialog.RadioButton5.Checked = True Then
                '1. запускаем генератор случайных чисел
                Dim randomObject As New Random() 'формируем индивидуальный id
                Dim randomNumber As Integer = randomObject.Next()
                'массив с пикетажем
                Dim ProfRezult As Profile = Nothing
                Dim ArrayFrofPK As String(,) = Nothing
                Dim CountArrayProfPK As Integer = 0
                'точки вертикальной геометрии красного профиля
                If ProfDialog.CheckBox1.Checked = True Or ProfDialog.CheckBox5.Checked = True Then
                    Dim FindVertGeomRed As Boolean = FuncFrofile.FuncArrayPKVertGeom(ProfileRed, ArrayFrofPK)
                    ProfRezult = ProfileRed
                End If
                'точки вертикальной геометрии черного профиля
                If ProfDialog.CheckBox2.Checked = True Then
                    Dim FindVertGeomRed As Boolean = FuncFrofile.FuncArrayPKVertGeom(ProfileBlask, ArrayFrofPK)
                End If
                'точки вертикальной геометрии дополнительного профиля
                If IsNothing(ProfileDop) = False Then
                    If ProfDialog.CheckBox5.Checked = True Then
                        Dim FindVertGeomRed As Boolean = FuncFrofile.FuncArrayPKVertGeom(ProfileDop, ArrayFrofPK)
                    End If
                End If
                'если есть шаг
                If ProfDialog.CheckBox4.Checked = True Then
                    'определяем старт
                    If IsArray(ArrayFrofPK) = True Then
                        CountArrayProfPK = ArrayFrofPK.Length
                        'получаем начальное значение цикла
                        Dim ChagRound As Double = ProfDialog.NumericUpDown1.Value
                        Dim IntPKStart As Double = ProfRezult.StartingStation + ChagRound
                        IntPKStart = Math.Ceiling(ProfRezult.StartingStation / ChagRound) 'делим на шаг и округляем вверх
                        IntPKStart = IntPKStart * ChagRound
                        If IntPKStart = ProfRezult.StartingStation Then IntPKStart = IntPKStart + ChagRound
                        'вычисляем конечное значение цикла
                        Dim IntPKEnd As Double = ProfRezult.EndingStation 'конечное значение цикла
                        IntPKEnd = Math.Floor(ProfRezult.EndingStation / ChagRound)
                        IntPKEnd = IntPKEnd * ChagRound
                        If IntPKEnd = ProfRezult.EndingStation Then IntPKEnd = IntPKEnd - ChagRound
                        'записываем в массив начальное значение профиля
                        ReDim Preserve ArrayFrofPK(1, CountArrayProfPK)
                        ArrayFrofPK(0, CountArrayProfPK) = ProfRezult.StartingStation
                        ArrayFrofPK(1, CountArrayProfPK) = "PIKET"
                        CountArrayProfPK += 1
                        For i As Double = IntPKStart To IntPKEnd Step ProfDialog.NumericUpDown1.Value
                            ReDim Preserve ArrayFrofPK(1, CountArrayProfPK)
                            ArrayFrofPK(0, CountArrayProfPK) = i
                            ArrayFrofPK(1, CountArrayProfPK) = "PIKET"
                            CountArrayProfPK += 1
                        Next i
                        'записываем последнее значение
                        ReDim Preserve ArrayFrofPK(1, CountArrayProfPK)
                        ArrayFrofPK(0, CountArrayProfPK) = ProfRezult.EndingStation
                        ArrayFrofPK(1, CountArrayProfPK) = "PIKET"
                    End If
                End If
                'сортируем массив с пикетами
                If IsArray(ArrayFrofPK) = True Then
                    Dim BoolSortArray As Boolean = MathFunction.FuncSortStrArray(ArrayFrofPK, 0)
                Else
                    MsgBox("Массив с пикетами пуст")
                    Exit Sub
                End If
                'начинаем перебор объектов
                Dim PrevPK As Double = -10000000
                For i As Double = 0 To ArrayFrofPK.GetUpperBound(1)
                    '1. берем пикет
                    Dim PKRedProfile As Double = ArrayFrofPK(0, i) 'пикет
                    Dim NameSegm As String = ArrayFrofPK(1, i) 'обозначение сегмента
                    'если такой пикет уже был то повторно его не обрабатываем
                    If PKRedProfile = PrevPK Then
                        GoTo line2
                    Else
                        PrevPK = PKRedProfile
                    End If
                    '2. Округляем пикет
                    Dim PKRedRound As Double = Math.Round(PKRedProfile, RoundZnH)
                    Dim HRed As Double = 0 'высота красного профиля
                    Dim HBlack As Double = 0 'высота черного профиля
                    Dim HDop As Double = 0 'высота дополнительного профиля
                    '3.находим высоту красного профиля
                    Dim boolRedH As Boolean = False
                    Try
                        HRed = ProfileRed.ElevationAt(PKRedRound)
                        boolRedH = True
                    Catch ex As System.ArgumentException
                        boolRedH = False
                    End Try
                    'если высота не найдена пытаемся определить может это начато трассы или конец
                    If boolRedH = False Then
                        Dim d1 As Double = Math.Abs(PKRedProfile - ProfileRed.StartingStation)
                        Dim d2 As Double = Math.Abs(PKRedProfile - ProfileRed.EndingStation)
                        If d1 < d2 Then
                            Try
                                PKRedRound = MathFunction.FuncTrimDigitCall(PKRedProfile, RoundZnH)
                                HRed = ProfileRed.ElevationAt(PKRedRound)
                                boolRedH = True
                            Catch ex As System.ArgumentException
                                GoTo line2
                            End Try
                        Else
                            Try
                                PKRedRound = MathFunction.FuncTrimDigitFloo(PKRedProfile, RoundZnH)
                                HRed = ProfileRed.ElevationAt(PKRedRound)
                                boolRedH = True
                            Catch ex As System.ArgumentException
                                GoTo line2
                            End Try
                        End If
                    End If
                    '4.находим высоту черного профиля
                    Dim boolBlackH As Boolean = False
                    Try
                        HBlack = ProfileBlask.ElevationAt(PKRedRound)
                        boolBlackH = True
                    Catch ex As System.ArgumentException
                        boolBlackH = False
                    End Try
                    'если высота не найдена пытаемся определить может это начато трассы или конец
                    If boolBlackH = False Then
                        Dim d1 As Double = Math.Abs(PKRedProfile - ProfileBlask.StartingStation)
                        Dim d2 As Double = Math.Abs(PKRedProfile - ProfileBlask.EndingStation)
                        If d1 < d2 Then
                            Try
                                PKRedRound = MathFunction.FuncTrimDigitCall(PKRedProfile, RoundZnH)
                                HBlack = ProfileBlask.ElevationAt(PKRedRound)
                                boolBlackH = True
                            Catch ex As System.ArgumentException
                                GoTo line2
                            End Try
                        Else
                            Try
                                PKRedRound = MathFunction.FuncTrimDigitFloo(PKRedProfile, RoundZnH)
                                HBlack = ProfileBlask.ElevationAt(PKRedRound)
                                boolRedH = True
                            Catch ex As System.ArgumentException
                                GoTo line2
                            End Try
                        End If
                    End If
                    '4.находим высоту дополнительного профиля
                    Dim boolDopH As Boolean = False
                    If IsNothing(ProfileDop) = False Then
                        Try
                            HDop = ProfileDop.ElevationAt(PKRedRound)
                            boolDopH = True
                        Catch ex As System.ArgumentException
                            boolDopH = False
                        End Try
                        'если высота не найдена пытаемся определить может это начато трассы или конец
                        If boolBlackH = False Then
                            Dim d1 As Double = Math.Abs(PKRedProfile - ProfileDop.StartingStation)
                            Dim d2 As Double = Math.Abs(PKRedProfile - ProfileDop.EndingStation)
                            If d1 < d2 Then
                                Try
                                    PKRedRound = MathFunction.FuncTrimDigitCall(PKRedProfile, RoundZnH)
                                    HDop = ProfileDop.ElevationAt(PKRedRound)
                                    boolDopH = True
                                Catch ex As System.ArgumentException
                                    GoTo line2
                                End Try
                            Else
                                Try
                                    PKRedRound = MathFunction.FuncTrimDigitFloo(PKRedProfile, RoundZnH)
                                    HBlack = ProfileDop.ElevationAt(PKRedRound)
                                    boolDopH = True
                                Catch ex As System.ArgumentException
                                    GoTo line2
                                End Try
                            End If
                        End If
                    End If

                    '4. вычисляем дельту
                    Dim delta As Double = 0
                    delta = Math.Round((HRed - HBlack), RoundZnH)
                    Dim delta1 As Double = 0
                    If ProfDialog.CheckBox7.Checked = True Then
                        delta1 = Math.Round((HRed - HDop), RoundZnH)
                    End If
                    '3. Инициализируем многострочный текст
                    Dim Mtext1 As MText = New MText
                    Mtext1.SetDatabaseDefaults()
                    Dim TextD As String = ""
                    If RoundZnH = 0 Then
                        TextD = delta
                    ElseIf RoundZnH = 1 Then
                        TextD = String.Format(CultureInfo.InvariantCulture, "{0:0.0}", Math.Abs(delta))
                    ElseIf RoundZnH = 2 Then
                        TextD = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Math.Abs(delta))
                    ElseIf RoundZnH = 3 Then
                        TextD = String.Format(CultureInfo.InvariantCulture, "{0:0.000}", Math.Abs(delta))
                    ElseIf RoundZnH = 4 Then
                        TextD = String.Format(CultureInfo.InvariantCulture, "{0:0.0000}", Math.Abs(delta))
                    End If
                    If ProfDialog.CheckBox7.Checked = True Then
                        If RoundZnH = 0 Then
                            TextD = delta1
                        ElseIf RoundZnH = 1 Then
                            TextD = String.Format(CultureInfo.InvariantCulture, "{0:0.0}", Math.Abs(delta1))
                        ElseIf RoundZnH = 2 Then
                            TextD = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Math.Abs(delta1))
                        ElseIf RoundZnH = 3 Then
                            TextD = String.Format(CultureInfo.InvariantCulture, "{0:0.000}", Math.Abs(delta1))
                        ElseIf RoundZnH = 4 Then
                            TextD = String.Format(CultureInfo.InvariantCulture, "{0:0.0000}", Math.Abs(delta1))
                        End If
                        TextD = "(" & TextD & ")"
                        'OtstupY = OtstupY + 3
                    End If
                    Mtext1.Contents = TextD 'записываем дельту по модулю
                    '4. находим координаты точки вставки текста
                    Dim x As Double = 0
                    Dim y As Double = 0
                    ProfileWiev.FindXYAtStationAndElevation(PKRedRound, HRed, x, y)
                    If delta < 0 Then
                        y = y - OtstupY 'если число отритцательное то записывам ниже красного профиля
                        Mtext1.Attachment = AttachmentPoint.MiddleRight
                    Else
                        y = y + OtstupY
                        Mtext1.Attachment = AttachmentPoint.MiddleLeft
                    End If
                    Dim Pt1 As Point3d = New Point3d(x, y, 0)
                    Mtext1.Location = Pt1
                    Mtext1.Rotation = Math.PI / 2
                    Mtext1.TextHeight = HText
                    If IsNothing(idStyleText) = False Then
                        Mtext1.TextStyleId = idStyleText
                    End If
                    If IsNothing(IDInsLayer) = False Then
                        If IDInsLayer.IsNull = False Then
                            Mtext1.Layer = LayerText
                        End If
                    End If
                    acBlkTblRec.AppendEntity(Mtext1)
                    tr.AddNewlyCreatedDBObject(Mtext1, True)
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'ЗАПИСЫВАМ В МЕТКУ РАСШИРЕННЫЕ ДАННЫЕ
                    'добавляем словарь в примитив
                    Dim extId As ObjectId = Mtext1.ExtensionDictionary
                    If extId.IsNull = True Then
                        Mtext1.UpgradeOpen()
                        Mtext1.CreateExtensionDictionary()
                        extId = Mtext1.ExtensionDictionary
                    End If
                    Dim dbExt As DBDictionary = tr.GetObject(extId, OpenMode.ForRead)
                    dbExt.UpgradeOpen()
                    Dim xRec As Xrecord = New Xrecord()
                    Dim rb As ResultBuffer = New ResultBuffer()
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Рабочие отметки")) 'имя начальной группы
                    rb.Add(New TypedValue(DxfCode.ExtendedDataReal, randomNumber)) 'номер группы индивидуальный
                    rb.Add(New TypedValue(DxfCode.ExtendedDataHandle, ProfileWiev.Handle)) 'вид профиля
                    rb.Add(New TypedValue(DxfCode.ExtendedDataHandle, ProfileRed.Handle)) 'красный профиль
                    rb.Add(New TypedValue(DxfCode.ExtendedDataHandle, ProfileBlask.Handle)) 'черный профиль
                    If IsNothing(ProfileDop) = True Then
                        rb.Add(New TypedValue(DxfCode.ExtendedDataHandle, 0)) 'дополнительный профиль
                    Else
                        rb.Add(New TypedValue(DxfCode.ExtendedDataHandle, ProfileDop.Handle)) 'дополнительный профиль
                    End If
                    rb.Add(New TypedValue(DxfCode.ExtendedDataReal, PKRedRound)) 'пикет
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, NameSegm)) 'пикет
                    xRec.Data = rb
                    dbExt.SetAt("WORKMARK", xRec)
                    tr.AddNewlyCreatedDBObject(xRec, True)
line2:
                Next i

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'удалить отметки рабочие
            ElseIf ProfDialog.RadioButton4.Checked = True Then
                'Удаление рабочих отметок
                '1. выбираем метку для образца
                Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите рабочую отметку: ")
                Dim resBlk As PromptEntityResult = m_Editor.GetEntity(opt)
                Dim pvId As ObjectId = resBlk.ObjectId
                Dim RndErase As Integer = 0
                Dim objMText As Autodesk.AutoCAD.DatabaseServices.MText = tr.GetObject(pvId, OpenMode.ForRead)
                If TypeOf objMText Is MText Then
                    Dim extId As ObjectId = objMText.ExtensionDictionary
                    If extId.IsNull = False Then
                        Dim extDict As DBDictionary = tr.GetObject(extId, OpenMode.ForRead)
                        If extDict.Contains("WORKMARK") Then
                            Dim xRec As Xrecord = tr.GetObject(extDict.GetAt("WORKMARK"), OpenMode.ForRead)
                            Dim Rec As TypedValue() = xRec.Data.AsArray()
                            RndErase = Rec(1).Value
                        End If
                    End If
                End If
                '1.формируем фильтр для многострочных текстов
                Dim ssetMText As SelectionSet = Nothing
                Dim ValType1(1) As TypedValue
                ValType1.SetValue(New TypedValue(DxfCode.Start, "MText"), 0)
                ValType1.SetValue(New TypedValue(DxfCode.Visibility, 0), 1)
                Dim FilterMText As SelectionFilter = New SelectionFilter(ValType1)
                Dim MTextSelect As PromptSelectionResult = m_Editor.SelectAll(FilterMText)
                If MTextSelect.Status = PromptStatus.OK Then
                    ssetMText = MTextSelect.Value
                    Dim countKounter2 As Integer = ssetMText.Count
                    For Each idMText As ObjectId In ssetMText.GetObjectIds
                        Dim objEnt As Autodesk.AutoCAD.DatabaseServices.MText = tr.GetObject(idMText, OpenMode.ForWrite)
                        Dim extId As ObjectId = objEnt.ExtensionDictionary
                        If extId.IsNull = False Then
                            Dim extDict As DBDictionary = tr.GetObject(extId, OpenMode.ForRead)
                            If extDict.Contains("WORKMARK") Then
                                Dim xRec As Xrecord = tr.GetObject(extDict.GetAt("WORKMARK"), OpenMode.ForRead)
                                Dim Rec As TypedValue() = xRec.Data.AsArray()
                                Dim Rnd As Integer = Rec(1).Value
                                If Rnd = RndErase Then
                                    objEnt.Erase()
                                End If
                            End If
                        End If
                    Next
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'вставка отметок вручную
            ElseIf ProfDialog.RadioButton2.Checked = True Then
                '1. выбираем метку для образца
                Dim optMetka As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите рабочую отметку в качестве образца: ")
                Dim resMetka As PromptEntityResult = m_Editor.GetEntity(optMetka)
                Dim RndRecMetka As Integer = 0 'индивидуальный номер метки
                Dim extId As ObjectId = Nothing
                If resMetka.Status = PromptStatus.OK Then
                    Dim IdMetka As ObjectId = resMetka.ObjectId
                    Dim objEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(IdMetka, OpenMode.ForRead)
                    extId = objEnt.ExtensionDictionary
                    If extId.IsNull = False Then
                        Dim extDict As DBDictionary = tr.GetObject(extId, OpenMode.ForRead)
                        If extDict.Contains("WORKMARK") Then
                            Dim xRec As Xrecord = tr.GetObject(extDict.GetAt("WORKMARK"), OpenMode.ForRead)
                            Dim Rec As TypedValue() = xRec.Data.AsArray()
                            'номер группы
                            RndRecMetka = Rec(1).Value
                            'ссылки на профиля
                            '==================================================================
                            'красный профиль
                            Dim HandlRedProfile As String = Rec(3).Value
                            Dim idPrvRed As ObjectId = AcadFunction.FuncReturnObjectID(HandlRedProfile.ToString)
                            If IsNothing(idPrvRed) = True Then
                                'заново запускаем диалог с выбором красного профиля
                                Dim selectRedResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите красный профиль:")
                                If selectRedResult.Status = PromptStatus.OK Then
                                    Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectRedResult.ObjectId, OpenMode.ForRead)
                                    ProfileRed = TryCast(dbObj, Profile)
                                End If
                            Else
                                ProfileRed = tr.GetObject(idPrvRed, OpenMode.ForRead)
                            End If
                            '==========================================================================
                            'черный профиль
                            Dim HandBlackProfile As String = Rec(4).Value
                            Dim idPrvBlack As ObjectId = AcadFunction.FuncReturnObjectID(HandBlackProfile.ToString)
                            If IsNothing(idPrvBlack) = True Then
                                'выбираем черный прфиль
                                Dim selectBlackResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите черный профиль:")
                                If selectBlackResult.Status = PromptStatus.OK Then
                                    Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectBlackResult.ObjectId, OpenMode.ForRead)
                                    ProfileBlask = TryCast(dbObj, Profile)
                                End If
                            Else
                                ProfileBlask = tr.GetObject(idPrvBlack, OpenMode.ForRead)
                            End If
                            '===============================================================================
                            'вид профиля
                            Dim HAndlProfileView As String = Rec(2).Value 'вид профиля
                            Dim idPrvView As ObjectId = AcadFunction.FuncReturnObjectID(HAndlProfileView.ToString)
                            If IsNothing(idPrvView) = True Then
                                'заново запускаем диалог с выбором вида профиля
                                'выбираем вид профиля
                                Dim selectWiewProfileResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите вид профиля:")
                                If selectWiewProfileResult.Status = PromptStatus.OK Then
                                    Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectWiewProfileResult.ObjectId, OpenMode.ForRead)
                                    ProfileWiev = TryCast(dbObj, ProfileView)
                                End If
                            Else
                                ProfileWiev = tr.GetObject(idPrvView, OpenMode.ForRead)
                            End If
                            '===================================================================================
                            'дополнительный профиль
                            Dim HAndlDopProfile As String = Rec(5).Value 'вид профиля
                            If Not (HAndlDopProfile Like "0") Then
                                Dim idPrvDop As ObjectId = AcadFunction.FuncReturnObjectID(HAndlDopProfile.ToString)
                                If IsNothing(idPrvDop) = True Then
                                    'заново запускаем диалог с выбором дополнительного профиля
                                    Dim selectWiewProfileResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите дополнительный профиль:")
                                    If selectWiewProfileResult.Status = PromptStatus.OK Then
                                        Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectWiewProfileResult.ObjectId, OpenMode.ForRead)
                                        ProfileDop = TryCast(dbObj, Profile)
                                    End If
                                Else
                                    ProfileDop = tr.GetObject(idPrvDop, OpenMode.ForRead)
                                End If
                            End If
                            '==================================================================================================
                            'считываем свойства метки
                            LayerText = objEnt.Layer
                            If TypeOf objEnt Is MText Then
                                Dim MText2 As MText = objEnt
                                idStyleText = MText2.TextStyleId
                                HText = MText2.TextHeight
                            End If
                            If IsNothing(ProfileBlask) = True Or IsNothing(ProfileRed) = True Or IsNothing(ProfileWiev) = True Then
                                MsgBox("Ошибка выбора данных. Макро прервано!!!")
                                Exit Sub
                            End If
                        Else
                            Dim selectRedResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите красный профиль:")
                            If selectRedResult.Status = PromptStatus.OK Then
                                Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectRedResult.ObjectId, OpenMode.ForRead)
                                ProfileRed = TryCast(dbObj, Profile)
                            End If
                            'выбираем черный прфиль
                            Dim selectBlackResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите черный профиль:")
                            If selectBlackResult.Status = PromptStatus.OK Then
                                Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectBlackResult.ObjectId, OpenMode.ForRead)
                                ProfileBlask = TryCast(dbObj, Profile)
                            End If
                            'выбираем вид профиля
                            Dim selectWiewProfileResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите вид профиля:")
                            If selectWiewProfileResult.Status = PromptStatus.OK Then
                                Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectWiewProfileResult.ObjectId, OpenMode.ForRead)
                                ProfileWiev = TryCast(dbObj, ProfileView)
                            End If
                            If ProfDialog.CheckBox7.Checked = True Then
                                'заново запускаем диалог с выбором дополнительного профиля
                                Dim selectWiewProfileResult1 As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите дополнительный профиль:")
                                If selectWiewProfileResult1.Status = PromptStatus.OK Then
                                    Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectWiewProfileResult1.ObjectId, OpenMode.ForRead)
                                    ProfileDop = TryCast(dbObj, Profile)
                                End If
                            End If
                        End If
                    Else
                        Dim selectRedResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите красный профиль:")
                        If selectRedResult.Status = PromptStatus.OK Then
                            Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectRedResult.ObjectId, OpenMode.ForRead)
                            ProfileRed = TryCast(dbObj, Profile)
                        End If
                        'выбираем черный прфиль
                        Dim selectBlackResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите черный профиль:")
                        If selectBlackResult.Status = PromptStatus.OK Then
                            Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectBlackResult.ObjectId, OpenMode.ForRead)
                            ProfileBlask = TryCast(dbObj, Profile)
                        End If
                        'выбираем вид профиля
                        Dim selectWiewProfileResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите вид профиля:")
                        If selectWiewProfileResult.Status = PromptStatus.OK Then
                            Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectWiewProfileResult.ObjectId, OpenMode.ForRead)
                            ProfileWiev = TryCast(dbObj, ProfileView)
                        End If
                        If ProfDialog.CheckBox7.Checked = True Then
                            'заново запускаем диалог с выбором дополнительного профиля
                            Dim selectWiewProfileResult1 As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите дополнительный профиль:")
                            If selectWiewProfileResult1.Status = PromptStatus.OK Then
                                Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectWiewProfileResult1.ObjectId, OpenMode.ForRead)
                                ProfileDop = TryCast(dbObj, Profile)
                            End If
                        End If
                    End If
                Else
                    Dim selectRedResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите красный профиль:")
                    If selectRedResult.Status = PromptStatus.OK Then
                        Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectRedResult.ObjectId, OpenMode.ForRead)
                        ProfileRed = TryCast(dbObj, Profile)
                    End If
                    'выбираем черный прфиль
                    Dim selectBlackResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите черный профиль:")
                    If selectBlackResult.Status = PromptStatus.OK Then
                        Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectBlackResult.ObjectId, OpenMode.ForRead)
                        ProfileBlask = TryCast(dbObj, Profile)
                    End If
                    'выбираем вид профиля
                    Dim selectWiewProfileResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите вид профиля:")
                    If selectWiewProfileResult.Status = PromptStatus.OK Then
                        Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectWiewProfileResult.ObjectId, OpenMode.ForRead)
                        ProfileWiev = TryCast(dbObj, ProfileView)
                    End If
                    If ProfDialog.CheckBox7.Checked = True Then
                        'заново запускаем диалог с выбором дополнительного профиля
                        Dim selectWiewProfileResult1 As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите дополнительный профиль:")
                        If selectWiewProfileResult1.Status = PromptStatus.OK Then
                            Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(selectWiewProfileResult1.ObjectId, OpenMode.ForRead)
                            ProfileDop = TryCast(dbObj, Profile)
                        End If
                    End If
                End If
                If IsNothing(ProfileBlask) = True Or IsNothing(ProfileRed) = True Or IsNothing(ProfileWiev) = True Then
                    MsgBox("Ошибка выбора данных. Макро прервано!!!")
                    Exit Sub
                End If
                'выбираем точку
                Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                pPtOpts.AllowNone = True
                pPtOpts.Message = vbLf & "Укажите точку на виде профиля: "
                Do
                    Dim pPtRes As PromptPointResult = m_Editor.GetPoint(pPtOpts)
                    If pPtRes.Status = PromptStatus.OK Then
                        Dim ptStart As Point3d = pPtRes.Value 'координаты точки начала трассы
                        Dim PK As Double = 0
                        Dim Elev As Double = 0
                        ProfileWiev.FindStationAndElevationAtXY(ptStart.X, ptStart.Y, PK, Elev)
                        '2. По пикету находим высоту красного профиля
                        Dim HRed As Double = ProfileRed.ElevationAt(PK)
                        'По пикету находим высоту черного профиля
                        Dim HBlack As Double = ProfileBlask.ElevationAt(PK)
                        Dim HDop As Double = 0
                        Dim delta1 As Double = 0
                        If IsNothing(ProfileDop) = False Then
                            HDop = ProfileDop.ElevationAt(PK)
                            delta1 = Math.Round((HRed - HDop), RoundZnH)
                        End If
                        'вычисляем дельту
                        Dim delta As Double = Math.Round((HRed - HBlack), RoundZnH)
                        '3. Инициализируем многострочный текст
                        Dim Mtext1 As MText = New MText
                        Mtext1.SetDatabaseDefaults()
                        'delta = Math.Abs(delta)
                        Dim DText As String = ""
                        If RoundZnH = 0 Then
                            DText = delta
                        ElseIf RoundZnH = 1 Then
                            DText = String.Format(CultureInfo.InvariantCulture, "{0:0.0}", Math.Abs(delta))
                        ElseIf RoundZnH = 2 Then
                            DText = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Math.Abs(delta))
                        ElseIf RoundZnH = 3 Then
                            DText = String.Format(CultureInfo.InvariantCulture, "{0:0.000}", Math.Abs(delta))
                        ElseIf RoundZnH = 4 Then
                            DText = String.Format(CultureInfo.InvariantCulture, "{0:0.0000}", Math.Abs(delta))
                        End If
                        If IsNothing(ProfileDop) = False Then
                            Dim TextD1 As String = ""
                            If RoundZnH = 0 Then
                                TextD1 = delta1
                            ElseIf RoundZnH = 1 Then
                                TextD1 = String.Format(CultureInfo.InvariantCulture, "{0:0.0}", Math.Abs(delta1))
                            ElseIf RoundZnH = 2 Then
                                TextD1 = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Math.Abs(delta1))
                            ElseIf RoundZnH = 3 Then
                                TextD1 = String.Format(CultureInfo.InvariantCulture, "{0:0.000}", Math.Abs(delta1))
                            ElseIf RoundZnH = 4 Then
                                TextD1 = String.Format(CultureInfo.InvariantCulture, "{0:0.0000}", Math.Abs(delta1))
                            End If
                            DText = DText & " (" & TextD1 & ")"
                        End If
                        Mtext1.Contents = DText
                        '4. находим координаты точки вставки текста
                        Dim x As Double = 0
                        Dim y As Double = 0
                        ProfileWiev.FindXYAtStationAndElevation(PK, HRed, x, y)
                        If delta < 0 Then
                            y = y - OtstupY 'если число отритцательное то записывам ниже красного профиля
                            Mtext1.Attachment = AttachmentPoint.MiddleRight
                        Else
                            y = y + OtstupY
                            Mtext1.Attachment = AttachmentPoint.MiddleLeft
                        End If
                        Dim Pt1 As Point3d = New Point3d(x, y, 0)
                        Mtext1.Location = Pt1
                        Mtext1.Rotation = Math.PI / 2
                        Mtext1.TextHeight = HText
                        If IsNothing(idStyleText) = False Then
                            Mtext1.TextStyleId = idStyleText
                        End If
                        If IsNothing(IDInsLayer) = False Then
                            If IDInsLayer.IsNull = False Then
                                Mtext1.Layer = LayerText
                            End If
                        End If
                        acBlkTblRec.AppendEntity(Mtext1)
                        tr.AddNewlyCreatedDBObject(Mtext1, True)
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'ЗАПИСЫВАМ В МЕТКУ РАСШИРЕННЫЕ ДАННЫЕ
                        'добавляем словарь в примитив
                        extId = Mtext1.ExtensionDictionary
                        If RndRecMetka = 0 Then
                            Dim randomObject As New Random() 'формируем индивидуальный id 
                            Dim randomNumber As Integer = randomObject.Next()
                            RndRecMetka = randomNumber
                        End If
                        If extId.IsNull = True Then
                            Mtext1.UpgradeOpen()
                            Mtext1.CreateExtensionDictionary()
                            extId = Mtext1.ExtensionDictionary
                        End If
                        Dim dbExt As DBDictionary = tr.GetObject(extId, OpenMode.ForRead)
                        dbExt.UpgradeOpen()
                        Dim xRec As Xrecord = New Xrecord()
                        Dim rb As ResultBuffer = New ResultBuffer()
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Рабочие отметки")) 'имя начальной группы
                        rb.Add(New TypedValue(DxfCode.ExtendedDataReal, RndRecMetka)) 'номер группы индивидуальный
                        rb.Add(New TypedValue(DxfCode.ExtendedDataHandle, ProfileWiev.Handle)) 'вид профиля
                        rb.Add(New TypedValue(DxfCode.ExtendedDataHandle, ProfileRed.Handle)) 'красный профиль
                        rb.Add(New TypedValue(DxfCode.ExtendedDataHandle, ProfileBlask.Handle)) 'черный профиль
                        If IsNothing(ProfileDop) = False Then
                            rb.Add(New TypedValue(DxfCode.ExtendedDataHandle, ProfileDop.Handle)) 'дополнительный профмль профиль
                        Else
                            rb.Add(New TypedValue(DxfCode.ExtendedDataHandle, 0)) 'дополнительный профмль профиль
                        End If
                        rb.Add(New TypedValue(DxfCode.ExtendedDataReal, PK)) 'пикет
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "PIKET")) 'пикет
                        xRec.Data = rb
                        dbExt.SetAt("WORKMARK", xRec)
                        tr.AddNewlyCreatedDBObject(xRec, True)
                    Else
                        Exit Do
                    End If
                Loop
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'редактирование рабочих отметок (обновление)
            ElseIf ProfDialog.RadioButton3.Checked = True Then
                '1. выбираем метку для образца
                Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите рабочую отметку: ")
                Dim resBlk As PromptEntityResult = m_Editor.GetEntity(opt)
                Dim pvId As ObjectId = resBlk.ObjectId
                Dim RndRec As Integer = 0 'индивидуальный номер метки
                Dim objMText As Autodesk.AutoCAD.DatabaseServices.MText = tr.GetObject(pvId, OpenMode.ForRead)
                If TypeOf objMText Is MText Then
                    Dim extId As ObjectId = objMText.ExtensionDictionary
                    If extId.IsNull = False Then
                        Dim extDict As DBDictionary = tr.GetObject(extId, OpenMode.ForRead)
                        If extDict.Contains("WORKMARK") Then
                            Dim xRec As Xrecord = tr.GetObject(extDict.GetAt("WORKMARK"), OpenMode.ForRead)
                            Dim Rec As TypedValue() = xRec.Data.AsArray()
                            'номер группы
                            RndRec = Rec(1).Value
                            'вид профиля
                            Dim HAndlProfileView As String = Rec(2).Value 'вид профиля
                            Dim idPrvView As ObjectId = AcadFunction.FuncReturnObjectID(HAndlProfileView.ToString)
                            ProfileWiev = tr.GetObject(idPrvView, OpenMode.ForWrite)
                            'красный профиль
                            Dim HandlRedProfile As String = Rec(3).Value
                            Dim idPrvRed As ObjectId = AcadFunction.FuncReturnObjectID(HandlRedProfile.ToString)
                            ProfileRed = tr.GetObject(idPrvRed, OpenMode.ForWrite)
                            'черный профиль
                            Dim HandBlackProfile As String = Rec(4).Value
                            Dim idPrvBlack As ObjectId = AcadFunction.FuncReturnObjectID(HandBlackProfile.ToString)
                            ProfileBlask = tr.GetObject(idPrvBlack, OpenMode.ForWrite)
                            '1.формируем фильтр для многострочных текстов
                            Dim ssetMText As SelectionSet = Nothing
                            Dim ValType1(1) As TypedValue
                            ValType1.SetValue(New TypedValue(DxfCode.Start, "MText"), 0)
                            ValType1.SetValue(New TypedValue(DxfCode.Visibility, 0), 1)
                            Dim FilterMText As SelectionFilter = New SelectionFilter(ValType1)
                            Dim MTextSelect As PromptSelectionResult = m_Editor.SelectAll(FilterMText)
                            If MTextSelect.Status = PromptStatus.OK Then
                                ssetMText = MTextSelect.Value
                                Dim countKounter2 As Integer = ssetMText.Count
                                For Each idMText As ObjectId In ssetMText.GetObjectIds
                                    Dim objEnt As Autodesk.AutoCAD.DatabaseServices.MText = tr.GetObject(idMText, OpenMode.ForWrite)
                                    extId = objEnt.ExtensionDictionary
                                    If extId.IsNull = False Then
                                        extDict = tr.GetObject(extId, OpenMode.ForRead)
                                        If extDict.Contains("WORKMARK") Then
                                            xRec = tr.GetObject(extDict.GetAt("WORKMARK"), OpenMode.ForRead)
                                            Rec = xRec.Data.AsArray()
                                            Dim Rnd As Integer = Rec(1).Value
                                            If Rnd = RndRec Then
                                                Dim PK As Double = Rec(5).Value
                                                '2. По пикету находим высоту красного профиля
                                                Try
                                                    Dim HRed As Double = ProfileRed.ElevationAt(PK)
                                                    'По пикету находим высоту черного профиля
                                                    Dim HBlack As Double = ProfileBlask.ElevationAt(PK)
                                                    'вычисляем дельту
                                                    Dim delta As Double = Math.Round((HRed - HBlack), RoundZnH)
                                                    'delta = Math.Abs(delta)
                                                    If RoundZnH = 0 Then
                                                        objEnt.Contents = delta
                                                    ElseIf RoundZnH = 1 Then
                                                        objEnt.Contents = String.Format(CultureInfo.InvariantCulture, "{0:0.0}", Math.Abs(delta))
                                                    ElseIf RoundZnH = 2 Then
                                                        objEnt.Contents = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Math.Abs(delta))
                                                    ElseIf RoundZnH = 3 Then
                                                        objEnt.Contents = String.Format(CultureInfo.InvariantCulture, "{0:0.000}", Math.Abs(delta))
                                                    ElseIf RoundZnH = 4 Then
                                                        objEnt.Contents = String.Format(CultureInfo.InvariantCulture, "{0:0.0000}", Math.Abs(delta))
                                                    End If
                                                Catch ex As System.ArgumentException
                                                End Try
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            End If
            tr.Commit()
        End Using
    End Sub
End Class
