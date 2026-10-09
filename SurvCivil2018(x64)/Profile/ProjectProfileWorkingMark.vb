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
Imports System.Globalization
'рабочие отметки на виде профиля
Public Class ProjectProfileWorkingMark
    <CommandMethod("ProjProfileWorkMarks")>
    Public Sub ProjectProfileWorkingMark()
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

        Dim objAlign As Autodesk.Civil.DatabaseServices.Alignment = Nothing 'объект типа трасса
        Dim ProfileWiev As Autodesk.Civil.DatabaseServices.ProfileView = Nothing 'вид профиля
        Dim ProfileRed As Autodesk.Civil.DatabaseServices.Profile = Nothing 'красный профиль
        Dim ProfileBlask As Autodesk.Civil.DatabaseServices.Profile = Nothing 'черный профиль
        Dim ProfileDop As Autodesk.Civil.DatabaseServices.Profile = Nothing 'дополнительный профиль
        '=====================================================================================================================
        'запуск диалогового окна
        Dim FormProjectProfileWorkMark As ProjProfileWorkingMark = New ProjProfileWorkingMark
        FormProjectProfileWorkMark.ShowDialog()
        If FormProjectProfileWorkMark.boolProf = False Then Exit Sub
        '=====================================================================================================================
        'отступ от оси
        Dim OtstupY As Double = FormProjectProfileWorkMark.NumericUpDown2.Value
        If FormProjectProfileWorkMark.ComboBox7.Text.Trim.Length > 0 Then
            OtstupY = OtstupY + 4.5
        End If
        '====================================================================================================================
        'тело программы
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            '=============================================================================================================
            'вид профиля
            ProfileWiev = FormProjectProfileWorkMark.ViewProfileObj
            If IsNothing(ProfileWiev) = True Then
                MsgBox("Вид профиля не найден. Макро прервано!!!")
                Exit Sub
            End If
            If ProfileWiev.Id.IsNull = True Then
                MsgBox("Вид профиля не найден. Макро прервано!!!")
                Exit Sub
            End If
            '================================================================================================================
            'выбираем красный профиль
            Dim NameRedProfile As String = FormProjectProfileWorkMark.ComboBox2.Text.Trim
            If IsNothing(NameRedProfile) = True Then
                MsgBox("Красный профиль не найден. Макро прервано!!!")
                Exit Sub
            End If
            '================================================================================================================
            'выбираем черный профиль
            Dim NameBlaskProfile As String = FormProjectProfileWorkMark.ComboBox3.Text.Trim
            If IsNothing(NameBlaskProfile) = True Then
                MsgBox("Черный профиль не найден. Макро прервано!!!")
                Exit Sub
            End If
            '================================================================================================================
            'выбираем дополнительный профиль
            Dim NameDopProfile As String = FormProjectProfileWorkMark.ComboBox7.Text.Trim
            '=================================================================================================================
            Dim idAlign As ObjectId = ProfileWiev.AlignmentId
            Dim Align As Alignment = tr.GetObject(idAlign, OpenMode.ForRead)
            Dim ProfileColl As ObjectIdCollection = Align.GetProfileIds
            If ProfileColl.Count > 1 Then
                For Each id As ObjectId In ProfileColl
                    Dim ObjTempProfile As Profile = tr.GetObject(id, OpenMode.ForRead)
                    If ObjTempProfile.Name Like NameRedProfile Then
                        ProfileRed = ObjTempProfile
                    ElseIf ObjTempProfile.Name Like NameBlaskProfile Then
                        ProfileBlask = ObjTempProfile
                    ElseIf ObjTempProfile.Name Like NameDopProfile Then
                        ProfileDop = ObjTempProfile
                    End If
                Next
            End If
            If IsNothing(ProfileBlask) = True Then
                MsgBox("Черный профиль не найден. Макро прервано!!!")
                Exit Sub
            End If
            If ProfileBlask.Id.IsNull = True Then
                MsgBox("Черный профиль не найден. Макро прервано!!!")
                Exit Sub
            End If
            If IsNothing(ProfileRed) = True Then
                MsgBox("Красный профиль не найден. Макро прервано!!!")
                Exit Sub
            End If
            If ProfileRed.Id.IsNull = True Then
                MsgBox("Красный профиль не найден. Макро прервано!!!")
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '===============================================================================================================================
            'работа с группой меток Civil 3d
            If FormProjectProfileWorkMark.RadioButton6.Checked = True Then
                Dim ProfileLabel As ProfileDataBandLabelGroup = Nothing
                'выберите метку
                Dim selectLabelResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите группу меток с рабочими отметками:")
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
                'метка Civil 3d
            ElseIf FormProjectProfileWorkMark.RadioButton5.Checked = True Then
                '1. запускаем генератор случайных чисел
                'массив с пикетажем
                Dim ProfRezult As Autodesk.Civil.DatabaseServices.Profile = Nothing
                Dim ArrayProfileObj As String(,) = Nothing
                Dim CountArrayProfileObj As Integer = 0
                'точки вертикальной геометрии красного профиля
                If FormProjectProfileWorkMark.CheckBox1.Checked = True Or FormProjectProfileWorkMark.CheckBox5.Checked = True Then
                    Dim FindVertGeomRed As Boolean = FuncAlignment.FuncReadVerticalGeometryToStatinName(ProfileRed, ArrayProfileObj)
                    ProfRezult = ProfileRed
                End If
                'точки вертикальной геометрии черного профиля
                If FormProjectProfileWorkMark.CheckBox2.Checked = True Then
                    Dim FindVertGeomRed As Boolean = FuncAlignment.FuncReadVerticalGeometryToStatinName(ProfileBlask, ArrayProfileObj)
                End If
                'точки вертикальной геометрии дополнительного профиля
                If IsNothing(ProfileDop) = False Then
                    If FormProjectProfileWorkMark.CheckBox5.Checked = True Then
                        Dim FindVertGeomRed As Boolean = FuncAlignment.FuncReadVerticalGeometryToStatinName(ProfileDop, ArrayProfileObj)
                    End If
                End If
                'если есть шаг
                If FormProjectProfileWorkMark.CheckBox4.Checked = True Then
                    'определяем старт
                    If IsArray(ArrayProfileObj) = True Then
                        CountArrayProfileObj = ArrayProfileObj.GetUpperBound(1) + 1
                        'получаем начальное значение цикла
                        Dim ChagRound As Double = FormProjectProfileWorkMark.NumericUpDown1.Value
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
                        ReDim Preserve ArrayProfileObj(1, CountArrayProfileObj)
                        ArrayProfileObj(0, CountArrayProfileObj) = ProfRezult.StartingStation
                        ArrayProfileObj(1, CountArrayProfileObj) = "PIKET"
                        CountArrayProfileObj += 1
                        For i As Double = IntPKStart To IntPKEnd Step FormProjectProfileWorkMark.NumericUpDown1.Value
                            ReDim Preserve ArrayProfileObj(1, CountArrayProfileObj)
                            ArrayProfileObj(0, CountArrayProfileObj) = i
                            ArrayProfileObj(1, CountArrayProfileObj) = "PIKET"
                            CountArrayProfileObj += 1
                        Next i
                        'записываем последнее значение
                        ReDim Preserve ArrayProfileObj(1, CountArrayProfileObj)
                        ArrayProfileObj(0, CountArrayProfileObj) = ProfRezult.EndingStation
                        ArrayProfileObj(1, CountArrayProfileObj) = "PIKET"
                    End If
                End If
                'сортируем массив с пикетами
                If IsArray(ArrayProfileObj) = True Then
                    If ArrayProfileObj.GetUpperBound(1) > 0 Then
                        Dim BoolSortArray As Boolean = MathFunction.FuncSortStrArray(ArrayProfileObj, 0)
                    End If
                Else
                    MsgBox("Массив с пикетами пуст")
                    Exit Sub
                End If
                '========================================================================================================
                'Стиль метки
                Dim NameStyleLabel As String = FormProjectProfileWorkMark.ComboBox6.Text
                Dim IDLabelStyle As ObjectId = CivilDoc.Styles.LabelStyles.ProfileViewLabelStyles.StationElevationLabelStyles(0)
                If IsNothing(NameStyleLabel) = False Then
                    If NameStyleLabel.Trim.Length > 0 Then
                        IDLabelStyle = CivilDoc.Styles.LabelStyles.ProfileViewLabelStyles.StationElevationLabelStyles.Item(NameStyleLabel)
                    End If
                End If
                '===========================================================================================================
                'стиль маркера
                Dim NameStyleMarker As String = FormProjectProfileWorkMark.ComboBox4.Text
                Dim IDMarkerStyle As ObjectId = CivilDoc.Styles.MarkerStyles.Item(0)
                If IsNothing(NameStyleMarker) = False Then
                    If NameStyleMarker.Trim.Length > 0 Then
                        IDMarkerStyle = CivilDoc.Styles.MarkerStyles.Item(NameStyleMarker)
                    End If
                End If
                'начинаем перебор объектов
                Dim PrevPK As Double = -10000000
                For i As Double = 0 To ArrayProfileObj.GetUpperBound(1)
                    '1. берем пикет
                    Dim PKRedProfile As Double = ArrayProfileObj(0, i) 'пикет
                    Dim NameSegm As String = ArrayProfileObj(1, i) 'обозначение сегмента
                    'если такой пикет уже был то повторно его не обрабатываем
                    If PKRedProfile = PrevPK Then
                        Continue For
                    Else
                        PrevPK = PKRedProfile
                    End If
                    '2. Округляем пикет
                    Dim HRed As Double = 0 'высота красного профиля
                    Dim HBlack As Double = 0 'высота черного профиля
                    Dim HDop As Double = 0 'высота дополнительного профиля
                    '3.находим высоту красного профиля
                    Dim boolRedH As Boolean = False
                    Try
                        HRed = ProfileRed.ElevationAt(PKRedProfile)
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
                                HRed = ProfileRed.ElevationAt(PKRedProfile)
                                boolRedH = True
                            Catch ex As System.ArgumentException
                                Continue For
                            End Try
                        Else
                            Try
                                HRed = ProfileRed.ElevationAt(PKRedProfile)
                                boolRedH = True
                            Catch ex As System.ArgumentException
                                Continue For
                            End Try
                        End If
                    End If
                    '4.находим высоту черного профиля
                    Dim boolBlackH As Boolean = False
                    Try
                        HBlack = ProfileBlask.ElevationAt(PKRedProfile)
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
                                HBlack = ProfileBlask.ElevationAt(PKRedProfile)
                                boolBlackH = True
                            Catch ex As System.ArgumentException
                                Continue For
                            End Try
                        Else
                            Try
                                HBlack = ProfileBlask.ElevationAt(PKRedProfile)
                                boolRedH = True
                            Catch ex As System.ArgumentException
                                Continue For
                            End Try
                        End If
                    End If
                    '4.находим высоту дополнительного профиля
                    Dim boolDopH As Boolean = False
                    If IsNothing(ProfileDop) = False Then
                        Try
                            HDop = ProfileDop.ElevationAt(PKRedProfile)
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
                                    HDop = ProfileDop.ElevationAt(PKRedProfile)
                                    boolDopH = True
                                Catch ex As System.ArgumentException
                                    Continue For
                                End Try
                            Else
                                Try
                                    HBlack = ProfileDop.ElevationAt(PKRedProfile)
                                    boolDopH = True
                                Catch ex As System.ArgumentException
                                    Continue For
                                End Try
                            End If
                        End If
                    End If
                    '4. вычисляем дельту
                    Dim delta As Double = 0
                    delta = HRed - HBlack
                    'Dim delta1 As Double = 0
                    'If FormProjectProfileWorkMark.CheckBox7.Checked = True Then
                    '    delta1 = Math.Round((HRed - HDop), RoundZnH)
                    'End If
                    'Dim x As Double = 0
                    'Dim y As Double = 0
                    'ProfileWiev.FindXYAtStationAndElevation(PKRedProfile, HRed, x, y)
                    If delta < 0 Then
                        HRed = HRed - OtstupY 'если число отритцательное то записывам ниже красного профиля
                    Else
                        HRed = HRed + OtstupY
                    End If
                    'Dim Pt1 As Point3d = New Point3d(x, y, 0)
                    Try
                        Dim IDElevProfileLabel As ObjectId = StationElevationLabel.Create(ProfileWiev.Id, IDLabelStyle, IDMarkerStyle, PrevPK, HRed)
                        Dim labelprofile As StationElevationLabel = tr.GetObject(IDElevProfileLabel, OpenMode.ForWrite)
                        labelprofile.Profile1Id = ProfileRed.Id
                        labelprofile.Profile2Id = ProfileBlask.Id
                    Catch ex As System.ArgumentOutOfRangeException
                    End Try
                Next
            End If
            tr.Commit()
        End Using
    End Sub
End Class
