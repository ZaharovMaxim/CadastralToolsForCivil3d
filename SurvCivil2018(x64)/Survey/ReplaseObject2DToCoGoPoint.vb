Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Public Class ReplaseObject2DToCoGoPoint
    '=============================================================================================
    'замена объектов автокад на точки CIVIL (без высот)
    <CommandMethod("ReplaseObject2DToCoGoPoint")>
    Public Sub ReplaseObject2DToCoGoPoint()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim strDWGName As String = acDoc.Name
        Try
            Dim objBase As Object = Application.GetSystemVariable("DWGTITLED")
            If System.Convert.ToInt16(objBase) > 0 Then
                Dim obj1 As Object = Application.GetSystemVariable("DBMOD")
                If Not (System.Convert.ToInt16(obj1) = 0) Then
                    acDoc.Database.SaveAs(strDWGName, True, DwgVersion.Current, acDoc.Database.SecurityParameters)
                End If
            End If
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try
        'вызываем диалоговое окно
        Dim FormReplceObject As ReplaseObj2DToCoGoPoint = New ReplaseObj2DToCoGoPoint
        FormReplceObject.ShowDialog()
        If FormReplceObject.FormLoad1 = False Then Exit Sub
        'начальные переменные
        Dim ArrayPoint As Double(,) = Nothing 'массив для координат точек
        Dim CountPoint As Integer = 0 'счетчик блоков
        Dim ArrayText As Double(,) = Nothing 'массив для текстов
        Dim CountText As Double = 0 'счетчик текстов
        'определяем условие со слоем
        Dim NameLayer As String = ""
        If FormReplceObject.CheckBox6.Checked = True Then
            NameLayer = FormReplceObject.ComboBox2.Text.Trim
        End If
        'запрос на номер точки
        Dim NumPoint As Integer = FormReplceObject.NumericUpDown4.Value
        'запрос на минимальные и максимальные отметки
        Dim MinH As Double = Nothing
        Dim MaxH As Double = Nothing
        If FormReplceObject.CheckBox5.Checked = True Then
            MinH = FormReplceObject.NumericUpDown3.Value
            MaxH = FormReplceObject.NumericUpDown2.Value
        End If
        'нажата кнопка преобразовать все объекты
        If FormReplceObject.RadioButton3.Checked = True Then
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Dim PointsCogo As CogoPointCollection = CivilDoc.CogoPoints
                Dim ssetBlk As SelectionSet = Nothing
                Dim ssetText As SelectionSet = Nothing
                Dim ssetMText As SelectionSet = Nothing
                Dim CollectionObject As ObjectIdCollection = New ObjectIdCollection
                Dim CollectionText As ObjectIdCollection = New ObjectIdCollection
                Dim CollectionMText As ObjectIdCollection = New ObjectIdCollection
                'максимальное рассояние от точки до блока
                Dim Lenght As Double = FormReplceObject.NumericUpDown1.Value
                'составляем фильтр для объектов
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If FormReplceObject.CheckBox6.Checked = True Then
                                If Not (acEnt.Layer Like NameLayer) Then
                                    Continue For
                                End If
                            End If
                            If TypeOf acEnt Is DBPoint Then
                                If FormReplceObject.RadioButton5.Checked = True Then
                                    Dim obj As DBPoint = acEnt
                                    ReDim Preserve ArrayPoint(1, CountPoint)
                                    ArrayPoint(0, CountPoint) = obj.Position.X
                                    ArrayPoint(1, CountPoint) = obj.Position.Y
                                    CountPoint += 1
                                End If
                            ElseIf TypeOf acEnt Is Circle Then
                                If FormReplceObject.RadioButton6.Checked = True Then
                                    Dim obj As Circle = acEnt
                                    ReDim Preserve ArrayPoint(1, CountPoint)
                                    ArrayPoint(0, CountPoint) = obj.Center.X
                                    ArrayPoint(1, CountPoint) = obj.Center.Y
                                    CountPoint += 1
                                End If
                            ElseIf TypeOf acEnt Is BlockReference Then
                                If FormReplceObject.RadioButton7.Checked = True Then
                                    Dim obj As BlockReference = acEnt
                                    If obj.Name Like FormReplceObject.TextBox1.Text.Trim Then
                                        ReDim Preserve ArrayPoint(1, CountPoint)
                                        ArrayPoint(0, CountPoint) = obj.Position.X
                                        ArrayPoint(1, CountPoint) = obj.Position.Y
                                        CountPoint += 1
                                    End If
                                End If
                            ElseIf TypeOf acEnt Is Hatch Then
                                If FormReplceObject.RadioButton1.Checked = True Then
                                    Dim obj As Hatch = acEnt
                                    Dim ext As Extents3d = obj.GeometricExtents
                                    ReDim Preserve ArrayPoint(1, CountPoint)
                                    ArrayPoint(0, CountPoint) = (ext.MinPoint.X + ext.MaxPoint.X) / 2
                                    ArrayPoint(1, CountPoint) = (ext.MinPoint.Y + ext.MaxPoint.Y) / 2
                                    CountPoint += 1
                                End If
                            ElseIf TypeOf acEnt Is DBText Then
                                If FormReplceObject.CheckBox1.Checked = True Then 'And FormReplceObject.RadioButton2.Checked = True 
                                    Dim text As DBText = acEnt
                                    Dim TextStr As String = text.TextString
                                    Dim ext As Extents3d = text.GeometricExtents
                                    Dim posT1 As Integer = text.TextString.IndexOf(".")
                                    Dim posT2 As Integer = text.TextString.IndexOf(",")
                                    Dim D As String = 0
                                    Dim M As String = 0
                                    Dim H As Double = -999
                                    If posT1 <> -1 Then
                                        D = Mid(TextStr, 1, posT1) 'извлекаем число до запятой
                                        M = Mid(TextStr, posT1 + 2) 'извлекаем число после запятой
                                        If IsNumeric(D) = True And IsNumeric(M) Then
                                            Dim dec As Integer = 10 ^ M.Length
                                            H = D + M / dec
                                        End If
                                    ElseIf posT2 <> -1 Then
                                        D = Mid(TextStr, 1, posT2) 'извлекаем число до запятой
                                        M = Mid(TextStr, posT2 + 2) 'извлекаем число после запятой
                                        If IsNumeric(D) = True And IsNumeric(M) Then
                                            Dim dec As Integer = 10 ^ M.Length
                                            H = D + M / dec
                                        End If
                                    Else 'если число целое тогда
                                        If IsNumeric(TextStr) Then
                                            If FormReplceObject.CheckBox3.Checked = False Then
                                                H = Val(TextStr)
                                            Else
                                                Continue For
                                            End If
                                        End If
                                    End If
                                    'проверяем высоту на допуск
                                    If FormReplceObject.CheckBox5.Checked = True Then
                                        If H > MinH And H < MaxH Then
                                            ReDim Preserve ArrayText(3, CountText)
                                            ArrayText(0, CountText) = (ext.MinPoint.X + ext.MaxPoint.X) / 2
                                            ArrayText(1, CountText) = (ext.MinPoint.Y + ext.MaxPoint.Y) / 2
                                            ArrayText(2, CountText) = H
                                            ArrayText(3, CountText) = 0
                                            CountText += 1
                                        End If
                                    Else
                                        ReDim Preserve ArrayText(3, CountText)
                                        ArrayText(0, CountText) = (ext.MinPoint.X + ext.MaxPoint.X) / 2
                                        ArrayText(1, CountText) = (ext.MinPoint.Y + ext.MaxPoint.Y) / 2
                                        ArrayText(2, CountText) = H
                                        ArrayText(3, CountText) = 0
                                        CountText += 1
                                    End If
                                End If
                            ElseIf TypeOf acEnt Is mText Then
                                If FormReplceObject.CheckBox2.Checked = True Then 'And FormReplceObject.RadioButton2.Checked = True 
                                    Dim text As MText = acEnt
                                    Dim dbObjColl As DBObjectCollection = New DBObjectCollection
                                    text.Explode(dbObjColl)
                                    If dbObjColl.Count > 0 Then
                                        For Each dbobj2 As Autodesk.AutoCAD.DatabaseServices.DBObject In dbObjColl
                                            If TypeOf dbobj2 Is DBText Then
                                                Dim TempTextPKobj As DBText = dbobj2
                                                Dim TextStr As String = TempTextPKobj.TextString
                                                Dim ext As Extents3d = TempTextPKobj.GeometricExtents
                                                Dim posT1 As Integer = TextStr.IndexOf(".")
                                                Dim posT2 As Integer = TextStr.IndexOf(",")
                                                Dim D As String = 0
                                                Dim M As String = 0
                                                Dim H As Double = -999
                                                If posT1 <> -1 Then
                                                    D = Mid(TextStr, 1, posT1) 'извлекаем число до запятой
                                                    M = Mid(TextStr, posT1 + 2) 'извлекаем число после запятой
                                                    If IsNumeric(D) = True And IsNumeric(M) Then
                                                        Dim dec As Integer = 10 ^ M.Length
                                                        H = D + M / dec
                                                    End If
                                                ElseIf posT2 <> -1 Then
                                                    D = Mid(TextStr, 1, posT2) 'извлекаем число до запятой
                                                    M = Mid(TextStr, posT2 + 2) 'извлекаем число после запятой
                                                    If IsNumeric(D) = True And IsNumeric(M) Then
                                                        Dim dec As Integer = 10 ^ M.Length
                                                        H = D + M / dec
                                                    End If
                                                Else 'если число целое тогда
                                                    If IsNumeric(TextStr) Then
                                                        If FormReplceObject.CheckBox3.Checked = False Then
                                                            H = Val(TextStr)
                                                        Else
                                                            Continue For
                                                        End If
                                                    End If
                                                End If
                                                'проверяем высоту на допуск
                                                If FormReplceObject.CheckBox5.Checked = True Then
                                                    If H > MinH And H < MaxH Then
                                                        ReDim Preserve ArrayText(3, CountText)
                                                        ArrayText(0, CountText) = (ext.MinPoint.X + ext.MaxPoint.X) / 2
                                                        ArrayText(1, CountText) = (ext.MinPoint.Y + ext.MaxPoint.Y) / 2
                                                        ArrayText(2, CountText) = H
                                                        ArrayText(3, CountText) = 0
                                                        CountText += 1
                                                    End If
                                                Else
                                                    ReDim Preserve ArrayText(3, CountText)
                                                    ArrayText(0, CountText) = (ext.MinPoint.X + ext.MaxPoint.X) / 2
                                                    ArrayText(1, CountText) = (ext.MinPoint.Y + ext.MaxPoint.Y) / 2
                                                    ArrayText(2, CountText) = H
                                                    ArrayText(3, CountText) = 0
                                                    CountText += 1
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If IsArray(ArrayPoint) = False Then
                    MsgBox("Фильтр не выявил ни одного объекта!!! Макро прервано.")
                    Exit Sub
                End If
                If IsArray(ArrayText) = False Then
                    MsgBox("Фильтр не выявил ни одного текста!!! Макро прервано.")
                    Exit Sub
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\'
                'тело функции
                For i As Integer = 0 To ArrayPoint.GetUpperBound(1)
                    Dim CoordPoint1 As Point3d = New Point3d(0, 0, 0)
                    Dim Dist As Double = 1000000000
                    Dim l As Double 'длина вектора
                    Dim ptObj As Point2d = New Point2d(ArrayPoint(0, i), ArrayPoint(1, i))
                    Dim posArray As Integer = -1
                    For j As Integer = 0 To ArrayText.GetUpperBound(1)
                        Dim st As Integer = ArrayText(3, j)
                        If st = 1 Then
                            Continue For
                        End If
                        Dim ptText As Point2d = New Point2d(ArrayText(0, j), ArrayText(1, j))
                        l = ptObj.GetVectorTo(ptText).Length
                        If l = 0 Then
                            CoordPoint1 = New Point3d(ArrayPoint(0, i), ArrayPoint(1, i), ArrayText(2, j))
                            Dist = 0
                            posArray = j
                            Exit For
                        Else
                            If Dist > l Then
                                CoordPoint1 = New Point3d(ArrayPoint(0, i), ArrayPoint(1, i), ArrayText(2, j))
                                Dist = l
                                posArray = j
                            End If
                        End If
                    Next j
                    If Dist <= FormReplceObject.NumericUpDown1.Value Then
                        If CivilDoc.CogoPoints.Contains(NumPoint) = False Then
                            Dim PointCogoID As ObjectId = CivilDoc.CogoPoints.Add(CoordPoint1, FormReplceObject.ComboBox1.Text, False) 'вставка одиночной точки
                            Dim PointCOGO2 As CogoPoint = PointCogoID.GetObject(OpenMode.ForWrite)
                            Try
                                If PointCOGO2.PointNumber <> NumPoint Then
                                    PointCOGO2.PointNumber = NumPoint
                                End If
                            Catch ex As System.ArgumentException
                            End Try
                            NumPoint += 1
                            ArrayText(3, posArray) = 1
                        Else
                            NumPoint += 1
                            For j As Integer = 0 To CivilDoc.CogoPoints.Count
                                If CivilDoc.CogoPoints.Contains(NumPoint) = False Then
                                    Exit For
                                End If
                                NumPoint += 1
                            Next j
                            Dim PointCogoID As ObjectId = CivilDoc.CogoPoints.Add(CoordPoint1, FormReplceObject.ComboBox1.Text, False) 'вставка одиночной точки
                            Dim PointCOGO2 As CogoPoint = PointCogoID.GetObject(OpenMode.ForWrite)
                            Try
                                If PointCOGO2.PointNumber <> NumPoint Then
                                    PointCOGO2.PointNumber = NumPoint
                                End If
                            Catch ex As System.ArgumentException
                            End Try
                            NumPoint += 1
                            ArrayText(3, posArray) = 1
                        End If
                    End If
                Next i
                tr.Commit()
            End Using
        Else
            Dim PointsCogo As CogoPointCollection = CivilDoc.CogoPoints
            Dim resBlk As PromptEntityResult = Nothing
            Dim optBlk As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите блок, точку или окружность: ")
            Dim optText As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите отметку точки: ")
            Dim Pt As Point3d = New Point3d(0, 0, 0)
            Dim TextStr As String = Nothing
            Do
                Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
                Using tr
                    Try
                        optBlk.AllowNone = True
                        resBlk = ed.GetEntity(optBlk)
                        If resBlk.Status = PromptStatus.OK Then
                            Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(resBlk.ObjectId, OpenMode.ForRead)
                            If TypeOf obj Is BlockReference Then
                                Dim objSelect As BlockReference = obj
                                Pt = objSelect.Position
                            ElseIf TypeOf obj Is Circle Then
                                Dim objSelect As Circle = obj
                                Pt = objSelect.Center
                            ElseIf TypeOf obj Is DBPoint Then
                                Dim objSelect As DBPoint = obj
                                Pt = objSelect.Position
                            Else
                                MsgBox("Ошибка!!! Выберите блок, окружность или точку Автокад")
                                GoTo line2
                            End If
                            'выбираем текстовой объект
                            optText.AllowNone = True
                            resBlk = ed.GetEntity(optText)
                            Dim Point2 As Point3d = New Point3d(0, 0, 0)
                            If resBlk.Status = PromptStatus.OK Then
                                Dim obj1 As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(resBlk.ObjectId, OpenMode.ForRead)
                                If TypeOf obj1 Is DBText Then
                                    Dim objSelect As DBText = obj1
                                    TextStr = objSelect.TextString
                                    Point2 = objSelect.Position
                                ElseIf TypeOf obj1 Is MText Then
                                    Dim objSelect As MText = obj1
                                    TextStr = objSelect.Text
                                    Point2 = objSelect.Location
                                Else
                                    MsgBox("Ошибка!!! Выберите однострочный или многострочный тексты Автокад")
                                    GoTo line2
                                End If
                                Dim posT As Integer = TextStr.IndexOf(",")
                                If posT <> -1 Then
                                    Dim D As Integer = Val(Mid(TextStr, 1, posT))
                                    Dim M As Integer = Val(Mid(TextStr, posT + 2))
                                    TextStr = D & "." & M
                                End If
                                Dim Elev As Double = Val(TextStr)
                                Dim CoordPoint1 As Point3d = New Point3d(Pt.X, Pt.Y, Elev)
                                Dim PointCogoID As ObjectId = CivilDoc.CogoPoints.Add(CoordPoint1, False) 'вставка одиночной точки
                                Dim PointCOGO2 As CogoPoint = PointCogoID.GetObject(OpenMode.ForWrite)
                                PointCOGO2.RawDescription = FormReplceObject.ComboBox1.Text
                                PointCOGO2.PointNumber = NumPoint
                            Else
                                Exit Do
                            End If
                        Else
                            Exit Do
                        End If
                    Catch
                    End Try
                    tr.Commit()
                End Using
line2:
            Loop
        End If
    End Sub
End Class

