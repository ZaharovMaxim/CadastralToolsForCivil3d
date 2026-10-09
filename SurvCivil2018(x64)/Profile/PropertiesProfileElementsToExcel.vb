Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class PropertiesProfileElementsToExcel
    '============================================================================================================
    'Характеристика элементов профиля
    <CommandMethod("PropertiesProfileElementsToExcel")>
    Public Sub PropertiesProfileElementsToExcel()
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim docBase As Database = docCol.MdiActiveDocument.Database
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim Align As Autodesk.Civil.DatabaseServices.Alignment = Nothing
        Dim ArrayExcel As String(,) = Nothing
        Dim CountArrayExcel As Integer = 0
        Try
            Dim strDWGName As String = doc.Name
            Dim objBase As Object = Application.GetSystemVariable("DWGTITLED")
            If System.Convert.ToInt16(objBase) > 0 Then
                Dim obj1 As Object = Application.GetSystemVariable("DBMOD")
                If Not (System.Convert.ToInt16(obj1) = 0) Then
                    doc.Database.SaveAs(strDWGName, True, DwgVersion.Current, doc.Database.SecurityParameters)
                End If
            End If
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim resBlk As PromptEntityResult = Nothing
            Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите профиль: ")
            resBlk = m_Editor.GetEntity(opt)
            Dim pvId As ObjectId = resBlk.ObjectId
            Dim obj As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(pvId, OpenMode.ForRead)
            If Not (TypeOf obj Is Autodesk.Civil.DatabaseServices.Profile) Then
                MsgBox("Необходимо выбрать объект типа ПРОФИЛЬ. Макро прервано!!!")
                Exit Sub
            End If
            '=================================================================================
            'получаем доступ к линии профиля
            Dim PProfil As Autodesk.Civil.DatabaseServices.Profile = obj

            Try
                Dim idAlign As ObjectId = PProfil.AlignmentId
                If IsNothing(idAlign) = False Or idAlign.IsNull = False Then
                    Align = tr.GetObject(idAlign, OpenMode.ForRead)
                End If
            Catch ex As Exception
            End Try
            Dim ProfileEntColl As ProfileEntityCollection = PProfil.Entities
            If ProfileEntColl.Count > 0 Then
                For i As Integer = 0 To ProfileEntColl.Count - 1
                    Dim ent As ProfileEntity = ProfileEntColl.Item(i)
                    If ent.Length > 0 Then
                        If ent.EntityType = ProfileEntityType.Tangent Then
                            Dim TEnt As ProfileTangent = ent
                            ReDim Preserve ArrayExcel(16, CountArrayExcel)
                            ArrayExcel(0, CountArrayExcel) = TEnt.StartStation
                            ArrayExcel(1, CountArrayExcel) = TEnt.EndStation
                            ArrayExcel(2, CountArrayExcel) = TEnt.Length
                            ArrayExcel(3, CountArrayExcel) = TEnt.StartElevation
                            ArrayExcel(4, CountArrayExcel) = TEnt.EndElevation
                            ArrayExcel(5, CountArrayExcel) = (TEnt.EndElevation - TEnt.StartElevation) / TEnt.Length
                            If IsNothing(Align) = False Then
                                Dim X As Double = -1
                                Dim Y As Double = -1
                                Try
                                    Align.PointLocation(TEnt.StartStation, 0, X, Y)
                                    ArrayExcel(6, CountArrayExcel) = X
                                    ArrayExcel(7, CountArrayExcel) = Y
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                                X = -1
                                Y = -1
                                Try
                                    Align.PointLocation(TEnt.EndStation, 0, X, Y)
                                    ArrayExcel(8, CountArrayExcel) = X
                                    ArrayExcel(9, CountArrayExcel) = Y
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            End If
                            ArrayExcel(16, CountArrayExcel) = 0
                            CountArrayExcel += 1
                        ElseIf ent.EntityType = ProfileEntityType.Circular Then
                            Dim CEnt As ProfileCircular = ent
                            ReDim Preserve ArrayExcel(16, CountArrayExcel)
                            ArrayExcel(0, CountArrayExcel) = CEnt.StartStation
                            ArrayExcel(1, CountArrayExcel) = CEnt.EndStation
                            ArrayExcel(2, CountArrayExcel) = CEnt.Length
                            ArrayExcel(3, CountArrayExcel) = CEnt.StartElevation
                            ArrayExcel(4, CountArrayExcel) = CEnt.EndElevation
                            ArrayExcel(5, CountArrayExcel) = (CEnt.EndElevation - CEnt.StartElevation) / CEnt.Length
                            If IsNothing(Align) = False Then
                                Dim X As Double = -1
                                Dim Y As Double = -1
                                Try
                                    Align.PointLocation(CEnt.StartStation, 0, X, Y)
                                    ArrayExcel(6, CountArrayExcel) = X
                                    ArrayExcel(7, CountArrayExcel) = Y
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                                X = -1
                                Y = -1
                                Try
                                    Align.PointLocation(CEnt.EndStation, 0, X, Y)
                                    ArrayExcel(8, CountArrayExcel) = X
                                    ArrayExcel(9, CountArrayExcel) = Y
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            End If
                            If CEnt.CurveType = VerticalCurveType.Crest Then
                                ArrayExcel(10, CountArrayExcel) = 770 'выпуклая кривая
                                ArrayExcel(11, CountArrayExcel) = CEnt.HighLowPointElevation 'высота наивысшей точки
                                ArrayExcel(12, CountArrayExcel) = CEnt.HighLowPointStation 'пикет наивысшей точки
                            ElseIf CEnt.CurveType = VerticalCurveType.Sag Then
                                ArrayExcel(10, CountArrayExcel) = 769 'вогнутая кривая
                                ArrayExcel(11, CountArrayExcel) = CEnt.HighLowPointElevation 'высота наивысшей точки
                                ArrayExcel(12, CountArrayExcel) = CEnt.HighLowPointStation 'пикет наивысшей точки
                            End If
                            If IsNothing(Align) = False Then
                                Dim X As Double = -1
                                Dim Y As Double = -1
                                Try
                                    'координаты наивысшей точки
                                    Align.PointLocation(CEnt.HighLowPointStation, 0, X, Y)
                                    ArrayExcel(13, CountArrayExcel) = X
                                    ArrayExcel(14, CountArrayExcel) = Y
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            End If
                            ArrayExcel(15, CountArrayExcel) = CEnt.Radius 'радиус кривой
                            ArrayExcel(16, CountArrayExcel) = 1
                            CountArrayExcel += 1

                        ElseIf ent.EntityType = ProfileEntityType.ParabolaAsymmetric Then
                            Dim AEnt As ProfileParabolaAsymmetric = ent
                            ReDim Preserve ArrayExcel(16, CountArrayExcel)
                            ArrayExcel(0, CountArrayExcel) = AEnt.StartStation
                            ArrayExcel(1, CountArrayExcel) = AEnt.EndStation
                            Try
                                ArrayExcel(2, CountArrayExcel) = AEnt.AsymmetricLength1
                            Catch ex As System.ArgumentException
                            End Try
                            If ArrayExcel(2, CountArrayExcel) = 0 Then
                                Try
                                    ArrayExcel(2, CountArrayExcel) = AEnt.AsymmetricLength2
                                Catch ex As System.ArgumentException
                                End Try
                            End If
                            ArrayExcel(3, CountArrayExcel) = AEnt.StartElevation
                            ArrayExcel(4, CountArrayExcel) = AEnt.EndElevation
                            ArrayExcel(5, CountArrayExcel) = (AEnt.EndElevation - AEnt.StartElevation) / AEnt.Length
                            If IsNothing(Align) = False Then
                                Dim X As Double = -1
                                Dim Y As Double = -1
                                Try
                                    Align.PointLocation(AEnt.StartStation, 0, X, Y)
                                    ArrayExcel(6, CountArrayExcel) = X
                                    ArrayExcel(7, CountArrayExcel) = Y
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                                X = -1
                                Y = -1
                                Try
                                    Align.PointLocation(AEnt.EndStation, 0, X, Y)
                                    ArrayExcel(8, CountArrayExcel) = X
                                    ArrayExcel(9, CountArrayExcel) = Y
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            End If
                            If AEnt.CurveType = VerticalCurveType.Crest Then
                                ArrayExcel(10, CountArrayExcel) = 770 'выпуклая кривая
                                ArrayExcel(11, CountArrayExcel) = AEnt.HighLowPointElevation 'высота наивысшей точки
                                ArrayExcel(12, CountArrayExcel) = AEnt.HighLowPointStation 'пикет наивысшей точки
                            ElseIf AEnt.CurveType = VerticalCurveType.Sag Then
                                ArrayExcel(10, CountArrayExcel) = 769 'вогнутая кривая
                                ArrayExcel(11, CountArrayExcel) = AEnt.HighLowPointElevation 'высота наивысшей точки
                                ArrayExcel(12, CountArrayExcel) = AEnt.HighLowPointStation 'пикет наивысшей точки
                            End If
                            If IsNothing(Align) = False Then
                                Dim X As Double = -1
                                Dim Y As Double = -1
                                Try
                                    'координаты наивысшей точки
                                    Align.PointLocation(AEnt.HighLowPointStation, 0, X, Y)
                                    ArrayExcel(13, CountArrayExcel) = X
                                    ArrayExcel(14, CountArrayExcel) = Y
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            End If
                            ArrayExcel(16, CountArrayExcel) = 2
                            CountArrayExcel += 1
                        ElseIf ent.EntityType = ProfileEntityType.ParabolaSymmetric Then
                            Dim AEnt As ProfileParabolaSymmetric = ent
                            ReDim Preserve ArrayExcel(16, CountArrayExcel)
                            ArrayExcel(0, CountArrayExcel) = AEnt.StartStation
                            ArrayExcel(1, CountArrayExcel) = AEnt.EndStation
                            Try
                                ArrayExcel(2, CountArrayExcel) = AEnt.Length
                            Catch ex As System.ArgumentException
                            End Try
                            ArrayExcel(3, CountArrayExcel) = AEnt.StartElevation
                            ArrayExcel(4, CountArrayExcel) = AEnt.EndElevation
                            ArrayExcel(5, CountArrayExcel) = (AEnt.EndElevation - AEnt.StartElevation) / AEnt.Length
                            If IsNothing(Align) = False Then
                                Dim X As Double = -1
                                Dim Y As Double = -1
                                Try
                                    Align.PointLocation(AEnt.StartStation, 0, X, Y)
                                    ArrayExcel(6, CountArrayExcel) = X
                                    ArrayExcel(7, CountArrayExcel) = Y
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                                X = -1
                                Y = -1
                                Try
                                    Align.PointLocation(AEnt.EndStation, 0, X, Y)
                                    ArrayExcel(8, CountArrayExcel) = X
                                    ArrayExcel(9, CountArrayExcel) = Y
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            End If
                            If AEnt.CurveType = VerticalCurveType.Crest Then
                                ArrayExcel(10, CountArrayExcel) = 770 'выпуклая кривая
                                ArrayExcel(11, CountArrayExcel) = AEnt.HighLowPointElevation 'высота наивысшей точки
                                ArrayExcel(12, CountArrayExcel) = AEnt.HighLowPointStation 'пикет наивысшей точки
                            ElseIf AEnt.CurveType = VerticalCurveType.Sag Then
                                ArrayExcel(10, CountArrayExcel) = 769 'вогнутая кривая
                                ArrayExcel(11, CountArrayExcel) = AEnt.HighLowPointElevation 'высота наивысшей точки
                                ArrayExcel(12, CountArrayExcel) = AEnt.HighLowPointStation 'пикет наивысшей точки
                            End If
                            If IsNothing(Align) = False Then
                                Dim X As Double = -1
                                Dim Y As Double = -1
                                Try
                                    'координаты наивысшей точки
                                    Align.PointLocation(AEnt.HighLowPointStation, 0, X, Y)
                                    ArrayExcel(13, CountArrayExcel) = X
                                    ArrayExcel(14, CountArrayExcel) = Y
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            End If
                            ArrayExcel(15, CountArrayExcel) = AEnt.Radius 'радиус кривой
                            ArrayExcel(16, CountArrayExcel) = 3
                            CountArrayExcel += 1
                        End If
                    End If
                Next i
            End If
        End Using
        'сортируем массив
        If IsArray(ArrayExcel) = True Then
            Dim boolSortArray As Boolean = MathFunction.FuncSortStrArray(ArrayExcel, 0)
        End If
        ' запускаем excel
        Dim xlApp As Excel.Application = Nothing
        Dim xlBook As Excel.Workbook = Nothing
        Dim xlSheets As Excel.Worksheet = Nothing
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. путь к блокам
        Dim PutchTables As String = "d:\AutodeskCivilTools\TemplateTables\Roads\Ведомость элементов продольного профиля\" 'пути для шаблона таблицы
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\Roads\Ведомость элементов продольного профиля\") = True Then
            PutchTables = parentPatch2 & "\TemplateTables\Roads\Ведомость элементов продольного профиля\"
        Else
            If IO.Directory.Exists(PutchTables) = False Then
                MsgBox("Не удалось найти папку" & parentPatch2 & "\TemplateTables\Roads\Ведомость элементов продольного профиля\")
                GoTo line1
            End If
        End If
        Dim NameFile As String = PutchTables & "Ведомость элементов продольного профиля.xlsx"
        If IO.File.Exists(NameFile) = False Then
            MsgBox("Шаблон не найден в папке: " & parentPatch2 & "\TemplateTables\Roads\Ведомость элементов продольного профиля\")
            GoTo line1
        End If
        xlApp = CreateObject("Excel.Application")
        xlBook = xlApp.Workbooks.Open(NameFile)
        xlApp.Visible = True
        Dim xlsheet As Excel.Worksheet = xlBook.Worksheets.Item("1") 'делаем ссылку на этот лист
        xlsheet.Activate()
        If IsArray(ArrayExcel) = True Then
            Dim CountRow As Integer = 3
            For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                'записываем сырой пикет
                Dim PK As Double = ArrayExcel(0, i)
                xlsheet.Cells(CountRow, 1).value = PK
                'записываем форматированные пикеты
                Try
                    xlsheet.Cells(CountRow, 2).value = "'" & Align.GetStationStringWithEquations(PK)
                Catch ex As System.ArgumentException
                    xlsheet.Cells(CountRow, 2).value = PK
                End Try
                PK = ArrayExcel(1, i)
                Try
                    xlsheet.Cells(CountRow, 3).value = "'" & Align.GetStationStringWithEquations(PK)
                Catch ex As System.ArgumentException
                    xlsheet.Cells(CountRow, 3).value = PK
                End Try
                'записываем длину и радиус
                Dim L As Double = ArrayExcel(2, i)
                Dim R As Double = ArrayExcel(15, i)
                xlsheet.Cells(CountRow, 4).value = L
                xlsheet.Cells(CountRow, 5).value = R
                'записываем отметки начала и конца элементов
                Dim H1 As Double = ArrayExcel(3, i)
                Dim H2 As Double = ArrayExcel(4, i)
                xlsheet.Cells(CountRow, 6).value = H1
                xlsheet.Cells(CountRow, 7).value = H2
                'записываем уклон
                Dim i1 As Double = ArrayExcel(5, i)
                i1 = Math.Abs(i1 * 1000)
                xlsheet.Cells(CountRow, 8).value = i1
                'записываем координаты начала элемента
                Dim X1 As Double = ArrayExcel(6, i)
                Dim Y1 As Double = ArrayExcel(7, i)
                xlsheet.Cells(CountRow, 9).value = X1
                xlsheet.Cells(CountRow, 10).value = Y1
                'записываем координаты конца элемента
                Dim X2 As Double = ArrayExcel(8, i)
                Dim Y2 As Double = ArrayExcel(9, i)
                xlsheet.Cells(CountRow, 11).value = X2
                xlsheet.Cells(CountRow, 12).value = Y2
                'записываем тип кривой
                Dim ind As Integer = ArrayExcel(7, i)
                If ind = 770 Then
                    xlsheet.Cells(CountRow, 13).value = "' Выпуклая кривая"
                ElseIf ind = 769 Then
                    xlsheet.Cells(CountRow, 13).value = "' Вогнутая кривая"
                Else
                    xlsheet.Cells(CountRow, 13).value = "' Прямой участок"
                End If
                'записываем верх кривой (отметка)
                Dim H As Double = ArrayExcel(11, i)
                xlsheet.Cells(CountRow, 14).value = H
                'записываем верх кривой (пикет)
                PK = ArrayExcel(12, i)
                Try
                    xlsheet.Cells(CountRow, 15).value = "'" & Align.GetStationStringWithEquations(PK)
                Catch ex As System.ArgumentException
                    xlsheet.Cells(CountRow, 15).value = PK
                End Try
                'записываем координаты верха кривой
                Dim X As Double = ArrayExcel(13, i)
                Dim Y As Double = ArrayExcel(14, i)
                xlsheet.Cells(CountRow, 16).value = X
                xlsheet.Cells(CountRow, 17).value = Y
                'записываем тип элемента
                Dim indEnt As Integer = ArrayExcel(16, i)
                If indEnt = 0 Then
                    xlsheet.Cells(CountRow, 18).value = "' Прямая"
                ElseIf indEnt = 1 Then
                    xlsheet.Cells(CountRow, 18).value = "' Дуга"
                ElseIf indEnt = 2 Then
                    xlsheet.Cells(CountRow, 18).value = "' Ассиметричная парабола"
                ElseIf indEnt = 3 Then
                    xlsheet.Cells(CountRow, 18).value = "' Симметричная парабола"
                End If
                CountRow += 1
            Next i
        End If
line1:
        xlSheets = Nothing
        xlBook = Nothing
        xlApp.Quit()
        xlApp = Nothing
    End Sub
End Class
