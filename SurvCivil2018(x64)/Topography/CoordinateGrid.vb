Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Public Class CoordinateGrid
    '=====================================================================================================
    'рисует координатную сетку
    <CommandMethod("CoordinateGrid")>
    Public Sub CoordinateGrid()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
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
        Dim FormSelectTemlate As SelectTemplateFiles = New SelectTemplateFiles
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
                        ReDim Preserve ArrayTemplates(CountArrayTemplates)
                        ArrayTemplates(CountArrayTemplates) = Path.GetFileName(fileDWG)
                        CountArrayTemplates += 1
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
        If IsArray(ArrayTemplates) = True Then
            If ArrayTemplates.Length > 1 Then
                'вызываем диалоговое окно с выбором шаблона
                FormSelectTemlate.ListBox1.DataSource = ArrayTemplates
                FormSelectTemlate.ShowDialog()
                If FormSelectTemlate.boolShow = False Then
                    Exit Sub
                Else
                    PutchTemplateFile = PutchTables & FormSelectTemlate.ListBox1.Text
                End If
            Else
                PutchTemplateFile = PutchTables & ArrayTemplates(0)
            End If
        End If
        Dim ArrayPropBlk As String(,) = Nothing
        Dim NameBlk As String = "Пересечение координатных линий"
        Dim IndColor As Integer = 92
        Dim NameLayer As String = "0"
        Dim ScaleBlk As Double = 1
        Dim idTypeLine As ObjectId = Nothing
        Dim boolPropBlkUPK As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateFile, NameBlk, ArrayPropBlk, "Блок", "Математическая основа")
        If IsArray(ArrayPropBlk) = True Then
            For i As Integer = 0 To ArrayPropBlk.GetUpperBound(1)
                Dim NameField As String = ArrayPropBlk(0, i)
                If NameField.Trim Like "Name" Then
                    NameBlk = ArrayPropBlk(1, i)
                    If IsNothing(NameBlk) = False Then
                        If NameBlk.Trim.Length > 0 Then
                            Dim boolInsRefBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchTemplateFile, NameBlk)
                        End If
                    End If
                ElseIf NameField.Trim Like "Color" Then
                    IndColor = Val(ArrayPropBlk(1, i))
                ElseIf NameField.Trim Like "Layer" Then
                    Dim TempNameLayer As String = ArrayPropBlk(1, i)
                    Dim boolInsRefLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplateFile, TempNameLayer)
                    If boolInsRefLayer = True Then
                        NameLayer = TempNameLayer
                    End If
                ElseIf NameField.Trim Like "Scale" Then
                    ScaleBlk = Val(ArrayPropBlk(1, i))
                End If
            Next i
        Else
            MsgBox("Характеристики блока не найдены. Макро прервано!!!")
            Exit Sub
        End If
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Try
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'тело программы
                    Dim SelectPoint1 As PromptPointResult = ed.GetPoint("Укажите нижний левый угол сетки: ")
                    If SelectPoint1.Status = PromptStatus.Cancel Then
                        Throw New System.Exception("Прервано пользователем")
                        Exit Sub 'если нажали ESC то выходим из программы
                    End If
                    Dim SelectPoint2 As PromptPointResult = ed.GetPoint("Укажите верхний правый угол сетки: ")
                    If SelectPoint1.Status = PromptStatus.Cancel Then
                        Throw New System.Exception("Прервано пользователем")
                        Exit Sub 'если нажали ESC то выходим из программы
                    End If
                    'находим масштаб чертежа
                    Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & "Укажите масштаб чертежа (500, 1000, 2000 ...): ")
                    pStrOpts.AllowSpaces = True
                    Dim pStrRes As PromptResult = acDoc.Editor.GetString(pStrOpts)
                    Dim MDraw As Double = Val(pStrRes.StringResult)
                    If MDraw > 0 Then
                        MDraw = Math.Round(MDraw, 0)
                    End If
                    Dim k As Integer = MDraw / 10 'находим расстояние между крестами
                    'находим положение левого нижнего креста
                    Dim Point1 As Point3d = SelectPoint1.Value
                    Dim point2 As Point3d = SelectPoint2.Value
                    If (Point1.X > point2.X And Point1.Y > point2.Y) Then
                        Point1 = SelectPoint2.Value
                        point2 = SelectPoint1.Value
                    End If
                    Dim KX0 As Integer = Int(Point1.X / k) 'находим число полных клеток для нижнего левого угла
                    Dim KY0 As Integer = Int(Point1.Y / k)
                    Dim NX As Integer = KX0 * k
                    Dim NY As Integer = KY0 * k
                    'находим число клеток для противоположной точки
                    Dim KX1 As Integer = Int(point2.X / k) 'находим число полных клеток для нижнего левого угла
                    Dim KY1 As Integer = Int(point2.Y / k)
                    Dim NX1 As Integer = KX1 - KX0 + 1
                    Dim NY1 As Integer = KY1 - KY0 + 1
                    ScaleBlk = ScaleBlk * (MDraw / 1000)
                    For i As Integer = 0 To NX1
                        Dim X As Integer = Int(NX + i * k)
                        For j As Integer = 0 To NY1
                            Dim Y As Integer = Int(NY + j * k)
                            Dim InsPoint As Point3d = New Point3d(X, Y, 0)
                            Dim insBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(InsPoint, NameBlk, NameLayer, IndColor, ScaleBlk)
                        Next j
                    Next i
                Catch
                    MsgBox("Ошибка. Макро прервано!!!")
                End Try
                acTrans.Commit()
            End Using
        Catch ex As Exception
        End Try
    End Sub
End Class
