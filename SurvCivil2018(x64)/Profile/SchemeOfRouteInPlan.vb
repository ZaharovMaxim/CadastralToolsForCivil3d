Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.Runtime
Imports Autodesk.Civil.Settings
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AECC.Interop.Land
Imports Autodesk.AECC.Interop.Survey
Imports Autodesk.AECC.Interop.UiLand
Imports Autodesk.AECC.Interop.UiSurvey
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
'SchemeOfRouteInPlan
Public Class SchemeOfRouteInPlan
    '=============================================================================================
    'схема трассы в плане
    <CommandMethod("SchemeOfRouteInPlan")>
    Public Sub SurvSubPofileTable()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'диалоговое окно
        Dim DialogProfil As SitProfil = New SitProfil
        DialogProfil.ShowDialog()
        If DialogProfil.stbool = False Then Exit Sub
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
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
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim Count As Integer = 0
        Try
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                'выбор исходной трассы
                Dim resBlk As PromptEntityResult = Nothing
                Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите исходную трассу: ")
                resBlk = ed.GetEntity(opt)
                Dim objID As ObjectId = resBlk.ObjectId
                Dim align As Alignment = tr.GetObject(objID, OpenMode.ForWrite)
                'выясняем начальный пикет
                Dim StartPiket As Double = align.StartingStation
                'масштаб плана
                Dim M1 As Integer = DialogProfil.NumericUpDown1.Value
                'мосштаб профиля
                Dim M2 As Integer = DialogProfil.NumericUpDown2.Value
                'ширина рамки профиля
                Dim LRamka As Double = DialogProfil.NumericUpDown3.Value
                'вычисляем коеффициент увеличения блоков
                Dim k As Double = M2 / M1
                'коэффициент растяжения
                Dim kR As Double = DialogProfil.NumericUpDown4.Value
                'вычисляем ширину рамки
                Dim Dist As Double = LRamka * M2 / 100
                'выясням координаты начального пикета трассы2
                'создаем слой PROF-Продольный профиль (элементы плана)
                Dim LayerPP As String = DialogProfil.ComboBox1.Text
                Dim IDLayerPP As ObjectId = AcadFunction.FuncAddLayer(LayerPP)
                'точка вставки 
                Dim StartPoint As Point2d = Nothing
                Dim SelectPoint1 As PromptPointResult = ed.GetPoint("Укажите точку вставки: ")
                If SelectPoint1.Status = PromptStatus.Cancel Then
                    Throw New System.Exception("Прервано пользователем")
                    Exit Sub 'если нажали ESC то выходим из программы
                End If
                If StartPiket <> 0 Then
                    StartPoint = New Point2d(SelectPoint1.Value.X - align.StartingStation, SelectPoint1.Value.Y)
                Else
                    StartPoint = New Point2d(SelectPoint1.Value.X, SelectPoint1.Value.Y)
                End If

                Dim ProfPt1 As Point2d = New Point2d(SelectPoint1.Value.X, SelectPoint1.Value.Y + Dist / 2)
                Dim ProfPt2 As Point2d = New Point2d(SelectPoint1.Value.X + align.Length, SelectPoint1.Value.Y + Dist / 2)
                Dim ProfPt3 As Point2d = New Point2d(SelectPoint1.Value.X + align.Length, SelectPoint1.Value.Y - Dist / 2)
                Dim ProfPt4 As Point2d = New Point2d(SelectPoint1.Value.X, SelectPoint1.Value.Y - Dist / 2)
                'рисуем границу отсечки
                Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline()
                acPoly.SetDatabaseDefaults()
                acPoly.AddVertexAt(0, ProfPt1, 0, 0, 0)
                acPoly.AddVertexAt(1, ProfPt2, 0, 0, 0)
                acPoly.AddVertexAt(2, ProfPt3, 0, 0, 0)
                acPoly.AddVertexAt(3, ProfPt4, 0, 0, 0)
                acPoly.Closed = True
                acPoly.ColorIndex = 1
                acPoly.Linetype = "Continuous"
                acPoly.LineWeight = LineWeight.LineWeight030
                acPoly.Layer = LayerPP
                acBlkTblRec.AppendEntity(acPoly)
                tr.AddNewlyCreatedDBObject(acPoly, True)
                'выбор объектов
                Dim Count1 As Integer = 0
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        Count += 1
                        Count1 += 1
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'трансформирование полилинии
                                If (TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline) Then
                                    If DialogProfil.CheckBox2.Checked = True Then
                                        Dim PLObj As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt
                                        'записываем характеристики линии
                                        Dim IndexColor As Integer = PLObj.ColorIndex 'цвет
                                        Dim TypeLinePl As String = PLObj.Linetype 'тип линии
                                        Dim LineWightPl As LineWeight = PLObj.LineWeight 'ширина линии
                                        Dim ClosePl As Boolean = PLObj.Closed 'замкнута или нет
                                        Dim ScalePl As Double = PLObj.LinetypeScale 'маштаб типа линии
                                        'масштабирование линии
                                        If DialogProfil.CheckBox15.Checked = True Then
                                            If TypeLinePl Like "Continuous" Then
                                                ScalePl = 1
                                            Else
                                                ScalePl = ScalePl * k
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'начинаем перебор вершин
                                        Dim ArrayPline As Double(,) = Nothing 'массив для координат полилинии
                                        Dim countVert As Integer = 0 'счетчик вершин
                                        For i As Integer = 0 To PLObj.NumberOfVertices - 1
                                            Dim Point1 As Point2d = PLObj.GetPoint2dAt(i)
                                            ReDim Preserve ArrayPline(1, i)
                                            ArrayPline(0, i) = Point1.X
                                            ArrayPline(1, i) = Point1.Y
                                            countVert += 1
                                        Next i
                                        FuncDrawElementsPlateToTpofile(ArrayPline, Dist, StartPoint, align, acPoly, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl, kR)
                                        Erase ArrayPline
                                        countVert = 0
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование 2д полилиний
                                ElseIf (TypeOf acEnt Is Polyline2d) Then
                                    If DialogProfil.CheckBox3.Checked = True Then
                                        Dim PL2dObj As Polyline2d = acEnt
                                        'записываем характеристики линии в массив
                                        Dim IndexColor As Integer = PL2dObj.ColorIndex 'цвет
                                        Dim TypeLinePl As String = PL2dObj.Linetype 'тип линии
                                        Dim LineWightPl As LineWeight = PL2dObj.LineWeight 'ширина линии
                                        Dim ClosePl As Boolean = PL2dObj.Closed 'замкнута или нет
                                        Dim ScalePl As Double = PL2dObj.LinetypeScale 'маштаб типа линии
                                        'масштабирование линии
                                        If DialogProfil.CheckBox15.Checked = True Then
                                            If TypeLinePl Like "Continuous" Then
                                                ScalePl = 1
                                            Else
                                                ScalePl = ScalePl * k
                                            End If
                                        End If
                                        'начинаем перебор вершин
                                        Dim ArrayPline As Double(,) = Nothing 'массив для координат полилинии
                                        Dim countVert As Integer = -1 'счетчик вершин
                                        For Each VertID As ObjectId In PL2dObj
                                            Dim v2d As Vertex2d = tr.GetObject(VertID, OpenMode.ForWrite)
                                            Dim Point1 As Point3d = v2d.Position
                                            countVert += 1
                                            ReDim Preserve ArrayPline(1, countVert)
                                            ArrayPline(0, countVert) = Point1.X
                                            ArrayPline(1, countVert) = Point1.Y
                                        Next
                                        FuncDrawElementsPlateToTpofile(ArrayPline, Dist, StartPoint, align, acPoly, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl, kR)
                                        Erase ArrayPline
                                        countVert = 0
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование 3д полилиний
                                ElseIf (TypeOf acEnt Is Polyline3d) Then
                                    If DialogProfil.CheckBox4.Checked = True Then
                                        Dim PL2dObj As Polyline3d = acEnt
                                        'записываем характеристики линии в массив
                                        Dim IndexColor As Integer = PL2dObj.ColorIndex 'цвет
                                        Dim TypeLinePl As String = PL2dObj.Linetype 'тип линии
                                        Dim LineWightPl As LineWeight = PL2dObj.LineWeight 'ширина линии
                                        Dim ClosePl As Boolean = PL2dObj.Closed 'замкнута или нет
                                        Dim ScalePl As Double = PL2dObj.LinetypeScale 'маштаб типа линии
                                        'масштабирование линии
                                        If DialogProfil.CheckBox15.Checked = True Then
                                            If TypeLinePl Like "Continuous" Then
                                                ScalePl = 1
                                            Else
                                                ScalePl = ScalePl * k
                                            End If
                                        End If
                                        'начинаем перебор вершин
                                        Dim ArrayPline As Double(,) = Nothing 'массив для координат полилинии
                                        Dim countVert As Integer = -1 'счетчик вершин
                                        For Each VertID As ObjectId In PL2dObj
                                            Dim v2d As PolylineVertex3d = tr.GetObject(VertID, OpenMode.ForWrite)
                                            Dim Point1 As Point3d = v2d.Position
                                            countVert += 1
                                            ReDim Preserve ArrayPline(1, countVert)
                                            ArrayPline(0, countVert) = Point1.X
                                            ArrayPline(1, countVert) = Point1.Y
                                        Next
                                        FuncDrawElementsPlateToTpofile(ArrayPline, Dist, StartPoint, align, acPoly, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl, kR)
                                        Erase ArrayPline
                                        countVert = 0
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование линий
                                ElseIf (TypeOf acEnt Is Line) Then
                                    If DialogProfil.CheckBox1.Checked = True Then
                                        Dim LinObj As Line = acEnt
                                        'записываем характеристики линии в массив
                                        Dim IndexColor As Integer = LinObj.ColorIndex 'цвет
                                        Dim TypeLinePl As String = LinObj.Linetype 'тип линии
                                        Dim LineWightPl As LineWeight = LinObj.LineWeight 'ширина линии
                                        Dim ClosePl As Boolean = False 'замкнута или нет
                                        Dim ScalePl As Double = LinObj.LinetypeScale 'маштаб типа линии
                                        'масштабирование линии
                                        If DialogProfil.CheckBox15.Checked = True Then
                                            If TypeLinePl Like "Continuous" Then
                                                ScalePl = 1
                                            Else
                                                ScalePl = ScalePl * k
                                            End If
                                        End If
                                        'начинаем перебор вершин
                                        Dim ArrayPline As Double(,) = Nothing 'массив для координат полилинии
                                        ReDim Preserve ArrayPline(1, 1)
                                        ArrayPline(0, 0) = LinObj.StartPoint.X
                                        ArrayPline(1, 0) = LinObj.StartPoint.Y
                                        ArrayPline(0, 1) = LinObj.EndPoint.X
                                        ArrayPline(1, 1) = LinObj.EndPoint.Y
                                        FuncDrawElementsPlateToTpofile(ArrayPline, Dist, StartPoint, align, acPoly, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl, kR)
                                        Erase ArrayPline
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование сплайнов
                                ElseIf (TypeOf acEnt Is Spline) Then
                                    If DialogProfil.CheckBox5.Checked = True Then
                                        Dim SpLine As Spline = acEnt
                                        'записываем характеристики линии в массив
                                        Dim IndexColor As Integer = SpLine.ColorIndex 'цвет
                                        Dim TypeLinePl As String = SpLine.Linetype 'тип линии
                                        Dim LineWightPl As LineWeight = SpLine.LineWeight 'ширина линии
                                        Dim ClosePl As Boolean = SpLine.Closed 'замкнута или нет
                                        Dim ScalePl As Double = SpLine.LinetypeScale 'маштаб типа линии
                                        'масштабирование линии
                                        If DialogProfil.CheckBox15.Checked = True Then
                                            If TypeLinePl Like "Continuous" Then
                                                ScalePl = 1
                                            Else
                                                ScalePl = ScalePl * k
                                            End If
                                        End If
                                        Dim ArrayPline As Double(,) = Nothing 'массив для координат полилинии
                                        Dim Spl As Double = SpLine.EndParam - SpLine.StartParam
                                        Dim intVert As Integer = Int(Spl / 2)
                                        If intVert > 1 Then
                                            For i As Integer = 0 To intVert
                                                Dim DistSpl As Double = SpLine.StartParam + i * 2
                                                Dim CoordSp As Point3d = SpLine.GetPointAtDist(DistSpl)
                                                ReDim Preserve ArrayPline(1, i)
                                                ArrayPline(0, i) = CoordSp.X
                                                ArrayPline(1, i) = CoordSp.Y
                                            Next i
                                            ReDim Preserve ArrayPline(1, intVert + 1)
                                            ArrayPline(0, intVert + 1) = SpLine.EndPoint.X
                                            ArrayPline(1, intVert + 1) = SpLine.EndPoint.Y
                                        Else
                                            Dim lenghtSp As Double = Spl / 4
                                            For i As Integer = 0 To 3
                                                Dim DistSpl As Double = (i + 1) * lenghtSp
                                                Dim CoordSp As Point3d = SpLine.GetPointAtDist(DistSpl)
                                                ReDim Preserve ArrayPline(1, i)
                                                ArrayPline(0, i) = CoordSp.X
                                                ArrayPline(1, i) = CoordSp.Y
                                            Next i
                                        End If
                                        FuncDrawElementsPlateToTpofile(ArrayPline, Dist, StartPoint, align, acPoly, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl, kR)
                                        Erase ArrayPline
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование дуг
                                ElseIf (TypeOf acEnt Is Arc) Then
                                    If DialogProfil.CheckBox6.Checked = True Then
                                        Dim ObjArc As Arc = acEnt
                                        'записываем характеристики линии в массив
                                        Dim IndexColor As Integer = ObjArc.ColorIndex 'цвет
                                        Dim TypeLinePl As String = ObjArc.Linetype 'тип линии
                                        Dim LineWightPl As LineWeight = ObjArc.LineWeight 'ширина линии
                                        Dim ClosePl As Boolean = False 'замкнута или нет
                                        Dim ScalePl As Double = ObjArc.LinetypeScale 'маштаб типа линии
                                        'масштабирование линии
                                        If DialogProfil.CheckBox15.Checked = True Then
                                            If TypeLinePl Like "Continuous" Then
                                                ScalePl = 1
                                            Else
                                                ScalePl = ScalePl * k
                                            End If
                                        End If
                                        Dim ArrayPline As Double(,) = Nothing 'массив для координат полилинии
                                        Dim lgArc As Double = ObjArc.Length
                                        Dim intVert As Integer = Int(lgArc / 2)
                                        Dim DeltaArc As Double = Math.Round(lgArc / intVert, 3)
                                        If intVert > 3 Then
                                            For i As Integer = 0 To intVert - 1
                                                Dim DistSpl As Double = i * DeltaArc
                                                Dim CoordSp As Point3d = ObjArc.GetPointAtDist(DistSpl)
                                                ReDim Preserve ArrayPline(1, i)
                                                ArrayPline(0, i) = CoordSp.X
                                                ArrayPline(1, i) = CoordSp.Y
                                            Next i
                                            ReDim Preserve ArrayPline(1, intVert + 1)
                                            ArrayPline(0, intVert + 1) = ObjArc.EndPoint.X
                                            ArrayPline(1, intVert + 1) = ObjArc.EndPoint.Y
                                        Else
                                            'дуга слишком маленькая делим ее на 3 точки
                                            'начальная точки
                                            ReDim Preserve ArrayPline(1, 2)
                                            ArrayPline(0, 0) = ObjArc.StartPoint.X
                                            ArrayPline(1, 0) = ObjArc.StartPoint.Y
                                            'середина
                                            Dim CoordSp As Point3d = ObjArc.GetPointAtDist(lgArc / 2)
                                            ArrayPline(0, 1) = CoordSp.X
                                            ArrayPline(1, 1) = CoordSp.Y
                                            'конечная точка
                                            ArrayPline(0, 2) = ObjArc.EndPoint.X
                                            ArrayPline(1, 2) = ObjArc.EndPoint.Y
                                        End If
                                        FuncDrawElementsPlateToTpofile(ArrayPline, Dist, StartPoint, align, acPoly, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl, kR)
                                        Erase ArrayPline
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование окружностей
                                ElseIf (TypeOf acEnt Is Circle) Then
                                    If DialogProfil.CheckBox8.Checked = True Then
                                        Dim ObjCirc As Circle = acEnt
                                        'записываем характеристики линии в массив
                                        Dim IndexColor As Integer = ObjCirc.ColorIndex 'цвет
                                        Dim TypeLinePl As String = ObjCirc.Linetype 'тип линии
                                        Dim LineWightPl As LineWeight = ObjCirc.LineWeight 'ширина линии
                                        Dim ClosePl As Boolean = True 'замкнута или нет
                                        Dim ScalePl As Double = ObjCirc.LinetypeScale 'маштаб типа линии
                                        'масштабирование линии
                                        If DialogProfil.CheckBox15.Checked = True Then
                                            If TypeLinePl Like "Continuous" Then
                                                ScalePl = 1
                                            Else
                                                ScalePl = ScalePl * k
                                            End If
                                        End If
                                        Dim ArrayPline As Double(,) = Nothing 'массив для координат полилинии
                                        Dim lgCirc As Double = 2 * Math.PI * ObjCirc.Radius
                                        Dim intVert As Integer = Int(lgCirc / 2)
                                        Dim DeltaArc As Double = Math.Round(lgCirc / intVert, 3)
                                        If intVert > 3 Then
                                            For i As Integer = 0 To intVert - 1
                                                Dim DistSpl As Double = i * DeltaArc
                                                Dim CoordSp As Point3d = ObjCirc.GetPointAtDist(DistSpl)
                                                ReDim Preserve ArrayPline(1, i)
                                                ArrayPline(0, i) = CoordSp.X
                                                ArrayPline(1, i) = CoordSp.Y
                                            Next i
                                            ReDim Preserve ArrayPline(1, intVert + 1)
                                            ArrayPline(0, intVert + 1) = ObjCirc.EndPoint.X
                                            ArrayPline(1, intVert + 1) = ObjCirc.EndPoint.Y
                                        Else
                                            Dim lenghtSp As Double = lgCirc / 8
                                            For i As Integer = 0 To 7
                                                Dim DistSpl As Double = (i + 1) * lenghtSp
                                                Dim CoordSp As Point3d = ObjCirc.GetPointAtDist(DistSpl)
                                                ReDim Preserve ArrayPline(1, i)
                                                ArrayPline(0, i) = CoordSp.X
                                                ArrayPline(1, i) = CoordSp.Y
                                            Next i
                                        End If
                                        FuncDrawElementsPlateToTpofile(ArrayPline, Dist, StartPoint, align, acPoly, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl, kR)
                                        Erase ArrayPline
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование однострочных текстов
                                ElseIf (TypeOf acEnt Is DBText) Then
                                    If DialogProfil.CheckBox13.Checked = True Then
                                        Dim ObjText As DBText = acEnt
                                        Dim offsetRez1 As Boolean = False 'найден или нет пикетаж
                                        Dim Angle As Double = ObjText.Rotation 'угол поворота
                                        Dim ptIns As Point3d = ObjText.Position 'точка вставки
                                        Dim ptAlign As Point3d = ObjText.AlignmentPoint 'точка выравнивания
                                        Dim vMode As TextVerticalMode = ObjText.VerticalMode 'вертикальное выравнивание
                                        Dim hMode As TextHorizontalMode = ObjText.HorizontalMode 'горизонтальное выравнивание
                                        Dim ptIns2D As Point2d = Nothing
                                        'определяем габаритный контейнер
                                        'временно делаем поворот текста 0
                                        ObjText.Rotation = 0
                                        Dim mPoint As Point3d = ObjText.GeometricExtents.MinPoint
                                        Dim mxPoint As Point3d = ObjText.GeometricExtents.MaxPoint
                                        Dim pt1 As Point2d = New Point2d(mPoint.X, mPoint.Y)
                                        Dim pt2 As Point2d = New Point2d(mxPoint.X, mPoint.Y)
                                        Dim lehgt As Double = pt1.GetVectorTo(pt2).Length
                                        ObjText.Rotation = Angle
                                        'определяем координаты точки выравнивания
                                        Dim AligPointY As Double = ptIns.Y + lehgt * Math.Sin(Angle)
                                        Dim AligPointX As Double = ptIns.X + lehgt * Math.Cos(Angle)
                                        Dim posText As Point3d = New Point3d(0, 0, 0)
                                        Dim posAlign As Point2d = New Point2d(0, 0)
                                        'точка вставки
                                        Dim Piket1 As Double = 0
                                        Dim offsetAlign1 As Double = 0
                                        Try
                                            align.StationOffset(ptIns.X, ptIns.Y, Piket1, offsetAlign1)
                                            If Math.Abs(offsetAlign1) <= (Dist / 2) * kR Then
                                                offsetRez1 = True
                                                Dim x1 As Double = StartPoint.X + Piket1
                                                Dim y1 As Double = StartPoint.Y - offsetAlign1 / kR
                                                posText = New Point3d(x1, y1, 0)
                                                ptIns2D = New Point2d(x1, y1)
                                            Else
                                                GoTo Line1
                                            End If
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            offsetRez1 = False
                                            GoTo Line1
                                        End Try
                                        'точка выравнивания
                                        Dim Piket2 As Double = 0
                                        Dim offsetAlign2 As Double = 0
                                        If Angle <> 0 Then
                                            Try
                                                align.StationOffset(AligPointX, AligPointY, Piket2, offsetAlign2)
                                                Dim x2 As Double = StartPoint.X + Piket2
                                                Dim y2 As Double = StartPoint.Y - offsetAlign2 / kR
                                                posAlign = New Point2d(x2, y2)
                                                Angle = ptIns2D.GetVectorTo(posAlign).Angle
                                                Select Case Angle
                                                    Case Math.PI / 2 To (3 * Math.PI) / 2
                                                        Angle = Angle + Math.PI
                                                        hMode = TextHorizontalMode.TextRight
                                                        'posText = New Point3d(posAlign.X, posAlign.Y, 0)
                                                End Select
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                GoTo Line2
                                            End Try
                                        End If

Line2:
                                        Dim acText As DBText = ObjText.Clone()
                                        acText.Position = posText
                                        acText.HorizontalMode = hMode
                                        acText.VerticalMode = vMode
                                        acText.Rotation = Angle
                                        acText.Layer = LayerPP
                                        'масштабирование текстов
                                        If DialogProfil.CheckBox17.Checked = True Then
                                            acText.Height = acText.Height * k
                                        End If
                                        acBlkTblRec.AppendEntity(acText)
                                        tr.AddNewlyCreatedDBObject(acText, True)
Line1:
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование многострочных текстов
                                ElseIf (TypeOf acEnt Is MText) Then
                                    If DialogProfil.CheckBox14.Checked = True Then
                                        Dim ObjMText As MText = acEnt
                                        Dim offsetRez1 As Boolean = False 'найден или нет пикетаж
                                        Dim Angle As Double = ObjMText.Rotation 'угол поворота
                                        Dim ptIns As Point3d = ObjMText.Location 'точка вставки
                                        Dim ptIns2D As Point2d = Nothing
                                        'определяем габаритный контейнер
                                        'временно делаем поворот текста 0
                                        ObjMText.Rotation = 0
                                        Dim mPoint As Point3d = ObjMText.GeometricExtents.MinPoint
                                        Dim mxPoint As Point3d = ObjMText.GeometricExtents.MaxPoint
                                        Dim lehgt As Double = mPoint.GetVectorTo(mxPoint).Length
                                        ObjMText.Rotation = Angle
                                        'определяем координаты точки выравнивания
                                        Dim AligPointY As Double = ptIns.Y + lehgt * Math.Sin(Angle)
                                        Dim AligPointX As Double = ptIns.X + lehgt * Math.Cos(Angle)
                                        Dim posText As Point3d = Nothing
                                        Dim posAlign As Point2d = Nothing
                                        'точка вставки
                                        Dim Piket1 As Double = 0
                                        Dim offsetAlign1 As Double = 0
                                        Try
                                            align.StationOffset(ptIns.X, ptIns.Y, Piket1, offsetAlign1)
                                            If Math.Abs(offsetAlign1) <= (Dist / 2) * kR Then
                                                offsetRez1 = True
                                                Dim x1 As Double = StartPoint.X + Piket1
                                                Dim y1 As Double = StartPoint.Y - offsetAlign1 / kR
                                                posText = New Point3d(x1, y1, 0)
                                                ptIns2D = New Point2d(x1, y1)
                                            Else
                                                GoTo Line3
                                            End If
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            offsetRez1 = False
                                            GoTo Line3
                                        End Try
                                        'точка выравнивания
                                        Dim Piket2 As Double = 0
                                        Dim offsetAlign2 As Double = 0
                                        If Angle <> 0 Then
                                            Try
                                                align.StationOffset(AligPointX, AligPointY, Piket2, offsetAlign2)
                                                Dim x2 As Double = StartPoint.X + Piket2
                                                Dim y2 As Double = StartPoint.Y - offsetAlign2 / kR
                                                posAlign = New Point2d(x2, y2)
                                                Angle = ptIns2D.GetVectorTo(posAlign).Angle
                                                Select Case Angle
                                                    Case Math.PI / 2 To (3 * Math.PI) / 2
                                                        Angle = Angle + Math.PI
                                                        posText = New Point3d(posAlign.X, posAlign.Y, 0)
                                                End Select
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                GoTo Line4
                                            End Try
                                        End If

Line4:
                                        Dim acMText As MText = ObjMText.Clone()
                                        acMText.Location = posText
                                        acMText.Rotation = Angle
                                        acMText.Layer = LayerPP
                                        'масштабирование текстов
                                        If DialogProfil.CheckBox17.Checked = True Then
                                            acMText.Height = acMText.Height * k
                                        End If
                                        acBlkTblRec.AppendEntity(acMText)
                                        tr.AddNewlyCreatedDBObject(acMText, True)
Line3:
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование блоков
                                ElseIf (TypeOf acEnt Is BlockReference) Then
                                    If DialogProfil.CheckBox12.Checked = True Then
                                        Dim ObjBlk As BlockReference = acEnt
                                        Dim Angle As Double = ObjBlk.Rotation 'угол поворота
                                        Dim ptIns As Point3d = ObjBlk.Position 'точка вставки
                                        If ObjBlk.IsDynamicBlock = False Then
                                            If Angle = 0 Then
                                                Dim Piket0 As Double = 0
                                                Dim offsetAlign0 As Double = 0
                                                Try
                                                    align.StationOffset(ptIns.X, ptIns.Y, Piket0, offsetAlign0)
                                                    If Math.Abs(offsetAlign0) <= (Dist / 2) * kR Then
                                                        Dim x1 As Double = StartPoint.X + Piket0
                                                        Dim y1 As Double = StartPoint.Y - offsetAlign0 / kR
                                                        Dim posBlk0 As Point3d = New Point3d(x1, y1, 0)
                                                        Dim ObjBlk2 As BlockReference = ObjBlk.Clone()
                                                        ObjBlk2.Position = posBlk0
                                                        ObjBlk2.Layer = LayerPP
                                                        'масштабирование блока
                                                        If DialogProfil.CheckBox16.Checked = True Then
                                                            ObjBlk2.ScaleFactors = ObjBlk2.ScaleFactors * k
                                                        End If
                                                        acBlkTblRec.AppendEntity(ObjBlk2)
                                                        tr.AddNewlyCreatedDBObject(ObjBlk2, True)
                                                    Else
                                                        GoTo Line7
                                                    End If
                                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                    GoTo Line7
                                                End Try
                                            Else
                                                Dim offsetRez1 As Boolean = False 'найден или нет пикетаж
                                                Dim ptIns2D As Point2d = Nothing
                                                'определяем габаритный контейнер
                                                'временно делаем поворот текста 0
                                                ObjBlk.Rotation = 0
                                                Dim mPoint As Point3d = ObjBlk.GeometricExtents.MinPoint
                                                Dim mxPoint As Point3d = ObjBlk.GeometricExtents.MaxPoint
                                                Dim lehgt As Double = mPoint.GetVectorTo(mxPoint).Length
                                                ObjBlk.Rotation = Angle
                                                'определяем координаты точки выравнивания
                                                Dim AligPointY As Double = ptIns.Y + lehgt * Math.Sin(Angle)
                                                Dim AligPointX As Double = ptIns.X + lehgt * Math.Cos(Angle)
                                                Dim posBlk As Point3d = Nothing
                                                Dim posAlign As Point2d = Nothing
                                                'точка вставки
                                                Dim Piket1 As Double = 0
                                                Dim offsetAlign1 As Double = 0
                                                Try
                                                    align.StationOffset(ptIns.X, ptIns.Y, Piket1, offsetAlign1)
                                                    If Math.Abs(offsetAlign1) <= (Dist / 2) * kR Then
                                                        offsetRez1 = True
                                                        Dim x1 As Double = StartPoint.X + Piket1
                                                        Dim y1 As Double = StartPoint.Y - offsetAlign1 / kR
                                                        posBlk = New Point3d(x1, y1, 0)
                                                        ptIns2D = New Point2d(x1, y1)
                                                    Else
                                                        GoTo Line7
                                                    End If
                                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                    offsetRez1 = False
                                                    GoTo Line7
                                                End Try
                                                'точка выравнивания
                                                Dim Piket2 As Double = 0
                                                Dim offsetAlign2 As Double = 0
                                                If Angle <> 0 Then
                                                    Try
                                                        align.StationOffset(AligPointX, AligPointY, Piket2, offsetAlign2)
                                                        Dim x2 As Double = StartPoint.X + Piket2
                                                        Dim y2 As Double = StartPoint.Y - offsetAlign2 / kR
                                                        posAlign = New Point2d(x2, y2)
                                                        Angle = ptIns2D.GetVectorTo(posAlign).Angle
                                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                        GoTo line6
                                                    End Try
                                                End If
line6:
                                                Dim ObjBlk2 As BlockReference = ObjBlk.Clone()
                                                ObjBlk2.Position = posBlk
                                                ObjBlk2.Rotation = Angle
                                                ObjBlk2.Layer = LayerPP
                                                'масштабирование блока
                                                If DialogProfil.CheckBox16.Checked = True Then
                                                    ObjBlk2.ScaleFactors = ObjBlk2.ScaleFactors * k
                                                End If
                                                acBlkTblRec.AppendEntity(ObjBlk2)
                                                tr.AddNewlyCreatedDBObject(ObjBlk2, True)
                                            End If
                                        Else
                                            Dim ptIns2D As Point2d = Nothing
                                            Dim blockDyn As BlockTableRecord = Nothing
                                            Dim blockRef As BlockReference = tr.GetObject(acEnt.ObjectId, OpenMode.ForRead)
                                            blockDyn = tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead)
                                            'забираем в коллекцию все параметры объекта
                                            Dim blkProp As DynamicBlockReferencePropertyCollection = ObjBlk.DynamicBlockReferencePropertyCollection
                                            Dim AngParam As Double = Nothing 'угол доворота динамического блока
                                            For Each obj As DynamicBlockReferenceProperty In blkProp
                                                Dim UnitParam As DynamicBlockReferencePropertyUnitsType = obj.UnitsType
                                                If UnitParam = DynamicBlockReferencePropertyUnitsType.Angular Then
                                                    AngParam = obj.Value
                                                    Angle = Angle + AngParam
                                                    Exit For
                                                End If
                                            Next
                                            Dim mPoint As Point3d = ObjBlk.GeometricExtents.MinPoint
                                            Dim mxPoint As Point3d = ObjBlk.GeometricExtents.MaxPoint
                                            Dim lehgt As Double = mPoint.GetVectorTo(mxPoint).Length
                                            'определяем координаты точки выравнивания
                                            Dim AligPointY As Double = ptIns.Y + lehgt * Math.Sin(Angle)
                                            Dim AligPointX As Double = ptIns.X + lehgt * Math.Cos(Angle)
                                            Dim posBlk As Point3d = Nothing
                                            Dim posAlign As Point2d = Nothing
                                            'точка вставки
                                            Dim Piket1 As Double = 0
                                            Dim offsetAlign1 As Double = 0
                                            Try
                                                align.StationOffset(ptIns.X, ptIns.Y, Piket1, offsetAlign1)
                                                If Math.Abs(offsetAlign1) <= (Dist / 2) * kR Then
                                                    Dim x1 As Double = StartPoint.X + Piket1
                                                    Dim y1 As Double = StartPoint.Y - offsetAlign1 / kR
                                                    posBlk = New Point3d(x1, y1, 0)
                                                    ptIns2D = New Point2d(x1, y1)
                                                Else
                                                    GoTo Line7
                                                End If
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                GoTo Line7
                                            End Try
                                            'точка выравнивания
                                            Dim Piket2 As Double = 0
                                            Dim offsetAlign2 As Double = 0
                                            If Angle <> 0 Then
                                                Try
                                                    align.StationOffset(AligPointX, AligPointY, Piket2, offsetAlign2)
                                                    Dim x2 As Double = StartPoint.X + Piket2
                                                    Dim y2 As Double = StartPoint.Y - offsetAlign2 / kR
                                                    posAlign = New Point2d(x2, y2)
                                                    Angle = (ptIns2D.GetVectorTo(posAlign).Angle)
                                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                    GoTo Line8
                                                End Try
                                            End If

Line8:
                                            Dim ObjBlk2 As BlockReference = ObjBlk.Clone()
                                            ObjBlk2.Position = posBlk
                                            ObjBlk2.Rotation = Angle
                                            If AngParam <> 0 Then
                                                Dim blockDyn2 As BlockTableRecord = tr.GetObject(ObjBlk2.DynamicBlockTableRecord, OpenMode.ForWrite)
                                                Dim blkProp2 As DynamicBlockReferencePropertyCollection = ObjBlk2.DynamicBlockReferencePropertyCollection
                                                For Each obj As DynamicBlockReferenceProperty In blkProp2
                                                    Dim UnitParam As DynamicBlockReferencePropertyUnitsType = obj.UnitsType
                                                    If UnitParam = DynamicBlockReferencePropertyUnitsType.Angular Then
                                                        Dim ang2 As Double = obj.Value
                                                        Dim anIst As Double = Angle - ang2
                                                        ObjBlk2.Rotation = anIst
                                                        'obj.Value = 0
                                                        Exit For
                                                    End If
                                                Next
                                            End If
                                            'масштабирование блока
                                            If DialogProfil.CheckBox16.Checked = True Then
                                                ObjBlk2.ScaleFactors = ObjBlk2.ScaleFactors * k
                                            End If
                                            ObjBlk2.Layer = LayerPP
                                            acBlkTblRec.AppendEntity(ObjBlk2)
                                            tr.AddNewlyCreatedDBObject(ObjBlk2, True)
                                        End If

Line7:
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование кого точек
                                ElseIf (TypeOf acEnt Is CogoPoint) Then
                                    If DialogProfil.CheckBox10.Checked = True Then
                                        Dim objCOGO As CogoPoint = acEnt
                                        Dim ptCogo As Point3d = objCOGO.Location
                                        Dim PosBlk As Point3d = Nothing
                                        Dim Piket1 As Double = 0
                                        Dim offsetAlign1 As Double = 0
                                        Try
                                            align.StationOffset(ptCogo.X, ptCogo.Y, Piket1, offsetAlign1)
                                            If Math.Abs(offsetAlign1) <= (Dist / 2) * kR Then
                                                Dim x1 As Double = StartPoint.X + Piket1
                                                Dim y1 As Double = StartPoint.Y - offsetAlign1 / kR
                                                PosBlk = New Point3d(x1, y1, 0)
                                            Else
                                                GoTo line10
                                            End If
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            GoTo line10
                                        End Try
                                        Dim acDBObjColl As DBObjectCollection = New DBObjectCollection()
                                        acEnt.Explode(acDBObjColl)
                                        For Each acEntExplode As Autodesk.AutoCAD.DatabaseServices.Entity In acDBObjColl
                                            If TypeOf acEntExplode Is BlockReference Then
                                                Dim blk As BlockReference = acEntExplode
                                                Dim blk2 As BlockReference = blk.Clone()
                                                blk2.Position = PosBlk
                                                blk2.Rotation = 0
                                                blk2.Layer = LayerPP
                                                'масштабирование блока
                                                If DialogProfil.CheckBox16.Checked = True Then
                                                    blk2.ScaleFactors = blk2.ScaleFactors * k
                                                End If
                                                acBlkTblRec.AppendEntity(blk2)
                                                tr.AddNewlyCreatedDBObject(blk2, True)
                                            End If
                                        Next
line10:
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование меток
                                ElseIf (TypeOf acEnt Is Label) Then
                                    If DialogProfil.CheckBox9.Checked = True Then
                                        Dim objLabel As Label = acEnt
                                        Dim posLabel As Point3d = objLabel.LabelLocation
                                        Dim AngleLabel As Double = objLabel.RotationAngle
                                        Dim Piket1 As Double = Nothing
                                        Dim offsetAlign1 As Double = Nothing
                                        Dim PosStartPointFrof As Point2d = Nothing 'координаты точки вставки метки на профиле
                                        Dim PosAlignPointProf As Point2d = Nothing 'координаты точки выравнивания на профиле
                                        If AngleLabel <> 0 Then
                                            objLabel.RotationAngle = 0
                                            Dim mPoint As Point3d = objLabel.GeometricExtents.MinPoint
                                            Dim mxPoint As Point3d = objLabel.GeometricExtents.MaxPoint
                                            Dim pt1 As Point2d = New Point2d(mPoint.X, mPoint.Y)
                                            Dim pt2 As Point2d = New Point2d(mxPoint.X, mPoint.Y)
                                            Dim lehgt As Double = pt1.GetVectorTo(pt2).Length
                                            objLabel.RotationAngle = AngleLabel
                                            'определяем координаты точки выравнивания
                                            Dim AligPointY As Double = posLabel.Y + lehgt * Math.Sin(AngleLabel)
                                            Dim AligPointX As Double = posLabel.X + lehgt * Math.Cos(AngleLabel)
                                            'определяем угол доворота
                                            Try
                                                'точка вставки метки (ккординаты на профиле)
                                                align.StationOffset(posLabel.X, posLabel.Y, Piket1, offsetAlign1)
                                                If Math.Abs(offsetAlign1) <= (Dist / 2) * kR Then
                                                    Dim x1 As Double = StartPoint.X + Piket1
                                                    Dim y1 As Double = StartPoint.Y - offsetAlign1 / kR
                                                    PosStartPointFrof = New Point2d(x1, y1)
                                                    Piket1 = Nothing
                                                    offsetAlign1 = Nothing
                                                Else
                                                    GoTo line11
                                                End If
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                GoTo line11
                                            End Try
                                            Try
                                                align.StationOffset(AligPointX, AligPointY, Piket1, offsetAlign1)
                                                Dim x2 As Double = StartPoint.X + Piket1
                                                Dim y2 As Double = StartPoint.Y - offsetAlign1 / kR
                                                PosAlignPointProf = New Point2d(x2, y2)
                                                Dim AngleLineProf As Double = PosStartPointFrof.GetVectorTo(PosAlignPointProf).Angle
                                                AngleLabel = (AngleLineProf)
                                                Piket1 = Nothing
                                                offsetAlign1 = Nothing
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException

                                            End Try
                                        End If
                                        'расчленяем метку 1 раз
                                        Dim acDBObjColl1 As DBObjectCollection = New DBObjectCollection()
                                        Dim BlkExplode As BlockReference = Nothing
                                        acEnt.Explode(acDBObjColl1)
                                        For Each acEntExplode As Autodesk.AutoCAD.DatabaseServices.Entity In acDBObjColl1
                                            If (TypeOf acEntExplode Is BlockReference) Then
                                                BlkExplode = acEntExplode
                                                Exit For
                                            End If
                                        Next
                                        'расчленяем метку 2 раз
                                        Dim acDBObjColl2 As DBObjectCollection = New DBObjectCollection()
                                        If IsNothing(BlkExplode) = False Then
                                            BlkExplode.Explode(acDBObjColl2)
                                            For Each acEntExplode As Autodesk.AutoCAD.DatabaseServices.Entity In acDBObjColl2
                                                If (TypeOf acEntExplode Is Line) Then
                                                    Dim LineExplode As Line = acEntExplode
                                                    Dim ArrayPline(,) As Double = {{LineExplode.StartPoint.X, LineExplode.EndPoint.X}, {LineExplode.StartPoint.Y, LineExplode.EndPoint.Y}}
                                                    FuncDrawElementsPlateToTpofile(ArrayPline, Dist, StartPoint, align, acPoly, LayerPP, LineExplode.ColorIndex, LineExplode.LineWeight, False, LineExplode.LinetypeScale, LineExplode.Linetype)
                                                    Erase ArrayPline
                                                ElseIf (TypeOf acEntExplode Is BlockReference) Then
                                                    Dim blkExplode1 As BlockReference = acEntExplode
                                                    Try
                                                        'точка вставки метки (ккординаты на профиле)
                                                        align.StationOffset(blkExplode1.Position.X, blkExplode1.Position.Y, Piket1, offsetAlign1)
                                                        If Math.Abs(offsetAlign1) <= (Dist / 2) * kR Then
                                                            Dim x1 As Double = StartPoint.X + Piket1
                                                            Dim y1 As Double = StartPoint.Y - offsetAlign1 / kR
                                                            Dim posBlkProf1 As Point3d = New Point3d(x1, y1, 0)
                                                            Piket1 = Nothing
                                                            offsetAlign1 = Nothing
                                                            Dim blkExplode2 As BlockReference = blkExplode1.Clone()
                                                            blkExplode2.Position = posBlkProf1
                                                            blkExplode2.Rotation = AngleLabel
                                                            blkExplode2.Layer = LayerPP
                                                            acBlkTblRec.AppendEntity(blkExplode2)
                                                            tr.AddNewlyCreatedDBObject(blkExplode2, True)
                                                        Else
                                                            GoTo line12
                                                        End If
                                                    Catch ex As Autodesk.Civil.PointNotOnEntityException

                                                    End Try
                                                ElseIf (TypeOf acEntExplode Is MText) Then
                                                    Dim MTextExplode1 As MText = acEntExplode
                                                    Try
                                                        'точка вставки метки (ккординаты на профиле)
                                                        align.StationOffset(MTextExplode1.Location.X, MTextExplode1.Location.Y, Piket1, offsetAlign1)
                                                        If Math.Abs(offsetAlign1) <= (Dist / 2) * kR Then
                                                            Dim x1 As Double = StartPoint.X + Piket1
                                                            Dim y1 As Double = StartPoint.Y - offsetAlign1 / kR
                                                            Dim posBlkProf1 As Point3d = New Point3d(x1, y1, 0)
                                                            Piket1 = Nothing
                                                            offsetAlign1 = Nothing
                                                            Dim MTextExplode2 As MText = MTextExplode1.Clone()
                                                            MTextExplode2.Location = posBlkProf1
                                                            Dim mtA As Double = MTextExplode2.Rotation
                                                            MTextExplode2.Rotation = AngleLabel
                                                            Select Case AngleLabel
                                                                Case Math.PI / 2 To (3 * Math.PI) / 2
                                                                    MTextExplode2.Attachment = AttachmentPoint.MiddleCenter
                                                                    MTextExplode2.Rotation = AngleLabel + Math.PI
                                                            End Select
                                                            MTextExplode2.Layer = LayerPP
                                                            acBlkTblRec.AppendEntity(MTextExplode2)
                                                            tr.AddNewlyCreatedDBObject(MTextExplode2, True)
                                                        Else
                                                            GoTo line12
                                                        End If
                                                    Catch ex As Autodesk.Civil.PointNotOnEntityException

                                                    End Try
                                                End If
line12:
                                            Next
                                        End If

line11:
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование труб
                                ElseIf (TypeOf acEnt Is Pipe) Then
                                    If DialogProfil.CheckBox11.Checked = True Then
                                        Dim ObjPipe As Pipe = acEnt
                                        'расчленяем метку 1 раз
                                        Dim acDBObjColl1 As DBObjectCollection = New DBObjectCollection()
                                        Dim BlkExplode As BlockReference = Nothing
                                        acEnt.Explode(acDBObjColl1)
                                        For Each acEntExplode As Autodesk.AutoCAD.DatabaseServices.Entity In acDBObjColl1
                                            If (TypeOf acEntExplode Is BlockReference) Then
                                                BlkExplode = acEntExplode
                                                Exit For
                                            End If
                                        Next
                                        'расчленяем метку 2 раз
                                        Dim acDBObjColl2 As DBObjectCollection = New DBObjectCollection()
                                        If IsNothing(BlkExplode) = False Then
                                            BlkExplode.Explode(acDBObjColl2)
                                            For Each acEntExplode As Autodesk.AutoCAD.DatabaseServices.Entity In acDBObjColl2
                                                If (TypeOf acEntExplode Is Line) Then
                                                    Dim LineExplode As Line = acEntExplode
                                                    Dim ArrayPline(,) As Double = {{LineExplode.StartPoint.X, LineExplode.EndPoint.X}, {LineExplode.StartPoint.Y, LineExplode.EndPoint.Y}}
                                                    FuncDrawElementsPlateToTpofile(ArrayPline, Dist, StartPoint, align, acPoly, LayerPP, LineExplode.ColorIndex, LineExplode.LineWeight, False, LineExplode.LinetypeScale, LineExplode.Linetype)
                                                    Erase ArrayPline
                                                End If
                                            Next
                                        End If
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование точек
                                ElseIf (TypeOf acEnt Is DBPoint) Then
                                    If DialogProfil.CheckBox7.Checked = True Then
                                        Dim ObjPoint As DBPoint = acEnt
                                        Dim Piket1 As Double = Nothing
                                        Dim offsetAlign1 As Double = Nothing
                                        Try
                                            '
                                            align.StationOffset(ObjPoint.Position.X, ObjPoint.Position.Y, Piket1, offsetAlign1)
                                            If Math.Abs(offsetAlign1) <= (Dist / 2) * kR Then
                                                Dim x1 As Double = StartPoint.X + Piket1
                                                Dim y1 As Double = StartPoint.Y - offsetAlign1 / kR
                                                Dim posBlkProf1 As Point3d = New Point3d(x1, y1, 0)
                                                Piket1 = Nothing
                                                offsetAlign1 = Nothing
                                                Dim DBPoint1 As DBPoint = ObjPoint.Clone()
                                                DBPoint1.Position = posBlkProf1
                                                DBPoint1.Layer = LayerPP
                                                acBlkTblRec.AppendEntity(DBPoint1)
                                                tr.AddNewlyCreatedDBObject(DBPoint1, True)
                                            End If
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException

                                        End Try
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'трансформирование точек
                                    'ElseIf (TypeOf acEnt Is Hatch) Then
                                    'Dim ObjPoint As Hatch = acEnt

                                End If
                            End If
                        End If
                    Next
                End If
                tr.Commit()
            End Using

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        locker.Dispose()
    End Sub
    'функция рисует элементы ситуации и рельефа в подпрофильной таблице
    Public Sub FuncDrawElementsPlateToTpofile(ByRef ArrayCoord(,) As Double, ByVal Dist As Double, ByVal StartPoint As Point2d, ByVal align As Alignment, ByVal acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline, ByVal LayerPP As String, ByVal IndexColor As Integer, ByVal LineWightPl As LineWeight, ByVal ClosePl As Boolean, ByVal ScalePl As Double, ByVal TypeLinePl As String, Optional ByVal kr As Double = 1)
        'начинаем перебор вершин
        Dim ArrayPline As Double(,) = Nothing 'массив для координат полилинии
        Dim countVert As Integer = 0 'счетчик вершин
        For i As Integer = 0 To ArrayCoord.GetUpperBound(1) - 1
            '1 точка
            Dim Point1 As Point2d = New Point2d(ArrayCoord(0, i), ArrayCoord(1, i))
            Dim Piket1 As Double = Nothing
            Dim offsetAlign1 As Double = Nothing
            Dim offsetRez1 As Boolean = False
            'находим пикет и смещение для 2 вершины полилинии
            Try
                align.StationOffset(Point1.X, Point1.Y, Piket1, offsetAlign1)
                offsetRez1 = True
            Catch ex As Autodesk.Civil.PointNotOnEntityException
                offsetRez1 = False
            End Try
            '2 точка
            Dim Point2 As Point2d = New Point2d(ArrayCoord(0, i + 1), ArrayCoord(1, i + 1))
            Dim Piket2 As Double = Nothing
            Dim offsetAlign2 As Double = Nothing
            Dim offsetRez2 As Boolean = False
            'находим пикет и смещение для 2 вершины полилинии
            Try
                align.StationOffset(Point2.X, Point2.Y, Piket2, offsetAlign2)
                offsetRez2 = True
            Catch ex As Autodesk.Civil.PointNotOnEntityException
                offsetRez2 = False
            End Try
            'начинаем анализ
            If offsetRez1 = True And offsetRez2 = True Then
                'обе вершины в пределах рамки (записываем первую вершину в массив)
                If Math.Abs(offsetAlign1) <= (Dist / 2) * kr And Math.Abs(offsetAlign2) <= (Dist / 2) * kr Then
                    ReDim Preserve ArrayPline(1, countVert)
                    ArrayPline(0, countVert) = StartPoint.X + Piket1
                    ArrayPline(1, countVert) = StartPoint.Y - offsetAlign1 / kr
                    countVert += 1
                    If i = ArrayCoord.GetUpperBound(1) - 1 Then
                        ReDim Preserve ArrayPline(1, countVert)
                        ArrayPline(0, countVert) = StartPoint.X + Piket2
                        ArrayPline(1, countVert) = StartPoint.Y - offsetAlign2 / kr
                        Dim IdAcPoly As ObjectId = AcadFunction.FuncDrawPolyline(ArrayPline, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl)
                        Erase ArrayPline
                        countVert = 0
                    End If
                    'если вторая вершина выше то записываем в массив первую и продолжаем
                ElseIf Math.Abs(offsetAlign1) <= (Dist / 2) * kr And Math.Abs(offsetAlign2) > (Dist / 2) * kr Then
                    ReDim Preserve ArrayPline(1, countVert)
                    ArrayPline(0, countVert) = StartPoint.X + Piket1
                    ArrayPline(1, countVert) = StartPoint.Y - offsetAlign1 / kr
                    countVert += 1
                    Dim intersectPoint As Point3dCollection = New Point3dCollection
                    Dim pt3d1 As Point3d = New Point3d(StartPoint.X + Piket1, StartPoint.Y - offsetAlign1 / kr, 0)
                    Dim pt3d2 As Point3d = New Point3d(StartPoint.X + Piket2, StartPoint.Y - offsetAlign2 / kr, 0)
                    Dim Line1 As Line = New Line(pt3d1, pt3d2)
                    Line1.IntersectWith(acPoly, Intersect.OnBothOperands, intersectPoint, 1, 0)
                    Dim lg As Double = 10000000000
                    If intersectPoint.Count > 0 Then
                        ReDim Preserve ArrayPline(1, countVert)
                        For Each ptIns As Point3d In intersectPoint
                            Dim Lenght As Double = pt3d1.GetVectorTo(ptIns).Length
                            If Lenght < lg Then
                                ArrayPline(0, countVert) = ptIns.X
                                ArrayPline(1, countVert) = ptIns.Y
                                lg = Lenght
                            End If
                        Next
                        countVert += 1
                    End If
                    Dim IdAcPoly As ObjectId = AcadFunction.FuncDrawPolyline(ArrayPline, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl)
                    Erase ArrayPline
                    countVert = 0
                    'если первая вершина выше второй
                ElseIf Math.Abs(offsetAlign1) > (Dist / 2) * kr And Math.Abs(offsetAlign2) <= (Dist / 2) * kr Then
                    Dim intersectPoint As Point3dCollection = New Point3dCollection
                    Dim pt3d1 As Point3d = New Point3d(StartPoint.X + Piket1, StartPoint.Y - offsetAlign1 / kr, 0)
                    Dim pt3d2 As Point3d = New Point3d(StartPoint.X + Piket2, StartPoint.Y - offsetAlign2 / kr, 0)
                    Dim Line1 As Line = New Line(pt3d1, pt3d2)
                    Line1.IntersectWith(acPoly, Intersect.ExtendThis, intersectPoint, 1, 0)
                    Dim lg As Double = 10000000000
                    If intersectPoint.Count > 0 Then
                        ReDim Preserve ArrayPline(1, countVert)
                        For Each ptIns As Point3d In intersectPoint
                            Dim Lenght As Double = pt3d1.GetVectorTo(ptIns).Length
                            If Lenght < lg Then
                                ArrayPline(0, countVert) = ptIns.X
                                ArrayPline(1, countVert) = ptIns.Y
                                lg = Lenght
                            End If
                        Next
                        countVert += 1
                    End If
                    If i = ArrayCoord.GetUpperBound(1) - 1 Then
                        ReDim Preserve ArrayPline(1, countVert)
                        ArrayPline(0, countVert) = StartPoint.X + Piket2
                        ArrayPline(1, countVert) = StartPoint.Y - offsetAlign2 / kr
                        Dim IdAcPoly As ObjectId = AcadFunction.FuncDrawPolyline(ArrayPline, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl)
                        Erase ArrayPline
                        countVert = 0
                    End If
                    'если обе вершины выше ограничения
                ElseIf Math.Abs(offsetAlign1) > (Dist / 2) * kr And Math.Abs(offsetAlign2) > (Dist / 2) * kr Then
                    If (offsetAlign1 < 0 And offsetAlign2 > 0) Or (offsetAlign1 > 0 And offsetAlign2 < 0) Then
                        Dim intersectPoint As Point3dCollection = New Point3dCollection
                        Dim pt3d1 As Point3d = New Point3d(StartPoint.X + Piket1, StartPoint.Y - offsetAlign1 / kr, 0)
                        Dim pt3d2 As Point3d = New Point3d(StartPoint.X + Piket2, StartPoint.Y - offsetAlign2 / kr, 0)
                        Dim Line1 As Line = New Line(pt3d1, pt3d2)
                        Line1.IntersectWith(acPoly, Intersect.ExtendThis, intersectPoint, 1, 0)
                        If intersectPoint.Count > 1 Then
                            ReDim Preserve ArrayPline(1, countVert)
                            ArrayPline(0, countVert) = intersectPoint.Item(0).X
                            ArrayPline(1, countVert) = intersectPoint.Item(0).Y
                            countVert += 1
                            ReDim Preserve ArrayPline(1, countVert)
                            ArrayPline(0, countVert) = intersectPoint.Item(1).X
                            ArrayPline(1, countVert) = intersectPoint.Item(1).Y
                        End If
                        Dim IdAcPoly As ObjectId = AcadFunction.FuncDrawPolyline(ArrayPline, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl)
                        Erase ArrayPline
                        countVert = 0
                    End If
                End If
                'если объект одним первым концом за пределами трассы
            ElseIf offsetRez1 = False And offsetRez2 = True Then
                If Math.Abs(offsetAlign2) <= Dist / 2 Then
                    ReDim Preserve ArrayPline(1, countVert)
                    Dim Pt2 As Double = Nothing
                    Dim Pt01 As Point3d = New Point3d(Point1.X, Point1.Y, 0) 'начальная точка линии (внешняя)
                    Dim Pt02 As Point3d = New Point3d(Point2.X, Point2.Y, 0) 'конечная точка линии (внутренняя)
                    Dim Line0 As Line = New Line(Pt01, Pt02) 'создаем линию по точкам
                    Dim pt1 As Point3d = New Point3d(0, 0, 0)
                    Try
                        'pt1 = align.GetPointAtDist(Piket2) 'точка для которой возможно вычислить пикетажное значение (второй конец)
                        Dim n As Double = 0
                        Dim e As Double = 0
                        align.PointLocation(Piket2, offsetAlign2, e, n)
                        pt1 = New Point3d(e, n, 0)
                    Catch ex As Autodesk.Civil.PointNotOnEntityException

                    End Try

                    'определям в начале или в конце трассы это пересечение
                    Dim l1 As Double = Pt01.GetVectorTo(align.StartPoint).Length
                    Dim l2 As Double = Pt01.GetVectorTo(align.EndPoint).Length
                    If l1 < l2 Then
                        Pt2 = align.StartingStation 'пикет начала трассы
                    Else
                        Pt2 = align.EndingStation 'пикет конца трассы
                    End If
                    'находим координаты 4 точек (координаты прямоугольника с которым возможно пересечение)
                    Dim x1 As Double = Nothing
                    Dim y1 As Double = Nothing
                    align.PointLocation(Piket2, Dist / 2, x1, y1)
                    Dim LNPt1 As Point3d = New Point3d(x1, y1, 0)

                    Dim x2 As Double = Nothing
                    Dim y2 As Double = Nothing
                    align.PointLocation(Pt2, Dist / 2, x2, y2)
                    Dim LNPt2 As Point3d = New Point3d(x2, y2, 0)

                    Dim x3 As Double = Nothing
                    Dim y3 As Double = Nothing
                    align.PointLocation(Pt2, -1 * Dist / 2, x3, y3)
                    Dim LNPt3 As Point3d = New Point3d(x3, y3, 0)

                    Dim x4 As Double = Nothing
                    Dim y4 As Double = Nothing
                    align.PointLocation(Piket2, -1 * Dist / 2, x4, y4)
                    Dim LNPt4 As Point3d = New Point3d(x4, y4, 0)

                    '3 варианта пересечения составляем 3 линии
                    Dim Line1 As Line = New Line(LNPt1, LNPt2)
                    Dim Line2 As Line = New Line(LNPt2, LNPt3)
                    Dim Line3 As Line = New Line(LNPt3, LNPt4)
                    'проверяем пересечки
                    Dim lg As Double = 100000000000
                    Dim intersectPoint1 As Point3dCollection = New Point3dCollection
                    '1линия
                    Line1.IntersectWith(Line0, Intersect.ExtendThis, intersectPoint1, 1, 0)
                    If intersectPoint1.Count > 0 Then
                        Dim ptIns As Point3d = intersectPoint1.Item(0)
                        Dim lenght1 As Double = Pt01.GetVectorTo(ptIns).Length
                        If lenght1 < lg Then
                            Try
                                align.StationOffset(ptIns.X, ptIns.Y, Piket1, offsetAlign1)
                                offsetRez1 = True
                                lg = lenght1
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                offsetRez1 = False
                            End Try
                        End If
                    End If
                    '3 линия
                    Dim intersectPoint2 As Point3dCollection = New Point3dCollection
                    Line3.IntersectWith(Line0, Intersect.ExtendThis, intersectPoint2, 1, 0)
                    If intersectPoint2.Count > 0 Then
                        Dim ptIns As Point3d = intersectPoint2.Item(0)
                        Dim lenght1 As Double = Pt01.GetVectorTo(ptIns).Length
                        If lenght1 < lg Then
                            Try
                                align.StationOffset(ptIns.X, ptIns.Y, Piket1, offsetAlign1)
                                offsetRez1 = True
                                lg = lenght1
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                offsetRez1 = False
                            End Try
                        End If
                    End If
                    '2 линия
                    Dim intersectPoint3 As Point3dCollection = New Point3dCollection
                    Line2.IntersectWith(Line0, Intersect.ExtendThis, intersectPoint3, 1, 0)
                    If intersectPoint3.Count > 0 Then
                        Dim ptIns As Point3d = intersectPoint3.Item(0)
                        Dim lenght1 As Double = Pt01.GetVectorTo(ptIns).Length
                        If lenght1 < lg Then
                            Try
                                align.StationOffset(ptIns.X, ptIns.Y, Piket1, offsetAlign1)
                                offsetRez1 = True
                                lg = lenght1
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                offsetRez1 = False
                            End Try
                        End If
                    End If
                    ArrayPline(0, countVert) = StartPoint.X + Piket1
                    ArrayPline(1, countVert) = StartPoint.Y - offsetAlign1 / kr
                    countVert += 1
                    If i = ArrayCoord.GetUpperBound(1) - 1 Then
                        ReDim Preserve ArrayPline(1, countVert)
                        ArrayPline(0, countVert) = StartPoint.X + Piket2
                        ArrayPline(1, countVert) = StartPoint.Y - offsetAlign2 / kr
                        Dim IdAcPoly As ObjectId = AcadFunction.FuncDrawPolyline(ArrayPline, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl)
                        Erase ArrayPline
                        countVert = 0
                    End If
                End If
                'если объект одним последним концом за пределами трассы
            ElseIf offsetRez1 = True And offsetRez2 = False Then
                If Math.Abs(offsetAlign1) <= Dist / 2 Then
                    Dim Pt2 As Double = Nothing
                    ReDim Preserve ArrayPline(1, countVert)
                    ArrayPline(0, countVert) = StartPoint.X + Piket1
                    ArrayPline(1, countVert) = StartPoint.Y - offsetAlign1 / kr
                    countVert += 1
                    ReDim Preserve ArrayPline(1, countVert)
                    Dim Pt01 As Point3d = New Point3d(Point1.X, Point1.Y, 0)
                    Dim Pt02 As Point3d = New Point3d(Point2.X, Point2.Y, 0)
                    Dim Line0 As Line = New Line(Pt01, Pt02)
                    Dim l1 As Double = Pt01.GetVectorTo(align.StartPoint).Length
                    Dim l2 As Double = Pt01.GetVectorTo(align.EndPoint).Length
                    If l1 < l2 Then
                        Pt2 = align.StartingStation 'координаты начала трассы
                    Else
                        Pt2 = align.EndingStation 'координаты конца трассы
                    End If
                    'находим координаты 4 точек
                    Dim x1 As Double = Nothing
                    Dim y1 As Double = Nothing
                    align.PointLocation(Piket1, Dist / 2, x1, y1)
                    Dim LNPt1 As Point3d = New Point3d(x1, y1, 0)

                    Dim x2 As Double = Nothing
                    Dim y2 As Double = Nothing
                    align.PointLocation(Pt2, Dist / 2, x2, y2)
                    Dim LNPt2 As Point3d = New Point3d(x2, y2, 0)

                    Dim x3 As Double = Nothing
                    Dim y3 As Double = Nothing
                    align.PointLocation(Pt2, -1 * Dist / 2, x3, y3)
                    Dim LNPt3 As Point3d = New Point3d(x3, y3, 0)

                    Dim x4 As Double = Nothing
                    Dim y4 As Double = Nothing
                    align.PointLocation(Piket1, -1 * Dist / 2, x4, y4)
                    Dim LNPt4 As Point3d = New Point3d(x4, y4, 0)

                    '3 варианта пересечения составляем 3 линии
                    Dim Line1 As Line = New Line(LNPt1, LNPt2)
                    Dim Line2 As Line = New Line(LNPt2, LNPt3)
                    Dim Line3 As Line = New Line(LNPt3, LNPt4)
                    'проверяем пересечки
                    Dim lg As Double = 100000000000
                    Dim intersectPoint1 As Point3dCollection = New Point3dCollection
                    '1линия
                    Line1.IntersectWith(Line0, Intersect.ExtendThis, intersectPoint1, 1, 0)
                    If intersectPoint1.Count > 0 Then
                        Dim ptIns As Point3d = intersectPoint1.Item(0)
                        Dim lenght1 As Double = Pt01.GetVectorTo(ptIns).Length
                        If lenght1 < lg Then
                            Try
                                align.StationOffset(ptIns.X, ptIns.Y, Piket2, offsetAlign2)
                                offsetRez1 = True
                                lg = lenght1
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                offsetRez1 = False
                            End Try
                        End If
                    End If
                    '3 линия
                    Dim intersectPoint2 As Point3dCollection = New Point3dCollection
                    Line3.IntersectWith(Line0, Intersect.ExtendThis, intersectPoint2, 1, 0)
                    If intersectPoint2.Count > 0 Then
                        Dim ptIns As Point3d = intersectPoint2.Item(0)
                        Dim lenght1 As Double = Pt01.GetVectorTo(ptIns).Length
                        If lenght1 < lg Then
                            Try
                                align.StationOffset(ptIns.X, ptIns.Y, Piket2, offsetAlign2)
                                offsetRez1 = True
                                lg = lenght1
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                offsetRez1 = False
                            End Try
                        End If
                    End If
                    '2 линия
                    Dim intersectPoint3 As Point3dCollection = New Point3dCollection
                    Line2.IntersectWith(Line0, Intersect.ExtendThis, intersectPoint3, 1, 0)
                    If intersectPoint3.Count > 0 Then
                        Dim ptIns As Point3d = intersectPoint3.Item(0)
                        Dim lenght1 As Double = Pt01.GetVectorTo(ptIns).Length
                        If lenght1 < lg Then
                            Try
                                align.StationOffset(ptIns.X, ptIns.Y, Piket2, offsetAlign2)
                                offsetRez1 = True
                                lg = lenght1
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                offsetRez1 = False
                            End Try
                        End If
                    End If
                    ArrayPline(0, countVert) = StartPoint.X + Piket2
                    ArrayPline(1, countVert) = StartPoint.Y - offsetAlign2 / kr
                    Dim IdAcPoly As ObjectId = AcadFunction.FuncDrawPolyline(ArrayPline, LayerPP, IndexColor, LineWightPl, ClosePl, ScalePl, TypeLinePl)
                    Erase ArrayPline
                    countVert = 0
                End If
            End If
        Next i
    End Sub
End Class
