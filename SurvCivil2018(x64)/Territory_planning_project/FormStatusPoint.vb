Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Collections.Generic
Imports System
Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.LayerManager.LayerFilter
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports Autodesk.AutoCAD.Colors
Imports System.Drawing.Graphics
Imports System.Drawing
Imports System.IO
Imports System.String
Public Class FormStatusPoint
    Dim PutshBlock As String = "k:\STPRGeoTools\ZemKadastr\MPlan\Blocks\"
    Public pointLayer As String() 'массив для слоев
    Public formload As Boolean = False
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim PointZakr As String() = {"026016000000 Нет закрепления", "026005000000 Угол (забора, изгороди, ограды)", "026006000000 Угол (здания сооружения)",
        "026001000000 Уголок металлический", "026004000000 Столб (деревянный, бетонный, кирпичный)", "026002000000 Труба (металлическая, асбоцементная, бетонная)",
        "026003000000 Пилон железобетонный", "026007000000 Примыкание забора к зданию", "026008000000  Марка, дюбель, штырь на горизонтальной поверхности",
        "026009000000	 Марка, дюбель, метал штырь на вертикальной поверхности", "026010000000 Угол поребрика", "026012000000 Швеллер (Т - образный, металлический)",
        "026013000000	 Маркировка на здании (сооружении, заборе)", "026014000000 Штырь, арматура, стержень", "026015000000	 Иные способы закрепления границ"}
        ComboBox2.DataSource = PointZakr
        Dim layerRez As Boolean = AcadFunction.FuncReadLayersDWG(pointLayer) 'массив с включенными слоями
        Dim pointLayer2 As String()
        ReDim pointLayer2(0)
        pointLayer2(0) = " "
        If IsArray(pointLayer) = True Then
            For i As Integer = 0 To pointLayer.GetUpperBound(0)
                ReDim Preserve pointLayer2(i + 1)
                pointLayer2(i + 1) = pointLayer(i)
            Next
        End If

        Dim arrayTextStyles As String() = {""}
        Dim boolfindtextstyle As Integer = AcadFunction.FuncReadListTextStyle(arrayTextStyles)
        ComboBox3.DataSource = arrayTextStyles

        'метод определения координат
        Dim MetodCoord As String() = {"Геодезический метод", "Метод спутниковых геодезических измерений", "Аналитический метод", "Картометрический метод", "Фотограмметрический метод", "Иное описание"}
        ComboBox1.DataSource = MetodCoord
    End Sub
    Private Sub Form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        Dim Hpoint1 As Integer = 10
        Dim Xpoint As Single = 5
        Dim Ypoint As Single = 5
        Dim HText1 As Integer = 15
        Dim SBrush As New System.Drawing.SolidBrush(System.Drawing.Color.Black)
        Dim TextFont As New System.Drawing.Font("Arial", HText1, FontStyle.Regular)
        Dim TextFont1 As New System.Drawing.Font("Arial", HText1, FontStyle.Underline)
        e.Graphics.DrawEllipse(Pens.Black, 200, 200, Hpoint1, Hpoint1)
        e.Graphics.FillEllipse(Brushes.Black, 200, 200, Hpoint1, Hpoint1)
        If RadioButton11.Checked = True Then
            e.Graphics.DrawString("н1", TextFont, Brushes.Black, 200 + Xpoint, 200 - Ypoint - HText1)
        ElseIf RadioButton12.Checked = True Then
            e.Graphics.DrawString("1", TextFont, Brushes.Black, 200 + Xpoint, 200 - Ypoint - HText1)
        ElseIf RadioButton13.Checked = True Then
            e.Graphics.DrawString("1", TextFont1, Brushes.Black, 200 + Xpoint, 200 - Ypoint - HText1)
        ElseIf RadioButton2.Checked = True Then
            e.Graphics.DrawString("т1", TextFont1, Brushes.Black, 200 + Xpoint, 200 - Ypoint - HText1)
        ElseIf RadioButton1.Checked = True Then
            Dim str As String = TextBox1.Text & "1"
            e.Graphics.DrawString(str, TextFont1, Brushes.Black, 200 + Xpoint, 200 - Ypoint - HText1)
        End If
    End Sub
    Private Sub RadioButton9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim invalid As Rectangle = New Rectangle(100, 100, 200, 200)
        Me.Invalidate(invalid)
    End Sub
    Private Sub RadioButton10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim invalid As Rectangle = New Rectangle(100, 100, 200, 200)
        Me.Invalidate(invalid)
    End Sub
    Private Sub RadioButton11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton11.CheckedChanged
        Dim invalid As Rectangle = New Rectangle(100, 100, 200, 200)
        Me.Invalidate(invalid)
    End Sub
    Private Sub RadioButton12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton12.CheckedChanged
        Dim invalid As Rectangle = New Rectangle(100, 100, 200, 200)
        Me.Invalidate(invalid)
    End Sub
    Private Sub RadioButton13_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton13.CheckedChanged
        Dim invalid As Rectangle = New Rectangle(100, 100, 200, 200)
        Me.Invalidate(invalid)
    End Sub

    Private Sub CheckBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.Click
        'погрешность определения точки (блокировка\разблокировка)
        If CheckBox1.Checked = True Then
            NumericUpDown6.Enabled = True
        Else
            NumericUpDown6.Enabled = False
        End If
    End Sub
    Private Sub CheckBox2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.Click
        'описание зактепления точки (блокировка\разблокировка)
        If CheckBox2.Checked = True Then
            ComboBox2.Enabled = True
        Else
            ComboBox2.Enabled = False
        End If
    End Sub
    Private Sub CheckBox5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox5.Click
        'список слоев точек (блокировка\разблокировка)
        If CheckBox5.Checked = True Then
            RadioButton11.Enabled = True
            RadioButton12.Enabled = True
            RadioButton13.Enabled = True
        Else
            RadioButton11.Enabled = False
            RadioButton12.Enabled = False
            RadioButton13.Enabled = False
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'нажата кнопка отмена
        formload = False
        Me.Hide()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'кнопка применить
        Dim count As Integer = 0
        Dim countStPoint As Integer = 0
        Dim NameBlock As String = ""
        Dim stPoint As String = "нет"
        Dim Mt As Double = 0
        Dim Zakr As String = ""
        Dim layerPoint As String = ""
        Dim colObj As ObjectIdCollection = New ObjectIdCollection
        Me.Hide()
        Dim mtPoint As Double = NumericUpDown6.Value
        Dim zakrPoint As String = ComboBox2.Text

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayMCounter As String(,) = Nothing
        Dim countArrayMCounter As Integer = 0
        Dim ArrayTablesXML As String(,) = Nothing 'массив с таблицами для фильтрации
        ReDim ArrayTablesXML(1, 6)
        ArrayTablesXML(0, 0) = "PROJECT_PARCEL"
        ArrayTablesXML(1, 0) = "Участок постоянного отвода"
        ArrayTablesXML(0, 1) = "TEMPORARY_PARCEL"
        ArrayTablesXML(1, 1) = "Участок временного отвода"
        ArrayTablesXML(0, 2) = "CADASTR_OSTROV"
        ArrayTablesXML(1, 2) = "Внутренний островок"
        ArrayTablesXML(0, 3) = "MPLAN_PARCEL"
        ArrayTablesXML(1, 3) = "Участок для постановки на ГКН"
        ArrayTablesXML(0, 4) = "PROJECT_BUILDING"
        ArrayTablesXML(1, 4) = "Здание или сооружение проектируемое"
        ArrayTablesXML(0, 5) = "TERRITORIAL_ZONE"
        ArrayTablesXML(1, 5) = "Зона с особыми условиями"
        ArrayTablesXML(0, 6) = "PROJECT_TREE"
        ArrayTablesXML(1, 6) = "Снос зеленых насаждений"
        Dim CountPoint As Integer = 0
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
                Dim pointStartPline As Point2d = Nothing
                Dim numberVertexStart As Integer = -1
                Dim pointEndPline As Point2d = Nothing
                Dim numberVertexEnd As Integer = -1
                If RadioButton3.Checked = True Then
                    Dim prPointRes1 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите контур (полилинию):")
                    If prPointRes1.Status = PromptStatus.OK Then
                        Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                        If IsDBNull(entObjID) = False Then
                            Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForRead)
                            If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Pline = entObj
                                numberVertexStart = 0
                                numberVertexEnd = Pline.NumberOfVertices - 1
                            End If
                        End If
                    Else
                        Exit Sub
                    End If
                Else
                    Dim prPointRes1 As PromptEntityResult = Nothing
                    prPointRes1 = ed.GetEntity(vbLf & "Выберите первую точку ближайшую к узлу полилинии:")
                    If prPointRes1.Status = PromptStatus.OK Then
                        Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                        If IsDBNull(entObjID) = False Then
                            Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForRead)
                            If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Pline = entObj
                                Dim point3d1 As Point3d = prPointRes1.PickedPoint 'смотрим координаты точки указания
                                Dim point2dPic As Point2d = point3d1.Convert2d(New Plane)
                                Dim Dist As Double = 1000000
                                For i As Integer = 0 To Pline.NumberOfVertices - 1
                                    Dim tempPoint As Point2d = Pline.GetPoint2dAt(i)
                                    Dim L As Double = tempPoint.GetVectorTo(point2dPic).Length
                                    If L < Dist Then
                                        Dist = L
                                        pointStartPline = tempPoint
                                        numberVertexStart = i
                                    End If
                                Next i
                            End If
                        End If
                    End If

                    prPointRes1 = ed.GetEntity(vbLf & "Выберите последнюю точку ближайшую к узлу полилинии:")
                    If prPointRes1.Status = PromptStatus.OK Then
                        Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                        If IsDBNull(entObjID) = False Then
                            Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForRead)
                            If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                If IsNothing(Pline) = False Then
                                    If Pline.Id <> entObjID Then
                                        MsgBox("Изменить статус точкам можно только если они находятся на одной и тойже полилинии. Макро превано!!!!")
                                        Exit Sub
                                    End If
                                Else
                                    Pline = entObj
                                End If
                                Dim point3d1 As Point3d = prPointRes1.PickedPoint 'смотрим координаты точки указания
                                Dim point2dPic As Point2d = point3d1.Convert2d(New Plane)
                                Dim Dist As Double = 1000000
                                For i As Integer = 0 To Pline.NumberOfVertices - 1
                                    Dim tempPoint As Point2d = Pline.GetPoint2dAt(i)
                                    Dim L As Double = tempPoint.GetVectorTo(point2dPic).Length
                                    If L < Dist Then
                                        Dist = L
                                        pointEndPline = tempPoint
                                        numberVertexEnd = i
                                    End If
                                Next i
                            End If
                        End If
                    End If
                End If
                If IsNothing(Pline) = False Then
                    If numberVertexStart = -1 Then
                        numberVertexStart = 0
                    End If
                    If numberVertexEnd = -1 Then
                        numberVertexEnd = Pline.NumberOfVertices - 1
                    End If
                End If

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'фильтруем точки
                Dim idTablePs As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties("CADASTRAL_POINT")
                Dim nameBlk As String = ComboBox1.Text
                Dim arrayNamePoint As String() = {"PROJECT_POINT", "TEMP_POINT"}
                Dim idPointColl As ObjectIdCollection = FuncCadastr.FuncFilterGeoPointByBoundObject(Pline, arrayNamePoint)
                If idPointColl.Count > 0 Then
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'меняем статус точки
                    For i As Integer = numberVertexStart To numberVertexEnd
                        Dim pt As Point2d = Pline.GetPoint2dAt(i)
                        'ищем геоточку в узле
                        Dim idGeoPoint As ObjectId = FuncCadastr.FuncFindGeoPoint(idPointColl, pt, 0.01)
                        If IsNothing(idGeoPoint) = False Then
                            If idGeoPoint.IsNull = False Then
                                'меняем статус
                                Dim acBlk As BlockReference = idGeoPoint.GetObject(OpenMode.ForWrite)
                                'считываем старые атрибуты
                                Dim ArrayOBJ1 As String(,) = Nothing 'формируем массив для записи PS
                                Dim countZ As Integer = 0
                                Dim ArrayData As String(,) = Nothing
                                Dim boolreadBlk As Boolean = AcadFunction.FuncReadAllBlockAttr(acBlk, ArrayData)
                                If IsArray(ArrayData) = True Then
                                    For j As Integer = 0 To ArrayData.GetUpperBound(1)
                                        Dim tagStr As String = ArrayData(0, j)
                                        If tagStr Like "Точность_положения" Or tagStr Like "Delta_geopoint" Or tagStr Like "DELTA" Then
                                            If CheckBox1.Checked = True Then
                                                ArrayData(1, j) = NumericUpDown6.Value

                                                ReDim Preserve ArrayOBJ1(2, countZ)
                                                ArrayOBJ1(0, countZ) = "CADASTRAL_POINT"
                                                ArrayOBJ1(1, countZ) = "Delta"
                                                ArrayOBJ1(2, countZ) = NumericUpDown6.Value
                                                countZ += 1
                                            End If
                                        ElseIf tagStr Like "Описание_закрепления" Or tagStr Like "GeopointZacrep" Or tagStr Like "FASTDESK" Then
                                            If CheckBox2.Checked = True Then
                                                ArrayData(1, j) = ComboBox2.Text

                                                ReDim Preserve ArrayOBJ1(2, countZ)
                                                ArrayOBJ1(0, countZ) = "CADASTRAL_POINT"
                                                ArrayOBJ1(1, countZ) = "FastDesk"
                                                ArrayOBJ1(2, countZ) = ComboBox2.Text
                                                countZ += 1
                                            End If
                                        ElseIf tagStr Like "Метод_определения_координат" Or tagStr Like "Geopoint_opred" Or tagStr Like "METHODCOORDINATES" Then
                                            If CheckBox3.Checked = True Then
                                                ArrayData(1, j) = ComboBox1.Text

                                                ReDim Preserve ArrayOBJ1(2, countZ)
                                                ArrayOBJ1(0, countZ) = "CADASTRAL_POINT"
                                                ArrayOBJ1(1, countZ) = "MethodCoordinates"
                                                ArrayOBJ1(2, countZ) = ComboBox1.Text
                                                countZ += 1
                                            End If
                                        ElseIf tagStr Like "Имя_точки" Or tagStr Like "NUM_GEOPOINT" Then
                                            If CheckBox5.Checked = True Then
                                                Dim numPoint As String = ArrayData(1, j)
                                                If numPoint.Trim.Length > 0 Then
                                                    Dim result As System.Text.RegularExpressions.MatchCollection = New System.Text.RegularExpressions.Regex("[0-9]+").Matches(numPoint)
                                                    If result.Count > 0 Then
                                                        Dim value As Integer = Val(result.Item(0).Value)
                                                        ReDim Preserve ArrayOBJ1(2, countZ)
                                                        ArrayOBJ1(0, countZ) = "CADASTRAL_POINT"
                                                        ArrayOBJ1(1, countZ) = "NUMBER"
                                                        countZ += 1
                                                        If RadioButton11.Checked = True Then
                                                            ArrayData(1, j) = "н" & value
                                                            ArrayData(4, j) = 0
                                                            ArrayOBJ1(2, countZ - 1) = "н" & value
                                                        ElseIf RadioButton12.Checked = True Then
                                                            ArrayData(1, j) = value
                                                            ArrayData(4, j) = 0
                                                            ArrayOBJ1(2, countZ - 1) = value
                                                        ElseIf RadioButton13.Checked = True Then
                                                            ArrayData(1, j) = "%%u" & value
                                                            ArrayData(4, j) = 0.35
                                                            ArrayOBJ1(2, countZ - 1) = "%%u" & value
                                                        ElseIf RadioButton2.Checked = True Then
                                                            ArrayData(1, j) = "т" & value
                                                            ArrayData(4, j) = 0
                                                            ArrayOBJ1(2, countZ - 1) = "т" & value
                                                        Else
                                                            ArrayData(1, j) = TextBox1.Text & value
                                                            ArrayData(4, j) = 0
                                                            ArrayOBJ1(2, countZ - 1) = TextBox1.Text & value
                                                        End If
                                                    End If
                                                End If
                                                If NumericUpDown1.Value > 0 Then
                                                    ArrayData(3, j) = NumericUpDown1.Value
                                                End If
                                                If ComboBox3.Text.Trim.Length > 0 Then
                                                    ArrayData(5, j) = ComboBox3.Text
                                                End If
                                                If CheckBox4.Checked = True Then
                                                    ArrayData(6, j) = NumericUpDown2.Value
                                                End If
                                            End If
                                        End If
                                    Next j
                                End If
                                Dim boolInsBlk As Boolean = AcadFunction.FuncWriteAllBlockAttr(acBlk, ArrayData)
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'записываем расширенные данные
                                If IsNothing(idTablePs) = False Then
                                    If idTablePs.IsNull = False Then
                                        Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(acBlk, idTablePs)
                                        Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acBlk, ArrayOBJ1)
                                    End If
                                End If
                                If boolInsBlk = True Then
                                    count += 1
                                End If
                            End If
                        End If
                    Next
                End If
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Не удалось преобразовать все точки.")
                tr.Commit()
                locker.Dispose()
                Exit Sub
            End Try
            tr.Commit()
            locker.Dispose()
            MsgBox("Всего преобразовано " & count & " точек.")
        End Using
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            TextBox1.Enabled = True
        Else
            TextBox1.Enabled = False
        End If
        Dim invalid As Rectangle = New Rectangle(100, 100, 200, 200)
        Me.Invalidate(invalid)
    End Sub
    'метод определения координат
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            ComboBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
        End If
    End Sub
    'изменить цвет
    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            NumericUpDown2.Enabled = True
        Else
            NumericUpDown2.Enabled = False
        End If
    End Sub

    Private Sub FormStatusPoint_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub
End Class