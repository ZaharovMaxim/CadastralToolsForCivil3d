Imports System.Collections.Generic
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports CadastralTools

''' <summary>
''' Выбор замкнутой полилинии на чертеже AutoCAD и чтение её вершин
''' в контур лесотаксационного выдела.
''' </summary>
Public Module MdPickContour

    ''' <summary>
    ''' Запросить у пользователя замкнутую полилинию и заполнить контур её вершинами.
    ''' </summary>
    ''' <param name="contour">Контур выдела (заполняется).</param>
    ''' <param name="deltaGeopoint">Погрешность положения точек, м.</param>
    ''' <param name="area">Площадь контура (кв. м), 0 при отказе.</param>
    ''' <returns>Пустая строка при успехе, иначе сообщение об ошибке/отказе.</returns>
    Public Function PickClosedPolyline(contour As ClsTaxationDescriptionXml.TdContourData,
                                       deltaGeopoint As String,
                                       ByRef area As Double) As String
        area = 0.0
        If contour Is Nothing Then Return "Контур не задан."
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        If doc Is Nothing Then Return "Нет активного документа AutoCAD."
        Dim ed As Editor = doc.Editor

        Dim opts As New PromptEntityOptions(vbLf & "Выберите замкнутую полилинию контура выдела")
        opts.SetRejectMessage("Можно выбрать только полилинию.")
        opts.AddAllowedClass(GetType(Polyline), True)
        opts.AddAllowedClass(GetType(Polyline2d), True)
        opts.AddAllowedClass(GetType(Polyline3d), True)
        Dim res As PromptEntityResult = ed.GetEntity(opts)
        If res.Status <> PromptStatus.OK Then Return "Полилиния не выбрана."

        Try
            Using tr As Transaction = doc.TransactionManager.StartTransaction()
                Dim pts As New List(Of Point3d)()
                Dim ent As Entity = DirectCast(tr.GetObject(res.ObjectId, OpenMode.ForRead), Entity)
                If TypeOf ent Is Polyline Then
                    Dim pl As Polyline = DirectCast(ent, Polyline)
                    If Not pl.Closed Then Return "Полилиния не замкнута."
                    For i As Integer = 0 To pl.NumberOfVertices - 1
                        pts.Add(pl.GetPoint3dAt(i))
                    Next
                    area = pl.Area
                ElseIf TypeOf ent Is Polyline2d Then
                    Dim pl2 As Polyline2d = DirectCast(ent, Polyline2d)
                    If Not pl2.Closed Then Return "Полилиния не замкнута."
                    For Each cv As Vertex2d In pl2
                        pts.Add(cv.Position)
                    Next
                    area = pl2.Area
                ElseIf TypeOf ent Is Polyline3d Then
                    Dim pl3 As Polyline3d = DirectCast(ent, Polyline3d)
                    If Not pl3.Closed Then Return "Полилиния не замкнута."
                    For Each pv As PolylineVertex3d In pl3
                        pts.Add(pv.Position)
                    Next
                    area = pl3.Area
                Else
                    Return "Выбранный объект не является полилинией."
                End If
                tr.Commit()

                If pts.Count < 3 Then Return "В полилинии меньше трёх вершин."

                ' Первая и последняя точки замкнутой полилинии совпадают — дубликат не пишем
                If pts.Count > 1 AndAlso pts(0).DistanceTo(pts(pts.Count - 1)) < 0.000001 Then
                    pts.RemoveAt(pts.Count - 1)
                End If

                contour.Ordinates.Clear()
                Dim n As Integer = 1
                For Each p In pts
                    Dim o As New ClsTaxationDescriptionXml.TdOrdinate()
                    o.X = p.X.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                    o.Y = p.Y.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                    o.Z = p.Z.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                    o.OrdNmb = n
                    o.DeltaGeopoint = deltaGeopoint
                    contour.Ordinates.Add(o)
                    n += 1
                Next
                Return ""
            End Using
        Catch ex As Exception
            Return "Ошибка чтения полилинии: " & ex.Message
        End Try
    End Function

End Module
