Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD

Public Class CreateDWGObject
    Private _db As Database = Nothing
    Public Sub New(ByRef DBDrawing As Database)
        _db = DBDrawing
    End Sub

    Public Property DBDrawing As Database
        Get
            Return _db
        End Get
        Set(value As Database)
            _db = value
        End Set
    End Property
    '===================================================================================================
    'создание полилинии из полигона
    Public Function createPolylineByPolygon(ByVal polygon As NetTopologySuite.Geometries.Polygon, Optional ByRef ringPolyline As Polyline = Nothing) As Polyline
        createPolylineByPolygon = Nothing
        If IsNothing(DBDrawing) = False Then
            If IsNothing(polygon) = False Then
                If polygon.Coordinates.Count > 2 Then
                    Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
                        Dim acBlkTbl As BlockTable = acTrans.GetObject(DBDrawing.BlockTableId, OpenMode.ForRead)
                        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        Dim listCoordPolugon As NetTopologySuite.Geometries.Coordinate() = polygon.Coordinates
                        If IsNothing(ringPolyline) = True Then
                            ringPolyline = New Polyline()
                            ringPolyline.SetDatabaseDefaults()
                            acBlkTblRec.AppendEntity(ringPolyline)
                            acTrans.AddNewlyCreatedDBObject(ringPolyline, True)
                        Else
                            If ringPolyline.NumberOfVertices > listCoordPolugon.Count - 1 Then
                                For i As Integer = ringPolyline.NumberOfVertices - 1 To listCoordPolugon.Count - 2 Step -1
                                    ringPolyline.RemoveVertexAt(i)
                                Next i
                            End If
                        End If

                        If listCoordPolugon.Length > 2 Then
                            For i As Integer = 0 To listCoordPolugon.Length - 2
                                Dim coord As NetTopologySuite.Geometries.Coordinate = listCoordPolugon(i)
                                If i <= ringPolyline.NumberOfVertices - 1 Then
                                    ringPolyline.SetPointAt(i, New Point2d(coord.X, coord.Y))
                                Else
                                    ringPolyline.AddVertexAt(i, New Point2d(coord.X, coord.Y), 0, 0, 0)
                                End If
                            Next i
                            Dim startVertex As NetTopologySuite.Geometries.Coordinate = listCoordPolugon(0)
                            Dim endVertex As NetTopologySuite.Geometries.Coordinate = listCoordPolugon(listCoordPolugon.Length - 1)
                            Dim distance As Double = startVertex.Distance(endVertex)
                            If distance > 0.01 Then
                                ringPolyline.AddVertexAt(ringPolyline.NumberOfVertices - 1, New Point2d(endVertex.X, endVertex.Y), 0, 0, 0)
                            End If
                            ringPolyline.Closed = True
                        End If
                        acTrans.Commit()
                    End Using
                End If
            End If
            Return ringPolyline
        End If
    End Function

    Public Function UpdatePolylineFromPolygon(ByVal polygon As NetTopologySuite.Geometries.Polygon) As Boolean
        If IsNothing(DBDrawing) = True Then Return False
        If IsNothing(polygon) = True Then Return False
        Using Tr As Transaction = DBDrawing.TransactionManager.StartTransaction()
            'обновояем внешний контур
            Dim extendPolygon As NetTopologySuite.Geometries.LineString = polygon.ExteriorRing
            Dim exterialData As CadastralTools.StructureElement = TryCast(polygon.UserData, CadastralTools.StructureElement)
            If IsNothing(exterialData) = False Then
                Dim exteralhandle As String = exterialData.IdObject
                If IsNothing(exteralhandle) = False Then
                    Dim id As ObjectId = AcadFunction.FuncReturnObjectID(exteralhandle)
                    If IsNothing(id) = False Then
                        If Not (id.IsNull OrElse id.IsErased) Then
                            ' --- Открываем полилинию для записи ---
                            Dim plineCounter As Polyline = TryCast(Tr.GetObject(id, OpenMode.ForWrite), Polyline)
                            If IsNothing(plineCounter) = False Then
                                ' --- Обновляем вершины ---
                                Dim coords As NetTopologySuite.Geometries.Coordinate() = extendPolygon.Coordinates
                                If Not (coords Is Nothing OrElse coords.Length < 4) Then
                                    ' Реальных вершин — n - 1 (последняя точка = копия первой)
                                    Dim vertexCount As Integer = coords.Length - 1
                                    ' Удаляем все существующие вершины
                                    ' --- 1. Удаляем лишние вершины (если полилиния длиннее) ---
                                    ' Удаляем с конца, пока не останется ровно vertexCount вершин.
                                    While plineCounter.NumberOfVertices > vertexCount
                                        plineCounter.RemoveVertexAt(plineCounter.NumberOfVertices - 1)
                                    End While
                                    For j As Integer = 0 To vertexCount - 1
                                        If j <= plineCounter.NumberOfVertices - 1 Then
                                            plineCounter.SetPointAt(j, New Point2d(coords(j).X, coords(j).Y))
                                        Else
                                            plineCounter.AddVertexAt(j + 1, New Point2d(coords(j).X, coords(j).Y), 0.0, 0.0, 0.0)
                                        End If
                                    Next
                                    ' Замыкаем полилинию
                                    plineCounter.Closed = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            'островки
            For i As Integer = 0 To polygon.NumInteriorRings - 1
                Dim userPolygon As NetTopologySuite.Geometries.LineString = polygon.InteriorRings(i)
                Dim userData As CadastralTools.StructureElement = TryCast(userPolygon.UserData, CadastralTools.StructureElement)
                If userData Is Nothing Then Continue For
                Dim handle As String = userData.IdObject
                If IsNothing(handle) = False Then
                    Dim id As ObjectId = AcadFunction.FuncReturnObjectID(handle)
                    If IsNothing(id) = False Then
                        If id.IsNull OrElse id.IsErased Then Continue For
                        ' --- Открываем полилинию для записи ---
                        Dim plineCounter As Polyline = TryCast(Tr.GetObject(id, OpenMode.ForWrite), Polyline)
                        If IsNothing(plineCounter) = True Then Continue For
                        ' --- Обновляем вершины ---
                        Dim coords As NetTopologySuite.Geometries.Coordinate() = userPolygon.Coordinates
                        If coords Is Nothing OrElse coords.Length < 4 Then Continue For
                        ' Реальных вершин — n - 1 (последняя точка = копия первой)
                        Dim vertexCount As Integer = coords.Length - 1
                        ' --- 1. Удаляем лишние вершины (если полилиния длиннее) ---
                        ' Удаляем с конца, пока не останется ровно vertexCount вершин.
                        While plineCounter.NumberOfVertices > vertexCount
                            plineCounter.RemoveVertexAt(plineCounter.NumberOfVertices - 1)
                        End While
                        For j As Integer = 0 To vertexCount - 1
                            If j <= plineCounter.NumberOfVertices - 1 Then
                                plineCounter.SetPointAt(j, New Point2d(coords(j).X, coords(j).Y))
                            Else
                                plineCounter.AddVertexAt(j + 1, New Point2d(coords(j).X, coords(j).Y), 0.0, 0.0, 0.0)
                            End If
                        Next
                        ' Замыкаем полилинию
                        plineCounter.Closed = True
                    End If
                End If
            Next
            Tr.Commit()
        End Using
        Return True
    End Function
    '===================================================================================================
    'обновление полилинии по линии LineString
    'Аналог UpdatePolylineFromPolygon: линия не замыкается принудительно — замкнутость
    'берётся из самой геометрии (LineString.IsClosed)
    Public Function UpdatePolylineFromLineString(ByVal line As NetTopologySuite.Geometries.LineString) As Boolean
        If IsNothing(DBDrawing) = True Then Return False
        If IsNothing(line) = True Then Return False
        Using Tr As Transaction = DBDrawing.TransactionManager.StartTransaction()
            Dim userData As CadastralTools.StructureElement = TryCast(line.UserData, CadastralTools.StructureElement)
            If IsNothing(userData) = False Then
                Dim handle As String = userData.IdObject
                If IsNothing(handle) = False Then
                    Dim id As ObjectId = AcadFunction.FuncReturnObjectID(handle)
                    If IsNothing(id) = False Then
                        If Not (id.IsNull OrElse id.IsErased) Then
                            ' --- Открываем полилинию для записи ---
                            Dim plineCounter As Polyline = TryCast(Tr.GetObject(id, OpenMode.ForWrite), Polyline)
                            If IsNothing(plineCounter) = False Then
                                ' --- Обновляем вершины ---
                                Dim coords As NetTopologySuite.Geometries.Coordinate() = line.Coordinates
                                UpdatePolylineVertices(plineCounter, coords, line.IsClosed)
                            End If
                        End If
                    End If
                End If
            End If
            Tr.Commit()
        End Using
        Return True
    End Function
    '===================================================================================================
    'переписывает вершины существующей полилинии по массиву координат
    'Если линия замкнута, последняя координата — копия первой, реальных вершин на одну меньше
    Private Shared Function UpdatePolylineVertices(ByVal plineCounter As Polyline, ByVal coords As NetTopologySuite.Geometries.Coordinate(), ByVal boolClosed As Boolean) As Boolean
        If IsNothing(plineCounter) OrElse coords Is Nothing Then Return False

        ' Линии достаточно двух точек (у кольца их минимум 4)
        Dim vertexCount As Integer = coords.Length
        If boolClosed = True Then vertexCount = coords.Length - 1
        If vertexCount < 2 Then Return False

        ' --- 1. Удаляем лишние вершины (если полилиния длиннее) ---
        ' Удаляем с конца, пока не останется ровно vertexCount вершин.
        While plineCounter.NumberOfVertices > vertexCount
            plineCounter.RemoveVertexAt(plineCounter.NumberOfVertices - 1)
        End While
        For j As Integer = 0 To vertexCount - 1
            If j <= plineCounter.NumberOfVertices - 1 Then
                plineCounter.SetPointAt(j, New Point2d(coords(j).X, coords(j).Y))
            Else
                plineCounter.AddVertexAt(j + 1, New Point2d(coords(j).X, coords(j).Y), 0.0, 0.0, 0.0)
            End If
        Next
        ' Замыкание берём из геометрии, а не выставляем всегда True
        plineCounter.Closed = boolClosed
        Return True
    End Function



    '==================================================================================================
    'создание полилинии из линии
    Public Function createPolylineByLineString(ByVal polygon As NetTopologySuite.Geometries.LineString, Optional ByRef polyline As Polyline = Nothing) As Polyline
        createPolylineByLineString = Nothing
        If IsNothing(DBDrawing) = False Then
            If IsNothing(polygon) = False Then
                If polygon.Coordinates.Count > 2 Then
                    Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
                        Dim acBlkTbl As BlockTable = acTrans.GetObject(DBDrawing.BlockTableId, OpenMode.ForRead)
                        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        If IsNothing(polyline) = True Then
                            polyline = New Polyline()
                            polyline.SetDatabaseDefaults()
                        Else
                            While polyline.NumberOfVertices > 0
                                polyline.RemoveVertexAt(0) ' Удаляем первую вершину, пока не останется ни одной
                            End While
                        End If
                        Dim listCoordPolugon As NetTopologySuite.Geometries.Coordinate() = polygon.Coordinates
                        If listCoordPolugon.Length > 1 Then
                            For i As Integer = 0 To listCoordPolugon.Length - 1
                                Dim coord As NetTopologySuite.Geometries.Coordinate = listCoordPolugon(i)
                                polyline.AddVertexAt(0, New Point2d(coord.X, coord.Y), 0, 0, 0)
                            Next i
                            acBlkTblRec.AppendEntity(polyline)
                            acTrans.AddNewlyCreatedDBObject(polyline, True)
                            polyline.Closed = True
                        End If
                        acTrans.Commit()
                    End Using
                End If
            End If
            Return polyline
        End If
    End Function
    '==================================================================================================
    'функция вставляет мультивыноску в чертеж (точка вставки, текст, смещениеX, смещениеY,) новая версия
    Public Function createLeader(ByVal firstVertex As Point3d, ByVal contextText As String, Optional lastVertex As Point3d = Nothing, Optional ByVal idStyleMLeader As ObjectId = Nothing, Optional ByVal NameLayer As String = "0", Optional ByVal M As Double = 1, Optional ByVal StyleTextID As ObjectId = Nothing, Optional ByVal IndexColorText As Double = -1, Optional ByVal Background As Boolean = True) As MLeader
        createLeader = Nothing
        Dim myMLeaderId As ObjectId = ObjectId.Null
        If IsNothing(DBDrawing) = False Then
            Try
                Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
                    Dim acBlkTbl As BlockTable = acTrans.GetObject(DBDrawing.BlockTableId, OpenMode.ForRead)
                    Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim mlStyles As DBDictionary = acTrans.GetObject(DBDrawing.MLeaderStyleDictionaryId, OpenMode.ForRead)
                    Dim myleaderStyleId As ObjectId = DBDrawing.MLeaderstyle
                    Dim mlstyle As MLeaderStyle = TryCast(acTrans.GetObject(myleaderStyleId, OpenMode.ForRead), MLeaderStyle)
                    'создаем новый MLeader
                    Dim myMLeader As MLeader = New MLeader()
                    myMLeader.SetDatabaseDefaults()
                    myMLeaderId = myMLeader.ObjectId
                    'назначаем стиль
                    If IsNothing(idStyleMLeader) = False Then
                        If idStyleMLeader.IsNull = False Then
                            If mlStyles.Contains(idStyleMLeader) = True Then
                                myMLeader.MLeaderStyle = idStyleMLeader
                                mlstyle = TryCast(acTrans.GetObject(idStyleMLeader, OpenMode.ForRead), MLeaderStyle)
                            End If
                        End If
                    End If
                    'назначаем положение
                    Dim myMText As MText = TryCast(mlstyle.DefaultMText.Clone(), MText)
                    Dim leaderIndex As Integer = myMLeader.AddLeader()
                    Dim leaderLineIndex As Integer = myMLeader.AddLeaderLine(leaderIndex)
                    If lastVertex.X = 0 And lastVertex.Y = 0 And lastVertex.Z = 0 Then
                        myMLeader.AddFirstVertex(leaderLineIndex, firstVertex)
                        myMLeader.RemoveLeaderLine(leaderLineIndex)
                        myMText.Location = firstVertex
                    Else
                        myMLeader.AddFirstVertex(leaderLineIndex, firstVertex)
                        myMLeader.AddLastVertex(leaderLineIndex, lastVertex)
                    End If
                    'назначаем содержимое
                    If myMText IsNot Nothing Then
                        If IsNothing(StyleTextID) = False Then
                            If StyleTextID.IsNull = False Then
                                Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(DBDrawing.TextStyleTableId, OpenMode.ForRead)
                                If acTextStyleTbl.Has(StyleTextID) = True Then
                                    Dim acTextStyleRec As TextStyleTableRecord = acTrans.GetObject(StyleTextID, OpenMode.ForRead)
                                    myMText.TextStyleId = StyleTextID
                                End If
                            End If
                        End If
                        myMLeader.ContentType = ContentType.MTextContent
                        myMText.Contents = contextText
                        myMText.TextHeight = myMText.TextHeight * M
                        myMLeader.MText = myMText
                        'характеристики текста
                        If IndexColorText > -1 Then
                            myMText.ColorIndex = IndexColorText
                        End If
                        myMText.BackgroundFill = Background
                    End If
                    Dim acLyrTbl As LayerTable = acTrans.GetObject(DBDrawing.LayerTableId, OpenMode.ForRead)
                    If acLyrTbl.Has(NameLayer) = True Then
                        myMLeader.Layer = NameLayer
                    End If
                    myMLeader.ArrowSize = myMLeader.ArrowSize * M
                    createLeader = myMLeader
                    acBlkTblRec.AppendEntity(myMLeader)
                    acTrans.AddNewlyCreatedDBObject(myMLeader, True)
                    acTrans.Commit()
                End Using
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
            End Try
        End If
    End Function
    '==================================================================================================
    '-===================================================================================
    'вставка в Чертеж блока с атрибутами точки
    Public Shared Function CreateBlock(ByVal insertPosition As Point3d,
                                   ByVal styleBlock As CadastralBlockStyle,
                                   ByVal listAttributes As List(Of AttributeCadastralBlock)) As BlockReference

        If styleBlock Is Nothing Then Return Nothing

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        If doc Is Nothing Then Return Nothing
        Dim db As Database = doc.Database

        Using tr As Transaction = doc.TransactionManager.StartTransaction()
            Try
                Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                If String.IsNullOrWhiteSpace(styleBlock.NameBlock) OrElse Not bt.Has(styleBlock.NameBlock) Then
                    tr.Abort()
                    Return Nothing
                End If

                Dim blockId As ObjectId = bt(styleBlock.NameBlock)
                Dim br As New BlockReference(insertPosition, blockId)

                br.Layer = styleBlock.NameLayer
                br.ScaleFactors = New Scale3d(styleBlock.ScaleBlock, styleBlock.ScaleBlock, styleBlock.ScaleBlock)
                br.ColorIndex = styleBlock.Color
                br.LinetypeScale = styleBlock.ScaleTypeLine
                br.Linetype = styleBlock.TypeLine
                br.LineWeight = styleBlock.WeightLine

                Dim curSpace As BlockTableRecord = DirectCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                curSpace.AppendEntity(br)
                tr.AddNewlyCreatedDBObject(br, True)

                ' --- Атрибуты ---
                Dim btrDef As BlockTableRecord = DirectCast(tr.GetObject(blockId, OpenMode.ForRead), BlockTableRecord)
                If btrDef.HasAttributeDefinitions Then
                    For Each id As ObjectId In btrDef
                        Dim ad As AttributeDefinition = TryCast(tr.GetObject(id, OpenMode.ForRead), AttributeDefinition)
                        If ad Is Nothing Then Continue For

                        ' Ищем пользовательское значение по тегу
                        Dim userAttrib As AttributeCadastralBlock = Nothing
                        If listAttributes IsNot Nothing Then
                            For Each ua In listAttributes
                                If String.Equals(ua.TagAttribute, ad.Tag, StringComparison.OrdinalIgnoreCase) Then
                                    userAttrib = ua
                                    Exit For
                                End If
                            Next
                        End If

                        ' Если для этого атрибута нет данных — не создаём ссылку
                        If userAttrib Is Nothing Then Continue For

                        Dim ar As New AttributeReference()
                        ar.SetAttributeFromBlock(ad, br.BlockTransform)
                        ar.Layer = "0"
                        ar.TextString = If(userAttrib.TextStringAttribute, String.Empty)
                        'ar.Visible = userAttrib.VisibleAttribute

                        ' Локальные координаты внутри блока -> WCS
                        Dim ptWcs As New Point3d(userAttrib.XAttribute, userAttrib.YAttribute, 0)

                        If userAttrib.VisibleAttribute Then
                            ar.Oblique = userAttrib.RotationAttribute
                            ar.Height = userAttrib.HeightAttribute
                            ar.ColorIndex = userAttrib.ColorAttribute

                            If userAttrib.HorizontalAlignment = 1 Then
                                ar.HorizontalMode = TextHorizontalMode.TextLeft
                            Else
                                ar.HorizontalMode = TextHorizontalMode.TextRight
                            End If

                            If userAttrib.VerticalAlignment = 1 Then
                                ar.VerticalMode = TextVerticalMode.TextBottom
                            Else
                                ar.VerticalMode = TextVerticalMode.TextTop
                            End If

                            ar.Position = ptWcs
                            ar.AlignmentPoint = ptWcs
                        Else
                            ar.Position = ptWcs
                        End If

                        br.AttributeCollection.AppendAttribute(ar)
                        tr.AddNewlyCreatedDBObject(ar, True)
                    Next
                End If

                tr.Commit()
                Return br

            Catch ex As System.Exception
                tr.Abort()
                Return Nothing
            End Try
        End Using
    End Function
End Class
