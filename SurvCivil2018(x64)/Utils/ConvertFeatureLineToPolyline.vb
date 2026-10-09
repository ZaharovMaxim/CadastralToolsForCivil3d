Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.Globalization
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Public Class ConvertFeatureLineToPolyline
    <CommandMethod("ReplaseFeatureLineToPolyline")>
    Public Sub ReplaseFeatureLineToPolyline()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'сохраняем документ
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
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
        Using tr
            Try
                Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                pKeyOpts.Message = vbLf & "В какой объект преобразовать? "
                pKeyOpts.Keywords.Add("Полилиния")
                pKeyOpts.Keywords.Add("3Dполилиния")
                Dim pKeyRes As PromptResult = acDoc.Editor.GetKeywords(pKeyOpts)
                '=========================================================================================================
                'выбираем участки
                Dim rng As Random = New Random()
                Dim acSSPrompt As PromptSelectionResult = Nothing
                acSSPrompt = acDoc.Editor.GetSelection
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is FeatureLine Then
                                    Dim ArrayData As String(,) = Nothing
                                    'считываем расширенные данные (Наборв характеристик)
                                    Dim boolReadData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acEnt, ArrayData)
                                    'доступ к Характерной линии
                                    Dim fLine As Autodesk.Civil.DatabaseServices.FeatureLine = acEnt
                                    Dim Pt3dColliction As Point3dCollection = fLine.GetPoints(FeatureLinePointType.PIPoint)
                                    'Доступ к стилю характерной линии
                                    Dim idStyle As ObjectId = fLine.StyleId
                                    Dim StyleFeatyreLine As Autodesk.Civil.DatabaseServices.Styles.FeatureLineStyle = tr.GetObject(idStyle, OpenMode.ForRead)
                                    Dim DStyle As Styles.DisplayStyle = StyleFeatyreLine.GetFeatureLineDisplayStylePlan 'доступ к стилю отображения в плане
                                    Dim NameLayer As String = DStyle.Layer
                                    Dim NameTypeLine As String = DStyle.Linetype
                                    '====================================================================================================================
                                    'тип линии
                                    If NameTypeLine Like "ByLayer" Or NameLayer Like "ByBlock" Or NameLayer Like "ПоСлою" Or NameLayer Like "ПоБлоку" Then
                                        Dim ArrayLayer As String() = Nothing
                                        Dim boolPropertyLayer As Boolean = AcadFunction.FuncReadLayerProperties(NameLayer, ArrayLayer)
                                        If IsArray(ArrayLayer) = True Then
                                            If ArrayLayer.Length > 5 Then
                                                Dim tempTypeLine As String = ArrayLayer(5)
                                                If IsNothing(tempTypeLine) = False Then
                                                    If tempTypeLine.Trim.Length > 0 Then
                                                        NameTypeLine = tempTypeLine
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    '====================================================================================================================
                                    'цвет
                                    Dim IndexColor As Integer = DStyle.Color.ColorIndex
                                    Dim WLine As Double = DStyle.Lineweight
                                    Dim ScaleLine As Double = DStyle.LinetypeScale

                                    Dim acPoly As Polyline = New Polyline()
                                    Dim acPoly3d As Polyline3d = New Polyline3d()
                                    If pKeyRes.Status Like "Полилиния" Then
                                        acPoly.SetDatabaseDefaults()
                                        Dim CountVertex As Integer = 0
                                        For Each pt3d As Point3d In Pt3dColliction
                                            acPoly.AddVertexAt(CountVertex, New Point2d(pt3d.X, pt3d.Y), 0, 0, 0)
                                            CountVertex += 1
                                        Next
                                        If fLine.Closed = True Then
                                            acPoly.Closed = True
                                        End If
                                        acPoly.Layer = NameLayer
                                        acPoly.Linetype = NameTypeLine
                                        acPoly.ColorIndex = IndexColor
                                        acPoly.LineWeight = WLine
                                        acPoly.LinetypeScale = ScaleLine

                                        acBlkTblRec.AppendEntity(acPoly)
                                        tr.AddNewlyCreatedDBObject(acPoly, True)
                                        If IsArray(ArrayData) = True Then
                                            For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                                Dim NameField As String = ArrayData(1, i)
                                                If NameField Like "LOCALID" Then
                                                    ArrayData(2, i) = rng.Next.ToString
                                                End If
                                            Next i
                                            Dim ArrayDataTables As String() = Nothing
                                            Dim boolReadTable As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayDataTables)
                                            If ArrayDataTables.Length > 0 Then
                                                For i As Integer = 0 To ArrayDataTables.Length - 1
                                                    Dim NameTblPs As String = ArrayDataTables(i)
                                                    Dim idTbl As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(NameTblPs)
                                                    If IsNothing(idTbl) = False Then
                                                        If idTbl.IsNull = False Then
                                                            Dim boolwrite As Boolean = FuncPropertySet.FuncAddDataSetProperties(acPoly, idTbl)
                                                        End If
                                                    End If
                                                Next i
                                            End If
                                            Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acPoly, ArrayData)
                                        End If
                                    Else
                                        acPoly3d.SetDatabaseDefaults()
                                        Dim CountVertex As Integer = 0
                                        For Each pt3d As Point3d In Pt3dColliction
                                            Dim vert3d As PolylineVertex3d = New PolylineVertex3d(pt3d)
                                            acPoly3d.AppendVertex(vert3d)
                                        Next
                                        If fLine.Closed = True Then
                                            acPoly3d.Closed = True
                                        End If
                                        acPoly3d.Layer = NameLayer
                                        acPoly3d.Linetype = NameTypeLine
                                        acPoly3d.ColorIndex = IndexColor
                                        acPoly3d.LineWeight = WLine
                                        acPoly3d.LinetypeScale = ScaleLine

                                        acBlkTblRec.AppendEntity(acPoly3d)
                                        tr.AddNewlyCreatedDBObject(acPoly3d, True)
                                        If IsArray(ArrayData) = True Then
                                            For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                                Dim NameField As String = ArrayData(1, i)
                                                If NameField Like "LOCALID" Then
                                                    ArrayData(2, i) = rng.Next.ToString
                                                End If
                                            Next i
                                            Dim ArrayDataTables As String() = Nothing
                                            Dim boolReadTable As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayDataTables)
                                            If ArrayDataTables.Length > 0 Then
                                                For i As Integer = 0 To ArrayDataTables.Length - 1
                                                    Dim NameTblPs As String = ArrayDataTables(i)
                                                    Dim idTbl As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(NameTblPs)
                                                    If IsNothing(idTbl) = False Then
                                                        If idTbl.IsNull = False Then
                                                            Dim boolwrite As Boolean = FuncPropertySet.FuncAddDataSetProperties(acPoly3d, idTbl)
                                                        End If
                                                    End If
                                                Next i
                                            End If
                                            Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acPoly3d, ArrayData)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            Catch ex As System.Exception
            End Try
            tr.Commit()
        End Using
    End Sub

End Class
