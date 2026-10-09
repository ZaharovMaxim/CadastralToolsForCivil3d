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

Public Class RoadSignEditor
    Private Function DeleteEntityFromBlock2(blockRefId As ObjectId, Optional boolColor As Boolean = False, Optional boolKM As Boolean = False, Optional boolPK As Boolean = False) As Boolean
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = trans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Try
                ' Получаем BlockReference
                Dim blockRef As BlockReference = trans.GetObject(blockRefId, OpenMode.ForWrite)
                ' Получаем BlockTableRecord
                Dim btr As BlockTableRecord = trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite)
                Dim rotateBlk As Double = blockRef.Rotation
                Dim transform As Matrix3d = blockRef.BlockTransform
                Dim insertBlk As Point3d = blockRef.Position
                Dim newInsertPoint As Point3d = New Point3d()
                ' Собираем все примитивы в блоке
                For Each id As ObjectId In btr
                    If Not id.IsNull AndAlso Not id.IsErased Then
                        Dim entity As Entity = TryCast(trans.GetObject(id, OpenMode.ForRead), Entity)
                        If entity IsNot Nothing Then
                            Dim gWidth As Double = 0
                            Dim startPoint As Point3d = New Point3d()
                            Dim endPoint As Point3d = New Point3d
                            Dim indexColor As Integer = 0
                            If TypeOf entity Is Polyline Then
                                Dim poly As Polyline = entity
                                If poly.NumberOfVertices = 2 Then
                                    gWidth = poly.ConstantWidth
                                    startPoint = poly.StartPoint
                                    endPoint = poly.EndPoint
                                    indexColor = poly.ColorIndex
                                Else
                                    Continue For
                                End If
                            ElseIf TypeOf entity Is Line Then
                                Dim acLine As Line = entity
                                startPoint = acLine.StartPoint
                                endPoint = acLine.EndPoint
                                indexColor = acLine.ColorIndex
                            Else
                                Continue For
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'делаем анализ пригодности линии к удалению
                            Dim boolErase As Boolean = False
                            If Math.Abs(startPoint.X) < 0.01 And Math.Abs(startPoint.Y) < 0.01 Then
                                newInsertPoint = endPoint.TransformBy(transform)
                                btr.Origin = endPoint
                                boolErase = True
                            ElseIf Math.Abs(endPoint.X) < 0.01 And Math.Abs(endPoint.Y) < 0.01 Then
                                newInsertPoint = startPoint.TransformBy(transform)
                                btr.Origin = startPoint
                                boolErase = True
                            ElseIf gWidth > 0 And indexColor = 0 Then
                                boolErase = True
                            End If
                            If boolErase = True Then
                                entity.UpgradeOpen()
                                entity.Erase()
                            End If
                        End If
                    End If
                Next
                'удаляем все остальное
                Dim textMleader As String = ""
                For Each id As ObjectId In btr
                    If Not id.IsNull AndAlso Not id.IsErased Then
                        Dim entity As Entity = TryCast(trans.GetObject(id, OpenMode.ForRead), Entity)
                        If entity IsNot Nothing Then
                            Dim stPoint As Point3d = New Point3d(-1, -1, -1)
                            Dim enPoint As Point3d = New Point3d(-1, -1, -1)
                            If TypeOf entity Is DBText Then
                                Dim acText As DBText = entity
                                Dim text As String = acText.TextString
                                If text.IndexOf("КМ") > -1 Then
                                    If boolKM = True Then
                                        entity.UpgradeOpen()
                                        entity.Erase()
                                        Continue For
                                    End If
                                ElseIf text.IndexOf("ПК") > -1 Then
                                    If boolPK = True Then
                                        entity.UpgradeOpen()
                                        entity.Erase()
                                        Continue For
                                    End If
                                End If
                            ElseIf TypeOf entity Is Hatch Then
                                If boolColor = True Then
                                    entity.UpgradeOpen()
                                    If Not (entity.ColorIndex = 255 Or entity.ColorIndex = 256) Then
                                        entity.ColorIndex = 253
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                '===============================================================================================
                'делаем мультивыноску

                blockRef.Position = newInsertPoint
                trans.Commit()
                Return True
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                trans.Abort()
                Return False
            End Try
        End Using
    End Function
    <CommandMethod("RoadSignEditor")>
    Public Sub RoadSignEditor()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database
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
        Dim formRoadSign As FormRoadSignEditor = New FormRoadSignEditor
        formRoadSign.ShowDialog()
        If formRoadSign.boolShow = False Then
            Exit Sub
        End If

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim namelayer As String = formRoadSign.ComboBox2.Text
            Dim idLayer As ObjectId = AcadFunction.FuncAddLayer(namelayer)
            If IsNothing(idLayer) = True Then namelayer = "0"
            If idLayer.IsNull = True Then namelayer = "0"

            Dim idMLeaderStyle As ObjectId = db.MLeaderstyle
            Dim nameStyleMLeader As String = formRoadSign.ComboBox1.Text
            Dim mlStyles As DBDictionary = tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead)

            If mlStyles.Contains(nameStyleMLeader) = True Then
                Dim styleId As ObjectId = mlStyles.GetAt(nameStyleMLeader)
            End If

            Dim ms As BlockTableRecord = DirectCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
            Dim acSSPrompt As PromptSelectionResult = doc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                If acSSet.Count > 0 Then
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Try
                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                                If TypeOf acEnt Is BlockReference Then
                                    Dim blk As BlockReference = acEnt
                                    Dim blkPos As Point3d = blk.Position
                                    Dim blkRot As Double = blk.Rotation
                                    Dim blkScale As Scale3d = blk.ScaleFactors
                                    Dim boolGreyZn As Boolean = formRoadSign.CheckBox1.Checked
                                    Dim boolDeletePK As Boolean = formRoadSign.CheckBox2.Checked
                                    Dim boolDeleteKM As Boolean = formRoadSign.CheckBox3.Checked
                                    Dim boolDelEntity As Boolean = DeleteEntityFromBlock2(blk.Id, boolGreyZn, boolDeleteKM, boolDeletePK)
                                    If boolDelEntity = True Then
                                        Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                                        Dim btrId As ObjectId = bt(blk.Name)
                                        Dim blkPos2 As Point3d = blk.Position
                                        If formRoadSign.RadioButton1.Checked = True Then
                                            ' 2. Создаем MLeader
                                            Dim mleader As New MLeader()
                                            mleader.MLeaderStyle = idMLeaderStyle
                                            mleader.BlockConnectionType = BlockConnectionType.ConnectBase
                                            mleader.ContentType = ContentType.BlockContent
                                            mleader.BlockContentId = btrId
                                            mleader.BlockPosition = New Point3d(blkPos2.X, blkPos2.Y, 0) ' Позиция блока
                                            mleader.BlockRotation = blkRot
                                            mleader.BlockScale = blkScale
                                            mleader.Layer = namelayer
                                            ' Добавляем точки линии-выноски
                                            Dim leaderIdx As Integer = mleader.AddLeader()
                                            Dim leaderLineIdx As Integer = mleader.AddLeaderLine(leaderIdx)
                                            mleader.AddFirstVertex(leaderLineIdx, blkPos)
                                            mleader.AddLastVertex(leaderLineIdx, blkPos2)
                                            ' 3. Добавляем в пространство модели
                                            ms.AppendEntity(mleader)
                                            tr.AddNewlyCreatedDBObject(mleader, True)
                                            blk.Erase()
                                        ElseIf formRoadSign.RadioButton2.Checked = True Then
                                            Dim acLine2 As New Line()
                                            acLine2.SetDatabaseDefaults()
                                            acLine2.StartPoint = blkPos
                                            acLine2.EndPoint = blkPos2
                                            blk.Layer = namelayer
                                            acLine2.Layer = namelayer
                                            ms.AppendEntity(acLine2)
                                            tr.AddNewlyCreatedDBObject(acLine2, True)
                                        Else
                                            blk.Layer = namelayer
                                        End If
                                    End If
                                End If
                            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                            End Try
                        End If
                    Next
                End If
            End If
            tr.Commit()
        End Using
    End Sub
End Class
