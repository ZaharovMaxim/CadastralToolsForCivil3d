Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class CreatingProfileLineFromPolyline
    '=============================================================================================
    'создание линии профиля из полилинии
    <CommandMethod("CreatingProfileLineFromPolyline")>
    Public Sub CreatingProfileLineFromPolyline()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
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
        Dim ColAlign As ObjectIdCollection = CivilDoc.GetAlignmentIds
        If ColAlign.Count = 0 Then
            MsgBox("На чертеже нет объектов типа ТРАССА. Макро прервано.")
            Exit Sub
        End If
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim count As Integer = 0
        Try
            Using tr
                Dim Align As Alignment = Nothing
                If ColAlign.Count > 1 Then
                    Dim ppr1 As PromptEntityResult = Ed.GetEntity("Выберите трассу: ")
                    Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                    If TypeOf ent1 Is Alignment Then
                        Align = ent1
                    Else
                        MsgBox("Выбранный объект не является трассой. Макро прервано.")
                        Exit Sub
                    End If
                Else
                    Align = tr.GetObject(ColAlign.Item(0), OpenMode.ForRead)
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем вид профиля
                Dim pw As ProfileView = Nothing
                Dim ppr2 As PromptEntityResult = Ed.GetEntity("Выберите вид профиля: ")
                Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr2.ObjectId, OpenMode.ForRead)
                If TypeOf ent2 Is ProfileView Then
                    pw = ent2
                    Dim NameAlign As String = pw.AlignmentName
                    If Not (Align.Name Like NameAlign) Then
                        MsgBox("Данный вид профиля был создан для другой трассы. Макро прервано.")
                        Exit Sub
                    End If
                Else
                    MsgBox("Выбранный объект не является Видом профиля. Макро прервано.")
                    Exit Sub
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'назначаем стиль профиля и стиль меток
                Dim layerId As ObjectId = Align.LayerId
                Dim styleId As ObjectId = CivilDoc.Styles.ProfileStyles.Item(0)
                Dim labelSetId As ObjectId = CivilDoc.Styles.LabelSetStyles.ProfileLabelSetStyles.Item(0)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'назначаем имя профиля
                Dim oProfile As Autodesk.Civil.DatabaseServices.Profile = Nothing
                Dim NameProf As String = ""
                Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & "Введите имя создаваемого профиля: ")
                pStrOpts.AllowSpaces = True
                Dim prPointRes1 As PromptResult = Ed.GetString(pStrOpts)
                If prPointRes1.Status = PromptStatus.OK Then
                    NameProf = prPointRes1.StringResult
                End If
                Try
                    Dim oProfileId As ObjectId = Autodesk.Civil.DatabaseServices.Profile.CreateByLayout(NameProf, Align.Id, layerId, styleId, labelSetId)
                    oProfile = tr.GetObject(oProfileId, OpenMode.ForWrite)
                Catch ex As System.ArgumentException
                    MsgBox("Не верное имя создаваемого профиля либо стиль профиля. Макро прервано.")
                    Exit Sub
                End Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем полилинию
                Dim ppr As PromptEntityResult = Ed.GetEntity("Выберите полилинию: ")
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr.ObjectId, OpenMode.ForRead)
                If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = ent
                    For i As Integer = 0 To Pline.NumberOfVertices - 2
                        Dim segType As SegmentType = Pline.GetSegmentType(i)
                        If segType = SegmentType.Arc Then
                            Dim arcSg As Autodesk.AutoCAD.Geometry.CircularArc2d = Pline.GetArcSegment2dAt(i)
                        ElseIf segType = SegmentType.Line Then
                            Dim sgLn As LineSegment2d = Pline.GetLineSegment2dAt(i)
                            Dim pt1 As Point2d = sgLn.StartPoint
                            Dim PK1 As Double = -1
                            Dim off1 As Double = -1
                            Try
                                pw.FindStationAndElevationAtXY(pt1.X, pt1.Y, PK1, off1)
                            Catch ex As Exception
                            End Try
                            Dim pt2 As Point2d = sgLn.EndPoint
                            Dim PK2 As Double = -1
                            Dim off2 As Double = -1
                            Try
                                pw.FindStationAndElevationAtXY(pt2.X, pt2.Y, PK2, off2)
                            Catch ex As Exception
                            End Try
                            If PK1 <> -1 And PK2 <> -1 And off1 <> -1 And off2 <> -1 Then
                                Dim newPt1 As Point2d = New Point2d(PK1, off1)
                                Dim newPt2 As Point2d = New Point2d(PK2, off2)
                                Dim tg As ProfileTangent = oProfile.Entities.AddFixedTangent(newPt1, newPt2)
                            End If
                        End If
                    Next
                End If
                tr.Commit()
            End Using
        Catch ex As Exception
        End Try
    End Sub
End Class
