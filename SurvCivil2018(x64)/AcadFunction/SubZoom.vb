Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Public Class SubZoom
    Public Shared Sub Zoom(ByVal pMin As Point3d, ByVal pMax As Point3d,
                ByVal pCenter As Point3d, ByVal dFactor As Double)
        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim nCurVport As Integer = System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"))

        '' Get the extents of the current space when no points 
        '' or only a center point is provided
        '' Check to see if Model space is current
        If acCurDb.TileMode = True Then
            If pMin.Equals(New Point3d()) = True And
                pMax.Equals(New Point3d()) = True Then

                pMin = acCurDb.Extmin
                pMax = acCurDb.Extmax
            End If
        Else
            '' Check to see if Paper space is current
            If nCurVport = 1 Then
                If pMin.Equals(New Point3d()) = True And
                   pMax.Equals(New Point3d()) = True Then

                    pMin = acCurDb.Pextmin
                    pMax = acCurDb.Pextmax
                End If
            Else
                '' Get the extents of Model space
                If pMin.Equals(New Point3d()) = True And
                   pMax.Equals(New Point3d()) = True Then

                    pMin = acCurDb.Extmin
                    pMax = acCurDb.Extmax
                End If
            End If
        End If

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Get the current view
            Using acView As ViewTableRecord = acDoc.Editor.GetCurrentView()
                Dim eExtents As Extents3d

                '' Translate WCS coordinates to DCS
                Dim matWCS2DCS As Matrix3d
                matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection)
                matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) * matWCS2DCS
                matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist,
                                               acView.ViewDirection,
                                               acView.Target) * matWCS2DCS

                '' If a center point is specified, define the 
                '' min and max point of the extents
                '' for Center and Scale modes
                If pCenter.DistanceTo(Point3d.Origin) <> 0 Then
                    pMin = New Point3d(pCenter.X - (acView.Width / 2),
                                       pCenter.Y - (acView.Height / 2), 0)

                    pMax = New Point3d((acView.Width / 2) + pCenter.X,
                                       (acView.Height / 2) + pCenter.Y, 0)
                End If

                '' Create an extents object using a line
                Using acLine As Line = New Line(pMin, pMax)
                    eExtents = New Extents3d(acLine.Bounds.Value.MinPoint,
                                             acLine.Bounds.Value.MaxPoint)
                End Using

                '' Calculate the ratio between the width and height of the current view
                Dim dViewRatio As Double
                dViewRatio = (acView.Width / acView.Height)

                '' Tranform the extents of the view
                matWCS2DCS = matWCS2DCS.Inverse()
                eExtents.TransformBy(matWCS2DCS)

                Dim dWidth As Double
                Dim dHeight As Double
                Dim pNewCentPt As Point2d

                '' Check to see if a center point was provided (Center and Scale modes)
                If pCenter.DistanceTo(Point3d.Origin) <> 0 Then
                    dWidth = acView.Width
                    dHeight = acView.Height

                    If dFactor = 0 Then
                        pCenter = pCenter.TransformBy(matWCS2DCS)
                    End If

                    pNewCentPt = New Point2d(pCenter.X, pCenter.Y)
                Else '' Working in Window, Extents and Limits mode
                    '' Calculate the new width and height of the current view
                    dWidth = eExtents.MaxPoint.X - eExtents.MinPoint.X
                    dHeight = eExtents.MaxPoint.Y - eExtents.MinPoint.Y

                    '' Get the center of the view
                    pNewCentPt = New Point2d(((eExtents.MaxPoint.X + eExtents.MinPoint.X) * 0.5),
                                             ((eExtents.MaxPoint.Y + eExtents.MinPoint.Y) * 0.5))
                End If

                '' Check to see if the new width fits in current window
                If dWidth > (dHeight * dViewRatio) Then dHeight = dWidth / dViewRatio

                '' Resize and scale the view
                If dFactor <> 0 Then
                    acView.Height = dHeight * dFactor
                    acView.Width = dWidth * dFactor
                End If

                '' Set the center of the view
                acView.CenterPoint = pNewCentPt

                '' Set the current view
                acDoc.Editor.SetCurrentView(acView)
            End Using

            '' Commit the changes
            acTrans.Commit()
        End Using
    End Sub

End Class
