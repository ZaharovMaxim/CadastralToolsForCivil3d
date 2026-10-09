Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System

Public Class Triangulate
    Public Function circum(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByRef xc As Double, ByRef yc As Double, ByRef r As Double) As Boolean
        Const eps As Double = 0.000001
        Const big As Double = 1000000000000.0
        Dim result As Boolean = True
        Dim m1, m2, mx1, mx2, my1, my2, dx, dy As Double

        If (Math.Abs(y1 - y2) < eps) AndAlso (Math.Abs(y2 - y3) < eps) Then
            result = False
            xc = x1
            yc = y1
            r = big
        Else

            If Math.Abs(y2 - y1) < eps Then
                m2 = -(x3 - x2) / (y3 - y2)
                mx2 = (x2 + x3) / 2
                my2 = (y2 + y3) / 2
                xc = (x2 + x1) / 2
                yc = m2 * (xc - mx2) + my2
            ElseIf Math.Abs(y3 - y2) < eps Then
                m1 = -(x2 - x1) / (y2 - y1)
                mx1 = (x1 + x2) / 2
                my1 = (y1 + y2) / 2
                xc = (x3 + x2) / 2
                yc = m1 * (xc - mx1) + my1
            Else
                m1 = -(x2 - x1) / (y2 - y1)
                m2 = -(x3 - x2) / (y3 - y2)

                If Math.Abs(m1 - m2) < eps Then
                    result = False
                    xc = x1
                    yc = y1
                    r = big
                Else
                    mx1 = (x1 + x2) / 2
                    mx2 = (x2 + x3) / 2
                    my1 = (y1 + y2) / 2
                    my2 = (y2 + y3) / 2
                    xc = (m1 * mx1 - m2 * mx2 + my2 - my1) / (m1 - m2)
                    yc = m1 * (xc - mx1) + my1
                End If
            End If
        End If

        dx = x2 - xc
        dy = y2 - yc
        r = dx * dx + dy * dy
        Return result
    End Function

    Private Enum OutputObjectType
        PolyFaceMesh = 1
        Solid3d = 2
        SubDMesh = 4
        All = 7
    End Enum

    Private Sub TriangulatePoints(ByVal objType As OutputObjectType, ByVal maxpoints As Integer)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim createSubDMesh As Boolean = (objType And OutputObjectType.SubDMesh) > 0
        Dim createPolyFaceMesh As Boolean = (objType And OutputObjectType.PolyFaceMesh) > 0
        Dim createSolid3d As Boolean = (objType And OutputObjectType.Solid3d) > 0
        Dim tvs As TypedValue() = {New TypedValue(0, "POINT")}
        Dim sf As SelectionFilter = New SelectionFilter(tvs)
        Dim pso As PromptSelectionOptions = New PromptSelectionOptions()
        pso.MessageForAdding = vbLf & "Select points:"
        pso.AllowDuplicates = False
        Dim psr As PromptSelectionResult = ed.GetSelection(pso, sf)
        If psr.Status = PromptStatus.[Error] Then Return
        If psr.Status = PromptStatus.Cancel Then Return
        Dim ss As SelectionSet = psr.Value
        Dim npts As Integer = ss.Count




        If npts < 3 Then
            ed.WriteMessage("Minimum of 3 points must be selected!")
            Return
        End If

        If npts > maxpoints Then
            ed.WriteMessage("Maximum number of points exceeded!")
            Return
        End If

        Dim zref As Double = 0.0

        If createSolid3d Then
            Dim ps As PromptDoubleResult = ed.GetDouble(vbLf & "Enter Z coordinate of reference plane:")
            If ps.Status <> PromptStatus.OK Then Return
            zref = ps.Value
        End If

        Dim i, j, k, ntri, ned, nouted As Integer, status1 As Integer = 0, status2 As Integer = 0
        Dim status As Boolean
        Dim ptx As Double() = New Double(maxpoints + 3 - 1) {}
        Dim pty As Double() = New Double(maxpoints + 3 - 1) {}
        Dim ptz As Double() = New Double(maxpoints + 3 - 1) {}
        Dim pt1 As Integer() = New Integer(maxpoints * 2 + 1 - 1) {}
        Dim pt2 As Integer() = New Integer(maxpoints * 2 + 1 - 1) {}
        Dim pt3 As Integer() = New Integer(maxpoints * 2 + 1 - 1) {}
        Dim cex As Double() = New Double(maxpoints * 2 + 1 - 1) {}
        Dim cey As Double() = New Double(maxpoints * 2 + 1 - 1) {}
        Dim rad As Double() = New Double(maxpoints * 2 + 1 - 1) {}
        Dim xmin, ymin, xmax, ymax, dx, dy, xmid, ymid As Double
        Dim ed1 As Integer() = New Integer(maxpoints * 2 + 1 - 1) {}
        Dim ed2 As Integer() = New Integer(maxpoints * 2 + 1 - 1) {}
        Dim outed1 As Integer() = Nothing
        If createSolid3d Then outed1 = New Integer(maxpoints + 1 - 1) {}
        Dim idarray As ObjectId() = ss.GetObjectIds()
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Dim arrayPt As Point3d() = Nothing
        Dim count As Integer = 0
        For j1 As Integer = 0 To ss.Count - 1
            Dim id As ObjectId = ss.GetObjectIds(j1)
            Dim pt As DBPoint = tr.GetObject(id, OpenMode.ForRead)
            ReDim Preserve arrayPt(count)
            arrayPt(count) = pt.Position
            count += 1
        Next j1
        Using tr

            k = 0
            For i = 0 To arrayPt.GetUpperBound(0)
                Dim pt As Point3d = arrayPt(k)
                ptx(i) = pt.X
                pty(i) = pt.Y
                ptz(i) = pt.Z

                For j = 0 To i - 1

                    If (ptx(i) = ptx(j)) AndAlso (pty(i) = pty(j)) Then
                        i -= 1
                        npts -= 1
                        status2 += 1
                    End If
                Next

                k += 1
            Next

            'For i = 0 To npts - 1
            '    ent = CType(tr.GetObject(idarray(k), OpenMode.ForRead, False), DBPoint)
            '    ptx(i) = ent.Position(0)
            '    pty(i) = ent.Position(1)
            '    ptz(i) = ent.Position(2)

            '    For j = 0 To i - 1

            '        If (ptx(i) = ptx(j)) AndAlso (pty(i) = pty(j)) Then
            '            i -= 1
            '            npts -= 1
            '            status2 += 1
            '        End If
            '    Next

            '    k += 1
            'Next

            tr.Commit()
        End Using

        If status2 > 0 Then ed.WriteMessage(vbLf & "Ignored {0} point(s) with same coordinates.", status2)
        xmin = ptx(0)
        xmax = xmin
        ymin = pty(0)
        ymax = ymin

        For i = 0 To npts - 1
            If ptx(i) < xmin Then xmin = ptx(i)
            If ptx(i) > xmax Then xmax = ptx(i)
            If pty(i) < xmin Then ymin = pty(i)
            If pty(i) > xmin Then ymax = pty(i)
        Next

        dx = xmax - xmin
        dy = ymax - ymin
        xmid = (xmin + xmax) / 2
        ymid = (ymin + ymax) / 2
        i = npts
        ptx(i) = xmid - (90 * (dx + dy)) - 100
        pty(i) = ymid - (50 * (dx + dy)) - 100
        ptz(i) = 0
        pt1(0) = i
        i += 1
        ptx(i) = xmid + (90 * (dx + dy)) + 100
        pty(i) = ymid - (50 * (dx + dy)) - 100
        ptz(i) = 0
        pt2(0) = i
        i += 1
        ptx(i) = xmid
        pty(i) = ymid + 100 * (dx + dy + 1)
        ptz(i) = 0
        pt3(0) = i
        ntri = 1
        circum(ptx(pt1(0)), pty(pt1(0)), ptx(pt2(0)), pty(pt2(0)), ptx(pt3(0)), pty(pt3(0)), cex(0), cey(0), rad(0))

        For i = 0 To npts - 1
            ned = 0
            xmin = ptx(i)
            ymin = pty(i)
            j = 0

            While j < ntri
                dx = cex(j) - xmin
                dy = cey(j) - ymin

                If ((dx * dx) + (dy * dy)) < rad(j) Then
                    ed1(ned) = pt1(j)
                    ed2(ned) = pt2(j)
                    ned += 1
                    ed1(ned) = pt2(j)
                    ed2(ned) = pt3(j)
                    ned += 1
                    ed1(ned) = pt3(j)
                    ed2(ned) = pt1(j)
                    ned += 1
                    ntri -= 1
                    pt1(j) = pt1(ntri)
                    pt2(j) = pt2(ntri)
                    pt3(j) = pt3(ntri)
                    cex(j) = cex(ntri)
                    cey(j) = cey(ntri)
                    rad(j) = rad(ntri)
                    j -= 1
                End If

                j += 1
            End While

            For j = 0 To ned - 1 - 1

                For k = j + 1 To ned - 1

                    If (ed1(j) = ed2(k)) AndAlso (ed2(j) = ed1(k)) Then
                        ed1(j) = -1
                        ed2(j) = -1
                        ed1(k) = -1
                        ed2(k) = -1
                    End If
                Next
            Next

            For j = 0 To ned - 1

                If (ed1(j) >= 0) AndAlso (ed2(j) >= 0) Then
                    pt1(ntri) = ed1(j)
                    pt2(ntri) = ed2(j)
                    pt3(ntri) = i
                    status = circum(ptx(pt1(ntri)), pty(pt1(ntri)), ptx(pt2(ntri)), pty(pt2(ntri)), ptx(pt3(ntri)), pty(pt3(ntri)), cex(ntri), cey(ntri), rad(ntri))

                    If Not status Then
                        status1 += 1
                    End If

                    ntri += 1
                End If
            Next
        Next

        i = 0
        nouted = 0

        While i < ntri

            If (pt1(i) >= npts) OrElse (pt2(i) >= npts) OrElse (pt3(i) >= npts) Then

                If createSolid3d Then

                    If (pt1(i) >= npts) AndAlso (pt2(i) < npts) AndAlso (pt3(i) < npts) Then
                        ed1(nouted) = pt2(i)
                        ed2(nouted) = pt3(i)
                        nouted += 1
                    End If

                    If (pt2(i) >= npts) AndAlso (pt1(i) < npts) AndAlso (pt3(i) < npts) Then
                        ed1(nouted) = pt3(i)
                        ed2(nouted) = pt1(i)
                        nouted += 1
                    End If

                    If (pt3(i) >= npts) AndAlso (pt1(i) < npts) AndAlso (pt2(i) < npts) Then
                        ed1(nouted) = pt1(i)
                        ed2(nouted) = pt2(i)
                        nouted += 1
                    End If
                End If

                ntri -= 1
                pt1(i) = pt1(ntri)
                pt2(i) = pt2(ntri)
                pt3(i) = pt3(ntri)
                cex(i) = cex(ntri)
                cey(i) = cey(ntri)
                rad(i) = rad(ntri)
                i -= 1
            End If

            i += 1
        End While

        If createSolid3d Then
            outed1(0) = 0

            For i = 1 To nouted - 1

                For j = 1 To nouted - 1

                    If ed2(outed1(i - 1)) = ed1(j) Then
                        outed1(i) = j
                        j = nouted
                    End If
                Next
            Next

            outed1(nouted) = 0
        End If

        tr = db.TransactionManager.StartTransaction()

        Using tr
            Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead, False), BlockTable)
            Dim btr As BlockTableRecord = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite, False), BlockTableRecord)

            If createPolyFaceMesh Then
                Dim pfm As PolyFaceMesh = New PolyFaceMesh()
                btr.AppendEntity(pfm)
                tr.AddNewlyCreatedDBObject(pfm, True)

                For i = 0 To npts - 1
                    Dim vert As PolyFaceMeshVertex = New PolyFaceMeshVertex(New Point3d(ptx(i), pty(i), ptz(i)))
                    pfm.AppendVertex(vert)
                    tr.AddNewlyCreatedDBObject(vert, True)
                Next

                For i = 0 To ntri - 1
                    Dim face As FaceRecord = New FaceRecord(CShort((pt1(i) + 1)), CShort((pt2(i) + 1)), CShort((pt3(i) + 1)), 0)
                    pfm.AppendFaceRecord(face)
                    tr.AddNewlyCreatedDBObject(face, True)
                Next
            End If

            If createSubDMesh OrElse createSolid3d Then
                Dim vertarray As Point3dCollection = New Point3dCollection()
                Dim facearray As Int32Collection = New Int32Collection()

                For i = 0 To npts - 1
                    vertarray.Add(New Point3d(ptx(i), pty(i), ptz(i)))
                Next

                If createSolid3d Then

                    For i = 0 To nouted - 1
                        vertarray.Add(New Point3d(ptx(ed1(outed1(i))), pty(ed1(outed1(i))), zref))
                    Next
                End If

                j = 0

                For i = 0 To ntri - 1
                    facearray.Add(3)
                    facearray.Add(pt1(i))
                    facearray.Add(pt2(i))
                    facearray.Add(pt3(i))
                Next

                If createSolid3d Then

                    For i = 0 To nouted - 1
                        facearray.Add(4)
                        k = outed1(i)
                        facearray.Add(ed1(k))
                        facearray.Add(ed2(k))

                        If i = nouted - 1 Then
                            facearray.Add(npts)
                        Else
                            facearray.Add(npts + i + 1)
                        End If

                        facearray.Add(npts + i)
                    Next

                    facearray.Add(nouted)

                    For i = 0 To nouted - 1
                        facearray.Add(npts + i)
                    Next
                End If

                Dim sdm As SubDMesh = New SubDMesh()
                sdm.SetDatabaseDefaults()
                sdm.SetSubDMesh(vertarray, facearray, 0)
                btr.AppendEntity(sdm)
                tr.AddNewlyCreatedDBObject(sdm, True)

                If createSolid3d Then
                    Dim sol As Solid3d = Nothing

                    Try
                        sol = sdm.ConvertToSolid(False, False)
                        btr.AppendEntity(sol)
                        tr.AddNewlyCreatedDBObject(sol, True)
                    Catch
                        ed.WriteMessage(vbLf & "Mesh was too complex to turn into a solid.")
                    End Try

                    If Not createSubDMesh Then sdm.[Erase]()
                End If
            End If

            tr.Commit()
        End Using

        If status1 > 0 Then ed.WriteMessage(vbLf & "Warning! {0} thin triangle(s) found!" & " Wrong result possible!", status1)
        Application.UpdateScreen()
    End Sub

    <CommandMethod("PFT")>
    Public Sub PolyFaceTriangulate()
        TriangulatePoints(OutputObjectType.PolyFaceMesh, 32767)
    End Sub

    <CommandMethod("SDT")>
    Public Sub SubDTriangulate()
        TriangulatePoints(OutputObjectType.SubDMesh, 200000)
    End Sub

    <CommandMethod("S3T")>
    Public Sub Solid3dTriangulate()
        TriangulatePoints(OutputObjectType.Solid3d, 200000)
    End Sub
End Class

