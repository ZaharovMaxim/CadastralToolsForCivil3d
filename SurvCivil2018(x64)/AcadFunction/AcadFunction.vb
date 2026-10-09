Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.IO
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD
Imports System.Globalization
Imports Autodesk.AutoCAD.EditorInput
Imports System.Text.RegularExpressions

Public Class AcadFunction
    '==================================================================================================================
    'работа с полилинией
    Structure PerVertexData
        Public pt As Autodesk.AutoCAD.Geometry.Point2d
        Public bulge As Double
        Public startWidth As Double
        Public endWidth As Double
    End Structure
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ОБЩИЕ ФУНКЦИИ
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '==================================================================================
    'функция по Hangle объекта возвращает его ObjectID
    Public Shared Function FuncReturnObjectID(ByVal HandleString As String) As ObjectId
        FuncReturnObjectID = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        If IsNothing(HandleString) = False Then
            If HandleString.Trim.Length > 1 Then
                Try
                    'получаем доступ к полилиниии в которую требуется внести изменения
                    Dim ln As Long = Convert.ToInt64(HandleString, 16)
                    Dim hn As Handle = New Handle(ln)
                    FuncReturnObjectID = db.GetObjectId(False, hn, 0)
                Catch ex As System.Exception
                End Try
            End If
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'работа с масштабами
    '-=====================================================================================
    '-Функция возвращает текущий масштаб аннотаций
    Public Shared Function FuncReadCurrentAnnotativeScaleDWG() As Double
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim anno As AnnotationScale = db.Cannoscale
        FuncReadCurrentAnnotativeScaleDWG = (1 / anno.Scale) * 1000
    End Function
    '=====================================================================================
    'Функция читает аннотации и возвращает массив с именами (название масштаба)
    Public Shared Function FuncReadNameAnnoScaleDWG(ByRef ScaleArray() As String) As Boolean
        FuncReadNameAnnoScaleDWG = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim CountScaleArray As Integer = 0
        Try
            Dim ocm As ObjectContextManager = db.ObjectContextManager
            Dim occ As ObjectContextCollection = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES")
            For Each Anno As AnnotationScale In occ
                ReDim Preserve ScaleArray(CountScaleArray)
                ScaleArray(CountScaleArray) = Anno.Name
                CountScaleArray += 1
            Next
            FuncReadNameAnnoScaleDWG = True
        Catch ex As Exception
        End Try
    End Function
    '-=====================================================================================
    '-Функция читает аннотации и возвращает массив с именами (название масштаба, еденицы листа, еденицы чертежа)
    Public Shared Function FuncReadNameAndUnitAnnotativeScale(ByRef ScaleArray(,) As String) As Boolean
        FuncReadNameAndUnitAnnotativeScale = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim CountScaleArray As Integer = 0
        Try
            Dim ocm As ObjectContextManager = db.ObjectContextManager
            Dim occ As ObjectContextCollection = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES")
            For Each Anno As AnnotationScale In occ
                ReDim Preserve ScaleArray(2, CountScaleArray)
                ScaleArray(0, CountScaleArray) = Anno.Name
                ScaleArray(1, CountScaleArray) = Anno.PaperUnits
                ScaleArray(2, CountScaleArray) = Anno.DrawingUnits
                CountScaleArray += 1
            Next
            FuncReadNameAndUnitAnnotativeScale = True
        Catch ex As Exception
        End Try
    End Function
    '=============================================================================
    'функция получает масштаб аннотаций выбранного чертежа
    Public Shared Function FuncReadCurrentAnnoDestDraw(ByVal PutshDraw As String) As Double
        FuncReadCurrentAnnoDestDraw = 1000
        Dim NameFile As String = ""
        Dim NameFullPatch As String = ""
        Try
            NameFile = Path.GetFileName(PutshDraw)
            NameFullPatch = Path.GetDirectoryName(PutshDraw)
        Catch ex As ArgumentException
            Return 0
        End Try
        Try
            Dim sourceFileName As String = Path.Combine(NameFullPatch, NameFile)
            Dim sourceDb As Database = New Database(False, True)
            sourceDb.ReadDwgFile(sourceFileName, FileShare.Read, True, String.Empty)
            Dim anno As AnnotationScale = sourceDb.Cannoscale
            FuncReadCurrentAnnoDestDraw = (1 / anno.Scale) * 1000
            sourceDb.CloseInput(True)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try
    End Function
    '-=====================================================================================
    '-Функция возвращает масштаб аннотации по его имени
    Public Shared Function FuncReadAnnotativeScaleDWGByName(ByVal NameAnnotativeScale As String) As AnnotationScale
        FuncReadAnnotativeScaleDWGByName = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Try
            Dim ocm As ObjectContextManager = db.ObjectContextManager
            Dim occ As ObjectContextCollection = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES")
            For Each Anno As AnnotationScale In occ
                If Anno.Name Like NameAnnotativeScale Then
                    FuncReadAnnotativeScaleDWGByName = Anno
                    Exit For
                End If
            Next
        Catch ex As System.Exception
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '==============================================================================================================================
    'работа со слоями
    'функция для вставки нового слоя
    Public Shared Function FuncAddLayer(ByVal LayerStr As String, Optional ByVal color1 As Short = Nothing, Optional ByVal idlinetype As ObjectId = Nothing, Optional ByVal wightline As Integer = Nothing) As ObjectId
        FuncAddLayer = Nothing
        Try
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acLyrTbl As LayerTable
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                Try
                    If acLyrTbl.Has(LayerStr) = False Then
                        Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()
                        acLyrTblRec.Name = LayerStr 'имя слоя
                        'цвет слоя
                        If IsNothing(color1) = False Then
                            Dim color2 As Autodesk.AutoCAD.Colors.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, color1)
                            acLyrTblRec.Color = color2
                        End If
                        'тип линии
                        If IsNothing(idlinetype) = False Then
                            acLyrTblRec.LinetypeObjectId = idlinetype
                        End If
                        'вес линии
                        If IsNothing(wightline) = False Then
                            acLyrTblRec.LineWeight = wightline
                        End If
                        acLyrTbl.UpgradeOpen()
                        acLyrTbl.Add(acLyrTblRec)
                        acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)
                        FuncAddLayer = acLyrTblRec.Id
                    Else
                        Return acLyrTbl.Item(LayerStr)
                    End If
                Catch
                    FuncAddLayer = Nothing
                End Try
                acTrans.Commit()
            End Using
        Catch
            FuncAddLayer = Nothing
        End Try
    End Function
    '========================================================================================
    'функция возвращает id стиля текста по его имени
    Public Shared Function FuncReturnIDTextStylebyName(ByVal NameTextStyle As String) As ObjectId
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        FuncReturnIDTextStylebyName = Nothing
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead)
                For Each id As ObjectId In acTextStyleTbl
                    Dim acTextStyleRec As TextStyleTableRecord = acTrans.GetObject(id, OpenMode.ForRead)
                    Dim NameStyle As String = acTextStyleRec.Name.Trim
                    If NameTextStyle Like NameStyle Then
                        FuncReturnIDTextStylebyName = id
                        Exit Function
                    End If
                Next
            End Using
        Catch ex As System.Exception
            Exit Function
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает возвращает id стиля мультивыноски по его имени
    Public Shared Function FuncReturnIDMLeaderStylebyName(ByVal NameMLeaderStyle As String) As ObjectId
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        FuncReturnIDMLeaderStylebyName = Nothing
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim mlstyles As DBDictionary = acTrans.GetObject(acCurDb.MLeaderStyleDictionaryId, OpenMode.ForRead)
                FuncReturnIDMLeaderStylebyName = mlstyles.Item(NameMLeaderStyle)
            End Using
        Catch ex As System.Exception
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает id стрелки мультивыноски по его имени
    Public Shared Function FuncGetArrowObjectId(ByVal newArrName As String) As ObjectId
        Dim arrObjId As ObjectId = ObjectId.Null
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim oldArrName As String = TryCast(Application.GetSystemVariable("DIMBLK"), String)
        Application.SetSystemVariable("DIMBLK", newArrName)
        If oldArrName.Length <> 0 Then Application.SetSystemVariable("DIMBLK", oldArrName)
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            arrObjId = bt(newArrName)
            tr.Commit()
        End Using
        Return arrObjId
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает id типа линии
    Public Shared Function FuncGetLineTypeId(ByVal nameLineType As String) As ObjectId
        FuncGetLineTypeId = Nothing
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Try
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim lt As LinetypeTable = CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead), LinetypeTable)
                Dim acTextStyleTbl As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
                Dim idTextStyle As ObjectId = Nothing
                If lt.Has(nameLineType) = True Then
                    Dim idTypeLine As ObjectId = lt.Item(nameLineType)
                    Return idTypeLine
                End If
            End Using
        Catch ex As System.Exception
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает код типа линии
    Public Shared Function FuncReturnTypeLineCode(ByVal recTypeLine As LinetypeTableRecord, Optional supportPatchFolder As String() = Nothing) As String
        FuncReturnTypeLineCode = "A,"
        Dim Str As String = "A,"
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'находим все штрихи и точки
            Dim lenD As Integer = recTypeLine.NumDashes
            If lenD > 0 Then
                Try
                    For i As Integer = 0 To lenD - 1
                        Dim l As Double = 0
                        Dim indexShape As Integer = 0
                        Dim textVal As String = ""
                        Dim nameShape As String = ""
                        'находим значения
                        l = recTypeLine.DashLengthAt(i) 'длина штриха
                        indexShape = recTypeLine.ShapeNumberAt(i) 'индекс формы (0 - нет формы)
                        Try
                            textVal = recTypeLine.TextAt(i) 'читаем текст
                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                        End Try
                        If indexShape > 0 Then
                            Dim idTextStyleShape As ObjectId = recTypeLine.ShapeStyleAt(i)
                            Dim textStyleShape As TextStyleTableRecord = tr.GetObject(idTextStyleShape, OpenMode.ForRead)
                            If textStyleShape.IsShapeFile = True Then
                                Using acTrans As Transaction = db.TransactionManager.StartTransaction()
                                    Dim Btr As BlockTableRecord = acTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, False)
                                    Dim pShape As Autodesk.AutoCAD.DatabaseServices.Shape = New Autodesk.AutoCAD.DatabaseServices.Shape(Point3d.Origin, 1, 0, 1)
                                    pShape.ShapeNumber = indexShape
                                    pShape.StyleId = idTextStyleShape
                                    Btr.AppendEntity(pShape)
                                    nameShape = pShape.Name
                                End Using
                                'имя формы не найдено, ищем имя формы в файле shp
                                If nameShape.Trim.Length = 0 Then
                                    If IsArray(supportPatchFolder) = False Then
                                        Dim boolFindFolderPatch As Boolean = FuncFiles.FuncGetSupportFolderPath(supportPatchFolder)
                                    End If
                                    If IsArray(supportPatchFolder) = True Then
                                        Dim nameFileSHP As String = textStyleShape.FileName.Replace(".shx", ".shp")
                                        Dim boolFindShp As Boolean = False
                                        Dim regMatch As String = "\*"
                                        For kd As Integer = 0 To supportPatchFolder.Length - 1
                                            Dim searFiles As String() = Directory.GetFiles(supportPatchFolder(kd), nameFileSHP)
                                            If IsArray(searFiles) = True Then
                                                If searFiles.Length > 0 Then
                                                    'читаем файл с формой
                                                    Dim ffileShp As String = searFiles(0)
                                                    If IO.File.Exists(ffileShp) = True Then
                                                        Dim sr As StreamReader = New StreamReader(ffileShp)
                                                        While sr.Peek <> -1
                                                            Dim st As String = sr.ReadLine().Trim
                                                            Dim flagread As String = Left(st, 1)
                                                            If Regex.IsMatch(flagread, regMatch) Then 'chr(42)
                                                                Dim rez As String() = st.Split(",")
                                                                If rez.GetUpperBound(0) > 1 Then
                                                                    Dim indForm As String = Mid(rez(0), 2)
                                                                    If indForm Like indexShape Then
                                                                        nameShape = rez(2)
                                                                        boolFindShp = True
                                                                        Exit While
                                                                    End If
                                                                End If
                                                            End If
                                                        End While
                                                        sr.Close()
                                                        If boolFindShp = True Then
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Next kd
                                    End If
                                End If
                            End If
                            Str = Str & l & ",[" & nameShape & "," & textStyleShape.FileName & ",S=" & recTypeLine.ShapeScaleAt(i) & ",R=" & recTypeLine.ShapeRotationAt(i) & ",X=" & recTypeLine.ShapeOffsetAt(i).X & ",Y=" & recTypeLine.ShapeOffsetAt(i).Y & "]"
                        ElseIf textVal.Trim.Length > 0 Then
                            Dim id As ObjectId = recTypeLine.ShapeStyleAt(i)
                            Dim acShape As TextStyleTableRecord = tr.GetObject(id, OpenMode.ForRead)
                            Dim styleName As String = acShape.Name
                            Str = Str & l & ",[" & Chr(34) & textVal & Chr(34) & "," & styleName & ",S=" & recTypeLine.ShapeScaleAt(i) & ",R=" & recTypeLine.ShapeRotationAt(i) & ",X=" & recTypeLine.ShapeOffsetAt(i).X & ",Y=" & recTypeLine.ShapeOffsetAt(i).Y & "]"
                        Else
                            Str = Str & l
                        End If
                        If i < lenD - 1 Then
                            Str = Str & ","
                        End If
                    Next i
                    FuncReturnTypeLineCode = Str
                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                End Try
            End If
        End Using
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция создает тип линии
    Public Shared Function FuncCreateTypeLineByCode(ByVal nameTypeLine As String, ByVal deskTypeLine As String, ByVal patternStr As String) As LinetypeTableRecord
        FuncCreateTypeLineByCode = Nothing
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim patternLine As String() = patternStr.Split(",")
        Dim supportPatchFolder As String() = Nothing
        Dim boolFindSupport As Boolean = FuncFiles.FuncGetSupportFolderPath(supportPatchFolder)
        If IsArray(patternLine) = True Then
            If patternLine.Length > 1 Then
                Dim lenghtPattern As Double = 0
                Dim numDashes As Integer = 0
                For i As Integer = 0 To patternLine.Length - 1
                    Dim valZnStr As String = patternLine(i)
                    If IsNumeric(valZnStr) Then
                        Dim zn As Double = Math.Abs(CDbl(valZnStr))
                        If zn > 0 Then
                            lenghtPattern = lenghtPattern + zn
                            numDashes += 1
                        End If
                    End If
                Next i
                Dim tr As Transaction = db.TransactionManager.StartTransaction()
                Using tr
                    Dim lt As LinetypeTable = CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForWrite), LinetypeTable)
                    Dim acTextStyleTbl As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
                    Dim idTextStyle As ObjectId = Nothing
                    If lt.Has(nameTypeLine) = True Then
                        Dim idTypeLine As ObjectId = lt.Item(nameTypeLine)
                        FuncCreateTypeLineByCode = idTypeLine.GetObject(OpenMode.ForRead)
                    Else
                        Dim ltr As LinetypeTableRecord = New LinetypeTableRecord()
                        ltr.Name = nameTypeLine
                        ltr.AsciiDescription = deskTypeLine
                        ltr.NumDashes = numDashes
                        ltr.PatternLength = lenghtPattern
                        Dim index As Integer = 0

                        For i As Integer = 1 To patternLine.Length - 1
                            Dim tempstr As String = patternLine(i).Trim
                            If IsNumeric(tempstr) = True Then
                                ltr.SetDashLengthAt(index, Val(tempstr))
                                index += 1
                            Else
                                '==================================================================
                                'сложный тип линии
                                If Left(tempstr, 1) Like "[[]" Then
                                    tempstr = Mid(tempstr, 2)
                                    tempstr = tempstr.Replace("""", "")
                                    Dim indexForm As Integer = 0
                                    Dim namefile As String = ""
                                    Dim S As Double = 1
                                    Dim R As Double = 0
                                    Dim X As Double = 0
                                    Dim Y As Double = 0
                                    If index > 0 Then
                                        index -= 1
                                    End If

                                    For j As Integer = i + 1 To patternLine.Length - 1
                                        Dim tempStr2 As String = patternLine(j).Trim
                                        Dim boolEnd As Boolean = False
                                        If Right(tempStr2, 1) Like "[]]" Then
                                            boolEnd = True
                                            tempStr2 = Mid(tempStr2, 1, tempStr2.Length - 2)
                                        End If
                                        If Left(tempStr2, 2) Like "S=" Or Left(tempStr2, 2) Like "s=" Then
                                            S = Val(Mid(tempStr2, 3))
                                        ElseIf Left(tempStr2, 2) Like "R=" Or Left(tempStr2, 2) Like "r=" Then
                                            R = Val(Mid(tempStr2, 3))
                                        ElseIf Left(tempStr2, 2) Like "U=" Or Left(tempStr2, 2) Like "u=" Then
                                            R = Val(Mid(tempStr2, 3))
                                        ElseIf Left(tempStr2, 2) Like "X=" Or Left(tempStr2, 2) Like "x=" Then
                                            X = Val(Mid(tempStr2, 3))
                                        ElseIf Left(tempStr2, 2) Like "Y=" Or Left(tempStr2, 2) Like "y=" Then
                                            Y = Val(Mid(tempStr2, 3))
                                        Else
                                            namefile = tempStr2
                                            If Right(namefile, 4) Like ".shx" Then
                                                For Each id As ObjectId In acTextStyleTbl
                                                    Dim recTextStyle As TextStyleTableRecord = id.GetObject(OpenMode.ForRead)
                                                    If recTextStyle.FileName Like namefile Then
                                                        idTextStyle = id
                                                        Exit For
                                                    End If
                                                Next
                                                If IsNothing(idTextStyle) = True OrElse idTextStyle.IsNull = True Then
                                                    idTextStyle = FuncCreateTextStyle("", namefile, False)
                                                End If
                                                Dim FindFile As String = namefile.Replace("shx", "shp")
                                                If IsArray(supportPatchFolder) = True Then
                                                    For k As Integer = 0 To supportPatchFolder.Length - 1
                                                        Dim namefiletemp As String = supportPatchFolder(k) & "\" & FindFile
                                                        If File.Exists(namefiletemp) = True Then
                                                            'ищем индекс формы
                                                            Dim inputshp As StreamReader = New StreamReader(namefiletemp, True)
                                                            Dim line2 As String
                                                            Dim regMatch As String = "\*"
                                                            Do Until inputshp.EndOfStream
                                                                line2 = inputshp.ReadLine().Trim 'считываем строку
                                                                If line2.Length > 0 Then
                                                                    Dim flagread As String = Left(line2, 1)
                                                                    If Regex.IsMatch(flagread, regMatch) Then 'chr(*)
                                                                        Dim arrayShp As String() = line2.Split(",")
                                                                        If IsArray(arrayShp) = True Then
                                                                            If arrayShp.GetUpperBound(0) > 1 Then
                                                                                If arrayShp(2).Trim Like tempstr Then
                                                                                    Dim StrNumb As String = arrayShp(0).Trim
                                                                                    If StrNumb.Length > 1 Then
                                                                                        Dim Numbshp As Integer = Val(Mid(StrNumb, 2))
                                                                                        If Numbshp > 0 Then
                                                                                            indexForm = Numbshp
                                                                                            Exit Do
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            Loop
                                                            inputshp.Close()
                                                        End If
                                                    Next k
                                                End If
                                            Else
                                                'ищем текстовый стиль
                                                idTextStyle = FuncReturnIDTextStylebyName(namefile)
                                                If IsNothing(idTextStyle) = True OrElse idTextStyle.IsNull = True Then
                                                    'создаем новый текстовый стиль
                                                    idTextStyle = FuncCreateTextStyle(namefile, "", False)
                                                End If
                                            End If
                                        End If
                                        If boolEnd = True Then
                                            i = j
                                            Exit For
                                        End If
                                    Next j
                                    If indexForm > 0 Then
                                        ltr.SetShapeNumberAt(index, indexForm)
                                    Else
                                        ltr.SetTextAt(index, tempstr)
                                    End If
                                    ltr.SetShapeStyleAt(index, idTextStyle)
                                    ltr.SetShapeScaleAt(index, Val(S))
                                    ltr.SetShapeRotationAt(index, Val(R))
                                    ltr.SetShapeOffsetAt(index, New Vector2d(X, Y))
                                    index += 1
                                    idTextStyle = Nothing
                                End If
                            End If
                        Next i
                        Dim ltId As ObjectId = lt.Add(ltr)
                        tr.AddNewlyCreatedDBObject(ltr, True)
                        FuncCreateTypeLineByCode = ltr
                    End If
                    tr.Commit()
                End Using
            End If
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'КОПИРОВАНИЕ СТИЛЕЙ ИЗ ЧЕРТЕЖА-ШАБЛОНА
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '=================================================================================================================================
    'функция копирует стили (LineType, MLineType, TextStyle, MLeaderType, DimStyle, TableStyle, PSTable, Block, MVBlock, Layer)
    Public Shared Function FuncCopyALLStylePSDestDraw(ByVal PathDraw As String, ByVal ArrayStyle As String(,)) As Boolean
        FuncCopyALLStylePSDestDraw = False
        If File.Exists(PathDraw) = False Then Exit Function
        If IsArray(ArrayStyle) = False Then Exit Function
        Try
            Dim sourceDb As Database = New Database(False, True)
            sourceDb.ReadDwgFile(PathDraw, FileShare.Read, True, String.Empty)
            Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
            Using sourceTr
                For i As Integer = 0 To ArrayStyle.GetUpperBound(1)
                    Try
                        Dim NameStyle As String = ArrayStyle(0, i)
                        If IsNothing(NameStyle) = True Then
                            Continue For
                        ElseIf NameStyle.Trim.Length = 0 Then
                            Continue For
                        Else
                            NameStyle = NameStyle.Trim
                        End If
                        Dim TypeObject As String = ArrayStyle(1, i)
                        If IsNothing(TypeObject) = True Then
                            Continue For
                        ElseIf TypeObject.Trim.Length = 0 Then
                            Continue For
                        Else
                            TypeObject = TypeObject.Trim
                        End If
                        Dim boolReplase As Boolean = Val(ArrayStyle(2, i))
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        If TypeObject Like "LineType" Then
                            Dim acLineTypTbl As LinetypeTable = sourceTr.GetObject(sourceDb.LinetypeTableId, OpenMode.ForRead)
                            If acLineTypTbl.Has(NameStyle) = True Then
                                Dim IDTypeLine As ObjectId = acLineTypTbl.Item(NameStyle)
                                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                ObjColID.Add(IDTypeLine)
                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                Dim db1 As Database = New Database(False, False)
                                Dim destDb1 As Database = acDoc.Database
                                Dim destDictId1 As ObjectId = destDb1.LinetypeTableId
                                Dim iMap1 As IdMapping = New IdMapping()
                                If boolReplase = True Then
                                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                Else
                                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                End If
                            End If
                        ElseIf TypeObject Like "MLineType" Then
                            Dim mlDict As DBDictionary = sourceTr.GetObject(sourceDb.MLStyleDictionaryId, OpenMode.ForRead)
                            If mlDict.Contains(NameStyle) = True Then
                                Dim IDTypeMLine As ObjectId = mlDict.Item(NameStyle)
                                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                ObjColID.Add(IDTypeMLine)
                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                Dim db1 As Database = New Database(False, False)
                                Dim destDb1 As Database = acDoc.Database
                                Dim destDictId1 As ObjectId = destDb1.MLStyleDictionaryId
                                Dim iMap1 As IdMapping = New IdMapping()
                                If boolReplase = True Then
                                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                Else
                                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                End If
                            End If
                        ElseIf TypeObject Like "TextStyle" Then
                            Dim acTextStyleTbl As TextStyleTable = sourceTr.GetObject(sourceDb.TextStyleTableId, OpenMode.ForRead)
                            If acTextStyleTbl.Has(NameStyle) = True Then
                                Dim IDTextStyle As ObjectId = acTextStyleTbl.Item(NameStyle)
                                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                ObjColID.Add(IDTextStyle)
                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                Dim db1 As Database = New Database(False, False)
                                Dim destDb1 As Database = acDoc.Database
                                Dim destDictId1 As ObjectId = destDb1.TextStyleTableId
                                Dim iMap1 As IdMapping = New IdMapping()
                                If boolReplase = True Then
                                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                Else
                                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                End If
                            End If
                        ElseIf TypeObject Like "MLeaderType" Then
                            Dim acMleaderStyleTbl As DBDictionary = sourceTr.GetObject(sourceDb.MLeaderStyleDictionaryId, OpenMode.ForRead)
                            Dim MLeaderStyleCollection As ObjectIdCollection = New ObjectIdCollection
                            For Each id As DBDictionaryEntry In acMleaderStyleTbl
                                If id.Key Like NameStyle Then
                                    MLeaderStyleCollection.Add(id.Value)
                                    Exit For
                                End If
                            Next
                            If MLeaderStyleCollection.Count > 0 Then
                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                Dim db1 As Database = New Database(False, False)
                                Dim destDb1 As Database = acDoc.Database
                                Dim destDictId1 As ObjectId = destDb1.MLeaderStyleDictionaryId
                                Dim iMap1 As IdMapping = New IdMapping()
                                If boolReplase = True Then
                                    destDb1.WblockCloneObjects(MLeaderStyleCollection, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                Else
                                    destDb1.WblockCloneObjects(MLeaderStyleCollection, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                End If
                            End If
                        ElseIf TypeObject Like "DimStyle" Then
                            Dim acDimStyleTbl As DimStyleTable = sourceTr.GetObject(sourceDb.DimStyleTableId, OpenMode.ForRead)
                            If acDimStyleTbl.Has(NameStyle) = True Then
                                Dim IDDimStyle As ObjectId = acDimStyleTbl.Item(NameStyle)
                                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                ObjColID.Add(IDDimStyle)
                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                Dim db1 As Database = New Database(False, False)
                                Dim destDb1 As Database = acDoc.Database
                                Dim destDictId1 As ObjectId = destDb1.DimStyleTableId
                                Dim iMap1 As IdMapping = New IdMapping()
                                If boolReplase = True Then
                                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                Else
                                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                End If
                            End If
                        ElseIf TypeObject Like "TableStyle" Then
                            Dim acTableStyleTbl As DBDictionary = sourceTr.GetObject(sourceDb.TableStyleDictionaryId, OpenMode.ForRead)
                            Dim TableStyleCollection As ObjectIdCollection = New ObjectIdCollection
                            For Each id As DBDictionaryEntry In acTableStyleTbl
                                If id.Key Like NameStyle Then
                                    TableStyleCollection.Add(id.Value)
                                    Exit For
                                End If
                            Next
                            If TableStyleCollection.Count > 0 Then
                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                Dim db1 As Database = New Database(False, False)
                                Dim destDb1 As Database = acDoc.Database
                                Dim destDictId1 As ObjectId = destDb1.TableStyleDictionaryId
                                Dim iMap1 As IdMapping = New IdMapping()
                                If boolReplase = True Then
                                    destDb1.WblockCloneObjects(TableStyleCollection, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                Else
                                    destDb1.WblockCloneObjects(TableStyleCollection, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                End If
                            End If

                        ElseIf TypeObject Like "PSTable" Then
                            Dim dictPropSetDef = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(sourceDb)
                            If dictPropSetDef.Has(NameStyle, sourceTr) Then
                                Dim IDPSTable As ObjectId = dictPropSetDef.GetAt(NameStyle)
                                Dim CollDictPropSet As ObjectIdCollection = New ObjectIdCollection
                                CollDictPropSet.Add(IDPSTable)
                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                Dim destDb As Database = acDoc.Database
                                Dim destTr As Transaction = destDb.TransactionManager.StartTransaction()
                                Using destTr
                                    Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(destDb)
                                    If dictPropSetDefTemp.Has(NameStyle, destTr) = False Then
                                        Dim helpme As Autodesk.AEC.ApplicationServices.Utility.CloningHelper = New Autodesk.AEC.ApplicationServices.Utility.CloningHelper
                                        Dim RXClassPS As RXClass = RXClass.GetClass(GetType(Autodesk.AEC.PropertyData.DatabaseServices.PropertySet))
                                        helpme.Clone(sourceDb, destDb, CollDictPropSet, RXClassPS, True)
                                    End If
                                    destTr.Commit()
                                End Using
                            End If
                        ElseIf TypeObject Like "Block" Then
                            Dim acBlockTbl As BlockTable = sourceTr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
                            If acBlockTbl.Has(NameStyle) = True Then
                                Dim IDBlk As ObjectId = acBlockTbl.Item(NameStyle)
                                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                ObjColID.Add(IDBlk)
                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                Dim destDb1 As Database = acDoc.Database
                                Dim destDictId1 As ObjectId = destDb1.BlockTableId
                                Dim iMap1 As IdMapping = New IdMapping()
                                If boolReplase = True Then
                                    destDb1.WblockCloneObjects(ObjColID, destDb1.BlockTableId, iMap1, DuplicateRecordCloning.Replace, False)
                                Else
                                    destDb1.WblockCloneObjects(ObjColID, destDb1.BlockTableId, iMap1, DuplicateRecordCloning.Ignore, False)
                                End If
                            End If
                        ElseIf TypeObject Like "MVBlock" Then
                            Dim acBlockTbl As BlockTable = sourceTr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
                            Dim mvbstyledict As Autodesk.AEC.DatabaseServices.DictionaryMultiViewBlockDefinition = New Autodesk.AEC.DatabaseServices.DictionaryMultiViewBlockDefinition(sourceDb)
                            If mvbstyledict.Has(NameStyle, sourceTr) = True Then
                                Dim bstyleID As ObjectId = mvbstyledict.GetAt(NameStyle)
                                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                ObjColID.Add(bstyleID)
                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                Dim destDb As Database = acDoc.Database
                                Dim destmvbstyledict As Autodesk.AEC.DatabaseServices.DictionaryMultiViewBlockDefinition = New Autodesk.AEC.DatabaseServices.DictionaryMultiViewBlockDefinition(destDb)
                                Dim iMap1 As IdMapping = New IdMapping()
                                If boolReplase = True Then
                                    destDb.WblockCloneObjects(ObjColID, destmvbstyledict.DictionaryId, iMap1, DuplicateRecordCloning.Replace, False)
                                Else
                                    destDb.WblockCloneObjects(ObjColID, destmvbstyledict.DictionaryId, iMap1, DuplicateRecordCloning.Ignore, False)
                                End If
                            End If
                        ElseIf TypeObject Like "Layer" Then
                            Dim acLyrTbl As LayerTable = sourceTr.GetObject(sourceDb.LayerTableId, OpenMode.ForRead)
                            If acLyrTbl.Has(NameStyle) = True Then
                                Dim IDLayer As ObjectId = acLyrTbl.Item(NameStyle)
                                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                ObjColID.Add(IDLayer)
                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                Dim destDb1 As Database = acDoc.Database
                                Dim destDictId1 As ObjectId = destDb1.LayerTableId
                                Dim iMap1 As IdMapping = New IdMapping()
                                If boolReplase = True Then
                                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                Else
                                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                End If
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                Next i
            End Using
            sourceDb.CloseInput(True)
        Catch ex As SystemException
        End Try
    End Function
    '===============================================================================================================================================================
    'функция копирует тип линии из другого чертежа
    Public Shared Function FuncCreateTypeLineDestDraw(ByVal PutshDraw As String, ByVal NameTypeLine As String, Optional boolReplase As Boolean = False) As Boolean
        FuncCreateTypeLineDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim acLineTypTbl As LinetypeTable = sourceTr.GetObject(sourceDb.LinetypeTableId, OpenMode.ForRead)
            If acLineTypTbl.Has(NameTypeLine) = True Then
                Dim IDTypeLine As ObjectId = acLineTypTbl.Item(NameTypeLine)
                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                ObjColID.Add(IDTypeLine)
                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim db1 As Database = New Database(False, False)
                Dim destDb1 As Database = acDoc.Database
                Dim destDictId1 As ObjectId = destDb1.LinetypeTableId
                Dim iMap1 As IdMapping = New IdMapping()
                If boolReplase = True Then
                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                Else
                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                End If
            Else
                Return False
            End If
            sourceTr.Commit()
        End Using
        sourceDb.CloseInput(True)
        FuncCreateTypeLineDestDraw = True
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает код типа линии из внешнего файла (supportPatchFolder - массив путей к вспомогательным фвйлам поддержки)
    Public Shared Function FuncReturnTypeLineCodeDestDraw(ByVal nameTypeLine As String, ByVal PutshDraw As String, Optional supportPatchFolder As String() = Nothing) As String
        FuncReturnTypeLineCodeDestDraw = "A"
        Dim str As String = "A"
        Dim sourceDb As Database = New Database(False, True)
        sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim acBlkTbl As BlockTable = sourceTr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = sourceTr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim acLineTypTbl As LinetypeTable = sourceTr.GetObject(sourceDb.LinetypeTableId, OpenMode.ForRead)
            If acLineTypTbl.Has(nameTypeLine) = True Then
                Dim IDTypeLine As ObjectId = acLineTypTbl.Item(nameTypeLine)
                Dim recTypeLine As LinetypeTableRecord = sourceTr.GetObject(IDTypeLine, OpenMode.ForRead)
                'находим все штрихи и точки
                Dim lenD As Integer = recTypeLine.NumDashes
                If lenD > 0 Then
                    Try
                        For i As Integer = 0 To lenD - 1
                            Dim l As Double = 0
                            Dim indexShape As Integer = 0
                            Dim textVal As String = ""
                            Dim nameShape As String = ""
                            'находим значения
                            l = recTypeLine.DashLengthAt(i) 'длина штриха
                            indexShape = recTypeLine.ShapeNumberAt(i) 'индекс формы (0 - нет формы)
                            Try
                                textVal = recTypeLine.TextAt(i) 'читаем текст
                            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                            End Try
                            If indexShape > 0 Then
                                Dim idTextStyleShape As ObjectId = recTypeLine.ShapeStyleAt(i)
                                Dim textStyleShape As TextStyleTableRecord = sourceTr.GetObject(idTextStyleShape, OpenMode.ForRead)
                                If textStyleShape.IsShapeFile = True Then
                                    Using acTrans As Transaction = sourceDb.TransactionManager.StartTransaction()
                                        Dim Btr As BlockTableRecord = acTrans.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForWrite, False)
                                        Dim pShape As Autodesk.AutoCAD.DatabaseServices.Shape = New Autodesk.AutoCAD.DatabaseServices.Shape(Point3d.Origin, 1, 0, 1)
                                        pShape.ShapeNumber = indexShape
                                        pShape.StyleId = idTextStyleShape
                                        Btr.AppendEntity(pShape)
                                        nameShape = pShape.Name
                                    End Using
                                    'имя формы не найдено, ищем имя формы в файле shp
                                    If nameShape.Trim.Length = 0 Then
                                        If IsArray(supportPatchFolder) = False Then
                                            Dim boolFindFolderPatch As Boolean = FuncFiles.FuncGetSupportFolderPath(supportPatchFolder)
                                        End If
                                        If IsArray(supportPatchFolder) = True Then
                                            Dim nameFileSHP As String = textStyleShape.FileName.Replace(".shx", ".shp")
                                            Dim boolFindShp As Boolean = False
                                            Dim regMatch As String = "\*"
                                            For kd As Integer = 0 To supportPatchFolder.Length - 1
                                                Dim searFiles As String() = Directory.GetFiles(supportPatchFolder(kd), nameFileSHP)
                                                If IsArray(searFiles) = True Then
                                                    If searFiles.Length > 0 Then
                                                        'читаем файл с формой
                                                        Dim ffileShp As String = searFiles(0)
                                                        If IO.File.Exists(ffileShp) = True Then
                                                            Dim sr As StreamReader = New StreamReader(ffileShp)
                                                            While sr.Peek <> -1
                                                                Dim st As String = sr.ReadLine().Trim
                                                                Dim flagread As String = Left(st, 1)
                                                                If Regex.IsMatch(flagread, regMatch) Then 'chr(42)
                                                                    Dim rez As String() = st.Split(",")
                                                                    If rez.GetUpperBound(0) > 1 Then
                                                                        Dim indForm As String = Mid(rez(0), 2)
                                                                        If indForm Like indexShape Then
                                                                            nameShape = rez(2)
                                                                            boolFindShp = True
                                                                            Exit While
                                                                        End If
                                                                    End If
                                                                End If
                                                            End While
                                                            sr.Close()
                                                            If boolFindShp = True Then
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next kd
                                        End If
                                    End If
                                End If
                                str = str & l & ",[" & nameShape & "," & textStyleShape.FileName & ",S=" & recTypeLine.ShapeScaleAt(i) & ",R=" & recTypeLine.ShapeRotationAt(i) & ",X=" & recTypeLine.ShapeOffsetAt(i).X & ",Y=" & recTypeLine.ShapeOffsetAt(i).Y & "]"
                            ElseIf textVal.Trim.Length > 0 Then
                                Dim id As ObjectId = recTypeLine.ShapeStyleAt(i)
                                Dim acShape As TextStyleTableRecord = sourceTr.GetObject(id, OpenMode.ForRead)
                                Dim styleName As String = acShape.Name
                                str = str & l & ",[" & Chr(34) & textVal & Chr(34) & "," & styleName & ",S=" & recTypeLine.ShapeScaleAt(i) & ",R=" & recTypeLine.ShapeRotationAt(i) & ",X=" & recTypeLine.ShapeOffsetAt(i).X & ",Y=" & recTypeLine.ShapeOffsetAt(i).Y & "]"
                            Else
                                str = str & l
                            End If
                            If i < lenD - 1 Then
                                str = str & ","
                            End If
                        Next i
                        FuncReturnTypeLineCodeDestDraw = str
                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                    End Try
                End If
            End If
        End Using
        sourceDb.CloseInput(True)
        Return FuncReturnTypeLineCodeDestDraw
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция копирует тип мультилинии из другого чертежа
    Public Shared Function FuncCreateTypeMLineDestDraw(ByVal PutshDraw As String, ByVal NameTypeMLine As String, Optional boolReplase As Boolean = False) As Boolean
        FuncCreateTypeMLineDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim mlDict As DBDictionary = sourceTr.GetObject(sourceDb.MLStyleDictionaryId, OpenMode.ForRead)
            If mlDict.Contains(NameTypeMLine) = True Then
                Dim IDTypeMLine As ObjectId = mlDict.Item(NameTypeMLine)
                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                ObjColID.Add(IDTypeMLine)
                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim db1 As Database = New Database(False, False)
                Dim destDb1 As Database = acDoc.Database
                Dim destDictId1 As ObjectId = destDb1.MLStyleDictionaryId
                Dim iMap1 As IdMapping = New IdMapping()
                If boolReplase = True Then
                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                Else
                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                End If
            Else
                Return False
            End If
            sourceTr.Commit()
        End Using
        sourceDb.CloseInput(True)
        FuncCreateTypeMLineDestDraw = True
    End Function
    '===========================================================================================================
    'функция копирует стиль текста из другого чертежа
    Public Shared Function FuncCreateTextStyleDestDraw(ByVal PutshDraw As String, ByVal NameTextStyle As String, Optional boolReplase As Boolean = False) As Boolean
        FuncCreateTextStyleDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim acTextStyleTbl As TextStyleTable = sourceTr.GetObject(sourceDb.TextStyleTableId, OpenMode.ForRead)
            If acTextStyleTbl.Has(NameTextStyle) = True Then
                Dim IDTextStyle As ObjectId = acTextStyleTbl.Item(NameTextStyle)
                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                ObjColID.Add(IDTextStyle)
                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim db1 As Database = New Database(False, False)
                Dim destDb1 As Database = acDoc.Database
                Dim destDictId1 As ObjectId = destDb1.TextStyleTableId
                Dim iMap1 As IdMapping = New IdMapping()
                If boolReplase = True Then
                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                Else
                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                End If
            Else
                Return False
            End If
            sourceTr.Commit()
        End Using
        sourceDb.CloseInput(True)
        FuncCreateTextStyleDestDraw = True
    End Function
    '===========================================================================================================
    'функция копирует стиль мультивыносок из другого чертежа
    Public Shared Function FuncCreateMLeaderStyleDestDraw(ByVal PutshDraw As String, ByVal NameMLeaderStyle As String, Optional boolReplase As Boolean = False) As Boolean
        FuncCreateMLeaderStyleDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim acMleaderStyleTbl As DBDictionary = sourceTr.GetObject(sourceDb.MLeaderStyleDictionaryId, OpenMode.ForRead)
            Dim MLeaderStyleCollection As ObjectIdCollection = New ObjectIdCollection
            For Each id As DBDictionaryEntry In acMleaderStyleTbl
                If id.Key Like NameMLeaderStyle Then
                    MLeaderStyleCollection.Add(id.Value)
                    Exit For
                End If
            Next
            If MLeaderStyleCollection.Count > 0 Then
                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim db1 As Database = New Database(False, False)
                Dim destDb1 As Database = acDoc.Database
                Dim destDictId1 As ObjectId = destDb1.MLeaderStyleDictionaryId
                Dim iMap1 As IdMapping = New IdMapping()
                If boolReplase = True Then
                    destDb1.WblockCloneObjects(MLeaderStyleCollection, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                Else
                    destDb1.WblockCloneObjects(MLeaderStyleCollection, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                End If
            Else
                Return False
            End If
            sourceTr.Commit()
        End Using
        sourceDb.CloseInput(True)
        FuncCreateMLeaderStyleDestDraw = True
    End Function
    '===========================================================================================================
    'функция копирует размерный стиль из другого чертежа
    Public Shared Function FuncCreateDimStyleDestDraw(ByVal PutshDraw As String, ByVal NameDimStyle As String, Optional boolReplase As Boolean = False) As Boolean
        FuncCreateDimStyleDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim acDimStyleTbl As DimStyleTable = sourceTr.GetObject(sourceDb.DimStyleTableId, OpenMode.ForRead)
            If acDimStyleTbl.Has(NameDimStyle) = True Then
                Dim IDDimStyle As ObjectId = acDimStyleTbl.Item(NameDimStyle)
                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                ObjColID.Add(IDDimStyle)
                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim db1 As Database = New Database(False, False)
                Dim destDb1 As Database = acDoc.Database
                Dim destDictId1 As ObjectId = destDb1.DimStyleTableId
                Dim iMap1 As IdMapping = New IdMapping()
                If boolReplase = True Then
                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                Else
                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                End If
            Else
                Return False
            End If
            sourceTr.Commit()
        End Using
        sourceDb.CloseInput(True)
        FuncCreateDimStyleDestDraw = True
    End Function
    '===========================================================================================================
    'функция копирует стиль таблицы из другого чертежа
    Public Shared Function FuncCreateTableStyleDestDraw(ByVal PutshDraw As String, ByVal NameTableStyle As String, Optional boolReplase As Boolean = False) As Boolean
        FuncCreateTableStyleDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim acTableStyleTbl As DBDictionary = sourceTr.GetObject(sourceDb.TableStyleDictionaryId, OpenMode.ForRead)
            Dim TableStyleCollection As ObjectIdCollection = New ObjectIdCollection
            For Each id As DBDictionaryEntry In acTableStyleTbl
                If id.Key Like NameTableStyle Then
                    TableStyleCollection.Add(id.Value)
                    Exit For
                End If
            Next
            If TableStyleCollection.Count > 0 Then
                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim db1 As Database = New Database(False, False)
                Dim destDb1 As Database = acDoc.Database
                Dim destDictId1 As ObjectId = destDb1.TableStyleDictionaryId
                Dim iMap1 As IdMapping = New IdMapping()
                If boolReplase = True Then
                    destDb1.WblockCloneObjects(TableStyleCollection, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                Else
                    destDb1.WblockCloneObjects(TableStyleCollection, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                End If
            End If
            sourceTr.Commit()
        End Using
        sourceDb.CloseInput(True)
        FuncCreateTableStyleDestDraw = True
    End Function
    '========================================================================================================================
    'функция копирует ключ-описатель для COGO точки из другого чертежа
    Public Shared Function FuncCopyDeskKeyDestDraw(ByVal PutshDraw As String, ByVal RawDesk As String, Optional boolReplase As Boolean = False) As Boolean
        FuncCopyDeskKeyDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim pdKeySetIds As ObjectIdCollection = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(sourceDb).SearchOrder
            If pdKeySetIds.Count > 0 Then
                For Each idKeySet As ObjectId In pdKeySetIds
                    Dim DeskKeySet As PointDescriptionKeySet = sourceTr.GetObject(idKeySet, OpenMode.ForRead)
                    Dim pdKeyIds As ObjectIdCollection = DeskKeySet.GetPointDescriptionKeyIds
                    For Each idDeskKey As ObjectId In pdKeyIds
                        If idDeskKey.IsNull = False Then
                            Dim DeskKey As PointDescriptionKey = sourceTr.GetObject(idDeskKey, OpenMode.ForRead)
                            If DeskKey.Code Like RawDesk Then
                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
                                Dim db1 As Database = New Database(False, False)
                                Dim destDb1 As Database = acDoc.Database
                                Dim destDictId1 As ObjectId = destDb1.LayerTableId
                                Dim iMap1 As IdMapping = New IdMapping()
                                'копируем слой
                                Dim IDLayer As ObjectId = DeskKey.LayerId
                                Dim acLyrTblRec As LayerTableRecord = sourceTr.GetObject(IDLayer, OpenMode.ForRead)
                                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                ObjColID.Add(DeskKey.LayerId)
                                destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                'копируем стиль метки
                                Dim LabelStylePoint As LabelStyle = sourceTr.GetObject(DeskKey.LabelStyleId, OpenMode.ForRead, False)
                                Try
                                    If boolReplase = False Then
                                        LabelStylePoint.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Ignore)
                                    Else
                                        LabelStylePoint.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Override)
                                    End If
                                Catch ex As System.InvalidOperationException
                                End Try
                                'копируем стиль точки
                                Dim StylePoint As PointStyle = sourceTr.GetObject(DeskKey.StyleId, OpenMode.ForRead, False)
                                Try
                                    If boolReplase = False Then
                                        StylePoint.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Ignore)
                                    Else
                                        StylePoint.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Override)
                                    End If
                                Catch ex As System.InvalidOperationException
                                End Try
                                'создаем новую ветку с категорией ключей описателей
                                Dim Tr As Transaction = destDb1.TransactionManager.StartTransaction()
                                Using Tr
                                    'создаем новую категорию ключей описателей
                                    Dim NewIdDeskSet As ObjectId = Nothing
                                    If PointDescriptionKeySetCollection.GetPointDescriptionKeySets(destDb1).Contains(DeskKeySet.Name) = False Then
                                        NewIdDeskSet = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(destDb1).Add(DeskKeySet.Name)
                                    Else
                                        NewIdDeskSet = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(destDb1).Item(DeskKeySet.Name)
                                    End If
                                    Dim DeskKeySetDest As PointDescriptionKeySet = Tr.GetObject(NewIdDeskSet, OpenMode.ForWrite)
                                    'в данной категории создаем ключ-описатель
                                    Try
                                        Dim idKey As ObjectId = DeskKeySetDest.Add(RawDesk)
                                        Dim DeskKeyDest As PointDescriptionKey = Tr.GetObject(idKey, OpenMode.ForWrite)
                                        'добавляем созданный слой
                                        Dim acLyrTbl As LayerTable = Tr.GetObject(destDb1.LayerTableId, OpenMode.ForRead)
                                        If acLyrTbl.Has(acLyrTblRec.Name) = True Then
                                            DeskKeyDest.LayerId = acLyrTbl.Item(acLyrTblRec.Name)
                                            DeskKeyDest.ApplyLayerId = True
                                        End If
                                        'добавляем стиль точки
                                        If CivilDoc.Styles.PointStyles.Contains(StylePoint.Name) = True Then
                                            DeskKeyDest.StyleId = CivilDoc.Styles.PointStyles.Item(StylePoint.Name)
                                            DeskKeyDest.ApplyLabelStyleId = True
                                        End If
                                        'добавляем стиль меток
                                        If CivilDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles.Contains(LabelStylePoint.Name) = True Then
                                            DeskKeyDest.LabelStyleId = CivilDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles.Item(LabelStylePoint.Name)
                                            DeskKeyDest.ApplyStyleId = True
                                        End If
                                        DeskKeyDest.Format = DeskKey.Format
                                    Catch ex As System.ArgumentException
                                        Return False
                                    End Try
                                    Tr.Commit()
                                End Using
                                FuncCopyDeskKeyDestDraw = True
                                Exit For
                            End If
                        End If
                    Next
                Next
            Else
                Return False
            End If
        End Using
        sourceDb.CloseInput(True)
    End Function
    '======================================================================================================================================
    'функция копирует группу точек из другого чертежа
    Public Shared Function FuncCopyGroupPointsDestDraw(ByVal PutshDraw As String, ByVal StrPointGroup As String, Optional ByVal boolReplace As Boolean = False) As Boolean
        FuncCopyGroupPointsDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilDocument.GetCivilDocument(sourceDb)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Try
                Dim IDPointGroup As ObjectId = CivilDoc.PointGroups.Item(StrPointGroup)
                Dim GroupPointUs As PointGroup = sourceTr.GetObject(IDPointGroup, OpenMode.ForRead, False)
                Dim PQ As StandardPointGroupQuery = GroupPointUs.GetQuery
                Dim UDPClass As String = GroupPointUs.UDPClassificationName
                'открываем текущий чертеж
                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim destCivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
                Dim db1 As Database = New Database(False, False)
                Dim destDb1 As Database = acDoc.Database
                Dim Tr As Transaction = destDb1.TransactionManager.StartTransaction()
                Using Tr
                    If destCivilDoc.PointGroups.Contains(GroupPointUs.Name) = False Then
                        Dim idNewGroup As ObjectId = destCivilDoc.PointGroups.Add(GroupPointUs.Name)
                        Dim NewGroupPoint As PointGroup = Tr.GetObject(idNewGroup, OpenMode.ForWrite)
                        NewGroupPoint.Description = GroupPointUs.Description
                        'копируем стиль метки
                        Dim LabelStylePoint As LabelStyle = sourceTr.GetObject(GroupPointUs.PointLabelStyleId, OpenMode.ForRead, False)
                        Try
                            If boolReplace = False Then
                                LabelStylePoint.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Ignore)
                            Else
                                LabelStylePoint.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Override)
                            End If
                        Catch ex As System.InvalidOperationException
                        End Try
                        'копируем стиль точки
                        Dim StylePoint As PointStyle = sourceTr.GetObject(GroupPointUs.PointStyleId, OpenMode.ForRead, False)
                        Try
                            If boolReplace = False Then
                                StylePoint.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Ignore)
                            Else
                                StylePoint.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Override)
                            End If
                        Catch ex As System.InvalidOperationException
                        End Try
                        'добавляем стиль точки
                        If destCivilDoc.Styles.PointStyles.Contains(StylePoint.Name) = True Then
                            NewGroupPoint.PointStyleId = destCivilDoc.Styles.PointStyles.Item(StylePoint.Name)
                        End If
                        'добавляем стиль меток
                        If destCivilDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles.Contains(LabelStylePoint.Name) = True Then
                            NewGroupPoint.PointLabelStyleId = destCivilDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles.Item(LabelStylePoint.Name)
                        End If
                        'добавляем настройки
                        Dim NewQuery As StandardPointGroupQuery = New StandardPointGroupQuery()
                        NewQuery.ExcludeElevations = PQ.ExcludeElevations
                        NewQuery.ExcludeFullDescriptions = PQ.ExcludeFullDescriptions
                        NewQuery.ExcludeNames = PQ.ExcludeNames
                        NewQuery.ExcludeNumbers = PQ.ExcludeNumbers
                        NewQuery.ExcludeRawDescriptions = PQ.ExcludeRawDescriptions
                        NewQuery.IncludeAllPoints = PQ.IncludeAllPoints
                        NewQuery.IncludeElevations = PQ.IncludeElevations
                        NewQuery.IncludeFullDescriptions = PQ.IncludeFullDescriptions
                        NewQuery.IncludeNames = PQ.IncludeNames
                        NewQuery.IncludeNumbers = PQ.IncludeNumbers
                        NewQuery.IncludeRawDescriptions = PQ.IncludeRawDescriptions
                        NewGroupPoint.SetQuery(NewQuery)
                        'добавляем пользовательскую классификацию
                        Dim UdpClassCol As UDPClassificationCollection = destCivilDoc.PointUDPClassifications
                        Dim boolClass As Boolean = False
                        If UdpClassCol.Count > 0 Then
                            For i As Integer = 0 To UdpClassCol.Count - 1
                                Dim nameUDP As UDPClassification = UdpClassCol.Item(i)
                                If nameUDP.Name Like UDPClass Then
                                    NewGroupPoint.UseCustomClassification(nameUDP.Name)
                                    boolClass = True
                                    Exit For
                                End If
                            Next i
                        End If
                        If boolClass = False Then
                            Dim udpClassification As UDPClassification = destCivilDoc.PointUDPClassifications.Add(UDPClass)
                            For Each IdUdp As UDP In CivilDoc.PointUDPs
                                If IsNothing(IdUdp) = False Then
                                    Try
                                        Dim nameClass As String = IdUdp.ClassificationName
                                        If nameClass Like UDPClass Then
                                            If TypeOf IdUdp Is UDPDouble Then
                                                Dim udpDbl As UDPDouble = IdUdp
                                                Dim attributeTypeInfoDB As AttributeTypeInfoDouble = New AttributeTypeInfoDouble(IdUdp.Name)
                                                attributeTypeInfoDB.Description = IdUdp.Description
                                                attributeTypeInfoDB.DefaultValue = IdUdp.DefaultValue
                                                attributeTypeInfoDB.UpperBoundValue = udpDbl.UpperBoundValue
                                                attributeTypeInfoDB.LowerBoundValue = udpDbl.LowerBoundValue
                                                udpClassification.CreateUDP(attributeTypeInfoDB)
                                            ElseIf TypeOf IdUdp Is UDPInteger Then
                                                Dim udpInt As UDPInteger = IdUdp
                                                Dim attributeTypeInfoDB As AttributeTypeInfoInt = New AttributeTypeInfoInt(IdUdp.Name)
                                                attributeTypeInfoDB.Description = IdUdp.Description
                                                attributeTypeInfoDB.DefaultValue = IdUdp.DefaultValue
                                                attributeTypeInfoDB.UpperBoundValue = udpInt.UpperBoundValue
                                                attributeTypeInfoDB.LowerBoundValue = udpInt.LowerBoundValue
                                                udpClassification.CreateUDP(attributeTypeInfoDB)
                                            ElseIf TypeOf IdUdp Is UDPString Then
                                                Dim udpStr As UDPString = IdUdp
                                                Dim attributeTypeInfoDB As AttributeTypeInfoString = New AttributeTypeInfoString(IdUdp.Name)
                                                attributeTypeInfoDB.Description = IdUdp.Description
                                                attributeTypeInfoDB.DefaultValue = IdUdp.DefaultValue
                                                udpClassification.CreateUDP(attributeTypeInfoDB)
                                            End If
                                        End If
                                    Catch ex As System.NullReferenceException
                                    End Try
                                End If
                            Next
                            NewGroupPoint.UseCustomClassification(udpClassification)
                        End If
                        NewGroupPoint.Update()
                    Else
                        Return False
                    End If
                    Tr.Commit()
                End Using
            Catch ex As System.ArgumentException
            End Try
        End Using
        sourceDb.CloseInput(True)
        FuncCopyGroupPointsDestDraw = True
    End Function
    '=================================================================================================================================
    'функция копирует PropertySET из другого чертежа (1-Полный путь к файлу-шаблону .dwg, 2-имя таблицы)
    Public Shared Function FuncCopyTablePSDestDraw(ByVal PathDraw As String, ByVal NamePSTable As String) As Boolean
        FuncCopyTablePSDestDraw = False
        Try
            Dim sourceDb As Database = New Database(False, True)
            sourceDb.ReadDwgFile(PathDraw, FileShare.Read, True, String.Empty)
            Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
            Using sourceTr
                Dim dictPropSetDef = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(sourceDb)
                If dictPropSetDef.Has(NamePSTable, sourceTr) Then
                    Dim IDPSTable As ObjectId = dictPropSetDef.GetAt(NamePSTable)
                    Dim CollDictPropSet As ObjectIdCollection = New ObjectIdCollection
                    CollDictPropSet.Add(IDPSTable)
                    Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim destDb As Database = acDoc.Database
                    Dim destTr As Transaction = destDb.TransactionManager.StartTransaction()
                    Using destTr
                        Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(destDb)
                        If dictPropSetDefTemp.Has(NamePSTable, destTr) = False Then
                            Dim helpme As Autodesk.AEC.ApplicationServices.Utility.CloningHelper = New Autodesk.AEC.ApplicationServices.Utility.CloningHelper
                            Dim RXClassPS As RXClass = RXClass.GetClass(GetType(Autodesk.AEC.PropertyData.DatabaseServices.PropertySet))
                            helpme.Clone(sourceDb, destDb, CollDictPropSet, RXClassPS, True)
                            FuncCopyTablePSDestDraw = True
                        End If
                        destTr.Commit()
                    End Using
                End If
            End Using
            sourceDb.CloseInput(True)
        Catch ex As SystemException
        End Try
    End Function
    '=================================================================================================================================
    'функция копирует стиль Характерной линии из другого чертежа
    Public Shared Function FuncCopyFeatureLineStyleDestDraw(ByVal PutshDraw As String, ByVal NameStyle As String, Optional ByVal boolReplace As Boolean = False) As Boolean
        FuncCopyFeatureLineStyleDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilDocument.GetCivilDocument(sourceDb)
            Dim idStyle As ObjectId = Nothing
            Try
                idStyle = CivilDoc.Styles.FeatureLineStyles.Item(NameStyle)
            Catch ex As System.ArgumentException
                Return False
            End Try
            If idStyle.IsNull = False Then
                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim ActiveCivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
                Dim db1 As Database = New Database(False, False)
                Dim destDb1 As Database = acDoc.Database
                Dim iMap1 As IdMapping = New IdMapping()
                'копируем стиль характерной линии
                If idStyle.IsNull = False Then
                    Dim FeatureLineStyle As FeatureLineStyle = sourceTr.GetObject(idStyle, OpenMode.ForRead, False)
                    Try
                        If boolReplace = False Then
                            FeatureLineStyle.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Ignore)
                        Else
                            FeatureLineStyle.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Override)
                        End If
                    Catch ex As System.InvalidOperationException
                        Return False
                    End Try
                End If
            End If
            sourceTr.Commit()
            FuncCopyFeatureLineStyleDestDraw = True
        End Using
        sourceDb.CloseInput(True)
    End Function
    '=================================================================================================================================
    'функция копирует стиль метки civi3d для линии
    Public Shared Function FuncCopyLabelLineStyleDestDraw(ByVal PutshDraw As String, ByVal NameStyle As String, Optional ByVal boolReplace As Boolean = False) As Boolean
        FuncCopyLabelLineStyleDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilDocument.GetCivilDocument(sourceDb)
            Dim idStyle As ObjectId = Nothing
            Try
                idStyle = CivilDoc.Styles.LabelStyles.GeneralLineLabelStyles.Item(NameStyle)
            Catch ex As System.ArgumentException
                Return False
            End Try
            If idStyle.IsNull = False Then
                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim ActiveCivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
                Dim db1 As Database = New Database(False, False)
                Dim destDb1 As Database = acDoc.Database
                Dim iMap1 As IdMapping = New IdMapping()
                'копируем стиль характерной линии
                If idStyle.IsNull = False Then
                    Dim LabelLineStyle As LabelStyle = sourceTr.GetObject(idStyle, OpenMode.ForRead, False)
                    Try
                        If boolReplace = False Then
                            LabelLineStyle.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Ignore)
                        Else
                            LabelLineStyle.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Override)
                        End If
                    Catch ex As System.InvalidOperationException
                        Return False
                    End Try
                End If
            End If
            sourceTr.Commit()
            FuncCopyLabelLineStyleDestDraw = True
        End Using
        sourceDb.CloseInput(True)
    End Function
    '=================================================================================================================================
    'функция копирует стиль метки civi3d для примечаний
    Public Shared Function FuncCopyNoteLabelStyleDestDraw(ByVal PutshDraw As String, ByVal NameStyle As String, Optional ByVal boolReplace As Boolean = False) As Boolean
        FuncCopyNoteLabelStyleDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try
        Try
            Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
            Using sourceTr
                Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilDocument.GetCivilDocument(sourceDb)
                Dim idStyle As ObjectId = Nothing
                Try
                    idStyle = CivilDoc.Styles.LabelStyles.GeneralNoteLabelStyles.Item(NameStyle)
                Catch ex As System.ArgumentException
                    Return False
                End Try
                If idStyle.IsNull = False Then
                    Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim ActiveCivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
                    Dim db1 As Database = New Database(False, False)
                    Dim destDb1 As Database = acDoc.Database
                    Dim iMap1 As IdMapping = New IdMapping()
                    'копируем стиль характерной линии
                    If idStyle.IsNull = False Then
                        Dim LabelNoteStyle As LabelStyle = sourceTr.GetObject(idStyle, OpenMode.ForRead, False)
                        Try
                            If boolReplace = False Then
                                LabelNoteStyle.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Ignore)
                            Else
                                LabelNoteStyle.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Override)
                            End If
                        Catch ex As System.InvalidOperationException
                            Return False
                        End Try
                    End If
                End If
                sourceTr.Commit()
                FuncCopyNoteLabelStyleDestDraw = True
            End Using
        Catch ex As System.Exception
        End Try
        sourceDb.CloseInput(True)
    End Function
    '=================================================================================================================================
    'функция копирует стиль маркера civi3d для примечаний
    Public Shared Function FuncCopyMarkerStyleDestDraw(ByVal PutshDraw As String, ByVal NameStyle As String, Optional ByVal boolReplace As Boolean = False) As Boolean
        FuncCopyMarkerStyleDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try

        Try
            Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
            Using sourceTr
                Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilDocument.GetCivilDocument(sourceDb)
                Dim idStyle As ObjectId = Nothing
                Try
                    idStyle = CivilDoc.Styles.MarkerStyles.Item(NameStyle)
                Catch ex As System.ArgumentException
                    Return False
                End Try
                If idStyle.IsNull = False Then
                    Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim ActiveCivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
                    Dim db1 As Database = New Database(False, False)
                    Dim destDb1 As Database = acDoc.Database
                    Dim iMap1 As IdMapping = New IdMapping()
                    'копируем стиль характерной линии
                    If idStyle.IsNull = False Then
                        Dim MarkerUserStyle As MarkerStyle = sourceTr.GetObject(idStyle, OpenMode.ForRead, False)
                        Try
                            If boolReplace = False Then
                                MarkerUserStyle.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Ignore)
                            Else
                                MarkerUserStyle.ExportTo(destDb1, Autodesk.Civil.StyleConflictResolverType.Override)
                            End If
                        Catch ex As System.InvalidOperationException
                            Return False
                        End Try
                    End If
                End If
                sourceTr.Commit()
                FuncCopyMarkerStyleDestDraw = True
            End Using
        Catch ex As System.Exception
        End Try
        sourceDb.CloseInput(True)
    End Function
    '=================================================================================================================
    'функция копирует определение блока из другого чертежа
    Public Shared Function FuncCopyBlockDestDraw(ByVal PatshTemplate As String, ByVal NameBlock As String, Optional boolReplace As Boolean = False) As Boolean
        FuncCopyBlockDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PatshTemplate, FileShare.Read, True, String.Empty)
            Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
            Using sourceTr
                Dim acBlockTbl As BlockTable = sourceTr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
                If acBlockTbl.Has(NameBlock) = True Then
                    Dim IDBlk As ObjectId = acBlockTbl.Item(NameBlock)
                    Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                    ObjColID.Add(IDBlk)
                    Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim destDb1 As Database = acDoc.Database
                    Dim destDictId1 As ObjectId = destDb1.BlockTableId
                    Dim iMap1 As IdMapping = New IdMapping()
                    If boolReplace = False Then
                        destDb1.WblockCloneObjects(ObjColID, destDb1.BlockTableId, iMap1, DuplicateRecordCloning.Ignore, False)
                    Else
                        destDb1.WblockCloneObjects(ObjColID, destDb1.BlockTableId, iMap1, DuplicateRecordCloning.Replace, False)
                    End If
                    FuncCopyBlockDestDraw = True
                Else
                    Return False
                End If
                sourceTr.Commit()
                FuncCopyBlockDestDraw = True
            End Using
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try
        sourceDb.CloseInput(True)
    End Function
    '=================================================================================================================
    'функция копирует определение мультивидовых блоков из другого чертежа
    Public Shared Function FuncCopyMVBlockDestDraw(ByVal PatshTemplate As String, ByVal NameMVBlock As String, Optional boolReplace As Boolean = False) As Boolean
        FuncCopyMVBlockDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PatshTemplate, FileShare.Read, True, String.Empty)
            Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
            Using sourceTr
                Dim acBlockTbl As BlockTable = sourceTr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
                Dim mvbstyledict As Autodesk.AEC.DatabaseServices.DictionaryMultiViewBlockDefinition = New Autodesk.AEC.DatabaseServices.DictionaryMultiViewBlockDefinition(sourceDb)
                If mvbstyledict.Has(NameMVBlock, sourceTr) = True Then
                    Dim bstyleID As ObjectId = mvbstyledict.GetAt(NameMVBlock)
                    Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                    ObjColID.Add(bstyleID)
                    Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim destDb As Database = acDoc.Database
                    Dim destmvbstyledict As Autodesk.AEC.DatabaseServices.DictionaryMultiViewBlockDefinition = New Autodesk.AEC.DatabaseServices.DictionaryMultiViewBlockDefinition(destDb)
                    Dim iMap1 As IdMapping = New IdMapping()
                    If boolReplace = False Then
                        destDb.WblockCloneObjects(ObjColID, destmvbstyledict.DictionaryId, iMap1, DuplicateRecordCloning.Ignore, False)
                    Else
                        destDb.WblockCloneObjects(ObjColID, destmvbstyledict.DictionaryId, iMap1, DuplicateRecordCloning.Replace, False)
                    End If
                    FuncCopyMVBlockDestDraw = True
                Else
                    Return False
                End If
                sourceTr.Commit()
                FuncCopyMVBlockDestDraw = True
            End Using
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try
        sourceDb.CloseInput(True)
    End Function
    '=============================================================================================================
    'функция копирует слой из другого чертежа
    Public Shared Function FuncCreateLayerDestDraw(ByVal PutshDraw As String, ByVal NameLayer As String, Optional boolReplace As Boolean = False) As Boolean
        FuncCreateLayerDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim acLyrTbl As LayerTable = sourceTr.GetObject(sourceDb.LayerTableId, OpenMode.ForRead)
            If acLyrTbl.Has(NameLayer) = True Then
                Dim IDLayer As ObjectId = acLyrTbl.Item(NameLayer)
                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                ObjColID.Add(IDLayer)
                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim destDb1 As Database = acDoc.Database
                Dim destDictId1 As ObjectId = destDb1.LayerTableId
                Dim iMap1 As IdMapping = New IdMapping()
                If boolReplace = False Then
                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                Else
                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                End If
            Else
                Return False
            End If
            sourceTr.Commit()
        End Using
        sourceDb.CloseInput(True)
        FuncCreateLayerDestDraw = True
    End Function
    '=========================================================================================================================================================
    'функция копирует линейные объекты из файла-шаблона по заданному массиву
    Public Shared Function FuncCopyLineStylesDestDraw(ByVal PutshDraw As String, ByVal ArrayNameObject As String(,), ByVal boolReplase As Boolean) As Boolean
        FuncCopyLineStylesDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return 0
        End Try
        If IsArray(ArrayNameObject) = True Then
            If ArrayNameObject.GetUpperBound(1) > 0 Then
                Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
                Using sourceTr
                    Dim acLineTypTbl As LinetypeTable = sourceTr.GetObject(sourceDb.LinetypeTableId, OpenMode.ForRead)
                    Dim acLyrTbl As LayerTable = sourceTr.GetObject(sourceDb.LayerTableId, OpenMode.ForRead)
                    Dim dictPropSetDef = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(sourceDb)
                    For i As Integer = 0 To ArrayNameObject.GetUpperBound(1)
                        Try
                            'копируем тип линии
                            Dim strFields As String = ArrayNameObject(0, i).Trim
                            Dim ValStr As String = ArrayNameObject(1, i).Trim
                            If strFields.Trim Like "Name" Then
                                If IsNothing(ValStr) = False Then
                                    If ValStr.Trim.Length > 0 Then
                                        If acLineTypTbl.Has(ValStr) = True Then
                                            Dim IDTypeLine As ObjectId = acLineTypTbl.Item(ValStr)
                                            Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                            ObjColID.Add(IDTypeLine)
                                            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                            Dim db1 As Database = New Database(False, False)
                                            Dim destDb1 As Database = acDoc.Database
                                            Dim destDictId1 As ObjectId = destDb1.LinetypeTableId
                                            Dim iMap1 As IdMapping = New IdMapping()
                                            If boolReplase = True Then
                                                destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                            Else
                                                destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf strFields.Trim Like "Layer" Then
                                'копируем слой
                                If IsNothing(ValStr) = False Then
                                    If ValStr.Trim.Length > 0 Then
                                        If acLyrTbl.Has(ValStr) = True Then
                                            Dim IDLayer As ObjectId = acLyrTbl.Item(ValStr)
                                            Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                            ObjColID.Add(IDLayer)
                                            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                            Dim destDb1 As Database = acDoc.Database
                                            Dim destDictId1 As ObjectId = destDb1.LayerTableId
                                            Dim iMap1 As IdMapping = New IdMapping()
                                            If boolReplase = False Then
                                                destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                            Else
                                                destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf strFields.Trim Like "PS" Then
                                'копируем таблицу PS
                                If IsNothing(ValStr) = False Then
                                    If ValStr.Trim.Length > 0 Then
                                        If dictPropSetDef.Has(ValStr, sourceTr) = True Then
                                            Dim IDPSTable As ObjectId = dictPropSetDef.GetAt(ValStr)
                                            Dim CollDictPropSet As ObjectIdCollection = New ObjectIdCollection
                                            CollDictPropSet.Add(IDPSTable)
                                            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                            Dim destDb As Database = acDoc.Database
                                            Dim destTr As Transaction = destDb.TransactionManager.StartTransaction()
                                            Using destTr
                                                Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(destDb)
                                                If dictPropSetDefTemp.Has(ValStr, destTr) = False Then
                                                    Dim helpme As Autodesk.AEC.ApplicationServices.Utility.CloningHelper = New Autodesk.AEC.ApplicationServices.Utility.CloningHelper
                                                    Dim RXClassPS As RXClass = RXClass.GetClass(GetType(Autodesk.AEC.PropertyData.DatabaseServices.PropertySet))
                                                    helpme.Clone(sourceDb, destDb, CollDictPropSet, RXClassPS, True)
                                                End If
                                                destTr.Commit()
                                            End Using
                                        End If
                                    End If
                                End If
                            End If
                        Catch ex As System.Exception
                        End Try
                    Next
                    sourceTr.Commit()
                End Using
            End If
        End If
        sourceDb.CloseInput(True)
        FuncCopyLineStylesDestDraw = True
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ВОЗВРАТ ВСЕХ ЗНАЧЕНИЙ ИМЕН СТИЛЕЙ И Т.П С АКТИВНОГО ЧЕРТЕЖА
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает массив со всеми стилями мультивыносок
    Public Shared Function FuncReadListMLeaderStyle(ByRef MLeaderStyleArray() As String) As Integer
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Count As Integer = 0
        If IsArray(MLeaderStyleArray) = True Then
            Count = MLeaderStyleArray.Length
        End If
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim mlstyles As DBDictionary = acTrans.GetObject(acCurDb.MLeaderStyleDictionaryId, OpenMode.ForRead)
                For Each item As DBDictionaryEntry In mlstyles
                    ReDim Preserve MLeaderStyleArray(Count)
                    MLeaderStyleArray(Count) = item.Key
                    Count += 1
                Next
                acTrans.Commit()
            End Using
        Catch
            FuncReadListMLeaderStyle = 0
            Exit Function
        End Try
        FuncReadListMLeaderStyle = Count
    End Function
    '========================================================================================
    'функция возвращает массив со всеми стилями текста
    Public Shared Function FuncReadListTextStyle(ByRef TextStyleArray() As String) As Integer
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Count As Integer = 0
        If IsArray(TextStyleArray) = True Then
            Count = TextStyleArray.Length
        End If
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead)
                For Each id As ObjectId In acTextStyleTbl
                    Dim acTextStyleRec As TextStyleTableRecord = acTrans.GetObject(id, OpenMode.ForRead)
                    ReDim Preserve TextStyleArray(Count)
                    TextStyleArray(Count) = acTextStyleRec.Name
                    Count += 1
                Next
                acTrans.Commit()
            End Using
        Catch
            FuncReadListTextStyle = 0
            Exit Function
        End Try
        FuncReadListTextStyle = Count
    End Function
    '========================================================================================
    'функция возвращает массив со всеми размерными стилями
    Public Shared Function FuncReadListDimStyle(ByRef DimStyleArray() As String) As Integer
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Count As Integer = 0
        If IsArray(DimStyleArray) = True Then
            Count = DimStyleArray.Length
        End If
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acDimStyleTbl As DimStyleTable = acTrans.GetObject(acCurDb.DimStyleTableId, OpenMode.ForRead)
                For Each id As ObjectId In acDimStyleTbl
                    Dim acDimStyleRec As DimStyleTableRecord = acTrans.GetObject(id, OpenMode.ForRead)
                    ReDim Preserve DimStyleArray(Count)
                    DimStyleArray(Count) = acDimStyleRec.Name
                    Count += 1
                Next
                acTrans.Commit()
            End Using
        Catch
            FuncReadListDimStyle = 0
            Exit Function
        End Try
        FuncReadListDimStyle = Count
    End Function
    '========================================================================================
    'функция возвращает массив со всеми слоями активного чертежа
    Public Shared Function FuncReadLayersDWG(ByRef LayerStr() As String, Optional ByRef currentLayer As String = "") As Integer
        FuncReadLayersDWG = 0
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Count As Integer = 0
        If IsArray(LayerStr) = True Then
            If LayerStr.Length > 0 Then
                Count = LayerStr.Length
            End If
        End If
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                For Each id As ObjectId In acLyrTbl
                    Dim acLyrTblRec As LayerTableRecord
                    acLyrTblRec = acTrans.GetObject(id, OpenMode.ForRead)
                    If acLyrTblRec.Id = acCurDb.Clayer Then
                        currentLayer = acLyrTblRec.Name
                    End If
                    ReDim Preserve LayerStr(Count)
                    LayerStr(Count) = acLyrTblRec.Name
                    Count += 1
                Next
                acTrans.Commit()
            End Using
        Catch
            Exit Function
        End Try
        FuncReadLayersDWG = Count
    End Function
    '========================================================================================
    'функция возвращает характеристики слоя
    Public Shared Function FuncReadLayerProperties(ByVal NameLayerStr As String, ByRef LayerProperties As String()) As Boolean
        FuncReadLayerProperties = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                For Each id As ObjectId In acLyrTbl
                    Dim acLyrTblRec As LayerTableRecord = acTrans.GetObject(id, OpenMode.ForRead)
                    If NameLayerStr.Trim Like acLyrTblRec.Name.Trim Then
                        ReDim Preserve LayerProperties(6)
                        LayerProperties(0) = acLyrTblRec.IsOff
                        LayerProperties(1) = acLyrTblRec.IsFrozen
                        LayerProperties(2) = acLyrTblRec.IsLocked
                        LayerProperties(3) = acLyrTblRec.IsPlottable
                        LayerProperties(4) = acLyrTblRec.Color.ColorIndex
                        Dim idLineType As ObjectId = acLyrTblRec.LinetypeObjectId
                        Dim TypeLineStyle As LinetypeTableRecord = idLineType.GetObject(OpenMode.ForRead)
                        LayerProperties(5) = TypeLineStyle.Name
                        LayerProperties(6) = acLyrTblRec.LineWeight
                    End If
                Next
                acTrans.Commit()
                FuncReadLayerProperties = True
            End Using
        Catch ex As System.Exception
            Exit Function
        End Try
    End Function
    '========================================================================================================================================
    'Функция позволяет найти текстовый стиль по строке пути к файлу форм
    Public Shared Function GetObjectIdOfTextStyleWithFileName(ByVal tst As TextStyleTable, ByVal shxFileName As String) As ObjectId
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Try
            Using Tr As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim FullShxFileName As String = HostApplicationServices.Current.FindFile(shxFileName, tst.Database, FindFileHint.FontFile).ToLower()
                For Each idtsr As ObjectId In tst
                    Dim tstr As TextStyleTableRecord = TryCast(Tr.GetObject(idtsr, OpenMode.ForRead), TextStyleTableRecord)
                    If Not tstr.IsShapeFile Then Continue For
                    Dim filename As String = tstr.FileName.ToLower()
                    Try
                        filename = HostApplicationServices.Current.FindFile(filename, tst.Database, FindFileHint.FontFile).ToLower()
                    Catch
                    End Try
                    If FullShxFileName = filename Then Return tstr.ObjectId
                Next
            End Using
        Catch ex As System.Exception
        End Try
        Return ObjectId.Null
    End Function
    '===================================================================================================================================
    'функция возвращает массив со всеми определениями блоков
    Public Shared Function FuncReadListBlockDef(ByRef BlockDefArray() As String) As Integer
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Count As Integer = 0
        If IsArray(BlockDefArray) = True Then
            Count = BlockDefArray.Length
        End If
        Try
            Using acDoc.LockDocument()
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim blockTable As BlockTable = CType(acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead), BlockTable)
                    For Each blockId As ObjectId In blockTable
                        Dim blockTableRecord As BlockTableRecord = CType(acTrans.GetObject(blockId, OpenMode.ForRead), BlockTableRecord)
                        ReDim Preserve BlockDefArray(Count)
                        BlockDefArray(Count) = blockTableRecord.Name
                        Count += 1
                    Next
                End Using
            End Using
            Return Count
        Catch exc As Exception
            Return 0
        End Try
    End Function
    '===================================================================================================================================
    'функция заносит в словарь цвет слоя)
    Public Shared Function FuncGetLayerColors() As Dictionary(Of String, Integer)
        Dim layerColors As New Dictionary(Of String, Integer)()
        Try
            ' Получаем текущий документ
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            ' Начинаем транзакцию
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                ' Открываем таблицу слоев
                Dim lt As LayerTable = trans.GetObject(db.LayerTableId, OpenMode.ForRead)
                ' Перебираем все слои
                For Each layerId As ObjectId In lt
                    Dim layer As LayerTableRecord = trans.GetObject(layerId, OpenMode.ForRead)
                    ' Получаем имя слоя
                    Dim layerName As String = layer.Name
                    ' Получаем цвет слоя
                    Dim color As Autodesk.AutoCAD.Colors.Color = layer.Color
                    Dim colorIndex As Integer
                    ' Преобразуем цвет в индекс (если это индексный цвет)
                    If color.IsByAci Or color.IsByBlock Then
                        colorIndex = color.ColorIndex
                    Else
                        ' Для не-индексных цветов можно использовать -1 или преобразовать в RGB
                        colorIndex = -1
                    End If
                    ' Добавляем в словарь
                    If Not layerColors.ContainsKey(layerName) Then
                        layerColors.Add(layerName, colorIndex)
                    End If
                Next
                ' Не сохраняем транзакцию, так как мы только читаем данные
                'trans.Commit()
            End Using
        Catch ex As Exception
            Application.ShowAlertDialog($"Ошибка: {ex.Message}")
        End Try
        Return layerColors
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ФУНКЦИИ ДЛЯ РАБОТЫ С ШАБЛОНОМ УСЛОВНЫХ ОБОЗНАЧЕНИЙ
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает все свойства объекта в указанной строке таблицы (поиск по 2 колонке "Описание объекта", TypeObjectUser- Полилиния, блок, размер и т.п.)
    Public Shared Function FuncReadPropertyObjectDestDWG(ByVal PutshDrawTemplate As String, ByVal DeskObject As String, ByRef ArrayTableTypeObject(,) As String, Optional TypeObjectUser As String = "", Optional NameTable As String = "") As Boolean
        FuncReadPropertyObjectDestDWG = False
        If IsArray(ArrayTableTypeObject) = True Then
            Erase ArrayTableTypeObject
        End If
        Dim CountArrayTableTypeObject As Integer = 0
        Dim sourceDb As Database = New Database(False, True)
        If File.Exists(PutshDrawTemplate) = False Then
            Exit Function
        End If
        Try
            sourceDb.ReadDwgFile(PutshDrawTemplate, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
            If DeskObject.Trim.Length > 0 Then
                For Each id As ObjectId In sourceBtr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Table Then
                        'получаем доступ к таблице
                        Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = ent
                        'смотрим чтобы у таблицы в обязательном порядке было строк больше 3 и столбцев более 2
                        If tbl.Rows.Count > 3 And tbl.Columns.Count > 2 Then
                            'читаем вторую стоку с названием названим таблицы
                            Dim NameTableRead As String = tbl.Cells(1, 0).Value
                            'если название не совпадает с запросом переходим к следующей таблице
                            If NameTable.Trim.Length > 0 Then
                                If Not (NameTable Like NameTableRead) Then
                                    Continue For
                                End If
                            End If
                            'читаем столбцы таблицы для определения позиций 2-х слолбцов (см. название ниже)
                            Dim PosDeskObject As Integer = -1
                            Dim PosTypeObject As Integer = -1
                            For j As Integer = 1 To tbl.Columns.Count - 1
                                Dim TempNameColumns As String = tbl.Cells(2, j).Value
                                If IsNothing(TempNameColumns) = False Then
                                    If TempNameColumns Like "Type" Then
                                        PosTypeObject = j
                                    ElseIf TempNameColumns Like "Description" Then
                                        PosDeskObject = j
                                    End If
                                    If PosDeskObject > -1 And PosTypeObject > -1 Then
                                        Exit For
                                    End If
                                End If
                            Next j
                            'цикл продолжается если столбцы найдены
                            If PosTypeObject > -1 And PosDeskObject > -1 Then
                                'читаем строки с объектами
                                For i As Integer = 3 To tbl.Rows.Count - 1
                                    Dim rowUser As Autodesk.AutoCAD.DatabaseServices.Row = tbl.Rows.Item(i)
                                    'если нет объединения столбцов то продолжаем цикл
                                    If rowUser.RightColumn > rowUser.LeftColumn Then
                                        Dim DeskObjectTemp As String = tbl.Cells(i, PosDeskObject).Value 'Наименование объекта
                                        Dim TypeObjectTemp As String = tbl.Cells(i, PosTypeObject).Value 'Тип объекта
                                        If IsNothing(DeskObjectTemp) = False And IsNothing(TypeObjectTemp) = False Then
                                            If TypeObjectUser.Trim.Length > 0 Then
                                                If Not (TypeObjectUser Like TypeObjectTemp) Then
                                                    Continue For
                                                End If
                                            End If
                                            'тип объекта нас устраивает заполняем массив с характеристиками
                                            If DeskObjectTemp.Trim Like DeskObject Then
                                                'заполняем массив
                                                For j As Integer = 1 To tbl.Columns.Count - 1
                                                    Dim TempNameUser As String = tbl.Cells(2, j).Value
                                                    Dim Value As String = tbl.Cells(i, j).Value
                                                    If IsNothing(Value) = False And IsNothing(TempNameUser) = False Then
                                                        ReDim Preserve ArrayTableTypeObject(1, CountArrayTableTypeObject)
                                                        ArrayTableTypeObject(0, CountArrayTableTypeObject) = TempNameUser
                                                        ArrayTableTypeObject(1, CountArrayTableTypeObject) = Value
                                                        CountArrayTableTypeObject += 1
                                                    End If
                                                Next j
                                                FuncReadPropertyObjectDestDWG = True
                                                GoTo line1
                                            End If
                                        End If
                                    End If
                                Next i
                            End If
                        End If
                    End If
                Next
            End If
line1:
        End Using
        sourceDb.CloseInput(True)
    End Function
    '========================================================================================================================================================================
    'функция возвращает все описания стилей блоков, типов линий, размеров и т.п из указанной таблицы
    Public Shared Function FuncReadDeskriptionObjectAllDestDWG(ByVal PutshDrawTemplate As String, ByVal NameTable As String, ByRef ArrayDeskriptionObject() As String, Optional TypeObjectUser As String = "", Optional DeskTable As String = "") As Boolean
        FuncReadDeskriptionObjectAllDestDWG = False
        Dim CountArray As Integer = 0
        If IsArray(ArrayDeskriptionObject) = True Then
            CountArray = ArrayDeskriptionObject.Length
        End If
        Dim sourceDb As Database = New Database(False, True)
        If File.Exists(PutshDrawTemplate) = False Then
            Exit Function
        End If
        Try
            sourceDb.ReadDwgFile(PutshDrawTemplate, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
            For Each id As ObjectId In sourceBtr
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Table Then
                    Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = ent
                    If tbl.Rows.Count > 3 And tbl.Columns.Count > 2 Then
                        Dim TempNameTbl As String = tbl.Cells(1, 0).Value 'читаем вторую строку (имя таблицы)
                        Dim TempDeskTbl As String = tbl.Cells(0, 0).Value 'читаем первую строку (описание таблицы)
                        If TempNameTbl.Trim Like NameTable Then
                            If TempDeskTbl.Trim.Length > 0 Then
                                If Not (TempDeskTbl.Trim Like DeskTable) Then
                                    Continue For
                                End If
                            End If
                            'читаем заголовки таблицы
                            Dim PosDeskObject As Integer = -1
                            Dim PosTypeObject As Integer = -1
                            Dim rowUser As Autodesk.AutoCAD.DatabaseServices.Row = tbl.Rows.Item(2)
                            If rowUser.RightColumn - rowUser.LeftColumn = tbl.Columns.Count - 1 Then
                                For j As Integer = 0 To tbl.Columns.Count - 1
                                    Dim DeskColTemp As String = tbl.Cells(2, j).Value 'Описание объекта
                                    If DeskColTemp.Trim Like "Description" Then
                                        PosDeskObject = j
                                    ElseIf DeskColTemp.Trim Like "Type" Then
                                        PosTypeObject = j
                                    End If
                                Next j
                            End If
                            'читаем строки таблицы
                            If PosDeskObject > -1 Then
                                For j As Integer = 3 To tbl.Rows.Count - 1
                                    rowUser = tbl.Rows.Item(j)
                                    If rowUser.RightColumn - rowUser.LeftColumn = tbl.Columns.Count - 1 Then
                                        Dim NameObjectTemp As String = tbl.Cells(j, PosDeskObject).Value 'Описание объекта
                                        Dim TypeObjectTemp As String = tbl.Cells(j, PosTypeObject).Value 'Тип объекта
                                        If IsNothing(NameObjectTemp) = False Then
                                            If Not (TypeObjectUser Like "") Then
                                                If IsNothing(TypeObjectTemp) = False Then
                                                    If Not (TypeObjectTemp.Trim Like TypeObjectUser) Then
                                                        Continue For
                                                    Else
                                                        ReDim Preserve ArrayDeskriptionObject(CountArray)
                                                        ArrayDeskriptionObject(CountArray) = NameObjectTemp
                                                        CountArray += 1
                                                    End If
                                                End If
                                            Else
                                                ReDim Preserve ArrayDeskriptionObject(CountArray)
                                                ArrayDeskriptionObject(CountArray) = NameObjectTemp
                                                CountArray += 1
                                            End If
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                    End If
                End If
            Next
        End Using
    End Function
    '========================================================================================================================================================================
    'функция возвращает все имена стилей блоков, типов линий, размеров и т.п из указанной таблицы
    Public Shared Function FuncReadNameObjectAllDestDWG(ByVal PutshDrawTemplate As String, ByVal NameTables As String(), ByRef ArrayTableNameObject() As String, Optional TypeObjectUser As String = "") As Boolean
        FuncReadNameObjectAllDestDWG = False
        Dim CountArray As Integer = 0
        If IsArray(ArrayTableNameObject) = True Then
            Erase ArrayTableNameObject
        End If
        If IsNothing(NameTables) = True Then
            Exit Function
        End If
        If IsArray(NameTables) = False Then
            Exit Function
        End If

        Dim CountArrayTableTypeObject As Integer = 0
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDrawTemplate, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
            For i As Integer = 0 To NameTables.Length - 1
                Dim nameTable As String = NameTables(i)
                For Each id As ObjectId In sourceBtr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Table Then
                        Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = ent
                        If tbl.Rows.Count > 3 And tbl.Columns.Count > 2 Then
                            Dim TempNameTbl As String = tbl.Cells(1, 0).Value 'читаем вторую строку
                            If TempNameTbl.Trim Like nameTable Then
                                'читаем заголовки таблицы
                                Dim PosNameObject As Integer = -1
                                Dim PosTypeObject As Integer = -1
                                Dim rowUser As Autodesk.AutoCAD.DatabaseServices.Row = tbl.Rows.Item(2)
                                If rowUser.RightColumn - rowUser.LeftColumn = tbl.Columns.Count - 1 Then
                                    For j As Integer = 0 To tbl.Columns.Count - 1
                                        Dim NameColTemp As String = tbl.Cells(2, j).Value 'Наименование объекта
                                        If NameColTemp.Trim Like "Name" Then
                                            PosNameObject = j
                                        ElseIf NameColTemp.Trim Like "Type" Then
                                            PosTypeObject = j
                                        End If
                                    Next j
                                End If
                                'читаем строки таблицы
                                If PosNameObject > -1 Then
                                    For j As Integer = 3 To tbl.Rows.Count - 1
                                        rowUser = tbl.Rows.Item(j)
                                        If rowUser.RightColumn - rowUser.LeftColumn = tbl.Columns.Count - 1 Then
                                            Dim NameObjectTemp As String = tbl.Cells(j, PosNameObject).Value 'Наименование объекта
                                            Dim TypeObjectTemp As String = tbl.Cells(j, PosTypeObject).Value 'Тип объекта
                                            If IsNothing(NameObjectTemp) = False Then
                                                If Not (TypeObjectUser Like "") Then
                                                    If IsNothing(TypeObjectTemp) = False Then
                                                        If Not (TypeObjectTemp.Trim Like TypeObjectUser) Then
                                                            Continue For
                                                        Else
                                                            ReDim Preserve ArrayTableNameObject(CountArray)
                                                            ArrayTableNameObject(CountArray) = NameObjectTemp
                                                            CountArray += 1
                                                        End If
                                                    End If
                                                Else
                                                    ReDim Preserve ArrayTableNameObject(CountArray)
                                                    ArrayTableNameObject(CountArray) = NameObjectTemp
                                                    CountArray += 1
                                                End If
                                            End If
                                        End If
                                    Next j
                                End If
                            End If
                        End If
                    End If
                Next
            Next i
        End Using
    End Function
    '========================================================================================================================================================================
    'функция возвращает все объекты из указанной таблицы (путь, имя таблицы (2 строка), опимание таблицы (1 строка),массив с возвращаемыми данными)
    Public Shared Function FuncReadTableDataDestDWG(ByVal PutshDrawTemplate As String, ByVal NameTable As String, ByVal DeskTable As String, ByRef ArrayDataTable(,) As String, ByVal TypeObject As String) As Boolean
        FuncReadTableDataDestDWG = False
        Dim CountArray As Integer = 0
        If IsArray(FuncReadTableDataDestDWG) = True Then
            Erase ArrayDataTable
        End If

        Dim CountArrayTableTypeObject As Integer = 0
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDrawTemplate, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
            For Each id As ObjectId In sourceBtr
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Table Then
                    Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = ent
                    If tbl.Rows.Count > 3 And tbl.Columns.Count > 2 Then
                        Dim DeskTbl As String = tbl.Cells(0, 0).Value 'читаем первую строку
                        Dim NameTbl As String = tbl.Cells(1, 0).Value 'читаем вторую строку
                        If NameTbl.Trim Like NameTable Then
                            If DeskTbl.Trim Like DeskTable Then
                                'читаем строки таблицы
                                Dim CountRows As Integer = 0
                                Dim CountColumns As Integer = tbl.Columns.Count - 2
                                Dim PosTypeObject As Integer = 0
                                For i As Integer = 1 To CountColumns
                                    Dim ValData As String = tbl.Cells(2, i).Value 'читаем заголовки
                                    If ValData Like "Тип объекта" Then
                                        PosTypeObject = i
                                        Exit For
                                    End If
                                Next

                                For j As Integer = 3 To tbl.Rows.Count - 1
                                    Dim TempvalData As String = tbl.Cells(j, PosTypeObject).Value
                                    If TempvalData Like TypeObject Then
                                        For k As Integer = 1 To tbl.Columns.Count - 1
                                            Dim DataField As String = tbl.Cells(2, k).Value 'читаем данные
                                            Dim ValData As String = tbl.Cells(j, k).Value 'читаем данные
                                            If IsNothing(ValData) = True Then ValData = ""
                                            If ValData Like "" And k = 1 Then
                                                Exit For
                                            End If
                                            ReDim Preserve ArrayDataTable(1, CountArray)
                                            ArrayDataTable(0, CountArray) = DataField
                                            ArrayDataTable(1, CountArray) = ValData
                                            CountArray += 1
                                        Next k
                                    End If
                                Next j
                            End If
                        End If
                    End If
                End If
            Next
        End Using
        FuncReadTableDataDestDWG = True
    End Function
    '========================================================================================================================================================================
    'функция возвращает все объекты из указанной таблицы (путь, имя таблицы (2 строка), опимание таблицы (1 строка),массив с возвращаемыми данными, 1 - элемент массива это шапка таблицы)
    Public Shared Function FuncReadTableDataTemplateDestDWG(ByVal PutshDrawTemplate As String, ByVal NameTable As String, ByVal DeskTable As String, ByRef ArrayDataTable(,) As String, ByVal TypeObject As String) As Boolean
        FuncReadTableDataTemplateDestDWG = False
        Dim CountArray As Integer = 0
        If IsArray(ArrayDataTable) = True Then
            Erase ArrayDataTable
        End If
        Dim CountArrayDataTable As Integer = 1

        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDrawTemplate, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
            For Each id As ObjectId In sourceBtr
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Table Then
                    Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = ent
                    If tbl.Rows.Count > 3 And tbl.Columns.Count > 2 Then
                        Dim DeskTbl As String = tbl.Cells(0, 0).TextString 'читаем первую строку
                        Dim NameTbl As String = tbl.Cells(1, 0).TextString 'читаем вторую строку
                        If NameTbl.Trim Like NameTable Then
                            If DeskTbl.Trim Like DeskTable Then
                                Dim CountColumns As Integer = tbl.Columns.Count 'количество столбцов
                                Dim PosTypeObject As Integer = 0
                                Erase ArrayDataTable
                                CountArrayDataTable = 1
                                ReDim Preserve ArrayDataTable(CountColumns - 1, 0)
                                For i As Integer = 0 To CountColumns - 1
                                    Dim ValData As String = tbl.Cells(2, i).TextString 'читаем заголовки столбцов
                                    ArrayDataTable(i, 0) = ValData
                                    If ValData Like "Type" Then
                                        PosTypeObject = i
                                    End If
                                Next
                                '================================================================================
                                'читаем остальные строки
                                For j As Integer = 3 To tbl.Rows.Count - 1
                                    Dim TempvalData As String = tbl.Cells(j, PosTypeObject).TextString
                                    If TempvalData Like TypeObject Then
                                        For k As Integer = 0 To tbl.Columns.Count - 1
                                            Dim ValData As String = tbl.Cells(j, k).TextString 'читаем данные
                                            If IsNothing(ValData) = True Then ValData = ""
                                            If ValData Like "" And k = 1 Then
                                                Exit For
                                            End If
                                            ReDim Preserve ArrayDataTable(CountColumns - 1, CountArrayDataTable)
                                            ArrayDataTable(k, CountArrayDataTable) = ValData
                                        Next k
                                    End If
                                    CountArrayDataTable += 1
                                Next j
                                If IsArray(ArrayDataTable) = True Then
                                    If ArrayDataTable.GetUpperBound(1) > 0 Then
                                        GoTo Line1
                                    Else
                                        Erase ArrayDataTable
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
Line1:
        End Using
        sourceDb.CloseInput(True)
        FuncReadTableDataTemplateDestDWG = True
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ПРОСТОЕ СОЗДАНИЕ ОБЪЕКТОВ ТИПА AUTOCAD
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '=================================================================================================

    '=====================================================================================================
    'функция вставляет многострочный текст
    Public Shared Function FuncDrawMText(ByVal pointIns As Point3d, ByVal MTextContext As String, ByVal HMText As Double, ByVal NameLayer As String, Optional ByVal StyleMText As ObjectId = Nothing, Optional ByVal AngleMText As Double = 0, Optional ByVal AttachMText As Integer = 10, Optional ByVal indColor As Integer = 7) As Boolean
        FuncDrawMText = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        If IsNothing(pointIns) = True Then
            Exit Function
        End If
        If MTextContext.Trim.Length = 0 Then
            Exit Function
        End If
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim acMText As MText = New MText()
                acMText.SetDatabaseDefaults()
                If HMText > 0 Then
                    acMText.TextHeight = HMText
                End If
                If IsNothing(pointIns) = False Then
                    acMText.Location = pointIns
                End If
                acMText.Contents = MTextContext
                acMText.Rotation = AngleMText
                If NameLayer.Trim.Length > 0 Then
                    Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                    If acLyrTbl.Has(NameLayer) = True Then
                        acMText.Layer = NameLayer
                    End If
                End If
                If IsNothing(StyleMText) = False Then
                    acMText.TextStyleId = StyleMText
                End If
                acMText.ColorIndex = indColor
                Try
                    If AttachMText = 22 Then
                        acMText.Attachment = AttachmentPoint.BottomCenter
                    ElseIf AttachMText = 24 Then
                        acMText.Attachment = AttachmentPoint.TopCenter
                    ElseIf AttachMText = 5 Then
                        acMText.Attachment = AttachmentPoint.MiddleCenter
                    Else
                        acMText.Attachment = AttachmentPoint.BaseLeft
                    End If
                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                End Try

                acBlkTblRec.AppendEntity(acMText)
                acTrans.AddNewlyCreatedDBObject(acMText, True)
                acTrans.Commit()
            End Using
        Catch ex As System.Exception
        End Try
        FuncDrawMText = True
    End Function
    '==================================================================================================
    'функция вставляет мультивыноску в чертеж (точка вставки, текст, смещениеX, смещениеY,
    Public Shared Function FuncDrawMleader(ByVal pointIns As Point2d, ByVal MTextIns As String, ByVal OffsetX As Double, ByVal offsetY As Double, Optional ByVal NameLayer As String = "0", Optional ByVal HText As Double = 0, Optional ByVal StyleMLeader As ObjectId = Nothing, Optional ByVal StyleTextID As ObjectId = Nothing, Optional ByVal IndexColorText As Double = -1, Optional ByVal Background As Boolean = True, Optional rotateTextByLine As Boolean = False) As ObjectId
        FuncDrawMleader = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim mlStyles As DBDictionary = acTrans.GetObject(acCurDb.MLeaderStyleDictionaryId, OpenMode.ForRead)
                Dim MLeader1 As MLeader = New MLeader
                MLeader1.SetDatabaseDefaults()
                Dim leaderIndex As Integer = MLeader1.AddLeader()
                Dim leaderLineIndex As Integer = MLeader1.AddLeaderLine(leaderIndex)
                MLeader1.AddFirstVertex(leaderLineIndex, New Point3d(pointIns.X, pointIns.Y, 0))
                MLeader1.AddLastVertex(leaderLineIndex, New Point3d(pointIns.X + OffsetX, pointIns.Y + offsetY, 0))
                MLeader1.ContentType = ContentType.MTextContent
                If IsNothing(StyleMLeader) = False Then
                    If StyleMLeader.IsNull = False Then
                        If mlStyles.Contains(StyleMLeader) = True Then
                            MLeader1.MLeaderStyle = StyleMLeader
                        End If
                    End If
                End If
                Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                If acLyrTbl.Has(NameLayer) = True Then
                    MLeader1.Layer = NameLayer
                End If
                'добавляем многострочный текст
                If MTextIns.Trim.Length > 0 Then
                    Dim acMText As MText = New MText()
                    acMText.SetDatabaseDefaults()
                    If HText > 0 Then
                        acMText.TextHeight = HText
                    End If
                    If IsNothing(StyleTextID) = False Then
                        If StyleTextID.IsNull = False Then
                            Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead)
                            If acTextStyleTbl.Has(StyleTextID) = True Then
                                Dim acTextStyleRec As TextStyleTableRecord = acTrans.GetObject(StyleTextID, OpenMode.ForRead)
                                acMText.TextStyleId = StyleTextID
                            End If
                        End If
                    End If
                    If IndexColorText > -1 Then
                        acMText.ColorIndex = IndexColorText
                    Else
                        If StyleMLeader.IsNull = False Then
                            Dim mlStylesTemp As MLeaderStyle = StyleMLeader.GetObject(OpenMode.ForRead)
                            acMText.ColorIndex = mlStylesTemp.DefaultMText.ColorIndex
                        Else
                            acMText.ColorIndex = 256
                        End If

                    End If
                    If rotateTextByLine = True Then
                        Dim pt1 As Point2d = MLeader1.GetFirstVertex(leaderIndex).Convert2d(New Plane)
                        Dim pt2 As Point2d = MLeader1.GetLastVertex(leaderIndex).Convert2d(New Plane)
                        acMText.Rotation = pt1.GetVectorTo(pt2).Angle
                    End If
                    acMText.BackgroundFill = Background
                    acMText.Contents = MTextIns
                    MLeader1.MText = acMText
                Else
                    If StyleMLeader.IsNull = False Then
                        Dim mlStylesTemp As MLeaderStyle = StyleMLeader.GetObject(OpenMode.ForRead)
                        MLeader1.MText = mlStylesTemp.DefaultMText
                        If IndexColorText > -1 Then
                            MLeader1.MText.ColorIndex = mlStylesTemp.DefaultMText.ColorIndex
                        End If
                    End If
                    MLeader1.MText.BackgroundFill = Background
                    If rotateTextByLine = True Then
                        Dim pt1 As Point2d = MLeader1.GetFirstVertex(leaderIndex).Convert2d(New Plane)
                        Dim pt2 As Point2d = MLeader1.GetLastVertex(leaderIndex).Convert2d(New Plane)
                        Dim userMtext As MText = MLeader1.MText
                        userMtext.Rotation = pt1.GetVectorTo(pt2).Angle
                        MLeader1.MText = userMtext
                    End If
                End If
                acBlkTblRec.AppendEntity(MLeader1)
                acTrans.AddNewlyCreatedDBObject(MLeader1, True)
                FuncDrawMleader = MLeader1.Id
                acTrans.Commit()
            End Using
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try
    End Function
    '==================================================================================================
    'функция вставляет мультивыноску в чертеж (точка вставки, текст, смещениеX, смещениеY,) новая версия
    Public Shared Function FuncDrawMleader2(ByVal firstVertex As Point3d, ByVal contextText As String, Optional lastVertex As Point3d = Nothing, Optional ByVal idStyleMLeader As ObjectId = Nothing, Optional ByVal NameLayer As String = "0", Optional ByVal M As Double = 1, Optional ByVal StyleTextID As ObjectId = Nothing, Optional ByVal IndexColorText As Double = -1, Optional ByVal Background As Boolean = True) As MLeader
        FuncDrawMleader2 = Nothing
        Dim myMLeaderId As ObjectId = ObjectId.Null
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim mlStyles As DBDictionary = acTrans.GetObject(acCurDb.MLeaderStyleDictionaryId, OpenMode.ForRead)
                Dim myleaderStyleId As ObjectId = acCurDb.MLeaderstyle
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
                            Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead)
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
                Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                If acLyrTbl.Has(NameLayer) = True Then
                    myMLeader.Layer = NameLayer
                End If
                myMLeader.ArrowSize = myMLeader.ArrowSize * M
                FuncDrawMleader2 = myMLeader
                acBlkTblRec.AppendEntity(myMLeader)
                acTrans.AddNewlyCreatedDBObject(myMLeader, True)
                acTrans.Commit()
            End Using
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try
    End Function
    '===================================================================================================
    'функция рисует полилинию 
    Public Shared Function FuncDrawPolylineByArrayProperties(ByVal ArrayVertex As Double(,), ArrayProperty As String(,), Optional ClosedPline As Boolean = False) As Polyline
        FuncDrawPolylineByArrayProperties = Nothing
        If IsArray(ArrayVertex) = False Then Exit Function
        If ArrayVertex.GetUpperBound(1) < 1 Then Exit Function

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Using Tr As Transaction = db.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = Tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = Tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline()
            acPoly.SetDatabaseDefaults()
            Dim PrevVertex As Point2d = Nothing
            Dim CountVertex As Integer = 0
            For i As Integer = 0 To ArrayVertex.GetUpperBound(1)
                Try
                    Dim Pt1 As Point2d = New Point2d(ArrayVertex(0, i), ArrayVertex(1, i))
                    If i > 0 Then
                        Dim Dist As Double = Pt1.GetVectorTo(PrevVertex).Length
                        If Dist > 0 Then
                            acPoly.AddVertexAt(CountVertex, Pt1, 0, 0, 0)
                            PrevVertex = Pt1
                            CountVertex += 1
                        End If
                    Else
                        acPoly.AddVertexAt(CountVertex, Pt1, 0, 0, 0)
                        PrevVertex = Pt1
                        CountVertex += 1
                    End If
                Catch
                End Try
            Next i
            acBlkTblRec.AppendEntity(acPoly)
            Tr.AddNewlyCreatedDBObject(acPoly, True)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'назначаем характеристики созданной полилинии
            If IsArray(ArrayProperty) = False Then
                If IsNothing(ArrayProperty) = False Then
                    For i As Integer = 0 To ArrayProperty.GetUpperBound(1)
                        Dim NameField As String = ArrayProperty(0, i)
                        If NameField.Trim Like "Layer" Then
                            Dim NameLayer As String = ArrayProperty(1, i)
                            If IsNothing(NameLayer) = False Then
                                If NameLayer.Trim.Length > 0 Then
                                    Dim acLyrTbl As LayerTable = Tr.GetObject(db.LayerTableId, OpenMode.ForRead)
                                    If acLyrTbl.Has(NameLayer) = True Then
                                        acPoly.Layer = NameLayer
                                    End If
                                End If
                            End If
                        ElseIf NameField.Trim Like "Color" Then
                            If IsNothing(ArrayProperty(1, i)) = False Then
                                Dim ColorLin As Integer = Val(ArrayProperty(1, i))
                                If ColorLin < 0 Then
                                    ColorLin = 0
                                ElseIf ColorLin > 256 Then
                                    ColorLin = 256
                                End If
                                acPoly.ColorIndex = ColorLin
                            End If
                        ElseIf NameField.Trim Like "LineScale" Then
                            If IsNothing(ArrayProperty(1, i)) = False Then
                                Dim ScaleLine As Integer = Val(ArrayProperty(1, i))
                                If ScaleLine > 0 Then
                                    acPoly.LinetypeScale = ScaleLine
                                End If
                            End If
                        ElseIf NameField.Trim Like "Name" Then
                            Dim NameTypeLine As String = ArrayProperty(1, i)
                            If IsNothing(NameTypeLine) = False Then
                                If NameTypeLine.Trim.Length > 0 Then
                                    Dim acLineTypTbl As LinetypeTable = Tr.GetObject(db.LinetypeTableId, OpenMode.ForRead)
                                    If acLineTypTbl.Has(NameTypeLine) = True Then
                                        acPoly.Linetype = NameTypeLine
                                    End If
                                End If
                            End If
                        ElseIf NameField.Trim Like "LineWidth" Then
                            If IsNothing(ArrayProperty(1, i)) = False Then
                                Dim WidthLin As Integer = Val(ArrayProperty(1, i))
                                If WidthLin > 0 Then
                                    acPoly.LineWeight = WidthLin
                                End If
                            End If
                        ElseIf NameField.Trim Like "LineGWidth" Then
                            If IsNothing(ArrayProperty(1, i)) = False Then
                                Dim Condlin As Double = Val(ArrayProperty(1, i))
                                If Condlin > 0 Then
                                    acPoly.ConstantWidth = Condlin
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
            acPoly.Closed = ClosedPline
            Tr.Commit()
            FuncDrawPolylineByArrayProperties = acPoly
        End Using
    End Function
    '===============================================================================================
    'функция рисует полилинию
    Public Shared Function FuncDrawPolyline(ByVal ArrayVertex As Double(,), Optional ByVal nameLayer As String = "0", Optional ByVal indColor As String = "7", Optional ByVal LW As LineWeight = LineWeight.LineWeight020, Optional ByVal BClose As Boolean = False, Optional ByVal ScalePl As Double = 1, Optional ByVal TypeLinePl As String = "") As ObjectId
        FuncDrawPolyline = Nothing
        If IsArray(ArrayVertex) = False Then Exit Function
        If ArrayVertex.GetUpperBound(1) < 1 Then Exit Function
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline()
                For i As Integer = 0 To ArrayVertex.GetUpperBound(1)
                    Try
                        Dim Pt1 As Point2d = New Point2d(ArrayVertex(0, i), ArrayVertex(1, i))
                        acPoly.AddVertexAt(i, Pt1, 0, 0, 0)
                    Catch
                    End Try
                Next i
                acPoly.Layer = nameLayer
                acPoly.ColorIndex = indColor
                acPoly.LineWeight = LW
                acPoly.Closed = BClose
                acPoly.LinetypeScale = ScalePl
                If TypeLinePl.Trim.Length > 0 Then
                    acPoly.Linetype = TypeLinePl
                End If
                acPoly.SetDatabaseDefaults()
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
                acTrans.Commit()
                FuncDrawPolyline = acPoly.ObjectId
            End Using
        Catch ex As System.Exception
        End Try
    End Function
    '===============================================================================================
    'функция рисует полилинию
    Public Shared Function createPolylineFromArrayCoordinates2(ByVal arrayCoordinates(,) As String) As Polyline
        ' Проверяем, что массив не пустой и содержит достаточно данных
        If arrayCoordinates Is Nothing OrElse arrayCoordinates.GetLength(0) < 2 Then
            Return Nothing
        End If

        ' Проверяем, что во второй размерности есть хотя бы 2 элемента (X и Y)
        If arrayCoordinates.GetLength(1) < 2 Then
            Return Nothing
        End If

        ' Создаем новую полилинию
        Dim polyline As New Polyline()

        Try
            ' Получаем количество точек (количество строк в первой размерности)
            Dim pointCount As Integer = arrayCoordinates.GetLength(0)

            ' Проходим по всем точкам (по строкам массива)
            For i As Integer = 0 To pointCount - 1
                ' Парсим координаты
                Dim x As Double
                Dim y As Double

                ' arrayCoordinates(i, 0) - X координата
                ' arrayCoordinates(i, 1) - Y координата
                If Not Double.TryParse(arrayCoordinates(i, 0), x) Then
                    Throw New ArgumentException($"Не удалось преобразовать X координату в строке {i + 1}: {arrayCoordinates(i, 0)}")
                End If

                If Not Double.TryParse(arrayCoordinates(i, 1), y) Then
                    Throw New ArgumentException($"Не удалось преобразовать Y координату в строке {i + 1}: {arrayCoordinates(i, 1)}")
                End If

                ' Создаем точку в 2D (Z = 0)
                Dim point2d As New Point2d(x, y)

                ' Добавляем вершину к полилинии
                polyline.AddVertexAt(i, point2d, 0, 0, 0)
            Next

            ' Если нужно замкнуть полилинию, раскомментируйте следующую строку
            ' polyline.Closed = True

        Catch ex As Exception
            ' В случае ошибки возвращаем Nothing
            Return Nothing
        End Try

        Return polyline
    End Function




    '===============================================================================================
    'функция рисует полилинию на основе Polygon
    '===============================================================================================
    'функция рисует полилинию
    Public Shared Function FuncDrawPolylineByPolygon(ByVal userGeometry As NetTopologySuite.Geometries.Geometry) As Polyline()
        Dim result As Polyline() = {}
        Dim countPoly As Integer = 0
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        If IsNothing(userGeometry) = False Then
            If userGeometry.Coordinates.Count > 0 Then
                Dim countPolygon As Integer = userGeometry.Boundary.NumGeometries
                If countPolygon > 0 Then
                    Try
                        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                            For i As Integer = 0 To countPolygon - 1
                                Dim tempGeo As NetTopologySuite.Geometries.Geometry = userGeometry.Boundary.GetGeometryN(i)
                                Dim coordPoly As List(Of NetTopologySuite.Geometries.Coordinate) = tempGeo.Coordinates.ToList
                                Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline()
                                For j As Integer = 0 To coordPoly.Count - 1
                                    Dim coodPoint As NetTopologySuite.Geometries.Coordinate = coordPoly.Item(j)
                                    acPoly.AddVertexAt(i, New Point2d(coodPoint.X, coodPoint.Y), 0, 0, 0)
                                Next j
                                If acPoly.Area > 0 Then
                                    acPoly.Closed = True
                                    acPoly.SetDatabaseDefaults()
                                    acBlkTblRec.AppendEntity(acPoly)
                                    acTrans.AddNewlyCreatedDBObject(acPoly, True)
                                    ReDim Preserve result(countPoly)
                                    result(countPoly) = acPoly
                                    countPoly += 1
                                End If
                            Next i
                            acTrans.Commit()
                        End Using
                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                    End Try
                End If
            End If
        End If
        Return result
    End Function

    '===============================================================================================
    'функция рисует мультилинию
    Public Shared Function FuncDrawMLine(ByVal ArrayVertex As Double(,), Optional ByVal nameLayer As String = "0", Optional ByVal indColor As String = "7", Optional ByVal LW As LineWeight = LineWeight.LineWeight020, Optional ByVal ScaleMl As Double = 1, Optional ByVal TypeMLine As String = "", Optional ByVal Jusific As Integer = 1) As ObjectId
        FuncDrawMLine = Nothing
        If IsArray(ArrayVertex) = False Then Exit Function
        If ArrayVertex.GetUpperBound(1) < 1 Then Exit Function
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim acMline As Mline = New Mline()
            acMline.Normal = New Vector3d(0.0, 0.0, 1.0)
            acMline.SetDatabaseDefaults()
            '====================================================================================================
            'стиль мультилинии
            If TypeMLine.Trim.Length > 0 Then
                Dim mlDict As DBDictionary = acTrans.GetObject(acCurDb.MLStyleDictionaryId, OpenMode.ForRead)
                If mlDict.Contains(TypeMLine) = True Then
                    Dim IDTypeMLine As ObjectId = mlDict.Item(TypeMLine)
                    acMline.Style = IDTypeMLine
                Else
                    acMline.Style = acCurDb.CmlstyleID
                End If
            Else
                acMline.Style = acCurDb.CmlstyleID
            End If
            '====================================================================================================
            'выравнивание
            If Jusific = 1 Then
                acMline.Justification = MlineJustification.Zero
            ElseIf Jusific = 0 Then
                acMline.Justification = MlineJustification.Top
            ElseIf Jusific = 2 Then
                acMline.Justification = MlineJustification.Bottom
            Else
                acMline.Justification = MlineJustification.Zero
            End If
            '====================================================================================================
            'масштаб
            If ScaleMl > 0 Then
                acMline.Scale = ScaleMl
            Else
                acMline.Scale = Math.Abs(ScaleMl)
            End If
            '=====================================================================================================
            'цвет
            Dim IndexColor As Integer = 256
            If indColor.Trim Like "ПоСлою" Then
                IndexColor = 256
            ElseIf indColor.Trim Like "ПоБлоку" Then
                IndexColor = 0
            Else
                If IsNumeric(indColor) = False Then
                    IndexColor = 256
                End If
            End If
            If IndexColor < 0 Then
                IndexColor = 0
            ElseIf IndexColor > 256 Then
                IndexColor = 256
            Else
                IndexColor = Val(indColor)
            End If
            acMline.ColorIndex = IndexColor
            '===================================================================================================
            'толщина
            acMline.LineWeight = LW
            '===================================================================================================
            'делаем линию
            For i As Integer = 0 To ArrayVertex.GetUpperBound(1)
                Try
                    Dim Pt1 As Point3d = New Point3d(ArrayVertex(0, i), ArrayVertex(1, i), 0)
                    acMline.AppendSegment(Pt1)
                Catch
                End Try
            Next i
            acBlkTblRec.AppendEntity(acMline)
            acTrans.AddNewlyCreatedDBObject(acMline, True)
            acTrans.Commit()
            FuncDrawMLine = acMline.ObjectId
        End Using
    End Function



    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'СОЗДАНИЕ ОБЪЕКТОВ ТИПА AUTOCAD ПО СВОЙСТВАМ ВЗЯТЫМ ИЗ ФАЙЛА-ЩАБЛОНА
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'работа с блоками
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ghjcnfz dcnfdrf ,kjrf
    Public Shared Function FuncInsertBlk(ByVal pointBlkXY As Point3d, ByVal NameBlock As String, Optional NameLayer As String = "0", Optional ColorBlk As Integer = 7, Optional ScaleBlkXY As Double = 1, Optional LineWight As Double = 20, Optional rot As Double = 0, Optional ScaleBlkZ As Double = 0) As BlockReference
        Dim blk As BlockReference = Nothing
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim tr As Transaction = doc.TransactionManager.StartTransaction()
        If IsNothing(NameBlock) = True Then Exit Function
        If NameBlock.Trim.Length = 0 Then Exit Function
        Using tr
            Try
                Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                If bt.Has(NameBlock) = True Then
                    Dim space As BlockTableRecord = DirectCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                    Dim btr As BlockTableRecord = DirectCast(tr.GetObject(bt(NameBlock), OpenMode.ForRead), BlockTableRecord)
                    blk = New BlockReference(pointBlkXY, btr.ObjectId)
                    space.AppendEntity(blk)
                    tr.AddNewlyCreatedDBObject(blk, True)
                    'присваиваем характеристики блоку
                    'слой
                    Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
                    If acLyrTbl.Has(NameLayer) = True Then
                        blk.Layer = NameLayer
                    End If
                    'цвет
                    If ColorBlk < 0 Then
                        ColorBlk = 0
                    ElseIf ColorBlk > 256 Then
                        ColorBlk = 256
                    End If
                    blk.ColorIndex = ColorBlk
                    'масштаб
                    If ScaleBlkXY > 0 Then
                        blk.ScaleFactors = New Scale3d(ScaleBlkXY, ScaleBlkXY, ScaleBlkXY)
                    End If
                    If ScaleBlkZ > 0 Then
                        blk.ScaleFactors = New Scale3d(ScaleBlkXY, ScaleBlkXY, ScaleBlkZ)
                    End If
                    'вес линии
                    If LineWight > 0 Then
                        blk.LineWeight = LineWight
                    Else
                        If LineWight = -2 Then
                            blk.LineWeight = LineWeight.ByBlock
                        Else
                            blk.LineWeight = LineWeight.ByLayer
                        End If
                    End If
                    'поворот блока
                    blk.Rotation = rot
                End If
            Catch
                tr.Abort()
                Return Nothing
            End Try
            tr.Commit()
        End Using
        Return blk
    End Function


    '===================================================================================
    'вставка в Чертеж блока (в пространство модели) со свойсвами определенными отдельными переменными
    Public Shared Function FuncInsertDrawingBlk(ByVal pointBlkXY As Point3d, ByVal NameBlock As String, Optional NameLayer As String = "0", Optional ColorBlk As Integer = 7, Optional ScaleBlkXY As Double = 1, Optional LineWight As Double = 20, Optional rot As Double = 0, Optional ScaleBlkZ As Double = 0) As BlockReference
        FuncInsertDrawingBlk = Nothing
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim tr As Transaction = doc.TransactionManager.StartTransaction()
        If IsNothing(NameBlock) = True Then Exit Function
        If NameBlock.Trim.Length = 0 Then Exit Function
        Using tr
            Try
                Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                If bt.Has(NameBlock) = True Then
                    Dim space As BlockTableRecord = DirectCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                    Dim btr As BlockTableRecord = DirectCast(tr.GetObject(bt(NameBlock), OpenMode.ForRead), BlockTableRecord)
                    Dim br As New BlockReference(pointBlkXY, btr.ObjectId)
                    space.AppendEntity(br)
                    tr.AddNewlyCreatedDBObject(br, True)
                    'присваиваем характеристики блоку
                    'слой
                    Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
                    If acLyrTbl.Has(NameLayer) = True Then
                        br.Layer = NameLayer
                    End If
                    'цвет
                    If ColorBlk < 0 Then
                        ColorBlk = 0
                    ElseIf ColorBlk > 256 Then
                        ColorBlk = 256
                    End If
                    br.ColorIndex = ColorBlk
                    'масштаб
                    If ScaleBlkXY > 0 Then
                        br.ScaleFactors = New Scale3d(ScaleBlkXY, ScaleBlkXY, ScaleBlkXY)
                    End If
                    If ScaleBlkZ > 0 Then
                        br.ScaleFactors = New Scale3d(ScaleBlkXY, ScaleBlkXY, ScaleBlkZ)
                    End If
                    'вес линии
                    If LineWight > 0 Then
                        br.LineWeight = LineWight
                    Else
                        If LineWight = -2 Then
                            br.LineWeight = LineWeight.ByBlock
                        Else
                            br.LineWeight = LineWeight.ByLayer
                        End If
                    End If
                    'поворот блока
                    br.Rotation = rot
                    If btr.HasAttributeDefinitions Then
                        Dim myAttColl As DatabaseServices.AttributeCollection = br.AttributeCollection
                        For Each entID As ObjectId In btr
                            Dim myEnt As DatabaseServices.Entity = entID.GetObject(OpenMode.ForWrite)
                            If TypeOf myEnt Is DatabaseServices.AttributeDefinition Then 'если это атрибут
                                'делаем на него ссылку
                                Dim myAttDef As DatabaseServices.AttributeDefinition = myEnt
                                'создаем новое определение атрибута
                                Dim myAttRef As New DatabaseServices.AttributeReference
                                'копируем значения в новый атрибут
                                myAttRef.SetAttributeFromBlock(myAttDef, br.BlockTransform)
                                myAttRef.Position = myAttDef.Position.TransformBy(br.BlockTransform)
                                myAttRef.TextString = myAttDef.TextString
                                myAttRef.Tag = myAttDef.Tag
                                myAttColl.AppendAttribute(myAttRef)
                                'проверяем атрибут на наличие поля
                                If myAttDef.HasFields = True Then
                                    Dim id As ObjectId = myAttDef.GetField() 'смотрим старый атрибут
                                    Dim fd As Field = tr.GetObject(id, OpenMode.ForRead) 'читаем его
                                    Dim CodStrFiled As String = fd.GetFieldCode 'получаем код поля
                                    'получаем id блока и отделяем от него скобки
                                    Dim strID As String = br.Id.ToString() 'получаем id вновь созданного блока
                                    strID = Mid(strID, 2, strID.Length - 2) 'убираем из строки первую и последние скобки
                                    Dim PosBlk As Point3d = br.Position 'смотрим обновленную позицию блока
                                    'строка поиска (ОТОБРАЖАТЬ КООРДИНАТЫ)
                                    Dim FindStr1 As String = "InsertionPoint \f "
                                    Dim ind1 As Integer = CodStrFiled.IndexOf(FindStr1)
                                    Dim indk As Integer = CodStrFiled.LastIndexOf(")")
                                    If ind1 <> -1 Then
                                        Dim str1 As String = "%<\AcObjProp Object(%<\_ObjId "
                                        Dim StartPos As Integer = ind1 + 18
                                        Dim SumbCount As Integer = indk - ind1 - 17
                                        Dim str2 As String = ">%).InsertionPoint \f " & Chr(34) & "%lu2%pt4%pr2" & Chr(34) & ">%"
                                        If SumbCount > 0 Then
                                            str2 = ">%).InsertionPoint \f " & Mid(CodStrFiled, ind1 + 18, indk - ind1 - 17) & ">%"
                                        End If
                                        Dim str As String = str1 + strID + str2
                                        Dim fldstr As String = String.Format(str, PosBlk.ToString())
                                        Dim field As Field = New Field(fldstr)
                                        field.Evaluate()
                                        Dim fieldEval As FieldEvaluationStatusResult = field.EvaluationStatus
                                        If fieldEval.Status = FieldEvaluationStatus.Success Then
                                            myAttRef.SetField(field)
                                            tr.AddNewlyCreatedDBObject(field, True)
                                        Else
                                            Dim Otm As Double = Math.Round(pointBlkXY.Z, 2)
                                            myAttDef.TextString = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Otm)
                                        End If
                                    End If
                                    'строка поиска (ДИНАМИЧЕСКИЙ ПАРАМЕТР ЛИНЕЙНЫЙ РАЗМЕР)
                                    Dim FindStr2 As String = ").UpdatedDistance \f"
                                    Dim ind2 As Integer = CodStrFiled.LastIndexOf(FindStr2)
                                    If ind2 <> -1 Then
                                        '1. ищем слово параметр
                                        Dim PosFindStrPar1 As String = CodStrFiled.LastIndexOf("Parameter(")
                                        Dim PosFlag As String = CodStrFiled.IndexOf("AcExpr")
                                        Dim str As String = ""
                                        If PosFlag > -1 Then
                                            Dim PosFindStrPar2 As String = CodStrFiled.LastIndexOf("\f")
                                            Dim str1 As String = "%<\AcObjProp Object(%<\_ObjId "
                                            Dim str2 As String = Mid(CodStrFiled, PosFindStrPar1, PosFindStrPar2 - PosFindStrPar1 - 2)
                                            str = str1 + strID + ">%)" + str2 + """" & ">%"
                                        Else
                                            Dim PosFindStrPar2 As String = CodStrFiled.LastIndexOf("\f")
                                            Dim str1 As String = "%<\AcObjProp Object(%<\_ObjId "
                                            Dim str2 As String = Mid(CodStrFiled, PosFindStrPar1)
                                            str = str1 + strID + ">%)" + str2 + ">%"
                                        End If
                                        Dim fldstr As String = String.Format(str, PosBlk.ToString())
                                        Dim field As Field = New Field(fldstr)
                                        field.Evaluate()
                                        Dim fieldEval As FieldEvaluationStatusResult = field.EvaluationStatus
                                        If fieldEval.Status = FieldEvaluationStatus.Success Then
                                            myAttRef.SetField(field)
                                            tr.AddNewlyCreatedDBObject(field, True)
                                        Else
                                            Dim Otm As Double = Math.Round(pointBlkXY.Z, 2)
                                            myAttDef.TextString = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Otm)
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If

                    FuncInsertDrawingBlk = br
                    tr.Commit()
                End If
            Catch
                FuncInsertDrawingBlk = Nothing
            End Try
        End Using
    End Function
    '===================================================================================
    'вставка в Чертеж блока (в пространство модели) в массиве должен быть ОБЯЗАТЕЛЬНЫЙ параметр, 0-имя блока,1-значение
    Public Shared Function FuncInsBlkByArrayProperties(ByVal pointBlkXY As Point3d, ArrayProperty As String(,), Optional rot As Double = 0) As BlockReference
        FuncInsBlkByArrayProperties = Nothing
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim tr As Transaction = doc.TransactionManager.StartTransaction()
        Dim NameBlock As String = ""
        If IsArray(ArrayProperty) = False Then
            Exit Function
        Else
            For i As Integer = 0 To ArrayProperty.GetUpperBound(1)
                Dim NameField As String = ArrayProperty(0, i)
                If IsNothing(NameField) = True Then Continue For
                Dim value As String = ArrayProperty(1, i)
                If IsNothing(value) = True Then Continue For
                If value.Trim.Length = 0 Then Continue For
                If NameField.Trim Like "Name" Then
                    NameBlock = ArrayProperty(1, i)
                    Exit For
                End If
            Next i
        End If
        If IsNothing(NameBlock) = True Then Exit Function
        If NameBlock.Trim.Length = 0 Then Exit Function
        Using tr
            Try
                Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                If bt.Has(NameBlock) = True Then
                    Dim space As BlockTableRecord = DirectCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                    Dim btr As BlockTableRecord = DirectCast(tr.GetObject(bt(NameBlock), OpenMode.ForRead), BlockTableRecord)
                    Dim br As New BlockReference(pointBlkXY, btr.ObjectId)
                    space.AppendEntity(br)
                    tr.AddNewlyCreatedDBObject(br, True)
                    'присваиваем характеристики блоку
                    For i As Integer = 0 To ArrayProperty.GetUpperBound(1)
                        Dim NameField As String = ArrayProperty(0, i)
                        If IsNothing(NameField) = True Then Continue For
                        Dim value As String = ArrayProperty(1, i)
                        If IsNothing(value) = True Then Continue For
                        If value.Trim.Length = 0 Then Continue For
                        If NameField.Trim Like "Layer" Then
                            Dim NameLayer As String = ArrayProperty(1, i)
                            If IsNothing(NameLayer) = False Then
                                If NameLayer.Trim.Length > 0 Then
                                    Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
                                    If acLyrTbl.Has(NameLayer) = True Then
                                        br.Layer = NameLayer
                                    End If
                                End If
                            End If
                        ElseIf NameField.Trim Like "Color" Then
                            If IsNothing(ArrayProperty(1, i)) = False Then
                                Dim ColorBlk As Integer = Val(ArrayProperty(1, i))
                                If ColorBlk < 0 Then
                                    ColorBlk = 0
                                ElseIf ColorBlk > 256 Then
                                    ColorBlk = 256
                                End If
                                br.ColorIndex = ColorBlk
                            End If
                        ElseIf NameField.Trim Like "Scale" Then
                            If IsNothing(ArrayProperty(1, i)) = False Then
                                Dim ScaleBlk As Integer = Val(ArrayProperty(1, i))
                                If ScaleBlk > 0 Then
                                    Dim anno As AnnotationScale = db.Cannoscale
                                    Dim DrawScale As Double = anno.DrawingUnits
                                    Dim ScaleBlkRez As Double = DrawScale * ScaleBlk
                                    br.ScaleFactors = New Scale3d(ScaleBlkRez, ScaleBlkRez, ScaleBlkRez)
                                End If
                            End If
                        ElseIf NameField.Trim Like "LineType" Then
                            Dim NameTypeLine As String = ArrayProperty(1, i)
                            If IsNothing(NameTypeLine) = False Then
                                If NameTypeLine.Trim.Length > 0 Then
                                    Dim acLineTypTbl As LinetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead)
                                    If acLineTypTbl.Has(NameTypeLine) = True Then
                                        br.Linetype = NameTypeLine
                                    End If
                                End If
                            End If
                        ElseIf NameField.Trim Like "LineWidth" Then
                            If IsNothing(ArrayProperty(1, i)) = False Then
                                Dim WidthBlk As Integer = Val(ArrayProperty(1, i))
                                If WidthBlk > 0 Then
                                    br.LineWeight = WidthBlk
                                End If
                            End If
                        End If
                    Next i
                    br.Rotation = rot
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'Вставляем атрибуты блока еcли они есть
                    If btr.HasAttributeDefinitions Then
                        Dim myAttColl As DatabaseServices.AttributeCollection = br.AttributeCollection
                        For Each entID As ObjectId In btr
                            Dim myEnt As DatabaseServices.Entity = entID.GetObject(OpenMode.ForWrite)
                            If TypeOf myEnt Is DatabaseServices.AttributeDefinition Then 'если это атрибут
                                'делаем на него ссылку
                                Dim myAttDef As DatabaseServices.AttributeDefinition = myEnt
                                'создаем новое определение атрибута
                                Dim myAttRef As New DatabaseServices.AttributeReference
                                'копируем значения в новый атрибут
                                myAttRef.SetAttributeFromBlock(myAttDef, br.BlockTransform)
                                myAttRef.Position = myAttDef.Position.TransformBy(br.BlockTransform)
                                myAttRef.Tag = myAttDef.Tag
                                myAttRef.TextString = myAttDef.TextString
                                For i As Integer = 0 To ArrayProperty.GetUpperBound(1)
                                    Dim NameField As String = ArrayProperty(0, i)
                                    If IsNothing(NameField) = True Then Continue For
                                    Dim value As String = ArrayProperty(1, i)
                                    If IsNothing(value) = True Then Continue For
                                    If NameField.Trim Like myAttRef.Tag Then
                                        myAttRef.TextString = value
                                        myAttRef.Rotation = 0
                                    End If
                                Next
                                myAttColl.AppendAttribute(myAttRef)
                                'проверяем атрибут на наличие поля
                                If myAttDef.HasFields = True Then
                                    Dim id As ObjectId = myAttDef.GetField() 'смотрим старый атрибут
                                    Dim fd As Field = tr.GetObject(id, OpenMode.ForRead) 'читаем его
                                    Dim CodStrFiled As String = fd.GetFieldCode 'получаем код поля
                                    'получаем id блока и отделяем от него скобки
                                    Dim strID As String = br.Id.ToString() 'получаем id вновь созданного блока
                                    strID = Mid(strID, 2, strID.Length - 2) 'убираем из строки первую и последние скобки
                                    Dim PosBlk As Point3d = br.Position 'смотрим обновленную позицию блока
                                    'строка поиска (ОТОБРАЖАТЬ КООРДИНАТЫ)
                                    Dim FindStr1 As String = "InsertionPoint \f "
                                    Dim ind1 As Integer = CodStrFiled.IndexOf(FindStr1)
                                    Dim indk As Integer = CodStrFiled.LastIndexOf(")")
                                    If ind1 <> -1 Then
                                        Dim str1 As String = "%<\AcObjProp Object(%<\_ObjId "
                                        Dim str2 As String = ">%).InsertionPoint \f " & Mid(CodStrFiled, ind1 + 18, indk - ind1 - 17) & ">%"
                                        Dim str As String = str1 + strID + str2
                                        Dim fldstr As String = String.Format(str, PosBlk.ToString())
                                        Dim field As Field = New Field(fldstr)
                                        field.Evaluate()
                                        Dim fieldEval As FieldEvaluationStatusResult = field.EvaluationStatus
                                        If fieldEval.Status = FieldEvaluationStatus.Success Then
                                            myAttRef.SetField(field)
                                            tr.AddNewlyCreatedDBObject(field, True)
                                        Else
                                            Dim Otm As Double = Math.Round(pointBlkXY.Z, 2)
                                            myAttDef.TextString = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Otm)
                                        End If
                                    End If
                                    'строка поиска (ДИНАМИЧЕСКИЙ ПАРАМЕТР ЛИНЕЙНЫЙ РАЗМЕР)
                                    Dim FindStr2 As String = ").UpdatedDistance \f"
                                    Dim ind2 As Integer = CodStrFiled.LastIndexOf(FindStr2)
                                    If ind2 <> -1 Then
                                        '1. ищем слово параметр
                                        Dim PosFindStrPar1 As String = CodStrFiled.LastIndexOf("Parameter(")
                                        Dim PosFlag As String = CodStrFiled.IndexOf("AcExpr")
                                        Dim str As String = ""
                                        If PosFlag > -1 Then
                                            Dim PosFindStrPar2 As String = CodStrFiled.LastIndexOf("\f")
                                            Dim str1 As String = "%<\AcObjProp Object(%<\_ObjId "
                                            Dim str2 As String = Mid(CodStrFiled, PosFindStrPar1, PosFindStrPar2 - PosFindStrPar1 - 2)
                                            str = str1 + strID + ">%)" + str2 + """" & ">%"
                                        Else
                                            Dim PosFindStrPar2 As String = CodStrFiled.LastIndexOf("\f")
                                            Dim str1 As String = "%<\AcObjProp Object(%<\_ObjId "
                                            Dim str2 As String = Mid(CodStrFiled, PosFindStrPar1)
                                            str = str1 + strID + ">%)" + str2 + ">%"
                                        End If
                                        Dim fldstr As String = String.Format(str, PosBlk.ToString())
                                        Dim field As Field = New Field(fldstr)
                                        field.Evaluate()
                                        Dim fieldEval As FieldEvaluationStatusResult = field.EvaluationStatus
                                        If fieldEval.Status = FieldEvaluationStatus.Success Then
                                            myAttRef.SetField(field)
                                            tr.AddNewlyCreatedDBObject(field, True)
                                        Else
                                            Dim Otm As Double = Math.Round(pointBlkXY.Z, 2)
                                            myAttDef.TextString = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Otm)
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                    FuncInsBlkByArrayProperties = br
                    tr.Commit()
                End If
            Catch
                FuncInsBlkByArrayProperties = Nothing
            End Try
        End Using
    End Function
    '===================================================================================
    'Функция считывает в массив все атрибуты блока (Массив 0-тег, 1-содержание атрибута, 2-описание атрибута, 3-высота атрибута, 4-угол наклона атрибута, 5-текстовый стиль, 6-цвет)
    Public Shared Function FuncReadAllBlockAttr(ByVal blkPoint As BlockReference, ByRef ArrayData As String(,)) As Boolean
        Dim Count As Integer = 0
        FuncReadAllBlockAttr = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Erase ArrayData
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = CType(acTrans.GetObject(bt(blkPoint.Name), OpenMode.ForRead), BlockTableRecord)
                Dim attCol As AttributeCollection = blkPoint.AttributeCollection
                For Each attId As ObjectId In attCol
                    Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForRead)
                    ReDim Preserve ArrayData(6, Count)
                    ArrayData(0, Count) = attRef.Tag 'имя поля
                    ArrayData(1, Count) = attRef.TextString 'значение
                    ArrayData(3, Count) = attRef.Height 'высота атрибута
                    ArrayData(4, Count) = attRef.Rotation 'угол атрибута
                    ArrayData(5, Count) = attRef.TextStyleName 'стиль текста
                    ArrayData(6, Count) = attRef.ColorIndex 'цвет
                    If btr.HasAttributeDefinitions Then
                        For Each ObjId As ObjectId In btr
                            Dim obj As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ObjId, OpenMode.ForRead)
                            If TypeOf obj Is AttributeDefinition Then
                                Dim AttDef As AttributeDefinition = obj
                                Dim TAG_string As String = AttDef.Tag
                                Dim PROMPT_String As String = AttDef.Prompt
                                If TAG_string Like attRef.Tag Then
                                    ArrayData(2, Count) = PROMPT_String 'описание поля
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                    Count += 1
                Next
            Catch
                Exit Function
            End Try
            FuncReadAllBlockAttr = True
        End Using
    End Function
    '-===================================================================================
    'Функция считывает определенный атрибут блока и возврящает его значение
    Public Shared Function FuncReadValueAttr(ByVal blkPoint As BlockReference, ByVal NameField As String) As String
        FuncReadValueAttr = ""
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim attCol As AttributeCollection = blkPoint.AttributeCollection
                If attCol.Count > 0 Then
                    For Each attId As ObjectId In attCol
                        Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForRead)
                        If attRef.Tag Like NameField Then
                            FuncReadValueAttr = attRef.TextString
                            Exit Function
                        End If
                    Next
                End If
            Catch
                Exit Function
            End Try
        End Using
    End Function
    '-===================================================================================
    'Функция записывает атрибуты блока (Массив 0-имя атрибута, 1-содержание атрибута, 3-высота атрибута, 4-угол наклона атрибута,5-текстовый стиль,6-цвет)
    Public Shared Function FuncWriteAllBlockAttr(ByVal blkPoint As BlockReference, ByRef ArrayAtribute As String(,)) As Boolean
        FuncWriteAllBlockAttr = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim attCol As AttributeCollection = blkPoint.AttributeCollection
                For Each attId As ObjectId In attCol
                    Try
                        Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForWrite)
                        Dim atributeTag As String = attRef.Tag
                        For i As Integer = 0 To ArrayAtribute.GetUpperBound(1)
                            Dim nameField As String = ArrayAtribute(0, i)
                            If nameField Like atributeTag Then
                                Dim valField As String = ArrayAtribute(1, i)
                                If IsNothing(valField) = False Then
                                    If valField.Trim.Length > 0 Then
                                        attRef.TextString = valField
                                        If attRef.Visible = True Then
                                            Dim heightAttr As Double = Val(ArrayAtribute(3, i))
                                            If heightAttr > 0 Then
                                                attRef.Height = heightAttr
                                            End If
                                            Dim angle As Double = Val(ArrayAtribute(4, i))
                                            attRef.Oblique = angle
                                            Dim textStyleAttr As String = ArrayAtribute(5, i)
                                            If textStyleAttr.Trim.Length > 0 Then
                                                Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead)
                                                If acTextStyleTbl.Has(textStyleAttr) = True Then
                                                    Dim idTextStyle As ObjectId = acTextStyleTbl.Item(textStyleAttr)
                                                    attRef.TextStyleId = idTextStyle
                                                End If
                                            End If
                                            Dim colorInd As Integer = Val(ArrayAtribute(6, i))
                                            If colorInd > -1 And colorInd < 257 Then
                                                attRef.ColorIndex = colorInd
                                            End If
                                        End If
                                    End If
                                End If
                                Exit For
                            End If
                        Next i
                    Catch ex As System.Exception
                    End Try
                Next
            Catch
                Exit Function
            End Try
            acTrans.Commit()
            FuncWriteAllBlockAttr = True
        End Using
    End Function
    '-===================================================================================
    'Функция возвращает имя блока
    Public Shared Function FuncReadNameBlk(ByVal blockRef As BlockReference) As String
        FuncReadNameBlk = ""
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim block As BlockTableRecord = Nothing
                If blockRef.IsDynamicBlock Then
                    block = TryCast(acTrans.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                Else
                    block = TryCast(acTrans.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                End If

                If block IsNot Nothing Then
                    FuncReadNameBlk = block.Name
                End If
            Catch ex As System.Exception
                Exit Function
            End Try
        End Using
    End Function
    '========================================================================================
    Public Shared Function SetRotationParameterSafe(blockRef As BlockReference, ByVal nameParameter As String, ByVal value As String) As Boolean
        Try
            ' Проверяем, является ли блок динамическим
            If blockRef Is Nothing Then
                Return False
            End If

            If Not blockRef.IsDynamicBlock Then
                Return False
            End If

            ' Получаем коллекцию свойств динамического блока
            Dim props As DynamicBlockReferencePropertyCollection = blockRef.DynamicBlockReferencePropertyCollection
            If props Is Nothing OrElse props.Count = 0 Then
                Return False
            End If

            ' Ищем параметр по имени
            For Each prop As DynamicBlockReferenceProperty In props
                If prop IsNot Nothing AndAlso
               String.Equals(prop.PropertyName, nameParameter, StringComparison.OrdinalIgnoreCase) Then

                    Try
                        ' Проверяем, можно ли записывать
                        If Not prop.ReadOnly Then

                            ' Проверяем тип значения
                            Dim currentValue As Object = prop.Value
                            Dim newValue As Object = Nothing

                            ' Определяем тип и преобразуем
                            If currentValue IsNot Nothing Then
                                Dim valueType As Type = currentValue.GetType()

                                If valueType Is GetType(Double) Then
                                    Dim doubleVal As Double
                                    If Double.TryParse(value, doubleVal) Then
                                        newValue = doubleVal
                                    End If
                                ElseIf valueType Is GetType(Integer) Then
                                    Dim intVal As Integer
                                    If Integer.TryParse(value, intVal) Then
                                        newValue = intVal
                                    End If
                                ElseIf valueType Is GetType(String) Then
                                    newValue = value
                                ElseIf valueType Is GetType(Boolean) Then
                                    Dim boolVal As Boolean
                                    If Boolean.TryParse(value, boolVal) Then
                                        newValue = boolVal
                                    End If
                                End If
                            End If

                            If newValue IsNot Nothing Then
                                prop.Value = newValue
                                Return True
                            End If
                        End If
                    Catch innerEx As Exception
                        ' Логируем ошибку для конкретного параметра
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                        vbCrLf & $"Ошибка при установке параметра '{nameParameter}': {innerEx.Message}")
                    End Try

                    Return False
                End If
            Next

            Return False

        Catch ex As Exception
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
            vbCrLf & $"Ошибка при установке параметра: {ex.Message}")
            Return False
        End Try
    End Function
    'функция меняет параметр угол поворота динамического блока
    Public Shared Function SetRotationParameter(blockRef As BlockReference, ByVal nameParameter As String, ByVal value As String) As Boolean
        ' Проверяем, является ли блок динамическим
        SetRotationParameter = False
        If Not blockRef.IsDynamicBlock Then
            Return False
        End If
        ' Получаем коллекцию свойств динамического блока
        Dim props As DynamicBlockReferencePropertyCollection = blockRef.DynamicBlockReferencePropertyCollection
        ' Ищем параметр по имени "Угол поворота"
        For Each prop As DynamicBlockReferenceProperty In props
            If prop.PropertyName = nameParameter Then
                If prop.UnitsType = DynamicBlockReferencePropertyUnitsType.NoUnits Then
                    prop.Value = value
                    Return True
                Else
                    prop.Value = Val(value)
                    Return True
                End If
            End If
        Next
    End Function
    Public Shared Function CopyBlockToPosition(sourceBlock As BlockReference, newPosition As Point3d, tr As Transaction) As BlockReference

        ' Получаем имя блока
        Dim blockName As String = sourceBlock.Name

        ' Получаем определение блока
        Dim bt As BlockTable = tr.GetObject(sourceBlock.Database.BlockTableId, OpenMode.ForRead)
        Dim btrId As ObjectId = bt(blockName)

        ' Создаем новую ссылку
        Dim newBlock As New BlockReference(newPosition, btrId)

        ' Копируем масштаб и поворот
        newBlock.ScaleFactors = sourceBlock.ScaleFactors
        newBlock.Rotation = sourceBlock.Rotation

        Return newBlock
    End Function
    Public Shared Function CopyDynamicBlock(sourceBlock As BlockReference, newPosition As Point3d, tr As Transaction) As BlockReference
        ' Получаем имя блока
        Dim blockName As String = sourceBlock.Name
        ' Получаем определение блока
        Dim bt As BlockTable = tr.GetObject(sourceBlock.Database.BlockTableId, OpenMode.ForRead)
        Dim btrId As ObjectId = bt(blockName)
        ' Создаем новую ссылку на блок
        Dim newBlock As New BlockReference(newPosition, btrId)
        ' Копируем основные свойства
        newBlock.Rotation = sourceBlock.Rotation
        newBlock.ScaleFactors = sourceBlock.ScaleFactors
        newBlock.Layer = sourceBlock.Layer
        newBlock.Color = sourceBlock.Color
        ' Копируем динамические свойства (если есть)
        If sourceBlock.IsDynamicBlock Then
            Dim sourceProps As DynamicBlockReferencePropertyCollection = sourceBlock.DynamicBlockReferencePropertyCollection
            Dim newProps As DynamicBlockReferencePropertyCollection = newBlock.DynamicBlockReferencePropertyCollection

            For Each sourceProp As DynamicBlockReferenceProperty In sourceProps
                For Each newProp As DynamicBlockReferenceProperty In newProps
                    If newProp.PropertyName = sourceProp.PropertyName Then
                        newProp.Value = sourceProp.Value
                        Exit For
                    End If
                Next
            Next
        End If
        Return newBlock
    End Function
    'функция выдергивает из блока полилинии и вставляет их в чертеж
    Public Shared Function ExtractPolylinesFromBlockAndDelete(ByVal blockId As ObjectId) As Polyline()
        Dim result As Polyline() = {}
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim extractedCount As Integer = 0
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' 1. Получаем ссылку на блок (BlockReference)
                Dim blockRef As BlockReference = trans.GetObject(blockId, OpenMode.ForRead)
                ' 2. Получаем определение блока (BlockTableRecord)
                Dim blockDefId As ObjectId = blockRef.BlockTableRecord
                Dim blockDef As BlockTableRecord = trans.GetObject(blockDefId, OpenMode.ForRead)
                ' 3. Создаем список для хранения ObjectId полилиний, которые нужно скопировать
                Dim polylineIds As New List(Of ObjectId)()

                ' 4. Перебираем все элементы в определении блока
                For Each entId As ObjectId In blockDef
                    Dim ent As DatabaseServices.Entity = trans.GetObject(entId, OpenMode.ForRead)
                    ' Проверяем, является ли элемент полилинией
                    If TypeOf ent Is Polyline Then
                        polylineIds.Add(entId)
                    ElseIf TypeOf ent Is Polyline2d Then
                        polylineIds.Add(entId)
                    ElseIf TypeOf ent Is Polyline3d Then
                        polylineIds.Add(entId)
                    End If
                Next
                ' 5. Если полилинии найдены, копируем их в пространство модели
                If polylineIds.Count > 0 Then
                    ' Открываем пространство модели для записи
                    Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim ms As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    ' Получаем матрицу трансформации блока
                    Dim transform As Matrix3d = blockRef.BlockTransform
                    ' Копируем каждую полилинию
                    For Each plId As ObjectId In polylineIds
                        Dim sourcePl As DatabaseServices.Entity = trans.GetObject(plId, OpenMode.ForRead)
                        ' Создаем копию полилинии
                        Dim clonedPl As DatabaseServices.Entity = sourcePl.Clone()
                        ' Применяем трансформацию блока (масштаб, поворот, смещение)
                        clonedPl.TransformBy(transform)
                        ' Добавляем в пространство модели
                        ms.AppendEntity(clonedPl)
                        trans.AddNewlyCreatedDBObject(clonedPl, True)
                        ReDim Preserve result(extractedCount)
                        result(extractedCount) = clonedPl
                        extractedCount += 1
                    Next
                End If
                ' 6. Удаляем исходный блок
                blockRef.UpgradeOpen()
                blockRef.Erase()
                trans.Commit()
                ed.WriteMessage(vbCrLf & $"Извлечено полилиний: {extractedCount}. Блок удален.")
            Catch ex As Exception
                ed.WriteMessage(vbCrLf & $"Ошибка: {ex.Message}")
                trans.Abort()
            End Try
        End Using
        Return result
    End Function

    '===================================================================================
    'Функция меняет размер атрибута блока и его позицию
    Public Shared Function setHeightAndPositionArraibute(ByVal blk As BlockReference, ByVal height As Double, ByVal pos As Point2d) As Boolean
        setHeightAndPositionArraibute = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = CType(acTrans.GetObject(bt(blk.Name), OpenMode.ForWrite), BlockTableRecord)
                Dim attCol As AttributeCollection = blk.AttributeCollection
                For Each attId As ObjectId In attCol
                    Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForWrite)
                    If attRef.Visible = True Then
                        If height > 0 Then
                            attRef.Height = height
                        End If
                        If pos.X <> 0 OrElse pos.Y <> 0 Then
                            Dim tempPos As Point3d = attRef.Position
                            tempPos = New Point3d(pos.X, pos.Y, tempPos.Z)
                            attRef.Position = tempPos
                        End If
                    End If
                Next
            Catch
                Exit Function
            End Try
            acTrans.Commit()
            setHeightAndPositionArraibute = True
        End Using
    End Function
    '===================================================================================
    'Функция меняет размер атрибута блока и его позицию 0-имя атрибута, 1-x, 2-y
    Public Shared Function setPositionAndPositionArraibute(ByVal blk As BlockReference, ByVal properties As String(,)) As Boolean
        setPositionAndPositionArraibute = False
        If IsArray(properties) = True Then
            If properties.GetUpperBound(0) > 1 Then
                If IsNothing(blk) = False Then
                    Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim acCurDb As Database = acDoc.Database
                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                        Try
                            Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                            Dim btr As BlockTableRecord = CType(acTrans.GetObject(bt(blk.Name), OpenMode.ForWrite), BlockTableRecord)
                            Dim attCol As AttributeCollection = blk.AttributeCollection
                            For Each attId As ObjectId In attCol
                                Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForWrite)
                                If attRef.Visible = True Then
                                    For i As Integer = 0 To properties.GetUpperBound(1)
                                        Dim nameAttr As String = properties(0, i)
                                        If IsNothing(nameAttr) = False Then
                                            If nameAttr Like attRef.Tag Then
                                                Dim tempPos As Point3d = attRef.Position
                                                Dim x As Double = blk.Position.X + Val(properties(1, i))
                                                Dim y As Double = blk.Position.Y + Val(properties(2, i))
                                                tempPos = New Point3d(x, y, tempPos.Z)
                                                attRef.Position = tempPos
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                        Catch
                            Exit Function
                        End Try
                        setPositionAndPositionArraibute = True
                        acTrans.Commit()
                    End Using
                End If
            End If
        End If
    End Function



    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '====================================================================================
    'функция создает текстовый стиль
    Public Shared Function FuncCreateTextStyle(ByVal NameTextStyle As String, ByVal NameFont As String, ByVal NameItalic As Boolean) As ObjectId
        FuncCreateTextStyle = Nothing
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Try
            Using tr
                Dim StText As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForWrite)
                If NameTextStyle.Trim.Length > 0 Then
                    If StText.Has(NameTextStyle) = True Then
                        Dim idTextStyle As ObjectId = StText.Item(NameTextStyle)
                        FuncCreateTextStyle = idTextStyle
                    Else
                        Dim TextStyleLine As TextStyleTableRecord = New TextStyleTableRecord()
                        TextStyleLine.Name = NameTextStyle
                        Dim acFont As Autodesk.AutoCAD.GraphicsInterface.FontDescriptor = TextStyleLine.Font
                        Dim acNewFont As Autodesk.AutoCAD.GraphicsInterface.FontDescriptor
                        If NameFont.Trim.Length = 0 Then
                            For Each fontFamily As System.Windows.Media.FontFamily In System.Windows.Media.Fonts.SystemFontFamilies
                                If fontFamily.Source Like "Arial" Then
                                    NameFont = fontFamily.Source
                                    Exit For
                                End If
                                NameFont = fontFamily.Source
                            Next
                        End If
                        If NameItalic = True Then
                            acNewFont = New Autodesk.AutoCAD.GraphicsInterface.FontDescriptor(NameFont, acFont.Bold = True, acFont.Italic = True, acFont.CharacterSet, acFont.PitchAndFamily)
                        Else
                            acNewFont = New Autodesk.AutoCAD.GraphicsInterface.FontDescriptor(NameFont, acFont.Bold = True, acFont.Italic = False, acFont.CharacterSet, acFont.PitchAndFamily)
                        End If
                        TextStyleLine.Font = acNewFont
                        Dim idTextStyle As ObjectId = StText.Add(TextStyleLine)
                        tr.AddNewlyCreatedDBObject(TextStyleLine, True)
                        FuncCreateTextStyle = idTextStyle
                    End If
                Else
                    Dim idTextStyle As ObjectId = GetObjectIdOfTextStyleWithFileName(StText, NameFont)
                    If idTextStyle.IsNull Then
                        Dim tstr As TextStyleTableRecord = New TextStyleTableRecord()
                        tstr.IsShapeFile = True
                        tstr.FileName = NameFont
                        idTextStyle = StText.Add(tstr)
                        tr.AddNewlyCreatedDBObject(tstr, True)
                    End If
                    FuncCreateTextStyle = idTextStyle
                End If
                tr.Commit()
            End Using
        Catch ex As System.Exception
            FuncCreateTextStyle = Nothing
            Exit Function
        End Try
    End Function

    '=========================================================================================
    'функция округляет вершины полилиний (всех 2д и 3д)
    Public Shared Function FuncPlineRoundVertex(ObjEnt As Autodesk.AutoCAD.DatabaseServices.Entity, Optional ByVal RoundZnXY As Integer = 2, Optional ByVal RoundZnH As Integer = 2) As Boolean
        FuncPlineRoundVertex = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            If ObjEnt.Id.IsNull = False Or ObjEnt.Id.IsErased = False Then
                If TypeOf ObjEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim pLine As Autodesk.AutoCAD.DatabaseServices.Polyline = ObjEnt
                    Try
                        For i As Integer = 0 To pLine.NumberOfVertices - 1
                            Dim x As Double = pLine.GetPoint2dAt(i).X
                            Dim y As Double = pLine.GetPoint2dAt(i).Y
                            Dim pt As Point2d = New Point2d(Math.Round(x, RoundZnXY), Math.Round(y, RoundZnXY))
                            pLine.SetPointAt(i, pt)
                        Next i
                        FuncPlineRoundVertex = True
                    Catch ex As Exception
                        FuncPlineRoundVertex = False
                    End Try
                ElseIf TypeOf ObjEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
                    Dim pLine2d As Autodesk.AutoCAD.DatabaseServices.Polyline2d = ObjEnt
                    Try
                        For Each vid As ObjectId In pLine2d
                            Dim v2d As Autodesk.AutoCAD.DatabaseServices.Vertex2d = acTrans.GetObject(vid, OpenMode.ForWrite)
                            v2d.Position = New Point3d(Math.Round(v2d.Position.X, RoundZnXY), Math.Round(v2d.Position.Y, RoundZnXY), 0)
                        Next
                        FuncPlineRoundVertex = True
                    Catch ex As Exception
                        FuncPlineRoundVertex = False
                    End Try
                ElseIf TypeOf ObjEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
                    Dim pLine3d As Autodesk.AutoCAD.DatabaseServices.Polyline3d = ObjEnt
                    Try
                        For Each vid As ObjectId In pLine3d
                            Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = acTrans.GetObject(vid, OpenMode.ForWrite)
                            v3d.Position = New Point3d(Math.Round(v3d.Position.X, RoundZnXY), Math.Round(v3d.Position.Y, RoundZnXY), Math.Round(v3d.Position.Z, RoundZnH))
                        Next
                        FuncPlineRoundVertex = True
                    Catch ex As Exception
                        FuncPlineRoundVertex = False
                    End Try
                End If
            End If
            acTrans.Commit()
        End Using
    End Function
    '=========================================================================================
    'функция удаляет совпадающие вершины (всех 2д и 3д)
    Public Shared Function FuncPlineDelDubleVertex(ByRef ObjEnt As Autodesk.AutoCAD.DatabaseServices.Entity, Optional ByVal DopuskXY As Double = 0, Optional ByVal DopuskH As Double = 0) As Boolean
        FuncPlineDelDubleVertex = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            If ObjEnt.Id.IsNull = False Or ObjEnt.Id.IsErased = False Then
                'обычная полилиния
                If TypeOf ObjEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim pLine As Autodesk.AutoCAD.DatabaseServices.Polyline = ObjEnt
                    Try
line1:
                        'проверяем на совпадение последней и первой вершины
                        Dim StartPt As Autodesk.AutoCAD.Geometry.Point2d = New Point2d(pLine.StartPoint.X, pLine.StartPoint.Y)
                        Dim EndPt As Autodesk.AutoCAD.Geometry.Point2d = pLine.GetPoint2dAt(pLine.NumberOfVertices - 1)
                        Dim dist As Double = StartPt.GetVectorTo(EndPt).Length
                        If dist <= DopuskXY And pLine.NumberOfVertices > 1 Then
                            pLine.RemoveVertexAt(pLine.NumberOfVertices - 1)
                            pLine.Closed = True
                            GoTo line1
                        End If
                        'делаем проверку для остальных вершин
                        For i As Integer = 0 To pLine.NumberOfVertices - 2
                            Dim seg As Autodesk.AutoCAD.Geometry.LineSegment2d = pLine.GetLineSegment2dAt(i)
                            If seg.Length <= DopuskXY Then
                                pLine.RemoveVertexAt(i + 1)
                                GoTo line1
                            End If
                        Next i
                        FuncPlineDelDubleVertex = True
                    Catch ex As Exception
                        FuncPlineDelDubleVertex = False
                    End Try
                    '2d полилиния
                ElseIf TypeOf ObjEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
                    Dim pLine2d As Autodesk.AutoCAD.DatabaseServices.Polyline2d = ObjEnt
                    Dim v2dPrew As Autodesk.AutoCAD.DatabaseServices.Vertex2d = Nothing
                    Dim countVert As Integer = 0
                    Try
line2:
                        For Each vid As ObjectId In pLine2d
                            Dim v2d As Autodesk.AutoCAD.DatabaseServices.Vertex2d = acTrans.GetObject(vid, OpenMode.ForWrite)
                            'если это первая вершина то берем ее значение и продолжаем цикл
                            If countVert = 0 Then
                                v2dPrew = v2d
                                countVert += 1
                                Continue For
                            End If
                            'сравниваем значение с предудущей вершиной
                            Dim seg As Double = v2d.Position.GetVectorTo(v2dPrew.Position).Length
                            If seg <= DopuskXY Then
                                v2d.Erase(True)
                            Else
                                v2dPrew = v2d
                                countVert += 1
                            End If
                        Next
                        'делаем проверку последней вершины
                        If v2dPrew.Position.X = pLine2d.StartPoint.X And v2dPrew.Position.Y = pLine2d.StartPoint.Y Then
                            v2dPrew.Erase(True)
                            pLine2d.Closed = True
                        End If
                        FuncPlineDelDubleVertex = True
                    Catch ex As Exception
                        FuncPlineDelDubleVertex = False
                    End Try
                ElseIf TypeOf ObjEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
                    Dim pLine3d As Autodesk.AutoCAD.DatabaseServices.Polyline3d = ObjEnt
                    Dim v3dPrew As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = Nothing
                    Dim countVert As Integer = 0
                    Try
line3:
                        For Each vid As ObjectId In pLine3d
                            Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = acTrans.GetObject(vid, OpenMode.ForWrite)
                            If countVert = 0 Then
                                v3dPrew = v3d
                                countVert += 1
                                Continue For
                            End If
                            Dim pt1 As Autodesk.AutoCAD.Geometry.Point2d = New Autodesk.AutoCAD.Geometry.Point2d(v3d.Position.X, v3d.Position.Y)
                            Dim pt2 As Autodesk.AutoCAD.Geometry.Point2d = New Autodesk.AutoCAD.Geometry.Point2d(v3dPrew.Position.X, v3dPrew.Position.Y)
                            Dim seg2d As Double = pt1.GetVectorTo(pt2).Length
                            Dim seg3d As Double = Math.Abs(v3d.Position.Z - v3dPrew.Position.Z)
                            If seg2d <= DopuskXY And seg3d <= seg3d Then
                                v3d.Erase()
                                countVert = 0
                                GoTo line3
                            Else
                                v3dPrew = v3d
                                countVert += 1
                            End If
                        Next
                        If v3dPrew.Position.X = pLine3d.StartPoint.X And v3dPrew.Position.Y = pLine3d.StartPoint.Y And v3dPrew.Position.Z = pLine3d.StartPoint.Z Then
                            v3dPrew.Erase(True)
                            pLine3d.Closed = True
                        End If
                        FuncPlineDelDubleVertex = True
                    Catch ex As Exception
                        FuncPlineDelDubleVertex = False
                    End Try
                End If
            End If
            acTrans.Commit()
        End Using
    End Function

    '========================================================================================
    'функция делает реверс полилинии
    Public Shared Function FuncReversPline(ByRef objEnt As Autodesk.AutoCAD.DatabaseServices.Entity) As Boolean
        FuncReversPline = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            If objEnt.Id.IsNull = False Or objEnt.Id.IsErased = False Then
                If TypeOf objEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim pl As Autodesk.AutoCAD.DatabaseServices.Polyline = objEnt
                    Try
                        Dim vertData As List(Of PerVertexData) = New List(Of PerVertexData)(pl.NumberOfVertices)
                        For i As Integer = 0 To pl.NumberOfVertices - 1
                            Dim pvd As PerVertexData = New PerVertexData()
                            pvd.bulge = (If(i > 0, pl.GetBulgeAt(i - 1), 0))
                            pvd.startWidth = (If(i > 0, pl.GetStartWidthAt(i - 1), 0))
                            pvd.endWidth = (If(i > 0, pl.GetEndWidthAt(i - 1), 0))
                            pvd.pt = pl.GetPoint2dAt(i)
                            vertData.Add(pvd)
                        Next
                        For i As Integer = 0 To pl.NumberOfVertices - 1
                            Dim pvd As PerVertexData = vertData(pl.NumberOfVertices - (i + 1))
                            pl.SetPointAt(i, pvd.pt)
                            pl.SetBulgeAt(i, -pvd.bulge)
                            pl.SetStartWidthAt(i, pvd.endWidth)
                            pl.SetEndWidthAt(i, pvd.startWidth)
                        Next
                    Catch ex As Exception
                    End Try
                End If
            End If
            acTrans.Commit()
        End Using
        FuncReversPline = True
    End Function

    '========================================================================================
    'функция делает начальной северозападный угол полилинии
    Public Shared Function FuncSVAnglePline(ByRef objEnt As Autodesk.AutoCAD.DatabaseServices.Entity) As Boolean
        FuncSVAnglePline = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            If objEnt.Id.IsNull = False Or objEnt.Id.IsErased = False Then
                If TypeOf objEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim pl As Autodesk.AutoCAD.DatabaseServices.Polyline = objEnt
                    Try
                        'выбираем габаритный контейнер для полилинии
                        Dim minCoord As Autodesk.AutoCAD.Geometry.Point3d = pl.GeometricExtents.MinPoint
                        Dim maxCoord As Autodesk.AutoCAD.Geometry.Point3d = pl.GeometricExtents.MaxPoint
                        Try
                            Dim vertData As List(Of PerVertexData) = New List(Of PerVertexData)(pl.NumberOfVertices)
                            Dim vertDataNew As List(Of PerVertexData) = New List(Of PerVertexData)(pl.NumberOfVertices)
                            Dim vSV2d As Point2d = New Point2d(minCoord.X, maxCoord.Y)
                            Dim numbVert As Integer = 0
                            Dim MinDist As Double = 1000000000
                            For i As Integer = 0 To pl.NumberOfVertices - 1
                                Dim pvd As PerVertexData = New PerVertexData()
                                pvd.bulge = (If(i > 0, pl.GetBulgeAt(i - 1), 0))
                                pvd.startWidth = (If(i > 0, pl.GetStartWidthAt(i - 1), 0))
                                pvd.endWidth = (If(i > 0, pl.GetEndWidthAt(i - 1), 0))
                                pvd.pt = pl.GetPoint2dAt(i)
                                Dim vdist As Double = pl.GetPoint2dAt(i).GetVectorTo(vSV2d).Length
                                If vdist < MinDist Then
                                    numbVert = i
                                    MinDist = vdist
                                End If
                                vertData.Add(pvd)
                            Next
                            If numbVert > 0 Then
                                '======================================================================
                                ' делаем сортировку вершин
                                Dim nVert As Integer = 0
                                For Each temppvd As PerVertexData In vertData
                                    If nVert >= numbVert Then
                                        vertDataNew.Add(temppvd)
                                        nVert += 1
                                    Else
                                        nVert += 1
                                    End If
                                Next
                                nVert = 0
                                For Each temppvd As PerVertexData In vertData
                                    If nVert < numbVert Then
                                        vertDataNew.Add(temppvd)
                                        nVert += 1
                                    End If
                                Next
                                '=====================================================================
                                'пересобираем полилинию
                                For i As Integer = 0 To pl.NumberOfVertices - 1
                                    Dim pvd As PerVertexData = vertDataNew(i)
                                    pl.SetPointAt(i, pvd.pt)
                                    pl.SetBulgeAt(i, -pvd.bulge)
                                    pl.SetStartWidthAt(i, pvd.endWidth)
                                    pl.SetEndWidthAt(i, pvd.startWidth)
                                Next
                            End If
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                    End Try
                End If
            End If
            acTrans.Commit()
        End Using
        FuncSVAnglePline = True
    End Function

    '===============================================================================================
    'функция заменяет дуговые сегменты 2д полилинии на линейные
    Public Shared Function FuncReplaseArcToLineSegmentPline(ByRef objEnt As Autodesk.AutoCAD.DatabaseServices.Entity) As Boolean
        FuncReplaseArcToLineSegmentPline = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            If objEnt.Id.IsNull = False Or objEnt.Id.IsErased = False Then
                If TypeOf objEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
                    Dim pl As Autodesk.AutoCAD.DatabaseServices.Polyline2d = objEnt
                End If
            End If
            acTrans.Commit()
        End Using
        FuncReplaseArcToLineSegmentPline = True
    End Function

    '===============================================================================================
    'функция вставляет вершину в полилинию
    Public Shared Function FuncInsertVertexToPolyline(ByRef obj As Autodesk.AutoCAD.DatabaseServices.Entity, ByVal pt3d As Point3d) As Boolean
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim chk As Integer = 1
        Try
            Dim tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Using tr
                If obj.IsWriteEnabled = False Then obj.UpgradeOpen()
                Dim lws As Polyline = TryCast(obj, Polyline)
                If lws IsNot Nothing Then
                    Dim pnt1 As Point3d = lws.GetClosestPointTo(pt3d, False)
                    Dim distPoint As Double = lws.GetDistAtPoint(pnt1)
                    For i As Integer = 0 To lws.NumberOfVertices - 2
                        Dim pt1 As Point3d = lws.GetPoint3dAt(i)
                        Dim pt2 As Point3d = lws.GetPoint3dAt(i + 1)
                        Dim dist1 As Double = lws.GetDistAtPoint(pt1)
                        Dim dist2 As Double = lws.GetDistAtPoint(pt2)
                        If dist1 = dist2 Then Continue For

                        If dist1 < distPoint And dist2 > distPoint Then
                            lws.AddVertexAt(i + 1, pt3d.Convert2d(New Plane), 0, 0, 0)
                            Exit For
                        End If
                    Next i
                End If
                tr.Commit()
            End Using
        Catch ex As System.Exception
        End Try
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '3d моделирование
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '=============================================================================================
    'функция создает цилиндр (Точка вставки-центр тела)
    Public Shared Function FuncCreateCylindr3d(ByVal height As Double, ByVal radiusTop As Double, ByVal radiusDown As Double, ByVal centrePoint As Point3d) As Solid3d
        FuncCreateCylindr3d = Nothing
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim acSol3D As Solid3d = New Solid3d()
                acSol3D.SetDatabaseDefaults()
                acSol3D.RecordHistory = True
                acSol3D.CreateFrustum(height, radiusDown, radiusDown, radiusTop)
                acSol3D.TransformBy(Autodesk.AutoCAD.Geometry.Matrix3d.Displacement(centrePoint - Point3d.Origin))
                acBlkTblRec.AppendEntity(acSol3D)
                acTrans.AddNewlyCreatedDBObject(acSol3D, True)
                acTrans.Commit()
                FuncCreateCylindr3d = acSol3D
            End Using
        Catch ex As System.Exception
        End Try
    End Function



    '=========================================================================================================
    'Отчет по образуемым земельным участкам
    <CommandMethod("TestAddObject")>
    Public Sub ListGovementOwnerParcels()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim Putch As String = "k:\Civil3DToolsUtility\FileResources\3d_Object\MVBlocks.dwg"
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim boolinsObject As Boolean = FuncCopyMVBlockDestDraw(Putch, "MBV-Растительность луговая", False)
            tr.Commit()
        End Using

    End Sub

End Class
