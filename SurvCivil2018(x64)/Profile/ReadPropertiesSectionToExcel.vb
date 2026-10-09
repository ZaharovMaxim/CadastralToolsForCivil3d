Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class ReadPropertiesSectionToExcel
    '=============================================================================================
    'Сформировать отчет характеристик поперечного профиля на основе меток Civil3d
    <CommandMethod("ReadPropertiesCivilSectionToExcel")>
    Public Sub ReadPropertiesCivilSectionToExcel()
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim docBase As Database = docCol.MdiActiveDocument.Database
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
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
        Dim ArrayExcel As String(,) = Nothing
        Dim CountArrayExcel As Integer = 0
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim resBlk As PromptEntityResult = Nothing
            Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите трассу: ")
            resBlk = m_Editor.GetEntity(opt)
            Dim pvId As ObjectId = resBlk.ObjectId
            Dim obj As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(pvId, OpenMode.ForRead)
            If Not (TypeOf obj Is Alignment) Then
                MsgBox("Необходимо выбрать объект типа ТРАССА. Макро прервано!!!")
                Exit Sub
            End If
            Dim align As Alignment = obj
            'находим все самплы
            Dim CollSamplLine As ObjectIdCollection = New ObjectIdCollection
            CollSamplLine = align.GetSampleLineGroupIds()
            If CollSamplLine.Count = 0 Then
                MsgBox("У выбранной трассы нет групп сечений. Макро прервано!!!")
                Exit Sub
            End If
            '==================================================================================================
            'указать начало и конец трассы
            Dim StartPK As Double = align.StartingStation
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            pPtOpts.AllowNone = True
            pPtOpts.Message = vbLf & "Укажите точку начала трассы "
            Dim pPtRes As PromptPointResult = m_Editor.GetPoint(pPtOpts)
            Dim ptStart As Point3d = pPtRes.Value 'координаты точки начала трассы
            If pPtRes.Status = PromptStatus.None Or pPtRes.Status = PromptStatus.Cancel Then 'нет выбора
                StartPK = align.StartingStation
            Else
                Try
                    Dim OfsetVedStartAlign As Double = 0
                    align.StationOffset(ptStart.X, ptStart.Y, StartPK, OfsetVedStartAlign)
                    Dim tempPtX As Double = -1
                    Dim tempPtY As Double = -1
                    align.PointLocation(StartPK, 0, tempPtX, tempPtY)
                    ptStart = New Point3d(tempPtX, tempPtY, 0)
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                    MsgBox("Не верно указана точка начала!!!")
                    locker.Dispose()
                    Exit Sub
                Catch ex As ArgumentException
                    MsgBox("Не верно указана точка начала!!!")
                    locker.Dispose()
                    Exit Sub
                End Try
            End If
            'запрашиваем точку конца трассы
            Dim EndPK As Double = align.EndingStation
            pPtOpts = New PromptPointOptions("")
            pPtOpts.AllowNone = True
            pPtOpts.Message = vbLf & "Укажите точку конца трассы "
            pPtRes = m_Editor.GetPoint(pPtOpts)
            Dim ptEnd As Point3d = pPtRes.Value 'координаты точки конца трассы
            If pPtRes.Status = PromptStatus.None Or pPtRes.Status = PromptStatus.Cancel Then
                EndPK = align.EndingStation
            Else
                Try
                    Dim OfsetVedEndAlign As Double = 0
                    align.StationOffset(ptEnd.X, ptEnd.Y, EndPK, OfsetVedEndAlign)
                    Dim tempPtX As Double = -1
                    Dim tempPtY As Double = -1
                    align.PointLocation(EndPK, 0, tempPtX, tempPtY)
                    ptEnd = New Point3d(tempPtX, tempPtY, 0)
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                    MsgBox("Не верно указана точка начала!!!")
                    Exit Sub
                Catch ex As ArgumentException
                    MsgBox("Не верно указана точка начала!!!")
                    Exit Sub
                End Try
            End If
            '======================================================================================
            'фильтруем самплы
            Dim CollLine As ObjectIdCollection = New ObjectIdCollection
            For Each ids As ObjectId In CollSamplLine
                Dim SLineG As SampleLineGroup = tr.GetObject(ids, OpenMode.ForRead)
                Dim TempSlineColl As ObjectIdCollection = New ObjectIdCollection
                TempSlineColl = SLineG.GetSampleLineIds
                If TempSlineColl.Count > 0 Then
                    For Each idtemp As ObjectId In TempSlineColl
                        CollLine.Add(idtemp)
                    Next
                End If
            Next
            For Each ids As ObjectId In CollLine
                Dim SLine As SampleLine = tr.GetObject(ids, OpenMode.ForRead)
                Dim PK As Double = SLine.Station
                If PK < StartPK Then
                    Continue For
                End If
                If PK > EndPK Then
                    Continue For
                End If
                'находим вид профиля
                Dim CollVPro As ObjectIdCollection = New ObjectIdCollection
                CollVPro = SLine.GetSectionViewIds()
                If CollVPro.Count > 0 Then
                    For Each idVProf As ObjectId In CollVPro
                        Dim VProf As SectionView = tr.GetObject(idVProf, OpenMode.ForRead)
                        Dim TV As SectionViewVolumeTableGroup = VProf.VolumeTables
                        Dim id As ObjectIdCollection = TV.GetVolumeTableIds
                        For Each id2 As ObjectId In id
                            Dim s As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(id2, OpenMode.ForRead)
                            Dim qw As Autodesk.Civil.DatabaseServices.SectionViewQuantityTakeoffTable = s
                            Dim db1 As DBObjectCollection = New DBObjectCollection
                            qw.Explode(db1)
                            If db1.Count > 0 Then
                                For Each dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject In db1
                                    If TypeOf dbObj Is BlockReference Then
                                        Dim blk As BlockReference = dbObj
                                        Dim db2 As DBObjectCollection = New DBObjectCollection
                                        blk.Explode(db2)
                                        If db2.Count > 0 Then
                                            For Each dbobj1 As Autodesk.AutoCAD.DatabaseServices.DBObject In db2
                                                Dim db3 As DBObjectCollection = New DBObjectCollection
                                                If TypeOf dbobj1 Is MText Then
                                                    Dim Mtxt As MText = dbobj1
                                                    Mtxt.Explode(db3)
                                                    If db3.Count > 0 Then
                                                        For Each dbobj2 As Autodesk.AutoCAD.DatabaseServices.DBObject In db3
                                                            If TypeOf dbobj2 Is DBText Then
                                                                Dim txt As DBText = dbobj2
                                                                Dim TextVal As String = txt.TextString
                                                                TextVal = TextVal.Replace("U+", "u")
                                                                TextVal = FuncFormatZN.FuncUnicodDecode(TextVal)
                                                                Dim Pref As String = ""
                                                                Dim val As String = ""
                                                                Dim pos As Integer = TextVal.IndexOf("=")
                                                                If pos > -1 Then
                                                                    Pref = Mid(TextVal, 1, pos - 1).Trim
                                                                    TextVal = Mid(TextVal, pos + 1)
                                                                End If
                                                                'вычленяем значение
                                                                Dim Group1 As String() = (From M As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(TextVal, "[0-9]+(\.[0-9]+)?") Select M.Groups(0).Value).ToArray()
                                                                If IsArray(Group1) = True Then
                                                                    val = Group1(Group1.Length - 1)
                                                                End If
                                                                If pos = -1 Then
                                                                    Dim pos2 As Integer = TextVal.IndexOf(val)
                                                                    If pos2 > -1 Then
                                                                        Pref = Mid(TextVal, 1, pos2 - 1)
                                                                    End If
                                                                End If
                                                                If Pref.Trim.Length > 0 And val.Trim.Length > 0 Then
                                                                    'заносим данные в массив
                                                                    '=======================================================================================
                                                                    'заносим сампл в массив 0-RawПК, 1-ПК, 2-Имя поля, 3-значение
                                                                    ReDim Preserve ArrayExcel(3, CountArrayExcel)
                                                                    ArrayExcel(0, CountArrayExcel) = PK
                                                                    Try
                                                                        ArrayExcel(1, CountArrayExcel) = align.GetStationStringWithEquations(PK)
                                                                    Catch ex As System.ArgumentException
                                                                    End Try
                                                                    ArrayExcel(2, CountArrayExcel) = Pref
                                                                    ArrayExcel(3, CountArrayExcel) = val
                                                                    CountArrayExcel += 1
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    Next
                End If
            Next
            locker.Dispose()
        End Using
        ' запускаем excel
        If IsArray(ArrayExcel) = True Then
            If ArrayExcel.GetUpperBound(1) > 2 Then
                Dim boolsort As Boolean = MathFunction.FuncSortStrArray(ArrayExcel, 0)
            End If
            Dim xlApp As Excel.Application = Nothing
            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Worksheet = Nothing
            xlApp = CreateObject("Excel.Application")
            xlBook = xlApp.Workbooks.Add()
            xlApp.Visible = True
            Dim xlsheet As Excel.Worksheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
            xlsheet.Activate()
            If IsArray(ArrayExcel) = True Then
                Dim CountRow As Integer = 1
                Dim PrefPK As Double = -0.0912
                Dim ArrayField As String() = Nothing
                Dim CountArrayField As Integer = 0
                For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                    Dim PK As Double = ArrayExcel(0, i)
                    Dim FormatPK As String = ArrayExcel(1, i)
                    Dim FieldStr As String = ArrayExcel(2, i)
                    Dim val As Double = CDbl(ArrayExcel(3, i))
                    If i = 0 Then
                        xlsheet.Cells(CountRow, 2).value = "ПК+..."
                    End If
                    If PK <> PrefPK Then
                        PrefPK = PK
                        CountRow += 2
                        xlsheet.Cells(CountRow, 1).value = PK
                        xlsheet.Cells(CountRow, 2).value = FormatPK
                    End If
                    If IsArray(ArrayField) = False Then
                        ReDim Preserve ArrayField(CountArrayField)
                        ArrayField(CountArrayField) = FieldStr
                        CountArrayField += 1
                        xlsheet.Cells(1, 3).value = "'" & FieldStr
                        xlsheet.Cells(CountRow, 3).value = val
                    Else
                        Dim boolFindField As Boolean = False
                        For j As Integer = 0 To ArrayField.Length - 1
                            Dim TempNameField As String = ArrayField(j)
                            If TempNameField Like FieldStr Then
                                xlsheet.Cells(CountRow, j + 3).value = val
                                boolFindField = True
                                Exit For
                            End If
                        Next j
                        If boolFindField = False Then
                            ReDim Preserve ArrayField(CountArrayField)
                            ArrayField(CountArrayField) = FieldStr
                            CountArrayField += 1
                            xlsheet.Cells(1, CountArrayField + 2).value = "'" & FieldStr
                            xlsheet.Cells(CountRow, CountArrayField + 2).value = val
                        End If
                    End If
                Next i
                xlSheets = Nothing
                xlBook = Nothing
                xlApp.Quit()
                xlApp = Nothing
            End If
        End If
    End Sub

    '============================================================================================================
    'Сформировать отчет характеристик поперечного профиля на основе многострочных текстов
    <CommandMethod("ReadPropertiesAcadSectionToExcel")>
    Public Sub ReadPropertiesAcadSectionToExcel()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ссылка на чертеж .dwg
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim Align As Autodesk.Civil.DatabaseServices.Alignment = Nothing
        Dim Lengh As Double = 70
        Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Ведите максимальное расстояние от значения пикета до характеристик поперечного профиля: ")
        pStrOpts.DefaultValue = Lengh
        Dim pStrRes As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts)
        If pStrRes.Status = PromptStatus.OK Then
            Lengh = pStrRes.Value
        Else
            Exit Sub
        End If
        '==========================================================================================================================
        Dim resBlk As PromptEntityResult = Nothing
        Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите текст (образец) со значением пикета на поперечнике: ")
        resBlk = ed.GetEntity(opt)
        Dim pvId As ObjectId = Nothing
        If Not (resBlk.Status = PromptStatus.OK) Then
            Exit Sub
        End If
        pvId = resBlk.ObjectId
        '==========================================================================================================================
        Dim resBlk1 As PromptEntityResult = Nothing
        Dim opt1 As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите текст (образец) со значениеми характеристик поперечника: ")
        resBlk1 = ed.GetEntity(opt1)
        Dim pvId1 As ObjectId = Nothing
        If Not (resBlk1.Status = PromptStatus.OK) Then
            Exit Sub
        End If
        pvId1 = resBlk1.ObjectId
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim obj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(pvId, OpenMode.ForRead)
            If Not (TypeOf obj Is DBText Or TypeOf obj Is MText) Then
                MsgBox("Необходимо выбрать Текст или МТекст. Макро прервано!!!")
                GoTo line1
            End If
            Dim obj1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(pvId1, OpenMode.ForRead)
            If Not (TypeOf obj1 Is DBText Or TypeOf obj1 Is MText) Then
                MsgBox("Необходимо выбрать Текст или МТекст. Макро прервано!!!")
                GoTo line1
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'фильтруем тексты со всего чертежа для значений пикетов
            Dim acSetBlk As SelectionSet = Nothing
            Dim ValType1(1) As TypedValue
            If TypeOf obj Is DBText Then
                ValType1.SetValue(New TypedValue(DxfCode.Start, "Text"), 0)
            Else
                ValType1.SetValue(New TypedValue(DxfCode.Start, "MText"), 0)
            End If
            ValType1.SetValue(New TypedValue(DxfCode.LayerName, obj.Layer), 1)
            Dim FilterParcel As SelectionFilter = New SelectionFilter(ValType1)
            Dim BlkSelect As PromptSelectionResult = ed.SelectAll(FilterParcel)
            If BlkSelect.Status = PromptStatus.OK Then
                acSetBlk = BlkSelect.Value
            Else
                MsgBox("Фильтр не смог выбрать ни одного значения ПК. Макро прервано!!!")
                Exit Sub
            End If

            Dim ArrayWord As String(,) = Nothing
            Dim CountArrayWord As Integer = 0
            Dim CountField As Integer = 200
            ReDim Preserve ArrayWord(CountField, CountArrayWord)
            ArrayWord(0, 0) = "-999999999"
            ArrayWord(1, 0) = "ПК+"
            ArrayWord(2, 0) = "X"
            ArrayWord(3, 0) = "Y"
            Dim ColumnPos As Integer = 0
            CountArrayWord += 1
            'заносим в массив пикеты 0-сырой ПК, 1- ПК, 2-X, 3=y
            For Each txtId As ObjectId In acSetBlk.GetObjectIds
                Dim acent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(txtId, OpenMode.ForWrite, False)
                Dim TextPK As String = ""
                Dim x As Double = -1
                Dim y As Double = -1
                If TypeOf acent Is MText Then
                    Dim mtext As MText = acent
                    Dim expObj As DBObjectCollection = New DBObjectCollection
                    mtext.Explode(expObj)
                    If expObj.Count > 0 Then
                        For Each dbobj2 As Autodesk.AutoCAD.DatabaseServices.DBObject In expObj
                            If TypeOf dbobj2 Is DBText Then
                                Dim TempTextPKobj As DBText = dbobj2
                                Dim TextVal As String = TempTextPKobj.TextString
                                If TextVal.IndexOf("+") > -1 Then
                                    TextPK = TextVal
                                    x = mtext.Location.X
                                    y = mtext.Location.Y
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                ElseIf TypeOf acent Is DBText Then
                    Dim TempTextPKobj As DBText = acent
                    If TempTextPKobj.TextString.IndexOf("+") > -1 Then
                        TextPK = TempTextPKobj.TextString
                        x = TempTextPKobj.Position.X
                        y = TempTextPKobj.Position.Y
                    End If
                End If
                'делаем разбор строки
                If TextPK.Trim.Length > 0 Then
                    If TextPK.IndexOf("ПК") > -1 Then
                        Dim TextConv As String = TextPK.Replace("ПК", "")
                        ReDim Preserve ArrayWord(CountField, CountArrayWord)
                        TextConv = TextConv.Replace("{\C256;", "")
                        TextConv = TextConv.Replace("}", "")
                        ArrayWord(1, CountArrayWord) = TextConv
                        TextConv = TextConv.Replace("+", "")
                        TextConv = TextConv.Replace(" ", "")
                        If IsNumeric(TextConv) = True Then
                            ArrayWord(0, CountArrayWord) = TextConv
                        Else
                            ArrayWord(0, CountArrayWord) = 0
                        End If
                        ArrayWord(2, CountArrayWord) = x
                        ArrayWord(3, CountArrayWord) = y
                        CountArrayWord += 1
                    End If
                End If
            Next
            'если массив не пустой сортируем массив
            If IsArray(ArrayWord) = True Then
                If ArrayWord.GetUpperBound(1) > 0 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayWord, 0)
                End If
            Else
                Exit Sub
            End If
            '///////////////////////////////////////////////////////////////////////////////
            'заносим в массив тексты с характеристиками
            Dim ArrayText As String(,) = Nothing
            Dim CountArrayText As Integer = 0
            Dim acSetBlk2 As SelectionSet = Nothing
            Dim ValType2(1) As TypedValue
            If TypeOf obj1 Is DBText Then
                ValType2.SetValue(New TypedValue(DxfCode.Start, "Text"), 0)
            Else
                ValType2.SetValue(New TypedValue(DxfCode.Start, "MText"), 0)
            End If
            ValType2.SetValue(New TypedValue(DxfCode.LayerName, obj1.Layer), 1)
            Dim FilterParcel2 As SelectionFilter = New SelectionFilter(ValType2)
            Dim BlkSelect2 As PromptSelectionResult = ed.SelectAll(FilterParcel2)
            If BlkSelect2.Status = PromptStatus.OK Then
                acSetBlk2 = BlkSelect2.Value
            Else
                MsgBox("Фильтр не смог выбрать значения поперечныхпрофилей. Макро прервано!!!")
                Exit Sub
            End If

            For Each txtId As ObjectId In acSetBlk2.GetObjectIds
                Dim acent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(txtId, OpenMode.ForWrite, False)
                Dim TextHarobj As String = ""
                Dim x As Double = -1
                Dim y As Double = -1
                If TypeOf acent Is MText Then
                    Dim mtext1 As MText = acent
                    Dim expObj As DBObjectCollection = New DBObjectCollection
                    mtext1.Explode(expObj)
                    If expObj.Count > 0 Then
                        For Each dbobj2 As Autodesk.AutoCAD.DatabaseServices.DBObject In expObj
                            If TypeOf dbobj2 Is DBText Then
                                Dim TempTextPKobj As DBText = dbobj2
                                Dim TextVal As String = TempTextPKobj.TextString
                                If TextVal.IndexOf("=") > -1 Then
                                    TextHarobj = TextVal
                                    x = TempTextPKobj.Position.X
                                    y = TempTextPKobj.Position.Y
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                ElseIf TypeOf acent Is DBText Then
                    Dim TextHar As DBText = acent
                    If TextHar.TextString.IndexOf("=") > -1 Then
                        TextHarobj = TextHar.TextString
                        x = TextHar.Position.X
                        y = TextHar.Position.Y
                    End If
                End If

                If TextHarobj.Trim.Length > 0 Then
                    ReDim Preserve ArrayText(4, CountArrayText)
                    ArrayText(0, CountArrayText) = TextHarobj
                    ArrayText(1, CountArrayText) = x
                    ArrayText(2, CountArrayText) = y
                    ArrayText(3, CountArrayText) = 0
                    CountArrayText += 1
                End If
            Next
            'находим тексты которые принадлежат ПК
            If IsArray(ArrayWord) = True And IsArray(ArrayText) = True Then
                For i As Integer = 1 To ArrayWord.GetUpperBound(1)
                    'точка ПК
                    Dim Pt1 As Point2d = New Point2d(Val(ArrayWord(2, i)), Val(ArrayWord(3, i)))
                    For j As Integer = 0 To ArrayText.GetUpperBound(1)
                        'точка текста
                        If ArrayText(3, j) = 0 Then
                            Dim Pt2 As Point2d = New Point2d(ArrayText(1, j), ArrayText(2, j))
                            If Pt1.GetVectorTo(Pt2).Length <= Lengh Then
                                Dim TextCont As String = ArrayText(0, j)
                                Dim Pos As Integer = TextCont.IndexOf("=")
                                If Pos > -1 Then
                                    Dim pref As String = Mid(TextCont, 1, Pos)
                                    pref = pref.Replace("{\C256;", "").Trim
                                    Dim post As String = Mid(TextCont, Pos + 2)
                                    Dim FinalPost As String = ""
                                    Dim TempArray As String() = post.Split
                                    If IsArray(TempArray) = True Then
                                        For i1 As Integer = 0 To TempArray.Length - 1
                                            Dim Tempst As String = TempArray(i1)
                                            If IsNumeric(Tempst) = True Then
                                                FinalPost = Tempst
                                                Exit For
                                            End If
                                        Next i1
                                    End If
                                    If FinalPost = "" Then
                                        post = post.Replace("кв.м", "").Trim
                                        post = post.Replace("}", "").Trim
                                        FinalPost = post
                                    End If
                                    Dim IndPole As Integer = 3
                                    Dim boolfindPole As Boolean = False
                                    For k As Integer = 0 To ArrayWord.GetUpperBound(0)
                                        Dim pole As String = ArrayWord(k, 0)
                                        If pole Like pref.Trim Then
                                            IndPole = k
                                            boolfindPole = True
                                            Exit For
                                        ElseIf IsNothing(pole) = True Then
                                            IndPole = k
                                            boolfindPole = False
                                            Exit For
                                        Else
                                            IndPole = k
                                        End If
                                    Next k
                                    If boolfindPole = True Then
                                        ArrayWord(IndPole, i) = FinalPost
                                    Else
                                        ArrayWord(IndPole, i) = FinalPost
                                        ArrayWord(IndPole, 0) = pref
                                    End If
                                    ArrayText(3, j) = 1
                                End If
                            End If
                        End If

                    Next j
                Next i
            End If
            If IsArray(ArrayWord) = True Then
                If ArrayWord.GetUpperBound(1) > 0 Then
                    'запускаем эксель
                    Dim xlApp As Excel.Application = Nothing
                    Dim xlBook As Excel.Workbook = Nothing
                    Dim xlSheets As Excel.Worksheet = Nothing
                    xlApp = CreateObject("Excel.Application")
                    xlBook = xlApp.Workbooks.Add()
                    xlApp.Visible = True
                    Dim xlsheet As Excel.Worksheet
                    xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
                    xlsheet.Activate()
                    Dim posEx As Integer = 1
                    For i As Integer = 0 To ArrayWord.GetUpperBound(0)
                        Dim val As String = ArrayWord(i, 0)
                        If IsNothing(val) = False Then
                            xlsheet.Cells(1, i + 1) = val
                        End If
                    Next i
                    For i As Integer = 1 To ArrayWord.GetUpperBound(1)
                        For j As Integer = 0 To ArrayWord.GetUpperBound(0)
                            Dim Rez As String = ArrayWord(j, i)
                            If IsNothing(Rez) = False Then
                                If Rez.Trim.Length > 0 Then
                                    xlsheet.Cells(posEx + 1, j + 1) = ArrayWord(j, i)
                                End If
                            End If
                        Next j
                        posEx += 2
                    Next i
                    xlsheet = Nothing
                    xlBook = Nothing
                    xlApp = Nothing
                End If
            End If
line1:
        End Using
    End Sub
End Class
