Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class InsertAssemblyByGroupsSection
    <CommandMethod("InsertAssemblyByGroupsSection")>
    Public Sub InsertAssemblyByGroupsSection()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim acCurDb As Database = acDoc.Database
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
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
        Dim resBlk As PromptEntityResult = Nothing
        Dim optBlk As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите сечение: ")
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim count As Integer = 0
        Try
            Using tr
                resBlk = m_Editor.GetEntity(optBlk)
                Dim objID As ObjectId = resBlk.ObjectId
                Dim Obj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(objID, OpenMode.ForRead)
                If Not (TypeOf Obj Is Autodesk.Civil.DatabaseServices.SampleLine) Then Exit Sub
                Dim SampleLineObj As Autodesk.Civil.DatabaseServices.SampleLine = Obj
                Dim IDSGroup As ObjectId = SampleLineObj.GroupId
                Dim GRoupLineObj As Autodesk.Civil.DatabaseServices.SampleLineGroup = tr.GetObject(IDSGroup, OpenMode.ForRead)
                Dim ObjCol As ObjectIdCollection = GRoupLineObj.GetSampleLineIds
                Dim ArraySamplLine As Double(,) = Nothing
                Dim countArraySampLine As Integer = 0
                For Each id As ObjectId In ObjCol
                    Dim sLine As Autodesk.Civil.DatabaseServices.SampleLine = tr.GetObject(id, OpenMode.ForRead)
                    ReDim Preserve ArraySamplLine(1, countArraySampLine)
                    ArraySamplLine(0, countArraySampLine) = sLine.Station
                    ArraySamplLine(1, countArraySampLine) = sLine.Number
                    countArraySampLine += 1
                Next
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'если сечения найдены тогда вставляем станции в корридор
                If IsArray(ArraySamplLine) = True Then
                    'выбираем коридор
                    Dim resBlk1 As PromptEntityResult = Nothing
                    Dim optBlk1 As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите коридор: ")
                    resBlk1 = m_Editor.GetEntity(optBlk1)
                    Dim objID1 As ObjectId = resBlk1.ObjectId
                    Dim Obj1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(objID1, OpenMode.ForRead)
                    If Not (TypeOf Obj1 Is Autodesk.Civil.DatabaseServices.Corridor) Then Exit Sub
                    Dim CorridorObj As Autodesk.Civil.DatabaseServices.Corridor = Obj1
                    'минимальный и максимальный пикеты для региона
                    Dim PKMin As Double = -1
                    Dim PKMax As Double = -1
                    Dim blineReg As BaselineRegion = Nothing 'регион
                    'перебираем массив с сечениями
                    For i As Integer = 0 To ArraySamplLine.GetUpperBound(1)
                        Dim PKSLine As Double = ArraySamplLine(0, i) 'пикет вставки сечения
                        Dim NumbSLine As Integer = ArraySamplLine(1, i) 'номер линии
                        'проверям находится сечение в регионе, (ищем нужный регион)
                        If Not (PKSLine >= PKMin And PKSLine <= PKMax) Then
                            'если точка не в диапазоне пикетов, ищем регион
                            Dim BLines As BaselineCollection = Nothing 'коллекция базовых линий
                            Dim BLine As Baseline = Nothing 'базовая линия
                            If IsNothing(CorridorObj) = False Then
                                BLines = CorridorObj.Baselines 'получаем коллекцию базовых линий
                                'если в коллекции более 1 базовой линии запускаем диалоговое окно с выбором
                                If BLines.Count > 1 Then
                                    Dim FormBline As AssemblyByGroupsSection = New AssemblyByGroupsSection
                                    Dim ArrayList As String() = Nothing
                                    Dim CountArrayList As Integer = 0
                                    For Each bLine1 As Baseline In BLines
                                        ReDim Preserve ArrayList(CountArrayList)
                                        ArrayList(CountArrayList) = bLine1.Name
                                        CountArrayList += 1
                                    Next
                                    FormBline.ListBox1.Items.AddRange(ArrayList)
                                    FormBline.ShowDialog()
                                    Dim NAmeBline As String = FormBline.ListBox1.SelectedItem
                                    For Each bLine1 As Baseline In BLines
                                        If bLine1.Name Like NAmeBline Then
                                            BLine = bLine1
                                            Exit For
                                        End If
                                    Next
                                Else
                                    BLine = BLines.Item(0)
                                End If
                                If IsNothing(BLine) = False Then
                                    'ищем регионы
                                    Dim blineRegs As BaselineRegionCollection = BLine.BaselineRegions
                                    For Each BlineRegion As BaselineRegion In blineRegs
                                        If (PKSLine >= BlineRegion.StartStation And PKSLine <= BlineRegion.EndStation) Then
                                            blineReg = BlineRegion
                                            PKMin = BlineRegion.StartStation
                                            PKMax = BlineRegion.EndStation
                                            GoTo line1
                                        End If
                                    Next
                                End If

                            End If
                        End If

line1:
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        If IsNothing(blineReg) = False Then
                            If (PKSLine >= PKMin And PKSLine <= PKMax) Then
                                Dim SortSt As Double() = blineReg.SortedStations
                                Dim boolSt As Boolean = False
                                If IsArray(SortSt) = True Then
                                    For j As Integer = 0 To SortSt.Length - 1
                                        If Math.Abs(SortSt(j) - PKSLine) <= 0.01 Then
                                            boolSt = True
                                            Exit For
                                        End If
                                    Next
                                End If
                                If boolSt = False Then
                                    blineReg.AddStation(PKSLine, "Сечение " & NumbSLine)
                                End If
                            End If
                        End If
                    Next
                End If
                tr.Commit()
                locker.Dispose()
            End Using
        Catch ex As System.Exception
            MsgBox("Ошибка. Макро прервано!!!")
        End Try
    End Sub
End Class
